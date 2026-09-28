"""明日值班提醒。

从 Outlook 导出的 CSV 月度班表中读取排班，找出「明天」的值班安排，
并通过企业微信群机器人 Webhook 发送提醒。

用法::

    python main.py                      # 提醒明天的值班
    python main.py --date 2026-09-26    # 提醒指定日期
    python main.py --dry-run            # 只打印消息，不发送

班表来源：默认从 SMB 共享读取（.env 里配 SMB_*），读不到自动回落到本地 outlook/。
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import logging
import logging.handlers
import os
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OUTLOOK_DIR = BASE_DIR / "outlook"
ENV_FILE = BASE_DIR / ".env"
CONTACTS_FILE = BASE_DIR / "contacts.json"
LOG_DIR = BASE_DIR / "logs"

WEBHOOK_BASE = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send"

# Outlook 导出的日期形如 1900/1/30，只有「日」有意义，月份需要按跨月推算
DATE_CELL_RE = re.compile(r"^\d{4}/\d{1,2}/(\d{1,2})$")
TITLE_RE = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月")
NOTE_RE = re.compile(r"^[*#]")
KEY_RE = re.compile(r"(key=)([^&\s]+)")
COLUMNS = 7

WEEKDAY_NAMES = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

logger = logging.getLogger("duty")


# --------------------------------------------------------------------------- #
# 环境与小工具
# --------------------------------------------------------------------------- #
def setup_logging(level: str = "INFO", log_dir: Path = LOG_DIR) -> None:
    """同时输出到控制台和 logs/duty-YYYY-MM-DD.log（按天轮转，保留 30 天）。"""
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.handlers.clear()
    logger.propagate = False

    fmt = logging.Formatter(
        fmt="%(asctime)s %(levelname)-7s [%(funcName)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console = logging.StreamHandler(stream=sys.stdout)
    console.setFormatter(fmt)
    logger.addHandler(console)

    # smbprotocol 关闭连接时会刷 "socket aborted by peer"，属于正常现象，屏蔽掉
    logging.getLogger("smbprotocol").setLevel(logging.ERROR)

    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        file_handler = logging.handlers.TimedRotatingFileHandler(
            log_dir / "duty.log",
            when="midnight",
            interval=1,
            backupCount=30,
            encoding="utf-8",
        )
        file_handler.suffix = "%Y-%m-%d"
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)
    except OSError as exc:  # 目录不可写时不阻断主流程
        logger.warning("无法写入日志文件 %s：%s", log_dir, exc)


def mask_url(url: str) -> str:
    """日志里隐藏 webhook key。"""
    return KEY_RE.sub(lambda m: m.group(1) + m.group(2)[:8] + "***", url)
def load_env(path: Path) -> dict[str, str]:
    """读取 .env（不引入第三方依赖）。"""
    env: dict[str, str] = {}
    if not path.exists():
        return env
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip().removeprefix("export ").strip()] = value.strip().strip("'\"")
    return env


def load_contacts(path: Path) -> dict[str, str]:
    """读取「姓名 -> 手机号」映射，用于 @ 提醒。

    支持两种格式::

        contacts.json  {"Jessie": "85259887686", ...}
        contacts.txt   Jessie<TAB>85259887686   （也支持逗号/冒号分隔）
    """
    if not path.exists():
        logger.warning("未找到通讯录 %s，消息中将不会 @ 任何人", path.name)
        return {}

    text = path.read_text(encoding="utf-8-sig").strip()
    if not text:
        return {}

    if path.suffix.lower() == ".json":
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            logger.error("通讯录 %s 格式错误：%s", path.name, exc)
            return {}
        if not isinstance(data, dict):
            logger.error("通讯录 %s 应为 {姓名: 手机号} 的对象", path.name)
            return {}
        return {str(k).strip(): str(v).strip() for k, v in data.items() if k}

    contacts: dict[str, str] = {}
    for line in text.splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        parts = re.split(r"[\t,:：]", line.strip(), maxsplit=1)
        if len(parts) == 2:
            name, mobile = parts[0].strip(), parts[1].strip()
            if name and mobile:
                contacts[name] = mobile
    return contacts


# --------------------------------------------------------------------------- #
# 班表来源：SMB 共享 / 本地目录
# --------------------------------------------------------------------------- #
@dataclass
class SmbConfig:
    server: str
    share: str
    subdir: str
    username: str
    password: str
    port: int = 445

    @classmethod
    def from_env(cls, env: dict[str, str]) -> "SmbConfig | None":
        server = env.get("SMB_SERVER", "").strip()
        if not server:
            return None
        # SMB_DIR 形如 "Settlement Internal\结算部排班表"：第一段是共享名
        raw_dir = env.get("SMB_DIR", "").strip().strip("\\/")
        share, _, subdir = raw_dir.partition("\\")
        share = env.get("SMB_SHARE", "").strip() or share
        if not share:
            logger.warning("SMB_DIR 未包含共享名，例如 SMB_DIR=ShareName\\子目录")
            return None
        return cls(
            server=server,
            share=share,
            subdir=subdir.replace("/", "\\"),
            username=env.get("SMB_USERNAME", "").strip(),
            password=env.get("SMB_PASSWORD", ""),
            port=int(env.get("SMB_PORT", "445") or 445),
        )

    def unc(self, *parts: str) -> str:
        path = "\\".join(p for p in (self.share, self.subdir, *parts) if p)
        return f"\\\\{self.server}\\{path}"

    def __str__(self) -> str:
        return f"\\\\{self.server}\\{self.share}\\{self.subdir}".rstrip("\\")


def smb_list_files(cfg: SmbConfig) -> list[tuple[str, float]]:
    """列出共享目录里的班表文件（xlsx/csv），返回 [(文件名, 修改时间)]，按时间倒序。"""
    import smbclient  # 延迟导入：本地模式下不强制依赖

    smbclient.register_session(
        cfg.server, username=cfg.username, password=cfg.password, port=cfg.port
    )
    try:
        entries = []
        with smbclient.scandir(cfg.unc()) as it:
            for entry in it:
                if entry.is_file() and entry.name.lower().endswith((".xlsx", ".xlsm", ".csv")):
                    entries.append((entry.name, entry.stat().st_mtime))
        entries.sort(key=lambda x: x[1], reverse=True)
        logger.info("SMB %s 下共 %d 个班表文件：%s", cfg, len(entries), ", ".join(n for n, _ in entries))
        return entries
    finally:
        smbclient.delete_session(cfg.server, port=cfg.port)


def smb_read_bytes(cfg: SmbConfig, name: str) -> bytes:
    """从 SMB 读取一个文件的原始字节。"""
    import smbclient

    smbclient.register_session(
        cfg.server, username=cfg.username, password=cfg.password, port=cfg.port
    )
    try:
        with smbclient.open_file(cfg.unc(name), mode="rb") as fh:
            data = fh.read()
        logger.info("已读取 SMB 文件 %s（%d 字节）", name, len(data))
        return data
    finally:
        smbclient.delete_session(cfg.server, port=cfg.port)


def _pick_names(names: list[str], target: dt.date) -> list[str]:
    """优先挑文件名含目标年份的班表（如 «结算部 排班表 2026.xlsx»），否则按名称倒序。"""
    year_names = [n for n in names if f"{target:%Y}" in n]
    return year_names or sorted(names, reverse=True)


def fetch_schedules(mode: str, csv_dir: Path, env: dict[str, str], target: dt.date):
    """取得 [(来源名, {日期: DaySchedule})]，已按优先级排序。默认走 SMB。"""
    if mode in ("auto", "smb"):
        cfg = SmbConfig.from_env(env)
        if cfg is None:
            logger.warning("未配置 SMB_SERVER / SMB_DIR，跳过 SMB 读取")
        else:
            try:
                names = [n for n, _ in smb_list_files(cfg)]
                docs: list[tuple[str, dict]] = []
                for name in _pick_names(names, target)[:2]:
                    data = smb_read_bytes(cfg, name)
                    if name.lower().endswith(".csv"):
                        docs.append((name, parse_csv_text(data.decode("utf-8-sig", errors="replace"), name)))
                    else:
                        docs.extend(parse_xlsx_bytes(data, name, target))
                if docs:
                    return docs
                logger.warning("SMB %s 下没有可用的班表文件", cfg)
            except Exception as exc:  # SMB 不可用时不影响本地兜底
                logger.exception("读取 SMB 失败：%s", exc)
                if mode == "smb":
                    raise

    candidates = sorted(csv_dir.glob("*.csv"))
    if not candidates:
        raise FileNotFoundError(f"目录中没有找到 CSV：{csv_dir}")
    logger.info("本地目录 %s 找到 %d 份班表：%s", csv_dir, len(candidates), ", ".join(p.name for p in candidates))
    picked = _pick_names([p.name for p in candidates], target)[:3]
    return [(name, parse_csv_text((csv_dir / name).read_text(encoding="utf-8-sig", errors="replace"), name)) for name in picked]


# --------------------------------------------------------------------------- #
# 班表解析（CSV / XLSX 共用一套「月历网格」解析）
# --------------------------------------------------------------------------- #
@dataclass
class DaySchedule:
    date: dt.date
    duties: list[tuple[str, str]] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def is_empty(self) -> bool:
        return not self.duties and not self.notes


def parse_cell(text: str) -> tuple[list[tuple[str, str]], list[str]]:
    """把单元格文本拆成 [(任务, 人员)] 和 [备注]。"""
    duties: list[tuple[str, str]] = []
    notes: list[str] = []
    for line in text.replace("\r\n", "\n").split("\n"):
        line = line.strip().lstrip("\ufeff")
        if not line:
            continue
        if NOTE_RE.match(line):
            notes.append(line.lstrip("*#").strip())
            continue
        if re.search(r"[:：]", line):
            task, person = re.split(r"[:：]", line, maxsplit=1)
            task, person = task.strip(), person.strip()
            if task and person:
                duties.append((task, person))
            continue
        notes.append(line)
    return duties, notes


def _resolve_week_dates(days: list[int | None], year: int, month: int, prev_day: int):
    """把一列「日」推算成完整日期（自动处理跨月）。"""
    dates: list[dt.date | None] = []
    for day in days:
        if day is None:
            dates.append(None)
            continue
        if prev_day and day + 15 < prev_day:  # 日期回跳 => 进入下个月
            month += 1
            if month > 12:
                month, year = 1, year + 1
        prev_day = day
        dates.append(dt.date(year, month, day))
    return dates, year, month, prev_day


def cell_day(value) -> int | None:
    """日期单元格 -> 日号。CSV 里是 "1900/1/30" 文本，XLSX 里是 date/datetime。"""
    if isinstance(value, dt.datetime):
        return value.day
    if isinstance(value, dt.date):
        return value.day
    if isinstance(value, str):
        m = DATE_CELL_RE.match(value.strip())
        if m:
            return int(m.group(1))
    return None


def cell_text(value) -> str:
    """非日期单元格 -> 文本。"""
    if isinstance(value, (dt.date, dt.datetime)):
        return ""
    return "" if value is None else str(value).strip()


def parse_grid(grid, source: str = "") -> dict[dt.date, DaySchedule]:
    """解析 Outlook 月历网格（7 列：日到六）为 {日期: DaySchedule}。

    网格第一行是标题（含年月），第二行是星期名；之后是「日期行 / 内容行」交替出现
    （CSV 里一个日期行会跟多行内容，XLSX 里通常一行日期一行内容）。
    """
    if not grid:
        return {}

    title = " ".join(cell_text(c) for c in grid[0])
    match = TITLE_RE.search(title)
    base_year = int(match.group(1)) if match else dt.date.today().year
    base_month = int(match.group(2)) if match else dt.date.today().month

    schedule: dict[dt.date, DaySchedule] = {}
    year, month = base_year, base_month
    prev_day = 0
    first_week = True
    week_dates: list[dt.date | None] = []
    buffers: list[str] = []

    def flush() -> None:
        for date, cell in zip(week_dates, buffers):
            if date is None or not cell.strip():
                continue
            duties, notes = parse_cell(cell)
            entry = schedule.setdefault(date, DaySchedule(date))
            entry.duties.extend(duties)
            entry.notes.extend(notes)

    for row in grid:
        body = list(row)[:COLUMNS] + [""] * (COLUMNS - min(len(row), COLUMNS))
        days = [cell_day(c) for c in body]
        if sum(d is not None for d in days) >= 2:  # 一行里有多个日期 => 日期行
            if first_week and days[0] not in (None, 1):
                month -= 1  # 首行不是 1 号 => 该周属于上个月
                if month < 1:
                    month, year = 12, year - 1
            first_week = False
            flush()
            week_dates, year, month, prev_day = _resolve_week_dates(days, year, month, prev_day)
            buffers = [""] * COLUMNS
            continue
        if not week_dates:
            continue
        for i, cell in enumerate(body):
            text = cell_text(cell)
            if text:
                buffers[i] = f"{buffers[i]}\n{text}" if buffers[i] else text
    flush()

    if schedule:
        dates = sorted(schedule)
        logger.info(
            "解析 %s：%d 条排班（%s ~ %s）",
            source or "CSV", len(schedule), dates[0], dates[-1],
        )
    else:
        logger.warning("解析 %s：没有识别到任何排班，请检查格式", source or "班表")
    return schedule


def parse_csv_text(text: str, source: str = "") -> dict[dt.date, DaySchedule]:
    """解析 Outlook 月历 CSV 文本。单元格内含换行，必须让 csv 模块按流解析。"""
    return parse_grid(list(csv.reader(io.StringIO(text, newline=""))), source)


def parse_xlsx_bytes(data: bytes, source: str, target: dt.date) -> list[tuple[str, dict]]:
    """解析 xlsx 班表：每个月份一个 sheet，返回 [(来源名, {日期: DaySchedule})]。

    优先解析目标月份所在 sheet，其次相邻月份（跨月边缘的日期会用得上）。
    """
    from python_calamine import CalamineWorkbook  # 延迟导入

    workbook = CalamineWorkbook.from_object(io.BytesIO(data))
    wanted = {(target.month - 1) % 12 or 12, target.month, target.month % 12 + 1}

    docs: list[tuple[str, dict]] = []
    for name in workbook.sheet_names:
        m = re.search(r"\d+", name)
        sheet_month = int(m.group()) if m else None
        if sheet_month is not None and sheet_month not in wanted:
            continue
        grid = workbook.get_sheet_by_name(name).to_python()
        label = f"{source}::{name}"
        docs.append((label, parse_grid(grid, label)))
        if len(docs) >= 3:
            break
    return docs


# --------------------------------------------------------------------------- #
# 消息构建与发送
# --------------------------------------------------------------------------- #
def build_message(
    target: dt.date,
    entry: DaySchedule | None,
    source: str,
    contacts: dict[str, str] | None = None,
) -> tuple[str, list[str]]:
    """生成纯文本消息，并返回需要 @ 的手机号列表。

    用手机号 @ 人只能走 text 类型（markdown 不支持 mentioned_mobile_list）。
    """
    contacts = contacts or {}
    lines = [
        "【明日值班提醒】",
        f"{target:%Y-%m-%d} {WEEKDAY_NAMES[target.weekday()]}",
        "",
    ]
    duties = entry.duties if entry else []
    notes = entry.notes if entry else []

    mobiles: list[str] = []
    if duties:
        lines.append("值班安排：")
        for task, person in duties:
            lines.append(f"· {task}：{person}")
            mobile = contacts.get(person.strip())
            if mobile and mobile not in mobiles:
                mobiles.append(mobile)
    else:
        lines.append("明天没有排到值班任务。")
    if notes:
        lines.append("")
        lines.append("备注：")
        for note in notes:
            lines.append(f"  · {note}")
    lines.append("")
    lines.append(f"来源：{source}")
    return "\n".join(lines), mobiles


def send_wechat_text(
    webhook_url: str, content: str, mentioned_mobile_list: list[str] | None = None
) -> None:
    payload = json.dumps(
        {
            "msgtype": "text",
            "text": {
                "content": content,
                "mentioned_mobile_list": mentioned_mobile_list or [],
            },
        },
        ensure_ascii=False,
    ).encode("utf-8")
    request = urllib.request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    logger.info(
        "发送企业微信 Webhook：%s（%d 字节，@ %s）",
        mask_url(webhook_url), len(payload), mentioned_mobile_list or "-",
    )
    with urllib.request.urlopen(request, timeout=10) as response:
        result = json.loads(response.read().decode("utf-8"))
    if result.get("errcode") != 0:
        raise RuntimeError(f"Webhook 返回错误：{result}")
    logger.info("Webhook 返回：errcode=%s errmsg=%s", result.get("errcode"), result.get("errmsg"))


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def collect_schedule(documents, target: dt.date) -> tuple[DaySchedule | None, str]:
    """在已解析好的 [(来源名, {日期: DaySchedule})] 中查找 ``target`` 的排班。"""
    for name, schedule in documents:
        entry = schedule.get(target)
        if entry and not entry.is_empty:
            logger.info("命中 %s 的排班，来源：%s", target, name)
            return entry, name
    logger.warning("%s 在所有班表中都没有排班记录", target)
    first = documents[0][0] if documents else "无班表"
    return None, f"{first}(未找到排班)"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="读取 Outlook 值班 CSV，提醒明天的值班安排")
    parser.add_argument("--date", help="提醒哪一天（YYYY-MM-DD），默认明天")
    parser.add_argument("--csv-dir", default=str(OUTLOOK_DIR), help="本地 CSV 目录（local 模式）")
    parser.add_argument(
        "--source", choices=["auto", "smb", "local"], default="auto", help="班表来源，默认先 SMB 后本地"
    )
    parser.add_argument("--dry-run", action="store_true", help="只打印消息，不发送 Webhook")
    parser.add_argument(
        "--contacts", default=str(CONTACTS_FILE), help=f"通讯录文件，默认 {CONTACTS_FILE.name}"
    )
    parser.add_argument("--log-dir", default=str(LOG_DIR), help=f"日志目录，默认 {LOG_DIR}")
    parser.add_argument("--log-level", default="INFO", help="DEBUG/INFO/WARNING/ERROR，默认 INFO")
    args = parser.parse_args(argv)

    setup_logging(args.log_level, Path(args.log_dir))
    logger.info("=== 值班提醒开始 ===")

    target = (
        dt.datetime.strptime(args.date, "%Y-%m-%d").date()
        if args.date
        else dt.date.today() + dt.timedelta(days=1)
    )
    logger.info("提醒日期：%s", target)

    env = {**load_env(ENV_FILE), **os.environ}
    try:
        documents = fetch_schedules(args.source, Path(args.csv_dir), env, target)
    except FileNotFoundError as exc:
        logger.error("%s", exc)
        return 1
    entry, source = collect_schedule(documents, target)

    contacts = load_contacts(Path(args.contacts))
    message, mobiles = build_message(target, entry, source, contacts)
    for name, mobile in contacts.items():
        if mobile in mobiles:
            logger.debug("将 @ %s（%s）", name, mobile)

    logger.info("消息内容：\n%s", message)

    if args.dry_run:
        logger.info("[dry-run] 未发送。")
        return 0

    webhook_url = env.get("WEBHOOK_URL") or ""
    if not webhook_url:
        webhook_key = env.get("WEBHOOK_KEY", "").strip()
        if not webhook_key:
            logger.error("缺少 WEBHOOK_KEY / WEBHOOK_URL，请在 .env 中配置")
            return 1
        webhook_url = f"{WEBHOOK_BASE}?key={webhook_key}"

    try:
        send_wechat_text(webhook_url, message, mobiles)
    except (urllib.error.URLError, RuntimeError, OSError) as exc:
        logger.exception("发送失败：%s", exc)
        return 1
    logger.info("=== 值班提醒完成 ===")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:  # 兜底：任何异常都写进日志，避免计划任务里静默失败
        logger.exception("运行异常终止")
        raise
