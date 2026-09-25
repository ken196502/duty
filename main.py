"""明日值班提醒。

从 Outlook 导出的 CSV 月度班表中读取排班，找出「明天」的值班安排，
并通过企业微信群机器人 Webhook 发送提醒。

用法::

    python main.py                      # 提醒明天的值班
    python main.py --date 2026-09-26    # 提醒指定日期
    python main.py --dry-run            # 只打印消息，不发送
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
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


def load_mentions(env: dict[str, str]) -> dict[str, str]:
    """可选的人员 -> 企业微信 userid 映射，用于在消息里 @ 到具体的人。

    .env 中可配置（两种写法都支持）::

        MENTIONS=Joey:zhangsan,Matthew:lisi
        MENTIONS={"Joey": "zhangsan"}
    """
    raw = env.get("MENTIONS", "")
    if not raw:
        return {}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        data = None
    if isinstance(data, dict):
        return {str(k).strip(): str(v).strip() for k, v in data.items()}
    mapping: dict[str, str] = {}
    for item in raw.split(","):
        if ":" in item:
            name, userid = item.split(":", 1)
            mapping[name.strip()] = userid.strip()
    return mapping


# --------------------------------------------------------------------------- #
# CSV 解析
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


def parse_csv(path: Path) -> dict[dt.date, DaySchedule]:
    """把一个 Outlook 月历 CSV 解析成 {日期: DaySchedule}。

    CSV 结构：第一行是标题（含年月），第二行是星期名，紧接着一行是日期
    （形如 1900/1/30，只有「日」有效）；之后每一行都是这一周 7 个单元格
    的续行，直到下一个日期行出现。
    """
    # 单元格内含换行，必须让 csv 模块读取原始流，不能用 read_text().splitlines()
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        return {}

    title = " ".join(rows[0])
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

    for row in rows:
        cells = (list(row) + [""] * (COLUMNS + 1))[: COLUMNS + 1]
        body = [cells[i].strip() for i in range(1, COLUMNS + 1)]
        if DATE_CELL_RE.match(body[0]):
            days: list[int | None] = []
            for cell in body:
                m = DATE_CELL_RE.match(cell)
                days.append(int(m.group(1)) if m else None)
            # 首行不是 1 号 => 该周属于上个月
            if first_week and days[0] not in (None, 1):
                month -= 1
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
            if cell:
                buffers[i] = f"{buffers[i]}\n{cell}" if buffers[i] else cell
    flush()

    if schedule:
        dates = sorted(schedule)
        logger.info(
            "解析 %s：%d 条排班（%s ~ %s）",
            path.name, len(schedule), dates[0], dates[-1],
        )
    else:
        logger.warning("解析 %s：没有识别到任何排班，请检查 CSV 格式", path.name)
    return schedule


# --------------------------------------------------------------------------- #
# 消息构建与发送
# --------------------------------------------------------------------------- #
def build_message(
    target: dt.date,
    entry: DaySchedule | None,
    source: str,
    mentions: dict[str, str] | None = None,
) -> str:
    mentions = mentions or {}
    lines = [
        "## 明日值班提醒",
        f"**日期**：{target:%Y-%m-%d} {WEEKDAY_NAMES[target.weekday()]}",
        "",
    ]
    duties = entry.duties if entry else []
    notes = entry.notes if entry else []
    if duties:
        lines.append("**值班安排：**")
        for task, person in duties:
            who = person
            userid = mentions.get(person)
            if userid:
                who = f"{person} (<@{userid}>)"
            lines.append(f"> {task}：**{who}**")
    else:
        lines.append("> 明天没有排到值班任务。")
    if notes:
        lines.append("")
        lines.append("**备注：**")
        for note in notes:
            lines.append(f"> {note}")
    lines.append("")
    lines.append(f'<font color="comment">来源：{source}</font>')
    return "\n".join(lines)


def send_wechat_markdown(webhook_url: str, content: str) -> None:
    payload = json.dumps(
        {"msgtype": "markdown", "markdown": {"content": content}},
        ensure_ascii=False,
    ).encode("utf-8")
    request = urllib.request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    logger.info("发送企业微信 Webhook：%s（%d 字节）", mask_url(webhook_url), len(payload))
    with urllib.request.urlopen(request, timeout=10) as response:
        result = json.loads(response.read().decode("utf-8"))
    if result.get("errcode") != 0:
        raise RuntimeError(f"Webhook 返回错误：{result}")
    logger.info("Webhook 返回：errcode=%s errmsg=%s", result.get("errcode"), result.get("errmsg"))


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def collect_schedule(csv_dir: Path, target: dt.date) -> tuple[DaySchedule | None, str]:
    """在所有 CSV 中查找 ``target`` 的排班，返回 (排班, 来源文件名)。"""
    candidates = sorted(csv_dir.glob("*.csv"))
    if not candidates:
        raise FileNotFoundError(f"目录中没有找到 CSV：{csv_dir}")
    logger.info("在 %s 找到 %d 份班表：%s", csv_dir, len(candidates), ", ".join(p.name for p in candidates))

    # 优先使用文件名正好对应目标月份的那份表
    preferred = f"{target:%Y-%m}.csv"
    candidates.sort(key=lambda p: (p.name != preferred, p.name))

    for path in candidates:
        entry = parse_csv(path).get(target)
        if entry and not entry.is_empty:
            logger.info("命中 %s 的排班，来源：%s", target, path.name)
            return entry, path.name
    logger.warning("%s 在所有班表中都没有排班记录", target)
    return None, f"{preferred}(未找到排班)"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="读取 Outlook 值班 CSV，提醒明天的值班安排")
    parser.add_argument("--date", help="提醒哪一天（YYYY-MM-DD），默认明天")
    parser.add_argument("--csv-dir", default=str(OUTLOOK_DIR), help="CSV 所在目录")
    parser.add_argument("--dry-run", action="store_true", help="只打印消息，不发送 Webhook")
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

    csv_dir = Path(args.csv_dir)
    try:
        entry, source = collect_schedule(csv_dir, target)
    except FileNotFoundError as exc:
        logger.error("%s", exc)
        return 1

    env = {**load_env(ENV_FILE), **os.environ}
    message = build_message(target, entry, source, load_mentions(env))

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
        send_wechat_markdown(webhook_url, message)
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
