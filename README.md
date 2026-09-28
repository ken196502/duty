# duty

从 SMB 共享上的「结算部排班表」xlsx 读取排班，提醒**明天**的值班安排，
并通过企业微信群机器人 Webhook 发送（会按手机号 @ 到值班的人）。

## 安装

```powershell
uv sync
```

## 环境配置（.env）

```ini
WEBHOOK_KEY=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
# 或直接写完整地址（二选一）
# WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxx

SMB_SERVER=10.202.7.240
SMB_USERNAME=nas_batch_operation
SMB_PASSWORD=xxxxxxxx
SMB_DIR=Settlement Internal\结算部排班表   # 第一段是共享名，后面是子目录
# 可选：SMB_SHARE=ShareName（SMB_DIR 只有子目录时用）、SMB_PORT=445
```

## 使用

```powershell
.\.venv\Scripts\python.exe main.py            # 提醒明天的值班并发送
.\.venv\Scripts\python.exe main.py --dry-run  # 只打印消息，不发送（自测用）
uv run main.py --date 2026-09-29              # 指定日期
uv run main.py --source local                 # 强制读本地 outlook/*.csv（SMB 不可用时的兜底）
uv run main.py --csv-dir D:\share             # 本地模式下的 CSV 目录
uv run main.py --contacts C:\duty\contacts.json
uv run main.py --log-dir D:\logs --log-level DEBUG
```

## 班表来源

1. **SMB（默认）**：列出 `SMB_DIR` 下的 `*.xlsx/*.csv`，取文件名含目标年份的那份；
   xlsx 按「N 月」sheet 解析，只解析目标月及相邻月（跨月边缘的日期会用得上）。
2. **本地兜底**：SMB 连不上或没配时，读 `outlook/*.csv`（Outlook 月视图导出）。

两种格式都是同一套「月历网格」结构：标题行含年月（`2026 年 9 月`），
其次是星期行，之后「日期行 / 内容行」交替；单元格内 `任务 : 人员` 为值班安排，
`*` 开头的行作为备注（如 `*Wesley 放假`）。

## @ 提醒（通讯录）

`contacts.json` 维护「姓名 -> 手机号」，值班人员会自动被 @：

```json
{"Jessie": "85259887686", "Joey": "85253004454"}
```

也支持 `contacts.txt`（tab / 逗号 / 冒号分隔），用 `--contacts` 指定。

> 群机器人只有 **text** 类型支持按手机号 @（`mentioned_mobile_list`），
> markdown 不支持，所以消息是纯文本格式。
> 如果群里没 @ 到人，说明通讯录登记的号码格式不同，把号码改成 `+85259887686`
> 或 `+852-59887686` 再试一次。

## 日志

同时输出到控制台和 `logs/duty.log`，按天轮转、保留 30 天（`logs/` 已加入 .gitignore）。
日志中会脱敏 webhook key（`key=ba055da0***`）。

## 定时运行

Windows 计划任务，每工作日 18:00 执行（注意用虚拟环境的 python）：

```powershell
schtasks /create /tn "DutyReminder" /tr "C:\Users\SimSettuser1\Documents\duty\.venv\Scripts\python.exe C:\Users\SimSettuser1\Documents\duty\main.py" /sc daily /st 18:00
```
