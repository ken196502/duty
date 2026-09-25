# duty

从 Outlook 导出的月度值班 CSV 中读取排班，提醒**明天**的值班安排，并通过企业微信群机器人 Webhook 发送。

## 环境准备

`.env` 中配置群机器人 key：

```ini
WEBHOOK_KEY=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
# 或直接写完整地址（二选一）
# WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxx

# 可选：人员 -> 企业微信 userid，配置后消息里会 @ 到具体的人
# MENTIONS=Joey:zhangsan,Matthew:lisi
```

## 使用

```powershell
python main.py                      # 提醒明天的值班并发送
python main.py --dry-run            # 只打印消息，不发送（自测用）
python main.py --date 2026-09-26    # 指定日期
python main.py --csv-dir D:\share   # 指定 CSV 目录
python main.py --log-dir D:\logs    # 指定日志目录
python main.py --log-level DEBUG    # 更详细的日志
```

## 日志

同时输出到控制台和 `logs/duty.log`，按天轮转、保留 30 天（`logs/` 已加入 .gitignore）。
日志中会脱敏 webhook key（`key=ba055da0***`）。

## CSV 约定

把 Outlook「月视图」导出的 CSV 放在 `outlook/` 下，文件名建议为 `YYYY-MM.csv`（如 `2026-09.csv`）。

- 首行标题需含年月份（`2026 年 9 月`）；跨月前后的日期会按 30/31 → 1 自动推算月份
- 单元格内 `任务 : 人员` 为值班安排，`*` 开头的行作为「备注」（如 `*Wesley 放假`）

## 定时运行

Windows 计划任务，每工作日 18:00 执行：

```powershell
schtasks /create /tn "DutyReminder" /tr "python C:\Users\SimSettuser1\Documents\duty\main.py" /sc daily /st 18:00
```
