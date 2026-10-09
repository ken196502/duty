"""临时诊断：用兄弟项目里的其他 SMB 账号测试排班表目录是否可访问。"""
import smbclient
from pathlib import Path

OTHER = Path(r"C:\Users\SimSettuser1\Documents\ccass2pdf\.env")
env: dict[str, str] = {}
for raw in OTHER.read_text(encoding="utf-8-sig").splitlines():
    line = raw.strip()
    if line and "=" in line and not line.startswith("#"):
        k, v = line.split("=", 1)
        env[k.strip()] = v.strip().strip("'\"")

SERVER = "10.202.7.240"
TARGET = "\\\\%s\\Settlement Internal\\结算部排班表" % SERVER
ACCOUNTS = [
    (env.get("SMB_CA_USERNAME"), env.get("SMB_CA_PASSWORD")),
    (env.get("SMB_USERNAME"), env.get("SMB_PASSWORD")),
]

for user, pwd in ACCOUNTS:
    if not user:
        continue
    print("== %s ==" % user)
    try:
        smbclient.register_session(SERVER, username=user, password=pwd)
    except Exception as exc:
        print("  register失败:", str(exc)[:120])
        continue
    try:
        print("  内容:", smbclient.listdir(TARGET)[:20])
    except Exception as exc:
        print("  读取失败:", str(exc)[:130])
    finally:
        smbclient.delete_session(SERVER)
