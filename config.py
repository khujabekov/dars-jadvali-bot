import os
from pathlib import Path
from dotenv import load_dotenv

# .env faylni yuklash
BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(dotenv_path=ENV_PATH)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

# Guruh ID (masalan: -1001234567890)
raw_group_id = os.getenv("GROUP_ID", "").strip()
try:
    GROUP_ID = int(raw_group_id) if raw_group_id else None
except ValueError:
    GROUP_ID = raw_group_id

# Adminlar ID ro'yxati (vergul bilan ajratilgan, masalan: 12345678,87654321)
raw_admins = os.getenv("ADMIN_IDS", "").strip()
ADMIN_IDS = []
if raw_admins:
    for item in raw_admins.split(","):
        item = item.strip()
        if item.isdigit():
            ADMIN_IDS.append(int(item))

# Jadval yuboriladigan vaqt (masalan: 07:30)
SEND_TIME = os.getenv("SEND_TIME", "07:30").strip()
try:
    SEND_HOUR, SEND_MINUTE = map(int, SEND_TIME.split(":"))
except Exception:
    SEND_HOUR, SEND_MINUTE = 7, 30

# Vaqt mintaqasi (default: Asia/Tashkent)
TIMEZONE = os.getenv("TIMEZONE", "Asia/Tashkent").strip()

# Yakshanba kuni xabar yuboriladimi? (Ha bo'lsa dam olish kuni haqida eslatma yuboradi, Yo'q bo'lsa jim turadi)
NOTIFY_ON_OFF_DAYS = os.getenv("NOTIFY_ON_OFF_DAYS", "false").lower() in ("true", "1", "yes", "ha")

# Fayllar yo'li
DATA_DIR = BASE_DIR / "data"
SCHEDULE_FILE = DATA_DIR / "jadval.xlsx"
TEMPLATE_FILE = DATA_DIR / "jadval_namuna.xlsx"

DATA_DIR.mkdir(parents=True, exist_ok=True)
