import os
from pathlib import Path
from dotenv import load_dotenv
import logging

# Load environment variables
env_path = Path(__file__).resolve().parent.parent / "config" / ".env"

load_dotenv(env_path)

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Create output and barcode directories
OUTPUT_DIR = Path("output").absolute()
BARCODE_DIR = Path("barcode_images").absolute()
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(BARCODE_DIR, exist_ok=True)

# PostgreSQL Database Configuration
DATABASE_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "port": int(os.getenv("DB_PORT", 5432)),
    "database": os.getenv("DB_NAME", "dl_generator"),
    "user": os.getenv("DB_USER", "postgres"),
    "password": os.getenv("DB_PASSWORD", "$PMadr!d1234")
}

# FTP Configuration
FTP_CONFIG = {
    "hostname": os.getenv("OMKT_FTP_HOSTNAME"),
    "port": int(os.getenv("OMKT_FTP_PORT", 21)),
    "username": os.getenv("OMKT_FTP_USERNAME"),
    "password": os.getenv("OMKT_FTP_PASSWORD")
}

# Lark/Feishu Configuration
APP_ID = os.getenv("APP_ID")
APP_SECRET = os.getenv("APP_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI")
AUTH_URL = f"https://open.larksuite.com/open-apis/authen/v1/authorize?app_id={APP_ID}&redirect_uri={REDIRECT_URI}"
TOKEN_URL = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
USER_ACCESS_TOKEN_URL = "https://open.larksuite.com/open-apis/authen/v1/oidc/access_token"
USER_INFO_URL = "https://open.larksuite.com/open-apis/authen/v1/user_info"
REFRESH_TOKEN_URL = "https://open.larksuite.com/open-apis/authen/v1/oidc/refresh_access_token"

# Google Sheets Configuration
SERVICE_ACCOUNT_JSON = "config/dl_automation_sheet.json"
SPREADSHEET_ID = "1M0Vmmf9HfPRB0oSeJR_xUAZPPpTJ4xsR3gYDQMvnu5k"
SHEET_NAME = "LetterHeads"

# Session Configuration
SESSION_FILE = Path("sessions.json")
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./app.db")