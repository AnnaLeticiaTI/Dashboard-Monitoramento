import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
AUTH_DIR = BASE_DIR / "auth_data"
EXPORT_DIR = Path("/tmp/dashboard_auditoria_exports") if os.environ.get("VERCEL") else BASE_DIR / "exports"
EXCEL_FILE = DATA_DIR / "dashboard.xlsx"
PAINT_EXCEL_FILE = DATA_DIR / "execucao_paint.xlsx"
CARTEIRA_EXCEL_FILE = DATA_DIR / "carteira_auditoria.xlsx"
AUTH_FILE = AUTH_DIR / "credentials.json"

APP_NAME = "Dashboard Auditoria"
VERSION = "3.5"
