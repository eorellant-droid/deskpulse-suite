#!/usr/bin/env python3
# =============================================================================
# DESKPULSE SUITE — v1.2.6-last-fixed16-04-206
# =============================================================================
# Implementación consolidada según la especificación de arquitectura v1.2.1.
# Corregido: backfill limitado, escritura atómica, mouse tracking, locks, diálogos transient.
# =============================================================================


# =============================================================================
# IMPORTS
# =============================================================================
import ctypes
import csv
import io
import json
import logging
import os
import queue
import re
import shutil
import subprocess
import sys
import socket
import threading
import time
import zipfile
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, List, Dict, Any

try:
    from zoneinfo import ZoneInfo
    _TZ_BOLIVIA = ZoneInfo("America/La_Paz")
except Exception:
    _TZ_BOLIVIA = timezone(timedelta(hours=-4))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Third-party (must be installed) ---
try:
    import bcrypt
except ImportError:
    bcrypt = None  # Handled gracefully at runtime

try:
    import win32crypt
    _DPAPI_OK = True
except ImportError:
    win32crypt = None
    _DPAPI_OK = False

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes as crypto_hashes
    from cryptography.hazmat.backends import default_backend
    _CRYPTO_OK = True
except ImportError:
    _CRYPTO_OK = False

try:
    import mss
    import mss.tools
    _MSS_OK = True
except ImportError:
    _MSS_OK = False

try:
    from PIL import Image
    _PIL_OK = True
except ImportError:
    _PIL_OK = False

try:
    import psutil
    _PSUTIL_OK = True
except ImportError:
    _PSUTIL_OK = False

try:
    import pynput.keyboard as _pk
    import pynput.mouse as _pm
    _PYNPUT_OK = True
except ImportError:
    _PYNPUT_OK = False

try:
    import win32gui
    import win32con
    import win32process
    _WIN32_OK = True
except ImportError:
    _WIN32_OK = False

try:
    import gspread
    from google.oauth2.service_account import Credentials as _GSACredentials
    _GSPREAD_OK = True
except ImportError:
    _GSPREAD_OK = False


# =============================================================================
# CONSTANTS & STYLES
# =============================================================================

APP_TITLE    = "DeskPulse Suite"
APP_VERSION  = "v1.2.7-production"
APP_FEATURES = "Secure Admin Import • OT Cap Logic • Midnight-Safe Sessions • UI Polish • Self Ignore"
WIN_W, WIN_H = 425, 525

SPLASH_W, SPLASH_H        = 480, 270
SPLASH_DURATION_MS        = 3000
SPLASH_FADE_MS            = 600
SPLASH_FADE_STEPS         = 15

APP_HOME = Path(os.getenv("LOCALAPPDATA", str(Path.home() / "AppData" / "Local"))) / "DeskPulseSuite"
APP_HOME.mkdir(parents=True, exist_ok=True)

SECURE_DIR = APP_HOME / "secure"
SECURE_DIR.mkdir(parents=True, exist_ok=True)

RECORDS_DIR = APP_HOME / "records"
EXPORTS_DIR = APP_HOME / "exports"

LEGACY_CONFIG_FILE = APP_HOME / "config.json"
LEGACY_CREDENTIALS_FILE = APP_HOME / "credentials.json"

CONFIG_FILE = SECURE_DIR / "config.dpapi"
CREDENTIALS_FILE = SECURE_DIR / "credentials.dpapi"

SESSION_STATE_FILE = APP_HOME / "session_state.json"
SHEETS_QUEUE_FILE = APP_HOME / "sheets_queue.json"

APP_ID = "DeskPulseSuite.AgentMonitor"
DPAPI_ENTROPY = b"DeskPulseSuite|DPAPI|v1"


def _runtime_asset_dirs() -> list[Path]:
    """Search order for bundled assets across source, onedir and onefile builds."""
    dirs: list[Path] = []
    try:
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            dirs.append(Path(meipass))
    except Exception:
        pass
    try:
        if getattr(sys, "frozen", False):
            dirs.append(Path(sys.executable).resolve().parent)
    except Exception:
        pass
    dirs.append(Path(__file__).resolve().parent)
    seen: set[str] = set()
    unique: list[Path] = []
    for item in dirs:
        key = str(item)
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique



def _resolve_runtime_asset(*names: str) -> Path:
    for base in _runtime_asset_dirs():
        for name in names:
            candidate = base / name
            if candidate.exists():
                return candidate
    return _runtime_asset_dirs()[0] / names[0]


SPLASH_IMAGE_FILE = _resolve_runtime_asset("splash.png", "splash.jpg", "splash.gif")
APP_ICON_FILE = _resolve_runtime_asset("app.ico")

SHEETS_RETRY_INTERVAL = 60   # seconds between retry attempts
SHEETS_MAX_RETRIES    = 10   # max retries per queued event before discard

ADMIN_MAX_ATTEMPTS    = 3
ADMIN_LOCKOUT_SECONDS = 30

# Colors — Standardized 4-color palette
C_NAVY            = "#0D2149"   # Dark blue
C_INK             = "#1A1A1A"   # Almost black
C_GRAY            = "#6B7280"   # Medium gray
C_LIGHT           = "#D1D5DB"   # Light gray
C_CHALK           = "#F3F4F6"   # Chalk white

C_BG              = C_CHALK
C_SURFACE         = C_CHALK
C_TEXT            = C_INK
C_TEXT_SECONDARY  = C_GRAY
C_MUTED           = C_GRAY
C_DISABLED_BG     = C_LIGHT
C_DISABLED_FG     = C_GRAY
C_ACCENT          = C_NAVY
C_ACCENT_SOFT     = C_NAVY
C_ACCENT_BG       = C_LIGHT
C_ACCENT_BORDER   = C_NAVY
C_ERROR           = C_INK
C_BLACK           = C_INK
C_BORDER          = C_LIGHT
C_BORDER_SUBTLE   = C_LIGHT
C_BUTTON_BORDER   = C_LIGHT
C_ON_ACCENT       = C_CHALK

# Semantic button colors
C_BUTTON          = C_GRAY       # Regular / cancel buttons
C_BUTTON_PRIMARY  = C_LIGHT      # Continue / accept buttons

# Dark sections
C_DARK_CARD       = C_NAVY
C_DARK_CARD2      = C_NAVY
C_DARK_TEXT       = C_CHALK
C_DARK_MUTED      = C_LIGHT
C_DARK_BADGE_BG   = C_NAVY
C_DARK_BADGE_BD   = C_NAVY
C_GREEN_DOT       = C_NAVY
C_STATUS_ONLINE   = "#16A34A"   # Green
C_STATUS_IDLE     = "#DC2626"   # Red
C_ALERT_BG        = "#FEE2E2"   # Soft red
C_ALERT_BORDER    = C_STATUS_IDLE
C_WORKED_CARD_BG  = C_NAVY
C_SESSION_BADGE_FG = C_CHALK

C_SPLASH_BG       = C_BG
C_SPLASH_PANEL    = C_SURFACE
C_SPLASH_ACCENT   = C_ACCENT
C_SPLASH_MUTED    = C_MUTED

# Fonts
FONT_FAMILY  = "Space Mono"
F_BASE       = (FONT_FAMILY, 10)
F_TITLE      = (FONT_FAMILY, 14, "bold")
F_BIG_STATUS = (FONT_FAMILY, 20, "bold")
F_MUTED      = (FONT_FAMILY, 10)

# Console logger
LOGGER = logging.getLogger("DeskPulseSuite")
if not LOGGER.handlers:
    _handler = logging.StreamHandler(sys.stdout)
    _handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S"))
    LOGGER.addHandler(_handler)
LOGGER.setLevel(logging.INFO)
LOGGER.propagate = False
F_AGENT_NAME = (FONT_FAMILY, 13, "bold")
F_BOLD       = (FONT_FAMILY, 10, "bold")

# Work modes
WORK_MODES: Dict[str, Dict] = {
    "In Office": {
        "interval_min": 5,
        "screenshots": "none",
        "desc": "Every 5 min, mandatory screenshots at clock in/out",
    },
    "Home Office": {
        "interval_min": 2,
        "screenshots": "inactivity",
        "desc": "Every 2 min, screenshot on inactivity + clock in/out",
    },
    "Training": {
        "interval_min": 10,
        "screenshots": "start_end",
        "desc": "Every 10 min, mandatory screenshots at clock in/out",
    },
}

HOME_OFFICE_IDLE_THRESHOLD_SEC = 120

# Time policy
REGULAR_WORKDAY_SEC = 8 * 60 * 60
OVERTIME_REQUESTERS = ("CLIENT", "TEAM LEADER", "OFFICE OPERATIONS")

ACTIVITY_RULES_DEFAULT = {
    "min_keystrokes": 20,
    "min_clicks":     30,
    "min_scroll":     15,
}

# App tracker exclusion constants
SYSTEM_PROC_PATHS = (
    r"C:\Windows\System32",
    r"C:\Windows\SysWOW64",
    r"C:\Windows\WinSxS",
    r"C:\Windows\SystemApps",
)
SYSTEM_PROC_NAMES = {
    "svchost.exe", "dwm.exe", "csrss.exe", "explorer.exe",
    "taskhostw.exe", "conhost.exe", "lsass.exe", "services.exe",
    "winlogon.exe", "wininit.exe", "smss.exe", "fontdrvhost.exe",
    "sihost.exe", "ctfmon.exe", "runtimebroker.exe",
}
APP_TRACKER_IGNORE_NAME_KEYWORDS = (
    "deskpulse",
)
APP_TRACKER_IGNORE_TITLE_KEYWORDS = (
    "deskpulse",
)


# =============================================================================
# DOMAIN — CONFIG & DATA MODELS
# =============================================================================

def _now_bolivia() -> datetime:
    """Returns current datetime in Bolivia timezone (UTC-4, fixed offset)."""
    return datetime.now(_TZ_BOLIVIA)


def _fmt(dt: Optional[datetime]) -> str:
    """Formats a datetime to ISO-like string for CSV/JSON storage."""
    if dt is None:
        return ""
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def _fmt_date(dt: Optional[datetime]) -> str:
    if dt is None:
        return ""
    return dt.strftime("%Y-%m-%d")


def _fmt_time(dt: Optional[datetime]) -> str:
    if dt is None:
        return ""
    return dt.strftime("%H:%M:%S")


def _extract_time(value: str) -> str:
    raw = (value or "").strip()
    if not raw:
        return ""
    if len(raw) >= 8 and raw[2] == ":" and raw[5] == ":":
        return raw[:8]
    if " " in raw:
        raw = raw.split(" ")[-1]
    return raw[:8]


def _extract_date(value: str) -> str:
    raw = (value or "").strip()
    if not raw:
        return ""
    if len(raw) >= 10 and raw[4] == "-" and raw[7] == "-":
        return raw[:10]
    return ""


def _format_short_session_id(dt: Optional[datetime] = None) -> str:
    return (dt or _now_bolivia()).strftime("%Y%m%d%H%M")


def _has_taken_lunch(session: "SessionData") -> bool:
    return bool(session.lunch_start or session.lunch_end or session.lunch_sec > 0)


def _format_app_list(apps) -> str:
    ordered = []
    seen = set()
    for app in apps or []:
        value = str(app).strip()
        if not value or value in seen:
            continue
        seen.add(value)
        ordered.append(value)
    if not ordered:
        return ""
    return "|" + "|".join(ordered) + "|"


def _sanitize_fs_name(value: str) -> str:
    raw = str(value or "").strip()
    raw = re.sub(r'[<>:"/\\|?*]+', "_", raw)
    raw = re.sub(r"\s+", " ", raw).strip(" .")
    return raw or "Unknown"


def _agent_folder_label(agent_id: str, agent_name: str) -> str:
    safe_id = _sanitize_fs_name(agent_id)
    safe_name = _sanitize_fs_name(agent_name)
    if safe_id and safe_name:
        return f"{safe_id} - {safe_name}"
    return safe_name or safe_id or "Unknown Agent"


def _worksheet_agent_title(agent_name: str, agent_id: str) -> str:
    name = str(agent_name or "").strip()
    code = str(agent_id or "").strip()
    if name and code:
        title = f"{name} ({code})"
    else:
        title = name or code or "DeskPulseSessions"
    title = re.sub(r"[\[\]\*\?/\:\\]", "_", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title[:100] or "DeskPulseSessions"


def _extract_sheet_id(value: str) -> str:
    raw = (value or "").strip()
    if not raw:
        return ""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", raw)
    if match:
        return match.group(1)
    return raw


def _normalize_sheet_header(value: str) -> str:
    raw = str(value or "").strip().lower().replace("-", " ").replace("_", " ")
    return re.sub(r"\s+", " ", raw)


def _bool_to_text(value: Any) -> str:
    return "yes" if bool(value) else "no"


def _internet_reachable(timeout: float = 2.0) -> bool:
    for host, port in (("8.8.8.8", 53), ("1.1.1.1", 53)):
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except OSError:
            continue
    return False


def _capture_connectivity_snapshot() -> Dict[str, Any]:
    network = NetworkDetector.detect()
    connection_type = str(network.get("connection_type", "none") or "none")
    network_name = str(network.get("network_name", "") or "")
    internet_access = False
    if connection_type != "none":
        internet_access = _internet_reachable()
    return {
        "connection_type": connection_type,
        "network_name": network_name,
        "internet_access": internet_access,
    }


def _should_ignore_tracked_app(proc_name: str, exe_path: str, title: str) -> bool:
    name = str(proc_name or "").strip().lower()
    exe = str(exe_path or "").strip().lower()
    ttl = str(title or "").strip().lower()
    for keyword in APP_TRACKER_IGNORE_NAME_KEYWORDS:
        if keyword and (keyword in name or keyword in exe):
            return True
    for keyword in APP_TRACKER_IGNORE_TITLE_KEYWORDS:
        if keyword and keyword in ttl:
            return True
    return False


DEFAULT_CONFIG: Dict[str, Any] = {
    "first_run_done": False,
    "agent": {
        "agent_id":     "",
        "agent_name":   "",
        "project_name": "",
    },
    "work_mode": "In Office",
    "activity_rule": dict(ACTIVITY_RULES_DEFAULT),
    "admin": {
        "username":        "admin",
        "password":        "admin123",
        "password_hashed": False,
        "export_password": "ChangeMe!",
    },
    "google_sheets": {
        "enabled":          False,
        "spreadsheet_link": "",
        "credentials_file": str(CREDENTIALS_FILE),
    },
    "config_import": {
        "source_file": "",
        "imported_at": "",
    },
}


CONFIG_IMPORT_KEY_ALIASES = {
    "agent id": "agent_id",
    "agent_id": "agent_id",
    "agentid": "agent_id",
    "id": "agent_id",
    "agent name": "agent_name",
    "agent_name": "agent_name",
    "agentname": "agent_name",
    "agetn name": "agent_name",
    "agetn_name": "agent_name",
    "project": "project_name",
    "project name": "project_name",
    "project_name": "project_name",
    "proyect": "project_name",
    "proyect_name": "project_name",
    "work mode": "work_mode",
    "work_mode": "work_mode",
    "workmode": "work_mode",
    "mode": "work_mode",
}

WORK_MODE_ALIASES = {
    "in office": "In Office",
    "office": "In Office",
    "in_office": "In Office",
    "home office": "Home Office",
    "homeoffice": "Home Office",
    "home_office": "Home Office",
    "remote": "Home Office",
    "training": "Training",
    "trainning": "Training",
}


def _normalize_import_key(value: str) -> str:
    raw = str(value or "").strip().lower().replace("-", " ").replace("_", " ")
    raw = re.sub(r"\s+", " ", raw)
    return CONFIG_IMPORT_KEY_ALIASES.get(raw, "")


def _normalize_work_mode(value: str) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    raw_key = raw.lower().replace("-", " ").replace("_", " ")
    raw_key = re.sub(r"\s+", " ", raw_key)
    return WORK_MODE_ALIASES.get(raw_key, raw)


def _build_import_profile(data: Dict[str, Any]) -> Dict[str, str]:
    profile = {
        "agent_id": "",
        "agent_name": "",
        "project_name": "",
        "work_mode": "",
    }

    if isinstance(data.get("agent"), dict):
        nested = data.get("agent", {})
        for key, value in nested.items():
            canon = _normalize_import_key(key)
            if canon in ("agent_id", "agent_name", "project_name"):
                profile[canon] = str(value or "").strip()

    for key, value in data.items():
        canon = _normalize_import_key(key)
        if not canon:
            continue
        if canon == "work_mode":
            profile[canon] = _normalize_work_mode(value)
        else:
            profile[canon] = str(value or "").strip()

    profile["work_mode"] = _normalize_work_mode(profile["work_mode"])

    missing = [field for field in ("agent_id", "agent_name", "project_name", "work_mode") if not profile[field]]
    if missing:
        labels = ", ".join(missing)
        raise ValueError(f"Missing required fields: {labels}")
    if profile["work_mode"] not in WORK_MODES:
        raise ValueError(
            "Invalid work_mode. Allowed values: In Office, Home Office, Training."
        )
    return profile


def load_import_settings_file(path: str) -> Dict[str, str]:
    raw = Path(path).read_text(encoding="utf-8-sig")
    stripped = raw.strip()
    if not stripped:
        raise ValueError("The selected file is empty.")

    try:
        data = json.loads(stripped)
        if not isinstance(data, dict):
            raise ValueError("Config file must contain a JSON object.")
        return _build_import_profile(data)
    except json.JSONDecodeError:
        pass

    data: Dict[str, Any] = {}
    for line in raw.splitlines():
        clean = line.strip()
        if not clean or clean.startswith("#") or clean.startswith("//"):
            continue
        if "=" in clean:
            key, value = clean.split("=", 1)
        elif ":" in clean:
            key, value = clean.split(":", 1)
        else:
            continue
        canon = _normalize_import_key(key)
        if not canon:
            continue
        data[canon] = value.strip().strip('"').strip("'")

    if not data:
        raise ValueError(
            "Unsupported config format. Use JSON, or text lines like Agent ID=1234."
        )
    return _build_import_profile(data)


def _config_is_ready(cfg: "AppConfig") -> bool:
    if not cfg.first_run_done:
        return False
    if not str(cfg.agent_id or "").strip():
        return False
    if not str(cfg.agent_name or "").strip():
        return False
    if not str(cfg.project_name or "").strip():
        return False
    if _normalize_work_mode(cfg.work_mode) not in WORK_MODES:
        return False
    return True


@dataclass
class AppConfig:
    first_run_done: bool = False
    agent_id:       str  = ""
    agent_name:     str  = ""
    project_name:   str  = ""
    work_mode:      str  = "In Office"
    activity_rule:  dict = field(default_factory=lambda: dict(ACTIVITY_RULES_DEFAULT))
    admin_username:       str  = "admin"
    admin_password:       str  = "admin123"
    admin_password_hashed: bool = False
    export_password:      str  = "ChangeMe!"
    sheets_enabled:         bool = False
    sheets_spreadsheet_link: str = ""
    sheets_credentials_file: str = str(CREDENTIALS_FILE)
    import_source_file:      str = ""
    import_imported_at:      str = ""


@dataclass
class SessionData:
    session_id:     str
    agent_id:       str
    agent_name:     str
    project:        str
    work_mode:      str
    session_type:   str            # Normal | Schedule Change | Overtime
    clock_in:       Optional[datetime] = None
    clock_out:      Optional[datetime] = None
    lunch_start:    Optional[datetime] = None
    lunch_end:      Optional[datetime] = None
    lunch_sec:      float = 0.0
    meeting_sec:    float = 0.0
    status:         str   = "IDLE"   # IDLE | ONLINE | LUNCH | MEETING
    total_samples:  int   = 0
    active_samples: int   = 0
    inactive_samples_effective: int = 0
    top_app:        str   = ""
    app_used:       str   = ""
    close_reason:   str   = ""
    overtime:       bool  = False
    ot_note:        str   = ""
    overtime_requested_by: str = ""
    payable_worked_sec:    float = 0.0
    overtime_duration_sec: float = 0.0
    clock_in_connection_type: str = ""
    clock_in_network_name: str = ""
    clock_in_internet: bool = False
    clock_in_schedule_lookup: str = ""
    clock_in_lookup_message: str = ""
    clock_in_work_mode_source: str = "config_json"
    clock_out_connection_type: str = ""
    clock_out_network_name: str = ""
    clock_out_internet: bool = False


# =============================================================================
# PERSISTENCE — CONFIG, STATE & CSV FILES
# =============================================================================

class ConfigManager:
    """Load / save config.json; provides typed access via AppConfig."""

    def __init__(self):
        APP_HOME.mkdir(parents=True, exist_ok=True)
        self._raw: Dict[str, Any] = {}
        self.cfg = AppConfig()
        self.load()

    def load(self):
        DPAPIStore.migrate_plain_json_if_needed(
            LEGACY_CONFIG_FILE,
            CONFIG_FILE,
            description="DeskPulseSuite Config"
        )

        if CONFIG_FILE.exists():
            self._raw = DPAPIStore.read_json(CONFIG_FILE)
        else:
            self._raw = json.loads(json.dumps(DEFAULT_CONFIG))
            self._save_raw()

        self._hydrate()

    def _hydrate(self):
        r = self._raw
        self.cfg.first_run_done           = r.get("first_run_done", False)
        ag = r.get("agent", {})
        self.cfg.agent_id                 = ag.get("agent_id", "")
        self.cfg.agent_name               = ag.get("agent_name", "")
        self.cfg.project_name             = ag.get("project_name", "")
        self.cfg.work_mode                = _normalize_work_mode(r.get("work_mode", "In Office")) or "In Office"
        self.cfg.activity_rule            = r.get("activity_rule", dict(ACTIVITY_RULES_DEFAULT))
        adm = r.get("admin", {})
        self.cfg.admin_username           = adm.get("username", "admin")
        self.cfg.admin_password           = adm.get("password", "admin123")
        self.cfg.admin_password_hashed    = adm.get("password_hashed", False)
        self.cfg.export_password          = adm.get("export_password", "ChangeMe!")
        gs = r.get("google_sheets", {})
        self.cfg.sheets_enabled           = gs.get("enabled", False)
        self.cfg.sheets_spreadsheet_link  = gs.get("spreadsheet_link", gs.get("spreadsheet_id", ""))

        saved_cred_file = gs.get("credentials_file", str(CREDENTIALS_FILE))
        try:
            saved_path = Path(saved_cred_file)
            if saved_path.resolve() == LEGACY_CREDENTIALS_FILE.resolve():
                saved_cred_file = str(CREDENTIALS_FILE)
        except Exception:
            if str(saved_cred_file).lower().endswith("credentials.json"):
                saved_cred_file = str(CREDENTIALS_FILE)

        self.cfg.sheets_credentials_file = saved_cred_file

        ci = r.get("config_import", {})
        self.cfg.import_source_file       = ci.get("source_file", "")
        self.cfg.import_imported_at       = ci.get("imported_at", "")

    def save_agent(self, agent_id: str, agent_name: str, project: str):
        self._raw.setdefault("agent", {})
        self._raw["agent"]["agent_id"]     = agent_id
        self._raw["agent"]["agent_name"]   = agent_name
        self._raw["agent"]["project_name"] = project
        self._raw["first_run_done"]         = True
        self._save_raw()
        self._hydrate()

    def save_work_mode(self, mode: str):
        self._raw["work_mode"] = _normalize_work_mode(mode) or "In Office"
        self._save_raw()
        self._hydrate()

    def save_imported_profile(self, profile: Dict[str, str], source_file: str = ""):
        self._raw.setdefault("agent", {})
        self._raw["agent"]["agent_id"]     = profile["agent_id"].strip()
        self._raw["agent"]["agent_name"]   = profile["agent_name"].strip()
        self._raw["agent"]["project_name"] = profile["project_name"].strip()
        self._raw["work_mode"]              = _normalize_work_mode(profile["work_mode"]) or "In Office"
        self._raw["first_run_done"]         = True
        self._raw.setdefault("config_import", {})
        self._raw["config_import"]["source_file"] = source_file or ""
        self._raw["config_import"]["imported_at"] = _now_bolivia().strftime("%Y-%m-%d %H:%M:%S")
        self._save_raw()
        self._hydrate()

    def save_credentials(self, new_user: str, new_pw_hash: str, export_password: str,
                         password_hashed: bool = True):
        self._raw.setdefault("admin", {})
        self._raw["admin"]["username"]        = new_user
        self._raw["admin"]["password"]        = new_pw_hash
        self._raw["admin"]["password_hashed"] = bool(password_hashed)
        self._raw["admin"]["export_password"] = export_password
        self._save_raw()
        self._hydrate()

    def save_google_sheets(self, enabled: bool, spreadsheet_link: str, credentials_file: str):
        self._raw.setdefault("google_sheets", {})
        self._raw["google_sheets"]["enabled"] = bool(enabled)
        self._raw["google_sheets"]["spreadsheet_link"] = spreadsheet_link.strip()
        self._raw["google_sheets"]["credentials_file"] = credentials_file or str(CREDENTIALS_FILE)
        self._save_raw()
        self._hydrate()

    def _save_raw(self):
        APP_HOME.mkdir(parents=True, exist_ok=True)
        SECURE_DIR.mkdir(parents=True, exist_ok=True)
        DPAPIStore.write_json(CONFIG_FILE, self._raw, description="DeskPulseSuite Config")


class StateManager:
    """Reads / writes session_state.json for crash recovery."""

    @staticmethod
    def save(session: SessionData):
        state = {
            "session_id":     session.session_id,
            "agent_id":       session.agent_id,
            "agent_name":     session.agent_name,
            "project":        session.project,
            "work_mode":      session.work_mode,
            "session_type":   session.session_type,
            "status":         session.status,
            "clock_in":       _fmt(session.clock_in),
            "lunch_start":    _fmt(session.lunch_start),
            "lunch_end":      _fmt(session.lunch_end),
            "lunch_sec":      session.lunch_sec,
            "meeting_sec":    session.meeting_sec,
            "total_samples":  session.total_samples,
            "active_samples": session.active_samples,
            "inactive_samples_effective": session.inactive_samples_effective,
            "app_used":       session.app_used,
            "top_app":        session.top_app,
            "clock_in_connection_type": session.clock_in_connection_type,
            "clock_in_network_name": session.clock_in_network_name,
            "clock_in_internet": session.clock_in_internet,
            "clock_in_schedule_lookup": session.clock_in_schedule_lookup,
            "clock_in_lookup_message": session.clock_in_lookup_message,
            "clock_in_work_mode_source": session.clock_in_work_mode_source,
        }
        APP_HOME.mkdir(parents=True, exist_ok=True)
        # Atomic write: write to temp then replace
        temp_file = SESSION_STATE_FILE.with_suffix(".tmp")
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
        temp_file.replace(SESSION_STATE_FILE)

    @staticmethod
    def load() -> Optional[Dict]:
        if SESSION_STATE_FILE.exists():
            with open(SESSION_STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    @staticmethod
    def delete():
        if SESSION_STATE_FILE.exists():
            SESSION_STATE_FILE.unlink()


class SessionLogger:
    """Writes session_log.csv, activity_samples.csv, and summary_day.csv with session totals."""

    @staticmethod
    def _session_dir(session: SessionData) -> Path:
        ci = session.clock_in or _now_bolivia()
        agent_folder = _agent_folder_label(session.agent_id, session.agent_name)
        return (RECORDS_DIR / agent_folder /
                ci.strftime("%Y") / ci.strftime("%m") /
                ci.strftime("%d") / session.session_id)

    @staticmethod
    def _screenshots_dir(session: SessionData) -> Path:
        return SessionLogger._session_dir(session) / "screenshots"

    @staticmethod
    def _ensure_dir(session: SessionData):
        SessionLogger._session_dir(session).mkdir(parents=True, exist_ok=True)
        SessionLogger._screenshots_dir(session).mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _ensure_csv_schema(path: Path, fieldnames: List[str]):
        if not path.exists():
            return True
        try:
            with open(path, "r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                current = reader.fieldnames or []
                if current == fieldnames:
                    return False
                rows = list(reader)
        except Exception:
            rows = []

        tmp = path.with_suffix(path.suffix + ".tmp")
        with open(tmp, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            for row in rows:
                migrated = {k: row.get(k, "") for k in fieldnames}

                ts = row.get("timestamp", "")

                if "sample_date" in migrated and not migrated["sample_date"]:
                    migrated["sample_date"] = row.get("date", "") or _extract_date(ts)

                if "sample_time" in migrated and not migrated["sample_time"]:
                    migrated["sample_time"] = row.get("time", "") or _extract_time(ts)

                if "total_time_worked" in migrated and not migrated["total_time_worked"]:
                    migrated["total_time_worked"] = row.get("net_worked_time", "")
                if "payable_time_worked" in migrated and not migrated["payable_time_worked"]:
                    migrated["payable_time_worked"] = row.get("net_worked_time", "")
                if "overtime_duration" in migrated and not migrated["overtime_duration"]:
                    migrated["overtime_duration"] = "00:00:00"
                if "overtime_requested_by" in migrated and not migrated["overtime_requested_by"]:
                    migrated["overtime_requested_by"] = ""
                w.writerow(migrated)
        tmp.replace(path)
        return False

    @staticmethod
    def append_sample(session: SessionData, row: Dict):
        SessionLogger._ensure_dir(session)
        path = SessionLogger._session_dir(session) / "activity_samples.csv"
        fieldnames = [
            "session_id",
            "agent_name",
            "sample_date",
            "sample_time",
            "status",
            "keystrokes",
            "clicks",
            "scroll_events",
            "app_used",
            "connection_type",
            "network_name",
            "activity_flag",
            "screenshot_taken",
        ]
        write_header = SessionLogger._ensure_csv_schema(path, fieldnames)
        with open(path, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            if write_header:
                w.writeheader()
            w.writerow(row)

    @staticmethod
    def write_session_log(session: SessionData, lunch_time_fmt: str,
                          meeting_time_fmt: str, net_worked_fmt: str,
                          payable_worked_fmt: str, overtime_duration_fmt: str):
        SessionLogger._ensure_dir(session)
        path = SessionLogger._session_dir(session) / "session_log.csv"
        fieldnames = [
            "session_id", "date", "agent_id", "agent_name", "project",
            "work_mode", "session_type",
            "clock_in", "clock_out",
            "clock_in_connection_type", "clock_in_network_name", "clock_in_internet",
            "clock_in_schedule_lookup", "clock_in_lookup_message", "clock_in_work_mode_source",
            "clock_out_connection_type", "clock_out_network_name", "clock_out_internet",
            "lunch_start", "lunch_end",
            "lunch_time", "meeting_time", "net_worked_time",
            "total_time_worked", "payable_time_worked", "overtime_duration",
            "overtime_requested_by",
            "overtime", "ot_note", "top_app", "app_used",
            "activity_avg", "close_reason",
        ]
        write_header = SessionLogger._ensure_csv_schema(path, fieldnames)
        with open(path, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            if write_header:
                w.writeheader()
            total = session.total_samples or 1
            w.writerow({
                "session_id":      session.session_id,
                "date":            _fmt_date(session.clock_in),
                "agent_id":        session.agent_id,
                "agent_name":      session.agent_name,
                "project":         session.project,
                "work_mode":       session.work_mode,
                "session_type":    session.session_type,
                "clock_in":        _fmt_time(session.clock_in),
                "clock_out":       _fmt_time(session.clock_out),
                "clock_in_connection_type": session.clock_in_connection_type,
                "clock_in_network_name": session.clock_in_network_name,
                "clock_in_internet": _bool_to_text(session.clock_in_internet),
                "clock_in_schedule_lookup": session.clock_in_schedule_lookup,
                "clock_in_lookup_message": session.clock_in_lookup_message,
                "clock_in_work_mode_source": session.clock_in_work_mode_source,
                "clock_out_connection_type": session.clock_out_connection_type,
                "clock_out_network_name": session.clock_out_network_name,
                "clock_out_internet": _bool_to_text(session.clock_out_internet),
                "lunch_start":     _fmt_time(session.lunch_start),
                "lunch_end":       _fmt_time(session.lunch_end),
                "lunch_time":      lunch_time_fmt,
                "meeting_time":    meeting_time_fmt,
                "net_worked_time": net_worked_fmt,
                "total_time_worked": net_worked_fmt,
                "payable_time_worked": payable_worked_fmt,
                "overtime_duration": overtime_duration_fmt,
                "overtime_requested_by": session.overtime_requested_by,
                "overtime":        session.overtime,
                "ot_note":         session.ot_note,
                "top_app":         session.top_app,
                "app_used":        session.app_used,
                "activity_avg":    f"{session.active_samples / total * 100:.1f}",
                "close_reason":    session.close_reason,
            })

    @staticmethod
    def write_summary_day(session: SessionData, lunch_time_fmt: str,
                          meeting_time_fmt: str, net_worked_fmt: str,
                          payable_worked_fmt: str, overtime_duration_fmt: str):
        ci = session.clock_in or _now_bolivia()
        agent_folder = _agent_folder_label(session.agent_id, session.agent_name)
        summary_path = (RECORDS_DIR / agent_folder /
                        ci.strftime("%Y") / ci.strftime("%m") /
                        ci.strftime("%d") / "summary_day.csv")
        fieldnames = [
            "session_id", "agent_id", "agent_name", "project", "work_mode", "date",
            "clock_in", "clock_out",
            "clock_in_connection_type", "clock_in_network_name", "clock_in_internet",
            "clock_in_schedule_lookup", "clock_in_lookup_message", "clock_in_work_mode_source",
            "clock_out_connection_type", "clock_out_network_name", "clock_out_internet",
            "total_lunch_time", "total_meeting_time", "net_worked_time",
            "total_time_worked", "payable_time_worked", "overtime_duration",
            "overtime_requested_by",
            "overtime", "top_app", "app_used", "activity_avg",
            "total_samples", "active_samples", "inactive_samples",
        ]
        write_header = SessionLogger._ensure_csv_schema(summary_path, fieldnames)
        total = session.total_samples or 1
        with open(summary_path, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            if write_header:
                w.writeheader()
            w.writerow({
                "session_id":         session.session_id,
                "agent_id":           session.agent_id,
                "agent_name":         session.agent_name,
                "project":            session.project,
                "work_mode":          session.work_mode,
                "date":               ci.strftime("%Y-%m-%d"),
                "clock_in":           _fmt_time(session.clock_in),
                "clock_out":          _fmt_time(session.clock_out),
                "clock_in_connection_type": session.clock_in_connection_type,
                "clock_in_network_name": session.clock_in_network_name,
                "clock_in_internet": _bool_to_text(session.clock_in_internet),
                "clock_in_schedule_lookup": session.clock_in_schedule_lookup,
                "clock_in_lookup_message": session.clock_in_lookup_message,
                "clock_in_work_mode_source": session.clock_in_work_mode_source,
                "clock_out_connection_type": session.clock_out_connection_type,
                "clock_out_network_name": session.clock_out_network_name,
                "clock_out_internet": _bool_to_text(session.clock_out_internet),
                "total_lunch_time":   lunch_time_fmt,
                "total_meeting_time": meeting_time_fmt,
                "net_worked_time":    net_worked_fmt,
                "total_time_worked":  net_worked_fmt,
                "payable_time_worked": payable_worked_fmt,
                "overtime_duration":  overtime_duration_fmt,
                "overtime_requested_by": session.overtime_requested_by,
                "overtime":           session.overtime,
                "top_app":            session.top_app,
                "app_used":           session.app_used,
                "activity_avg":       f"{session.active_samples / total * 100:.1f}",
                "total_samples":      session.total_samples,
                "active_samples":     session.active_samples,
                "inactive_samples":   session.inactive_samples_effective,
            })
        return summary_path


# =============================================================================
# SECURITY — BCRYPT & AES-256-GCM EXPORT ENCRYPTION
# =============================================================================

class SecurityUtils:

    @staticmethod
    def hash_password(plain: str) -> str:
        if bcrypt is None:
            raise RuntimeError("bcrypt not installed")
        return bcrypt.hashpw(plain.encode(), bcrypt.gensalt()).decode()

    @staticmethod
    def verify_password(plain: str, hashed_or_plain: str, is_hashed: bool) -> bool:
        if is_hashed:
            if bcrypt is None:
                return False
            try:
                return bcrypt.checkpw(plain.encode(), hashed_or_plain.encode())
            except Exception:
                return False
        return plain == hashed_or_plain

    @staticmethod
    def encrypt_file(src_path: Path, dest_path: Path, password: str):
        """AES-256-GCM with PBKDF2-SHA256 (260 000 iterations).
        File format: salt(16) + nonce(12) + ciphertext+tag"""
        if not _CRYPTO_OK:
            raise RuntimeError("cryptography package not installed")
        salt  = os.urandom(16)
        nonce = os.urandom(12)
        kdf = PBKDF2HMAC(
            algorithm=crypto_hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=260_000,
            backend=default_backend(),
        )
        key  = kdf.derive(password.encode())
        data = src_path.read_bytes()
        ct   = AESGCM(key).encrypt(nonce, data, None)
        dest_path.write_bytes(salt + nonce + ct)

class DPAPIStore:
    """Protects local files with Windows DPAPI (CurrentUser scope)."""

    @staticmethod
    def _require():
        if not _DPAPI_OK:
            raise RuntimeError("pywin32 is required for DPAPI support (win32crypt).")

    @staticmethod
    def protect_bytes(data: bytes, description: str = "DeskPulseSuite") -> bytes:
        DPAPIStore._require()
        if not isinstance(data, (bytes, bytearray)):
            raise TypeError("data must be bytes")
        return win32crypt.CryptProtectData(
            bytes(data),
            description,
            DPAPI_ENTROPY,
            None,
            None,
            0,   # CurrentUser scope; do NOT use CRYPTPROTECT_LOCAL_MACHINE here
        )

    @staticmethod
    def unprotect_bytes(blob: bytes) -> bytes:
        DPAPIStore._require()
        if not isinstance(blob, (bytes, bytearray)):
            raise TypeError("blob must be bytes")
        _desc, data = win32crypt.CryptUnprotectData(
            bytes(blob),
            DPAPI_ENTROPY,
            None,
            None,
            0,
        )
        return data

    @staticmethod
    def write_json(path: Path, payload: Dict[str, Any], description: str = "DeskPulseSuite JSON"):
        raw = json.dumps(payload, indent=2).encode("utf-8")
        enc = DPAPIStore.protect_bytes(raw, description=description)
        path.parent.mkdir(parents=True, exist_ok=True)

        tmp = path.with_suffix(path.suffix + ".tmp")
        with open(tmp, "wb") as f:
            f.write(enc)
        tmp.replace(path)

    @staticmethod
    def read_json(path: Path) -> Dict[str, Any]:
        blob = path.read_bytes()
        raw = DPAPIStore.unprotect_bytes(blob)
        return json.loads(raw.decode("utf-8"))

    @staticmethod
    def migrate_plain_json_if_needed(legacy_path: Path, encrypted_path: Path, description: str):
        if encrypted_path.exists():
            return
        if not legacy_path.exists():
            return

        data = json.loads(legacy_path.read_text(encoding="utf-8"))
        DPAPIStore.write_json(encrypted_path, data, description=description)

        backup = legacy_path.with_suffix(legacy_path.suffix + ".migrated.bak")
        try:
            legacy_path.replace(backup)
        except Exception:
            try:
                legacy_path.unlink()
            except Exception:
                pass

# =============================================================================
# MONITORING — INPUT MONITOR
# =============================================================================

class InputMonitor:
    """Tracks keystrokes, clicks, scroll events, and mouse coverage.
    All counters are thread-safe and reset on snapshot()."""

    def __init__(self):
        self._lock       = threading.Lock()
        self._keystrokes = 0
        self._clicks     = 0
        self._scrolls    = 0
        self._kb_listener  = None
        self._ms_listener  = None

    def start(self):
        if not _PYNPUT_OK:
            return
        self._kb_listener = _pk.Listener(on_press=self._on_key)
        self._ms_listener = _pm.Listener(
            on_click=self._on_click,
            on_scroll=self._on_scroll,
        )
        self._kb_listener.daemon = True
        self._ms_listener.daemon = True
        self._kb_listener.start()
        self._ms_listener.start()

    def stop(self):
        if self._kb_listener:
            self._kb_listener.stop()
        if self._ms_listener:
            self._ms_listener.stop()

    def _on_key(self, key):
        with self._lock:
            self._keystrokes += 1

    def _on_click(self, x, y, button, pressed):
        if pressed:
            with self._lock:
                self._clicks += 1

    def _on_scroll(self, x, y, dx, dy):
        with self._lock:
            self._scrolls += 1


    def snapshot(self) -> Dict:
        with self._lock:
            ks = self._keystrokes
            cl = self._clicks
            sc = self._scrolls
            self._keystrokes = 0
            self._clicks     = 0
            self._scrolls    = 0
        return {
            "keystrokes":    ks,
            "clicks":        cl,
            "scroll_events": sc,
        }


# =============================================================================
# MONITORING — IDLE DETECTOR
# =============================================================================

# Windows idle tracking removed by business rule.
# This application now behaves strictly as a periodic sample logger.


# =============================================================================
# MONITORING — APP TRACKER
# =============================================================================

class AppTracker:
    """Tracks the foreground app/window and preserves the last valid non-DeskPulse app."""

    _last_valid_app: str = ""

    @classmethod
    def _fallback(cls) -> List[str]:
        return [cls._last_valid_app] if cls._last_valid_app else []

    @classmethod
    def get_visible_apps(cls) -> List[str]:
        if not (_WIN32_OK and _PSUTIL_OK):
            return cls._fallback()

        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return cls._fallback()

            if not win32gui.IsWindowVisible(hwnd):
                return cls._fallback()

            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if not title:
                return cls._fallback()

            # Exclude tool windows
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            if ex_style & win32con.WS_EX_TOOLWINDOW:
                return cls._fallback()

            # Exclude owned windows (popups)
            if win32gui.GetWindow(hwnd, win32con.GW_OWNER):
                return cls._fallback()

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)

            exe = (proc.exe() or "").strip()
            name = (proc.name() or "").strip()
            name_lower = name.lower()
            exe_lower = exe.lower()

            # Hard exclusions for Windows/system stuff
            extra_system_names = {
                "systemsettings.exe",
                "applicationframehost.exe",
                "searchhost.exe",
                "shellexperiencehost.exe",
                "startmenuexperiencehost.exe",
                "textinputhost.exe",
                "lockapp.exe",
                "widgets.exe",
            }
            extra_system_paths = (
                r"c:\windows\system32",
                r"c:\windows\syswow64",
                r"c:\windows\winsxs",
                r"c:\windows\systemapps",
                r"c:\windows\immersivecontrolpanel",
            )

            if name_lower in SYSTEM_PROC_NAMES or name_lower in extra_system_names:
                return cls._fallback()

            if any(p.lower() in exe_lower for p in SYSTEM_PROC_PATHS):
                return cls._fallback()

            if any(p in exe_lower for p in extra_system_paths):
                return cls._fallback()

            if _should_ignore_tracked_app(name, exe, title):
                return cls._fallback()

            app_label = f"{name} - {title}"
            cls._last_valid_app = app_label
            return [app_label]

        except Exception:
            return cls._fallback()


# =============================================================================
# MONITORING — NETWORK DETECTOR
# =============================================================================

class NetworkDetector:
    """Detects connection type (wifi/ethernet/none) and network name."""

    @staticmethod
    def detect() -> Dict[str, str]:
        if not _PSUTIL_OK:
            return {"connection_type": "none", "network_name": ""}
        try:
            stats   = psutil.net_if_stats()
            addrs   = psutil.net_if_addrs()
            wifi_if = None
            eth_if  = None

            for iface, st in stats.items():
                if not st.isup:
                    continue
                iface_lower = iface.lower()
                # Skip loopback and known virtual adapters
                if any(x in iface_lower for x in ("loopback", "lo", "virtual",
                                                    "vmware", "vbox", "hyper")):
                    continue
                ips = [a.address for a in addrs.get(iface, [])
                       if a.family == 2 and  # AF_INET
                       not a.address.startswith("169.254") and
                       not a.address.startswith("127.")]
                if not ips:
                    continue
                if any(x in iface_lower for x in ("wi-fi", "wifi", "wireless", "wlan")):
                    if wifi_if is None:
                        wifi_if = iface
                elif eth_if is None:
                    eth_if = iface

            # Business rule: if both Ethernet and Wi-Fi are active, Ethernet wins.
            if eth_if:
                return {"connection_type": "ethernet", "network_name": eth_if}
            if wifi_if:
                ssid = NetworkDetector._get_ssid(wifi_if)
                return {"connection_type": "wifi", "network_name": ssid}
        except Exception:
            pass
        return {"connection_type": "none", "network_name": ""}

    @staticmethod
    def _get_ssid(iface: str) -> str:
        try:
            result = subprocess.run(
                ["netsh", "wlan", "show", "interfaces"],
                capture_output=True, text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            for line in result.stdout.splitlines():
                if "SSID" in line and "BSSID" not in line:
                    return line.split(":", 1)[-1].strip()
        except Exception:
            pass
        return ""


# =============================================================================
# MONITORING — SCREENSHOT CAPTURE
# =============================================================================

class ScreenshotCapture:

    @staticmethod
    def take(session: SessionData, reason: str) -> Optional[str]:
        if not (_MSS_OK and _PIL_OK):
            return None
        try:
            shot_dir = SessionLogger._screenshots_dir(session)
            shot_dir.mkdir(parents=True, exist_ok=True)
            ts   = _now_bolivia().strftime("%Y%m%d_%H%M%S")
            name = f"{ts}_{reason}.png"
            path = shot_dir / name

            with mss.mss() as sct:
                monitors = sct.monitors[1:]  # skip the "all" monitor
                if len(monitors) == 1:
                    img = sct.grab(monitors[0])
                    Image.frombytes("RGB", img.size, img.bgra, "raw", "BGRX").save(str(path))
                else:
                    images = [Image.frombytes("RGB", sct.grab(m).size,
                                               sct.grab(m).bgra, "raw", "BGRX")
                              for m in monitors]
                    total_w = sum(i.width for i in images)
                    max_h   = max(i.height for i in images)
                    combined = Image.new("RGB", (total_w, max_h))
                    x_off = 0
                    for img in images:
                        combined.paste(img, (x_off, 0))
                        x_off += img.width
                    combined.save(str(path))
            return str(path)
        except Exception:
            return None


# =============================================================================
# MONITORING — SAMPLE WORKER
# =============================================================================

class SampleWorker:
    """Background timer that fires every N minutes per work mode.
    Collects input, apps, network; evaluates activity; writes CSV row;
    updates session state file for recovery."""

    def __init__(self, session: SessionData, cfg: AppConfig,
                 on_tick_callback=None):
        self.session          = session
        self.cfg              = cfg
        self.on_tick          = on_tick_callback
        self._timer: Optional[threading.Timer] = None
        self._stopped         = False
        self._input           = InputMonitor()
        self._app_counts: Dict[str, int] = {}
        self._last_sample_at: Optional[datetime] = None
        self._meeting_start = None   # Initialize meeting start tracker

    def start(self):
        self._last_sample_at = _now_bolivia()
        self._input.start()
        self._schedule()

    def stop(self):
        self._stopped = True
        if self._timer:
            self._timer.cancel()
        self._input.stop()

    def _interval_seconds(self) -> int:
        return WORK_MODES[self.session.work_mode]["interval_min"] * 60

    def _backfill_missed_samples(self, now: datetime):
        if not self._last_sample_at:
            self._last_sample_at = now
            return

        interval_sec = self._interval_seconds()
        elapsed_sec = max(0.0, (now - self._last_sample_at).total_seconds())
        missed = int(elapsed_sec // interval_sec) - 1
        if missed <= 0:
            return

        # Limit backfill to prevent massive log generation after long suspension
        MAX_BACKFILL = 10
        if missed > MAX_BACKFILL:
            LOGGER.warning(
                "Missed %d samples, limiting backfill to %d",
                missed, MAX_BACKFILL
            )
            missed = MAX_BACKFILL

        for idx in range(1, missed + 1):
            sample_time = self._last_sample_at + timedelta(seconds=interval_sec * idx)
            row = {
                "session_id":       self.session.session_id,
                "agent_name":       self.session.agent_name,
                "sample_date":      _fmt_date(sample_time),
                "sample_time":      _fmt_time(sample_time),
                "status":           self.session.status,
                "keystrokes":       0,
                "clicks":           0,
                "scroll_events":    0,
                "app_used":         "",
                "connection_type":  "none",
                "network_name":     "",
                "activity_flag":    "INACTIVE",
                "screenshot_taken": "no",
            }
            SessionLogger.append_sample(self.session, row)
            self.session.total_samples += 1
            LOGGER.info(
                "SAMPLE-BACKFILL | agent=%s | session=%s | status=%s | activity=INACTIVE | ts=%s",
                self.session.agent_id,
                self.session.session_id,
                self.session.status,
                _fmt(sample_time),
            )

    def _schedule(self):
        if self._stopped:
            return
        interval_sec = self._interval_seconds()
        self._timer = threading.Timer(interval_sec, self.tick)
        self._timer.daemon = True
        self._timer.start()

    def tick(self, reason: str = "scheduled"):
        if self._stopped:
            return
        try:
            self._collect_and_write(reason)
        finally:
            self._schedule()

    def _collect_and_write(self, reason: str):
        mode       = self.session.work_mode
        status     = self.session.status
        snap       = self._input.snapshot()
        rules      = self.cfg.activity_rule
        now        = _now_bolivia()
        self._backfill_missed_samples(now)

        # Determine what to capture per mode
        capture_scroll = (mode in ("In Office", "Home Office"))

        apps       = AppTracker.get_visible_apps()
        net_info   = NetworkDetector.detect()

        if not capture_scroll:
            snap["scroll_events"] = 0

        # Accumulate top_app counts
        for app in apps:
            self._app_counts[app] = self._app_counts.get(app, 0) + 1

        # Activity flag
        if status == "MEETING":
            activity_flag = "ACTIVE"  # forced per spec
        else:
            min_keystrokes = int(rules.get("min_keystrokes", 20) or 20)
            min_clicks = int(rules.get("min_clicks", 30) or 30)
            min_scroll = int(rules.get("min_scroll", 15) or 15)
            active = (
                snap["keystrokes"] >= min_keystrokes or
                snap["clicks"] >= min_clicks or
                snap["scroll_events"] > min_scroll
            )
            
            activity_flag = "ACTIVE" if active else "INACTIVE"

        # Screenshot logic
        screenshot_taken = "no"
        screenshot_path  = ""
        mode_ss = WORK_MODES[mode]["screenshots"]

        if mode_ss == "inactivity":
            if activity_flag == "INACTIVE":
                p = ScreenshotCapture.take(self.session, reason)
                if p:
                    screenshot_taken = "yes"
                    screenshot_path  = p
        # "start_end" screenshots are handled explicitly in clockin/clockout

        # Update session totals
        self.session.total_samples += 1
        if activity_flag == "ACTIVE":
            self.session.active_samples += 1
        elif status != "LUNCH":
            self.session.inactive_samples_effective += 1

        # Update top_app / app_used on session
        if self._app_counts:
            self.session.top_app = max(self._app_counts, key=self._app_counts.get)
            self.session.app_used = _format_app_list(self._app_counts.keys())

        # Write activity sample row
        row = {
            "session_id":       self.session.session_id,
            "agent_name":       self.session.agent_name,
            "sample_date":      _fmt_date(now),
            "sample_time":      _fmt_time(now),
            "status":           status,
            "keystrokes":       snap["keystrokes"],
            "clicks":           snap["clicks"],
            "scroll_events":    snap["scroll_events"],
            "app_used":         _format_app_list(apps),
            "connection_type":  net_info["connection_type"],
            "network_name":     net_info["network_name"],
            "activity_flag":    activity_flag,
            "screenshot_taken": screenshot_taken,
        }
        SessionLogger.append_sample(self.session, row)
        LOGGER.info(
            "SAMPLE | agent=%s | session=%s | status=%s | activity=%s | keys=%s | clicks=%s | scroll=%s | apps=%s | screenshot=%s",
            self.session.agent_id,
            self.session.session_id,
            status,
            activity_flag,
            snap["keystrokes"],
            snap["clicks"],
            snap["scroll_events"],
            row["app_used"] or "|",
            screenshot_taken,
        )

        # Persist state for recovery
        StateManager.save(self.session)
        self._last_sample_at = now

        if self.on_tick:
            self.on_tick()

    def set_status(self, new_status: str, now: Optional[datetime] = None) -> bool:
        """Called when status changes. Accumulates closed block seconds and blocks invalid transitions."""
        now = now or _now_bolivia()
        old = self.session.status

        allowed = {
            "ONLINE": {"ONLINE", "LUNCH", "MEETING"},
            "LUNCH": {"LUNCH", "ONLINE"},
            "MEETING": {"MEETING", "ONLINE"},
            "IDLE": {"ONLINE", "IDLE"},
        }
        if new_status not in allowed.get(old, {old}):
            LOGGER.warning("Invalid status transition blocked: %s -> %s", old, new_status)
            return False

        if old == "LUNCH" and self.session.lunch_start:
            self.session.lunch_sec += (now - self.session.lunch_start).total_seconds()
            self.session.lunch_end = now
        if old == "MEETING":
            if self._meeting_start:
                self.session.meeting_sec += (now - self._meeting_start).total_seconds()
                self._meeting_start = None

        self.session.status = new_status

        if new_status == "LUNCH":
            self.session.lunch_start = now
            self.session.lunch_end = None
        if new_status == "MEETING":
            self._meeting_start = now

        StateManager.save(self.session)
        return True


# =============================================================================
# GOOGLE SHEETS SYNC
# =============================================================================

class SheetsWorksheetNotFoundError(Exception):
    """Raised when the expected agent worksheet does not exist."""


class SheetsSync:
    """Sends session rows to Google Sheets. Queues on failure, retries."""

    SCHEDULE_WORKSHEET = "AGENTS_SCHEDULE"

    SCOPES = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    HEADERS = [
        "session_id", "date", "project_name", "agent_name", "session_type",
        "clock_in", "clock_out", "lunch_start", "lunch_end", "status", "net_worked_time",
    ]

    def __init__(self, cfg: AppConfig):
        self.cfg         = cfg
        self._queue: List[Dict] = []
        self._thread     = None
        self._lock       = threading.Lock()
        self._gc         = None
        self.last_error   = ""
        self._load_queue()
        self._start_retry_thread()

    def _load_queue(self):
        if SHEETS_QUEUE_FILE.exists():
            with open(SHEETS_QUEUE_FILE, "r", encoding="utf-8") as f:
                self._queue = json.load(f)

    def _save_queue(self):
        APP_HOME.mkdir(parents=True, exist_ok=True)
        with open(SHEETS_QUEUE_FILE, "w", encoding="utf-8") as f:
            json.dump(self._queue, f, indent=2)

    def refresh_config(self, cfg: AppConfig):
        self.cfg = cfg
        self._gc = None

    def _get_client(self):
        if not (_GSPREAD_OK and self.cfg.sheets_enabled):
            return None
        if self._gc:
            return self._gc

        try:
            raw_path = (self.cfg.sheets_credentials_file or "").strip()
            legacy_path = LEGACY_CREDENTIALS_FILE
            secure_path = CREDENTIALS_FILE

            if not raw_path:
                cred_path = secure_path
            else:
                cred_path = Path(raw_path)
                try:
                    if cred_path.resolve() == legacy_path.resolve():
                        cred_path = secure_path
                except Exception:
                    if str(cred_path).lower().endswith("credentials.json"):
                        cred_path = secure_path

            DPAPIStore.migrate_plain_json_if_needed(
                legacy_path,
                cred_path,
                description="DeskPulseSuite Google Credentials"
            )

            if not cred_path.exists():
                raise FileNotFoundError(f"Encrypted credentials file not found: {cred_path}")

            cred_info = DPAPIStore.read_json(cred_path)
            creds = _GSACredentials.from_service_account_info(
                cred_info,
                scopes=self.SCOPES
            )
            self._gc = gspread.authorize(creds)
            return self._gc
        except Exception as exc:
            LOGGER.exception("Google Sheets client init failed: %s", exc)
            return None

    def _spreadsheet(self, gc):
        spreadsheet_id = _extract_sheet_id(self.cfg.sheets_spreadsheet_link)
        if not spreadsheet_id:
            return None
        return gc.open_by_key(spreadsheet_id)

    def _worksheet(self, gc):
        sh = self._spreadsheet(gc)
        if sh is None:
            raise RuntimeError("Google Sheets spreadsheet is not configured.")
        title = _worksheet_agent_title(self.cfg.agent_name, self.cfg.agent_id)
        try:
            ws = sh.worksheet(title)
        except gspread.WorksheetNotFound as exc:
            raise SheetsWorksheetNotFoundError(
                f"ERROR USER NOT FOUND\n\nWorksheet '{title}' was not found in Google Sheets."
            ) from exc
        try:
            ws.resize(rows=max(ws.row_count, 1000), cols=len(self.HEADERS))
        except Exception:
            pass
        try:
            current_header = ws.row_values(1)
        except Exception:
            current_header = []
        if current_header != self.HEADERS:
            ws.update(range_name="A1", values=[self.HEADERS])
        return ws

    def lookup_agent_schedule(self, agent_id: str) -> Dict[str, str]:
        agent_id = str(agent_id or "").strip()
        if not agent_id:
            return {
                "status": "missing_agent_id",
                "message": "Agent ID is empty.",
                "work_mode": "",
            }
        if not self.cfg.sheets_enabled:
            return {
                "status": "disabled",
                "message": "Google Sheets sync is disabled.",
                "work_mode": "",
            }
        if not _extract_sheet_id(self.cfg.sheets_spreadsheet_link):
            return {
                "status": "spreadsheet_not_configured",
                "message": "Google Sheets spreadsheet is not configured.",
                "work_mode": "",
            }
        gc = self._get_client()
        if gc is None:
            return {
                "status": "client_unavailable",
                "message": "Google Sheets client could not be initialized.",
                "work_mode": "",
            }
        try:
            sh = self._spreadsheet(gc)
            if sh is None:
                return {
                    "status": "spreadsheet_not_configured",
                    "message": "Google Sheets spreadsheet is not configured.",
                    "work_mode": "",
                }
            try:
                ws = sh.worksheet(self.SCHEDULE_WORKSHEET)
            except gspread.WorksheetNotFound:
                return {
                    "status": "worksheet_not_found",
                    "message": f"Worksheet '{self.SCHEDULE_WORKSHEET}' was not found.",
                    "work_mode": "",
                }

            values = ws.get_all_values()
            if not values:
                return {
                    "status": "worksheet_empty",
                    "message": f"Worksheet '{self.SCHEDULE_WORKSHEET}' is empty.",
                    "work_mode": "",
                }

            headers = values[0]
            header_map = {_normalize_sheet_header(name): idx for idx, name in enumerate(headers)}
            agent_idx = header_map.get("agent id")
            work_mode_idx = header_map.get("work mode")

            if agent_idx is None or work_mode_idx is None:
                return {
                    "status": "invalid_headers",
                    "message": "AGENTS_SCHEDULE must include AGENT_ID and WORK MODE columns.",
                    "work_mode": "",
                }

            agent_id_cmp = agent_id.strip().upper()
            for row in values[1:]:
                row_agent_id = (row[agent_idx].strip() if agent_idx < len(row) else "")
                if row_agent_id.upper() != agent_id_cmp:
                    continue

                raw_work_mode = row[work_mode_idx].strip() if work_mode_idx < len(row) else ""
                if not raw_work_mode:
                    return {
                        "status": "empty_work_mode",
                        "message": f"AGENT_ID {agent_id} has an empty WORK MODE in AGENTS_SCHEDULE.",
                        "work_mode": "",
                    }

                normalized_mode = _normalize_work_mode(raw_work_mode)
                if normalized_mode not in WORK_MODES:
                    return {
                        "status": "invalid_work_mode",
                        "message": f"AGENT_ID {agent_id} has invalid WORK MODE '{raw_work_mode}'.",
                        "work_mode": "",
                    }

                return {
                    "status": "ok",
                    "message": f"AGENT_ID {agent_id} resolved from AGENTS_SCHEDULE.",
                    "work_mode": normalized_mode,
                }

            return {
                "status": "agent_not_found",
                "message": f"AGENT_ID {agent_id} was not found in AGENTS_SCHEDULE.",
                "work_mode": "",
            }
        except Exception as exc:
            return {
                "status": "error",
                "message": str(exc),
                "work_mode": "",
            }

    def _entry_from_session(self, event_type: str, session: SessionData) -> Dict[str, Any]:
        del event_type
        net_sec = 0.0
        if session.clock_in:
            end_dt = session.clock_out or _now_bolivia()
            net_sec = max(0.0, (end_dt - session.clock_in).total_seconds() - session.lunch_sec)
        return {
            "session_id": session.session_id,
            "date": _fmt_date(session.clock_in or _now_bolivia()),
            "project_name": session.project,
            "agent_name": session.agent_name,
            "session_type": session.session_type,
            "clock_in": _fmt_time(session.clock_in),
            "clock_out": _fmt_time(session.clock_out),
            "lunch_start": _fmt_time(session.lunch_start),
            "lunch_end": _fmt_time(session.lunch_end),
            "status": session.status,
            "net_worked_time": _sec_to_hms(net_sec),
            "_retries": 0,
        }

    def send_event(self, event_type: str, session: SessionData) -> bool:
        if not self.cfg.sheets_enabled:
            return True
        entry = self._entry_from_session(event_type, session)
        success, retryable, error_message = self._push(entry)
        self.last_error = error_message
        if not success and retryable:
            with self._lock:
                self._queue.append(entry)
                self._save_queue()
        return success

    def _push(self, entry: Dict) -> tuple[bool, bool, str]:
        try:
            gc = self._get_client()
            if gc is None:
                return False, True, "Google Sheets client unavailable."
            ws = self._worksheet(gc)

            values = ws.get_all_values()
            row_idx = None
            for idx, row in enumerate(values[1:], start=2):
                row_session_id = row[0] if len(row) > 0 else ""
                if row_session_id == entry["session_id"]:
                    row_idx = idx
                    break

            row_data = [entry.get(k, "") for k in self.HEADERS]
            if row_idx:
                ws.update(range_name=f"A{row_idx}", values=[row_data])
            else:
                ws.append_row(row_data)
            return True, False, ""
        except SheetsWorksheetNotFoundError as exc:
            return False, False, str(exc)
        except Exception as exc:
            return False, True, str(exc)

    def _start_retry_thread(self):
        self._thread = threading.Thread(target=self._retry_loop, daemon=True)
        self._thread.start()

    def _retry_loop(self):
        while True:
            time.sleep(SHEETS_RETRY_INTERVAL)
            with self._lock:
                remaining = []
                for entry in self._queue:
                    if entry.get("_retries", 0) >= SHEETS_MAX_RETRIES:
                        continue
                    success, retryable, error_message = self._push(entry)
                    if not success:
                        self.last_error = error_message
                        if retryable:
                            entry["_retries"] = entry.get("_retries", 0) + 1
                            remaining.append(entry)
                self._queue = remaining
                self._save_queue()


# =============================================================================
# UI — THEME & HELPERS
# =============================================================================

def _enable_dpi_awareness():
    """Enables Per-Monitor DPI awareness on Windows."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def _center_window(win: tk.Toplevel | tk.Tk, w: int, h: int):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x  = (sw - w) // 2
    y  = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")


def _apply_windows_app_id():
    """Helps Windows taskbar show the bundled app icon consistently."""
    if os.name != "nt":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


def _apply_window_icon(win: tk.Toplevel | tk.Tk):
    """Applies the portable app icon to any window if the asset exists."""
    try:
        if APP_ICON_FILE.exists() and APP_ICON_FILE.suffix.lower() == ".ico":
            win.iconbitmap(str(APP_ICON_FILE))
    except Exception:
        pass


def _create_dialog(parent: tk.Toplevel | tk.Tk, title: str, w: int, h: int, *, topmost: bool = False) -> tk.Toplevel:
    """Create dialogs hidden first to reduce visible flicker on transient windows."""
    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    _apply_window_icon(dlg)
    dlg.title(title)
    dlg.resizable(False, False)
    dlg.transient(parent)
    if topmost:
        try:
            dlg.attributes("-topmost", True)
        except Exception:
            pass
    dlg.configure(bg=C_BG)
    _center_window(dlg, w, h)
    return dlg


def _show_dialog(dlg: tk.Toplevel):
    dlg.update_idletasks()
    dlg.deiconify()
    try:
        dlg.lift()
        dlg.focus_force()
    except Exception:
        pass
    dlg.grab_set()


def _apply_theme(root: tk.Tk):
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", font=F_BASE, background=C_BG, foreground=C_TEXT)

    style.configure("Title.TLabel",
                     font=F_TITLE, background=C_BG, foreground=C_TEXT)
    style.configure("BigStatus.TLabel",
                     font=F_BIG_STATUS, background=C_BG, foreground=C_ACCENT)
    style.configure("Muted.TLabel",
                     font=F_MUTED, background=C_BG, foreground=C_MUTED)
    style.configure("Card.TLabelframe",
                     background=C_SURFACE, relief="solid", borderwidth=1,
                     bordercolor=C_BORDER, lightcolor=C_BORDER, darkcolor=C_BORDER)
    style.configure("Card.TLabelframe.Label",
                     background=C_SURFACE, foreground=C_TEXT_SECONDARY, font=F_BOLD)
    style.configure("Card.TFrame", background=C_SURFACE)

    style.configure("Accent.TButton",
                     background=C_BUTTON_PRIMARY, foreground=C_TEXT,
                     font=F_BOLD, borderwidth=1, focusthickness=0,
                     focuscolor="", relief="solid", padding=(10, 8),
                     bordercolor=C_BUTTON_BORDER, lightcolor=C_BUTTON_BORDER, darkcolor=C_BUTTON_BORDER)
    style.map("Accent.TButton",
              background=[("active", C_ACCENT), ("pressed", C_ACCENT),
                          ("disabled", C_DISABLED_BG)],
              foreground=[("disabled", C_DISABLED_FG),
                          ("active", C_ON_ACCENT), ("pressed", C_ON_ACCENT)])

    style.configure("TButton", font=F_BOLD, borderwidth=1, focusthickness=0,
                    focuscolor="", relief="solid", padding=(10, 6),
                    background=C_BUTTON, foreground=C_TEXT,
                    bordercolor=C_BUTTON_BORDER, lightcolor=C_BUTTON_BORDER, darkcolor=C_BUTTON_BORDER)
    style.map("TButton",
              background=[("active", C_ACCENT), ("pressed", C_ACCENT),
                          ("disabled", C_DISABLED_BG)],
              foreground=[("active", C_ON_ACCENT), ("pressed", C_ON_ACCENT),
                          ("disabled", C_DISABLED_FG)])

    style.configure("TEntry",
                    fieldbackground=C_SURFACE,
                    foreground=C_TEXT,
                    bordercolor=C_BORDER,
                    lightcolor=C_BORDER,
                    darkcolor=C_BORDER,
                    insertcolor=C_TEXT)
    style.configure("TCombobox",
                    fieldbackground=C_SURFACE,
                    background=C_SURFACE,
                    foreground=C_TEXT,
                    arrowcolor=C_TEXT,
                    bordercolor=C_BORDER,
                    lightcolor=C_BORDER,
                    darkcolor=C_BORDER)
    style.map("TCombobox",
              fieldbackground=[("readonly", C_SURFACE)],
              background=[("readonly", C_SURFACE)],
              foreground=[("readonly", C_TEXT)])
    style.configure("TNotebook", background=C_BG, borderwidth=0)
    style.configure("TNotebook.Tab", font=F_BASE, padding=[10, 4], background=C_SURFACE, foreground=C_TEXT)
    style.map("TNotebook.Tab",
              background=[("selected", C_ACCENT), ("active", C_SURFACE)],
              foreground=[("selected", C_ON_ACCENT), ("active", C_TEXT)])
    style.configure("Treeview",
                     background=C_SURFACE, foreground=C_TEXT,
                     fieldbackground=C_SURFACE, font=F_BASE,
                     bordercolor=C_BORDER, lightcolor=C_BORDER, darkcolor=C_BORDER)
    style.map("Treeview", background=[("selected", C_ACCENT)], foreground=[("selected", C_ON_ACCENT)])
    style.configure("Treeview.Heading", font=F_BOLD, background=C_SURFACE, foreground=C_TEXT)


def _sec_to_hms(seconds: float) -> str:
    seconds = int(seconds)
    h, r = divmod(seconds, 3600)
    m, s = divmod(r, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


# =============================================================================
# UI — SPLASH SCREEN
# =============================================================================

class SplashScreen:

    def __init__(self, root: tk.Tk, on_done):
        self.root    = root
        self.on_done = on_done
        self.win     = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.configure(bg=C_SPLASH_BG)
        self.win.attributes("-topmost", True)
        _apply_window_icon(self.win)
        _center_window(self.win, SPLASH_W, SPLASH_H)

        # Siempre usa el fallback de texto — no depende de archivos externos
        self._fallback_label()

        self.root.after(SPLASH_DURATION_MS, self._start_fade)

    def _fallback_label(self):
        container = tk.Frame(self.win, bg=C_SPLASH_BG, padx=22, pady=20)
        container.pack(expand=True, fill="both")

        panel = tk.Frame(container, bg=C_SPLASH_PANEL, highlightthickness=1,
                         highlightbackground=C_ACCENT)
        panel.pack(expand=True, fill="both")

        tk.Frame(panel, bg=C_SPLASH_ACCENT, height=5).pack(fill="x")

        body = tk.Frame(panel, bg=C_SPLASH_PANEL, padx=24, pady=26)
        body.pack(expand=True, fill="both")

        if SPLASH_IMAGE_FILE.exists():
            try:
                self._splash_img = tk.PhotoImage(file=str(SPLASH_IMAGE_FILE))
                tk.Label(body, image=self._splash_img,
                         bg=C_SPLASH_PANEL).pack(pady=(0, 10))
            except Exception:
                self._splash_img = None

        tk.Label(body, text=APP_TITLE,
                 font=("Segoe UI", 28, "bold"),
                 bg=C_SPLASH_PANEL, fg=C_ON_ACCENT).pack(anchor="center")
        tk.Label(body, text="Activity Tracker",
                 font=("Segoe UI", 11),
                 bg=C_SPLASH_PANEL, fg=C_SPLASH_MUTED).pack(pady=(4, 8))
        tk.Label(body, text="Unified four-color palette",
                 font=("Segoe UI", 10),
                 bg=C_SPLASH_PANEL, fg=C_SPLASH_MUTED).pack()

    def _start_fade(self):
        self._fade_step = 0
        self._fade()

    def _fade(self):
        self._fade_step += 1
        alpha = 1.0 - (self._fade_step / SPLASH_FADE_STEPS)
        if alpha <= 0:
            self.win.destroy()
            self.on_done()
            return
        self.win.attributes("-alpha", max(alpha, 0))
        step_ms = SPLASH_FADE_MS // SPLASH_FADE_STEPS
        self.root.after(step_ms, self._fade)


# =============================================================================
# UI — START VIEW
# =============================================================================

class StartView(tk.Frame):

    def __init__(self, parent, cfg_mgr: ConfigManager, navigate):
        super().__init__(parent, bg=C_BG)
        self.cfg_mgr  = cfg_mgr
        self.navigate = navigate
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        center = tk.Frame(self, bg=C_BG)
        center.grid(row=0, column=0, sticky="nsew", padx=16, pady=(20, 8))
        center.columnconfigure(0, weight=1)
        center.rowconfigure(0, weight=1)

        content = tk.Frame(center, bg=C_BG)
        content.grid(row=0, column=0)

        ttk.Label(content, text="Welcome to", style="Muted.TLabel").grid(
            row=0, column=0, pady=(0, 2))
        ttk.Label(content, text=APP_TITLE, style="Title.TLabel").grid(
            row=1, column=0, pady=(0, 6))
        ttk.Label(content, text="Activity tracking for remote and on-site agents",
                  style="Muted.TLabel", justify="center", anchor="center").grid(
            row=2, column=0, pady=(0, 8))

        ttk.Button(content, text="Start", style="Accent.TButton",
                   command=self._on_start, width=18).grid(
            row=3, column=0, pady=(10, 0))

        bottom = tk.Frame(self, bg=C_BG)
        bottom.grid(row=1, column=0, sticky="sew", padx=16, pady=16)
        bottom.columnconfigure(1, weight=1)

        ttk.Button(bottom, text="Admin Panel",
                   command=lambda: self.navigate("admin_login")).grid(
            row=0, column=0, sticky="w")
        ttk.Label(bottom, text=APP_VERSION, style="Muted.TLabel").grid(
            row=0, column=2, sticky="e")

    def _on_start(self):
        cfg = self.cfg_mgr.cfg
        if not _config_is_ready(cfg):
            messagebox.showwarning(
                APP_TITLE,
                "This workstation has no valid agent configuration.\n\n"
                "Please contact your admin and import the assigned config file "
                "from Admin Panel.",
            )
            return
        self.navigate("agent")


# =============================================================================
# UI — ADMIN LOGIN VIEW
# =============================================================================

class AdminLoginView(tk.Frame):

    def __init__(self, parent, cfg_mgr: ConfigManager, navigate):
        super().__init__(parent, bg=C_BG)
        self.cfg_mgr   = cfg_mgr
        self.navigate  = navigate
        self._attempts = 0
        self._locked   = False
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        tk.Frame(self, height=60, bg=C_BG).grid(row=0, column=0)
        ttk.Label(self, text="Admin Login", style="Title.TLabel").grid(
            row=1, column=0, pady=(0, 24))

        form = tk.Frame(self, bg=C_BG)
        form.grid(row=2, column=0, padx=60)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Username").grid(row=0, column=0, sticky="w", pady=4)
        self._user_var = tk.StringVar()
        ttk.Entry(form, textvariable=self._user_var, width=24).grid(
            row=0, column=1, pady=4, padx=(8, 0))

        ttk.Label(form, text="Password").grid(row=1, column=0, sticky="w", pady=4)
        self._pw_var = tk.StringVar()
        self._pw_entry = ttk.Entry(form, textvariable=self._pw_var,
                                   show="*", width=24)
        self._pw_entry.grid(row=1, column=1, pady=4, padx=(8, 0))
        self._pw_entry.bind("<Return>", lambda _: self._login())

        self._err_lbl = ttk.Label(form, text="", foreground=C_ERROR,
                                   background=C_BG)
        self._err_lbl.grid(row=2, column=0, columnspan=2, pady=(8, 0))

        btns = tk.Frame(self, bg=C_BG)
        btns.grid(row=3, column=0, pady=20)
        ttk.Button(btns, text="Cancel",
                   command=lambda: self.navigate("start")).pack(
            side="left", padx=8)
        ttk.Button(btns, text="Login", style="Accent.TButton",
                   command=self._login).pack(side="left", padx=8)

    def _login(self):
        if self._locked:
            return
        cfg  = self.cfg_mgr.cfg
        user = self._user_var.get().strip()
        pw   = self._pw_var.get()
        if user != cfg.admin_username:
            self._fail()
            return
        ok = SecurityUtils.verify_password(pw, cfg.admin_password,
                                            cfg.admin_password_hashed)
        if ok:
            self._attempts = 0
            self.navigate("admin_console")
        else:
            self._fail()

    def _fail(self):
        self._attempts += 1
        remaining = ADMIN_MAX_ATTEMPTS - self._attempts
        if self._attempts >= ADMIN_MAX_ATTEMPTS:
            self._locked = True
            self._countdown(ADMIN_LOCKOUT_SECONDS)
        else:
            self._err_lbl.config(
                text=f"Invalid credentials. {remaining} attempt(s) left.")

    def _countdown(self, sec: int):
        if sec <= 0:
            self._locked   = False
            self._attempts = 0
            self._err_lbl.config(text="")
            return
        self._err_lbl.config(
            text=f"Too many attempts. Retry in {sec}s…")
        self.after(1000, lambda: self._countdown(sec - 1))


# =============================================================================
# UI — ADMIN CONSOLE VIEW
# =============================================================================

class AdminConsoleView(tk.Frame):

    def __init__(self, parent, cfg_mgr: ConfigManager, navigate):
        super().__init__(parent, bg=C_BG)
        self.cfg_mgr  = cfg_mgr
        self.navigate = navigate
        self._selected_credentials_source = None
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        hdr = tk.Frame(self, bg=C_BG)
        hdr.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 0))
        hdr.columnconfigure(1, weight=1)
        ttk.Button(hdr, text="← Close",
                   command=lambda: self.navigate("start")).grid(
            row=0, column=0, sticky="w")
        ttk.Label(hdr, text="Admin Console", style="Title.TLabel").grid(
            row=0, column=1, sticky="w", padx=12)
        ttk.Label(self, text="Changes apply to new sessions",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w",
                                              padx=10, pady=(0, 8))

        nb = ttk.Notebook(self)
        nb.grid(row=2, column=0, sticky="nsew", padx=10, pady=4)
        self.rowconfigure(2, weight=1)

        nb.add(self._tab_settings(nb),    text="Settings")
        nb.add(self._tab_credentials(nb), text="Credentials")

    # ── Tab: Settings ───────────────────────────────────────────────────────
    def _tab_settings(self, parent):
        f = tk.Frame(parent, bg=C_BG, padx=16, pady=12)
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="Manual agent editing disabled in this build.",
                  font=F_BOLD, background=C_BG).grid(row=0, column=0, columnspan=2,
                                                     sticky="w")
        ttk.Label(
            f,
            text=(
                "Import one assigned settings file with Agent ID, Agent Name, "
                "Project and Work Mode. Supported formats: JSON, or TXT with "
                "JSON / key=value lines."
            ),
            style="Muted.TLabel",
            wraplength=360,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 10))

        self._settings_status = ttk.Label(f, text="", style="Muted.TLabel")
        self._settings_status.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self._import_info_lbl = ttk.Label(f, text="", style="Muted.TLabel",
                                          wraplength=360, justify="left")
        self._import_info_lbl.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 12))

        preview = tk.Frame(f, bg=C_BG)
        preview.grid(row=4, column=0, columnspan=2, sticky="ew")
        preview.columnconfigure(1, weight=1)

        labels = [
            ("Agent ID", "_settings_agent_id"),
            ("Agent Name", "_settings_agent_name"),
            ("Project", "_settings_project"),
            ("Work Mode", "_settings_work_mode"),
        ]
        for idx, (label, attr) in enumerate(labels):
            ttk.Label(preview, text=label).grid(row=idx, column=0, sticky="w", pady=4)
            var = tk.StringVar(value="—")
            setattr(self, attr, var)
            ttk.Label(preview, textvariable=var, font=F_BOLD,
                      background=C_BG).grid(row=idx, column=1, sticky="w", padx=10, pady=4)

        btn_row = tk.Frame(f, bg=C_BG)
        btn_row.grid(row=5, column=0, columnspan=2, sticky="w", pady=(16, 0))
        ttk.Button(btn_row, text="Import Settings File", style="Accent.TButton",
                   command=self._import_settings_file).pack(side="left")
        ttk.Button(btn_row, text="Reload Current Config",
                   command=self._refresh_settings_preview).pack(side="left", padx=(8, 0))

        self._refresh_settings_preview()
        return f

    def _refresh_settings_preview(self):
        self.cfg_mgr.load()
        cfg = self.cfg_mgr.cfg
        self._settings_agent_id.set(cfg.agent_id or "—")
        self._settings_agent_name.set(cfg.agent_name or "—")
        self._settings_project.set(cfg.project_name or "—")
        self._settings_work_mode.set(cfg.work_mode or "—")

        if _config_is_ready(cfg):
            self._settings_status.config(text="Configuration loaded and ready.")
        else:
            self._settings_status.config(text="No valid imported configuration loaded yet.")

        import_lines = []
        if cfg.import_source_file:
            import_lines.append(f"Source: {cfg.import_source_file}")
        if cfg.import_imported_at:
            import_lines.append(f"Imported at: {cfg.import_imported_at}")
        self._import_info_lbl.config(
            text="\n".join(import_lines) if import_lines else "No import history available."
        )

    def _import_settings_file(self):
        selected = filedialog.askopenfilename(
            title="Import agent settings",
            filetypes=[
                ("Config files", "*.json *.txt *.cfg *.ini"),
                ("JSON files", "*.json"),
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ],
        )
        if not selected:
            return
        try:
            profile = load_import_settings_file(selected)
            self.cfg_mgr.save_imported_profile(profile, source_file=selected)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not import config file:\n{exc}")
            return
        self._refresh_settings_preview()
        messagebox.showinfo(
            APP_TITLE,
            "Agent configuration imported successfully.\n\n"
            f"Agent ID: {profile['agent_id']}\n"
            f"Agent Name: {profile['agent_name']}\n"
            f"Project: {profile['project_name']}\n"
            f"Work Mode: {profile['work_mode']}",
        )

    # ── Tab: Credentials ────────────────────────────────────────────────────
    def _tab_credentials(self, parent):
        f = tk.Frame(parent, bg=C_BG, padx=16, pady=12)
        cfg = self.cfg_mgr.cfg

        ttk.Label(f, text="Current username:", background=C_BG).pack(anchor="w")
        ttk.Label(f, text=cfg.admin_username, font=F_BOLD,
                  background=C_BG).pack(anchor="w", pady=(0, 8))
        ttk.Separator(f).pack(fill="x", pady=4)

        def _row(label, var_name, value="", show="*"):
            ttk.Label(f, text=label, background=C_BG).pack(anchor="w", pady=(4, 0))
            var = tk.StringVar(value=value)
            setattr(self, var_name, var)
            ttk.Entry(f, textvariable=var, show=show, width=34).pack(anchor="w", fill="x")

        _row("Current password", "_cred_cur_pw", "")
        ttk.Separator(f).pack(fill="x", pady=8)
        _row("New username", "_cred_new_user", cfg.admin_username, show="")
        _row("New password", "_cred_new_pw", "")
        _row("Confirm new password", "_cred_confirm_pw", "")
        _row("Export password", "_cred_export_pw", cfg.export_password, show="")

        ttk.Separator(f).pack(fill="x", pady=8)

        self._gs_enabled = tk.BooleanVar(value=cfg.sheets_enabled)
        ttk.Checkbutton(f, text="Enable Google Sheets sync",
                        variable=self._gs_enabled).pack(anchor="w", pady=(2, 6))

        ttk.Label(f, text="Google Sheets link", background=C_BG).pack(anchor="w")
        self._gs_link = tk.StringVar(value=cfg.sheets_spreadsheet_link)
        ttk.Entry(f, textvariable=self._gs_link, width=34).pack(anchor="w", fill="x")

        ttk.Label(f, text="credentials.json stored in LOCALAPPDATA",
                  background=C_BG).pack(anchor="w", pady=(8, 0))
        cred_row = tk.Frame(f, bg=C_BG)
        cred_row.pack(fill="x", pady=(2, 0))
        self._gs_cred_path = tk.StringVar(value=cfg.sheets_credentials_file)
        ttk.Entry(cred_row, textvariable=self._gs_cred_path, width=26).pack(
            side="left", fill="x", expand=True)
        ttk.Button(cred_row, text="Select JSON",
                   command=self._select_credentials_json).pack(side="left", padx=(6, 0))

        ttk.Button(f, text="Save Credentials", style="Accent.TButton",
                   command=self._save_credentials).pack(pady=12, ipady=4)
        return f

    def _select_credentials_json(self):
        selected = filedialog.askopenfilename(
            title="Select credentials.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not selected:
            return
        self._selected_credentials_source = selected
        self._gs_cred_path.set(str(CREDENTIALS_FILE))

    def _save_credentials(self):
        cfg      = self.cfg_mgr.cfg
        cur_pw   = self._cred_cur_pw.get()
        new_user = self._cred_new_user.get().strip() or cfg.admin_username
        new_pw   = self._cred_new_pw.get()
        confirm  = self._cred_confirm_pw.get()
        exp_pw   = self._cred_export_pw.get().strip() or cfg.export_password
        gs_enabled = self._gs_enabled.get()
        gs_link    = self._gs_link.get().strip()
        cred_dest  = str(CREDENTIALS_FILE)

        if not SecurityUtils.verify_password(cur_pw, cfg.admin_password,
                                            cfg.admin_password_hashed):
            messagebox.showerror(APP_TITLE, "Current password is incorrect.")
            return
        if new_pw and new_pw != confirm:
            messagebox.showerror(APP_TITLE, "New password and confirmation do not match.")
            return
        if not exp_pw:
            messagebox.showerror(APP_TITLE, "Export password cannot be empty.")
            return

        final_password_hash = cfg.admin_password
        final_hashed_flag   = cfg.admin_password_hashed
        if new_pw:
            if bcrypt is None:
                messagebox.showerror(APP_TITLE, "bcrypt not installed — cannot hash password.")
                return
            final_password_hash = SecurityUtils.hash_password(new_pw)
            final_hashed_flag   = True
        elif not cfg.admin_password_hashed:
            if bcrypt is None:
                messagebox.showerror(APP_TITLE, "bcrypt not installed — cannot hash password.")
                return
            final_password_hash = SecurityUtils.hash_password(cur_pw)
            final_hashed_flag   = True

        if gs_enabled and not _extract_sheet_id(gs_link):
            messagebox.showerror(
                APP_TITLE,
                "Please provide a valid Google Sheets link or spreadsheet ID."
            )
            return

        if self._selected_credentials_source:
            try:
                raw_json = json.loads(
                    Path(self._selected_credentials_source).read_text(encoding="utf-8")
                )
                DPAPIStore.write_json(
                    CREDENTIALS_FILE,
                    raw_json,
                    description="DeskPulseSuite Google Credentials"
                )
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"Could not encrypt credentials.json:\n{exc}")
                return
        elif gs_enabled and not Path(cfg.sheets_credentials_file).exists():
            messagebox.showerror(
                APP_TITLE,
                "Select credentials.json before enabling Google Sheets."
            )
            return

        self.cfg_mgr.save_credentials(
            new_user,
            final_password_hash,
            exp_pw,
            password_hashed=final_hashed_flag,
        )
        self.cfg_mgr.save_google_sheets(gs_enabled, gs_link, cred_dest)
        self._gs_cred_path.set(str(CREDENTIALS_FILE))
        self._selected_credentials_source = ""
        messagebox.showinfo(APP_TITLE, "Credentials updated successfully.")


# =============================================================================
# UI — DIALOGS
# =============================================================================

class RecoveryDialog:
    """Popup shown when a recoverable session state is found."""

    def __init__(self, parent, state: Dict, on_new, on_resume):
        self.result = None
        dlg = _create_dialog(parent, "Recover Previous Session", 470, 240, topmost=True)

        tk.Label(dlg, text="Recover Previous Session",
                 font=F_TITLE, bg=C_BG).pack(pady=(16, 4))
        ci = state.get("clock_in", "Unknown time")
        tk.Label(dlg, text=f"Session started at: {ci}",
                 font=F_BOLD, bg=C_BG).pack(pady=4)
        tk.Label(dlg,
                 text="An unsaved session was found.\n"
                      "Would you like to resume it or start a new session?",
                 bg=C_BG, fg=C_MUTED, wraplength=400, justify="center").pack(pady=8)

        btns = tk.Frame(dlg, bg=C_BG)
        btns.pack(pady=12)
        ttk.Button(btns, text="Start New Session",
                   command=lambda: (on_new(), dlg.destroy())).pack(
            side="left", padx=10)
        ttk.Button(btns, text="Resume Session", style="Accent.TButton",
                   command=lambda: (on_resume(), dlg.destroy())).pack(
            side="left", padx=10)
        _show_dialog(dlg)
        dlg.wait_window()


class SessionTypeDialog:
    """Popup to choose Normal / Schedule Change / Overtime before Clock In."""

    def __init__(self, parent):
        self.session_type: Optional[str] = None
        dlg = _create_dialog(parent, "Start Session", 440, 290)

        tk.Label(dlg, text="Start Session", font=F_TITLE, bg=C_BG).pack(pady=(16, 8))

        var = tk.StringVar(value="Normal")
        options = [
            ("Normal",          "Regular scheduled shift"),
            ("Schedule Change", "Schedule was modified for today"),
            ("Overtime",        "Extra hours requested by supervisors"),
        ]
        for val, desc in options:
            row = tk.Frame(dlg, bg=C_BG)
            row.pack(fill="x", padx=30, pady=3)
            ttk.Radiobutton(row, text=val, variable=var, value=val).pack(side="left")
            tk.Label(row, text=f"— {desc}", bg=C_BG, fg=C_MUTED).pack(side="left")

        btns = tk.Frame(dlg, bg=C_BG)
        btns.pack(pady=16)

        def _confirm():
            self.session_type = var.get()
            dlg.destroy()

        ttk.Button(btns, text="Cancel",
                   command=dlg.destroy).pack(side="left", padx=8)
        ttk.Button(btns, text="Clock In", style="Accent.TButton",
                   command=_confirm).pack(side="left", padx=8)
        _show_dialog(dlg)
        dlg.wait_window()


class ClockOutConfirmDialog:
    """Confirm clock-out and optionally report overtime excess over 8 hours."""

    def __init__(self, parent, session: SessionData, net_sec: float):
        self.confirmed = False
        self.ot_reported = False
        self.ot_note = ""
        self.ot_requested_by = ""
        self._session = session
        self._net_sec = max(0.0, float(net_sec))
        self._parent = parent
        self._show_confirm()

    def _show_confirm(self):
        dlg = _create_dialog(self._parent, "Clock Out", 420, 240)

        tk.Label(dlg, text="Clock Out",
                 font=F_TITLE, bg=C_BG).pack(pady=(16, 8))
        tk.Label(dlg,
                 text="Are you sure you want to clock out?\nMake sure all your tasks are completed.",
                 bg=C_BG, fg=C_MUTED, wraplength=380,
                 justify="center").pack(pady=8)

        confirm_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            dlg,
            text="I confirm I am ready to clock out",
            variable=confirm_var,
            command=lambda: btn_co.config(
                state=(tk.NORMAL if confirm_var.get() else tk.DISABLED)
            ),
        ).pack(pady=(4, 8))

        btns = tk.Frame(dlg, bg=C_BG)
        btns.pack(pady=8)

        def _yes():
            if not confirm_var.get():
                return
            self.confirmed = True
            dlg.destroy()

        ttk.Button(btns, text="Cancel",
                   command=dlg.destroy).pack(side="left", padx=8)
        btn_co = ttk.Button(btns, text="Clock Out", style="Accent.TButton",
                            command=_yes, state=tk.DISABLED)
        btn_co.pack(side="left", padx=8)
        _show_dialog(dlg)
        dlg.wait_window()

        if self.confirmed and self._net_sec > REGULAR_WORKDAY_SEC:
            self._show_ot_step1()

    def _show_ot_step1(self):
        dlg = _create_dialog(self._parent, "Overtime Alert", 470, 250)

        total_fmt = _sec_to_hms(self._net_sec)
        excess_fmt = _sec_to_hms(self._net_sec - REGULAR_WORKDAY_SEC)

        tk.Label(dlg, text="Overtime Alert",
                 font=F_TITLE, bg=C_BG).pack(pady=(16, 8))
        tk.Label(
            dlg,
            text=(
                "This session lasted more than 8 hours.\n\n"
                f"Total worked time: {total_fmt}\n"
                f"Do you want to report {excess_fmt} as overtime?"
            ),
            bg=C_BG, fg=C_MUTED, justify="center", wraplength=410,
        ).pack(pady=8)

        btns = tk.Frame(dlg, bg=C_BG)
        btns.pack(pady=16)

        def _no():
            dlg.destroy()

        def _yes():
            dlg.destroy()
            self._show_ot_step2(excess_fmt)

        ttk.Button(btns, text="No, keep payable time at 8h",
                   command=_no).pack(side="left", padx=8)
        ttk.Button(btns, text="Yes, report overtime",
                   style="Accent.TButton", command=_yes).pack(side="left", padx=8)
        _show_dialog(dlg)
        dlg.wait_window()

    def _show_ot_step2(self, excess_fmt: str):
        dlg = _create_dialog(self._parent, "Report Overtime", 470, 320)

        tk.Label(dlg, text="Report Overtime",
                 font=F_TITLE, bg=C_BG).pack(pady=(16, 8))
        tk.Label(
            dlg,
            text=(
                f"Overtime to report: {excess_fmt}\n"
                "Select who requested or authorized it, then confirm."
            ),
            bg=C_BG, fg=C_MUTED, justify="center", wraplength=420,
        ).pack(pady=(4, 10))

        requester_var = tk.StringVar(value="")
        confirm_var = tk.BooleanVar(value=False)

        form = tk.Frame(dlg, bg=C_BG)
        form.pack(padx=18, pady=6, fill="x")

        tk.Label(form, text="Requested by", bg=C_BG, anchor="w").pack(fill="x")
        requester_combo = ttk.Combobox(
            form, textvariable=requester_var, values=OVERTIME_REQUESTERS,
            state="readonly", width=28
        )
        requester_combo.pack(fill="x", pady=(4, 10))

        ttk.Checkbutton(
            form,
            text="I confirm this overtime was requested / authorized.",
            variable=confirm_var,
        ).pack(anchor="w")

        err_lbl = tk.Label(dlg, text="", fg=C_ERROR, bg=C_BG)
        err_lbl.pack(pady=(8, 0))

        btns = tk.Frame(dlg, bg=C_BG)
        btns.pack(pady=14)

        def _refresh_confirm(*_):
            ready = confirm_var.get() and requester_var.get().strip() in OVERTIME_REQUESTERS
            if ready:
                confirm_btn.state(["!disabled"])
                err_lbl.config(text="")
            else:
                confirm_btn.state(["disabled"])

        def _back():
            dlg.destroy()
            self._show_ot_step1()

        def _submit():
            if requester_var.get().strip() not in OVERTIME_REQUESTERS:
                err_lbl.config(text="Select who requested the overtime.")
                return
            if not confirm_var.get():
                err_lbl.config(text="You must confirm before submitting.")
                return
            self.ot_reported = True
            self.ot_requested_by = requester_var.get().strip()
            self.ot_note = f"Requested by: {self.ot_requested_by}"
            dlg.destroy()

        ttk.Button(btns, text="Back", command=_back).pack(side="left", padx=8)
        confirm_btn = ttk.Button(btns, text="Confirm overtime",
                                 style="Accent.TButton", command=_submit)
        confirm_btn.pack(side="left", padx=8)
        confirm_btn.state(["disabled"])

        requester_combo.bind("<<ComboboxSelected>>", _refresh_confirm)
        confirm_var.trace_add("write", lambda *_: _refresh_confirm())
        _show_dialog(dlg)
        dlg.wait_window()


class AppCloseConfirmDialog:
    """Hard confirmation dialog for closing the app from the window X button."""

    def __init__(self, parent, has_active_session: bool):
        self.confirmed = False
        self._parent = parent
        self._has_active_session = has_active_session
        self._show()

    def _show(self):
        dlg = _create_dialog(self._parent, "Close Application", 420, 190)

        if self._has_active_session:
            # Full confirmation with checkbox when session is active
            _center_window(dlg, 470, 260)
            tk.Label(dlg, text="Close Application",
                     font=F_TITLE, bg=C_BG).pack(pady=(16, 8))
            tk.Label(
                dlg,
                text="Are you sure you want to close DeskPulse Suite?\nThe active session will be clocked out automatically.",
                bg=C_BG, fg=C_MUTED, wraplength=420, justify="center",
            ).pack(pady=(4, 10))

            confirm_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(
                dlg,
                text="I confirm that I am sure",
                variable=confirm_var,
                command=lambda: btn_accept.config(
                    state=(tk.NORMAL if confirm_var.get() else tk.DISABLED)
                ),
            ).pack(pady=(0, 6))

            tk.Label(dlg, text="Active session detected: closing will force Clock Out.",
                     bg=C_BG, fg=C_ERROR).pack()

            btns = tk.Frame(dlg, bg=C_BG)
            btns.pack(pady=14)

            def _accept():
                if not confirm_var.get():
                    return
                self.confirmed = True
                dlg.destroy()

            ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
            btn_accept = ttk.Button(btns, text="Close App", style="Accent.TButton",
                                    command=_accept, state=tk.DISABLED)
            btn_accept.pack(side="left", padx=8)

        else:
            # Simple confirmation without checkbox when no active session
            _center_window(dlg, 420, 190)
            tk.Label(dlg, text="Close Application",
                     font=F_TITLE, bg=C_BG).pack(pady=(16, 8))
            tk.Label(
                dlg,
                text="Are you sure you want to close DeskPulse Suite?",
                bg=C_BG, fg=C_MUTED, wraplength=380, justify="center",
            ).pack(pady=(4, 16))

            btns = tk.Frame(dlg, bg=C_BG)
            btns.pack(pady=8)

            def _accept_simple():
                self.confirmed = True
                dlg.destroy()

            ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
            ttk.Button(btns, text="Close App", style="Accent.TButton",
                       command=_accept_simple).pack(side="left", padx=8)

        _show_dialog(dlg)
        dlg.wait_window()


class SessionClosedDialog:
    """Styled dialog shown after a session is successfully closed — mirrors the design in the spec."""

    def __init__(self, parent, session: SessionData,
                 clock_in_fmt: str, clock_out_fmt: str,
                 worked_fmt: str, lunch_fmt: str,
                 inactive_fmt: str):
        dlg = _create_dialog(parent, "Session Closed", 420, 520)

        # ── Dark header ─────────────────────────────────────────────────────
        header = tk.Frame(dlg, bg=C_DARK_CARD, padx=18, pady=18)
        header.pack(fill="x")
        header.columnconfigure(1, weight=1)

        # Checkmark circle
        check_frame = tk.Frame(header, bg=C_ACCENT,
                               width=44, height=44,
                               highlightthickness=1,
                               highlightbackground=C_DARK_BADGE_BD)
        check_frame.pack_propagate(False)
        check_frame.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 14))
        tk.Label(check_frame, text="✓", font=(FONT_FAMILY, 18, "bold"),
                 fg=C_DARK_TEXT, bg=C_ACCENT).place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(header, text="Session Ended",
                 font=(FONT_FAMILY, 14, "bold"),
                 fg=C_DARK_TEXT, bg=C_DARK_CARD).grid(row=0, column=1, sticky="w")
        tk.Label(header, text="Activity Report",
                 font=(FONT_FAMILY, 10),
                 fg=C_DARK_MUTED, bg=C_DARK_CARD).grid(row=1, column=1, sticky="w", pady=(2, 0))

        # ── White body ───────────────────────────────────────────────────────
        body = tk.Frame(dlg, bg=C_SURFACE, padx=18, pady=18)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        # Metric cards row
        # WORKED — dark card
        worked_card = tk.Frame(body, bg=C_WORKED_CARD_BG, padx=14, pady=12,
                               highlightthickness=0)
        worked_card.grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=(0, 16))
        tk.Label(worked_card, text="SESSION DURATION", font=(FONT_FAMILY, 9, "bold"),
                 fg=C_DARK_MUTED, bg=C_WORKED_CARD_BG).pack(anchor="w")
        tk.Label(worked_card, text=worked_fmt, font=(FONT_FAMILY, 18, "bold"),
                 fg=C_DARK_TEXT, bg=C_WORKED_CARD_BG).pack(anchor="w", pady=(4, 0))

        # INACTIVE — white card with soft pulse
        inactive_card = tk.Frame(
            body,
            bg=C_SURFACE,
            padx=14,
            pady=12,
            highlightthickness=1,
            highlightbackground=C_STATUS_IDLE,
        )
        inactive_card.grid(row=0, column=1, sticky="ew", padx=(6, 0), pady=(0, 16))

        inactive_title = tk.Label(
            inactive_card,
            text="INACTIVE TIME",
            font=(FONT_FAMILY, 9, "bold"),
            fg=C_STATUS_IDLE,
            bg=C_SURFACE,
        )
        inactive_title.pack(anchor="w")

        inactive_value = tk.Label(
            inactive_card,
            text=inactive_fmt,
            font=(FONT_FAMILY, 18, "bold"),
            fg=C_STATUS_IDLE,
            bg=C_SURFACE,
        )
        inactive_value.pack(anchor="w", pady=(4, 0))

        def _pulse_inactive():
            try:
                current = inactive_card.cget("highlightbackground")
                next_color = C_BORDER if current == C_STATUS_IDLE else C_STATUS_IDLE
                inactive_card.config(highlightbackground=next_color)
                dlg.after(550, _pulse_inactive)
            except tk.TclError:
                pass

        dlg.after(550, _pulse_inactive)

        # Info rows
        sep_color = C_BORDER_SUBTLE

        def _info_row(parent, row_idx, label, value, badge=False, badge_color=None):
            tk.Frame(parent, bg=sep_color, height=1).grid(
                row=row_idx * 2, column=0, columnspan=2, sticky="ew", pady=(0, 0))
            tk.Label(parent, text=label, font=F_BASE,
                     fg=C_TEXT_SECONDARY, bg=C_SURFACE,
                     anchor="w").grid(row=row_idx * 2 + 1, column=0,
                                      sticky="w", pady=8)
            if badge:
                bc = badge_color or C_ACCENT_BG
                bd = C_ACCENT_BORDER if bc == C_ACCENT_BG else C_BORDER
                pill = tk.Label(parent, text=value, font=F_BOLD,
                                fg=C_ACCENT_SOFT if bc == C_ACCENT_BG else C_TEXT_SECONDARY,
                                bg=bc, padx=10, pady=3,
                                highlightthickness=1, highlightbackground=bd)
                pill.grid(row=row_idx * 2 + 1, column=1, sticky="e", pady=8)
            else:
                tk.Label(parent, text=value, font=F_BOLD,
                         fg=C_TEXT, bg=C_SURFACE,
                         anchor="e").grid(row=row_idx * 2 + 1, column=1,
                                          sticky="e", pady=8)

        info = tk.Frame(body, bg=C_SURFACE)
        info.grid(row=1, column=0, columnspan=2, sticky="ew")
        info.columnconfigure(0, weight=1)
        info.columnconfigure(1, weight=1)

        _info_row(info, 0, "Clock In",  clock_in_fmt)
        _info_row(info, 1, "Clock Out", clock_out_fmt)
        _info_row(info, 2, "Lunch Duration", lunch_fmt)
        _info_row(info, 3, "Type",      session.session_type or "Normal",  badge=True)
        ot_label = "Yes" if session.overtime else "No"
        _info_row(info, 4, "Overtime",  ot_label, badge=True,
                  badge_color=(C_ACCENT_BG if session.overtime else C_BG))

        # Close button
        tk.Frame(body, bg=C_BORDER_SUBTLE, height=1).grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        ttk.Button(body, text="Close", style="Accent.TButton",
                   command=dlg.destroy).grid(row=3, column=1, sticky="e", pady=12)
        _show_dialog(dlg)
        dlg.wait_window()


class LunchConfirmDialog:
    """Checkbox-gated confirmation for starting or ending Lunch."""

    def __init__(self, parent, starting: bool):
        self.confirmed = False
        action = "start" if starting else "end"
        msg_body = (
            "You will be marked as Lunch. Make sure\nyou are ready to take your break."
            if starting else
            "You will return to Online status\nfrom your lunch break."
        )
        dlg = _create_dialog(parent, f"{'Start' if starting else 'End'} Lunch", 400, 220)

        tk.Label(dlg, text=f"{'Start' if starting else 'End'} Lunch",
                 font=F_TITLE, bg=C_BG).pack(pady=(16, 8))
        tk.Label(dlg, text=msg_body, bg=C_BG, fg=C_MUTED,
                 wraplength=360, justify="center").pack(pady=6)

        confirm_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            dlg,
            text=f"I confirm I want to {action} Lunch",
            variable=confirm_var,
            command=lambda: btn_ok.config(
                state=(tk.NORMAL if confirm_var.get() else tk.DISABLED)
            ),
        ).pack(pady=(6, 8))

        btns = tk.Frame(dlg, bg=C_BG)
        btns.pack(pady=8)

        def _ok():
            if not confirm_var.get():
                return
            self.confirmed = True
            dlg.destroy()

        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        btn_ok = ttk.Button(btns, text="Continue", style="Accent.TButton",
                            command=_ok, state=tk.DISABLED)
        btn_ok.pack(side="left", padx=8)
        _show_dialog(dlg)
        dlg.wait_window()


class ExportDialog:
    """Export the full day folder encrypted with AES-256-GCM."""

    @staticmethod
    def _build_day_zip_bytes(day_dir: Path, root_name: str) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(day_dir.rglob("*")):
                if path.is_file():
                    arcname = f"{root_name}/{path.relative_to(day_dir).as_posix()}"
                    zf.write(path, arcname=arcname)
        return buffer.getvalue()

    @staticmethod
    def _encrypt_bytes(data: bytes, dest_path: Path, password: str):
        if not _CRYPTO_OK:
            raise RuntimeError("cryptography package not installed")

        salt = os.urandom(16)
        nonce = os.urandom(12)

        kdf = PBKDF2HMAC(
            algorithm=crypto_hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=260_000,
            backend=default_backend(),
        )
        key = kdf.derive(password.encode())
        ct = AESGCM(key).encrypt(nonce, data, None)

        dest_path.write_bytes(salt + nonce + ct)

    def __init__(self, parent, cfg_mgr: ConfigManager):
        cfg = cfg_mgr.cfg
        if not cfg.export_password or cfg.export_password == "ChangeMe!":
            messagebox.showwarning(
                APP_TITLE,
                "Export password is not configured. "
                "Please set it in Admin > Credentials first."
            )
            return

        dlg = tk.Toplevel(parent)
        _apply_window_icon(dlg)
        dlg.title("Export Day Package")
        dlg.resizable(False, False)
        dlg.transient(parent)
        _center_window(dlg, 420, 220)
        dlg.configure(bg=C_BG)
        dlg.grab_set()

        tk.Label(
            dlg,
            text="Export Encrypted Day Package",
            font=F_TITLE,
            bg=C_BG
        ).pack(pady=(16, 8))

        form = tk.Frame(dlg, bg=C_BG)
        form.pack(padx=20)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Date (YYYY-MM-DD)").grid(
            row=0, column=0, sticky="w", pady=4
        )
        date_var = tk.StringVar(value=_now_bolivia().strftime("%Y-%m-%d"))
        ttk.Entry(form, textvariable=date_var, width=16).grid(
            row=0, column=1, sticky="w", padx=8
        )

        ttk.Label(form, text="Destination folder").grid(
            row=1, column=0, sticky="w", pady=4
        )
        default_export_dir = Path.home()
        dest_var = tk.StringVar(value=str(default_export_dir))
        ttk.Entry(form, textvariable=dest_var, width=22).grid(
            row=1, column=1, sticky="w", padx=(8, 0)
        )
        ttk.Button(
            form,
            text="Browse",
            command=lambda: dest_var.set(
                filedialog.askdirectory(initialdir=str(default_export_dir)) or dest_var.get()
            )
        ).grid(row=1, column=2, padx=4)

        btns = tk.Frame(dlg, bg=C_BG)
        btns.pack(pady=12)

        def _export():
            date_str = date_var.get().strip()
            dest = dest_var.get().strip()

            if not dest:
                messagebox.showwarning(APP_TITLE, "Please select a destination folder.")
                return

            try:
                dt = datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror(APP_TITLE, "Invalid date format. Use YYYY-MM-DD.")
                return

            agent_id = cfg.agent_id.strip()
            agent_name = cfg.agent_name.strip()
            agent_folder = _agent_folder_label(agent_id, agent_name)
            day_dir = (
                RECORDS_DIR / agent_folder /
                dt.strftime("%Y") / dt.strftime("%m") / dt.strftime("%d")
            )
            legacy_day_dir = (
                RECORDS_DIR / agent_id /
                dt.strftime("%Y") / dt.strftime("%m") / dt.strftime("%d")
            )
            if (not day_dir.exists() or not day_dir.is_dir()) and legacy_day_dir.exists() and legacy_day_dir.is_dir():
                day_dir = legacy_day_dir

            if not day_dir.exists() or not day_dir.is_dir():
                messagebox.showerror(APP_TITLE, f"No records found for {date_str}.")
                return

            files = [p for p in day_dir.rglob("*") if p.is_file()]
            if not files:
                messagebox.showerror(APP_TITLE, f"No files found for {date_str}.")
                return

            out_base = Path(dest).expanduser()
            try:
                out_base.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                messagebox.showerror(
                    APP_TITLE,
                    f"Cannot access destination folder:\n{exc}"
                )
                return

            out_name = f"{agent_id}_{date_str}_day_package.enc"
            out_path = out_base / out_name

            try:
                zip_bytes = self._build_day_zip_bytes(
                    day_dir=day_dir,
                    root_name=f"{agent_id}_{date_str}"
                )
                self._encrypt_bytes(
                    data=zip_bytes,
                    dest_path=out_path,
                    password=cfg.export_password
                )
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"Export failed:\n{exc}")
                return

            messagebox.showinfo(
                APP_TITLE,
                f"Encrypted export created:\n{out_path}"
            )
            dlg.destroy()

        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        ttk.Button(btns, text="Export", style="Accent.TButton",
                   command=_export).pack(side="left", padx=8)

        dlg.wait_window()


# =============================================================================
# UI — AGENT VIEW
# =============================================================================

class AgentView(tk.Frame):

    def __init__(self, parent, cfg_mgr: ConfigManager,
                 sheets: SheetsSync, navigate):
        super().__init__(parent, bg=C_BG)
        self.cfg_mgr  = cfg_mgr
        self.sheets   = sheets
        self.navigate = navigate
        self._session: Optional[SessionData] = None
        self._worker:  Optional[SampleWorker] = None
        self._tick_job = None
        self._pulse_job = None
        self._pulse_phase = False
        self._state_lock = threading.Lock()  # protect state transitions
        self._build()
        self.after(300, self._check_recovery)

    def _make_action_button(self, parent, text: str, command, width: int = 11, compact: bool = False):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=C_BUTTON,
            fg=C_TEXT,
            activebackground=C_ACCENT,
            activeforeground=C_ON_ACCENT,
            disabledforeground=C_DISABLED_FG,
            relief="solid",
            bd=1,
            width=width,
            padx=6,
            pady=(6 if compact else 9),
            font=F_BOLD,
            cursor="hand2",
            highlightthickness=1,
            highlightbackground=C_BUTTON_BORDER,
            highlightcolor=C_ACCENT_BORDER,
            takefocus=True,
        )
    def _set_button_visual(self, btn: tk.Button, *, enabled: bool, selected: bool = False, text: Optional[str] = None):
        if text is not None:
            btn.config(text=text)
        if selected:
            btn.config(
                bg=C_ACCENT,
                fg=C_ON_ACCENT,
                activebackground=C_ACCENT_SOFT,
                activeforeground=C_ON_ACCENT,
                highlightbackground=C_ACCENT_BORDER,
                highlightcolor=C_ACCENT_BORDER,
            )
        else:
            btn.config(
                bg=C_DISABLED_BG if not enabled else C_BUTTON,
                fg=C_DISABLED_FG if not enabled else C_TEXT,
                activebackground=C_ACCENT,
                activeforeground=C_ON_ACCENT,
                highlightbackground=C_BUTTON_BORDER,
                highlightcolor=C_ACCENT_BORDER,
            )
        btn.config(state=(tk.NORMAL if enabled else tk.DISABLED), cursor=("hand2" if enabled else "arrow"))

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)

        # ── Top bar ──────────────────────────────────────────────────────────
        top_wrap = tk.Frame(self, bg=C_SURFACE, highlightthickness=1, highlightbackground=C_BORDER, bd=0)
        top_wrap.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        top_wrap.columnconfigure(0, weight=1)

        top = tk.Frame(top_wrap, bg=C_SURFACE, padx=10, pady=10, highlightthickness=0, bd=0)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        tk.Label(top, text="DESKPULSE", font=(FONT_FAMILY, 11, "bold"), fg=C_TEXT,
                 bg=C_SURFACE).grid(row=0, column=0, sticky="w")
        actions = tk.Frame(top, bg=C_SURFACE)
        actions.grid(row=0, column=2, sticky="e")
        self._btn_back = self._make_action_button(actions, "← Back", self._safe_back, width=9, compact=True)
        self._btn_back.pack(side="left", padx=(0, 8))
        self._btn_export = self._make_action_button(actions, "Export ↗", self._open_export, width=10, compact=True)
        self._btn_export.pack(side="left")

        # ── Agent card / session type ────────────────────────────────────────
        cfg = self.cfg_mgr.cfg
        card = tk.Frame(self, bg=C_DARK_CARD, padx=18, pady=16, highlightthickness=1,
                        highlightbackground=C_DARK_BADGE_BD, highlightcolor=C_DARK_BADGE_BD)
        card.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        card.columnconfigure(0, weight=1)

        self._session_badge_lbl = tk.Label(
            card,
            text="S. PENDING CLOCK IN",
            font=F_BOLD,
            fg=C_SESSION_BADGE_FG,
            bg=C_DARK_BADGE_BG,
            padx=10,
            pady=4,
            highlightthickness=1,
            highlightbackground=C_DARK_BADGE_BD,
            highlightcolor=C_DARK_BADGE_BD,
        )
        self._session_badge_lbl.grid(row=0, column=0, sticky="w")

        self._agent_lbl = tk.Label(
            card,
            text=f"{cfg.agent_name or '—'}",
            font=F_AGENT_NAME,
            fg=C_DARK_TEXT,
            bg=C_DARK_CARD,
        )
        self._agent_lbl.grid(row=1, column=0, sticky="w", pady=(14, 0))

        self._mode_lbl = tk.Label(
            card,
            text=f"{cfg.project_name or '—'} • {cfg.work_mode or '—'}",
            font=F_MUTED,
            fg=C_DARK_MUTED,
            bg=C_DARK_CARD,
        )
        self._mode_lbl.grid(row=2, column=0, sticky="w", pady=(4, 0))

        # ── Status box ───────────────────────────────────────────────────────
        status_card = tk.Frame(self, bg=C_SURFACE, padx=16, pady=16, highlightthickness=1,
                               highlightbackground=C_BORDER, highlightcolor=C_BORDER)
        status_card.grid(row=2, column=0, sticky="ew", padx=10, pady=(4, 0))
        status_card.columnconfigure(0, weight=1)

        # Status header row: label left, indicator badge right
        status_hdr = tk.Frame(status_card, bg=C_SURFACE)
        status_hdr.grid(row=0, column=0, sticky="ew")
        status_hdr.columnconfigure(0, weight=1)

        tk.Label(status_hdr, text="STATUS", font=F_BOLD,
                 fg=C_MUTED, bg=C_SURFACE).grid(row=0, column=0, sticky="w")

        self._status_indicator = tk.Label(
            status_hdr,
            text="● IDLE",
            font=(FONT_FAMILY, 9, "bold"),
            fg=C_MUTED,
            bg=C_SURFACE,
            padx=8, pady=3,
            highlightthickness=1,
            highlightbackground=C_BORDER,
        )
        self._status_indicator.grid(row=0, column=1, sticky="e")

        self._status_lbl = tk.Label(status_card, text="IDLE", font=(FONT_FAMILY, 24, "bold"),
                                    fg=C_MUTED, bg=C_SURFACE)
        self._status_lbl.grid(row=1, column=0, sticky="w", pady=(10, 2))

        self._status_elapsed_lbl = tk.Label(status_card, text="", font=(FONT_FAMILY, 10),
                                            fg=C_MUTED, bg=C_SURFACE)
        self._status_elapsed_lbl.grid(row=2, column=0, sticky="w", pady=(0, 12))

        # Counter mini-cards
        totals = tk.Frame(status_card, bg=C_SURFACE)
        totals.grid(row=3, column=0, sticky="ew")
        for i in range(3):
            totals.columnconfigure(i, weight=1)

        def _mini_card(parent, col):
            f = tk.Frame(parent, bg=C_SURFACE, padx=10, pady=8,
                         highlightthickness=1,
                         highlightbackground=C_BORDER,
                         highlightcolor=C_ACCENT_BORDER)
            f.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 4, 0))
            return f

        card_online  = _mini_card(totals, 0)
        card_lunch   = _mini_card(totals, 1)
        card_meeting = _mini_card(totals, 2)

        self._online_total_caption = tk.Label(card_online, text="ONLINE", font=(FONT_FAMILY, 9, "bold"),
                                              fg=C_MUTED, bg=C_SURFACE)
        self._online_total_caption.pack(anchor="w")
        self._online_total_lbl = tk.Label(card_online, text="00:00:00", font=(FONT_FAMILY, 13, "bold"),
                                          fg=C_DISABLED_FG, bg=C_SURFACE)
        self._online_total_lbl.pack(anchor="w", pady=(4, 0))

        self._lunch_total_caption = tk.Label(card_lunch, text="LUNCH", font=(FONT_FAMILY, 9, "bold"),
                                             fg=C_MUTED, bg=C_SURFACE)
        self._lunch_total_caption.pack(anchor="w")
        self._lunch_total_lbl = tk.Label(card_lunch, text="00:00:00", font=(FONT_FAMILY, 13, "bold"),
                                         fg=C_DISABLED_FG, bg=C_SURFACE)
        self._lunch_total_lbl.pack(anchor="w", pady=(4, 0))

        self._meeting_total_caption = tk.Label(card_meeting, text="MEETING", font=(FONT_FAMILY, 9, "bold"),
                                               fg=C_MUTED, bg=C_SURFACE)
        self._meeting_total_caption.pack(anchor="w")
        self._meeting_total_lbl = tk.Label(card_meeting, text="00:00:00", font=(FONT_FAMILY, 13, "bold"),
                                           fg=C_DISABLED_FG, bg=C_SURFACE)
        self._meeting_total_lbl.pack(anchor="w", pady=(4, 0))

        # Store references to mini-card frames for border coloring
        self._card_online  = card_online
        self._card_lunch   = card_lunch
        self._card_meeting = card_meeting

        self._update_session_header()

        tk.Frame(self, bg=C_BG).grid(row=4, column=0, sticky="nsew")

        btns = tk.Frame(self, bg=C_BG)
        btns.grid(row=5, column=0, sticky="sew", padx=10, pady=(10, 16))
        for i in range(4):
            btns.columnconfigure(i, weight=1)

        self._btn_ci = self._make_action_button(btns, "CLOCK\nIN", self._clock_in, width=10)
        self._btn_ln = self._make_action_button(btns, "LUNCH", self._toggle_lunch, width=10)
        self._btn_mt = self._make_action_button(btns, "MEETING", self._toggle_meeting, width=10)
        self._btn_co = self._make_action_button(btns, "CLOCK\nOUT", self._clock_out, width=10)

        for i, btn in enumerate((self._btn_ci, self._btn_ln, self._btn_mt, self._btn_co)):
            btn.grid(row=0, column=i, sticky="ew", padx=4)

        self._refresh_action_buttons()
        self._schedule_status_pulse()

    def _update_session_header(self):
        cfg = self.cfg_mgr.cfg
        current_type = (self._session.session_type if self._session else "Pending Clock In")
        badge_text = f"S. {str(current_type).upper()}"
        self._session_badge_lbl.config(text=badge_text,
                                       fg=C_SESSION_BADGE_FG,
                                       bg=C_DARK_BADGE_BG,
                                       highlightbackground=C_DARK_BADGE_BD,
                                       highlightcolor=C_DARK_BADGE_BD)
        agent_name = self._session.agent_name if self._session else cfg.agent_name
        project_name = self._session.project if self._session else cfg.project_name
        work_mode = self._session.work_mode if self._session else cfg.work_mode
        self._agent_lbl.config(text=f"{agent_name or '—'}")
        self._mode_lbl.config(text=f"{project_name or '—'} • {work_mode or '—'}")

    # ── Recovery ─────────────────────────────────────────────────────────────
    def _check_recovery(self):
        state = StateManager.load()
        if state is None:
            return
        if state.get("clock_out"):
            StateManager.delete()
            return

        def _new():
            self._write_abandoned_log(state)
            StateManager.delete()

        def _resume():
            self._restore_from_state(state)

        RecoveryDialog(self, state, _new, _resume)

    def _write_abandoned_log(self, state: Dict):
        ci = state.get("clock_in", "")
        try:
            ci_dt = datetime.strptime(ci, "%Y-%m-%d %H:%M:%S").replace(
                tzinfo=_TZ_BOLIVIA)
        except Exception:
            ci_dt = _now_bolivia()
        clock_out_snapshot = _capture_connectivity_snapshot()
        sess = SessionData(
            session_id   = state["session_id"],
            agent_id     = state["agent_id"],
            agent_name   = state["agent_name"],
            project      = state["project"],
            work_mode    = state["work_mode"],
            session_type = state["session_type"],
            clock_in     = ci_dt,
            clock_out    = _now_bolivia(),
            close_reason = "abandoned_at_recovery",
            clock_in_connection_type = state.get("clock_in_connection_type", ""),
            clock_in_network_name = state.get("clock_in_network_name", ""),
            clock_in_internet = bool(state.get("clock_in_internet", False)),
            clock_in_schedule_lookup = state.get("clock_in_schedule_lookup", ""),
            clock_in_lookup_message = state.get("clock_in_lookup_message", ""),
            clock_in_work_mode_source = state.get("clock_in_work_mode_source", "config_json"),
            clock_out_connection_type = clock_out_snapshot.get("connection_type", "none"),
            clock_out_network_name = clock_out_snapshot.get("network_name", ""),
            clock_out_internet = bool(clock_out_snapshot.get("internet_access", False)),
        )
        SessionLogger.write_session_log(sess, "00:00:00", "00:00:00", "00:00:00", "00:00:00", "00:00:00")

    def _restore_from_state(self, state: Dict):
        def _parse_dt(s):
            if not s:
                return None
            try:
                return datetime.strptime(s, "%Y-%m-%d %H:%M:%S").replace(
                    tzinfo=_TZ_BOLIVIA)
            except Exception:
                return None

        sess = SessionData(
            session_id     = state["session_id"],
            agent_id       = state["agent_id"],
            agent_name     = state["agent_name"],
            project        = state["project"],
            work_mode      = state["work_mode"],
            session_type   = state["session_type"],
            status         = state.get("status", "ONLINE"),
            clock_in       = _parse_dt(state.get("clock_in")),
            lunch_start    = _parse_dt(state.get("lunch_start")),
            lunch_end      = _parse_dt(state.get("lunch_end")),
            lunch_sec      = float(state.get("lunch_sec", 0)),
            meeting_sec    = float(state.get("meeting_sec", 0)),
            total_samples  = int(state.get("total_samples", 0)),
            active_samples = int(state.get("active_samples", 0)),
            inactive_samples_effective = int(state.get("inactive_samples_effective", 0)),
            app_used       = state.get("app_used", ""),
            top_app        = state.get("top_app", ""),
            clock_in_connection_type = state.get("clock_in_connection_type", ""),
            clock_in_network_name = state.get("clock_in_network_name", ""),
            clock_in_internet = bool(state.get("clock_in_internet", False)),
            clock_in_schedule_lookup = state.get("clock_in_schedule_lookup", ""),
            clock_in_lookup_message = state.get("clock_in_lookup_message", ""),
            clock_in_work_mode_source = state.get("clock_in_work_mode_source", "config_json"),
        )
        self._session = sess
        self._start_worker()
        self._refresh_status_visuals()
        self._update_session_header()
        self._refresh_action_buttons()
        self._start_tick()

    # ── Clock In ─────────────────────────────────────────────────────────────
    def _clock_in(self):
        with self._state_lock:
            if self._session:
                return
            dlg = SessionTypeDialog(self)
            if not dlg.session_type:
                return
            now = _now_bolivia()
            cfg = self.cfg_mgr.cfg
            sid = _format_short_session_id(now)

            clock_in_snapshot = _capture_connectivity_snapshot()
            lookup = self.sheets.lookup_agent_schedule(cfg.agent_id)
            lookup_status = str(lookup.get("status", "error") or "error")
            lookup_message = str(lookup.get("message", "") or "")
            effective_work_mode = cfg.work_mode
            work_mode_source = "config_json"
            alert_message = ""

            if lookup_status == "ok":
                effective_work_mode = _normalize_work_mode(lookup.get("work_mode", "")) or cfg.work_mode
                work_mode_source = "sheet"
            else:
                if not clock_in_snapshot.get("internet_access", False):
                    work_mode_source = "local_fallback_no_internet"
                    alert_message = (
                        "No internet connection was detected during Clock In.\n\n"
                        "The session will continue with the local configuration from the imported file.\n"
                        "This event will be recorded in the daily summary."
                    )
                elif lookup_status == "agent_not_found":
                    work_mode_source = "local_fallback_agent_not_found"
                    alert_message = (
                        "This AGENT_ID was not found in AGENTS_SCHEDULE.\n\n"
                        "The session will continue with the local configuration.\n"
                        "Please notify your supervisor because the schedule lookup failed."
                    )
                else:
                    work_mode_source = f"local_fallback_{lookup_status}"
                    alert_message = (
                        "The app could not validate your Work Mode from AGENTS_SCHEDULE.\n\n"
                        "The session will continue with the local configuration from the imported file.\n"
                        f"Reason: {lookup_message or lookup_status}"
                    )

            self._session = SessionData(
                session_id   = sid,
                agent_id     = cfg.agent_id,
                agent_name   = cfg.agent_name,
                project      = cfg.project_name,
                work_mode    = effective_work_mode,
                session_type = dlg.session_type,
                clock_in     = now,
                status       = "ONLINE",
                clock_in_connection_type = clock_in_snapshot.get("connection_type", "none"),
                clock_in_network_name = clock_in_snapshot.get("network_name", ""),
                clock_in_internet = bool(clock_in_snapshot.get("internet_access", False)),
                clock_in_schedule_lookup = lookup_status,
                clock_in_lookup_message = lookup_message,
                clock_in_work_mode_source = work_mode_source,
            )
            SessionLogger._ensure_dir(self._session)
            self._start_worker()

            shot_path = ScreenshotCapture.take(self._session, "session_start")
            LOGGER.info(
                "CLOCK IN | system_time=%s | bolivia_time=%s | agent=%s | session=%s | mode=%s | type=%s | internet=%s | lookup=%s | source=%s | screenshot=%s",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                _fmt(now),
                self._session.agent_id,
                self._session.session_id,
                self._session.work_mode,
                self._session.session_type,
                _bool_to_text(self._session.clock_in_internet),
                self._session.clock_in_schedule_lookup,
                self._session.clock_in_work_mode_source,
                "yes" if shot_path else "no",
            )

            clock_in_sync_ok = self.sheets.send_event("clock_in", self._session)
            self._refresh_status_visuals()
            self._update_session_header()
            self._refresh_action_buttons()
            self._start_tick()

            if not clock_in_sync_ok and self.sheets.last_error:
                messagebox.showerror(APP_TITLE, self.sheets.last_error)
            elif alert_message:
                messagebox.showwarning(APP_TITLE, alert_message)

    # ── Lunch / Meeting toggle ───────────────────────────────────────────────
    def _toggle_lunch(self):
        with self._state_lock:
            if not self._session:
                messagebox.showinfo(APP_TITLE, "Please Clock In first.")
                return
            if self._session.status == "LUNCH":
                dlg = LunchConfirmDialog(self, starting=False)
                if not dlg.confirmed:
                    return
                if self._worker.set_status("ONLINE"):
                    LOGGER.info("LUNCH END | agent=%s | session=%s | at=%s",
                                self._session.agent_id, self._session.session_id, _fmt(_now_bolivia()))
                    self.sheets.send_event("lunch_end", self._session)
                    self._refresh_status_visuals()
                    self._refresh_action_buttons()
                return
            if self._session.status == "MEETING":
                messagebox.showinfo(APP_TITLE, "End Meeting first before starting Lunch.")
                return
            if _has_taken_lunch(self._session):
                messagebox.showinfo(APP_TITLE, "Only one lunch is allowed per session.")
                return
            dlg = LunchConfirmDialog(self, starting=True)
            if not dlg.confirmed:
                return
            if self._worker.set_status("LUNCH"):
                LOGGER.info("LUNCH START | agent=%s | session=%s | at=%s",
                            self._session.agent_id, self._session.session_id, _fmt(_now_bolivia()))
                self.sheets.send_event("lunch_start", self._session)
                self._refresh_status_visuals()
                self._refresh_action_buttons()

    def _toggle_meeting(self):
        with self._state_lock:
            if not self._session:
                messagebox.showinfo(APP_TITLE, "Please Clock In first.")
                return
            if self._session.status == "LUNCH":
                messagebox.showinfo(APP_TITLE, "End Lunch first before starting Meeting.")
                return
            if self._session.status == "MEETING":
                if self._worker.set_status("ONLINE"):
                    LOGGER.info("MEETING END | agent=%s | session=%s | at=%s",
                                self._session.agent_id, self._session.session_id, _fmt(_now_bolivia()))
                    self.sheets.send_event("meeting_end", self._session)
                    self._refresh_status_visuals()
                    self._refresh_action_buttons()
            else:
                if self._worker.set_status("MEETING"):
                    LOGGER.info("MEETING START | agent=%s | session=%s | at=%s",
                                self._session.agent_id, self._session.session_id, _fmt(_now_bolivia()))
                    self.sheets.send_event("meeting_start", self._session)
                    self._refresh_status_visuals()
                    self._refresh_action_buttons()

    # ── Clock Out ────────────────────────────────────────────────────────────
    def _clock_out(self, force: bool = False):
        with self._state_lock:
            if not self._session:
                if not force:
                    messagebox.showinfo(APP_TITLE, "No active session.")
                return
            if not force and self._session.status in ("LUNCH", "MEETING"):
                messagebox.showinfo(APP_TITLE, f"End {self._session.status.title()} before Clock Out.")
                return

            now = _now_bolivia()
            ci  = self._session.clock_in or now
            clock_out_snapshot = _capture_connectivity_snapshot()

            if self._session.status == "LUNCH" and self._session.lunch_start:
                self._session.lunch_sec += (now - self._session.lunch_start).total_seconds()
                self._session.lunch_end = now
            if self._session.status == "MEETING":
                if hasattr(self._worker, "_meeting_start") and self._worker._meeting_start:
                    self._session.meeting_sec += (
                        now - self._worker._meeting_start).total_seconds()

            total_sec = (now - ci).total_seconds()
            net_sec   = max(0, total_sec - self._session.lunch_sec)

            if not force:
                dlg = ClockOutConfirmDialog(self, self._session, net_sec)
                if not dlg.confirmed:
                    return
                self._session.overtime = dlg.ot_reported
                self._session.overtime_requested_by = dlg.ot_requested_by
                self._session.ot_note = dlg.ot_note
                self._session.close_reason = ""
            else:
                self._session.overtime = False
                self._session.overtime_requested_by = ""
                self._session.ot_note = ""
                self._session.close_reason = "window_close_auto_clockout"

            self._session.payable_worked_sec = min(net_sec, REGULAR_WORKDAY_SEC)
            self._session.overtime_duration_sec = (
                max(0.0, net_sec - REGULAR_WORKDAY_SEC) if self._session.overtime else 0.0
            )

            self._session.clock_out = now
            self._session.status    = "IDLE"
            self._session.clock_out_connection_type = clock_out_snapshot.get("connection_type", "none")
            self._session.clock_out_network_name = clock_out_snapshot.get("network_name", "")
            self._session.clock_out_internet = bool(clock_out_snapshot.get("internet_access", False))

            if self._worker:
                self._worker.stop()
                self._worker = None

            shot_path = ScreenshotCapture.take(self._session, "session_end")

            lunch_fmt     = _sec_to_hms(self._session.lunch_sec)
            meeting_fmt   = _sec_to_hms(self._session.meeting_sec)
            net_fmt       = _sec_to_hms(net_sec)
            payable_fmt   = _sec_to_hms(self._session.payable_worked_sec)
            overtime_fmt  = _sec_to_hms(self._session.overtime_duration_sec)
            inactive_samples = max(0, self._session.inactive_samples_effective)
            inactive_sec = inactive_samples * (WORK_MODES[self._session.work_mode]["interval_min"] * 60)
            inactive_fmt = _sec_to_hms(inactive_sec)

            SessionLogger.write_summary_day(self._session, lunch_fmt, meeting_fmt, net_fmt, payable_fmt, overtime_fmt)
            SessionLogger.write_session_log(self._session, lunch_fmt, meeting_fmt, net_fmt, payable_fmt, overtime_fmt)
            StateManager.delete()
            LOGGER.info(
                "%s | system_time=%s | bolivia_time=%s | agent=%s | session=%s | worked=%s | internet=%s | screenshot=%s",
                "CLOCK OUT AUTO" if force else "CLOCK OUT",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                _fmt(now),
                self._session.agent_id,
                self._session.session_id,
                net_fmt,
                _bool_to_text(self._session.clock_out_internet),
                "yes" if shot_path else "no",
            )
            self.sheets.send_event("clock_out", self._session)

            if self._tick_job:
                self.after_cancel(self._tick_job)
                self._tick_job = None

            self._refresh_status_visuals()
            self._online_total_lbl.config(text="00:00:00")
            self._lunch_total_lbl.config(text="00:00:00")
            self._meeting_total_lbl.config(text="00:00:00")

            if not force:
                SessionClosedDialog(
                    self,
                    session=self._session,
                    clock_in_fmt=_fmt_time(self._session.clock_in),
                    clock_out_fmt=_fmt_time(now),
                    worked_fmt=net_fmt,
                    lunch_fmt=lunch_fmt,
                    inactive_fmt=inactive_fmt,
                )

            self._session = None
            self._update_session_header()

    # ── Background worker & tick ─────────────────────────────────────────────
    def _start_worker(self):
        self._worker = SampleWorker(self._session, self.cfg_mgr.cfg)
        self._worker.start()
        self._refresh_action_buttons()

    def _start_tick(self):
        self._tick_job = self.after(1000, self._update_totals)

    def _refresh_action_buttons(self):
        status = self._session.status if self._session else "IDLE"
        has_session = self._session is not None
        has_lunch = _has_taken_lunch(self._session) if self._session else False

        self._set_button_visual(self._btn_ci, enabled=not has_session)
        self._set_button_visual(self._btn_co, enabled=(has_session and status == "ONLINE"))

        if not has_session:
            self._set_button_visual(self._btn_ln, enabled=False, text="Lunch")
            self._set_button_visual(self._btn_mt, enabled=False, text="Meeting")
            return

        lunch_enabled = status in ("ONLINE", "LUNCH") and (status == "LUNCH" or not has_lunch)
        meeting_enabled = status in ("ONLINE", "MEETING")

        self._set_button_visual(
            self._btn_ln,
            enabled=lunch_enabled,
            selected=(status == "LUNCH"),
            text=("End Lunch" if status == "LUNCH" else "Lunch"),
        )
        self._set_button_visual(
            self._btn_mt,
            enabled=meeting_enabled,
            selected=(status == "MEETING"),
            text=("End Meeting" if status == "MEETING" else "Meeting"),
        )

    def _update_totals(self):
        if not self._session or not self._session.clock_in:
            return
        now   = _now_bolivia()
        ci    = self._session.clock_in
        total = (now - ci).total_seconds()

        lunch_live = self._session.lunch_sec
        if self._session.status == "LUNCH" and self._session.lunch_start:
            lunch_live += (now - self._session.lunch_start).total_seconds()

        meeting_live = self._session.meeting_sec
        if (self._session.status == "MEETING" and
                self._worker and hasattr(self._worker, "_meeting_start") and
                self._worker._meeting_start):
            meeting_live += (now - self._worker._meeting_start).total_seconds()

        online = max(0, total - lunch_live)
        self._online_total_lbl.config(text=_sec_to_hms(online))
        self._lunch_total_lbl.config(text=_sec_to_hms(lunch_live))
        self._meeting_total_lbl.config(text=_sec_to_hms(meeting_live))

        # Update elapsed time label under big status
        status = self._session.status
        if status == "ONLINE":
            self._status_elapsed_lbl.config(text=_sec_to_hms(online))
        elif status == "LUNCH":
            self._status_elapsed_lbl.config(text=_sec_to_hms(lunch_live))
        elif status == "MEETING":
            self._status_elapsed_lbl.config(text=_sec_to_hms(meeting_live))
        else:
            self._status_elapsed_lbl.config(text="")

        self._refresh_status_visuals()
        self._tick_job = self.after(1000, self._update_totals)

    def _refresh_status_visuals(self):
        status = self._session.status if self._session else "IDLE"
        pulse_color = C_ACCENT_SOFT if self._pulse_phase else C_ACCENT

        if status == "ONLINE":
            self._status_lbl.config(text="ONLINE", fg=C_STATUS_ONLINE)
            self._status_indicator.config(
                text="● ONLINE", fg=C_STATUS_ONLINE,
                highlightbackground=C_ACCENT_BORDER)
        elif status == "MEETING":
            self._status_lbl.config(text="MEETING", fg=pulse_color)
            self._status_indicator.config(
                text="● MEETING", fg=C_ACCENT,
                highlightbackground=C_ACCENT_BORDER)
        elif status == "LUNCH":
            self._status_lbl.config(text="LUNCH", fg=C_TEXT)
            self._status_indicator.config(
                text="● LUNCH", fg=C_TEXT_SECONDARY,
                highlightbackground=C_BORDER)
        else:
            self._status_lbl.config(text="IDLE", fg=C_STATUS_IDLE)
            self._status_indicator.config(
                text="● IDLE", fg=C_STATUS_IDLE,
                highlightbackground=C_BORDER)

        # Mini-card borders: highlight active card
        active_border = C_ACCENT_BORDER
        inactive_border = C_BORDER
        self._card_online.config(
            highlightbackground=active_border if status == "ONLINE" else inactive_border)
        self._card_lunch.config(
            highlightbackground=active_border if status == "LUNCH" else inactive_border)
        self._card_meeting.config(
            highlightbackground=active_border if status == "MEETING" else inactive_border)

        self._online_total_caption.config(fg=C_MUTED)
        self._meeting_total_caption.config(fg=C_MUTED)
        self._lunch_total_caption.config(fg=C_MUTED)

        self._online_total_lbl.config(
            fg=C_TEXT if self._online_total_lbl.cget("text") != "00:00:00" else C_DISABLED_FG
        )
        self._meeting_total_lbl.config(
            fg=C_TEXT if self._meeting_total_lbl.cget("text") != "00:00:00" else C_DISABLED_FG
        )
        self._lunch_total_lbl.config(
            fg=C_TEXT if self._lunch_total_lbl.cget("text") != "00:00:00" else C_DISABLED_FG
        )

    def _schedule_status_pulse(self):
        self._pulse_phase = not self._pulse_phase
        self._refresh_status_visuals()
        self._pulse_job = self.after(850, self._schedule_status_pulse)

    # ── Helpers ──────────────────────────────────────────────────────────────
    def _open_export(self):
        ExportDialog(self, self.cfg_mgr)

    def _safe_back(self):
        if self._session:
            messagebox.showwarning(APP_TITLE,
                "Please Clock Out before going back.")
            return
        if self._tick_job:
            self.after_cancel(self._tick_job)
        if self._pulse_job:
            self.after_cancel(self._pulse_job)
            self._pulse_job = None
        self.navigate("start")

    def force_clockout_if_active(self):
        """Called by main window when user closes via X."""
        if self._session:
            self._clock_out(force=True)


# =============================================================================
# UI — MAIN APPLICATION WINDOW
# =============================================================================

class App(tk.Tk):

    def __init__(self):
        _apply_windows_app_id()
        super().__init__()
        _enable_dpi_awareness()
        self.title(APP_TITLE)
        self.geometry(f"{WIN_W}x{WIN_H}")
        self.resizable(True, True)

        _apply_window_icon(self)

        _center_window(self, WIN_W, WIN_H)
        self.configure(bg=C_BG)
        _apply_theme(self)

        self._cfg_mgr = ConfigManager()
        self._sheets  = SheetsSync(self._cfg_mgr.cfg)

        self._frames: Dict[str, tk.Frame] = {}
        self._agent_view: Optional[AgentView] = None

        self.withdraw()
        self._show_splash()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _show_splash(self):
        SplashScreen(self, self._after_splash)

    def _after_splash(self):
        self.deiconify()
        self._navigate("start")

    def _navigate(self, target: str):
        # Prevent navigation if there is an active session in AgentView
        if self._agent_view and self._agent_view._session is not None:
            messagebox.showwarning(
                APP_TITLE,
                "Cannot navigate while a session is active.\nPlease clock out first."
            )
            return

        for f in self._frames.values():
            f.destroy()
        self._frames.clear()
        self._agent_view = None

        frame: tk.Frame
        if target == "start":
            frame = StartView(self, self._cfg_mgr, self._navigate)
        elif target == "admin_login":
            frame = AdminLoginView(self, self._cfg_mgr, self._navigate)
        elif target == "admin_console":
            self._cfg_mgr.load()
            self._sheets.refresh_config(self._cfg_mgr.cfg)
            frame = AdminConsoleView(self, self._cfg_mgr, self._navigate)
        elif target == "agent":
            self._cfg_mgr.load()
            self._sheets.refresh_config(self._cfg_mgr.cfg)
            frame = AgentView(self, self._cfg_mgr, self._sheets, self._navigate)
            self._agent_view = frame
            self.geometry(f"425x525")
            _center_window(self, 425, 525)
        else:
            return

        frame.place(relwidth=1, relheight=1)
        self._frames[target] = frame

    def _on_close(self):
        dlg = AppCloseConfirmDialog(self, has_active_session=bool(self._agent_view and self._agent_view._session))
        if not dlg.confirmed:
            return
        if self._agent_view:
            self._agent_view.force_clockout_if_active()
        self.destroy()

# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()