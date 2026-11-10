# -*- coding: utf-8 -*-
"""
===============================================================================
  OFFICE MANAGEMENT TOOL of <your CA firm>  (single-file Windows application)
  The firm name is taken from Invoicing > Firm Master, e.g.
  "Office Management Tool of Suthar & Suthar".
===============================================================================
  Modules
    1. Chat & Communication : team chat, group chat, file sharing,
                              voice call (1-to-1) and screen sharing
    2. Attendance & Leave   : check-in / check-out, leave application,
                              2 paid leaves per month, leave balance,
                              salary slip (PDF)
    3. Reimbursement        : Conveyance, Printing, Xerox, Office Purchases,
                              Other Expenses; month-end PDF + e-mail to Manager
    4. Litigation Tracker   : Income Tax, GST, VAT, Service Tax cases;
                              hearing / submission-wise log & report PDF,
                              desktop notifications, hearing calendar (+ .ics
                              export for Outlook/Google), fees, GST invoices,
                              receipts with TDS, out-of-pocket expenses
    5. Salary + Reimbursement payout : month-end statement (net salary +
                              approved reimbursement = total payout), PDF/CSV
    6. Task Tracker         : anyone can assign tasks to anyone, due date &
                              urgency priority, colour codes (green > 7 days,
                              yellow 4-7 days, red <= 3 days / overdue),
                              notifications, challenges remarks, billing on
                              completion -> auto entry in Billing tab,
                              automatic recurring tasks (GST / TDS returns)

  How it works on a LAN
    * ONE computer (normally the Manager's) ticks "Host server on this PC"
      while logging in. That PC keeps the central database.
    * Every other team member opens the same file, clicks "Find Server"
      (or types the host PC's IP address) and logs in.
    * First login on the host PC:  username = admin   password = admin@123
      (change it immediately from Settings).

  Run : open this file in IDLE and press F5   (or:  python suthar_suthar_app.py)
  Headless server only (optional):  python suthar_suthar_app.py --server

  Manual install (only if auto-install fails):
      pip install PyQt5 reportlab sounddevice
===============================================================================
"""

# =============================================================================
# 1. AUTO-INSTALL OF REQUIRED LIBRARIES  (runs before any third-party import)
# =============================================================================
import sys
import os
import subprocess
import importlib
import importlib.util
import site

# (import name, pip name, mandatory?)
REQUIRED_PACKAGES = [
    ("PyQt5", "PyQt5", True),          # user interface
    ("reportlab", "reportlab", True),  # PDF salary slip / claim statement
    ("sounddevice", "sounddevice", False),  # voice call (optional)
]


def _python_cmd():
    """Return python.exe (not pythonw.exe) so pip works when launched from IDLE."""
    exe = sys.executable or "python"
    if os.name == "nt" and exe.lower().endswith("pythonw.exe"):
        cand = exe[:-len("pythonw.exe")] + "python.exe"
        if os.path.exists(cand):
            exe = cand
    return exe


def _run(args):
    flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    try:
        return subprocess.run(args, capture_output=True, text=True, creationflags=flags)
    except Exception as exc:  # pragma: no cover
        class R:  # minimal stand-in result
            returncode, stdout, stderr = 1, "", str(exc)
        return R()


def ensure_packages():
    missing = [(m, p, req) for m, p, req in REQUIRED_PACKAGES
               if importlib.util.find_spec(m) is None]
    if not missing:
        return
    py = _python_cmd()
    print("Suthar & Suthar: installing missing libraries ->",
          ", ".join(p for _, p, _ in missing), "\n(please wait, first run only)...")
    if importlib.util.find_spec("pip") is None:
        _run([py, "-m", "ensurepip", "--upgrade"])
    for mod, pkg, req in missing:
        r = _run([py, "-m", "pip", "install", "--disable-pip-version-check", pkg])
        if r.returncode != 0:  # no admin rights -> try per-user install
            r = _run([py, "-m", "pip", "install", "--user", "--disable-pip-version-check", pkg])
        print(("  [OK]   " if r.returncode == 0 else "  [FAIL] ") + pkg)
        if r.returncode != 0:
            print((r.stderr or "")[-600:])
    importlib.invalidate_caches()
    try:
        usp = site.getusersitepackages()
        if os.path.isdir(usp) and usp not in sys.path:
            site.addsitedir(usp)
    except Exception:
        pass
    still = [p for m, p, req in missing if req and importlib.util.find_spec(m) is None]
    if still:
        print("\nCould not install:", ", ".join(still))
        print("Please run manually in Command Prompt:\n    pip install PyQt5 reportlab sounddevice")
        raise SystemExit(1)


ensure_packages()

# =============================================================================
# 2. IMPORTS
# =============================================================================
import json
import socket
import struct
import threading
import queue
import sqlite3
import hashlib
import hmac
import secrets
import re
import time
import calendar
import html
import uuid
import smtplib
import csv
import ssl
from datetime import datetime, date, timedelta
from email.message import EmailMessage

from PyQt5.QtCore import (Qt, QObject, pyqtSignal, QThread, QTimer, QDate, QUrl,
                          QByteArray, QBuffer, QIODevice)
from PyQt5.QtGui import QPixmap, QIcon, QColor, QPainter, QDesktopServices, QFont, QTextCharFormat
from PyQt5.QtWidgets import (
    QSystemTrayIcon, QMenu, QCalendarWidget, QTabWidget,
    QApplication, QMainWindow, QWidget, QDialog, QLabel, QPushButton, QLineEdit,
    QComboBox, QDateEdit, QDoubleSpinBox, QSpinBox, QCheckBox, QTextBrowser,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout,
    QGroupBox, QStackedWidget, QFrame, QFileDialog, QMessageBox, QSplitter,
    QSizePolicy, QScrollArea, QTextEdit, QButtonGroup)

try:
    import sounddevice as sd
    AUDIO_OK = True
except Exception:  # library missing or PortAudio not found
    sd = None
    AUDIO_OK = False

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors as rl_colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

# =============================================================================
# 3. CONSTANTS & SMALL HELPERS
# =============================================================================
PRODUCT_NAME = "Office Management Tool"
FIRM_DEFAULT = "Suthar & Suthar"            # used only until the Firm Master is filled in
FIRM_NAME = FIRM_DEFAULT
APP_NAME = f"{PRODUCT_NAME} of {FIRM_NAME}"


def set_firm(name):
    """Update the firm name used everywhere: 'Office Management Tool of <firm>'."""
    global FIRM_NAME, APP_NAME
    FIRM_NAME = (name or "").strip() or FIRM_DEFAULT
    APP_NAME = f"{PRODUCT_NAME} of {FIRM_NAME}"
    return APP_NAME
APP_VERSION = "1.0"
PORT = 50505
DISCOVERY_PORT = 50506
MAX_FILE = 200 * 1024 * 1024          # 200 MB per shared file
PAID_LEAVE_PER_MONTH = 2              # firm policy
WEEKLY_OFF = {6}                      # Sunday (Mon=0 ... Sun=6)
LEAVE_TYPES = ["Casual Leave", "Sick Leave", "Personal Leave", "Emergency Leave"]
REIMB_CATEGORIES = ["Conveyance", "Printing", "Xerox", "Office Purchases", "Other Expenses"]
DEFAULT_ADMIN = ("admin", "admin@123")

# ---- Litigation tracker masters (forum lists are editable in the case form) ----
LIT_LAWS = ["Income Tax", "GST", "VAT", "Service Tax"]
LIT_FORUMS = {
    "Income Tax": ["Assessing Officer", "Faceless Assessment Unit / NaFAC", "CPC Bengaluru",
                   "Dispute Resolution Panel (DRP)", "JCIT (Appeals)", "CIT (Appeals) / NFAC",
                   "ITAT", "High Court", "Supreme Court", "Rectification / Revision (PCIT)"],
    "GST": ["Proper Officer / Adjudicating Authority", "Appellate Authority (Appeals)",
            "GST Appellate Tribunal (GSTAT)", "AAR / AAAR", "High Court", "Supreme Court"],
    "VAT": ["Assessing Authority", "Appellate Authority (DC Appeals)", "VAT Tribunal / Tax Board",
            "High Court", "Supreme Court"],
    "Service Tax": ["Adjudicating Authority", "Commissioner (Appeals)", "CESTAT", "High Court",
                    "Supreme Court"],
}
LIT_STATUSES = ["Open", "Reply / Submission Pending", "Under Appeal", "Stay Granted",
                "Closed - Favourable", "Closed - Partly Allowed", "Closed - Unfavourable", "Withdrawn"]
LIT_OPEN = LIT_STATUSES[:4]
LIT_PTYPES = ["Notice Received", "Submission Filed", "Hearing Attended", "Adjournment",
              "Order Received", "Appeal Filed", "Other Update"]
LAW_COLORS = {"Income Tax": "#1e88e5", "GST": "#e91e63", "VAT": "#11998e", "Service Tax": "#ff6a00"}
PAY_MODES = ["NEFT / RTGS / IMPS", "UPI", "Cheque", "Cash"]

# ---- Task tracker masters ----
TASK_REPEATS = ["None", "Monthly", "Quarterly", "Half-Yearly", "Yearly"]
REPEAT_MONTHS = {"Monthly": 1, "Quarterly": 3, "Half-Yearly": 6, "Yearly": 12}
BILL_OPTIONS = ["Billing Due", "Billing Raised", "Yearly Billing", "No Billing"]
BILL_STATUSES = ["To be Billed", "Billed", "Received", "Written Off"]
TASK_PRESETS = ["GSTR-1 filing", "GSTR-3B filing", "GSTR-9 / 9C annual return", "TDS return (24Q / 26Q)",
                "TDS payment", "Advance tax computation", "Income tax return (ITR)", "Tax audit (Form 3CD)",
                "Statutory audit", "ROC filing (AOC-4 / MGT-7)", "Accounting & bookkeeping",
                "Reply to notice", "Client meeting"]
TASK_GREEN, TASK_YELLOW, TASK_RED, TASK_GREY = "#c8f7d4", "#fff3b0", "#ffcdd2", "#eceff1"
INVOICE_OVERDUE_DAYS = 30
# GST registration of the CA firm (Firm Master):
#   Regular      -> GST charged: CGST 9% + SGST 9% (intra-state) or IGST 18% (inter-state)  -> "TAX INVOICE"
#   Composition  -> no GST can be collected from client                                    -> "BILL OF SUPPLY"
#   Unregistered -> no GST charged                                                         -> "INVOICE"
GST_TYPES = ["Regular", "Composition", "Unregistered"]
FU_MODES = ["Phone call", "E-mail", "WhatsApp", "Personal visit", "SMS", "Other"]
FU_GAP_DAYS = 7          # default gap before the next payment follow-up
GST_DOC_TITLE = {"Regular": "TAX INVOICE", "Composition": "BILL OF SUPPLY", "Unregistered": "INVOICE"}

# GST state codes (first two digits of GSTIN) - used for place of supply / CGST+SGST vs IGST
STATE_CODES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh", "05": "Uttarakhand",
    "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura",
    "17": "Meghalaya", "18": "Assam", "19": "West Bengal", "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh",
    "23": "Madhya Pradesh", "24": "Gujarat", "26": "Dadra & Nagar Haveli and Daman & Diu", "27": "Maharashtra",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry",
    "35": "Andaman & Nicobar Islands", "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh",
    "97": "Other Territory"}


def state_of(gstin):
    code = str(gstin or "")[:2]
    return f"{STATE_CODES[code]} ({code})" if code in STATE_CODES else ""


def add_months(d, n, day=None):
    """Add n months to date d, keeping 'day' (clamped to month length)."""
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return date(y, m, min(day or d.day, calendar.monthrange(y, m)[1]))


def task_colour(days_left, status):
    """Green > 7 days, Yellow 4-7 days, Red <= 3 days or overdue."""
    if status in ("Completed", "Cancelled"):
        return TASK_GREY, "Done"
    if days_left <= 3:
        return TASK_RED, "High"
    if days_left <= 7:
        return TASK_YELLOW, "Medium"
    return TASK_GREEN, "Low"

if os.name == "nt":
    DATA_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "SutharSuthar")
else:
    DATA_DIR = os.path.join(os.path.expanduser("~"), ".sutharsuthar")
HOME = os.path.expanduser("~")
REPORT_DIR = os.path.join(HOME, "Documents", "Suthar & Suthar Reports")
DOWNLOAD_DIR = os.path.join(HOME, "Downloads", "Suthar & Suthar")
CONFIG_FILE = os.path.join(DATA_DIR, "client_config.json")


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def month_bounds(y, m):
    return date(y, m, 1), date(y, m, calendar.monthrange(y, m)[1])


def working_days(d1, d2):
    """Count days between d1..d2 (inclusive) excluding weekly-off."""
    n, d = 0, d1
    while d <= d2:
        if d.weekday() not in WEEKLY_OFF:
            n += 1
        d += timedelta(days=1)
    return n


def inr(x):
    """Indian digit grouping: 12,34,567.89"""
    try:
        x = float(x or 0)
    except Exception:
        x = 0.0
    neg = x < 0
    s = f"{abs(round(x, 2)):.2f}"
    ip, dp = s.split(".")
    if len(ip) > 3:
        last3, rest = ip[-3:], ip[:-3]
        groups = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        ip = ",".join(groups) + "," + last3
    return ("-" if neg else "") + ip + "." + dp


_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
         "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen",
         "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _words(n):
    if n < 20:
        return _ONES[n]
    if n < 100:
        return _TENS[n // 10] + ((" " + _ONES[n % 10]) if n % 10 else "")
    if n < 1000:
        return _ONES[n // 100] + " Hundred" + ((" " + _words(n % 100)) if n % 100 else "")
    for div, name in ((10 ** 7, "Crore"), (10 ** 5, "Lakh"), (1000, "Thousand")):
        if n >= div:
            return _words(n // div) + " " + name + ((" " + _words(n % div)) if n % div else "")
    return ""


def amount_in_words(x):
    x = round(float(x or 0), 2)
    rupees = int(x)
    paise = int(round((x - rupees) * 100))
    s = "Rupees " + (_words(rupees) if rupees else "Zero")
    if paise:
        s += " and " + _words(paise) + " Paise"
    return s + " Only"


def dmy_s(iso):
    try:
        return date.fromisoformat(str(iso)[:10]).strftime("%d-%m-%Y")
    except Exception:
        return str(iso or "")


def safe_filename(name):
    name = os.path.basename(str(name or "file"))
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name).strip(" .")
    return (name or "file")[:150]


def local_ips():
    ips = set()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ips.add(s.getsockname()[0])
        s.close()
    except Exception:
        pass
    try:
        for ip in socket.gethostbyname_ex(socket.gethostname())[2]:
            if not ip.startswith("127."):
                ips.add(ip)
    except Exception:
        pass
    return sorted(ips) or ["127.0.0.1"]


def invoice_folder(cfg):
    """Dedicated invoice folder for this PC (set in Settings / Invoicing). Default: Reports\\Invoices."""
    return (cfg or {}).get("invoice_dir") or os.path.join(REPORT_DIR, "Invoices")


def invoice_filename(inv):
    """Name of client_Invoice number.pdf  e.g.  Sarwan S Gupta_SS-2026-27-0001.pdf"""
    client = re.sub(r"\s+", " ", str(inv.get("client_name") or "Client")).strip()
    client = client.replace("/", "-").replace("\\", "-")      # M/s ABC -> M-s ABC (no folder separators)
    return safe_filename(f"{client}_{str(inv['inv_no']).replace('/', '-')}") + ".pdf"


def open_path(path):
    QDesktopServices.openUrl(QUrl.fromLocalFile(path))


def load_cfg():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_cfg(cfg):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# =============================================================================
# 4. NETWORK FRAMING  (4-byte header length + 4-byte payload length + JSON + bytes)
# =============================================================================
def send_frame(sock, header, payload=b""):
    h = json.dumps(header).encode("utf-8")
    sock.sendall(struct.pack("!II", len(h), len(payload)) + h)
    if payload:
        sock.sendall(payload)


def recv_exact(sock, n):
    buf = bytearray(n)
    view = memoryview(buf)
    pos = 0
    while pos < n:
        r = sock.recv_into(view[pos:], min(n - pos, 1 << 20))
        if r == 0:
            raise ConnectionError("connection closed")
        pos += r
    return bytes(buf)


def recv_frame(sock):
    hl, pl = struct.unpack("!II", recv_exact(sock, 8))
    if hl > 2 * 1024 * 1024 or pl > MAX_FILE + 1024 * 1024:
        raise ConnectionError("frame too large")
    header = json.loads(recv_exact(sock, hl).decode("utf-8"))
    payload = recv_exact(sock, pl) if pl else b""
    return header, payload


# =============================================================================
# 5. SERVER : DATABASE
# =============================================================================
SCHEMA = """
CREATE TABLE IF NOT EXISTS users(
    username TEXT PRIMARY KEY, full_name TEXT, pw_hash TEXT, salt TEXT,
    role TEXT DEFAULT 'member', designation TEXT DEFAULT '', department TEXT DEFAULT '',
    email TEXT DEFAULT '', salary REAL DEFAULT 0, pan TEXT DEFAULT '', bank TEXT DEFAULT '',
    doj TEXT DEFAULT '', pt REAL DEFAULT 0, tds REAL DEFAULT 0, active INTEGER DEFAULT 1);
CREATE TABLE IF NOT EXISTS messages(
    id INTEGER PRIMARY KEY AUTOINCREMENT, sender TEXT, recipient TEXT, text TEXT,
    file_id INTEGER, ts TEXT);
CREATE TABLE IF NOT EXISTS files(
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, stored TEXT, size INTEGER,
    uploader TEXT, recipient TEXT, purpose TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS attendance(
    id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, day TEXT, check_in TEXT,
    check_out TEXT DEFAULT '', mode TEXT DEFAULT 'Office', UNIQUE(username, day));
CREATE TABLE IF NOT EXISTS leaves(
    id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, from_date TEXT, to_date TEXT,
    half_day INTEGER DEFAULT 0, days REAL, ltype TEXT, reason TEXT,
    status TEXT DEFAULT 'Pending', applied_on TEXT, decided_by TEXT DEFAULT '',
    remark TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS reimb(
    id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, day TEXT, category TEXT,
    description TEXT, amount REAL, file_id INTEGER, status TEXT DEFAULT 'Pending',
    submitted_on TEXT, decided_by TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS cases(
    id INTEGER PRIMARY KEY AUTOINCREMENT, law TEXT, client_name TEXT, client_pan TEXT DEFAULT '',
    client_gstin TEXT DEFAULT '', client_address TEXT DEFAULT '', client_email TEXT DEFAULT '',
    period TEXT DEFAULT '', ref_no TEXT DEFAULT '', forum TEXT DEFAULT '', section TEXT DEFAULT '',
    nature TEXT DEFAULT '', demand REAL DEFAULT 0, assigned_to TEXT DEFAULT '', status TEXT DEFAULT 'Open',
    next_hearing TEXT DEFAULT '', next_purpose TEXT DEFAULT '', remarks TEXT DEFAULT '',
    created_by TEXT, created_on TEXT, updated_on TEXT);
CREATE TABLE IF NOT EXISTS proceedings(
    id INTEGER PRIMARY KEY AUTOINCREMENT, case_id INTEGER, pdate TEXT, ptype TEXT, details TEXT,
    outcome TEXT DEFAULT '', next_date TEXT DEFAULT '', file_id INTEGER, by_user TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS fees(
    id INTEGER PRIMARY KEY AUTOINCREMENT, case_id INTEGER, fdate TEXT, description TEXT, amount REAL,
    invoice_id INTEGER, by_user TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS ope(
    id INTEGER PRIMARY KEY AUTOINCREMENT, case_id INTEGER, edate TEXT, description TEXT, amount REAL,
    billable INTEGER DEFAULT 1, invoice_id INTEGER, by_user TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS invoices(
    id INTEGER PRIMARY KEY AUTOINCREMENT, inv_no TEXT UNIQUE, case_id INTEGER, inv_date TEXT,
    fee_total REAL, ope_total REAL, pure_agent INTEGER, gst_rate REAL, gst_mode TEXT, cgst REAL,
    sgst REAL, igst REAL, total REAL, received REAL DEFAULT 0, tds REAL DEFAULT 0,
    status TEXT DEFAULT 'Unpaid', created_by TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS receipts(
    id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id INTEGER, rdate TEXT, amount REAL, tds REAL,
    mode TEXT, ref TEXT, by_user TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS tasks(
    id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, description TEXT DEFAULT '', client TEXT DEFAULT '',
    assigned_to TEXT, created_by TEXT, due_date TEXT, urgent INTEGER DEFAULT 0, status TEXT DEFAULT 'Open',
    challenges TEXT DEFAULT '', billing_option TEXT DEFAULT '', billing_amount REAL DEFAULT 0,
    series_id INTEGER, created_on TEXT, completed_on TEXT DEFAULT '', completed_by TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS task_series(
    id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, description TEXT, client TEXT, assigned_to TEXT,
    created_by TEXT, frequency TEXT, day INTEGER, urgent INTEGER DEFAULT 0, next_due TEXT,
    active INTEGER DEFAULT 1, created_on TEXT);
CREATE TABLE IF NOT EXISTS billing(
    id INTEGER PRIMARY KEY AUTOINCREMENT, task_id INTEGER, client TEXT, description TEXT, amount REAL,
    btype TEXT, status TEXT, ref TEXT DEFAULT '', created_on TEXT, updated_by TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS clients(
    id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, gstin TEXT DEFAULT '', pan TEXT DEFAULT '',
    address TEXT DEFAULT '', email TEXT DEFAULT '', phone TEXT DEFAULT '', contact TEXT DEFAULT '',
    created_on TEXT);
CREATE TABLE IF NOT EXISTS followups(
    id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id INTEGER, fdate TEXT, mode TEXT, remarks TEXT DEFAULT '',
    promise_date TEXT DEFAULT '', next_date TEXT DEFAULT '', by_user TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS item_access(
    ref_type TEXT, ref_id INTEGER, username TEXT, can_view INTEGER DEFAULT 1, can_fees INTEGER DEFAULT 0,
    PRIMARY KEY(ref_type, ref_id, username));
CREATE TABLE IF NOT EXISTS docs(
    id INTEGER PRIMARY KEY AUTOINCREMENT, ref_type TEXT, ref_id INTEGER, file_id INTEGER, title TEXT DEFAULT '',
    by_user TEXT, ts TEXT);
CREATE TABLE IF NOT EXISTS inv_lines(
    id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id INTEGER, description TEXT, sac TEXT, amount REAL,
    taxable INTEGER DEFAULT 1);
"""
# columns added in later versions (safe ALTER TABLE on existing databases)
MIGRATIONS = [("users", "phone", "TEXT DEFAULT ''"), ("cases", "est_fee", "REAL DEFAULT 0"),
              ("cases", "client_phone", "TEXT DEFAULT ''"), ("tasks", "est_amount", "REAL DEFAULT 0"),
              ("task_series", "est_amount", "REAL DEFAULT 0"), ("billing", "invoice_id", "INTEGER"),
              ("invoices", "source", "TEXT DEFAULT 'case'"), ("invoices", "billing_id", "INTEGER"),
              ("invoices", "client_name", "TEXT DEFAULT ''"), ("invoices", "client_gstin", "TEXT DEFAULT ''"),
              ("invoices", "client_pan", "TEXT DEFAULT ''"), ("invoices", "client_address", "TEXT DEFAULT ''"),
              ("invoices", "client_email", "TEXT DEFAULT ''"), ("invoices", "client_phone", "TEXT DEFAULT ''"),
              ("invoices", "notes", "TEXT DEFAULT ''"), ("receipts", "rcpt_no", "TEXT DEFAULT ''"),
              ("invoices", "service_details", "TEXT DEFAULT ''"), ("cases", "restricted", "INTEGER DEFAULT 0"),
              ("tasks", "hide_fees", "INTEGER DEFAULT 0"), ("users", "must_change", "INTEGER DEFAULT 0"),
              ("invoices", "gst_type", "TEXT DEFAULT 'Regular'")]
SETTING_KEYS = ("company_name", "company_address", "manager_email", "firm_gstin", "firm_pan",
                "firm_bank", "invoice_prefix", "gst_rate", "reminder_days", "firm_city", "firm_state",
                "firm_phone", "firm_email", "firm_website", "firm_frn", "bank_name", "bank_acc", "bank_ifsc",
                "bank_branch", "upi_id", "sac_code", "terms", "signatory", "firm_logo", "firm_tagline",
                "firm_sign", "sign_auto", "gst_type")
SETTING_DEFAULTS = {"company_name": FIRM_DEFAULT, "company_address": "", "manager_email": "", "firm_gstin": "",
                    "firm_pan": "", "firm_bank": "", "invoice_prefix": "SS", "gst_rate": "18",
                    "reminder_days": "3", "sac_code": "9982", "signatory": "Partner",
                    "firm_tagline": "Chartered Accountants", "sign_auto": "1", "gst_type": "Regular",
                    "terms": "Payment due within 15 days of invoice date. Please quote the invoice number "
                             "while making payment."}
LONG_SETTINGS = {"company_address": 600, "terms": 1500, "firm_logo": 700000, "firm_sign": 700000}


def hash_pw(password, salt=None):
    salt = salt or secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), 120000)
    return h.hex(), salt


class Database:
    def __init__(self, path):
        self.lock = threading.RLock()
        self.con = sqlite3.connect(path, check_same_thread=False)
        self.con.row_factory = sqlite3.Row
        with self.lock:
            self.con.executescript(SCHEMA)
            self.con.commit()
        if not self.q("SELECT 1 FROM users LIMIT 1"):
            h, s = hash_pw(DEFAULT_ADMIN[1])
            self.x("INSERT INTO users(username, full_name, pw_hash, salt, role, designation)"
                   " VALUES(?,?,?,?,?,?)", (DEFAULT_ADMIN[0], "Administrator", h, s, "admin", "Manager"))
        for k, v in SETTING_DEFAULTS.items():
            self.x("INSERT OR IGNORE INTO settings(key, value) VALUES(?,?)", (k, v))
        with self.lock:
            for table, col, decl in MIGRATIONS:
                cols = [r[1] for r in self.con.execute(f"PRAGMA table_info({table})").fetchall()]
                if col not in cols:
                    self.con.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl}")
            self.con.commit()

    def q(self, sql, args=()):
        with self.lock:
            return [dict(r) for r in self.con.execute(sql, args).fetchall()]

    def one(self, sql, args=()):
        r = self.q(sql, args)
        return r[0] if r else None

    def x(self, sql, args=()):
        with self.lock:
            cur = self.con.execute(sql, args)
            self.con.commit()
            return cur.lastrowid

    def settings(self, with_logo=False):
        s = {r["key"]: r["value"] for r in self.q("SELECT key, value FROM settings")}
        if not with_logo:
            s.pop("firm_logo", None)
            s.pop("firm_sign", None)
            s["has_logo"] = "1" if self.one("SELECT 1 FROM settings WHERE key='firm_logo' AND value!=''") else ""
        return s


# =============================================================================
# 6. SERVER : CONNECTION HANDLING & BUSINESS RULES
# =============================================================================
PUBLIC_USER_FIELDS = ("username", "full_name", "role", "designation", "department", "email",
                      "salary", "pan", "bank", "doj", "pt", "tds", "active", "phone")


def public_user(u):
    return {k: u.get(k) for k in PUBLIC_USER_FIELDS}


class ClientConn:
    def __init__(self, sock, addr):
        self.sock, self.addr = sock, addr
        self.lock = threading.Lock()
        self.user = None
        self.alive = True

    def send(self, header, payload=b""):
        if not self.alive:
            return
        try:
            with self.lock:
                send_frame(self.sock, header, payload)
        except Exception:
            self.alive = False

    def close(self):
        self.alive = False
        try:
            self.sock.shutdown(socket.SHUT_RDWR)
        except Exception:
            pass
        try:
            self.sock.close()
        except Exception:
            pass


class Server:
    RELAY_KINDS = {"invite", "accept", "decline", "busy", "end", "audio", "frame"}

    def __init__(self, data_dir=DATA_DIR, port=PORT):
        self.data_dir = os.path.join(data_dir, "server")
        self.files_dir = os.path.join(self.data_dir, "files")
        os.makedirs(self.files_dir, exist_ok=True)
        self.db = Database(os.path.join(self.data_dir, "suthar_suthar.db"))
        self.port = port
        self.clients = {}            # username -> ClientConn
        self.clients_lock = threading.Lock()
        self.running = False
        self.sock = None
        self.udp = None

    # ---------------------------------------------------------------- lifecycle
    def start(self):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        if os.name != "nt":
            self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.sock.bind(("0.0.0.0", self.port))
        self.sock.listen(50)
        self.running = True
        threading.Thread(target=self._accept_loop, daemon=True).start()
        threading.Thread(target=self._recurring_loop, daemon=True).start()
        try:
            self.udp = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            if os.name != "nt":
                self.udp.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.udp.bind(("0.0.0.0", DISCOVERY_PORT))
            threading.Thread(target=self._discovery_loop, daemon=True).start()
        except OSError:
            self.udp = None  # discovery optional

    def stop(self):
        self.running = False
        for s in (self.sock, self.udp):
            try:
                s and s.close()
            except Exception:
                pass
        with self.clients_lock:
            for c in list(self.clients.values()):
                c.close()
            self.clients.clear()

    def _discovery_loop(self):
        while self.running:
            try:
                data, addr = self.udp.recvfrom(1024)
                if data == b"SUTHAR_DISCOVER":
                    msg = f"SUTHAR_SERVER|{self.port}|{socket.gethostname()}".encode()
                    self.udp.sendto(msg, addr)
            except Exception:
                if not self.running:
                    break
                time.sleep(0.2)

    def _accept_loop(self):
        while self.running:
            try:
                s, addr = self.sock.accept()
                s.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
                threading.Thread(target=self._client_loop, args=(ClientConn(s, addr),),
                                 daemon=True).start()
            except Exception:
                if not self.running:
                    break
                time.sleep(0.2)

    def _client_loop(self, conn):
        try:
            while self.running and conn.alive:
                hdr, payload = recv_frame(conn.sock)
                self._dispatch(conn, hdr, payload)
        except Exception:
            pass
        finally:
            conn.close()
            if conn.user:
                uname = conn.user["username"]
                with self.clients_lock:
                    if self.clients.get(uname) is conn:
                        del self.clients[uname]
                self.broadcast({"type": "push", "event": "presence"})

    def _dispatch(self, conn, hdr, payload):
        t = str(hdr.get("type", ""))
        rid = hdr.get("rid")
        fn = getattr(self, "h_" + t, None) if re.fullmatch(r"[a-z_]{2,40}", t) else None
        try:
            if fn is None:
                raise ValueError("Unknown request")
            if conn.user is None and t != "login":
                raise PermissionError("Please log in first")
            res = fn(conn, hdr, payload)
            data, out = (res if isinstance(res, tuple) else (res, b""))
            if rid is not None:
                conn.send({"type": "resp", "rid": rid, "ok": True, "data": data}, out)
        except Exception as e:
            if rid is not None:
                msg = str(e) if isinstance(e, (ValueError, PermissionError)) else f"Server error: {e}"
                conn.send({"type": "resp", "rid": rid, "ok": False, "error": msg})

    # ----------------------------------------------------------------- helpers
    def broadcast(self, header, payload=b"", exclude=None):
        with self.clients_lock:
            targets = [c for u, c in self.clients.items() if u != exclude]
        for c in targets:
            c.send(header, payload)

    def send_to(self, username, header, payload=b""):
        with self.clients_lock:
            c = self.clients.get(username)
        if c:
            c.send(header, payload)
            return True
        return False

    def notify_admins(self, text, ref=None):
        for a in self.db.q("SELECT username FROM users WHERE role='admin' AND active=1"):
            self.send_to(a["username"], {"type": "push", "event": "notify", "text": text, "ref": ref})

    @staticmethod
    def me(conn):
        return conn.user["username"]

    @staticmethod
    def is_admin(conn):
        return conn.user.get("role") == "admin"

    def need_admin(self, conn):
        if not self.is_admin(conn):
            raise PermissionError("Only the Manager / Admin can do this")

    def target(self, conn, hdr, allow_all=False):
        """Resolve 'username' param. Members can only see themselves."""
        u = str(hdr.get("username") or self.me(conn))
        if u == "ALL":
            if not (allow_all and self.is_admin(conn)):
                raise PermissionError("Not allowed")
            return None
        if u != self.me(conn) and not self.is_admin(conn):
            raise PermissionError("You can view only your own records")
        return u

    @staticmethod
    def parse_date(s, label="date"):
        try:
            return date.fromisoformat(str(s)[:10])
        except Exception:
            raise ValueError(f"Invalid {label}")

    @staticmethod
    def parse_month(s):
        m = re.fullmatch(r"(\d{4})-(\d{2})", str(s or ""))
        if not m or not (1 <= int(m.group(2)) <= 12):
            raise ValueError("Invalid month")
        return int(m.group(1)), int(m.group(2))

    def name_of(self, username):
        r = self.db.one("SELECT full_name FROM users WHERE username=?", (username,))
        return (r and r["full_name"]) or username

    def store_file(self, name, data, uploader, recipient, purpose):
        name = safe_filename(name)
        stored = uuid.uuid4().hex + "_" + name
        with open(os.path.join(self.files_dir, stored), "wb") as f:
            f.write(data)
        return self.db.x("INSERT INTO files(name, stored, size, uploader, recipient, purpose, ts)"
                         " VALUES(?,?,?,?,?,?,?)",
                         (name, stored, len(data), uploader, recipient, purpose, now_str()))

    def month_leave(self, username, y, m):
        first, last = month_bounds(y, m)
        rows = self.db.q("SELECT * FROM leaves WHERE username=? AND status IN ('Approved','Pending')"
                         " AND from_date<=? AND to_date>=?",
                         (username, last.isoformat(), first.isoformat()))
        approved = pending = 0.0
        for r in rows:
            f, t = date.fromisoformat(r["from_date"]), date.fromisoformat(r["to_date"])
            days = 0.5 if r["half_day"] else working_days(max(f, first), min(t, last))
            if r["status"] == "Approved":
                approved += days
            else:
                pending += days
        paid = min(float(PAID_LEAVE_PER_MONTH), approved)
        return {"month": f"{y:04d}-{m:02d}", "quota": PAID_LEAVE_PER_MONTH, "taken": approved,
                "paid": paid, "lwp": round(approved - paid, 2),
                "balance": round(PAID_LEAVE_PER_MONTH - paid, 2), "pending": pending}

    # ================================================================ handlers
    def h_login(self, conn, hdr, payload):
        uname = str(hdr.get("username", "")).strip().lower()
        pw = str(hdr.get("password", ""))
        u = self.db.one("SELECT * FROM users WHERE lower(username)=?", (uname,))
        ok = False
        if u and u["active"]:
            h, _ = hash_pw(pw, u["salt"])
            ok = hmac.compare_digest(h, u["pw_hash"])
        if not ok:
            time.sleep(1.0)  # slow down password guessing
            raise PermissionError("Invalid username or password")
        conn.user = u
        with self.clients_lock:
            old = self.clients.get(u["username"])
            self.clients[u["username"]] = conn
        if old and old is not conn:
            old.send({"type": "push", "event": "kicked"})
            old.close()
        self.broadcast({"type": "push", "event": "presence"}, exclude=u["username"])
        default_pw = u["username"] == DEFAULT_ADMIN[0] and hmac.compare_digest(
            hash_pw(DEFAULT_ADMIN[1], u["salt"])[0], u["pw_hash"])
        return {"user": public_user(u), "settings": self.db.settings(),
                "server_time": now_str(), "default_password": default_pw, "must_change": bool(u.get("must_change"))}

    def h_ping(self, conn, hdr, payload):
        return {"server_time": now_str()}

    def h_users(self, conn, hdr, payload):
        with self.clients_lock:
            online = set(self.clients.keys())
        rows = self.db.q("SELECT username, full_name, role, designation, department, email, phone FROM users"
                         " WHERE active=1 ORDER BY full_name")
        for r in rows:
            r["online"] = r["username"] in online
        return rows

    def h_me(self, conn, hdr, payload):
        u = self.db.one("SELECT * FROM users WHERE username=?", (self.me(conn),))
        return {"user": public_user(u), "settings": self.db.settings()}

    # ----------------------------------------------------------------- chat
    def _msg_out(self, mid):
        m = self.db.one("SELECT m.*, f.name AS file_name, f.size AS file_size FROM messages m"
                        " LEFT JOIN files f ON f.id=m.file_id WHERE m.id=?", (mid,))
        m["sender_name"] = self.name_of(m["sender"])
        return m

    def _deliver(self, m):
        hdr = {"type": "push", "event": "chat", "msg": m}
        if m["recipient"] == "ALL":
            self.broadcast(hdr)
        else:
            self.send_to(m["sender"], hdr)
            if m["recipient"] != m["sender"]:
                self.send_to(m["recipient"], hdr)

    def _check_recipient(self, to):
        if to != "ALL" and not self.db.one("SELECT 1 FROM users WHERE username=? AND active=1", (to,)):
            raise ValueError("Recipient not found")

    def h_chat_history(self, conn, hdr, payload):
        peer, me = str(hdr.get("peer", "ALL")), self.me(conn)
        base = ("SELECT m.*, f.name AS file_name, f.size AS file_size FROM messages m"
                " LEFT JOIN files f ON f.id=m.file_id WHERE ")
        if peer == "ALL":
            rows = self.db.q(base + "m.recipient='ALL' ORDER BY m.id DESC LIMIT 300")
        else:
            rows = self.db.q(base + "((m.sender=? AND m.recipient=?) OR (m.sender=? AND m.recipient=?))"
                             " ORDER BY m.id DESC LIMIT 300", (me, peer, peer, me))
        names = {u["username"]: u["full_name"] for u in self.db.q("SELECT username, full_name FROM users")}
        for r in rows:
            r["sender_name"] = names.get(r["sender"], r["sender"])
        return rows[::-1]

    def h_chat_send(self, conn, hdr, payload):
        to = str(hdr.get("to", "ALL"))
        text = str(hdr.get("text", "")).strip()[:5000]
        if not text:
            raise ValueError("Empty message")
        self._check_recipient(to)
        mid = self.db.x("INSERT INTO messages(sender, recipient, text, file_id, ts) VALUES(?,?,?,?,?)",
                        (self.me(conn), to, text, None, now_str()))
        m = self._msg_out(mid)
        self._deliver(m)
        return m

    def h_file_upload(self, conn, hdr, payload):
        to = str(hdr.get("to", "ALL"))
        self._check_recipient(to)
        if not payload:
            raise ValueError("Empty file")
        fid = self.store_file(hdr.get("name"), payload, self.me(conn), to, "chat")
        mid = self.db.x("INSERT INTO messages(sender, recipient, text, file_id, ts) VALUES(?,?,?,?,?)",
                        (self.me(conn), to, "", fid, now_str()))
        m = self._msg_out(mid)
        self._deliver(m)
        return m

    def h_file_download(self, conn, hdr, payload):
        f = self.db.one("SELECT * FROM files WHERE id=?", (int(hdr.get("file_id", 0)),))
        if not f:
            raise ValueError("File not found")
        me = self.me(conn)
        if not (self.is_admin(conn) or f["uploader"] == me or f["recipient"] in (me, "ALL")):
            raise PermissionError("You do not have access to this file")
        path = os.path.join(self.files_dir, f["stored"])
        if not os.path.isfile(path):
            raise ValueError("File missing on server")
        with open(path, "rb") as fh:
            return {"name": f["name"], "size": f["size"]}, fh.read()

    # ------------------------------------------------- voice / screen relay
    def h_relay(self, conn, hdr, payload):
        kind, mode = str(hdr.get("kind")), str(hdr.get("mode"))
        if kind not in self.RELAY_KINDS or mode not in ("call", "screen"):
            raise ValueError("Bad relay")
        to = str(hdr.get("to", ""))
        out = {"type": "push", "event": "relay", "from": self.me(conn), "kind": kind,
               "mode": mode, "sid": str(hdr.get("sid", ""))[:64],
               "from_name": conn.user.get("full_name")}
        if to == "ALL":
            if mode != "screen":
                raise ValueError("Group voice calls are not supported")
            with self.clients_lock:
                n = len([u for u in self.clients if u != self.me(conn)])
            if n == 0 and hdr.get("rid") is not None:
                raise ValueError("No other team member is online")
            self.broadcast(out, payload, exclude=self.me(conn))
            return {"delivered": n}
        if not self.send_to(to, out, payload):
            raise ValueError("This team member is offline")
        return {"delivered": 1}

    # ------------------------------------------------------------ attendance
    def h_att_mark(self, conn, hdr, payload):
        me, today, ts = self.me(conn), date.today().isoformat(), datetime.now().strftime("%H:%M:%S")
        mode = hdr.get("mode") if hdr.get("mode") in ("Office", "Work From Home", "Client Visit") else "Office"
        rec = self.db.one("SELECT * FROM attendance WHERE username=? AND day=?", (me, today))
        if hdr.get("action") == "in":
            if rec:
                raise ValueError(f"Already checked in today at {rec['check_in']}")
            self.db.x("INSERT INTO attendance(username, day, check_in, mode) VALUES(?,?,?,?)",
                      (me, today, ts, mode))
        elif hdr.get("action") == "out":
            if not rec:
                raise ValueError("Please check in first")
            self.db.x("UPDATE attendance SET check_out=? WHERE id=?", (ts, rec["id"]))
        else:
            raise ValueError("Bad action")
        return self.db.one("SELECT * FROM attendance WHERE username=? AND day=?", (me, today))

    def h_att_today(self, conn, hdr, payload):
        return self.db.one("SELECT * FROM attendance WHERE username=? AND day=?",
                           (self.me(conn), date.today().isoformat()))

    def h_att_list(self, conn, hdr, payload):
        u = self.target(conn, hdr, allow_all=True)
        y, m = self.parse_month(hdr.get("month"))
        like = f"{y:04d}-{m:02d}-%"
        sql = ("SELECT a.*, u.full_name FROM attendance a LEFT JOIN users u ON u.username=a.username"
               " WHERE a.day LIKE ?")
        args = [like]
        if u:
            sql += " AND a.username=?"
            args.append(u)
        return self.db.q(sql + " ORDER BY a.day, u.full_name", args)

    # ----------------------------------------------------------------- leave
    def h_leave_apply(self, conn, hdr, payload):
        me = self.me(conn)
        f = self.parse_date(hdr.get("from_date"), "From date")
        t = self.parse_date(hdr.get("to_date"), "To date")
        if t < f:
            raise ValueError("'To' date cannot be before 'From' date")
        if (t - f).days > 90:
            raise ValueError("Leave period cannot exceed 90 days")
        half = 1 if hdr.get("half_day") else 0
        if half and f != t:
            raise ValueError("Half-day leave must be for a single date")
        ltype = hdr.get("ltype") if hdr.get("ltype") in LEAVE_TYPES else LEAVE_TYPES[0]
        reason = str(hdr.get("reason", "")).strip()[:500]
        if not reason:
            raise ValueError("Please enter the reason for leave")
        days = 0.5 if half else working_days(f, t)
        if days <= 0:
            raise ValueError("Selected dates fall on weekly off (Sunday)")
        clash = self.db.one("SELECT id FROM leaves WHERE username=? AND status IN ('Pending','Approved')"
                            " AND from_date<=? AND to_date>=?", (me, t.isoformat(), f.isoformat()))
        if clash:
            raise ValueError(f"These dates overlap with leave application #{clash['id']}")
        lid = self.db.x("INSERT INTO leaves(username, from_date, to_date, half_day, days, ltype, reason,"
                        " applied_on) VALUES(?,?,?,?,?,?,?,?)",
                        (me, f.isoformat(), t.isoformat(), half, days, ltype, reason, now_str()))
        self.notify_admins(f"New leave application #{lid} from {conn.user['full_name']} ({days} day(s))",
                           ref={"kind": "page", "page": "leave"})
        return {"id": lid, "days": days}

    def h_leave_list(self, conn, hdr, payload):
        u = self.target(conn, hdr, allow_all=True)
        year = int(hdr.get("year") or date.today().year)
        sql = ("SELECT l.*, u.full_name FROM leaves l LEFT JOIN users u ON u.username=l.username"
               " WHERE (substr(l.from_date,1,4)=? OR substr(l.to_date,1,4)=?)")
        args = [str(year), str(year)]
        if u:
            sql += " AND l.username=?"
            args.append(u)
        return self.db.q(sql + " ORDER BY l.from_date DESC", args)

    def h_leave_decide(self, conn, hdr, payload):
        self.need_admin(conn)
        status = hdr.get("status")
        if status not in ("Approved", "Rejected"):
            raise ValueError("Bad status")
        lv = self.db.one("SELECT * FROM leaves WHERE id=?", (int(hdr.get("id", 0)),))
        if not lv:
            raise ValueError("Leave not found")
        if lv["status"] not in ("Pending", "Approved"):
            raise ValueError(f"Leave is already {lv['status']}")
        self.db.x("UPDATE leaves SET status=?, decided_by=?, remark=? WHERE id=?",
                  (status, self.me(conn), str(hdr.get("remark", ""))[:300], lv["id"]))
        self.send_to(lv["username"], {"type": "push", "event": "notify", "ref": {"kind": "page", "page": "leave"},
                                      "text": f"Your leave #{lv['id']} ({lv['from_date']} to "
                                              f"{lv['to_date']}) has been {status}"})
        return {"id": lv["id"], "status": status}

    def h_leave_cancel(self, conn, hdr, payload):
        lv = self.db.one("SELECT * FROM leaves WHERE id=?", (int(hdr.get("id", 0)),))
        if not lv or lv["username"] != self.me(conn):
            raise ValueError("Leave not found")
        if lv["status"] != "Pending":
            raise ValueError("Only pending applications can be cancelled")
        self.db.x("UPDATE leaves SET status='Cancelled' WHERE id=?", (lv["id"],))
        return {"id": lv["id"]}

    def h_leave_summary(self, conn, hdr, payload):
        u = self.target(conn, hdr)
        year = int(hdr.get("year") or date.today().year)
        return [self.month_leave(u, year, m) for m in range(1, 13)]

    # ------------------------------------------------------------- salary
    def h_payslip(self, conn, hdr, payload):
        u = self.target(conn, hdr)
        y, m = self.parse_month(hdr.get("month"))
        user = self.db.one("SELECT * FROM users WHERE username=?", (u,))
        if not user:
            raise ValueError("User not found")
        lv = self.month_leave(u, y, m)
        dim = calendar.monthrange(y, m)[1]
        present = self.db.one("SELECT COUNT(*) AS n FROM attendance WHERE username=? AND day LIKE ?",
                              (u, f"{y:04d}-{m:02d}-%"))["n"]
        reimb = self.db.one("SELECT COALESCE(SUM(amount),0) AS s FROM reimb WHERE username=? AND"
                            " status='Approved' AND day LIKE ?", (u, f"{y:04d}-{m:02d}-%"))["s"]
        gross = round(float(user["salary"] or 0), 2)
        basic = round(gross * 0.50, 2)
        hra = round(gross * 0.20, 2)
        special = round(gross - basic - hra, 2)
        lwp_amt = round(gross / dim * lv["lwp"], 2) if dim else 0.0
        pt, tds = round(float(user["pt"] or 0), 2), round(float(user["tds"] or 0), 2)
        total_ded = round(lwp_amt + pt + tds, 2)
        net = round(gross - total_ded + float(reimb), 2)
        return {"user": public_user(user), "settings": self.db.settings(), "month": f"{y:04d}-{m:02d}",
                "days_in_month": dim, "paid_days": dim - lv["lwp"], "present_days": present,
                "leave": lv, "earnings": [["Basic Salary", basic], ["House Rent Allowance", hra],
                                          ["Special Allowance", special]],
                "deductions": [["Loss of Pay (%.1f day(s))" % lv["lwp"], lwp_amt],
                               ["Professional Tax", pt], ["TDS (Income Tax)", tds]],
                "gross": gross, "total_deductions": total_ded, "reimbursement": round(float(reimb), 2),
                "net": net, "net_words": amount_in_words(net), "generated": now_str()}

    # ------------------------------------------------------- reimbursement
    def h_reimb_add(self, conn, hdr, payload):
        d = self.parse_date(hdr.get("day"))
        if d > date.today():
            raise ValueError("Expense date cannot be in the future")
        cat = hdr.get("category")
        if cat not in REIMB_CATEGORIES:
            raise ValueError("Invalid category")
        try:
            amt = round(float(hdr.get("amount")), 2)
        except Exception:
            raise ValueError("Invalid amount")
        if not (0 < amt <= 1000000):
            raise ValueError("Amount must be between 0.01 and 10,00,000")
        desc = str(hdr.get("description", "")).strip()[:300]
        if not desc:
            raise ValueError("Please enter a description")
        fid = None
        if payload:
            fid = self.store_file(hdr.get("bill_name") or "bill", payload, self.me(conn), "ADMIN", "bill")
        rid = self.db.x("INSERT INTO reimb(username, day, category, description, amount, file_id,"
                        " submitted_on) VALUES(?,?,?,?,?,?,?)",
                        (self.me(conn), d.isoformat(), cat, desc, amt, fid, now_str()))
        self.notify_admins(f"New reimbursement claim #{rid} from {conn.user['full_name']}: Rs. {inr(amt)}",
                           ref={"kind": "page", "page": "reimb"})
        return {"id": rid}

    def h_reimb_list(self, conn, hdr, payload):
        u = self.target(conn, hdr, allow_all=True)
        y, m = self.parse_month(hdr.get("month"))
        sql = ("SELECT r.*, u.full_name, u.designation FROM reimb r LEFT JOIN users u"
               " ON u.username=r.username WHERE r.day LIKE ?")
        args = [f"{y:04d}-{m:02d}-%"]
        if u:
            sql += " AND r.username=?"
            args.append(u)
        return {"rows": self.db.q(sql + " ORDER BY u.full_name, r.day", args),
                "settings": self.db.settings()}

    def h_reimb_decide(self, conn, hdr, payload):
        self.need_admin(conn)
        status = hdr.get("status")
        if status not in ("Approved", "Rejected"):
            raise ValueError("Bad status")
        r = self.db.one("SELECT * FROM reimb WHERE id=?", (int(hdr.get("id", 0)),))
        if not r:
            raise ValueError("Claim not found")
        self.db.x("UPDATE reimb SET status=?, decided_by=? WHERE id=?", (status, self.me(conn), r["id"]))
        self.send_to(r["username"], {"type": "push", "event": "notify", "ref": {"kind": "page", "page": "reimb"},
                                     "text": f"Reimbursement claim #{r['id']} (Rs. {inr(r['amount'])})"
                                             f" has been {status}"})
        return {"id": r["id"]}

    def h_reimb_delete(self, conn, hdr, payload):
        r = self.db.one("SELECT * FROM reimb WHERE id=?", (int(hdr.get("id", 0)),))
        if not r or r["username"] != self.me(conn):
            raise ValueError("Claim not found")
        if r["status"] != "Pending":
            raise ValueError("Only pending claims can be deleted")
        self.db.x("DELETE FROM reimb WHERE id=?", (r["id"],))
        return {"id": r["id"]}

    # ------------------------------------------------------------ dashboard
    def h_dashboard(self, conn, hdr, payload):
        me, t = self.me(conn), date.today()
        ym = f"{t.year:04d}-{t.month:02d}-%"
        out = {"today": self.db.one("SELECT * FROM attendance WHERE username=? AND day=?", (me, t.isoformat())),
               "leave": self.month_leave(me, t.year, t.month),
               "reimb": self.db.one("SELECT COALESCE(SUM(amount),0) AS s, COUNT(*) AS n FROM reimb"
                                    " WHERE username=? AND day LIKE ? AND status!='Rejected'", (me, ym)),
               "present": self.db.one("SELECT COUNT(*) AS n FROM attendance WHERE username=? AND day LIKE ?",
                                      (me, ym))["n"]}
        with self.clients_lock:
            out["online"] = len(self.clients)
        out["team"] = self.db.one("SELECT COUNT(*) AS n FROM users WHERE active=1")["n"]
        rem = self.h_reminders(conn, {}, b"")
        out["hearings"] = len(rem["upcoming"])
        out["hearing_days"] = rem["days"]
        out["overdue"] = len(rem["overdue"])
        mt = self.db.q("SELECT due_date FROM tasks WHERE assigned_to=? AND status IN ('Open','In Progress')", (me,))
        out["tasks_open"] = len(mt)
        out["tasks_red"] = sum(1 for x in mt if (date.fromisoformat(x["due_date"]) - t).days <= 3)
        if self.is_admin(conn):
            out["pending_leaves"] = self.db.one("SELECT COUNT(*) AS n FROM leaves WHERE status='Pending'")["n"]
            out["pending_reimb"] = self.db.one("SELECT COUNT(*) AS n FROM reimb WHERE status='Pending'")["n"]
            out["present_today"] = self.db.one("SELECT COUNT(*) AS n FROM attendance WHERE day=?",
                                               (t.isoformat(),))["n"]
        return out

    # ------------------------------------------------------- admin / users
    def h_user_list_full(self, conn, hdr, payload):
        self.need_admin(conn)
        return [public_user(u) for u in self.db.q("SELECT * FROM users ORDER BY full_name")]

    def h_user_save(self, conn, hdr, payload):
        self.need_admin(conn)
        d = hdr.get("user") or {}
        uname = str(d.get("username", "")).strip().lower()
        if not re.fullmatch(r"[a-z0-9_.]{3,30}", uname):
            raise ValueError("Username: 3-30 characters, letters/digits/._ only")
        full = str(d.get("full_name", "")).strip()[:80]
        if not full:
            raise ValueError("Full name is required")
        role = "admin" if d.get("role") == "admin" else "member"
        vals = dict(full_name=full, role=role,
                    designation=str(d.get("designation", ""))[:60], department=str(d.get("department", ""))[:60],
                    email=str(d.get("email", ""))[:120], pan=str(d.get("pan", "")).upper()[:10],
                    phone=re.sub(r"[^0-9+]", "", str(d.get("phone", "")))[:15],
                    bank=str(d.get("bank", ""))[:40], doj=str(d.get("doj", ""))[:10],
                    salary=max(0.0, float(d.get("salary") or 0)), pt=max(0.0, float(d.get("pt") or 0)),
                    tds=max(0.0, float(d.get("tds") or 0)), active=1 if d.get("active", 1) else 0)
        pw = str(d.get("password", ""))
        exists = self.db.one("SELECT * FROM users WHERE username=?", (uname,))
        if hdr.get("new"):
            if exists:
                raise ValueError("Username already exists")
            if len(pw) < 6:
                raise ValueError("Password must have at least 6 characters")
            h, s = hash_pw(pw)
            cols = ["username", "pw_hash", "salt"] + list(vals.keys())
            self.db.x(f"INSERT INTO users({','.join(cols)}) VALUES({','.join('?' * len(cols))})",
                      [uname, h, s] + list(vals.values()))
            self.db.x("UPDATE users SET must_change=1 WHERE username=?", (uname,))   # own password at 1st login
        else:
            if not exists:
                raise ValueError("User not found")
            if uname == self.me(conn) and (role != "admin" or not vals["active"]):
                raise ValueError("You cannot remove your own admin rights or deactivate yourself")
            if exists["role"] == "admin" and (role != "admin" or not vals["active"]) and \
                    not [a for a in self._active_admins() if a != uname]:
                raise ValueError("At least one active admin must remain")
            self.db.x("UPDATE users SET " + ", ".join(f"{k}=?" for k in vals) + " WHERE username=?",
                      list(vals.values()) + [uname])
            if pw:
                if len(pw) < 6:
                    raise ValueError("Password must have at least 6 characters")
                h, s = hash_pw(pw)
                self.db.x("UPDATE users SET pw_hash=?, salt=?, must_change=? WHERE username=?",
                          (h, s, 0 if uname == self.me(conn) else 1, uname))
        self.broadcast({"type": "push", "event": "presence"})
        return {"username": uname}

    def h_change_password(self, conn, hdr, payload):
        u = self.db.one("SELECT * FROM users WHERE username=?", (self.me(conn),))
        if not hmac.compare_digest(hash_pw(str(hdr.get("old", "")), u["salt"])[0], u["pw_hash"]):
            raise ValueError("Current password is incorrect")
        new = str(hdr.get("new", ""))
        if len(new) < 6:
            raise ValueError("New password must have at least 6 characters")
        if hmac.compare_digest(hash_pw(new, u["salt"])[0], u["pw_hash"]):
            raise ValueError("New password must be different from the current one")
        h, s = hash_pw(new)
        self.db.x("UPDATE users SET pw_hash=?, salt=?, must_change=0 WHERE username=?", (h, s, u["username"]))
        return {"ok": True}

    def _active_admins(self):
        return [r["username"] for r in self.db.q("SELECT username FROM users WHERE role='admin' AND active=1")]

    def h_user_reset_pw(self, conn, hdr, payload):
        """Only the Manager / Admin can reset a password. User must set a new one at next login."""
        self.need_admin(conn)
        uname = str(hdr.get("username", ""))
        u = self.db.one("SELECT * FROM users WHERE username=?", (uname,))
        if not u:
            raise ValueError("User not found")
        pw = str(hdr.get("password") or "")
        if not pw:
            pw = "".join(secrets.choice("ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789") for _ in range(8))
        if len(pw) < 6:
            raise ValueError("Password must have at least 6 characters")
        h, s = hash_pw(pw)
        self.db.x("UPDATE users SET pw_hash=?, salt=?, must_change=? WHERE username=?",
                  (h, s, 0 if uname == self.me(conn) else 1, uname))
        return {"username": uname, "password": pw}

    def h_user_set_role(self, conn, hdr, payload):
        """Appoint / remove an Admin (Manager). At least one active admin must always remain."""
        self.need_admin(conn)
        uname, role = str(hdr.get("username", "")), ("admin" if hdr.get("role") == "admin" else "member")
        u = self.db.one("SELECT * FROM users WHERE username=?", (uname,))
        if not u:
            raise ValueError("User not found")
        if role == "member":
            if uname == self.me(conn):
                raise ValueError("You cannot remove your own admin rights. Ask another admin to do it.")
            if u["role"] == "admin" and len([a for a in self._active_admins() if a != uname]) == 0:
                raise ValueError("At least one admin must remain")
        self.db.x("UPDATE users SET role=? WHERE username=?", (role, uname))
        self.send_to(uname, {"type": "push", "event": "notify",
                             "text": ("You have been appointed as Admin. Please log out and log in again to get admin "
                                      "options." if role == "admin" else "Your admin rights have been removed. Please "
                                                                          "log in again.")})
        self.broadcast({"type": "push", "event": "presence"})
        return {"username": uname, "role": role}

    def h_settings_set(self, conn, hdr, payload):
        self.need_admin(conn)
        for k in SETTING_KEYS:
            if k in hdr:
                v = str(hdr[k]).strip()[:LONG_SETTINGS.get(k, 300)]
                if k in ("firm_logo", "firm_sign") and v and not re.fullmatch(r"[A-Za-z0-9+/=]+", v):
                    raise ValueError("Invalid image data")
                if k == "gst_rate":
                    v = str(min(40.0, max(0.0, float(v or 0))))
                if k == "gst_type" and v not in GST_TYPES:
                    raise ValueError("GST registration must be Regular, Composition or Unregistered")
                if k == "reminder_days":
                    v = str(min(30, max(0, int(float(v or 3)))))
                self.db.x("INSERT OR REPLACE INTO settings(key, value) VALUES(?,?)", (k, v))
        return self.db.settings()

    # ======================================================== LITIGATION
    def _case(self, cid):
        c = self.db.one("SELECT * FROM cases WHERE id=?", (int(cid or 0),))
        if not c:
            raise ValueError("Case not found")
        return c

    def _lit_push(self, text=None, users=(), exclude=None, ref=None):
        self.broadcast({"type": "push", "event": "lit"})
        for u in set(users):
            if u and u != exclude and text:
                self.send_to(u, {"type": "push", "event": "notify", "text": text, "ref": ref})

    # ------------------------------------------------ confidentiality / access rights
    def _acc(self, ref_type, ref_id, username):
        return self.db.one("SELECT * FROM item_access WHERE ref_type=? AND ref_id=? AND username=?",
                           (ref_type, ref_id, username))

    def _case_visible(self, conn, c):
        """Confidential case: only Manager(s), the handler and members given 'can view' right."""
        if self.is_admin(conn) or not c.get("restricted"):
            return True
        me = self.me(conn)
        if c.get("assigned_to") == me:
            return True
        a = self._acc("case", c["id"], me)
        return bool(a and a["can_view"])

    def _case_v(self, conn, cid):
        c = self._case(cid)
        if not self._case_visible(conn, c):
            raise PermissionError("This case is confidential. Please contact the Manager.")
        return c

    def _fees_right(self, conn, ref_type, ref_id):
        if self.is_admin(conn):
            return True
        a = self._acc(ref_type, ref_id, self.me(conn))
        return bool(a and a["can_fees"])

    def _task_fee_visible(self, conn, t):
        if self.is_admin(conn):
            return True
        if not t.get("hide_fees"):
            return self.me(conn) in (t["created_by"], t["assigned_to"]) or self._fees_right(conn, "task", t["id"])
        return self._fees_right(conn, "task", t["id"])

    def _mask_task(self, conn, t):
        if not self._task_fee_visible(conn, t):
            t["est_amount"] = t["billing_amount"] = None
            t["fee_hidden"] = True
        return t

    def h_access_get(self, conn, hdr, payload):
        self.need_admin(conn)
        rt, rid = hdr.get("ref_type"), int(hdr.get("ref_id") or 0)
        if rt == "case":
            item = self._case(rid)
            flags = {"restricted": item["restricted"] or 0, "hide_fees": 1}
        elif rt == "task":
            item = self._task(rid)
            flags = {"restricted": 0, "hide_fees": item["hide_fees"] or 0}
        else:
            raise ValueError("Bad type")
        rows = {r["username"]: r for r in self.db.q("SELECT * FROM item_access WHERE ref_type=? AND ref_id=?", (rt, rid))}
        users = []
        for u in self.db.q("SELECT username, full_name, role FROM users WHERE active=1 ORDER BY full_name"):
            if u["role"] == "admin":
                continue
            a = rows.get(u["username"])
            users.append({"username": u["username"], "full_name": u["full_name"],
                          "can_view": a["can_view"] if a else 0, "can_fees": a["can_fees"] if a else 0})
        return dict(flags, users=users, assigned_to=item.get("assigned_to"),
                    created_by=item.get("created_by"))

    def h_access_set(self, conn, hdr, payload):
        self.need_admin(conn)
        rt, rid = hdr.get("ref_type"), int(hdr.get("ref_id") or 0)
        if rt == "case":
            c = self._case(rid)
            self.db.x("UPDATE cases SET restricted=? WHERE id=?", (1 if hdr.get("restricted") else 0, rid))
            label = f"case #{rid} ({c['client_name']})"
        elif rt == "task":
            t = self._task(rid)
            self.db.x("UPDATE tasks SET hide_fees=? WHERE id=?", (1 if hdr.get("hide_fees") else 0, rid))
            label = f"task '{t['title']}'"
        else:
            raise ValueError("Bad type")
        old = {r["username"] for r in self.db.q("SELECT username FROM item_access WHERE ref_type=? AND ref_id=? AND"
                                                 " can_view=1", (rt, rid))}
        self.db.x("DELETE FROM item_access WHERE ref_type=? AND ref_id=?", (rt, rid))
        granted = []
        for u in hdr.get("users") or []:
            cv, cf = 1 if u.get("can_view") else 0, 1 if u.get("can_fees") else 0
            if cv or cf:
                self.db.x("INSERT INTO item_access(ref_type, ref_id, username, can_view, can_fees) VALUES(?,?,?,?,?)",
                          (rt, rid, str(u.get("username"))[:30], 1 if (cv or cf) else 0, cf))
                if u.get("username") not in old:
                    granted.append(u.get("username"))
        ref = {"kind": rt, "id": rid}
        (self._lit_push if rt == "case" else self._task_push)(
            f"The Manager has given you access to {label}", granted, exclude=self.me(conn), ref=ref)
        return {"ok": True}

    def _admins(self):
        return [a["username"] for a in self.db.q("SELECT username FROM users WHERE role='admin' AND active=1")]

    @staticmethod
    def _money(v, label="Amount", allow_zero=False):
        try:
            v = round(float(v), 2)
        except Exception:
            raise ValueError(f"Invalid {label}")
        if v < 0 or (v == 0 and not allow_zero) or v > 1e10:
            raise ValueError(f"{label} must be greater than zero")
        return v

    def h_case_save(self, conn, hdr, payload):
        d = hdr.get("case") or {}
        if d.get("law") not in LIT_LAWS:
            raise ValueError("Select the law (Income Tax / GST / VAT / Service Tax)")
        client = str(d.get("client_name", "")).strip()[:120]
        if not client:
            raise ValueError("Client name is required")
        nh = str(d.get("next_hearing") or "")
        if nh:
            nh = self.parse_date(nh, "next hearing date").isoformat()
        assigned = str(d.get("assigned_to") or "")
        if assigned and not self.db.one("SELECT 1 FROM users WHERE username=?", (assigned,)):
            raise ValueError("Assigned team member not found")
        vals = dict(law=d["law"], client_name=client,
                    client_pan=str(d.get("client_pan", "")).upper().strip()[:10],
                    client_gstin=str(d.get("client_gstin", "")).upper().strip()[:15],
                    client_address=str(d.get("client_address", ""))[:300],
                    client_email=str(d.get("client_email", ""))[:120], period=str(d.get("period", ""))[:40],
                    client_phone=re.sub(r"[^0-9+]", "", str(d.get("client_phone", "")))[:15],
                    est_fee=self._money(d.get("est_fee") or 0, "Estimated fee", True),
                    ref_no=str(d.get("ref_no", ""))[:120], forum=str(d.get("forum", ""))[:120],
                    section=str(d.get("section", ""))[:120], nature=str(d.get("nature", ""))[:300],
                    demand=self._money(d.get("demand") or 0, "Demand", True), assigned_to=assigned,
                    status=d.get("status") if d.get("status") in LIT_STATUSES else "Open",
                    next_hearing=nh, next_purpose=str(d.get("next_purpose", ""))[:200],
                    remarks=str(d.get("remarks", ""))[:1000], updated_on=now_str())
        cid = hdr.get("id")
        if cid:
            old = self._case_v(conn, cid)
            if not self.is_admin(conn):
                vals["est_fee"] = old["est_fee"] or 0       # only the Manager sets fee estimates
            self.db.x("UPDATE cases SET " + ", ".join(f"{k}=?" for k in vals) + " WHERE id=?",
                      list(vals.values()) + [old["id"]])
            cid = old["id"]
            newly = assigned and assigned != old["assigned_to"]
        else:
            if not self.is_admin(conn):
                vals["est_fee"] = 0
            vals.update(created_by=self.me(conn), created_on=now_str())
            cid = self.db.x(f"INSERT INTO cases({','.join(vals)}) VALUES({','.join('?' * len(vals))})",
                            list(vals.values()))
            newly = bool(assigned)
        self._lit_push(f"Case #{cid} ({client} - {d['law']}) has been assigned to you"
                       + (f". Next hearing: {nh}" if nh else ""),
                       [assigned] if newly else [], exclude=self.me(conn), ref={"kind": "case", "id": cid})
        return {"id": cid}

    def h_case_list(self, conn, hdr, payload):
        adm = self.is_admin(conn)
        sql = ("SELECT c.*, u.full_name AS assigned_name,"
               " (SELECT COALESCE(SUM(total),0) FROM invoices i WHERE i.case_id=c.id AND i.status!='Cancelled') AS billed,"
               " (SELECT COALESCE(SUM(received+tds),0) FROM invoices i WHERE i.case_id=c.id AND i.status!='Cancelled') AS realised"
               " FROM cases c LEFT JOIN users u ON u.username=c.assigned_to WHERE 1=1")
        args = []
        if hdr.get("law") in LIT_LAWS:
            sql += " AND c.law=?"
            args.append(hdr["law"])
        grp = hdr.get("group", "Open")
        qm = ",".join("?" * len(LIT_OPEN))
        if grp == "Open":
            sql += f" AND c.status IN ({qm})"
            args += LIT_OPEN
        elif grp == "Closed":
            sql += f" AND c.status NOT IN ({qm})"
            args += LIT_OPEN
        s = str(hdr.get("search", "")).strip()
        if s:
            sql += (" AND (c.client_name LIKE ? OR c.client_pan LIKE ? OR c.client_gstin LIKE ? OR"
                    " c.ref_no LIKE ? OR c.period LIKE ? OR c.section LIKE ? OR c.forum LIKE ?)")
            args += [f"%{s}%"] * 7
        rows = self.db.q(sql + " ORDER BY CASE WHEN c.next_hearing='' THEN 1 ELSE 0 END, c.next_hearing, c.id", args)
        rows = [r for r in rows if self._case_visible(conn, r)]
        for r in rows:
            r["outstanding"] = round(r["billed"] - r["realised"], 2) if adm else None
            if not adm:
                r["billed"] = r["realised"] = r["est_fee"] = None
        t = date.today()
        allopen = [c for c in self.db.q(f"SELECT * FROM cases WHERE status IN ({qm})", LIT_OPEN)
                   if self._case_visible(conn, c)]
        summ = {"open": len(allopen),
                "next7": sum(1 for c in allopen if c["next_hearing"] and
                             t.isoformat() <= c["next_hearing"] <= (t + timedelta(days=7)).isoformat()),
                "overdue": sum(1 for c in allopen if c["next_hearing"] and c["next_hearing"] < t.isoformat()),
                "demand": round(sum(c["demand"] or 0 for c in allopen), 2)}
        if adm:
            r = self.db.one("SELECT COALESCE(SUM(total-received-tds),0) AS o FROM invoices WHERE status!='Cancelled'")
            summ["outstanding"] = round(r["o"], 2)
        return {"rows": rows, "summary": summ}

    def h_case_get(self, conn, hdr, payload):
        c = self._case_v(conn, hdr.get("id"))
        c["assigned_name"] = self.name_of(c["assigned_to"]) if c["assigned_to"] else ""
        if self._fees_right(conn, "case", c["id"]):
            c["billed_fees"] = self.db.one("SELECT COALESCE(SUM(amount),0) AS s FROM fees WHERE case_id=? AND"
                                           " invoice_id IN (SELECT id FROM invoices WHERE status!='Cancelled')",
                                           (c["id"],))["s"]
            c["unbilled_fees"] = self.db.one("SELECT COALESCE(SUM(amount),0) AS s FROM fees WHERE case_id=? AND"
                                             " invoice_id IS NULL", (c["id"],))["s"]
            iv = self.db.one("SELECT COALESCE(SUM(total),0) AS b, COALESCE(SUM(received+tds),0) AS r FROM invoices"
                             " WHERE case_id=? AND status!='Cancelled'", (c["id"],))
            c["inv_billed"], c["inv_received"] = round(iv["b"], 2), round(iv["r"], 2)
            c["balance_due"] = round(iv["b"] - iv["r"], 2)
        else:
            c["est_fee"] = None
        out = {"case": c, "settings": self.db.settings(),
               "proceedings": self.db.q("SELECT p.*, u.full_name AS by_name, f.name AS file_name FROM proceedings p"
                                        " LEFT JOIN users u ON u.username=p.by_user LEFT JOIN files f ON f.id=p.file_id"
                                        " WHERE p.case_id=? ORDER BY p.pdate, p.id", (c["id"],)),
               "ope": self.db.q("SELECT o.*, u.full_name AS by_name, i.inv_no FROM ope o LEFT JOIN users u"
                                " ON u.username=o.by_user LEFT JOIN invoices i ON i.id=o.invoice_id"
                                " WHERE o.case_id=? ORDER BY o.edate", (c["id"],))}
        if self.is_admin(conn):
            out["fees"] = self.db.q("SELECT f.*, i.inv_no FROM fees f LEFT JOIN invoices i ON i.id=f.invoice_id"
                                    " WHERE f.case_id=? ORDER BY f.fdate", (c["id"],))
            out["invoices"] = self.db.q("SELECT * FROM invoices WHERE case_id=? ORDER BY inv_date", (c["id"],))
        return out

    def h_case_delete(self, conn, hdr, payload):
        self.need_admin(conn)
        c = self._case(hdr.get("id"))
        if self.db.one("SELECT 1 FROM invoices WHERE case_id=? AND status!='Cancelled'", (c["id"],)):
            raise ValueError("This case has invoices. Cancel them first (or mark the case Closed instead).")
        for t in ("proceedings", "fees", "ope", "invoices"):
            self.db.x(f"DELETE FROM {t} WHERE case_id=?", (c["id"],))
        self.db.x("DELETE FROM cases WHERE id=?", (c["id"],))
        self._lit_push()
        return {"id": c["id"]}

    def h_proc_add(self, conn, hdr, payload):
        c = self._case_v(conn, hdr.get("case_id"))
        pdate = self.parse_date(hdr.get("pdate"))
        ptype = hdr.get("ptype") if hdr.get("ptype") in LIT_PTYPES else "Other Update"
        details = str(hdr.get("details", "")).strip()[:4000]
        if not details:
            raise ValueError("Please write the hearing / submission report")
        nd = str(hdr.get("next_date") or "")
        if nd:
            nd = self.parse_date(nd, "next hearing date").isoformat()
        fid = None
        if payload:
            fid = self.store_file(hdr.get("file_name") or "document", payload, self.me(conn), "ALL", "lit")
        self.db.x("INSERT INTO proceedings(case_id, pdate, ptype, details, outcome, next_date, file_id,"
                  " by_user, ts) VALUES(?,?,?,?,?,?,?,?,?)",
                  (c["id"], pdate.isoformat(), ptype, details, str(hdr.get("outcome", ""))[:500], nd, fid,
                   self.me(conn), now_str()))
        upd = {"updated_on": now_str()}
        if nd:
            upd.update(next_hearing=nd, next_purpose=str(hdr.get("next_purpose", ""))[:200])
        elif hdr.get("clear_next"):
            upd.update(next_hearing="", next_purpose="")
        if hdr.get("new_status") in LIT_STATUSES:
            upd["status"] = hdr["new_status"]
            if hdr["new_status"] not in LIT_OPEN:
                upd.update(next_hearing="", next_purpose="")
        self.db.x("UPDATE cases SET " + ", ".join(f"{k}=?" for k in upd) + " WHERE id=?",
                  list(upd.values()) + [c["id"]])
        msg = (f"Case #{c['id']} {c['client_name']} ({c['law']}): {ptype} updated by {conn.user['full_name']}"
               + (f". Next hearing {date.fromisoformat(nd).strftime('%d-%m-%Y')}" if nd else ""))
        self._lit_push(msg, [c["assigned_to"]] + self._admins(), exclude=self.me(conn),
                       ref={"kind": "case", "id": c["id"]})
        return {"case_id": c["id"]}

    def h_proc_delete(self, conn, hdr, payload):
        self.need_admin(conn)
        self.db.x("DELETE FROM proceedings WHERE id=?", (int(hdr.get("id", 0)),))
        self._lit_push()
        return {}

    def h_hearings(self, conn, hdr, payload):
        s = self.parse_date(hdr.get("start")).isoformat()
        e = self.parse_date(hdr.get("end")).isoformat()
        qm = ",".join("?" * len(LIT_OPEN))
        ev = []
        for c in self.db.q(f"SELECT * FROM cases WHERE status IN ({qm}) AND next_hearing!='' AND next_hearing"
                           " BETWEEN ? AND ? ORDER BY next_hearing", LIT_OPEN + [s, e]):
            if not self._case_visible(conn, c):
                continue
            ev.append({"date": c["next_hearing"], "kind": "Hearing", "case_id": c["id"], "client": c["client_name"],
                       "law": c["law"], "forum": c["forum"], "text": c["next_purpose"], "period": c["period"],
                       "assigned_to": c["assigned_to"], "ref_no": c["ref_no"]})
        if hdr.get("history", True):
            for p in self.db.q("SELECT p.*, c.client_name, c.law, c.forum, c.period, c.restricted, c.assigned_to"
                               " FROM proceedings p JOIN cases c"
                               " ON c.id=p.case_id WHERE p.pdate BETWEEN ? AND ? ORDER BY p.pdate", (s, e)):
                if not self._case_visible(conn, dict(p, id=p["case_id"])):
                    continue
                ev.append({"date": p["pdate"], "kind": p["ptype"], "case_id": p["case_id"], "client": p["client_name"],
                           "law": p["law"], "forum": p["forum"], "text": p["details"][:200], "period": p["period"]})
        return ev

    def h_reminders(self, conn, hdr, payload):
        st = self.db.settings()
        days = int(float(st.get("reminder_days") or 3))
        t = date.today()
        qm = ",".join("?" * len(LIT_OPEN))
        sql = f"SELECT * FROM cases WHERE status IN ({qm}) AND next_hearing!=''"
        args = list(LIT_OPEN)
        if not self.is_admin(conn):
            sql += " AND (assigned_to=? OR assigned_to='')"
            args.append(self.me(conn))
        rows = [r for r in self.db.q(sql + " ORDER BY next_hearing", args) if self._case_visible(conn, r)]
        up = [r for r in rows if t.isoformat() <= r["next_hearing"] <= (t + timedelta(days=days)).isoformat()]
        od = [r for r in rows if r["next_hearing"] < t.isoformat()]
        tasks = self.db.q("SELECT * FROM tasks WHERE assigned_to=? AND status IN ('Open','In Progress')"
                          " AND due_date<=? ORDER BY urgent DESC, due_date",
                          (self.me(conn), (t + timedelta(days=3)).isoformat()))
        inv = self.h_invoice_list(conn, {"overdue": True}, b"") if self.is_admin(conn) else []
        fus = self.h_invoice_list(conn, {"followup_due": True}, b"") if self.is_admin(conn) else []
        return {"days": days, "upcoming": up, "overdue": od, "tasks": tasks, "invoices": inv, "followups": fus}

    def h_fee_add(self, conn, hdr, payload):
        self.need_admin(conn)
        c = self._case(hdr.get("case_id"))
        desc = str(hdr.get("description", "")).strip()[:300]
        if not desc:
            raise ValueError("Enter the fee description")
        fid = self.db.x("INSERT INTO fees(case_id, fdate, description, amount, by_user, ts) VALUES(?,?,?,?,?,?)",
                        (c["id"], self.parse_date(hdr.get("fdate")).isoformat(), desc,
                         self._money(hdr.get("amount")), self.me(conn), now_str()))
        self._lit_push()
        return {"id": fid}

    def h_fee_delete(self, conn, hdr, payload):
        self.need_admin(conn)
        f = self.db.one("SELECT * FROM fees WHERE id=?", (int(hdr.get("id", 0)),))
        if not f:
            raise ValueError("Fee entry not found")
        if f["invoice_id"]:
            raise ValueError("This fee is already invoiced. Cancel the invoice first.")
        self.db.x("DELETE FROM fees WHERE id=?", (f["id"],))
        return {}

    def h_ope_add(self, conn, hdr, payload):
        c = self._case_v(conn, hdr.get("case_id"))
        d = self.parse_date(hdr.get("edate"))
        if d > date.today():
            raise ValueError("Expense date cannot be in the future")
        desc = str(hdr.get("description", "")).strip()[:300]
        if not desc:
            raise ValueError("Enter the expense description")
        oid = self.db.x("INSERT INTO ope(case_id, edate, description, amount, billable, by_user, ts)"
                        " VALUES(?,?,?,?,?,?,?)", (c["id"], d.isoformat(), desc, self._money(hdr.get("amount")),
                                                    1 if hdr.get("billable", True) else 0, self.me(conn), now_str()))
        self._lit_push()
        return {"id": oid}

    def h_ope_delete(self, conn, hdr, payload):
        o = self.db.one("SELECT * FROM ope WHERE id=?", (int(hdr.get("id", 0)),))
        if not o or not (self.is_admin(conn) or o["by_user"] == self.me(conn)):
            raise ValueError("Expense not found")
        if o["invoice_id"]:
            raise ValueError("This expense is already invoiced. Cancel the invoice first.")
        self.db.x("DELETE FROM ope WHERE id=?", (o["id"],))
        return {}

    # ------------------------------------------------ invoicing (common)
    def _next_inv_no(self, d):
        st = self.db.settings()
        prefix = re.sub(r"[^A-Za-z0-9-]", "", st.get("invoice_prefix") or "SS")[:8] or "SS"
        fy = f"{d.year}-{(d.year + 1) % 100:02d}" if d.month >= 4 else f"{d.year - 1}-{d.year % 100:02d}"
        seq = len(self.db.q("SELECT id FROM invoices WHERE inv_no LIKE ?", (f"{prefix}/{fy}/%",))) + 1
        while self.db.one("SELECT 1 FROM invoices WHERE inv_no=?", (f"{prefix}/{fy}/{seq:04d}",)):
            seq += 1
        return f"{prefix}/{fy}/{seq:04d}"

    @staticmethod
    def _gst(taxable, rate, mode):
        if mode == "NONE":
            return 0.0, 0.0, 0.0
        if mode == "IGST":
            return 0.0, 0.0, round(taxable * rate / 100, 2)
        half = round(taxable * rate / 200, 2)
        return half, half, 0.0

    def _gst_args(self, hdr):
        """GST is charged only when the firm is registered as a Regular taxpayer (Firm Master)."""
        if (self.db.settings().get("gst_type") or "Regular") != "Regular":
            return "NONE", 0.0
        mode = "IGST" if hdr.get("gst_mode") == "IGST" else "CGST+SGST"
        rate = float(hdr.get("gst_rate") or 0)
        if not 0 <= rate <= 40:
            raise ValueError("Invalid GST rate")
        return mode, rate

    def h_invoice_create(self, conn, hdr, payload):
        """Litigation invoice from un-billed fees + recoverable out-of-pocket expenses of a case."""
        self.need_admin(conn)
        c = self._case(hdr.get("case_id"))
        d = self.parse_date(hdr.get("inv_date"), "invoice date")
        mode, rate = self._gst_args(hdr)
        pure = 1 if hdr.get("pure_agent", True) else 0
        with self.db.lock:
            fees = self.db.q("SELECT * FROM fees WHERE case_id=? AND invoice_id IS NULL", (c["id"],))
            opes = self.db.q("SELECT * FROM ope WHERE case_id=? AND invoice_id IS NULL AND billable=1", (c["id"],))
            if not fees and not opes:
                raise ValueError("No un-billed fees or billable out-of-pocket expenses in this case")
            fee_total = round(sum(f["amount"] for f in fees), 2)
            ope_total = round(sum(o["amount"] for o in opes), 2)
            cgst, sgst, igst = self._gst(fee_total + (0 if pure else ope_total), rate, mode)
            total = round(fee_total + ope_total + cgst + sgst + igst, 2)
            inv_no = self._next_inv_no(d)
            iid = self.db.x("INSERT INTO invoices(inv_no, case_id, inv_date, fee_total, ope_total, pure_agent,"
                            " gst_rate, gst_mode, cgst, sgst, igst, total, created_by, ts, source, client_name,"
                            " client_gstin, client_pan, client_address, client_email, client_phone)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (inv_no, c["id"], d.isoformat(), fee_total, ope_total, pure, rate, mode, cgst, sgst,
                             igst, total, self.me(conn), now_str(), "case", c["client_name"], c["client_gstin"],
                             c["client_pan"], c["client_address"], c["client_email"], c.get("client_phone", "")))
            self.db.x("UPDATE invoices SET service_details=?, gst_type=? WHERE id=?",
                      (str(hdr.get("service_details", "")).strip()[:3000],
                       self.db.settings().get("gst_type") or "Regular", iid))
            for f in fees:
                self.db.x("UPDATE fees SET invoice_id=? WHERE id=?", (iid, f["id"]))
            for o in opes:
                self.db.x("UPDATE ope SET invoice_id=? WHERE id=?", (iid, o["id"]))
        self._lit_push()
        return {"id": iid, "inv_no": inv_no, "total": total}

    def h_invoice_manual(self, conn, hdr, payload):
        """System invoice for a completed task (billing entry) or any other professional service."""
        self.need_admin(conn)
        cl = hdr.get("client") or {}
        name = str(cl.get("name", "")).strip()[:120]
        if not name:
            raise ValueError("Client name is required")
        d = self.parse_date(hdr.get("inv_date"), "invoice date")
        mode, rate = self._gst_args(hdr)
        lines = []
        for ln in (hdr.get("lines") or [])[:50]:
            desc = str(ln.get("description", "")).strip()[:300]
            amt = round(float(ln.get("amount") or 0), 2)
            if desc and amt > 0:
                lines.append((desc, str(ln.get("sac", ""))[:10], amt, 1 if ln.get("taxable", True) else 0))
        if not lines:
            raise ValueError("Add at least one line with description and amount")
        taxable = round(sum(x[2] for x in lines if x[3]), 2)
        other = round(sum(x[2] for x in lines if not x[3]), 2)
        cgst, sgst, igst = self._gst(taxable, rate, mode)
        total = round(taxable + other + cgst + sgst + igst, 2)
        bid = hdr.get("billing_id")
        with self.db.lock:
            b = None
            if bid:
                b = self.db.one("SELECT * FROM billing WHERE id=?", (int(bid),))
                if not b:
                    raise ValueError("Billing entry not found")
                if b["invoice_id"] and self.db.one("SELECT 1 FROM invoices WHERE id=? AND status!='Cancelled'",
                                                   (b["invoice_id"],)):
                    raise ValueError("An invoice is already generated for this billing entry")
            inv_no = self._next_inv_no(d)
            iid = self.db.x("INSERT INTO invoices(inv_no, case_id, inv_date, fee_total, ope_total, pure_agent,"
                            " gst_rate, gst_mode, cgst, sgst, igst, total, created_by, ts, source, billing_id,"
                            " client_name, client_gstin, client_pan, client_address, client_email, client_phone, notes)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (inv_no, None, d.isoformat(), taxable, other, 1, rate, mode, cgst, sgst, igst, total,
                             self.me(conn), now_str(), "task" if b else "manual", b["id"] if b else None, name,
                             str(cl.get("gstin", "")).upper()[:15], str(cl.get("pan", "")).upper()[:10],
                             str(cl.get("address", ""))[:300], str(cl.get("email", ""))[:120],
                             re.sub(r"[^0-9+]", "", str(cl.get("phone", "")))[:15], str(hdr.get("notes", ""))[:500]))
            self.db.x("UPDATE invoices SET service_details=?, gst_type=? WHERE id=?",
                      (str(hdr.get("service_details", "")).strip()[:3000],
                       self.db.settings().get("gst_type") or "Regular", iid))
            for desc, sac, amt, tx in lines:
                self.db.x("INSERT INTO inv_lines(invoice_id, description, sac, amount, taxable) VALUES(?,?,?,?,?)",
                          (iid, desc, sac, amt, tx))
            if b:
                self.db.x("UPDATE billing SET status='Billed', ref=?, invoice_id=?, amount=?, updated_by=? WHERE id=?",
                          (inv_no, iid, taxable, self.me(conn), b["id"]))
            if hdr.get("save_client"):
                self._client_upsert(cl)
        self._lit_push()
        self._task_push()
        return {"id": iid, "inv_no": inv_no, "total": total}

    def h_invoice_get(self, conn, hdr, payload):
        self.need_admin(conn)
        inv = self.db.one("SELECT * FROM invoices WHERE id=?", (int(hdr.get("id", 0)),))
        if not inv:
            raise ValueError("Invoice not found")
        st = self.db.settings(with_logo=True)
        sac = st.get("sac_code") or "9982"
        case = self.db.one("SELECT * FROM cases WHERE id=?", (inv["case_id"],)) if inv["case_id"] else None
        lines = []
        if case:
            for f in self.db.q("SELECT * FROM fees WHERE invoice_id=? ORDER BY fdate", (inv["id"],)):
                lines.append({"description": f["description"], "date": f["fdate"], "sac": sac,
                              "amount": f["amount"], "taxable": 1})
            for o in self.db.q("SELECT * FROM ope WHERE invoice_id=? ORDER BY edate", (inv["id"],)):
                lines.append({"description": ("Out-of-pocket expense: " if not inv["pure_agent"] else "")
                              + o["description"], "date": o["edate"], "sac": sac if not inv["pure_agent"] else "",
                              "amount": o["amount"], "taxable": 0 if inv["pure_agent"] else 1})
            for k in ("name", "gstin", "pan", "address", "email", "phone"):   # fill old invoices from case
                if not inv.get("client_" + k):
                    inv["client_" + k] = case.get("client_" + k if k != "name" else "client_name", "") or ""
        else:
            for ln in self.db.q("SELECT * FROM inv_lines WHERE invoice_id=? ORDER BY id", (inv["id"],)):
                lines.append({"description": ln["description"], "date": "", "sac": ln["sac"] or sac,
                              "amount": ln["amount"], "taxable": ln["taxable"]})
        inv["outstanding"] = 0.0 if inv["status"] == "Cancelled" else round(inv["total"] - inv["received"] - inv["tds"], 2)
        inv["age_days"] = (date.today() - date.fromisoformat(inv["inv_date"])).days
        return {"invoice": inv, "case": case, "settings": st, "lines": lines,
                "receipts": self.db.q("SELECT * FROM receipts WHERE invoice_id=? ORDER BY rdate", (inv["id"],))}

    def h_invoice_list(self, conn, hdr, payload):
        self.need_admin(conn)
        sql = ("SELECT i.*, COALESCE(NULLIF(i.client_name,''), c.client_name) AS cname, c.law, c.period,"
               " COALESCE(NULLIF(i.client_email,''), c.client_email) AS cemail,"
               " COALESCE(NULLIF(i.client_phone,''), c.client_phone) AS cphone"
               " FROM invoices i LEFT JOIN cases c ON c.id=i.case_id WHERE 1=1")
        args = []
        if hdr.get("status") in ("Unpaid", "Part Paid", "Paid", "Cancelled"):
            sql += " AND i.status=?"
            args.append(hdr["status"])
        if hdr.get("source") in ("case", "task", "manual"):
            sql += " AND i.source=?"
            args.append(hdr["source"])
        rows = self.db.q(sql + " ORDER BY i.inv_date DESC, i.id DESC", args)
        today = date.today()
        fu = {f["invoice_id"]: f for f in self.db.q(
            "SELECT invoice_id, COUNT(*) AS n, MAX(fdate) AS last,"
            " (SELECT next_date FROM followups f2 WHERE f2.invoice_id=f.invoice_id ORDER BY f2.id DESC LIMIT 1) AS nxt,"
            " (SELECT promise_date FROM followups f3 WHERE f3.invoice_id=f.invoice_id ORDER BY f3.id DESC LIMIT 1) AS prm"
            " FROM followups f GROUP BY invoice_id")}
        out = []
        for r in rows:
            f = fu.get(r["id"]) or {}
            r["fu_count"], r["fu_last"] = f.get("n", 0), f.get("last") or ""
            r["fu_next"], r["fu_promise"] = f.get("nxt") or "", f.get("prm") or ""
            r["client_name"], r["client_email"], r["client_phone"] = r.pop("cname"), r.pop("cemail"), r.pop("cphone")
            r["outstanding"] = 0.0 if r["status"] == "Cancelled" else round(r["total"] - r["received"] - r["tds"], 2)
            r["age_days"] = (today - date.fromisoformat(r["inv_date"])).days
            r["law"] = r["law"] or ("Task billing" if r["source"] == "task" else "Professional services")
            r["fu_due"] = bool(r["outstanding"] > 0.009 and r["fu_next"] and r["fu_next"] <= today.isoformat())
            if hdr.get("overdue") and not (r["outstanding"] > 0.009 and r["age_days"] > INVOICE_OVERDUE_DAYS):
                continue
            if hdr.get("followup_due") and not r["fu_due"]:
                continue
            out.append(r)
        return out

    # ---------------------------------------------- payment follow-ups
    def h_followup_add(self, conn, hdr, payload):
        """Log a payment follow-up (call / e-mail / WhatsApp / visit), promised date and next follow-up date."""
        self.need_admin(conn)
        inv = self.db.one("SELECT * FROM invoices WHERE id=?", (int(hdr.get("invoice_id", 0)),))
        if not inv or inv["status"] == "Cancelled":
            raise ValueError("Invoice not found / cancelled")
        mode = hdr.get("mode") if hdr.get("mode") in FU_MODES else "Other"
        fd = self.parse_date(hdr.get("fdate") or date.today().isoformat(), "follow-up date")
        pd = self.parse_date(hdr["promise_date"], "promised date").isoformat() if hdr.get("promise_date") else ""
        nd = self.parse_date(hdr["next_date"], "next follow-up date").isoformat() if hdr.get("next_date") else ""
        if nd and nd < fd.isoformat():
            raise ValueError("Next follow-up date cannot be before the follow-up date")
        fid = self.db.x("INSERT INTO followups(invoice_id, fdate, mode, remarks, promise_date, next_date, by_user, ts)"
                        " VALUES(?,?,?,?,?,?,?,?)", (inv["id"], fd.isoformat(), mode,
                                                     str(hdr.get("remarks", "")).strip()[:1000], pd, nd,
                                                     self.me(conn), now_str()))
        self._lit_push()
        return {"id": fid}

    def h_followup_list(self, conn, hdr, payload):
        self.need_admin(conn)
        return self.db.q("SELECT f.*, u.full_name AS by_name FROM followups f LEFT JOIN users u ON u.username=f.by_user"
                         " WHERE f.invoice_id=? ORDER BY f.fdate DESC, f.id DESC", (int(hdr.get("invoice_id", 0)),))

    def h_followup_delete(self, conn, hdr, payload):
        self.need_admin(conn)
        self.db.x("DELETE FROM followups WHERE id=?", (int(hdr.get("id", 0)),))
        return {}

    def h_invoice_details(self, conn, hdr, payload):
        """Add / change the 'additional details of service' printed on an invoice."""
        self.need_admin(conn)
        inv = self.db.one("SELECT id FROM invoices WHERE id=?", (int(hdr.get("id", 0)),))
        if not inv:
            raise ValueError("Invoice not found")
        self.db.x("UPDATE invoices SET service_details=? WHERE id=?",
                  (str(hdr.get("service_details", "")).strip()[:3000], inv["id"]))
        return {"id": inv["id"]}

    def h_invoice_cancel(self, conn, hdr, payload):
        self.need_admin(conn)
        inv = self.db.one("SELECT * FROM invoices WHERE id=?", (int(hdr.get("id", 0)),))
        if not inv:
            raise ValueError("Invoice not found")
        if inv["received"] or inv["tds"]:
            raise ValueError("Receipts are recorded against this invoice; it cannot be cancelled")
        self.db.x("UPDATE invoices SET status='Cancelled' WHERE id=?", (inv["id"],))
        self.db.x("UPDATE fees SET invoice_id=NULL WHERE invoice_id=?", (inv["id"],))
        self.db.x("UPDATE ope SET invoice_id=NULL WHERE invoice_id=?", (inv["id"],))
        if inv.get("billing_id"):
            self.db.x("UPDATE billing SET status='To be Billed', ref='', invoice_id=NULL WHERE id=?",
                      (inv["billing_id"],))
        self._lit_push()
        self._task_push()
        return {}

    # ---------------------------------------------------- client master
    def _client_upsert(self, cl):
        name = str(cl.get("name", "")).strip()[:120]
        if not name:
            raise ValueError("Client name is required")
        vals = (str(cl.get("gstin", "")).upper().strip()[:15], str(cl.get("pan", "")).upper().strip()[:10],
                str(cl.get("address", ""))[:300], str(cl.get("email", "")).strip()[:120],
                re.sub(r"[^0-9+]", "", str(cl.get("phone", "")))[:15], str(cl.get("contact", ""))[:80])
        old = self.db.one("SELECT id FROM clients WHERE lower(name)=lower(?)", (name,))
        if old:
            self.db.x("UPDATE clients SET name=?, gstin=?, pan=?, address=?, email=?, phone=?, contact=? WHERE id=?",
                      (name,) + vals + (old["id"],))
            return old["id"]
        return self.db.x("INSERT INTO clients(name, gstin, pan, address, email, phone, contact, created_on)"
                         " VALUES(?,?,?,?,?,?,?,?)", (name,) + vals + (now_str(),))

    def h_firm_get(self, conn, hdr, payload):
        self.need_admin(conn)
        return self.db.settings(with_logo=True)

    def h_client_list(self, conn, hdr, payload):
        return self.db.q("SELECT * FROM clients ORDER BY name")

    def h_client_save(self, conn, hdr, payload):
        cl = hdr.get("client") or {}
        if cl.get("id"):
            old = self.db.one("SELECT * FROM clients WHERE id=?", (int(cl["id"]),))
            if not old:
                raise ValueError("Client not found")
            if old["name"].lower() != str(cl.get("name", "")).strip().lower() and \
                    self.db.one("SELECT 1 FROM clients WHERE lower(name)=lower(?)", (str(cl.get("name", "")).strip(),)):
                raise ValueError("Another client with this name already exists")
            self.db.x("UPDATE clients SET name=? WHERE id=?", (str(cl.get("name", "")).strip()[:120], old["id"]))
        return {"id": self._client_upsert(cl)}

    def h_client_delete(self, conn, hdr, payload):
        self.need_admin(conn)
        self.db.x("DELETE FROM clients WHERE id=?", (int(hdr.get("id", 0)),))
        return {}

    def h_receipt_add(self, conn, hdr, payload):
        self.need_admin(conn)
        inv = self.db.one("SELECT * FROM invoices WHERE id=?", (int(hdr.get("invoice_id", 0)),))
        if not inv or inv["status"] == "Cancelled":
            raise ValueError("Invoice not found / cancelled")
        amt = self._money(hdr.get("amount") or 0, "Amount received", True)
        tds = self._money(hdr.get("tds") or 0, "TDS", True)
        if amt + tds <= 0:
            raise ValueError("Enter amount received and/or TDS deducted")
        outstanding = round(inv["total"] - inv["received"] - inv["tds"], 2)
        if amt + tds > outstanding + 0.01:
            raise ValueError(f"Receipt exceeds outstanding amount Rs. {inr(outstanding)}")
        rd = self.parse_date(hdr.get("rdate"))
        fy = f"{rd.year}-{(rd.year + 1) % 100:02d}" if rd.month >= 4 else f"{rd.year - 1}-{rd.year % 100:02d}"
        with self.db.lock:
            seq = len(self.db.q("SELECT id FROM receipts WHERE rcpt_no LIKE ?", (f"RCPT/{fy}/%",))) + 1
            while self.db.one("SELECT 1 FROM receipts WHERE rcpt_no=?", (f"RCPT/{fy}/{seq:04d}",)):
                seq += 1
            rcpt_no = f"RCPT/{fy}/{seq:04d}"
            rcid = self.db.x("INSERT INTO receipts(invoice_id, rdate, amount, tds, mode, ref, by_user, ts, rcpt_no)"
                             " VALUES(?,?,?,?,?,?,?,?,?)",
                             (inv["id"], rd.isoformat(), amt, tds, str(hdr.get("mode", ""))[:40],
                              str(hdr.get("ref", ""))[:80], self.me(conn), now_str(), rcpt_no))
        status = self._recalc_invoice(inv["id"])
        return {"status": status, "id": rcid, "rcpt_no": rcpt_no}

    def _recalc_invoice(self, iid):
        """Re-total receipts of an invoice -> received, TDS, status (Unpaid / Part Paid / Paid)."""
        inv = self.db.one("SELECT * FROM invoices WHERE id=?", (iid,))
        r = self.db.one("SELECT COALESCE(SUM(amount),0) AS a, COALESCE(SUM(tds),0) AS t FROM receipts WHERE invoice_id=?",
                        (iid,))
        rec, td = round(r["a"], 2), round(r["t"], 2)
        status = "Unpaid" if rec + td <= 0.009 else ("Paid" if inv["total"] - rec - td <= 0.01 else "Part Paid")
        self.db.x("UPDATE invoices SET received=?, tds=?, status=? WHERE id=?", (rec, td, status, iid))
        if inv.get("billing_id"):
            self.db.x("UPDATE billing SET status=? WHERE id=?", ("Received" if status == "Paid" else "Billed",
                                                                 inv["billing_id"]))
            self._task_push()
        self._lit_push()
        return status

    def h_receipt_update(self, conn, hdr, payload):
        """Correct a payment already recorded (amount / TDS / date / mode / reference)."""
        self.need_admin(conn)
        rc = self.db.one("SELECT * FROM receipts WHERE id=?", (int(hdr.get("id", 0)),))
        if not rc:
            raise ValueError("Receipt not found")
        inv = self.db.one("SELECT * FROM invoices WHERE id=?", (rc["invoice_id"],))
        amt = self._money(hdr.get("amount") or 0, "Amount received", True)
        tds = self._money(hdr.get("tds") or 0, "TDS", True)
        if amt + tds <= 0:
            raise ValueError("Enter amount received and/or TDS deducted")
        others = self.db.one("SELECT COALESCE(SUM(amount+tds),0) AS s FROM receipts WHERE invoice_id=? AND id!=?",
                             (inv["id"], rc["id"]))["s"]
        if others + amt + tds > inv["total"] + 0.01:
            raise ValueError(f"Total receipts would exceed the invoice value Rs. {inr(inv['total'])}. "
                             f"Maximum for this receipt: Rs. {inr(inv['total'] - others)}")
        self.db.x("UPDATE receipts SET rdate=?, amount=?, tds=?, mode=?, ref=? WHERE id=?",
                  (self.parse_date(hdr.get("rdate")).isoformat(), amt, tds, str(hdr.get("mode", ""))[:40],
                   str(hdr.get("ref", ""))[:80], rc["id"]))
        return {"status": self._recalc_invoice(inv["id"]), "id": rc["id"], "rcpt_no": rc["rcpt_no"]}

    def h_receipt_delete(self, conn, hdr, payload):
        self.need_admin(conn)
        rc = self.db.one("SELECT * FROM receipts WHERE id=?", (int(hdr.get("id", 0)),))
        if not rc:
            raise ValueError("Receipt not found")
        self.db.x("DELETE FROM receipts WHERE id=?", (rc["id"],))
        return {"status": self._recalc_invoice(rc["invoice_id"])}

    def h_receipt_get(self, conn, hdr, payload):
        self.need_admin(conn)
        rc = self.db.one("SELECT * FROM receipts WHERE id=?", (int(hdr.get("id", 0)),))
        if not rc:
            raise ValueError("Receipt not found")
        d = self.h_invoice_get(conn, {"id": rc["invoice_id"]}, b"")
        upto = [x for x in d["receipts"] if x["id"] <= rc["id"]]
        rc["received_upto"] = round(sum(x["amount"] + x["tds"] for x in upto), 2)
        rc["balance_after"] = round(d["invoice"]["total"] - rc["received_upto"], 2)
        return {"receipt": rc, "invoice": d["invoice"], "settings": d["settings"]}

    def h_receipt_list(self, conn, hdr, payload):
        self.need_admin(conn)
        rows = self.db.q("SELECT r.*, i.inv_no, i.inv_date, i.total, i.case_id, i.client_name, i.client_email,"
                         " i.client_phone, c.client_name AS case_client FROM receipts r JOIN invoices i ON i.id=r.invoice_id"
                         " LEFT JOIN cases c ON c.id=i.case_id ORDER BY r.rdate DESC, r.id DESC")
        for r in rows:
            r["client_name"] = r["client_name"] or r.pop("case_client") or ""
        return rows

    def h_firm_logo(self, conn, hdr, payload):
        return {"logo": self.db.settings(with_logo=True).get("firm_logo", "")}

    # ---------------------------------------------------- case / task documents
    def _doc_check(self, conn, ref_type, ref_id):
        if ref_type == "case":
            self._case_v(conn, ref_id)
        elif ref_type == "task":
            t = self._task(ref_id)
            a = self._acc("task", t["id"], self.me(conn))
            if not (a and a["can_view"]):
                self._task_perm(conn, t)
        else:
            raise ValueError("Bad document type")

    def h_doc_upload(self, conn, hdr, payload):
        rt, rid = hdr.get("ref_type"), int(hdr.get("ref_id") or 0)
        self._doc_check(conn, rt, rid)
        if not payload:
            raise ValueError("Empty file")
        if len(payload) > 100 * 1024 * 1024:
            raise ValueError("File must be under 100 MB")
        fid = self.store_file(hdr.get("name"), payload, self.me(conn), "ALL", rt)
        did = self.db.x("INSERT INTO docs(ref_type, ref_id, file_id, title, by_user, ts) VALUES(?,?,?,?,?,?)",
                        (rt, rid, fid, str(hdr.get("title", ""))[:200], self.me(conn), now_str()))
        (self._lit_push if rt == "case" else self._task_push)()
        return {"id": did}

    def h_doc_list(self, conn, hdr, payload):
        rt, rid = hdr.get("ref_type"), int(hdr.get("ref_id") or 0)
        self._doc_check(conn, rt, rid)
        rows = self.db.q("SELECT d.*, f.name, f.size, u.full_name AS by_name, 'Uploaded document' AS source FROM docs d"
                         " JOIN files f ON f.id=d.file_id LEFT JOIN users u ON u.username=d.by_user"
                         " WHERE d.ref_type=? AND d.ref_id=? ORDER BY d.id DESC", (rt, rid))
        if rt == "case":     # also show attachments given with hearing / submission entries
            for p in self.db.q("SELECT p.id, p.file_id, p.ts, p.by_user, p.ptype, p.pdate, f.name, f.size,"
                               " u.full_name AS by_name FROM proceedings p JOIN files f ON f.id=p.file_id"
                               " LEFT JOIN users u ON u.username=p.by_user WHERE p.case_id=?", (rid,)):
                rows.append({"id": -p["id"], "file_id": p["file_id"], "name": p["name"], "size": p["size"],
                             "by_name": p["by_name"], "by_user": p["by_user"], "ts": p["ts"],
                             "title": f"{p['ptype']} ({dmy_s(p['pdate'])})", "source": "Hearing / submission entry"})
        return rows

    def h_doc_delete(self, conn, hdr, payload):
        d = self.db.one("SELECT * FROM docs WHERE id=?", (int(hdr.get("id", 0)),))
        if not d:
            raise ValueError("Document not found (hearing attachments are removed with their entry)")
        if not (self.is_admin(conn) or d["by_user"] == self.me(conn)):
            raise PermissionError("Only the person who uploaded it or the Manager can delete a document")
        self.db.x("DELETE FROM docs WHERE id=?", (d["id"],))
        (self._lit_push if d["ref_type"] == "case" else self._task_push)()
        return {}

    def h_task_get(self, conn, hdr, payload):
        t = self._task(hdr.get("id"))
        a = self._acc("task", t["id"], self.me(conn))
        if not (a and a["can_view"]):
            self._task_perm(conn, t)
        t["view_only"] = not (self.is_admin(conn) or self.me(conn) in (t["created_by"], t["assigned_to"]))
        t["assigned_name"], t["creator_name"] = self.name_of(t["assigned_to"]), self.name_of(t["created_by"])
        t["days_left"] = (date.fromisoformat(t["due_date"]) - date.today()).days
        t["colour"], t["priority"] = task_colour(t["days_left"], t["status"])
        return self._mask_task(conn, t)

    # ===================================================== MONTHLY PAYOUT
    def h_payout(self, conn, hdr, payload):
        """Salary + approved reimbursement for the month = total payout."""
        y, m = self.parse_month(hdr.get("month"))
        month = f"{y:04d}-{m:02d}"
        if hdr.get("username") == "ALL":
            self.need_admin(conn)
            users = [u["username"] for u in self.db.q("SELECT username FROM users WHERE active=1 ORDER BY full_name")]
        else:
            users = [self.target(conn, hdr)]
        rows = []
        for u in users:
            ps = self.h_payslip(conn, {"username": u, "month": month}, b"")
            pend = self.db.one("SELECT COALESCE(SUM(amount),0) AS s FROM reimb WHERE username=? AND status='Pending'"
                               " AND day LIKE ?", (u, month + "-%"))["s"]
            rows.append({"username": u, "full_name": ps["user"]["full_name"],
                         "designation": ps["user"]["designation"], "bank": ps["user"]["bank"],
                         "gross": ps["gross"], "lwp_days": ps["leave"]["lwp"], "lwp_amt": ps["deductions"][0][1],
                         "pt": ps["deductions"][1][1], "tds": ps["deductions"][2][1],
                         "net_salary": round(ps["gross"] - ps["total_deductions"], 2),
                         "reimb": ps["reimbursement"], "total": ps["net"], "pending_reimb": round(pend, 2)})
        if hdr.get("username") == "ALL":
            rows = [r for r in rows if r["gross"] or r["reimb"] or r["pending_reimb"]]
        return {"rows": rows, "month": month, "settings": self.db.settings(), "generated": now_str()}

    # ======================================================== TASK TRACKER
    def _task(self, tid):
        t = self.db.one("SELECT * FROM tasks WHERE id=?", (int(tid or 0),))
        if not t:
            raise ValueError("Task not found")
        return t

    def _task_perm(self, conn, t):
        if not (self.is_admin(conn) or self.me(conn) in (t["created_by"], t["assigned_to"])):
            raise PermissionError("Only the creator, the assignee or the Manager can change this task")

    def _task_push(self, text=None, users=(), exclude=None, ref=None):
        self.broadcast({"type": "push", "event": "task"})
        for u in set(users):
            if u and u != exclude and text:
                self.send_to(u, {"type": "push", "event": "notify", "text": text, "ref": ref})

    def _task_insert(self, d, creator):
        tid = self.db.x("INSERT INTO tasks(title, description, client, assigned_to, created_by, due_date, urgent,"
                        " series_id, created_on, est_amount) VALUES(?,?,?,?,?,?,?,?,?,?)",
                        (d["title"], d.get("description", ""), d.get("client", ""), d["assigned_to"], creator,
                         d["due_date"], 1 if d.get("urgent") else 0, d.get("series_id"), now_str(),
                         float(d.get("est_amount") or 0)))
        due = date.fromisoformat(d["due_date"]).strftime("%d-%m-%Y")
        self._task_push(f"New task for you: '{d['title']}'" + (f" ({d['client']})" if d.get("client") else "")
                        + f" — due {due}" + (" — URGENT" if d.get("urgent") else "")
                        + f" — from {self.name_of(creator)}", [d["assigned_to"]], exclude=creator,
                        ref={"kind": "task", "id": tid})
        return tid

    def _gen_recurring(self):
        """Create this period's task for every active recurring series (GST / TDS returns etc.)."""
        today = date.today()
        with self.db.lock:
            for s in self.db.q("SELECT * FROM task_series WHERE active=1"):
                nd, made = date.fromisoformat(s["next_due"]), 0
                while date(nd.year, nd.month, 1) <= today and made < 12:
                    if self.db.one("SELECT 1 FROM tasks WHERE series_id=? AND due_date=?", (s["id"], nd.isoformat())):
                        nd = add_months(nd, REPEAT_MONTHS.get(s["frequency"], 1), s["day"])
                        made += 1
                        continue
                    self._task_insert(dict(title=s["title"], description=s["description"], client=s["client"],
                                           assigned_to=s["assigned_to"], due_date=nd.isoformat(),
                                           urgent=s["urgent"], series_id=s["id"], est_amount=s["est_amount"]),
                                      s["created_by"])
                    nd = add_months(nd, REPEAT_MONTHS.get(s["frequency"], 1), s["day"])
                    made += 1
                if made:
                    self.db.x("UPDATE task_series SET next_due=? WHERE id=?", (nd.isoformat(), s["id"]))

    def _recurring_loop(self):
        while self.running:
            try:
                self._gen_recurring()
            except Exception:
                pass
            for _ in range(1800):          # every 30 minutes
                if not self.running:
                    return
                time.sleep(1)

    def h_task_save(self, conn, hdr, payload):
        d = hdr.get("task") or {}
        me = self.me(conn)
        title = str(d.get("title", "")).strip()[:200]
        if not title:
            raise ValueError("Task title is required")
        assigned = str(d.get("assigned_to") or me)
        if not self.db.one("SELECT 1 FROM users WHERE username=? AND active=1", (assigned,)):
            raise ValueError("Assigned team member not found")
        due = self.parse_date(d.get("due_date"), "due date")
        if hdr.get("id"):
            _t = self._task(hdr.get("id"))
            if not self._task_fee_visible(conn, _t):
                d["est_amount"] = _t["est_amount"]     # member cannot see / change a confidential fee
        vals = dict(title=title, description=str(d.get("description", ""))[:2000],
                    client=str(d.get("client", "")).strip()[:120], assigned_to=assigned,
                    due_date=due.isoformat(), urgent=1 if d.get("urgent") else 0,
                    est_amount=self._money(d.get("est_amount") or 0, "Estimated billing", True))
        tid = hdr.get("id")
        if tid:
            t = self._task(tid)
            self._task_perm(conn, t)
            self.db.x("UPDATE tasks SET " + ", ".join(f"{k}=?" for k in vals) + " WHERE id=?",
                      list(vals.values()) + [t["id"]])
            if assigned != t["assigned_to"]:
                self._task_push(f"Task '{title}' has been assigned to you by {conn.user['full_name']} — due "
                                f"{due:%d-%m-%Y}", [assigned], exclude=me,
                                ref={"kind": "task", "id": t["id"]})
            else:
                self._task_push(f"Task '{title}' was updated by {conn.user['full_name']}",
                                [t["assigned_to"], t["created_by"]], exclude=me, ref={"kind": "task", "id": t["id"]})
            return {"id": t["id"]}
        repeat = d.get("repeat") if d.get("repeat") in TASK_REPEATS else "None"
        sid = None
        if repeat != "None":
            sid = self.db.x("INSERT INTO task_series(title, description, client, assigned_to, created_by, frequency,"
                            " day, urgent, next_due, created_on, est_amount) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (title, vals["description"], vals["client"], assigned, me, repeat, due.day,
                             vals["urgent"], add_months(due, REPEAT_MONTHS[repeat]).isoformat(), now_str(),
                             vals["est_amount"]))
        vals["series_id"] = sid
        return {"id": self._task_insert(vals, me), "series_id": sid}

    def h_task_list(self, conn, hdr, payload):
        self._gen_recurring()
        me, scope = self.me(conn), hdr.get("scope", "mine")
        sql = ("SELECT t.*, a.full_name AS assigned_name, c.full_name AS creator_name FROM tasks t"
               " LEFT JOIN users a ON a.username=t.assigned_to LEFT JOIN users c ON c.username=t.created_by WHERE 1=1")
        args = []
        if scope == "mine":
            sql += " AND t.assigned_to=?"
            args.append(me)
        elif scope == "byme":
            sql += " AND t.created_by=? AND t.assigned_to!=?"
            args += [me, me]
        elif scope == "all":
            self.need_admin(conn)
        elif scope == "shared":
            sql += (" AND t.id IN (SELECT ref_id FROM item_access WHERE ref_type='task' AND username=? AND can_view=1)"
                    " AND t.assigned_to!=? AND t.created_by!=?")
            args += [me, me, me]
        else:
            raise ValueError("Bad scope")
        st = hdr.get("status", "Open")
        if st == "Open":
            sql += " AND t.status IN ('Open','In Progress')"
        elif st == "Completed":
            sql += " AND t.status='Completed'"
        elif st == "Cancelled":
            sql += " AND t.status='Cancelled'"
        rows = self.db.q(sql, args)
        today = date.today()
        for r in rows:
            r["days_left"] = (date.fromisoformat(r["due_date"]) - today).days
            r["colour"], r["priority"] = task_colour(r["days_left"], r["status"])
            self._mask_task(conn, r)
        done = lambda r: r["status"] in ("Completed", "Cancelled")
        rows.sort(key=lambda r: (done(r), -r["urgent"] if not done(r) else 0, r["due_date"]))
        return rows

    def h_task_status(self, conn, hdr, payload):
        t = self._task(hdr.get("id"))
        self._task_perm(conn, t)
        st = hdr.get("status")
        if st not in ("Open", "In Progress", "Cancelled"):
            raise ValueError("Bad status")
        if t["status"] == "Completed":
            raise ValueError("Task is already completed")
        self.db.x("UPDATE tasks SET status=? WHERE id=?", (st, t["id"]))
        self._task_push(f"Task '{t['title']}' marked {st} by {conn.user['full_name']}",
                        [t["assigned_to"], t["created_by"]], exclude=self.me(conn),
                        ref={"kind": "task", "id": t["id"]})
        return {"id": t["id"]}

    def h_task_complete(self, conn, hdr, payload):
        t = self._task(hdr.get("id"))
        self._task_perm(conn, t)
        if t["status"] in ("Completed", "Cancelled"):
            raise ValueError(f"Task is already {t['status']}")
        opt = hdr.get("billing_option")
        if opt not in BILL_OPTIONS:
            raise ValueError("Select a billing option")
        amt = 0.0 if opt == "No Billing" else self._money(hdr.get("billing_amount") or 0, "Billing amount", True)
        if opt != "No Billing" and not self._task_fee_visible(conn, t):
            amt = float(t["est_amount"] or 0)          # fee is confidential: use Manager's estimate
        ch = str(hdr.get("challenges", "")).strip()[:2000]
        chall = t["challenges"] + (f"\n[{now_str()[:16]} {conn.user['full_name']}] {ch}" if ch else "")
        self.db.x("UPDATE tasks SET status='Completed', completed_on=?, completed_by=?, billing_option=?,"
                  " billing_amount=?, challenges=? WHERE id=?",
                  (now_str(), self.me(conn), opt, amt, chall.strip(), t["id"]))
        if opt != "No Billing":
            self.db.x("INSERT INTO billing(task_id, client, description, amount, btype, status, created_on, updated_by)"
                      " VALUES(?,?,?,?,?,?,?,?)",
                      (t["id"], t["client"], t["title"], amt, opt, "Billed" if opt == "Billing Raised" else "To be Billed",
                       now_str(), self.me(conn)))
        msg = (f"Task completed: '{t['title']}'" + (f" ({t['client']})" if t["client"] else "") +
               f" by {conn.user['full_name']} — {opt}" + (f" Rs. {inr(amt)}" if amt else ""))
        self._task_push(msg, [t["created_by"]] + self._admins(), exclude=self.me(conn),
                        ref={"kind": "task", "id": t["id"]})
        return {"id": t["id"], "billing": opt != "No Billing"}

    def h_task_remark(self, conn, hdr, payload):
        t = self._task(hdr.get("id"))
        self._task_perm(conn, t)
        txt = str(hdr.get("text", "")).strip()[:2000]
        if not txt:
            raise ValueError("Remark is empty")
        chall = (t["challenges"] + f"\n[{now_str()[:16]} {conn.user['full_name']}] {txt}").strip()
        self.db.x("UPDATE tasks SET challenges=? WHERE id=?", (chall, t["id"]))
        self._task_push(f"Remark on task '{t['title']}' by {conn.user['full_name']}: {txt[:120]}",
                        [t["assigned_to"], t["created_by"]], exclude=self.me(conn),
                        ref={"kind": "task", "id": t["id"]})
        return {"id": t["id"]}

    def h_task_delete(self, conn, hdr, payload):
        t = self._task(hdr.get("id"))
        if not (self.is_admin(conn) or t["created_by"] == self.me(conn)):
            raise PermissionError("Only the creator or the Manager can delete a task")
        if self.db.one("SELECT 1 FROM billing WHERE task_id=? AND status!='To be Billed'", (t["id"],)):
            raise ValueError("Billing has been raised for this task; it cannot be deleted")
        self.db.x("DELETE FROM billing WHERE task_id=?", (t["id"],))
        self.db.x("DELETE FROM tasks WHERE id=?", (t["id"],))
        self._task_push()
        return {}

    def h_series_list(self, conn, hdr, payload):
        sql = ("SELECT s.*, a.full_name AS assigned_name, c.full_name AS creator_name FROM task_series s"
               " LEFT JOIN users a ON a.username=s.assigned_to LEFT JOIN users c ON c.username=s.created_by"
               " WHERE s.active=1")
        args = []
        if not self.is_admin(conn):
            sql += " AND (s.created_by=? OR s.assigned_to=?)"
            args = [self.me(conn)] * 2
        return self.db.q(sql + " ORDER BY s.next_due", args)

    def h_series_stop(self, conn, hdr, payload):
        s = self.db.one("SELECT * FROM task_series WHERE id=?", (int(hdr.get("id", 0)),))
        if not s:
            raise ValueError("Recurring task not found")
        if not (self.is_admin(conn) or s["created_by"] == self.me(conn)):
            raise PermissionError("Only the creator or the Manager can stop a recurring task")
        self.db.x("UPDATE task_series SET active=0 WHERE id=?", (s["id"],))
        self._task_push()
        return {}

    def h_billing_list(self, conn, hdr, payload):
        sql = ("SELECT b.*, t.title, t.assigned_to, t.completed_on, t.est_amount, u.full_name AS done_by,"
               " i.inv_no, i.status AS inv_status, i.total AS inv_total, i.received AS inv_received,"
               " i.tds AS inv_tds, i.inv_date, i.client_name AS inv_client FROM billing b LEFT JOIN tasks t ON t.id=b.task_id"
               " LEFT JOIN users u ON u.username=t.completed_by LEFT JOIN invoices i ON i.id=b.invoice_id WHERE 1=1")
        args = []
        if not self.is_admin(conn):
            sql += " AND (t.assigned_to=? OR t.created_by=?)"
            args = [self.me(conn)] * 2
        if hdr.get("status") in BILL_STATUSES:
            sql += " AND b.status=?"
            args.append(hdr["status"])
        if hdr.get("btype") in BILL_OPTIONS:
            sql += " AND b.btype=?"
            args.append(hdr["btype"])
        rows = self.db.q(sql + " ORDER BY CASE b.status WHEN 'To be Billed' THEN 0 ELSE 1 END, b.id DESC", args)
        if not self.is_admin(conn):
            for b in rows:
                t = self.db.one("SELECT * FROM tasks WHERE id=?", (b["task_id"],))
                if t and not self._task_fee_visible(conn, t):
                    for k in ("amount", "est_amount", "inv_total", "inv_received", "inv_tds"):
                        b[k] = None
                    b["fee_hidden"] = True
        return rows

    def h_billing_update(self, conn, hdr, payload):
        self.need_admin(conn)
        b = self.db.one("SELECT * FROM billing WHERE id=?", (int(hdr.get("id", 0)),))
        if not b:
            raise ValueError("Billing entry not found")
        st = hdr.get("status") if hdr.get("status") in BILL_STATUSES else b["status"]
        amt = self._money(hdr.get("amount", b["amount"]) or 0, "Amount", True)
        self.db.x("UPDATE billing SET status=?, ref=?, amount=?, updated_by=? WHERE id=?",
                  (st, str(hdr.get("ref", b["ref"]))[:80], amt, self.me(conn), b["id"]))
        self._task_push()
        return {"id": b["id"]}


def discover_servers(timeout=1.5):
    found = {}
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        s.settimeout(0.3)
        s.sendto(b"SUTHAR_DISCOVER", ("255.255.255.255", DISCOVERY_PORT))
        try:
            s.sendto(b"SUTHAR_DISCOVER", ("127.0.0.1", DISCOVERY_PORT))
        except Exception:
            pass
        end = time.time() + timeout
        while time.time() < end:
            try:
                data, addr = s.recvfrom(1024)
                parts = data.decode(errors="ignore").split("|")
                if parts[0] == "SUTHAR_SERVER":
                    found[addr[0]] = (addr[0], int(parts[1]), parts[2] if len(parts) > 2 else "")
            except socket.timeout:
                continue
        s.close()
    except Exception:
        pass
    return list(found.values())


# =============================================================================
# 7. CLIENT : NETWORK
# =============================================================================
class NetClient(QObject):
    push = pyqtSignal(object, object)       # header, payload
    disconnected = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.sock = None
        self.sendq = queue.Queue()
        self.pending = {}
        self.plock = threading.Lock()
        self.rid = 0
        self.alive = False
        self.audio_hook = None
        self.frames_pending = 0
        self.host = ""

    def connect_to(self, host, port, timeout=6):
        self.host = host
        self.sock = socket.create_connection((host, port), timeout=timeout)
        self.sock.settimeout(None)
        self.sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
        self.alive = True
        threading.Thread(target=self._reader, daemon=True).start()
        threading.Thread(target=self._writer, daemon=True).start()

    def _writer(self):
        while self.alive:
            item = self.sendq.get()
            if item is None:
                break
            hdr, payload = item
            try:
                send_frame(self.sock, hdr, payload)
            except Exception as e:
                self._fail(str(e))
                break
            finally:
                if hdr.get("kind") == "frame":
                    self.frames_pending = max(0, self.frames_pending - 1)

    def _reader(self):
        try:
            while self.alive:
                hdr, payload = recv_frame(self.sock)
                if hdr.get("type") == "resp":
                    with self.plock:
                        slot = self.pending.pop(hdr.get("rid"), None)
                    if slot:
                        slot["res"] = (hdr, payload)
                        slot["ev"].set()
                elif hdr.get("event") == "relay" and hdr.get("kind") == "audio":
                    if self.audio_hook:
                        try:
                            self.audio_hook(hdr, payload)
                        except Exception:
                            pass
                else:
                    self.push.emit(hdr, payload)
        except Exception as e:
            self._fail(str(e))

    def _fail(self, why):
        if not self.alive:
            return
        self.alive = False
        with self.plock:
            for slot in self.pending.values():
                slot["ev"].set()
            self.pending.clear()
        try:
            self.sock.close()
        except Exception:
            pass
        try:
            self.disconnected.emit(why)
        except RuntimeError:   # window already destroyed (app closing)
            pass

    def send(self, type_, payload=b"", **kw):
        if not self.alive:
            return
        hdr = dict(kw, type=type_)
        if hdr.get("kind") == "frame":
            self.frames_pending += 1
        self.sendq.put((hdr, payload))

    def request(self, type_, payload=b"", timeout=25, **kw):
        if not self.alive:
            raise ConnectionError("Not connected to server")
        with self.plock:
            self.rid += 1
            rid = self.rid
            slot = {"ev": threading.Event(), "res": None}
            self.pending[rid] = slot
        self.sendq.put((dict(kw, type=type_, rid=rid), payload))
        if not slot["ev"].wait(timeout):
            with self.plock:
                self.pending.pop(rid, None)
            raise TimeoutError("Server did not respond in time")
        if slot["res"] is None:
            raise ConnectionError("Connection to server lost")
        hdr, out = slot["res"]
        if not hdr.get("ok"):
            raise RuntimeError(hdr.get("error", "Request failed"))
        return hdr.get("data"), out

    def call(self, type_, **kw):
        """Request returning only data."""
        return self.request(type_, **kw)[0]

    def close(self):
        self.alive = False
        self.sendq.put(None)
        try:
            self.sock.shutdown(socket.SHUT_RDWR)
        except Exception:
            pass
        try:
            self.sock.close()
        except Exception:
            pass


class Task(QThread):
    done = pyqtSignal(object, object)

    def __init__(self, fn):
        super().__init__()
        self.fn = fn

    def run(self):
        try:
            self.done.emit(self.fn(), None)
        except Exception as e:
            self.done.emit(None, e)


_TASKS = set()


def run_task(fn, callback):
    """Run fn() in background thread; callback(result, error) on GUI thread."""
    t = Task(fn)
    _TASKS.add(t)

    def finished(res, err):
        _TASKS.discard(t)
        callback(res, err)
    t.done.connect(finished)
    t.start()
    return t


# =============================================================================
# 8. PDF REPORTS
# =============================================================================
# Letterhead palette: deep teal-navy firm name, orange accent line, light grey-blue band
PDF_PURPLE = rl_colors.HexColor("#1f4e5f")      # navy (headings, table headers, totals)
PDF_PINK = rl_colors.HexColor("#e67e22")        # orange accent
PDF_LIGHT = rl_colors.HexColor("#f3f6f8")       # zebra rows
PDF_BAND = rl_colors.HexColor("#eef2f5")        # letterhead band
PDF_HI = "#fdf0e3"                              # light orange highlight
PDF_BRAND = {"logo": ""}                        # firm logo cache (filled by the client at login)


def _pdf_styles():
    ss = getSampleStyleSheet()
    return {
        "co": ParagraphStyle("co", parent=ss["Title"], fontSize=20, textColor=PDF_PURPLE, spaceAfter=2),
        "addr": ParagraphStyle("addr", parent=ss["Normal"], fontSize=9, alignment=1,
                               textColor=rl_colors.HexColor("#555555")),
        "title": ParagraphStyle("t", parent=ss["Heading2"], alignment=1, textColor=rl_colors.white,
                                backColor=PDF_PINK, borderPadding=5, spaceBefore=8, spaceAfter=10),
        "n": ParagraphStyle("n", parent=ss["Normal"], fontSize=9, leading=12),
        "b": ParagraphStyle("b", parent=ss["Normal"], fontSize=9.5, leading=13, fontName="Helvetica-Bold"),
        "small": ParagraphStyle("s", parent=ss["Normal"], fontSize=7.5,
                                textColor=rl_colors.HexColor("#777777"), alignment=1),
    }


def _spaced(text):
    """C H A R T E R E D   A C C O U N T A N T S  (letter-spaced tagline)."""
    return "&nbsp;&nbsp;&nbsp;".join("&nbsp;".join(html.escape(ch) for ch in w) for w in str(text).split())


def _header(width, story, st, s, title):
    """Firm letterhead used on every PDF: logo + firm name + tagline on a light band, orange rule,
    contact strip and the document title."""
    from reportlab.platypus.flowables import HRFlowable
    E = lambda x: html.escape(str(x or ""))
    firm = (s.get("company_name") or FIRM_NAME).upper()
    tag = s.get("firm_tagline", "Chartered Accountants")
    name_st = ParagraphStyle("lhn", parent=st["n"], alignment=2, fontName="Times-Bold", fontSize=22, leading=27,
                             textColor=PDF_PURPLE)
    tag_st = ParagraphStyle("lht", parent=st["n"], alignment=2, fontName="Helvetica-Bold", fontSize=8.5,
                            leading=14, textColor=PDF_PINK)
    right = [Paragraph(E(firm), name_st)] + ([Paragraph(_spaced(tag.upper()), tag_st)] if tag else [])
    logo = _logo_flowable(s.get("firm_logo") or PDF_BRAND.get("logo"), max_h=19 * mm, max_w=34 * mm)
    lw = 40 * mm if logo else 0
    band = Table([[logo, right]] if logo else [[right]], colWidths=[lw, width - lw] if logo else [width])
    style = [("BACKGROUND", (0, 0), (-1, -1), PDF_BAND), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
             ("RIGHTPADDING", (-1, 0), (-1, 0), 12), ("LEFTPADDING", (0, 0), (0, 0), 8)]
    if logo:
        style += [("BACKGROUND", (0, 0), (0, 0), rl_colors.white), ("ALIGN", (0, 0), (0, 0), "CENTER")]
    band.setStyle(TableStyle(style))
    story += [band, HRFlowable(width="100%", thickness=2.2, color=PDF_PINK, spaceBefore=5, spaceAfter=3)]
    addr = ", ".join(x for x in (s.get("company_address"), s.get("firm_city"), s.get("firm_state")) if x)
    bits = [E(addr)] if addr else []
    bits += [x for x in (f"Ph: {E(s.get('firm_phone'))}" if s.get("firm_phone") else "",
                         E(s.get("firm_email")), E(s.get("firm_website"))) if x]
    ids = [x for x in (f"GSTIN: {E(s.get('firm_gstin'))}" if s.get("firm_gstin") else "",
                       f"PAN: {E(s.get('firm_pan'))}" if s.get("firm_pan") else "",
                       f"FRN: {E(s.get('firm_frn'))}" if s.get("firm_frn") else "") if x]
    strip = " &nbsp;|&nbsp; ".join(bits) + ("<br/>" + " &nbsp;|&nbsp; ".join(ids) if ids else "")
    if strip:
        story.append(Paragraph(strip, ParagraphStyle("lhs", parent=st["n"], alignment=1, fontSize=8, leading=10.5,
                                                     textColor=rl_colors.HexColor("#5b6b75"))))
    tt = Table([[Paragraph(E(title), ParagraphStyle("ttl", parent=st["n"], alignment=1, fontName="Helvetica-Bold",
                                                    fontSize=12, leading=15, textColor=PDF_PURPLE))]],
               colWidths=[width])
    tt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), PDF_BAND), ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                            ("LINEBELOW", (0, 0), (-1, -1), 1.2, PDF_PINK)]))
    story += [Spacer(1, 6), tt, Spacer(1, 8)]


def _grid(data, widths, header=True, zebra=True):
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    style = [("FONT", (0, 0), (-1, -1), "Helvetica", 9),
             ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#cfd9df")),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    if header:
        style += [("BACKGROUND", (0, 0), (-1, 0), PDF_PURPLE),
                  ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
                  ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9)]
    if zebra:
        for i in range(1 if header else 0, len(data)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), PDF_LIGHT))
    t.setStyle(TableStyle(style))
    return t


def month_label(ym):
    y, m = map(int, ym.split("-"))
    return f"{calendar.month_name[m]} {y}"


def make_payslip_pdf(d, path):
    st = _pdf_styles()
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm,
                            topMargin=14 * mm, bottomMargin=14 * mm,
                            title=f"Salary Slip {d['month']}", author=APP_NAME)
    u, lv = d["user"], d["leave"]
    story = []
    _header(doc.width, story, st, d["settings"], f"SALARY SLIP FOR THE MONTH OF {month_label(d['month']).upper()}")
    P = lambda s: Paragraph(html.escape(str(s if s not in (None, "") else "-")), st["n"])
    B = lambda s: Paragraph(html.escape(str(s)), st["b"])
    info = [[B("Employee Name"), P(u["full_name"]), B("Employee ID"), P(u["username"])],
            [B("Designation"), P(u["designation"]), B("Department"), P(u["department"])],
            [B("PAN"), P(u["pan"]), B("Bank A/c No."), P(u["bank"])],
            [B("Date of Joining"), P(u["doj"]), B("Days in Month"), P(d["days_in_month"])],
            [B("Paid Days"), P(f"{d['paid_days']:g}"), B("Days Present (attendance)"), P(d["present_days"])]]
    story += [_grid(info, [36 * mm, 51 * mm, 40 * mm, 51 * mm], header=False), Spacer(1, 8)]
    rows = [["Earnings", "Amount (Rs.)", "Deductions", "Amount (Rs.)"]]
    n = max(len(d["earnings"]), len(d["deductions"]))
    for i in range(n):
        e = d["earnings"][i] if i < len(d["earnings"]) else ["", ""]
        k = d["deductions"][i] if i < len(d["deductions"]) else ["", ""]
        rows.append([e[0], inr(e[1]) if e[0] else "", k[0], inr(k[1]) if k[0] else ""])
    rows.append(["Gross Earnings", inr(d["gross"]), "Total Deductions", inr(d["total_deductions"])])
    t = _grid(rows, [50 * mm, 39 * mm, 50 * mm, 39 * mm])
    t.setStyle(TableStyle([("ALIGN", (1, 1), (1, -1), "RIGHT"), ("ALIGN", (3, 1), (3, -1), "RIGHT"),
                           ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9),
                           ("BACKGROUND", (0, -1), (-1, -1), rl_colors.HexColor(PDF_HI))]))
    story += [t, Spacer(1, 8)]
    net = [["Gross Earnings", inr(d["gross"])], ["Less: Total Deductions", inr(d["total_deductions"])],
           ["Add: Approved Reimbursements", inr(d["reimbursement"])], ["NET PAY", inr(d["net"])]]
    t = _grid(net, [139 * mm, 39 * mm], header=False, zebra=False)
    t.setStyle(TableStyle([("ALIGN", (1, 0), (1, -1), "RIGHT"),
                           ("BACKGROUND", (0, -1), (-1, -1), PDF_PURPLE),
                           ("TEXTCOLOR", (0, -1), (-1, -1), rl_colors.white),
                           ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 11)]))
    story += [t, Spacer(1, 4), Paragraph("<b>Net Pay in words:</b> " + html.escape(d["net_words"]), st["n"]),
              Spacer(1, 10)]
    lvrows = [["Leave Month", "Paid Quota", "Availed", "Paid Used", "LWP Days", "Paid Balance"],
              [month_label(d["month"]), f"{lv['quota']:g}", f"{lv['taken']:g}", f"{lv['paid']:g}",
               f"{lv['lwp']:g}", f"{lv['balance']:g}"]]
    story += [_grid(lvrows, [38 * mm, 28 * mm, 26 * mm, 28 * mm, 28 * mm, 30 * mm]), Spacer(1, 30)]
    sig = Table([["Employee Signature", "", "For " + (d["settings"].get("company_name") or FIRM_NAME)],
                 ["", "", "Authorised Signatory"]], colWidths=[60 * mm, 58 * mm, 60 * mm])
    sig.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 9),
                             ("LINEABOVE", (0, 0), (0, 0), 0.6, rl_colors.black),
                             ("ALIGN", (2, 0), (2, -1), "RIGHT")]))
    story += [sig, Spacer(1, 16),
              Paragraph(f"This is a computer-generated salary slip. Generated on {d['generated']} "
                        f"via {APP_NAME}.", st["small"])]
    doc.build(story)
    return path


def make_reimb_pdf(rows, settings, month, member_label, path):
    st = _pdf_styles()
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm,
                            topMargin=14 * mm, bottomMargin=14 * mm,
                            title=f"Reimbursement Claim {month}", author=APP_NAME)
    story = []
    _header(doc.width, story, st, settings, f"REIMBURSEMENT CLAIM STATEMENT - {month_label(month).upper()}")
    story += [Paragraph(f"<b>Claimant:</b> {html.escape(member_label)} &nbsp;&nbsp;&nbsp; "
                        f"<b>Period:</b> {month_label(month)} &nbsp;&nbsp;&nbsp; "
                        f"<b>Prepared on:</b> {datetime.now().strftime('%d-%m-%Y %H:%M')}", st["n"]),
              Spacer(1, 8)]
    multi = len({r["username"] for r in rows}) > 1
    head = ["#", "Date"] + (["Member"] if multi else []) + ["Category", "Description", "Bill", "Status",
                                                            "Amount (Rs.)"]
    data = [head]
    claimed = approved = 0.0
    bycat = {c: [0.0, 0.0] for c in REIMB_CATEGORIES}
    for i, r in enumerate(rows, 1):
        d = date.fromisoformat(r["day"]).strftime("%d-%m-%Y")
        line = [str(i), d] + ([Paragraph(html.escape(r.get("full_name") or r["username"]), st["n"])]
                              if multi else [])
        line += [r["category"], Paragraph(html.escape(r["description"]), st["n"]),
                 "Yes" if r.get("file_id") else "No", r["status"], inr(r["amount"])]
        data.append(line)
        if r["status"] != "Rejected":
            claimed += r["amount"]
            bycat.setdefault(r["category"], [0.0, 0.0])[0] += r["amount"]
        if r["status"] == "Approved":
            approved += r["amount"]
            bycat[r["category"]][1] += r["amount"]
    if len(data) == 1:
        data.append(["-"] * len(head))
    data.append([""] * (len(head) - 2) + ["TOTAL*", inr(claimed)])
    widths = ([8 * mm, 20 * mm] + ([30 * mm] if multi else []) +
              [26 * mm, (58 if multi else 88) * mm, 11 * mm, 20 * mm, 27 * mm])
    t = _grid(data, widths)
    t.setStyle(TableStyle([("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
                           ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9),
                           ("BACKGROUND", (0, -1), (-1, -1), rl_colors.HexColor(PDF_HI))]))
    story += [t, Paragraph("* Total excludes rejected claims.", st["small"]), Spacer(1, 10)]
    cat = [["Category-wise Summary", "Claimed (Rs.)", "Approved (Rs.)"]]
    for c, (a, b) in bycat.items():
        cat.append([c, inr(a), inr(b)])
    cat.append(["TOTAL", inr(claimed), inr(approved)])
    t = _grid(cat, [80 * mm, 45 * mm, 45 * mm])
    t.setStyle(TableStyle([("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                           ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9)]))
    story += [t, Spacer(1, 6),
              Paragraph("<b>Amount claimed in words:</b> " + amount_in_words(claimed), st["n"]),
              Spacer(1, 8),
              Paragraph("Declaration: I hereby certify that the above expenses were actually incurred "
                        "for official purposes of the firm and have not been claimed earlier.", st["n"]),
              Spacer(1, 34)]
    sig = Table([["Claimant Signature", "", "Approved by (Manager)"]], colWidths=[60 * mm, 62 * mm, 60 * mm])
    sig.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 9),
                             ("LINEABOVE", (0, 0), (0, 0), 0.6, rl_colors.black),
                             ("LINEABOVE", (2, 0), (2, 0), 0.6, rl_colors.black),
                             ("ALIGN", (2, 0), (2, 0), "CENTER")]))
    story += [sig, Spacer(1, 14), Paragraph(f"Generated via {APP_NAME}.", st["small"])]
    doc.build(story)
    return path, claimed, approved


def dmy(iso):
    try:
        return date.fromisoformat(str(iso)[:10]).strftime("%d-%m-%Y")
    except Exception:
        return str(iso or "")


def _logo_flowable(b64, max_h=22 * mm, max_w=45 * mm):
    """Firm logo (base64 PNG/JPG from Firm Master) as a reportlab Image, or None."""
    if not b64:
        return None
    try:
        import base64
        import io
        from reportlab.platypus import Image as RLImage
        from reportlab.lib.utils import ImageReader
        raw = base64.b64decode(b64)
        iw, ih = ImageReader(io.BytesIO(raw)).getSize()
        scale = min(max_h / ih, max_w / iw)
        return RLImage(io.BytesIO(raw), width=iw * scale, height=ih * scale)
    except Exception:
        return None


def firm_bank_text(s):
    parts = []
    if s.get("bank_name"):
        parts.append(f"Bank: {s['bank_name']}")
    if s.get("bank_acc"):
        parts.append(f"A/c No.: {s['bank_acc']}")
    if s.get("bank_ifsc"):
        parts.append(f"IFSC: {s['bank_ifsc']}")
    if s.get("bank_branch"):
        parts.append(f"Branch: {s['bank_branch']}")
    if s.get("upi_id"):
        parts.append(f"UPI: {s['upi_id']}")
    return " | ".join(parts) or (s.get("firm_bank") or "")


def _sign_block(s, firm, width, bold_colour=None):
    """'For <firm>' + signature image (if uploaded & enabled in Firm Master) + signatory name."""
    img = None
    if s.get("sign_auto", "1") == "1" and s.get("firm_sign"):
        img = _logo_flowable(s["firm_sign"], max_h=17 * mm, max_w=48 * mm)
    rows = [["", "For " + firm], ["", img if img else ""], ["", s.get("signatory") or "Authorised Signatory"]]
    t = Table(rows, colWidths=[width - 70 * mm, 70 * mm])
    st = [("FONT", (0, 0), (-1, -1), "Helvetica", 9), ("ALIGN", (1, 0), (1, -1), "RIGHT"),
          ("FONT", (1, 0), (1, 0), "Helvetica-Bold", 9.5), ("VALIGN", (0, 0), (-1, -1), "BOTTOM")]
    if bold_colour is not None:
        st.append(("TEXTCOLOR", (1, 0), (1, 0), bold_colour))
    if not img:
        st.append(("TOPPADDING", (0, 1), (-1, 1), 22))
    t.setStyle(TableStyle(st))
    return t


def make_invoice_pdf(d, path):
    st = _pdf_styles()
    inv, s, case = d["invoice"], d["settings"], d.get("case")
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=12 * mm,
                            bottomMargin=12 * mm, title=f"Invoice {inv['inv_no']}", author=APP_NAME)
    E = lambda x: html.escape(str(x or ""))
    P = lambda x: Paragraph(x, st["n"])
    firm = s.get("company_name") or FIRM_NAME
    story = []
    gtype = inv.get("gst_type") or "Regular"
    _header(doc.width, story, st, s, GST_DOC_TITLE.get(gtype, "TAX INVOICE")
            + ("  (CANCELLED)" if inv["status"] == "Cancelled" else ""))
    if gtype == "Composition":
        story += [Paragraph("<b>Composition taxable person, not eligible to collect tax on supplies</b>",
                            ParagraphStyle("cmp", parent=st["n"], alignment=1, textColor=PDF_PINK, fontSize=9.5)),
                  Spacer(1, 6)]
    # ---- invoice meta + bill to
    pos = state_of(inv.get("client_gstin")) or (s.get("firm_state") or state_of(s.get("firm_gstin")))
    meta = (f"<b>Invoice No.:</b> {E(inv['inv_no'])}<br/><b>Invoice Date:</b> {dmy(inv['inv_date'])}<br/>"
            f"<b>Place of Supply:</b> {E(pos or '-')}" + ("<br/><b>Reverse charge:</b> No" if gtype == "Regular" else ""))
    if case:
        meta += f"<br/><b>Case Ref.:</b> #{case['id']} {E(case['law'])} {E(case['period'])}"
    bill = (f"<b>Bill To:</b><br/><b>{E(inv.get('client_name'))}</b><br/>{E(inv.get('client_address'))}<br/>"
            f"GSTIN: {E(inv.get('client_gstin') or '-')} &nbsp; PAN: {E(inv.get('client_pan') or '-')}")
    if case:
        matter = (f"<b>Matter:</b><br/>{E(case['law'])} - {E(case['forum'])}<br/>{E(case['section'])} "
                  f"{E(case['nature'])}<br/>Ref/Notice No.: {E(case['ref_no'] or '-')}")
    else:
        matter = f"<b>Particulars:</b><br/>{E(inv.get('notes') or 'Professional services')}"
    t = _grid([[P(bill), P(meta)], [P(matter), P("")]] if case else [[P(bill), P(meta)], [P(matter), P("")]],
              [100 * mm, 80 * mm], header=False, zebra=False)
    t.setStyle(TableStyle([("SPAN", (0, 1), (1, 1))]))
    story += [t, Spacer(1, 8)]
    # ---- lines
    rows = [["#", "Description of Service", "SAC", "Amount (Rs.)"]]
    taxable_lines = [x for x in d["lines"] if x["taxable"]]
    other_lines = [x for x in d["lines"] if not x["taxable"]]
    for i, x in enumerate(taxable_lines, 1):
        dt = f" <font size=7 color='#777777'>({dmy(x['date'])})</font>" if x.get("date") else ""
        rows.append([str(i), P(E(x["description"]) + dt), x.get("sac") or "", inr(x["amount"])])
    taxable = round(sum(x["amount"] for x in taxable_lines), 2)
    if inv["gst_mode"] == "NONE":
        rows.append(["", "Value of services" + (" (GST not applicable)" if gtype != "Regular" else ""), "",
                     inr(taxable)])
    else:
        rows.append(["", "Taxable Value", "", inr(taxable)])
    if inv["gst_mode"] == "NONE":
        pass
    elif inv["gst_mode"] == "IGST":
        rows.append(["", f"IGST @ {inv['gst_rate']:g}%", "", inr(inv["igst"])])
    else:
        rows.append(["", f"CGST @ {inv['gst_rate'] / 2:g}%", "", inr(inv["cgst"])])
        rows.append(["", f"SGST @ {inv['gst_rate'] / 2:g}%", "", inr(inv["sgst"])])
    if other_lines:
        rows.append(["", P("<b>Reimbursement of expenses incurred as pure agent (not part of taxable value)</b>"),
                     "", ""])
        for x in other_lines:
            rows.append(["", P((dmy(x["date"]) + " - " if x.get("date") else "") + E(x["description"])), "",
                         inr(x["amount"])])
    rows.append(["", "INVOICE TOTAL", "", inr(inv["total"])])
    tot_row = len(rows) - 1
    paid = inv["received"] or inv["tds"]
    if paid:
        bal = round(inv["total"] - inv["received"] - inv["tds"], 2)
        rows.append(["", "Less: Amount received", "", inr(inv["received"])])
        if inv["tds"]:
            rows.append(["", "Less: TDS deducted by client", "", inr(inv["tds"])])
        rows.append(["", "BALANCE DUE" if bal > 0.01 else "BALANCE DUE (Fully paid)", "", inr(max(0, bal))])
    t = _grid(rows, [10 * mm, 120 * mm, 18 * mm, 32 * mm])
    ts = [("ALIGN", (3, 1), (3, -1), "RIGHT"),
          ("FONT", (0, tot_row), (-1, tot_row), "Helvetica-Bold", 10.5),
          ("BACKGROUND", (0, tot_row), (-1, tot_row), PDF_PURPLE),
          ("TEXTCOLOR", (0, tot_row), (-1, tot_row), rl_colors.white)]
    if paid:
        ts += [("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 10.5),
               ("BACKGROUND", (0, -1), (-1, -1), rl_colors.HexColor(PDF_HI)),
               ("TEXTCOLOR", (0, -1), (-1, -1), PDF_PURPLE)]
    t.setStyle(TableStyle(ts))
    story += [t, Spacer(1, 6)]
    if (inv.get("service_details") or "").strip():
        dt = Table([[Paragraph("<b>Details of services rendered</b>", ParagraphStyle(
                        "sdh", parent=st["n"], textColor=PDF_PURPLE, fontSize=9.5))],
                    [Paragraph(E(inv["service_details"]).replace("\n", "<br/>"), st["n"])]], colWidths=[doc.width])
        dt.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, rl_colors.HexColor("#cfd9df")),
                                ("BACKGROUND", (0, 0), (-1, 0), PDF_BAND),
                                ("LINEBELOW", (0, 0), (-1, 0), 0.8, PDF_PINK),
                                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
        story += [dt, Spacer(1, 6)]
    story += [P("<b>Amount in words:</b> " + amount_in_words(inv["total"])), Spacer(1, 6)]
    bank = firm_bank_text(s)
    if bank:
        story += [P("<b>Bank details for payment:</b> " + E(bank)), Spacer(1, 4)]
    if s.get("terms"):
        story += [P("<b>Terms:</b> " + E(s["terms"])), Spacer(1, 4)]
    if gtype == "Unregistered":
        story.append(Paragraph("The supplier is not registered under GST; hence no GST is charged on this invoice.",
                               st["small"]))
    elif gtype == "Composition":
        story.append(Paragraph("Bill of supply issued by a composition taxable person; no tax is collected on this "
                               "supply.", st["small"]))
    if other_lines:
        story.append(Paragraph("Expenses shown as reimbursement were incurred on behalf of the client as a pure agent "
                               "and are excluded from the value of supply (Rule 33, CGST Rules, 2017).", st["small"]))
    story.append(Spacer(1, 14))
    sig = _sign_block(s, firm, doc.width, PDF_PURPLE)
    story += [sig, Spacer(1, 8), Paragraph("This is a computer-generated invoice.", st["small"])]
    doc.build(story)
    return path


def make_receipt_pdf(d, path):
    st = _pdf_styles()
    rc, inv, s = d["receipt"], d["invoice"], d["settings"]
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=12 * mm,
                            bottomMargin=12 * mm, title=f"Receipt {rc['rcpt_no']}", author=APP_NAME)
    E = lambda x: html.escape(str(x or ""))
    P = lambda x: Paragraph(x, st["n"])
    firm = s.get("company_name") or FIRM_NAME
    story = []
    _header(doc.width, story, st, s, "PAYMENT RECEIPT")
    meta = [[P(f"<b>Receipt No.:</b> {E(rc['rcpt_no'] or rc['id'])}"), P(f"<b>Receipt Date:</b> {dmy(rc['rdate'])}")],
            [P(f"<b>Against Invoice No.:</b> {E(inv['inv_no'])}"), P(f"<b>Invoice Date:</b> {dmy(inv['inv_date'])}")]]
    story += [_grid(meta, [doc.width / 2] * 2, header=False), Spacer(1, 12)]
    body = ParagraphStyle("rb", parent=st["n"], fontSize=10.5, leading=17)
    mode = E(rc["mode"]) + (f" (Ref: {E(rc['ref'])})" if rc.get("ref") else "")
    story += [Paragraph(f"Received with thanks from <b>{E(inv.get('client_name'))}</b>"
                        + (f" (GSTIN: {E(inv['client_gstin'])})" if inv.get("client_gstin") else "")
                        + f" a sum of <b>Rs. {inr(rc['amount'])}</b> (<i>{amount_in_words(rc['amount'])}</i>) by "
                        f"{mode} towards our Invoice No. <b>{E(inv['inv_no'])}</b> dated {dmy(inv['inv_date'])}.", body),
              Spacer(1, 12)]
    rows = [["Particulars", "Amount (Rs.)"], ["Invoice value", inr(inv["total"])],
            ["Amount received (this receipt)", inr(rc["amount"])],
            ["TDS deducted by client (this receipt)", inr(rc["tds"])],
            ["Total received till date (incl. TDS)", inr(rc["received_upto"])],
            ["Balance outstanding", inr(max(0, rc["balance_after"]))]]
    t = _grid(rows, [doc.width - 45 * mm, 45 * mm])
    t.setStyle(TableStyle([("ALIGN", (1, 1), (1, -1), "RIGHT"),
                           ("BACKGROUND", (0, 2), (-1, 2), rl_colors.HexColor(PDF_HI)),
                           ("FONT", (0, 2), (-1, 2), "Helvetica-Bold", 10),
                           ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9.5),
                           ("TEXTCOLOR", (0, -1), (-1, -1), PDF_PURPLE)]))
    story += [t, Spacer(1, 8)]
    notes = []
    if rc["tds"]:
        notes.append(f"TDS of Rs. {inr(rc['tds'])} is accepted subject to receipt of the TDS certificate and its "
                     "reflection in our tax credit statement.")
    if "Cheque" in (rc["mode"] or ""):
        notes.append("Subject to realisation of the cheque.")
    if rc["balance_after"] <= 0.01:
        notes.append("The above invoice stands fully settled. Thank you for the payment.")
    for n in notes:
        story.append(P("• " + E(n)))
    story.append(Spacer(1, 22))
    sig = _sign_block(s, firm, doc.width, PDF_PURPLE)
    story += [sig, Spacer(1, 10), Paragraph("This is a computer-generated receipt.", st["small"])]
    doc.build(story)
    return path


def make_hearing_note_pdf(d, path, proc=None, docs=None):
    """Hearing note: proc given -> note of that hearing / submission; proc None -> brief for the next hearing."""
    st = _pdf_styles()
    c, s = d["case"], d["settings"]
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=12 * mm,
                            bottomMargin=12 * mm, title=f"Hearing Note Case {c['id']}", author=APP_NAME)
    E = lambda x: html.escape(str(x if x not in (None, "") else "-"))
    P = lambda x: Paragraph(x, st["n"])
    B = lambda x: Paragraph(f"<b>{x}</b>", st["n"])
    W = doc.width
    story = []
    if proc:
        title = f"HEARING NOTE - {proc['ptype'].upper()} ON {dmy(proc['pdate'])}"
    else:
        title = "HEARING BRIEF - " + (f"HEARING ON {dmy(c['next_hearing'])}" if c.get("next_hearing") else "NEXT HEARING")
    _header(W, story, st, s, title)
    info = [[B("Client"), P(E(c["client_name"])), B("Law / Forum"), P(E(c["law"]) + " - " + E(c["forum"]))],
            [B("AY / Period"), P(E(c["period"])), B("Section"), P(E(c["section"]))],
            [B("Ref / Notice No."), P(E(c["ref_no"])), B("Demand (Rs.)"), P(inr(c["demand"]))],
            [B("Issue"), P(E(c["nature"])), B("Handled by"), P(E(c.get("assigned_name")))],
            [B("Status"), P(E(c["status"])), B("PAN / GSTIN"), P(E(c["client_pan"]) + " / " + E(c["client_gstin"]))]]
    story += [_grid(info, [30 * mm, W / 2 - 30 * mm, 30 * mm, W / 2 - 30 * mm], header=False), Spacer(1, 10)]
    sec = ParagraphStyle("sec", parent=st["n"], fontName="Helvetica-Bold", fontSize=10.5, leading=14,
                         textColor=PDF_PURPLE, spaceBefore=4, spaceAfter=4)
    body = ParagraphStyle("hb", parent=st["n"], fontSize=10, leading=14.5)
    procs = d.get("proceedings", [])
    if proc:
        det = [[B("Date"), P(dmy(proc["pdate"])), B("Type"), P(E(proc["ptype"]))],
               [B("Attended / filed by"), P(E(proc.get("by_name") or proc["by_user"])), B("Outcome"), P(E(proc["outcome"]))],
               [B("Next date"), P(dmy(proc["next_date"]) if proc["next_date"] else "-"),
                B("Attachment"), P(E(proc.get("file_name")))]]
        story += [_grid(det, [30 * mm, W / 2 - 30 * mm, 30 * mm, W / 2 - 30 * mm], header=False), Spacer(1, 8),
                  Paragraph("Proceedings / submissions made", sec),
                  Paragraph(E(proc["details"]).replace("\n", "<br/>"), body), Spacer(1, 8)]
        earlier = [p for p in procs if (p["pdate"], p["id"]) < (proc["pdate"], proc["id"])][-5:]
    else:
        story += [Paragraph("Purpose of hearing", sec),
                  Paragraph(E(c.get("next_purpose")) + (f" &nbsp; (on {dmy(c['next_hearing'])})" if c.get("next_hearing")
                                                         else ""), body), Spacer(1, 6)]
        earlier = procs[-8:]
    if earlier:
        rows = [["Date", "Type", "Summary", "Outcome"]]
        for p in earlier:
            txt = p["details"] if len(p["details"]) < 260 else p["details"][:257] + "..."
            rows.append([dmy(p["pdate"]), P(E(p["ptype"])), P(E(txt)), P(E(p["outcome"]))])
        story += [Paragraph("Earlier proceedings" if proc else "Chronology of proceedings", sec),
                  _grid(rows, [22 * mm, 30 * mm, W - 92 * mm, 40 * mm]), Spacer(1, 8)]
    if docs:
        story += [Paragraph("Documents on file", sec),
                  P("<br/>".join("• " + E(x["name"]) + (f" - {E(x['title'])}" if x.get("title") else "") for x in docs[:15])),
                  Spacer(1, 6)]
    if not proc:
        story.append(Paragraph("Points to argue / notes", sec))
        lines = Table([[""] for _ in range(9)], colWidths=[W], rowHeights=[8 * mm] * 9)
        lines.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#cfd9df"))]))
        story += [lines, Spacer(1, 6)]
    story.append(Spacer(1, 26))
    sig = Table([["Attended / prepared by", "", "Reviewed by (Partner)"]], colWidths=[60 * mm, W - 120 * mm, 60 * mm])
    sig.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 9),
                             ("LINEABOVE", (0, 0), (0, 0), 0.6, PDF_PURPLE), ("LINEABOVE", (2, 0), (2, 0), 0.6, PDF_PURPLE),
                             ("ALIGN", (2, 0), (2, 0), "CENTER")]))
    story += [sig, Spacer(1, 10),
              Paragraph(f"Prepared on {datetime.now().strftime('%d-%m-%Y %H:%M')} via {APP_NAME} Litigation Tracker. "
                        "Confidential - for internal use.", st["small"])]
    doc.build(story)
    return path


def make_case_report_pdf(d, path):
    st = _pdf_styles()
    c, s = d["case"], d["settings"]
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=13 * mm,
                            bottomMargin=13 * mm, title=f"Case Report #{c['id']}", author=APP_NAME)
    E = lambda x: html.escape(str(x if x not in (None, "") else "-"))
    P = lambda x: Paragraph(x, st["n"])
    B = lambda x: Paragraph(f"<b>{x}</b>", st["n"])
    story = []
    _header(doc.width, story, st, s, f"LITIGATION STATUS REPORT - CASE #{c['id']} ({c['law'].upper()})")
    info_rows = [[B("Client"), P(E(c["client_name"])), B("PAN / GSTIN"), P(E(c["client_pan"]) + " / " + E(c["client_gstin"]))],
                 [B("AY / FY / Period"), P(E(c["period"])), B("Forum"), P(E(c["forum"]))],
                 [B("Section"), P(E(c["section"])), B("Ref / Notice No."), P(E(c["ref_no"]))],
                 [B("Issue"), P(E(c["nature"])), B("Demand (Rs.)"), P(inr(c["demand"]))],
                 [B("Status"), P(E(c["status"])), B("Handled by"), P(E(c.get("assigned_name")))],
                 [B("Next Hearing"), P(dmy(c["next_hearing"]) if c["next_hearing"] else "-"),
                  B("Purpose"), P(E(c["next_purpose"]))]]
    story += [_grid(info_rows, [32 * mm, 58 * mm, 32 * mm, 60 * mm], header=False), Spacer(1, 10)]
    rows = [["#", "Date", "Type", "Hearing / Submission Report", "Outcome", "Next Date"]]
    for i, p in enumerate(d["proceedings"], 1):
        rows.append([str(i), dmy(p["pdate"]), P(E(p["ptype"])),
                     P(E(p["details"]).replace("\n", "<br/>") +
                       f"<br/><font size=7 color='#777777'>by {E(p.get('by_name') or p['by_user'])}</font>"),
                     P(E(p["outcome"])), dmy(p["next_date"]) if p["next_date"] else "-"])
    if len(rows) == 1:
        rows.append(["-", "-", "-", "No proceedings recorded yet", "-", "-"])
    story += [Paragraph("<b>Chronology of hearings &amp; submissions</b>", st["b"]), Spacer(1, 4),
              _grid(rows, [8 * mm, 20 * mm, 26 * mm, 76 * mm, 29 * mm, 23 * mm]), Spacer(1, 8)]
    if c["remarks"] and c["remarks"] != "-":
        story.append(P("<b>Remarks:</b> " + E(c["remarks"])))
    story += [Spacer(1, 10), Paragraph(f"Generated on {datetime.now().strftime('%d-%m-%Y %H:%M')} via "
                                       f"{APP_NAME} Litigation Tracker.", st["small"])]
    doc.build(story)
    return path


def make_lit_mis_pdf(rows, settings, path, title):
    from reportlab.lib.pagesizes import landscape
    st = _pdf_styles()
    doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=10 * mm, rightMargin=10 * mm,
                            topMargin=10 * mm, bottomMargin=10 * mm, title="Litigation MIS", author=APP_NAME)
    P = lambda x: Paragraph(html.escape(str(x or "-")), st["n"])
    story = []
    _header(doc.width, story, st, settings, title)
    data = [["#", "Law", "Client", "AY / Period", "Forum", "Section / Issue", "Demand (Rs.)", "Handled by",
             "Next Hearing", "Purpose", "Status"]]
    total = 0.0
    for r in rows:
        data.append([str(r["id"]), r["law"], P(r["client_name"]), P(r["period"]), P(r["forum"]),
                     P(f"{r['section']} {r['nature']}".strip()), inr(r["demand"]), P(r.get("assigned_name")),
                     dmy(r["next_hearing"]) if r["next_hearing"] else "-", P(r["next_purpose"]), P(r["status"])])
        total += r["demand"] or 0
    data.append(["", "", "TOTAL", "", "", "", inr(total), "", "", "", ""])
    t = _grid(data, [9 * mm, 21 * mm, 38 * mm, 20 * mm, 32 * mm, 37 * mm, 25 * mm, 25 * mm, 23 * mm, 23 * mm, 24 * mm])
    t.setStyle(TableStyle([("ALIGN", (6, 1), (6, -1), "RIGHT"), ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9)]))
    story += [t, Spacer(1, 6), Paragraph(f"Generated on {datetime.now().strftime('%d-%m-%Y %H:%M')} via "
                                         f"{APP_NAME} Litigation Tracker.", st["small"])]
    doc.build(story)
    return path


def make_payout_pdf(d, path):
    from reportlab.lib.pagesizes import landscape
    st = _pdf_styles()
    doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=10 * mm, rightMargin=10 * mm,
                            topMargin=10 * mm, bottomMargin=10 * mm, title=f"Payout {d['month']}", author=APP_NAME)
    P = lambda x: Paragraph(html.escape(str(x or "-")), st["n"])
    story = []
    _header(doc.width, story, st, d["settings"], f"MONTHLY PAYOUT STATEMENT (SALARY + REIMBURSEMENT) - "
                                      f"{month_label(d['month']).upper()}")
    head = ["#", "Team Member", "Designation", "Bank A/c", "Gross\nSalary", "LWP\nDays", "LWP\nDed.", "PT", "TDS",
            "Net\nSalary", "Reimburse-\nment", "TOTAL\nPAYOUT"]
    data = [head]
    keys = ("gross", "lwp_amt", "pt", "tds", "net_salary", "reimb", "total")
    tot = dict.fromkeys(keys, 0.0)
    for i, r in enumerate(d["rows"], 1):
        data.append([str(i), P(r["full_name"]), P(r["designation"]), P(r["bank"]), inr(r["gross"]),
                     f"{r['lwp_days']:g}", inr(r["lwp_amt"]), inr(r["pt"]), inr(r["tds"]), inr(r["net_salary"]),
                     inr(r["reimb"]), inr(r["total"])])
        for k in keys:
            tot[k] += r[k]
    data.append(["", "TOTAL", "", "", inr(tot["gross"]), "", inr(tot["lwp_amt"]), inr(tot["pt"]), inr(tot["tds"]),
                 inr(tot["net_salary"]), inr(tot["reimb"]), inr(tot["total"])])
    t = _grid(data, [8 * mm, 40 * mm, 30 * mm, 30 * mm, 24 * mm, 14 * mm, 20 * mm, 16 * mm, 18 * mm, 24 * mm,
                     24 * mm, 29 * mm])
    t.setStyle(TableStyle([("ALIGN", (4, 1), (-1, -1), "RIGHT"),
                           ("FONT", (0, -1), (-1, -1), "Helvetica-Bold", 9),
                           ("BACKGROUND", (-1, 1), (-1, -1), rl_colors.HexColor(PDF_HI)),
                           ("BACKGROUND", (0, -1), (-1, -1), PDF_PURPLE),
                           ("TEXTCOLOR", (0, -1), (-1, -1), rl_colors.white)]))
    story += [t, Spacer(1, 6),
              Paragraph(f"<b>Total payout in words:</b> {amount_in_words(tot['total'])}", st["n"]), Spacer(1, 4),
              Paragraph("Net Salary = Gross - Loss of Pay - Professional Tax - TDS. Reimbursement = claims approved "
                        "for the month. Pending claims are not included.", st["small"]), Spacer(1, 30)]
    sig = Table([["Prepared by", "", "Approved by (Manager)"]], colWidths=[70 * mm, 137 * mm, 70 * mm])
    sig.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 9),
                             ("LINEABOVE", (0, 0), (0, 0), 0.6, rl_colors.black),
                             ("LINEABOVE", (2, 0), (2, 0), 0.6, rl_colors.black), ("ALIGN", (2, 0), (2, 0), "CENTER")]))
    story += [sig, Spacer(1, 8), Paragraph(f"Generated on {d['generated']} via {APP_NAME}.", st["small"])]
    doc.build(story)
    return path


def make_ics(events, path):
    def esc(x):
        return str(x or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")
    stamp = time.strftime("%Y%m%dT%H%M%SZ", time.gmtime())
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Suthar & Suthar//Litigation Tracker//EN",
             "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
    for e in events:
        d = date.fromisoformat(e["date"])
        lines += ["BEGIN:VEVENT", f"UID:ss-case{e['case_id']}-{d:%Y%m%d}@sutharsuthar.local", f"DTSTAMP:{stamp}",
                  f"DTSTART;VALUE=DATE:{d:%Y%m%d}", f"DTEND;VALUE=DATE:{d + timedelta(days=1):%Y%m%d}",
                  "SUMMARY:" + esc(f"Hearing: {e['client']} ({e['law']}) - {e['forum']}"),
                  "DESCRIPTION:" + esc(f"Case #{e['case_id']} | {e.get('period', '')} | Ref {e.get('ref_no', '')}"
                                       f" | Purpose: {e.get('text', '')}"),
                  "BEGIN:VALARM", "TRIGGER:-P1D", "ACTION:DISPLAY", "DESCRIPTION:Hearing tomorrow", "END:VALARM",
                  "BEGIN:VALARM", "TRIGGER:-PT2H", "ACTION:DISPLAY", "DESCRIPTION:Hearing today", "END:VALARM",
                  "END:VEVENT"]
    lines.append("END:VCALENDAR")
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")
    return path


def send_email(cfg, password, to_addr, subject, body, attachment):
    msg = EmailMessage()
    msg["From"] = cfg.get("smtp_user")
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)
    if attachment:
        with open(attachment, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="pdf",
                               filename=os.path.basename(attachment))
    host, port = cfg.get("smtp_host"), int(cfg.get("smtp_port") or 587)
    sec = cfg.get("smtp_security", "STARTTLS")
    ctx = ssl.create_default_context()
    if sec == "SSL":
        server = smtplib.SMTP_SSL(host, port, context=ctx, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
        if sec == "STARTTLS":
            server.starttls(context=ctx)
    try:
        if password:
            server.login(cfg.get("smtp_user"), password)
        server.send_message(msg)
    finally:
        try:
            server.quit()
        except Exception:
            pass


# =============================================================================
# 9. GUI : THEME & WIDGET HELPERS
# =============================================================================
QSS = """
* { font-family: 'Segoe UI', 'Arial'; font-size: 10pt; }
QWidget { color: #2e2a47; }
QMainWindow, QDialog, #Page { background: #faf8ff; }
QToolTip { background: white; color: #2e2a47; border: 1px solid #ddd6fe; padding: 4px; }
#Sidebar { background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f3e8ff, stop:0.5 #fdf2f8, stop:1 #e0f7fa);
    border-right: 1px solid #ede9fe; }
#Sidebar QPushButton { color: #4c4570; background: transparent; border: none; text-align: left;
    padding: 10px 14px; font-size: 10.5pt; border-radius: 12px; }
#Sidebar QPushButton:hover { background: rgba(167,139,250,0.18); color: #5b21b6; }
#Sidebar QPushButton:checked { background: white; color: #6d28d9; font-weight: bold;
    border-left: 4px solid #a78bfa; }
#Brand { color: #6d28d9; font-size: 16pt; font-weight: 800; }
#BrandSub { color: #a78bfa; font-size: 8.5pt; font-weight: 600; }
#UserBadge { color: #4c1d95; background: rgba(255,255,255,0.85); border: 1px solid #ede9fe;
    border-radius: 12px; padding: 8px; }
QLabel#PageTitle { font-size: 19pt; font-weight: 800; color: #6d28d9; }
QLabel#PageSub { color: #8b84a8; }
QPushButton { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #8b5cf6, stop:1 #ec4899);
    color: white; border: none; border-radius: 10px; padding: 8px 16px; font-weight: 600; }
QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7c3aed, stop:1 #db2777); }
QPushButton:pressed { padding-top: 9px; }
QPushButton:disabled { background: #d6d3e6; color: #f5f5f5; }
QPushButton[kind="green"] { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #10b981, stop:1 #34d399); }
QPushButton[kind="red"] { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #f43f5e, stop:1 #fb7185); }
QPushButton[kind="orange"] { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #f97316, stop:1 #fb923c); }
QPushButton[kind="blue"] { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3b82f6, stop:1 #38bdf8); }
QPushButton[kind="gray"] { background: #a5a3c2; }
QLineEdit, QComboBox, QDateEdit, QDoubleSpinBox, QSpinBox, QTextEdit {
    background: white; border: 1.5px solid #e9e5fb; border-radius: 9px; padding: 5px 8px; }
QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QDoubleSpinBox:focus, QSpinBox:focus, QTextEdit:focus {
    border: 1.5px solid #a78bfa; background: #fefcff; }
QGroupBox { background: white; border: 1px solid #efeafc; border-radius: 16px; margin-top: 20px;
    padding: 14px 12px 12px 12px; font-weight: 700; color: #6d28d9; font-size: 10.5pt; }
QGroupBox::title { subcontrol-origin: margin; left: 14px; padding: 3px 10px; color: #5b21b6;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #ede9fe, stop:1 #fce7f3); border-radius: 8px; }
QTableWidget { background: white; border: 1px solid #efeafc; border-radius: 12px; gridline-color: #f3f0fb;
    selection-background-color: #ede9fe; selection-color: #2e2a47; alternate-background-color: #fcfbff; }
QHeaderView::section { background: #f3efff; color: #5b21b6; padding: 7px; border: none;
    border-bottom: 2px solid #ddd6fe; font-weight: 700; }
QListWidget { background: white; border: 1px solid #efeafc; border-radius: 14px; padding: 4px; }
QListWidget::item { padding: 9px 6px; border-radius: 9px; }
QListWidget::item:selected { background: #ede9fe; color: #4c1d95; }
QTextBrowser { background: #fdfcff; border: 1px solid #efeafc; border-radius: 14px; padding: 6px; }
QStatusBar { background: #f5f3ff; color: #5b21b6; border-top: 1px solid #ede9fe; }
QStatusBar QLabel { color: #5b21b6; }
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #ddd6fe; border-radius: 4px; min-height: 30px; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: #ddd6fe; border-radius: 4px; min-width: 30px; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QCheckBox { spacing: 6px; }
QTabWidget::pane { border: 1px solid #efeafc; border-radius: 12px; background: white; top: -1px; }
QTabBar::tab { background: #f5f3ff; color: #6d28d9; padding: 8px 18px; margin-right: 4px;
    border-top-left-radius: 10px; border-top-right-radius: 10px; font-weight: 600; }
QTabBar::tab:selected { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #c4b5fd, stop:1 #f9a8d4);
    color: #3b0764; }
QCalendarWidget QWidget#qt_calendar_navigationbar { background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
    stop:0 #c4b5fd, stop:1 #f9a8d4); }
QCalendarWidget QToolButton { color: #3b0764; background: transparent; font-weight: 700; font-size: 11pt; }
QCalendarWidget QAbstractItemView { selection-background-color: #8b5cf6; selection-color: white; }
QMenu { background: white; border: 1px solid #ede9fe; }
QMenu::item:selected { background: #ede9fe; color: #4c1d95; }
"""

CARD_COLORS = [("#ff6a00", "#ee0979"), ("#11998e", "#38ef7d"), ("#4a00e0", "#8e2de2"),
               ("#1e88e5", "#26c6da"), ("#f7971e", "#ffd200"), ("#e91e63", "#ff8a80"),
               ("#00b09b", "#96c93d"), ("#654ea3", "#eaafc8")]


def btn(text, kind=None, slot=None):
    b = QPushButton(text)
    if kind:
        b.setProperty("kind", kind)
    b.setCursor(Qt.PointingHandCursor)
    if slot:
        b.clicked.connect(slot)
    return b


def title_block(title, sub=""):
    w = QWidget()
    lay = QVBoxLayout(w)
    lay.setContentsMargins(0, 0, 0, 6)
    t = QLabel(title)
    t.setObjectName("PageTitle")
    lay.addWidget(t)
    if sub:
        s = QLabel(sub)
        s.setObjectName("PageSub")
        s.setWordWrap(True)
        lay.addWidget(s)
    return w


def make_table(headers, stretch_col=None):
    t = QTableWidget(0, len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setSelectionMode(QAbstractItemView.SingleSelection)
    t.setAlternatingRowColors(True)
    t.verticalHeader().setVisible(False)
    h = t.horizontalHeader()
    h.setSectionResizeMode(QHeaderView.ResizeToContents)
    if stretch_col is not None:
        h.setSectionResizeMode(stretch_col, QHeaderView.Stretch)
    else:
        h.setStretchLastSection(True)
    return t


STATUS_COLORS = {"Approved": "#11998e", "Rejected": "#e53935", "Pending": "#ff8f00", "Cancelled": "#8a8fb0"}


def fill_table(t, rows, keys=None):
    """rows: list of lists; first hidden data = row[0] stored as UserRole of col 0."""
    t.setRowCount(0)
    for r in rows:
        i = t.rowCount()
        t.insertRow(i)
        for j, v in enumerate(r):
            txt = "" if v is None else (f"{v:g}" if isinstance(v, float) else str(v))
            it = QTableWidgetItem(txt)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            if str(v) in STATUS_COLORS:
                it.setForeground(QColor(STATUS_COLORS[str(v)]))
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            t.setItem(i, j, it)
        if keys is not None:
            t.item(i, 0).setData(Qt.UserRole, keys[i])


def selected_key(t):
    r = t.currentRow()
    if r < 0 or not t.item(r, 0):
        return None
    return t.item(r, 0).data(Qt.UserRole)


def dot_icon(color):
    pm = QPixmap(14, 14)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QColor(color))
    p.setPen(Qt.NoPen)
    p.drawEllipse(1, 1, 12, 12)
    p.end()
    return QIcon(pm)


def app_icon():
    pm = QPixmap(64, 64)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QColor("#7b2ff7"))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(0, 0, 64, 64, 14, 14)
    p.setPen(QColor("white"))
    f = QFont("Segoe UI", 20)
    f.setBold(True)
    p.setFont(f)
    p.drawText(pm.rect(), Qt.AlignCenter, "S&S")
    p.end()
    return QIcon(pm)


def mix_white(hexc, f):
    """Blend a colour with white (f = 0..1, share of white) -> light pastel."""
    c = QColor(hexc)
    return QColor(int(c.red() + (255 - c.red()) * f), int(c.green() + (255 - c.green()) * f),
                  int(c.blue() + (255 - c.blue()) * f)).name()


class Card(QFrame):
    """Light pastel card: soft gradient, coloured headline value."""

    def __init__(self, title, colors, icon=""):
        super().__init__()
        self.setObjectName("Card")
        c1, c2 = colors
        self.accent = QColor(c1).darker(118).name()
        self.setStyleSheet(f"QFrame#Card {{ border-radius: 18px; border: 1px solid {mix_white(c1, 0.70)};"
                           f" background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {mix_white(c1, 0.86)},"
                           f" stop:1 {mix_white(c2, 0.80)}); }} QLabel {{ background: transparent; }}")
        self.setMinimumHeight(110)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 12, 16, 12)
        t = QLabel(f"{icon}  {title}")
        t.setStyleSheet("font-size: 10pt; font-weight: 600; color: #4b4466;")
        self.val = QLabel("—")
        self.vsize = 21
        self.val.setStyleSheet(f"font-size: 21pt; font-weight: 800; color: {self.accent};")
        self.sub = QLabel("")
        self.sub.setStyleSheet("font-size: 8.5pt; color: #6b6487;")
        self.sub.setWordWrap(True)
        lay.addWidget(t)
        lay.addWidget(self.val)
        lay.addWidget(self.sub)
        lay.addStretch()

    def small(self):
        self.val.setStyleSheet(f"font-size: 16pt; font-weight: 800; color: {self.accent};")

    def set(self, value, sub=""):
        self.val.setText(str(value))
        self.sub.setText(sub)


def month_edit():
    e = QDateEdit(QDate.currentDate())
    e.setDisplayFormat("MMMM yyyy")
    e.setCalendarPopup(False)
    return e


def ym_of(e):
    d = e.date()
    return f"{d.year():04d}-{d.month():02d}"


def date_edit(d=None):
    e = QDateEdit(d or QDate.currentDate())
    e.setDisplayFormat("dd-MM-yyyy")
    e.setCalendarPopup(True)
    return e


def err(parent, e, title="Error"):
    QMessageBox.warning(parent, f"{APP_NAME} - {title}", str(e))


def info(parent, text, title="Information"):
    QMessageBox.information(parent, f"{APP_NAME} - {title}", text)


class Page(QWidget):
    def __init__(self, main):
        super().__init__()
        self.main = main
        self.client = main.client
        self.setObjectName("Page")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        inner.setObjectName("Page")
        self.lay = QVBoxLayout(inner)
        self.lay.setContentsMargins(24, 18, 24, 18)
        self.lay.setSpacing(10)
        scroll.setWidget(inner)
        outer.addWidget(scroll)

    @property
    def me(self):
        return self.main.user["username"]

    @property
    def admin(self):
        return self.main.user["role"] == "admin"

    def user_combo(self, include_all=True):
        cb = QComboBox()
        cb.setMinimumWidth(200)
        self.main.combos.append((cb, include_all))
        self.main.fill_user_combo(cb, include_all)
        if not self.admin:
            cb.setEnabled(False)
        return cb

    def refresh(self):
        pass


# =============================================================================
# 10. GUI : LOGIN
# =============================================================================
class LoginDialog(QDialog):
    def __init__(self, cfg, server):
        super().__init__()
        self.cfg, self.server = cfg, server
        self.client = self.data = None
        self.setWindowTitle(f"{APP_NAME} - Login")
        self.setWindowIcon(app_icon())
        self.setMinimumWidth(470)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 16)
        firm = cfg.get("firm_name") or FIRM_NAME
        banner = QLabel(f"<div style='font-size:21pt;font-weight:800'>{PRODUCT_NAME}</div>"
                        f"<div style='font-size:13pt;font-weight:700;color:#7c3aed'>of {html.escape(firm)}</div>"
                        "<div style='font-size:9.5pt'>Chat • Attendance • Leave • Reimbursement • Litigation • Tasks • "
                        "Invoicing</div><div style='font-size:9pt;color:#a855f7'>Welcome back — let's make today count ✨</div>")
        banner.setAlignment(Qt.AlignCenter)
        banner.setMinimumHeight(130)
        banner.setStyleSheet("color:#5b21b6; background: qlineargradient(x1:0, y1:0, x2:1, y2:1, "
                             "stop:0 #ede9fe, stop:0.5 #fce7f3, stop:1 #e0f2fe);")
        lay.addWidget(banner)
        form_w = QWidget()
        form = QFormLayout(form_w)
        form.setContentsMargins(26, 14, 26, 4)
        form.setVerticalSpacing(10)
        self.host_cb = QCheckBox("Host the server on this PC  (Manager / main computer)")
        self.host_cb.setChecked(bool(server) or cfg.get("host_mode", False))
        if server:
            self.host_cb.setEnabled(False)
            self.host_cb.setText("Server is running on this PC")
        self.host_cb.toggled.connect(self._toggle)
        form.addRow(self.host_cb)
        row = QHBoxLayout()
        self.ip = QLineEdit(cfg.get("server_ip", ""))
        self.ip.setPlaceholderText("Server PC IP, e.g. 192.168.1.10")
        self.find_btn = btn("Find Server", "blue", self.find)
        row.addWidget(self.ip)
        row.addWidget(self.find_btn)
        form.addRow("Server IP:", row)
        self.user = QLineEdit(cfg.get("username", ""))
        self.user.setPlaceholderText("username")
        self.pw = QLineEdit()
        self.pw.setEchoMode(QLineEdit.Password)
        self.pw.setPlaceholderText("password")
        self.pw.returnPressed.connect(self.login)
        form.addRow("Username:", self.user)
        form.addRow("Password:", self.pw)
        lay.addWidget(form_w)
        self.hint = QLabel()
        self.hint.setWordWrap(True)
        self.hint.setStyleSheet("color:#6b6f8e; padding: 0 26px; font-size: 8.5pt;")
        lay.addWidget(self.hint)
        fg = QLabel("Forgot password? Only the Manager / Admin can reset it (Admin › Reset Password).")
        fg.setStyleSheet("color:#a78bfa; padding: 0 26px; font-size: 8pt;")
        fg.setWordWrap(True)
        lay.addWidget(fg)
        b = btn("  Login  ", None, self.login)
        b.setMinimumHeight(40)
        bl = QHBoxLayout()
        bl.setContentsMargins(26, 6, 26, 0)
        bl.addWidget(b)
        lay.addLayout(bl)
        self._toggle(self.host_cb.isChecked())
        (self.pw if self.user.text() else self.user).setFocus()

    def _toggle(self, on):
        self.ip.setEnabled(not on)
        self.find_btn.setEnabled(not on)
        if on:
            self.hint.setText("This PC will hold the central database. Team members connect to: "
                              + ", ".join(local_ips()) + f"  (port {PORT}).  First-time login: "
                              "admin / admin@123. Allow the app if Windows Firewall asks.")
        else:
            self.hint.setText("Click 'Find Server' to locate the host PC on your network automatically, "
                              "or type its IP address.")

    def find(self):
        self.find_btn.setEnabled(False)
        self.find_btn.setText("Searching...")

        def done(res, e):
            self.find_btn.setEnabled(True)
            self.find_btn.setText("Find Server")
            if res:
                self.ip.setText(res[0][0])
                info(self, "Server found: " + ", ".join(f"{r[0]} ({r[2]})" for r in res))
            else:
                err(self, "No server found. Make sure the host PC is running the app with "
                          "'Host the server' ticked, or type its IP address.", "Not found")
        run_task(discover_servers, done)

    def login(self):
        host_mode = self.host_cb.isChecked()
        if not self.user.text().strip() or not self.pw.text():
            return err(self, "Please enter username and password")
        if host_mode and not self.server:
            try:
                srv = Server()
                srv.start()
                self.server = srv
            except OSError as e:
                return err(self, f"Could not start the server on port {PORT}.\n"
                                 f"Is another copy already hosting on this PC?\n\n{e}")
        host = "127.0.0.1" if host_mode else self.ip.text().strip()
        if not host:
            return err(self, "Please enter the server IP address")
        client = NetClient()
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            client.connect_to(host, PORT)
            data, _ = client.request("login", username=self.user.text().strip(), password=self.pw.text())
        except Exception as e:
            client.close()
            QApplication.restoreOverrideCursor()
            msg = str(e)
            if isinstance(e, OSError) and not isinstance(e, ConnectionError):
                msg = f"Cannot reach server at {host}:{PORT}.\n{e}"
            return err(self, msg, "Login failed")
        QApplication.restoreOverrideCursor()
        if data.get("must_change"):
            dlg = ForceChangePwDialog(self, data["user"]["full_name"])
            if not dlg.exec_():
                client.close()
                return
            try:
                client.call("change_password", old=self.pw.text(), new=dlg.p1.text())
            except Exception as e:
                client.close()
                return err(self, e)
            info(self, "New password saved. Please use it from next time.")
        self.cfg.update(server_ip=self.ip.text().strip(), username=self.user.text().strip(),
                        host_mode=host_mode)
        save_cfg(self.cfg)
        self.client, self.data = client, data
        self.accept()


# =============================================================================
# 11. GUI : DASHBOARD
# =============================================================================
class DashboardPage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.hello = title_block(f"Welcome, {main.user['full_name']}",
                                 datetime.now().strftime("%A, %d %B %Y"))
        self.lay.addWidget(self.hello)
        self.thought = QLabel()
        self.thought.setWordWrap(True)
        self.thought.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #fdf4ff, stop:1 #ecfeff);"
                                   " border: 1px solid #f5d0fe; border-radius: 14px; padding: 10px 14px;"
                                   " color: #6b21a8; font-size: 10.5pt;")
        self.lay.addWidget(self.thought)
        self._greet()
        grid = QGridLayout()
        grid.setSpacing(14)
        self.c_att = Card("Today's Attendance", CARD_COLORS[1], "🕘")
        self.c_bal = Card("Paid Leave Balance (this month)", CARD_COLORS[2], "🌴")
        self.c_lwp = Card("Loss of Pay Days (this month)", CARD_COLORS[0], "⚠")
        self.c_rb = Card("Reimbursement Claimed (this month)", CARD_COLORS[3], "🧾")
        self.c_on = Card("Team Online", CARD_COLORS[6], "👥")
        self.c_pr = Card("Days Present (this month)", CARD_COLORS[4], "✅")
        self.c_hr = Card("Upcoming Hearings", ("#c2185b", "#7b1fa2"), "⚖")
        self.c_tk = Card("My Open Tasks", ("#e53935", "#ff8f00"), "✅")
        cards = [self.c_att, self.c_bal, self.c_lwp, self.c_rb, self.c_on, self.c_pr, self.c_hr, self.c_tk]
        if self.admin:
            self.c_pl = Card("Pending Leave Approvals", CARD_COLORS[5], "📝")
            self.c_pre = Card("Pending Reimbursement Claims", CARD_COLORS[7], "💰")
            cards += [self.c_pl, self.c_pre]
        for i, c in enumerate(cards):
            grid.addWidget(c, i // 3, i % 3)
        self.lay.addLayout(grid)
        g = QGroupBox("Quick Actions")
        h = QHBoxLayout(g)
        h.addWidget(btn("✔ Check In", "green", lambda: self.main.pages["att"].mark("in")))
        h.addWidget(btn("⏏ Check Out", "red", lambda: self.main.pages["att"].mark("out")))
        h.addWidget(btn("🌴 Apply Leave", "orange", lambda: self.main.go("leave")))
        h.addWidget(btn("🧾 New Claim", "blue", lambda: self.main.go("reimb")))
        h.addWidget(btn("💬 Open Chat", None, lambda: self.main.go("chat")))
        h.addWidget(btn("⚖ Litigation", "gray", lambda: self.main.go("lit")))
        h.addWidget(btn("✅ Tasks", "green", lambda: self.main.go("tasks")))
        self.lay.addWidget(g)
        pol = QLabel(f"<b>Leave policy:</b> {PAID_LEAVE_PER_MONTH} paid leaves per calendar month "
                     "(non-cumulative). Approved leave beyond this is treated as Loss of Pay (LWP). "
                     "Sundays are weekly off and are not counted as leave.")
        pol.setWordWrap(True)
        pol.setStyleSheet("background:#fff4e5; border:1px solid #ffd59a; border-radius:10px; padding:10px;")
        self.lay.addWidget(pol)
        self.lay.addStretch()

    THOUGHTS = [
        "Small steps every day add up to big results. 🌱",
        "Clear books, calm mind — let's close today's work neatly. 📒",
        "Every deadline met is a promise kept to a client. 🤝",
        "Ask, share, help — great firms are built by great teams. 💬",
        "Quality first; speed follows with practice. ✨",
        "A quick update to the team saves an hour of follow-ups. ⏱",
        "Today's preparation is tomorrow's confident hearing. ⚖",
        "Celebrate the small wins — a filed return is a job well done! 🎉",
        "Stay curious: every notice is a chance to learn something new. 📚",
        "Take a short break, stretch, and come back stronger. 🧘",
        "Your attention to detail is what clients trust us for. 🔍",
        "Plan the day, then work the plan. 🗓",
    ]

    def _greet(self):
        h = datetime.now().hour
        g = ("Good morning", "☀️") if h < 12 else (("Good afternoon", "🌤") if h < 17 else ("Good evening", "🌙"))
        first = (self.main.user["full_name"] or "").split(" ")[0]
        self.hello.findChild(QLabel, "PageTitle").setText(f"{g[0]}, {first} {g[1]}")
        self.hello.findChild(QLabel, "PageSub").setText(datetime.now().strftime("%A, %d %B %Y") +
                                                        "  •  Have a wonderful, productive day!")
        self.thought.setText("💡 <b>Thought for today:</b> " +
                             self.THOUGHTS[date.today().toordinal() % len(self.THOUGHTS)])

    def refresh(self):
        self._greet()
        try:
            d = self.client.call("dashboard")
        except Exception as e:
            return self.main.status(str(e))
        t = d.get("today")
        if t:
            self.c_att.set("Checked In" if not t.get("check_out") else "Completed",
                           f"In: {t['check_in']}   Out: {t.get('check_out') or '--'}   ({t['mode']})")
        else:
            self.c_att.set("Not Marked", "Click 'Check In' to mark attendance")
        lv = d["leave"]
        self.c_bal.set(f"{lv['balance']:g} / {lv['quota']}", f"Availed {lv['taken']:g}, pending {lv['pending']:g}")
        self.c_lwp.set(f"{lv['lwp']:g}", "Leave beyond paid quota")
        self.c_rb.set("Rs. " + inr(d["reimb"]["s"]), f"{d['reimb']['n']} claim(s) (excl. rejected)")
        self.c_on.set(f"{d['online']} / {d['team']}", "members connected now")
        self.c_pr.set(d["present"], "attendance entries")
        self.c_hr.set(d["hearings"], f"in next {d['hearing_days']} day(s); {d['overdue']} past date(s) to update")
        self.c_tk.set(d["tasks_open"], f"{d['tasks_red']} due within 3 days / overdue")
        if self.admin:
            self.c_pl.set(d["pending_leaves"], "awaiting your decision")
            self.c_pre.set(d["pending_reimb"], f"Present today: {d['present_today']}")


# =============================================================================
# 12. GUI : CHAT, FILES, VOICE, SCREEN
# =============================================================================
class ChatPage(QWidget):
    def __init__(self, main):
        super().__init__()
        self.main, self.client = main, main.client
        self.setObjectName("Page")
        self.current = "ALL"
        self.messages = []
        self.unread = {}
        lay = QHBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 16)
        split = QSplitter()
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 8, 0)
        lt = QLabel("Team")
        lt.setObjectName("PageTitle")
        ll.addWidget(lt)
        self.list = QListWidget()
        self.list.setIconSize(QPixmap(14, 14).size())
        self.list.currentItemChanged.connect(self._select)
        ll.addWidget(self.list)
        split.addWidget(left)
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(8, 0, 0, 0)
        top = QHBoxLayout()
        self.peer_lbl = QLabel("Everyone (Group)")
        self.peer_lbl.setObjectName("PageTitle")
        top.addWidget(self.peer_lbl)
        top.addStretch()
        self.call_btn = btn("📞 Voice Call", "green", self.voice_call)
        self.screen_btn = btn("🖥 Share Screen", "blue", self.share_screen)
        self.file_btn = btn("📎 Send File", "orange", self.send_file)
        for b in (self.call_btn, self.screen_btn, self.file_btn):
            top.addWidget(b)
        rl.addLayout(top)
        self.view = QTextBrowser()
        self.view.setOpenLinks(False)
        self.view.anchorClicked.connect(self._link)
        rl.addWidget(self.view, 1)
        bottom = QHBoxLayout()
        self.input = QLineEdit()
        self.input.setPlaceholderText("Type a message and press Enter...")
        self.input.setMinimumHeight(38)
        self.input.returnPressed.connect(self.send_text)
        bottom.addWidget(self.input)
        bottom.addWidget(btn("Send ➤", None, self.send_text))
        rl.addLayout(bottom)
        self.progress = QLabel("")
        self.progress.setStyleSheet("color:#8e2de2;")
        rl.addWidget(self.progress)
        split.addWidget(right)
        split.setSizes([260, 800])
        lay.addWidget(split)

    def refresh(self):
        self.refresh_users()
        self.load_history()

    def refresh_users(self):
        users = self.main.users
        self.list.blockSignals(True)
        self.list.clear()
        entries = [("ALL", "Everyone (Group)", True, "#ff6a00")]
        for u in users:
            if u["username"] == self.main.user["username"]:
                continue
            entries.append((u["username"], u["full_name"] + (f"  ·  {u['designation']}" if u["designation"] else ""),
                            u["online"], "#22c55e" if u["online"] else "#b0b3c6"))
        sel = None
        for key, label, online, color in entries:
            n = self.unread.get(key, 0)
            it = QListWidgetItem(dot_icon(color), label + (f"   ({n} new)" if n else ""))
            it.setData(Qt.UserRole, key)
            if n:
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            self.list.addItem(it)
            if key == self.current:
                sel = it
        if sel:
            self.list.setCurrentItem(sel)
        self.list.blockSignals(False)

    def _select(self, item, prev=None):
        if not item:
            return
        self.current = item.data(Qt.UserRole)
        self.unread.pop(self.current, None)
        self.peer_lbl.setText("Everyone (Group)" if self.current == "ALL" else self.main.name(self.current))
        self.call_btn.setEnabled(self.current != "ALL")
        self.load_history()
        QTimer.singleShot(0, self.refresh_users)

    def load_history(self):
        try:
            self.messages = self.client.call("chat_history", peer=self.current)
        except Exception as e:
            self.messages = []
            self.main.status(str(e))
        self.render()

    def _bubble(self, m):
        mine = m["sender"] == self.main.user["username"]
        ts = (m.get("ts") or "")[5:16]
        if m.get("file_id"):
            size = (m.get("file_size") or 0) / 1024
            sz = f"{size / 1024:.1f} MB" if size > 1024 else f"{size:.0f} KB"
            body = (f"📎 <a href='dl:{m['file_id']}' style='color:{'#fff59d' if mine else '#7b2ff7'};'>"
                    f"<b>{html.escape(m.get('file_name') or 'file')}</b></a> ({sz})<br>"
                    f"<span style='font-size:8pt'>click to download</span>")
        else:
            body = html.escape(m["text"]).replace("\n", "<br>")
        bg, fg = ("#7b2ff7", "#ffffff") if mine else ("#ffffff", "#1f2340")
        who = "" if mine else f"<span style='color:#e91e63;font-weight:bold;font-size:8.5pt'>{html.escape(m.get('sender_name') or m['sender'])}</span><br>"
        align = "right" if mine else "left"
        return (f"<table width='100%' cellspacing='0' cellpadding='3'><tr><td align='{align}'>"
                f"<table cellspacing='0' cellpadding='9' bgcolor='{bg}' style='border:1px solid #e3e5f4;'>"
                f"<tr><td style='color:{fg}'>{who}{body}<br>"
                f"<span style='font-size:7.5pt;color:{'#e6dcff' if mine else '#9a9db5'}'>{ts}</span>"
                f"</td></tr></table></td></tr></table>")

    def render(self):
        if not self.messages:
            self.view.setHtml("<p align='center' style='color:#9a9db5'><br><br>No messages yet. Say hello! 👋</p>")
            return
        self.view.setHtml("".join(self._bubble(m) for m in self.messages))
        sb = self.view.verticalScrollBar()
        sb.setValue(sb.maximum())

    def on_message(self, m):
        me = self.main.user["username"]
        key = "ALL" if m["recipient"] == "ALL" else (m["recipient"] if m["sender"] == me else m["sender"])
        if key == self.current:
            self.messages.append(m)
            self.render()
        else:
            self.unread[key] = self.unread.get(key, 0) + 1
            self.refresh_users()
        if m["sender"] != me:
            if self.main.stack.currentWidget() is not self or key != self.current:
                self.main.status(f"💬 New message from {m.get('sender_name')}")
                QApplication.beep()
            QApplication.alert(self.main)

    def send_text(self):
        text = self.input.text().strip()
        if not text:
            return
        try:
            self.client.call("chat_send", to=self.current, text=text)
            self.input.clear()
        except Exception as e:
            err(self, e)

    def send_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select file to share")
        if not path:
            return
        size = os.path.getsize(path)
        if size > MAX_FILE:
            return err(self, f"File is larger than {MAX_FILE // (1024 * 1024)} MB")
        if size == 0:
            return err(self, "File is empty")
        to, name = self.current, os.path.basename(path)
        self.progress.setText(f"Uploading {name} ...")
        self.file_btn.setEnabled(False)

        def work():
            with open(path, "rb") as f:
                data = f.read()
            return self.client.request("file_upload", data, timeout=900, to=to, name=name)

        def done(res, e):
            self.file_btn.setEnabled(True)
            self.progress.setText("" if not e else "")
            if e:
                err(self, e, "Upload failed")
            else:
                self.main.status(f"File '{name}' shared")
        run_task(work, done)

    def _link(self, url):
        s = url.toString()
        if s.startswith("dl:"):
            try:
                fid = int(s[3:])
            except ValueError:
                return
            self.main.download_file(fid, self.progress)

    def voice_call(self):
        if self.current == "ALL":
            return
        self.main.media.start_call(self.current)

    def share_screen(self):
        self.main.media.start_share(self.current)


class AudioEngine:
    RATE = 16000
    BLOCK = 640  # 40 ms of 16-bit mono audio

    def __init__(self, send_fn):
        self.send_fn = send_fn
        self.buf = bytearray()
        self.lock = threading.Lock()
        self.muted = False
        self.inp = self.out = None

    def start(self):
        self.inp = sd.RawInputStream(samplerate=self.RATE, channels=1, dtype="int16",
                                     blocksize=self.BLOCK, callback=self._in_cb)
        self.out = sd.RawOutputStream(samplerate=self.RATE, channels=1, dtype="int16",
                                      blocksize=self.BLOCK, callback=self._out_cb)
        self.out.start()
        self.inp.start()

    def _in_cb(self, indata, frames, t, status):
        if not self.muted:
            self.send_fn(bytes(indata))

    def _out_cb(self, outdata, frames, t, status):
        n = len(outdata)
        with self.lock:
            chunk = bytes(self.buf[:n])
            del self.buf[:n]
        if len(chunk) < n:
            chunk += b"\x00" * (n - len(chunk))
        outdata[:] = chunk

    def feed(self, data):
        with self.lock:
            self.buf.extend(data)
            limit = self.RATE * 2 // 2   # keep max 0.5 s buffered (low latency)
            if len(self.buf) > limit:
                del self.buf[:len(self.buf) - self.RATE * 2 // 5]

    def stop(self):
        for s in (self.inp, self.out):
            try:
                if s:
                    s.stop()
                    s.close()
            except Exception:
                pass
        self.inp = self.out = None


class CallDialog(QDialog):
    def __init__(self, media, peer_name):
        super().__init__(media.main)
        self.media = media
        self.setWindowTitle(f"Voice call - {peer_name}")
        self.setWindowIcon(app_icon())
        self.setMinimumWidth(340)
        self.setStyleSheet("QDialog { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #11998e,"
                           " stop:1 #4a00e0); } QLabel { color: white; }")
        lay = QVBoxLayout(self)
        n = QLabel(f"📞  {peer_name}")
        n.setStyleSheet("font-size: 16pt; font-weight: 800;")
        n.setAlignment(Qt.AlignCenter)
        self.st = QLabel("Calling...")
        self.st.setAlignment(Qt.AlignCenter)
        self.st.setStyleSheet("font-size: 11pt;")
        lay.addWidget(n)
        lay.addWidget(self.st)
        tip = QLabel("Tip: use headphones to avoid echo")
        tip.setAlignment(Qt.AlignCenter)
        tip.setStyleSheet("font-size: 8pt;")
        lay.addWidget(tip)
        h = QHBoxLayout()
        self.mute = btn("🎤 Mute", "orange", self._mute)
        self.mute.setCheckable(True)
        h.addWidget(self.mute)
        h.addWidget(btn("Hang Up", "red", self.close))
        lay.addLayout(h)
        self.started = None
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)

    def connected(self):
        self.started = time.time()
        self.timer.start(1000)
        self._tick()

    def _tick(self):
        s = int(time.time() - self.started)
        self.st.setText(f"Connected  {s // 60:02d}:{s % 60:02d}")

    def _mute(self):
        if self.media.audio:
            self.media.audio.muted = self.mute.isChecked()
        self.mute.setText("🔇 Unmute" if self.mute.isChecked() else "🎤 Mute")

    def closeEvent(self, ev):
        self.timer.stop()
        self.media.end_call(notify=True, from_dialog=True)
        super().closeEvent(ev)


class ViewerWindow(QWidget):
    def __init__(self, media, sid, sharer, sharer_name):
        super().__init__()
        self.media, self.sid, self.sharer = media, sid, sharer
        self.setWindowTitle(f"Screen shared by {sharer_name} - {APP_NAME}")
        self.setWindowIcon(app_icon())
        self.resize(1100, 680)
        self.setStyleSheet("background:#14112b;")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(4, 4, 4, 4)
        self.lbl = QLabel("Waiting for screen...")
        self.lbl.setAlignment(Qt.AlignCenter)
        self.lbl.setStyleSheet("color:white; font-size: 12pt;")
        self.lbl.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        lay.addWidget(self.lbl)
        self.pix = None
        self.closing_remote = False

    def set_frame(self, data):
        pm = QPixmap()
        if pm.loadFromData(data, "JPG"):
            self.pix = pm
            self._update()

    def _update(self):
        if self.pix:
            self.lbl.setPixmap(self.pix.scaled(self.lbl.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))

    def resizeEvent(self, ev):
        self._update()
        super().resizeEvent(ev)

    def closeEvent(self, ev):
        self.media.viewer_closed(self.sid, notify=not self.closing_remote)
        super().closeEvent(ev)


class ShareBar(QWidget):
    def __init__(self, media, target_name):
        super().__init__(None, Qt.WindowStaysOnTopHint | Qt.Tool)
        self.media = media
        self.setWindowTitle("Screen sharing")
        self.setStyleSheet("QWidget { background: #e53935; } QLabel { color: white; font-weight: 600; }")
        h = QHBoxLayout(self)
        self.lbl = QLabel(f"🔴 Sharing screen with {target_name}")
        h.addWidget(self.lbl)
        h.addWidget(btn("Stop Sharing", "gray", self.close))
        self.target_name = target_name

    def set_viewers(self, n):
        self.lbl.setText(f"🔴 Sharing screen with {self.target_name}  —  {n} viewer(s)")

    def closeEvent(self, ev):
        self.media.stop_share()
        super().closeEvent(ev)


class MediaManager(QObject):
    def __init__(self, main):
        super().__init__()
        self.main, self.client = main, main.client
        self.call = None        # {'sid','peer','box','dlg'}
        self.audio = None
        self.share = None       # {'sid','target','viewers','timer','bar'}
        self.viewers = {}       # sid -> ViewerWindow
        self.client.audio_hook = self._audio_in

    def _audio_in(self, hdr, payload):          # network thread
        c, a = self.call, self.audio
        if c and a and hdr.get("sid") == c["sid"]:
            a.feed(payload)

    # ------------------------------------------------------------- voice
    def start_call(self, peer):
        if not AUDIO_OK:
            return err(self.main, "Voice calling needs the 'sounddevice' library and a microphone.\n"
                                  "Run:  pip install sounddevice")
        if self.call:
            return err(self.main, "You are already in a call")
        sid = uuid.uuid4().hex
        try:
            self.client.call("relay", to=peer, kind="invite", mode="call", sid=sid)
        except Exception as e:
            return err(self.main, e, "Call failed")
        dlg = CallDialog(self, self.main.name(peer))
        self.call = {"sid": sid, "peer": peer, "dlg": dlg, "box": None}
        dlg.show()

    def _start_audio(self):
        c = self.call
        try:
            self.audio = AudioEngine(lambda data, sid=c["sid"], peer=c["peer"]:
                                     self.client.send("relay", data, to=peer, kind="audio",
                                                      mode="call", sid=sid))
            self.audio.start()
            return True
        except Exception as e:
            self.audio = None
            err(self.main, f"Could not open microphone / speaker:\n{e}", "Audio error")
            self.end_call(notify=True)
            return False

    def end_call(self, notify=True, from_dialog=False):
        c = self.call
        if not c:
            return
        self.call = None
        if notify:
            self.client.send("relay", to=c["peer"], kind="end", mode="call", sid=c["sid"])
        if self.audio:
            self.audio.stop()
            self.audio = None
        if c.get("box"):
            c["box"].done(QMessageBox.No)
        if c.get("dlg") and not from_dialog:
            c["dlg"].close()   # closeEvent -> end_call() returns at once (call already cleared)

    # ------------------------------------------------------------- screen
    def start_share(self, target):
        if self.share:
            return err(self.main, "You are already sharing your screen")
        sid = uuid.uuid4().hex
        try:
            self.client.call("relay", to=target, kind="invite", mode="screen", sid=sid)
        except Exception as e:
            return err(self.main, e, "Screen share failed")
        name = "Everyone" if target == "ALL" else self.main.name(target)
        timer = QTimer(self)
        timer.timeout.connect(self._grab)
        bar = ShareBar(self, name)
        self.share = {"sid": sid, "target": target, "viewers": set(), "timer": timer, "bar": bar}
        timer.start(250)  # ~4 frames per second
        bar.show()
        g = QApplication.primaryScreen().availableGeometry()
        bar.move(g.center().x() - bar.width() // 2, g.top() + 4)
        self.main.status("Waiting for team member(s) to accept the screen share...")

    def _grab(self):
        s = self.share
        if not s or not s["viewers"] or self.client.frames_pending > len(s["viewers"]):
            return
        try:
            pix = QApplication.primaryScreen().grabWindow(0)
            if pix.isNull():
                return
            if pix.width() > 1600:
                pix = pix.scaledToWidth(1600, Qt.SmoothTransformation)
            ba = QByteArray()
            buf = QBuffer(ba)
            buf.open(QIODevice.WriteOnly)
            pix.save(buf, "JPG", 60)
            data = bytes(ba)
        except Exception:
            return
        for v in list(s["viewers"]):
            self.client.send("relay", data, to=v, kind="frame", mode="screen", sid=s["sid"])

    def stop_share(self):
        s = self.share
        if not s:
            return
        self.share = None
        s["timer"].stop()
        targets = s["viewers"] if s["viewers"] else ({s["target"]} if s["target"] != "ALL" else set())
        if s["target"] == "ALL" and not s["viewers"]:
            self.client.send("relay", to="ALL", kind="end", mode="screen", sid=s["sid"])
        for v in targets:
            self.client.send("relay", to=v, kind="end", mode="screen", sid=s["sid"])
        if s["bar"].isVisible():
            s["bar"].hide()
        self.main.status("Screen sharing stopped")

    def viewer_closed(self, sid, notify=True):
        w = self.viewers.pop(sid, None)
        if w and notify:
            self.client.send("relay", to=w.sharer, kind="end", mode="screen", sid=sid)

    # ------------------------------------------------------- relay events
    def handle(self, h, payload):
        kind, mode, sid, frm = h.get("kind"), h.get("mode"), h.get("sid"), h.get("from")
        name = h.get("from_name") or frm
        if mode == "call":
            c = self.call
            if kind == "invite":
                if c:
                    self.client.send("relay", to=frm, kind="busy", mode="call", sid=sid)
                    return
                box = QMessageBox(QMessageBox.Question, f"{APP_NAME} - Incoming call",
                                  f"📞  {name} is calling you.\n\nAccept the voice call?",
                                  QMessageBox.Yes | QMessageBox.No, self.main)
                self.call = {"sid": sid, "peer": frm, "box": box, "dlg": None}
                QApplication.alert(self.main)
                QApplication.beep()
                ans = box.exec_()
                if not self.call or self.call["sid"] != sid:
                    self.main.status(f"Missed call from {name}")
                    return
                self.call["box"] = None
                if ans != QMessageBox.Yes:
                    self.client.send("relay", to=frm, kind="decline", mode="call", sid=sid)
                    self.call = None
                    return
                if not AUDIO_OK:
                    self.client.send("relay", to=frm, kind="decline", mode="call", sid=sid)
                    self.call = None
                    return err(self.main, "Voice calling needs the 'sounddevice' library (pip install sounddevice)")
                self.client.send("relay", to=frm, kind="accept", mode="call", sid=sid)
                dlg = CallDialog(self, name)
                self.call["dlg"] = dlg
                if self._start_audio():
                    dlg.show()
                    dlg.connected()
            elif c and c["sid"] == sid:
                if kind == "accept":
                    if self._start_audio() and c.get("dlg"):
                        c["dlg"].connected()
                elif kind in ("decline", "busy"):
                    self.end_call(notify=False)
                    info(self.main, f"{name} is busy on another call" if kind == "busy"
                         else f"{name} declined the call")
                elif kind == "end":
                    self.end_call(notify=False)
                    self.main.status(f"Call with {name} ended")
        elif mode == "screen":
            if kind == "frame":
                w = self.viewers.get(sid)
                if w:
                    w.set_frame(payload)
            elif kind == "invite":
                ans = QMessageBox.question(self.main, f"{APP_NAME} - Screen share",
                                           f"🖥  {name} wants to share their screen with you.\n\nView now?")
                if ans == QMessageBox.Yes:
                    w = ViewerWindow(self, sid, frm, name)
                    self.viewers[sid] = w
                    w.show()
                    self.client.send("relay", to=frm, kind="accept", mode="screen", sid=sid)
                else:
                    self.client.send("relay", to=frm, kind="decline", mode="screen", sid=sid)
            elif kind == "accept" and self.share and self.share["sid"] == sid:
                self.share["viewers"].add(frm)
                self.share["bar"].set_viewers(len(self.share["viewers"]))
                self.main.status(f"{name} is now viewing your screen")
            elif kind == "decline" and self.share and self.share["sid"] == sid:
                self.main.status(f"{name} declined the screen share")
                if self.share["target"] != "ALL":
                    self.share["bar"].close()
            elif kind == "end":
                if sid in self.viewers:
                    w = self.viewers[sid]
                    w.closing_remote = True
                    w.close()
                    self.main.status(f"{name} stopped sharing the screen")
                elif self.share and self.share["sid"] == sid:
                    self.share["viewers"].discard(frm)
                    self.share["bar"].set_viewers(len(self.share["viewers"]))

    def stop_all(self):
        self.end_call(notify=True)
        if self.share:
            self.share["bar"].close()
        for sid, w in list(self.viewers.items()):
            w.close()


# =============================================================================
# 13. GUI : ATTENDANCE
# =============================================================================
class AttendancePage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Attendance", "Mark your daily attendance. Time is taken from the "
                                                     "server PC so that all entries are consistent."))
        g = QGroupBox("Today")
        h = QHBoxLayout(g)
        self.clock = QLabel()
        self.clock.setStyleSheet("font-size: 22pt; font-weight: 800; color: #e91e63;")
        h.addWidget(self.clock)
        v = QVBoxLayout()
        self.today_lbl = QLabel("—")
        self.today_lbl.setStyleSheet("font-size: 11pt; color:#1f2340; font-weight:600;")
        v.addWidget(self.today_lbl)
        mrow = QHBoxLayout()
        mrow.addWidget(QLabel("Work mode:"))
        self.mode = QComboBox()
        self.mode.addItems(["Office", "Work From Home", "Client Visit"])
        mrow.addWidget(self.mode)
        mrow.addStretch()
        v.addLayout(mrow)
        h.addLayout(v, 1)
        b1 = btn("✔  CHECK IN", "green", lambda: self.mark("in"))
        b2 = btn("⏏  CHECK OUT", "red", lambda: self.mark("out"))
        for b in (b1, b2):
            b.setMinimumHeight(48)
            b.setMinimumWidth(150)
            h.addWidget(b)
        self.lay.addWidget(g)
        t = QTimer(self)
        t.timeout.connect(lambda: self.clock.setText(datetime.now().strftime("%I:%M:%S %p")))
        t.start(1000)
        self.clock.setText(datetime.now().strftime("%I:%M:%S %p"))

        g2 = QGroupBox("Attendance Register")
        v2 = QVBoxLayout(g2)
        f = QHBoxLayout()
        f.addWidget(QLabel("Member:"))
        self.ucb = self.user_combo(True)
        f.addWidget(self.ucb)
        f.addWidget(QLabel("Month:"))
        self.month = month_edit()
        f.addWidget(self.month)
        f.addWidget(btn("Show", "blue", self.load))
        f.addWidget(btn("Export CSV", "orange", self.export))
        f.addStretch()
        v2.addLayout(f)
        self.summary = QLabel("")
        self.summary.setStyleSheet("color:#4a00e0; font-weight:600;")
        v2.addWidget(self.summary)
        self.table = make_table(["Date", "Day", "Member", "Check In", "Check Out", "Hours", "Mode"], 2)
        self.table.setMinimumHeight(330)
        v2.addWidget(self.table)
        self.lay.addWidget(g2)
        self.rows = []

    def refresh(self):
        try:
            t = self.client.call("att_today")
        except Exception as e:
            return self.main.status(str(e))
        if t:
            self.today_lbl.setText(f"Checked in at {t['check_in']}  •  Check out: {t.get('check_out') or '--'}"
                                   f"  •  {t['mode']}")
        else:
            self.today_lbl.setText("You have not checked in today")
        self.load()

    def mark(self, action):
        try:
            r = self.client.call("att_mark", action=action, mode=self.mode.currentText())
            self.main.status(("Checked in at " + r["check_in"]) if action == "in"
                             else ("Checked out at " + r["check_out"]))
            info(self, ("✔ Attendance marked. Check-in time: " + r["check_in"]) if action == "in"
                 else ("✔ Check-out recorded at " + r["check_out"]))
        except Exception as e:
            err(self, e)
        self.refresh()
        self.main.pages["dash"].refresh()

    @staticmethod
    def hours(a, b):
        try:
            t1 = datetime.strptime(a, "%H:%M:%S")
            t2 = datetime.strptime(b, "%H:%M:%S")
            s = (t2 - t1).seconds
            return f"{s // 3600}h {(s % 3600) // 60:02d}m"
        except Exception:
            return ""

    def load(self):
        try:
            self.rows = self.client.call("att_list", username=self.ucb.currentData(), month=ym_of(self.month))
        except Exception as e:
            return err(self, e)
        out = []
        for r in self.rows:
            d = date.fromisoformat(r["day"])
            out.append([d.strftime("%d-%m-%Y"), d.strftime("%A"), r.get("full_name") or r["username"],
                        r["check_in"], r.get("check_out") or "", self.hours(r["check_in"], r.get("check_out")),
                        r["mode"]])
        fill_table(self.table, out)
        wfh = sum(1 for r in self.rows if r["mode"] == "Work From Home")
        self.summary.setText(f"Total entries: {len(self.rows)}   |   Office: "
                             f"{len(self.rows) - wfh}   |   Work From Home: {wfh}")

    def export(self):
        if not self.rows:
            return err(self, "Nothing to export")
        os.makedirs(REPORT_DIR, exist_ok=True)
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", os.path.join(
            REPORT_DIR, f"Attendance_{ym_of(self.month)}.csv"), "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Date", "Member", "Username", "Check In", "Check Out", "Hours", "Mode"])
            for r in self.rows:
                w.writerow([r["day"], r.get("full_name"), r["username"], r["check_in"], r.get("check_out"),
                            self.hours(r["check_in"], r.get("check_out")), r["mode"]])
        self.main.status("Saved " + path)
        open_path(path)


# =============================================================================
# 14. GUI : LEAVE
# =============================================================================
class LeavePage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Leave Portal", f"{PAID_LEAVE_PER_MONTH} paid leaves per month. "
                                       "Approved leave beyond the monthly quota is Loss of Pay (LWP)."))
        top = QHBoxLayout()
        g = QGroupBox("Apply for Leave")
        f = QFormLayout(g)
        self.f_from = date_edit()
        self.f_to = date_edit()
        self.half = QCheckBox("Half day (single date only)")
        self.ltype = QComboBox()
        self.ltype.addItems(LEAVE_TYPES)
        self.reason = QLineEdit()
        self.reason.setPlaceholderText("Reason for leave")
        self.preview = QLabel("")
        self.preview.setStyleSheet("color:#e91e63; font-weight:700;")
        for w in (self.f_from, self.f_to):
            w.dateChanged.connect(self._preview)
        self.half.toggled.connect(self._preview)
        f.addRow("From:", self.f_from)
        f.addRow("To:", self.f_to)
        f.addRow("", self.half)
        f.addRow("Leave type:", self.ltype)
        f.addRow("Reason:", self.reason)
        f.addRow("Leave days:", self.preview)
        f.addRow(btn("Submit Application", "orange", self.apply))
        top.addWidget(g, 1)
        right = QVBoxLayout()
        sel = QHBoxLayout()
        sel.addWidget(QLabel("Member:"))
        self.ucb = self.user_combo(False)
        sel.addWidget(self.ucb)
        sel.addWidget(QLabel("Year:"))
        self.year = QSpinBox()
        self.year.setRange(2020, 2100)
        self.year.setValue(date.today().year)
        sel.addWidget(self.year)
        sel.addWidget(btn("Show", "blue", self.load))
        sel.addStretch()
        right.addLayout(sel)
        cg = QGridLayout()
        self.k_quota = Card("Paid Quota (month)", CARD_COLORS[2], "📅")
        self.k_used = Card("Paid Leave Used", CARD_COLORS[3], "🌴")
        self.k_lwp = Card("Loss of Pay Days", CARD_COLORS[0], "⚠")
        self.k_bal = Card("Paid Leave Balance", CARD_COLORS[1], "✅")
        for i, c in enumerate((self.k_quota, self.k_used, self.k_lwp, self.k_bal)):
            cg.addWidget(c, i // 2, i % 2)
        right.addLayout(cg)
        top.addLayout(right, 1)
        self.lay.addLayout(top)
        g2 = QGroupBox("Month-wise Leave Summary")
        v2 = QVBoxLayout(g2)
        self.sum_table = make_table(["Month", "Paid Quota", "Leave Taken (approved)", "Paid Leave",
                                     "Loss of Pay", "Paid Balance", "Pending Approval"], 0)
        self.sum_table.setMinimumHeight(260)
        v2.addWidget(self.sum_table)
        self.lay.addWidget(g2)
        g3 = QGroupBox("Leave Applications")
        v3 = QVBoxLayout(g3)
        hb = QHBoxLayout()
        if self.admin:
            hb.addWidget(QLabel("Show:"))
            self.scope = QComboBox()
            self.scope.addItem("All team members", "ALL")
            self.scope.addItem("Selected member only", "ONE")
            self.scope.currentIndexChanged.connect(self.load)
            hb.addWidget(self.scope)
            hb.addWidget(btn("✔ Approve", "green", lambda: self.decide("Approved")))
            hb.addWidget(btn("✖ Reject", "red", lambda: self.decide("Rejected")))
        hb.addWidget(btn("Cancel My Pending Leave", "gray", self.cancel))
        hb.addStretch()
        v3.addLayout(hb)
        self.app_table = make_table(["#", "Member", "From", "To", "Days", "Type", "Reason", "Status",
                                     "Decided By", "Applied On"], 6)
        self.app_table.setMinimumHeight(260)
        v3.addWidget(self.app_table)
        self.lay.addWidget(g3)
        self._preview()

    def _preview(self):
        f, t = self.f_from.date().toPyDate(), self.f_to.date().toPyDate()
        if self.half.isChecked():
            self.preview.setText("0.5 day" if f == t else "Half day needs same From & To date")
        elif t < f:
            self.preview.setText("'To' date is before 'From' date")
        else:
            self.preview.setText(f"{working_days(f, t)} working day(s)  (Sundays excluded)")

    def refresh(self):
        self.load()

    def load(self):
        u = self.ucb.currentData() or self.me
        try:
            summ = self.client.call("leave_summary", username=u, year=self.year.value())
            scope = "ALL" if (self.admin and self.scope.currentData() == "ALL") else u
            apps = self.client.call("leave_list", username=scope, year=self.year.value())
        except Exception as e:
            return err(self, e)
        t = date.today()
        cur = summ[t.month - 1] if self.year.value() == t.year else summ[0]
        self.k_quota.set(f"{cur['quota']}", month_label(cur["month"]))
        self.k_used.set(f"{cur['paid']:g}", f"Total approved: {cur['taken']:g}")
        self.k_lwp.set(f"{cur['lwp']:g}", "Deducted in salary")
        self.k_bal.set(f"{cur['balance']:g}", f"Pending approval: {cur['pending']:g}")
        rows = [[month_label(s["month"]), s["quota"], s["taken"], s["paid"], s["lwp"], s["balance"], s["pending"]]
                for s in summ]
        tot = [sum(s[k] for s in summ) for k in ("taken", "paid", "lwp", "pending")]
        rows.append(["TOTAL (Year)", PAID_LEAVE_PER_MONTH * 12, tot[0], tot[1], tot[2], "", tot[3]])
        fill_table(self.sum_table, rows)
        fill_table(self.app_table, [[a["id"], a.get("full_name") or a["username"],
                                     date.fromisoformat(a["from_date"]).strftime("%d-%m-%Y"),
                                     date.fromisoformat(a["to_date"]).strftime("%d-%m-%Y"), a["days"],
                                     a["ltype"] + (" (Half)" if a["half_day"] else ""), a["reason"],
                                     a["status"], a["decided_by"], a["applied_on"][:16]] for a in apps],
                   keys=[a["id"] for a in apps])

    def apply(self):
        try:
            r = self.client.call("leave_apply", from_date=self.f_from.date().toPyDate().isoformat(),
                                 to_date=self.f_to.date().toPyDate().isoformat(),
                                 half_day=self.half.isChecked(), ltype=self.ltype.currentText(),
                                 reason=self.reason.text())
            info(self, f"Leave application #{r['id']} submitted for {r['days']:g} day(s).\n"
                       "It will be counted once approved by the Manager.")
            self.reason.clear()
            self.load()
        except Exception as e:
            err(self, e)

    def decide(self, status):
        lid = selected_key(self.app_table)
        if not lid:
            return err(self, "Please select an application")
        try:
            self.client.call("leave_decide", id=lid, status=status)
            self.load()
            self.main.status(f"Leave #{lid} {status}")
        except Exception as e:
            err(self, e)

    def cancel(self):
        lid = selected_key(self.app_table)
        if not lid:
            return err(self, "Please select your pending application")
        if QMessageBox.question(self, APP_NAME, f"Cancel leave application #{lid}?") != QMessageBox.Yes:
            return
        try:
            self.client.call("leave_cancel", id=lid)
            self.load()
        except Exception as e:
            err(self, e)


# =============================================================================
# 15. GUI : SALARY SLIP
# =============================================================================
class SalaryPage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Salary Slip & Payout", "Generate a salary slip PDF for any month and the "
                                       "month-end payout (salary + approved reimbursement). "
                                       "Loss of Pay = (Monthly gross ÷ calendar days in month) × LWP days."))
        g = QGroupBox("Select")
        h = QHBoxLayout(g)
        h.addWidget(QLabel("Member:"))
        self.ucb = self.user_combo(False)
        h.addWidget(self.ucb)
        h.addWidget(QLabel("Month:"))
        self.month = month_edit()
        h.addWidget(self.month)
        h.addWidget(btn("Preview", "blue", self.preview))
        h.addWidget(btn("📄 Generate PDF", "orange", self.pdf))
        h.addStretch()
        self.lay.addWidget(g)
        g2 = QGroupBox("Preview")
        v = QVBoxLayout(g2)
        self.table = make_table(["Particulars", "Amount (Rs.) / Value"], 0)
        self.table.setMinimumHeight(430)
        v.addWidget(self.table)
        self.words = QLabel("")
        self.words.setStyleSheet("color:#4a00e0; font-weight:700;")
        self.words.setWordWrap(True)
        v.addWidget(self.words)
        self.lay.addWidget(g2)
        g3 = QGroupBox("Month-end Payout — Salary + Reimbursement")
        v3 = QVBoxLayout(g3)
        h3 = QHBoxLayout()
        h3.addWidget(QLabel("Month:"))
        self.pmonth = month_edit()
        h3.addWidget(self.pmonth)
        h3.addWidget(btn("⟳ Generate Payout", "green", self.load_payout))
        h3.addWidget(btn("📄 Payout PDF", "orange", self.payout_pdf))
        h3.addWidget(btn("Export CSV", "blue", self.payout_csv))
        h3.addStretch()
        v3.addLayout(h3)
        self.ptable = make_table(["Team Member", "Gross Salary", "LWP Days", "LWP Ded.", "PT", "TDS", "Net Salary",
                                  "Reimbursement (approved)", "TOTAL PAYOUT", "Reimb. Pending"])
        self.ptable.setMinimumHeight(240)
        v3.addWidget(self.ptable)
        self.ptot = QLabel("")
        self.ptot.setWordWrap(True)
        self.ptot.setStyleSheet("background:#f1ecff; border-radius:8px; padding:8px; color:#2a0a7a;")
        v3.addWidget(self.ptot)
        self.lay.addWidget(g3)
        self.pdata = None
        self.data = None

    def refresh(self):
        self.preview()
        self.load_payout()

    def load_payout(self):
        try:
            self.pdata = d = self.client.call("payout", username="ALL" if self.admin else self.me,
                                              month=ym_of(self.pmonth))
        except Exception as e:
            return self.main.status(str(e))
        rows = d["rows"]
        out = [[r["full_name"], inr(r["gross"]), r["lwp_days"], inr(r["lwp_amt"]), inr(r["pt"]), inr(r["tds"]),
                inr(r["net_salary"]), inr(r["reimb"]), inr(r["total"]), inr(r["pending_reimb"])] for r in rows]
        s = lambda k: sum(r[k] for r in rows)
        out.append(["TOTAL", inr(s("gross")), "", inr(s("lwp_amt")), inr(s("pt")), inr(s("tds")),
                    inr(s("net_salary")), inr(s("reimb")), inr(s("total")), inr(s("pending_reimb"))])
        fill_table(self.ptable, out)
        n = self.ptable.rowCount()
        for i in range(n):
            for j in range(1, 10):
                self.ptable.item(i, j).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.ptable.item(i, 8).setBackground(QColor("#ffe0ef"))
            f = self.ptable.item(i, 8).font()
            f.setBold(True)
            self.ptable.item(i, 8).setFont(f)
        for j in range(10):
            it = self.ptable.item(n - 1, j)
            it.setBackground(QColor("#4a00e0"))
            it.setForeground(QColor("white"))
        self.ptot.setText(f"<b>{month_label(d['month'])}:</b> Net salary Rs. {inr(s('net_salary'))} + Approved "
                          f"reimbursement Rs. {inr(s('reimb'))} = <b>Total payout Rs. {inr(s('total'))}</b>"
                          f"<br>Reimbursement still pending approval: Rs. {inr(s('pending_reimb'))} "
                          "(approve in the Reimbursement page to include it).")

    def payout_pdf(self):
        self.load_payout()
        if not self.pdata or not self.pdata["rows"]:
            return err(self, "No data for this month")
        os.makedirs(REPORT_DIR, exist_ok=True)
        p, _ = QFileDialog.getSaveFileName(self, "Save payout statement", os.path.join(
            REPORT_DIR, f"Payout_{self.pdata['month']}.pdf"), "PDF (*.pdf)")
        if p:
            try:
                make_payout_pdf(self.pdata, p)
                open_path(p)
            except PermissionError:
                err(self, "Cannot write the file. Is it open in a PDF viewer?")
            except Exception as e:
                err(self, e)

    def payout_csv(self):
        if not self.pdata or not self.pdata["rows"]:
            return err(self, "Generate the payout first")
        os.makedirs(REPORT_DIR, exist_ok=True)
        p, _ = QFileDialog.getSaveFileName(self, "Save CSV", os.path.join(
            REPORT_DIR, f"Payout_{self.pdata['month']}.csv"), "CSV (*.csv)")
        if not p:
            return
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Member", "Username", "Bank A/c", "Gross", "LWP Days", "LWP Deduction", "PT", "TDS",
                        "Net Salary", "Reimbursement", "Total Payout", "Reimbursement Pending"])
            for r in self.pdata["rows"]:
                w.writerow([r["full_name"], r["username"], r["bank"], r["gross"], r["lwp_days"], r["lwp_amt"],
                            r["pt"], r["tds"], r["net_salary"], r["reimb"], r["total"], r["pending_reimb"]])
        open_path(p)

    def fetch(self):
        self.data = self.client.call("payslip", username=self.ucb.currentData() or self.me,
                                     month=ym_of(self.month))
        return self.data

    def preview(self):
        try:
            d = self.fetch()
        except Exception as e:
            return err(self, e)
        u, lv = d["user"], d["leave"]
        rows = [["Employee", f"{u['full_name']} ({u['username']})"], ["Designation", u["designation"] or "-"],
                ["Month", month_label(d["month"])], ["Days in month", d["days_in_month"]],
                ["Paid days", f"{d['paid_days']:g}"], ["Days present (attendance)", d["present_days"]],
                ["Leave availed / Paid / LWP", f"{lv['taken']:g} / {lv['paid']:g} / {lv['lwp']:g}"],
                ["— EARNINGS —", ""]]
        rows += [[k, inr(v)] for k, v in d["earnings"]]
        rows += [["Gross Earnings", inr(d["gross"])], ["— DEDUCTIONS —", ""]]
        rows += [[k, inr(v)] for k, v in d["deductions"]]
        rows += [["Total Deductions", inr(d["total_deductions"])],
                 ["Add: Approved Reimbursements", inr(d["reimbursement"])], ["NET PAY", inr(d["net"])]]
        fill_table(self.table, rows)
        last = self.table.rowCount() - 1
        for c in range(2):
            it = self.table.item(last, c)
            it.setBackground(QColor("#4a00e0"))
            it.setForeground(QColor("white"))
        self.words.setText(d["net_words"])
        if not d["gross"]:
            self.words.setText(d["net_words"] + "   (Monthly salary not set — Admin can set it in Admin › Team Members)")

    def pdf(self):
        try:
            d = self.fetch()
        except Exception as e:
            return err(self, e)
        os.makedirs(REPORT_DIR, exist_ok=True)
        default = os.path.join(REPORT_DIR, f"Salary_Slip_{d['user']['username']}_{d['month']}.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "Save salary slip", default, "PDF (*.pdf)")
        if not path:
            return
        try:
            make_payslip_pdf(d, path)
        except PermissionError:
            return err(self, "Cannot write the file. Is it open in a PDF viewer? Close it and retry.")
        except Exception as e:
            return err(self, e)
        self.main.status("Salary slip saved: " + path)
        open_path(path)


# =============================================================================
# 16. GUI : REIMBURSEMENT
# =============================================================================
class EmailDialog(QDialog):
    def __init__(self, parent, cfg, to_addr, subject, body, attachment):
        super().__init__(parent)
        self.cfg, self.attachment = cfg, attachment
        self.setWindowTitle(f"{APP_NAME} - Send E-mail")
        self.setMinimumWidth(520)
        f = QFormLayout(self)
        self.to = QLineEdit(to_addr)
        self.subj = QLineEdit(subject)
        self.body = QTextEdit(body)
        self.body.setMinimumHeight(120)
        self.host = QLineEdit(cfg.get("smtp_host", "smtp.gmail.com"))
        self.port = QSpinBox()
        self.port.setRange(1, 65535)
        self.port.setValue(int(cfg.get("smtp_port", 587)))
        self.sec = QComboBox()
        self.sec.addItems(["STARTTLS", "SSL", "None"])
        self.sec.setCurrentText(cfg.get("smtp_security", "STARTTLS"))
        self.user = QLineEdit(cfg.get("smtp_user", ""))
        self.user.setPlaceholderText("your e-mail id (sender)")
        self.pw = QLineEdit()
        self.pw.setEchoMode(QLineEdit.Password)
        self.pw.setPlaceholderText("e-mail password / App Password (not saved)")
        f.addRow("To:", self.to)
        f.addRow("Subject:", self.subj)
        f.addRow("Message:", self.body)
        f.addRow(QLabel(f"<b>Attachment:</b> {html.escape(os.path.basename(attachment)) if attachment else 'none'}"))
        f.addRow("SMTP server:", self.host)
        f.addRow("Port:", self.port)
        f.addRow("Security:", self.sec)
        f.addRow("Sender e-mail:", self.user)
        f.addRow("Password:", self.pw)
        note = QLabel("Gmail: smtp.gmail.com, 587, STARTTLS + a 16-character App Password. "
                      "Outlook/Office 365: smtp.office365.com, 587, STARTTLS.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#6b6f8e; font-size:8.5pt;")
        f.addRow(note)
        h = QHBoxLayout()
        self.send_btn = btn("✉ Send", "green", self.send)
        h.addWidget(self.send_btn)
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)

    def send(self):
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", self.to.text().strip()):
            return err(self, "Enter a valid recipient e-mail address")
        if not self.user.text().strip() or not self.host.text().strip():
            return err(self, "Enter SMTP server and sender e-mail")
        self.cfg.update(smtp_host=self.host.text().strip(), smtp_port=self.port.value(),
                        smtp_security=self.sec.currentText(), smtp_user=self.user.text().strip())
        save_cfg(self.cfg)
        self.send_btn.setEnabled(False)
        self.send_btn.setText("Sending...")
        args = (dict(self.cfg), self.pw.text(), self.to.text().strip(), self.subj.text(),
                self.body.toPlainText(), self.attachment)

        def done(res, e):
            self.send_btn.setEnabled(True)
            self.send_btn.setText("✉ Send")
            if e:
                err(self, f"E-mail could not be sent:\n{e}")
            else:
                info(self, "E-mail sent successfully to " + args[2])
                self.accept()
        run_task(lambda: send_email(*args), done)


class ReimbPage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Reimbursement Claims", "Submit expenses with bills. At month-end "
                                       "generate the claim statement PDF and e-mail it to the Manager."))
        g = QGroupBox("New Claim")
        grid = QGridLayout(g)
        self.day = date_edit()
        self.day.setMaximumDate(QDate.currentDate())
        self.cat = QComboBox()
        self.cat.addItems(REIMB_CATEGORIES)
        self.amt = QDoubleSpinBox()
        self.amt.setRange(0, 1000000)
        self.amt.setDecimals(2)
        self.amt.setPrefix("Rs. ")
        self.amt.setGroupSeparatorShown(True)
        self.desc = QLineEdit()
        self.desc.setPlaceholderText("e.g. Auto fare to client office (Jaipur) / 200 pages xerox for audit file")
        self.bill_path = ""
        self.bill_lbl = QLabel("No bill attached")
        self.bill_lbl.setStyleSheet("color:#6b6f8e;")
        grid.addWidget(QLabel("Date:"), 0, 0)
        grid.addWidget(self.day, 0, 1)
        grid.addWidget(QLabel("Category:"), 0, 2)
        grid.addWidget(self.cat, 0, 3)
        grid.addWidget(QLabel("Amount:"), 0, 4)
        grid.addWidget(self.amt, 0, 5)
        grid.addWidget(QLabel("Description:"), 1, 0)
        grid.addWidget(self.desc, 1, 1, 1, 5)
        grid.addWidget(btn("📎 Attach Bill (optional)", "blue", self.attach), 2, 0, 1, 2)
        grid.addWidget(self.bill_lbl, 2, 2, 1, 2)
        grid.addWidget(btn("Submit Claim", "orange", self.submit), 2, 4, 1, 2)
        self.lay.addWidget(g)
        g2 = QGroupBox("Claims Register")
        v = QVBoxLayout(g2)
        f = QHBoxLayout()
        f.addWidget(QLabel("Member:"))
        self.ucb = self.user_combo(True)
        f.addWidget(self.ucb)
        f.addWidget(QLabel("Month:"))
        self.month = month_edit()
        f.addWidget(self.month)
        f.addWidget(btn("Show", "blue", self.load))
        f.addStretch()
        v.addLayout(f)
        a = QHBoxLayout()
        if self.admin:
            a.addWidget(btn("✔ Approve", "green", lambda: self.decide("Approved")))
            a.addWidget(btn("✖ Reject", "red", lambda: self.decide("Rejected")))
        a.addWidget(btn("Delete My Pending Claim", "gray", self.delete))
        a.addWidget(btn("⬇ Download Bill", "blue", self.download_bill))
        a.addStretch()
        a.addWidget(btn("📄 Generate Monthly PDF", "orange", self.make_pdf))
        a.addWidget(btn("✉ E-mail PDF to Manager", None, self.email_pdf))
        v.addLayout(a)
        self.table = make_table(["#", "Date", "Member", "Category", "Description", "Amount (Rs.)", "Bill",
                                 "Status", "Decided By"], 4)
        self.table.setMinimumHeight(300)
        v.addWidget(self.table)
        self.totals = QLabel("")
        self.totals.setWordWrap(True)
        self.totals.setStyleSheet("background:#f1ecff; border-radius:8px; padding:8px; color:#2a0a7a;")
        v.addWidget(self.totals)
        self.lay.addWidget(g2)
        self.rows, self.settings = [], {}
        self.last_pdf = None

    def refresh(self):
        self.load()

    def attach(self):
        p, _ = QFileDialog.getOpenFileName(self, "Select bill / receipt", "",
                                           "Bills (*.pdf *.jpg *.jpeg *.png);;All files (*.*)")
        if p:
            if os.path.getsize(p) > 20 * 1024 * 1024:
                return err(self, "Bill file must be under 20 MB")
            self.bill_path = p
            self.bill_lbl.setText("📎 " + os.path.basename(p))

    def submit(self):
        payload = b""
        if self.bill_path:
            try:
                with open(self.bill_path, "rb") as fh:
                    payload = fh.read()
            except Exception as e:
                return err(self, e)
        try:
            r = self.client.request("reimb_add", payload, timeout=120,
                                    day=self.day.date().toPyDate().isoformat(), category=self.cat.currentText(),
                                    amount=self.amt.value(), description=self.desc.text(),
                                    bill_name=os.path.basename(self.bill_path) if self.bill_path else "")[0]
        except Exception as e:
            return err(self, e)
        info(self, f"Claim #{r['id']} submitted for Rs. {inr(self.amt.value())}")
        self.desc.clear()
        self.amt.setValue(0)
        self.bill_path = ""
        self.bill_lbl.setText("No bill attached")
        self.load()

    def load(self):
        try:
            d = self.client.call("reimb_list", username=self.ucb.currentData(), month=ym_of(self.month))
        except Exception as e:
            return err(self, e)
        self.rows, self.settings = d["rows"], d["settings"]
        fill_table(self.table, [[r["id"], date.fromisoformat(r["day"]).strftime("%d-%m-%Y"),
                                 r.get("full_name") or r["username"], r["category"], r["description"],
                                 inr(r["amount"]), "Yes" if r["file_id"] else "No", r["status"], r["decided_by"]]
                                for r in self.rows], keys=[r["id"] for r in self.rows])
        for i in range(self.table.rowCount()):
            self.table.item(i, 5).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        bycat = {}
        for r in self.rows:
            if r["status"] != "Rejected":
                bycat[r["category"]] = bycat.get(r["category"], 0) + r["amount"]
        tot = sum(bycat.values())
        appr = sum(r["amount"] for r in self.rows if r["status"] == "Approved")
        self.totals.setText("  |  ".join(f"<b>{k}</b>: Rs. {inr(v)}" for k, v in bycat.items()) +
                            f"<br><b>Total claimed: Rs. {inr(tot)}</b> &nbsp;&nbsp; Approved: Rs. {inr(appr)}"
                            f" &nbsp;&nbsp; Pending: Rs. {inr(sum(r['amount'] for r in self.rows if r['status'] == 'Pending'))}")

    def decide(self, status):
        rid = selected_key(self.table)
        if not rid:
            return err(self, "Please select a claim")
        try:
            self.client.call("reimb_decide", id=rid, status=status)
            self.load()
        except Exception as e:
            err(self, e)

    def delete(self):
        rid = selected_key(self.table)
        if not rid:
            return err(self, "Please select your pending claim")
        if QMessageBox.question(self, APP_NAME, f"Delete claim #{rid}?") != QMessageBox.Yes:
            return
        try:
            self.client.call("reimb_delete", id=rid)
            self.load()
        except Exception as e:
            err(self, e)

    def download_bill(self):
        rid = selected_key(self.table)
        r = next((x for x in self.rows if x["id"] == rid), None)
        if not r:
            return err(self, "Please select a claim")
        if not r["file_id"]:
            return err(self, "No bill attached to this claim")
        self.main.download_file(r["file_id"])

    def make_pdf(self, ask_path=True):
        self.load()
        if not self.rows:
            err(self, "There are no claims for this month")
            return None
        sel = self.ucb.currentData()
        label = "All Team Members" if sel == "ALL" else (self.rows[0].get("full_name") or self.rows[0]["username"])
        tag = "ALL" if sel == "ALL" else self.rows[0]["username"]
        os.makedirs(REPORT_DIR, exist_ok=True)
        path = os.path.join(REPORT_DIR, f"Reimbursement_{tag}_{ym_of(self.month)}.pdf")
        if ask_path:
            path, _ = QFileDialog.getSaveFileName(self, "Save claim statement", path, "PDF (*.pdf)")
            if not path:
                return None
        try:
            _, claimed, _ = make_reimb_pdf(self.rows, self.settings, ym_of(self.month), label, path)
        except PermissionError:
            err(self, "Cannot write the file. Is it open in a PDF viewer? Close it and retry.")
            return None
        except Exception as e:
            err(self, e)
            return None
        self.last_pdf = (path, label, claimed)
        if ask_path:
            open_path(path)
        self.main.status("Claim statement saved: " + path)
        return self.last_pdf

    def email_pdf(self):
        res = self.make_pdf(ask_path=False)
        if not res:
            return
        path, label, claimed = res
        ml = month_label(ym_of(self.month))
        body = (f"Dear Sir/Madam,\n\nPlease find attached the reimbursement claim statement of {label} "
                f"for {ml} amounting to Rs. {inr(claimed)}.\n\nKindly approve and arrange reimbursement.\n\n"
                f"Regards,\n{self.main.user['full_name']}\n{self.settings.get('company_name', FIRM_NAME)}")
        EmailDialog(self, self.main.cfg, self.settings.get("manager_email", ""),
                    f"Reimbursement Claim - {label} - {ml}", body, path).exec_()


# =============================================================================
# 16A. GUI : LITIGATION TRACKER
# =============================================================================
def opt_date_row(iso=""):
    """Checkbox + date edit (for optional dates). Returns (widget, getter)."""
    w = QWidget()
    h = QHBoxLayout(w)
    h.setContentsMargins(0, 0, 0, 0)
    cb = QCheckBox("Scheduled")
    de = date_edit(QDate.fromString(iso, "yyyy-MM-dd") if iso else QDate.currentDate().addDays(7))
    cb.setChecked(bool(iso))
    de.setEnabled(bool(iso))
    cb.toggled.connect(de.setEnabled)
    h.addWidget(cb)
    h.addWidget(de, 1)
    return w, (lambda: de.date().toPyDate().isoformat() if cb.isChecked() else "")


class CaseDialog(QDialog):
    def __init__(self, parent, users, c=None):
        super().__init__(parent)
        c = c or {}
        self.setWindowTitle(f"{APP_NAME} - {'Edit Case #%s' % c['id'] if c.get('id') else 'New Litigation Case'}")
        self.setMinimumWidth(600)
        f = QFormLayout(self)
        self.w = {}

        def le(k, ph=""):
            e = QLineEdit(str(c.get(k) or ""))
            e.setPlaceholderText(ph)
            self.w[k] = e
            return e
        self.law = QComboBox()
        self.law.addItems(LIT_LAWS)
        self.law.setCurrentText(c.get("law") or LIT_LAWS[0])
        self.forum = QComboBox()
        self.forum.setEditable(True)
        self.law.currentTextChanged.connect(self._forums)
        self._forums(self.law.currentText())
        if c.get("forum"):
            self.forum.setCurrentText(c["forum"])
        f.addRow("Law*:", self.law)
        f.addRow("Client name*:", le("client_name"))
        row = QHBoxLayout()
        row.addWidget(le("client_pan", "PAN"))
        row.addWidget(le("client_gstin", "GSTIN"))
        f.addRow("Client PAN / GSTIN:", row)
        f.addRow("Client address:", le("client_address", "used on invoice"))
        f.addRow("Client e-mail:", le("client_email"))
        f.addRow("Client mobile / WhatsApp:", le("client_phone", "10-digit mobile"))
        f.addRow("AY / FY / Period:", le("period", "e.g. AY 2023-24  /  FY 2019-20  /  Apr-17 to Mar-18"))
        f.addRow("Notice / DIN / Ref no.:", le("ref_no"))
        f.addRow("Forum / Authority:", self.forum)
        f.addRow("Section:", le("section", "e.g. 143(3) / 148A / 73 / 74 / 61"))
        f.addRow("Issue / Nature of dispute:", le("nature", "e.g. ITC mismatch GSTR-2B vs 3B"))
        self.demand = QDoubleSpinBox()
        self.demand.setRange(0, 1e10)
        self.demand.setDecimals(2)
        self.demand.setPrefix("Rs. ")
        self.demand.setGroupSeparatorShown(True)
        self.demand.setValue(float(c.get("demand") or 0))
        f.addRow("Demand / amount in dispute:", self.demand)
        self.est = QDoubleSpinBox()
        self.est.setRange(0, 1e9)
        self.est.setDecimals(2)
        self.est.setPrefix("Rs. ")
        self.est.setGroupSeparatorShown(True)
        self.est.setValue(float(c.get("est_fee") or 0))
        if parent.main.user["role"] == "admin":
            f.addRow("Estimated professional fee:", self.est)
        self.assign = QComboBox()
        self.assign.addItem("— Not assigned —", "")
        for u in users:
            self.assign.addItem(u["full_name"], u["username"])
        i = self.assign.findData(c.get("assigned_to") or "")
        self.assign.setCurrentIndex(max(0, i))
        f.addRow("Handled by:", self.assign)
        self.status = QComboBox()
        self.status.addItems(LIT_STATUSES)
        self.status.setCurrentText(c.get("status") or "Open")
        f.addRow("Status:", self.status)
        nh_w, self.nh = opt_date_row(c.get("next_hearing") or "")
        f.addRow("Next hearing / due date:", nh_w)
        f.addRow("Purpose of next date:", le("next_purpose", "e.g. Personal hearing / Reply to SCN due"))
        self.remarks = QTextEdit(c.get("remarks") or "")
        self.remarks.setMaximumHeight(70)
        f.addRow("Remarks:", self.remarks)
        h = QHBoxLayout()
        h.addWidget(btn("Save Case", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)

    def _forums(self, law):
        cur = self.forum.currentText()
        self.forum.clear()
        self.forum.addItems(LIT_FORUMS.get(law, []))
        if cur and cur not in LIT_FORUMS.get(law, []):
            self.forum.setCurrentText(cur)

    def value(self):
        d = {k: e.text().strip() for k, e in self.w.items()}
        d.update(law=self.law.currentText(), forum=self.forum.currentText().strip(), demand=self.demand.value(),
                 est_fee=self.est.value(),
                 assigned_to=self.assign.currentData(), status=self.status.currentText(),
                 next_hearing=self.nh(), remarks=self.remarks.toPlainText().strip())
        return d


class ProceedingDialog(QDialog):
    def __init__(self, parent, c):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - Hearing / Submission update - Case #{c['id']}")
        self.setMinimumWidth(580)
        self.file_path = ""
        f = QFormLayout(self)
        f.addRow(QLabel(f"<b>{html.escape(c['client_name'])}</b> &nbsp; {html.escape(c['law'])} &nbsp; "
                        f"{html.escape(c['period'] or '')} &nbsp; {html.escape(c['forum'] or '')}"))
        self.pdate = date_edit()
        self.ptype = QComboBox()
        self.ptype.addItems(LIT_PTYPES)
        self.ptype.setCurrentText("Hearing Attended")
        self.details = QTextEdit()
        self.details.setPlaceholderText("Report: who attended, what was submitted / argued, officer's "
                                        "observations, documents called for...")
        self.details.setMinimumHeight(120)
        self.outcome = QLineEdit()
        self.outcome.setPlaceholderText("e.g. Adjourned / Order reserved / Demand reduced to Rs. ...")
        nd_w, self.nd = opt_date_row("")
        self.purpose = QLineEdit(c.get("next_purpose") or "")
        self.purpose.setPlaceholderText("Purpose of next date")
        self.status = QComboBox()
        self.status.addItem("— No change (%s) —" % c["status"], "")
        for s in LIT_STATUSES:
            self.status.addItem(s, s)
        self.clear_next = QCheckBox("No further date (clear next hearing)")
        self.att = QLabel("No document attached")
        self.att.setStyleSheet("color:#6b6f8e;")
        f.addRow("Date:", self.pdate)
        f.addRow("Type:", self.ptype)
        f.addRow("Report*:", self.details)
        f.addRow("Outcome:", self.outcome)
        f.addRow("Next hearing date:", nd_w)
        f.addRow("Next purpose:", self.purpose)
        f.addRow("", self.clear_next)
        f.addRow("Update case status:", self.status)
        ab = QHBoxLayout()
        ab.addWidget(btn("📎 Attach submission / order", "blue", self._attach))
        ab.addWidget(self.att, 1)
        f.addRow(ab)
        h = QHBoxLayout()
        h.addWidget(btn("Save Update", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)

    def _attach(self):
        p, _ = QFileDialog.getOpenFileName(self, "Attach document")
        if p:
            if os.path.getsize(p) > 50 * 1024 * 1024:
                return err(self, "Attachment must be under 50 MB")
            self.file_path = p
            self.att.setText("📎 " + os.path.basename(p))

    def value(self):
        return dict(pdate=self.pdate.date().toPyDate().isoformat(), ptype=self.ptype.currentText(),
                    details=self.details.toPlainText().strip(), outcome=self.outcome.text().strip(),
                    next_date=self.nd(), next_purpose=self.purpose.text().strip(),
                    clear_next=self.clear_next.isChecked(), new_status=self.status.currentData(),
                    file_name=os.path.basename(self.file_path) if self.file_path else "")


class EntryDialog(QDialog):
    """Generic date + description + amount dialog (fees / out-of-pocket expenses)."""

    def __init__(self, parent, title, with_billable=False, desc_ph=""):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - {title}")
        self.setMinimumWidth(460)
        f = QFormLayout(self)
        self.d = date_edit()
        self.desc = QLineEdit()
        self.desc.setPlaceholderText(desc_ph)
        self.amt = QDoubleSpinBox()
        self.amt.setRange(0, 1e9)
        self.amt.setDecimals(2)
        self.amt.setPrefix("Rs. ")
        self.amt.setGroupSeparatorShown(True)
        f.addRow("Date:", self.d)
        f.addRow("Description*:", self.desc)
        f.addRow("Amount*:", self.amt)
        self.billable = QCheckBox("Recoverable from client (include in next invoice)")
        self.billable.setChecked(True)
        if with_billable:
            f.addRow("", self.billable)
        h = QHBoxLayout()
        h.addWidget(btn("Save", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)


class InvoiceDialog(QDialog):
    """Raise a tax invoice for a litigation case: optionally add a fee line now, then bill all un-billed items."""

    def __init__(self, parent, settings, fee_total, ope_total, client_gstin, est=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - Raise Invoice to Client")
        self.setMinimumWidth(680)
        self.fee_total, self.ope_total = fee_total, ope_total
        f = QFormLayout(self)
        bal_est = 0.0
        if est:
            bal_est = max(0.0, round((est.get("est_fee") or 0) - est.get("billed_fees", 0) - fee_total, 2))
            el = QLabel(f"<b>{html.escape(est.get('client_name', ''))}</b> &nbsp; {html.escape(est.get('law', ''))} "
                        f"{html.escape(est.get('period') or '')}<br>Estimated fee: <b>Rs. {inr(est.get('est_fee') or 0)}"
                        f"</b> &nbsp;|&nbsp; Already billed: Rs. {inr(est.get('billed_fees', 0))} &nbsp;|&nbsp; "
                        f"Balance of estimate: <b>Rs. {inr(bal_est + fee_total)}</b>")
            el.setWordWrap(True)
            el.setStyleSheet("background:#fff4e5; border-radius:8px; padding:6px;")
            f.addRow(el)
        fl = QLabel(f"Fee entries not yet billed: <b>Rs. {inr(fee_total)}</b> &nbsp; | &nbsp; "
                    f"Recoverable out-of-pocket expenses: <b>Rs. {inr(ope_total)}</b>")
        fl.setWordWrap(True)
        f.addRow(fl)
        g = QGroupBox("Add professional fee now (optional)")
        gf = QFormLayout(g)
        self.fdesc = QLineEdit(f"Professional fees - {est.get('law', '')} matter before {est.get('forum', '')}"
                               f" {est.get('period') or ''}".strip() if est else "Professional fees")
        self.famt = QDoubleSpinBox()
        self.famt.setRange(0, 1e9)
        self.famt.setDecimals(2)
        self.famt.setPrefix("Rs. ")
        self.famt.setGroupSeparatorShown(True)
        self.famt.setValue(bal_est if fee_total == 0 else 0)
        gf.addRow("Description:", self.fdesc)
        gf.addRow("Fee amount:", self.famt)
        f.addRow(g)
        self.d = date_edit()
        self.mode = QComboBox()
        self.mode.addItems(["CGST+SGST", "IGST"])
        fs, cs = (settings.get("firm_gstin") or "")[:2], (client_gstin or "")[:2]
        if fs and cs and fs != cs:
            self.mode.setCurrentText("IGST")
        self.rate = QDoubleSpinBox()
        self.rate.setRange(0, 40)
        self.rate.setSuffix(" %")
        self.rate.setValue(float(settings.get("gst_rate") or 18))
        self.gst_type = settings.get("gst_type") or "Regular"
        if self.gst_type != "Regular":
            self.mode.addItem("No GST")
            self.mode.setCurrentText("No GST")
            self.mode.setEnabled(False)
            self.rate.setValue(0)
            self.rate.setEnabled(False)
            self.mode.setToolTip(f"Firm is {self.gst_type} under GST (Firm Master) — no GST is charged")
        self.pure = QCheckBox("Expenses as pure-agent reimbursement (no GST)")
        self.pure.setChecked(True)
        self.prev = QLabel()
        self.prev.setStyleSheet("color:#6d28d9; font-weight:700; font-size:11pt;")
        for w in (self.mode.currentTextChanged, self.rate.valueChanged, self.pure.toggled, self.famt.valueChanged):
            w.connect(self._calc)
        f.addRow("Invoice date:", self.d)
        f.addRow("GST type:", self.mode)
        f.addRow("GST rate:", self.rate)
        if self.gst_type != "Regular":
            gn = QLabel(f"ℹ Your firm is <b>{self.gst_type}</b> under GST (Firm Master) — no GST will be charged; "
                        f"the document will be titled <b>{GST_DOC_TITLE[self.gst_type].title()}</b>.")
            gn.setWordWrap(True)
            gn.setStyleSheet("background:#ecfeff; border-radius:8px; padding:6px; color:#155e75;")
            f.addRow(gn)
        f.addRow("", self.pure)
        self.sdet = QTextEdit()
        self.sdet.setPlaceholderText("Additional details of service (optional) — e.g. drafted & filed appeal in Form "
                                     "36, prepared paper book, attended 3 hearings before ITAT Bench 'B'...")
        self.sdet.setMaximumHeight(90)
        f.addRow("Service details:", self.sdet)
        f.addRow("Invoice value:", self.prev)
        h = QHBoxLayout()
        h.addWidget(btn("🧾 Raise Invoice", "green", self._ok))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)
        self._calc()

    def _calc(self, *a):
        fees = self.fee_total + self.famt.value()
        taxable = fees + (0 if self.pure.isChecked() else self.ope_total)
        gst = round(taxable * self.rate.value() / 100, 2)
        self.prev.setText(f"Taxable Rs. {inr(taxable)} + GST Rs. {inr(gst)} → Total Rs. {inr(fees + self.ope_total + gst)}")

    def _ok(self):
        if self.fee_total + self.famt.value() + self.ope_total <= 0:
            return err(self, "Enter the fee amount (nothing else is pending to bill in this case).")
        if self.famt.value() > 0 and not self.fdesc.text().strip():
            return err(self, "Enter the fee description")
        self.accept()

    def value(self):
        return dict(inv_date=self.d.date().toPyDate().isoformat(), gst_mode=self.mode.currentText(),
                    gst_rate=self.rate.value(), pure_agent=self.pure.isChecked(),
                    service_details=self.sdet.toPlainText().strip())


class ReceiptDialog(QDialog):
    """Update payment received against an invoice (new receipt, or correct an existing one)."""

    def __init__(self, parent, inv, rc=None):
        super().__init__(parent)
        self.inv, self.rc = inv, rc
        self.prior = round(inv["received"] + inv["tds"] - ((rc["amount"] + rc["tds"]) if rc else 0), 2)
        bal_before = round(inv["total"] - self.prior, 2)
        self.setWindowTitle(f"{APP_NAME} - {'Edit' if rc else 'Update'} Payment Received - {inv['inv_no']}")
        self.setMinimumWidth(500)
        f = QFormLayout(self)
        summ = QLabel(f"<b>{html.escape(inv.get('client_name') or '')}</b> &nbsp; Invoice <b>{html.escape(inv['inv_no'])}"
                      f"</b> dated {dmy(inv['inv_date'])}<br>Invoice value: <b>Rs. {inr(inv['total'])}</b> &nbsp;|&nbsp; "
                      f"Received so far (incl. TDS){' - other receipts' if rc else ''}: Rs. {inr(self.prior)} &nbsp;|&nbsp; "
                      f"<b style='color:#c2410c'>Balance due: Rs. {inr(bal_before)}</b>")
        summ.setWordWrap(True)
        summ.setStyleSheet("background:#f5f3ff; border-radius:8px; padding:8px;")
        f.addRow(summ)

        def money(v):
            sp = QDoubleSpinBox()
            sp.setRange(0, 1e10)
            sp.setDecimals(2)
            sp.setPrefix("Rs. ")
            sp.setGroupSeparatorShown(True)
            sp.setValue(v)
            return sp
        self.d = date_edit(QDate.fromString(rc["rdate"], "yyyy-MM-dd") if rc else None)
        self.d.setMaximumDate(QDate.currentDate())
        self.amt = money(rc["amount"] if rc else bal_before)
        self.tds = money(rc["tds"] if rc else 0)
        self.mode = QComboBox()
        self.mode.addItems(PAY_MODES)
        if rc and rc.get("mode") in PAY_MODES:
            self.mode.setCurrentText(rc["mode"])
        self.ref = QLineEdit(rc.get("ref", "") if rc else "")
        self.ref.setPlaceholderText("UTR / cheque no. / UPI ref")
        self.after = QLabel()
        self.after.setStyleSheet("font-weight:700; font-size:11pt;")
        self.amt.valueChanged.connect(self._calc)
        self.tds.valueChanged.connect(self._calc)
        f.addRow("Date of receipt:", self.d)
        f.addRow("Amount received:", self.amt)
        f.addRow("TDS deducted by client:", self.tds)
        f.addRow("Mode:", self.mode)
        f.addRow("Reference:", self.ref)
        f.addRow("Balance due after this:", self.after)
        h = QHBoxLayout()
        h.addWidget(btn("💾 Save && Generate Receipt", "green", self._ok))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)
        self._calc()

    def _calc(self, *a):
        bal = round(self.inv["total"] - self.prior - self.amt.value() - self.tds.value(), 2)
        if bal < -0.01:
            self.after.setText(f"<span style='color:#dc2626'>Exceeds invoice by Rs. {inr(-bal)}</span>")
        elif bal <= 0.01:
            self.after.setText("<span style='color:#059669'>Rs. 0.00 — invoice fully paid ✔</span>")
        else:
            self.after.setText(f"<span style='color:#c2410c'>Rs. {inr(bal)} (part payment)</span>")

    def _ok(self):
        if self.amt.value() + self.tds.value() <= 0:
            return err(self, "Enter the amount received (and/or TDS)")
        if self.inv["total"] - self.prior - self.amt.value() - self.tds.value() < -0.01:
            return err(self, "Amount is more than the balance due on this invoice")
        self.accept()

    def value(self):
        return dict(rdate=self.d.date().toPyDate().isoformat(), amount=self.amt.value(), tds=self.tds.value(),
                    mode=self.mode.currentText(), ref=self.ref.text().strip())


class CaseFileDialog(QDialog):
    def __init__(self, page, case_id):
        super().__init__(page.main)
        self.page, self.main, self.client, self.cid = page, page.main, page.client, case_id
        self.admin = self.main.user["role"] == "admin"
        self.setWindowTitle(f"{APP_NAME} - Case File #{case_id}")
        self.setWindowIcon(app_icon())
        self.resize(1150, 740)
        lay = QVBoxLayout(self)
        self.head = QLabel()
        self.head.setWordWrap(True)
        self.head.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.head)
        top = QHBoxLayout()
        top.addWidget(btn("✎ Edit Case", "blue", self.edit))
        top.addWidget(btn("➕ Add Hearing / Submission", "orange", self.add_proc))
        top.addWidget(btn("📄 Case Status Report PDF", None, self.report))
        top.addWidget(btn("📨 Send Details / Reminder", "green", self.share))
        top.addWidget(btn("🖨 Hearing Brief PDF", "blue", lambda: self.main.hearing_note(self.cid)))
        if self.admin:
            top.addWidget(btn("🧾 Raise Invoice", "orange", self.make_invoice))
            top.addWidget(btn("🔒 Access", "gray", self.access))
        top.addStretch()
        top.addWidget(btn("⟳ Refresh", "gray", self.refresh))
        lay.addLayout(top)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs, 1)
        # proceedings
        w = QWidget()
        v = QVBoxLayout(w)
        self.t_proc = make_table(["Date", "Type", "Hearing / Submission Report", "Outcome", "Next Date", "By",
                                  "Attachment"], 2)
        self.t_proc.setWordWrap(True)
        v.addWidget(self.t_proc)
        h = QHBoxLayout()
        h.addWidget(btn("🖨 Print Hearing Note (selected entry)", "orange", self.print_note))
        h.addWidget(btn("⬇ Download Attachment", "blue", self.dl_att))
        if self.admin:
            h.addWidget(btn("Delete Entry", "red", self.del_proc))
        h.addStretch()
        v.addLayout(h)
        self.tabs.addTab(w, "🗂 Hearings && Submissions")
        # OPE
        w = QWidget()
        v = QVBoxLayout(w)
        self.t_ope = make_table(["#", "Date", "Description", "Amount (Rs.)", "Recoverable", "Invoice", "By"], 2)
        v.addWidget(self.t_ope)
        self.ope_sum = QLabel()
        v.addWidget(self.ope_sum)
        h = QHBoxLayout()
        h.addWidget(btn("➕ Add Out-of-Pocket Expense", "orange", self.add_ope))
        h.addWidget(btn("Delete", "red", self.del_ope))
        h.addStretch()
        v.addLayout(h)
        self.tabs.addTab(w, "🧾 Out-of-Pocket Expenses")
        if self.admin:
            w = QWidget()
            v = QVBoxLayout(w)
            self.t_fee = make_table(["#", "Date", "Description", "Amount (Rs.)", "Invoice"], 2)
            v.addWidget(self.t_fee)
            h = QHBoxLayout()
            h.addWidget(btn("➕ Add Fee", "green", self.add_fee))
            h.addWidget(btn("Delete Un-billed Fee", "red", self.del_fee))
            h.addWidget(btn("➕ Fee = Balance of Estimate", "orange", self.fee_from_est))
            h.addStretch()
            h.addWidget(btn("🧾 Create Tax Invoice", None, self.make_invoice))
            v.addLayout(h)
            self.tabs.addTab(w, "💵 Fees")
            w = QWidget()
            v = QVBoxLayout(w)
            self.t_inv = make_table(["Invoice No.", "Date", "Fees", "OPE", "GST", "Total", "Received", "TDS",
                                     "Balance Due", "Status"], 0)
            v.addWidget(self.t_inv)
            h = QHBoxLayout()
            h.addWidget(btn("📄 Invoice PDF", "orange", lambda: self.main.invoice_pdf(selected_key(self.t_inv))))
            h.addWidget(btn("📨 Send Invoice", "blue", lambda: self.main.send_invoice(selected_key(self.t_inv))))
            h.addWidget(btn("🧾 Raise Invoice", "orange", self.make_invoice))
            h.addWidget(btn("₹ Payment Received", "green", self.receipt))
            h.addWidget(btn("Cancel Invoice", "red", self.cancel_inv))
            h.addStretch()
            v.addLayout(h)
            self.tabs.addTab(w, "📑 Invoices && Receipts")
        self.docs = DocsWidget(self.main, "case", case_id)
        self.tabs.addTab(self.docs, "📁 Case Documents")
        self.data = None
        self.refresh()

    def refresh(self):
        try:
            self.data = d = self.client.call("case_get", id=self.cid)
        except Exception as e:
            return err(self, e)
        c = d["case"]
        col = LAW_COLORS.get(c["law"], "#4a00e0")
        nh = (f"<b style='color:#e91e63'>{dmy(c['next_hearing'])}</b> — {html.escape(c['next_purpose'] or '')}"
              if c["next_hearing"] else "—")
        E = lambda x: html.escape(str(x or "—"))
        self.head.setText(
            f"<div style='font-size:15pt;font-weight:800;color:{col}'>#{c['id']} &nbsp; {E(c['client_name'])}"
            f" &nbsp;<span style='font-size:10pt;color:white;background:{col}'>&nbsp;{E(c['law'])}&nbsp;</span></div>"
            f"<b>Period:</b> {E(c['period'])} &nbsp; <b>Forum:</b> {E(c['forum'])} &nbsp; <b>Section:</b> "
            f"{E(c['section'])} &nbsp; <b>Ref:</b> {E(c['ref_no'])}<br><b>Issue:</b> {E(c['nature'])} &nbsp; "
            f"<b>Demand:</b> Rs. {inr(c['demand'])} &nbsp; <b>Status:</b> {E(c['status'])} &nbsp; "
            f"<b>Handled by:</b> {E(c.get('assigned_name'))}<br><b>Next hearing:</b> {nh}"
            + (f"<br><b>Invoiced (incl. GST):</b> Rs. {inr(c['inv_billed'])} &nbsp; <b>Received:</b> Rs. "
               f"{inr(c['inv_received'])} &nbsp; <b style='color:#c2410c'>Balance due: Rs. {inr(c['balance_due'])}</b>"
               if c.get("inv_billed") is not None else "")
            + ("<br><span style='color:#b91c1c'><b>🔒 Confidential case</b> — visible only to permitted members</span>"
               if c.get("restricted") else "")
            + (f"<br><b>Estimated fee:</b> Rs. {inr(c['est_fee'])} &nbsp; <b>Billed:</b> Rs. {inr(c['billed_fees'])}"
               f" &nbsp; <b>Un-billed fee entries:</b> Rs. {inr(c['unbilled_fees'])} &nbsp; <b>Balance of estimate:</b>"
               f" Rs. {inr(c['est_fee'] - c['billed_fees'] - c['unbilled_fees'])}" if c.get("inv_billed") is not None
               else ""))
        fill_table(self.t_proc, [[dmy(p["pdate"]), p["ptype"], p["details"], p["outcome"],
                                  dmy(p["next_date"]) if p["next_date"] else "", p.get("by_name") or p["by_user"],
                                  ("📎 " + p["file_name"]) if p.get("file_name") else ""] for p in d["proceedings"]],
                   keys=[p["id"] for p in d["proceedings"]])
        self.t_proc.resizeRowsToContents()
        fill_table(self.t_ope, [[o["id"], dmy(o["edate"]), o["description"], inr(o["amount"]),
                                 "Yes" if o["billable"] else "No", o.get("inv_no") or "",
                                 o.get("by_name") or o["by_user"]] for o in d["ope"]], keys=[o["id"] for o in d["ope"]])
        tot = sum(o["amount"] for o in d["ope"])
        unb = sum(o["amount"] for o in d["ope"] if o["billable"] and not o["invoice_id"])
        self.ope_sum.setText(f"<b>Total out-of-pocket: Rs. {inr(tot)}</b> &nbsp;&nbsp; Recoverable, not yet "
                             f"invoiced: Rs. {inr(unb)}")
        if self.admin:
            fill_table(self.t_fee, [[f["id"], dmy(f["fdate"]), f["description"], inr(f["amount"]),
                                     f.get("inv_no") or "Un-billed"] for f in d["fees"]],
                       keys=[f["id"] for f in d["fees"]])
            fill_table(self.t_inv, [[i["inv_no"], dmy(i["inv_date"]), inr(i["fee_total"]), inr(i["ope_total"]),
                                     inr(i["cgst"] + i["sgst"] + i["igst"]), inr(i["total"]), inr(i["received"]),
                                     inr(i["tds"]),
                                     inr(0 if i["status"] == "Cancelled" else i["total"] - i["received"] - i["tds"]),
                                     i["status"]] for i in d["invoices"]], keys=[i["id"] for i in d["invoices"]])

    def edit(self):
        self.page.edit_case(self.cid)
        self.refresh()

    def share(self):
        self.main.share_case(self.data["case"])

    def access(self):
        c = self.data["case"]
        if AccessDialog(self.main, "case", self.cid, f"Case #{c['id']} — {c['client_name']} ({c['law']})").exec_():
            self.refresh()

    def fee_from_est(self):
        c = self.data["case"]
        bal = round((c["est_fee"] or 0) - c["billed_fees"] - c["unbilled_fees"], 2)
        if bal <= 0:
            return err(self, "No balance left in the estimated fee (set the estimate in Edit Case).")
        if QMessageBox.question(self, APP_NAME, f"Add a fee entry of Rs. {inr(bal)} (balance of estimate)?") \
                == QMessageBox.Yes:
            try:
                self.client.call("fee_add", case_id=self.cid, fdate=date.today().isoformat(),
                                 description=f"Professional fees - {c['law']} {c['forum']} {c['period']}".strip(),
                                 amount=bal)
                self.refresh()
            except Exception as e:
                err(self, e)

    def add_proc(self):
        if self.page.add_proc(self.data["case"]):
            self.refresh()

    def report(self):
        os.makedirs(REPORT_DIR, exist_ok=True)
        c = self.data["case"]
        p, _ = QFileDialog.getSaveFileName(self, "Save case report", os.path.join(
            REPORT_DIR, f"Case_{c['id']}_{safe_filename(c['client_name'])}_Report.pdf"), "PDF (*.pdf)")
        if p:
            try:
                make_case_report_pdf(self.data, p)
                open_path(p)
            except Exception as e:
                err(self, e)

    def print_note(self):
        pid = selected_key(self.t_proc)
        if not pid:
            return err(self, "Select a hearing / submission entry (or use 'Hearing Brief PDF' for the next hearing)")
        self.main.hearing_note(self.cid, pid)

    def dl_att(self):
        pid = selected_key(self.t_proc)
        p = next((x for x in self.data["proceedings"] if x["id"] == pid), None)
        if not p or not p.get("file_id"):
            return err(self, "Select an entry that has an attachment")
        self.main.download_file(p["file_id"])

    def del_proc(self):
        pid = selected_key(self.t_proc)
        if pid and QMessageBox.question(self, APP_NAME, "Delete this entry?") == QMessageBox.Yes:
            try:
                self.client.call("proc_delete", id=pid)
                self.refresh()
            except Exception as e:
                err(self, e)

    def add_ope(self):
        dlg = EntryDialog(self, "Out-of-Pocket Expense", True, "e.g. Travel to ITAT Jaipur / Paper book printing")
        if dlg.exec_():
            try:
                self.client.call("ope_add", case_id=self.cid, edate=dlg.d.date().toPyDate().isoformat(),
                                 description=dlg.desc.text(), amount=dlg.amt.value(),
                                 billable=dlg.billable.isChecked())
                self.refresh()
            except Exception as e:
                err(self, e)

    def del_ope(self):
        oid = selected_key(self.t_ope)
        if oid and QMessageBox.question(self, APP_NAME, "Delete this expense?") == QMessageBox.Yes:
            try:
                self.client.call("ope_delete", id=oid)
                self.refresh()
            except Exception as e:
                err(self, e)

    def add_fee(self):
        dlg = EntryDialog(self, "Professional Fee", False, "e.g. Drafting & filing reply to SCN u/s 73")
        if dlg.exec_():
            try:
                self.client.call("fee_add", case_id=self.cid, fdate=dlg.d.date().toPyDate().isoformat(),
                                 description=dlg.desc.text(), amount=dlg.amt.value())
                self.refresh()
            except Exception as e:
                err(self, e)

    def del_fee(self):
        fid = selected_key(self.t_fee)
        if fid and QMessageBox.question(self, APP_NAME, "Delete this fee entry?") == QMessageBox.Yes:
            try:
                self.client.call("fee_delete", id=fid)
                self.refresh()
            except Exception as e:
                err(self, e)

    def make_invoice(self):
        """Raise invoice to client (adds a fee line now if entered, then bills all un-billed fees / expenses)."""
        d = self.data
        if not d:
            return
        fees = sum(f["amount"] for f in d["fees"] if not f["invoice_id"])
        ope = sum(o["amount"] for o in d["ope"] if o["billable"] and not o["invoice_id"])
        dlg = InvoiceDialog(self, d["settings"], fees, ope, d["case"]["client_gstin"], d["case"])
        if not dlg.exec_():
            return False
        try:
            if dlg.famt.value() > 0:
                self.client.call("fee_add", case_id=self.cid, fdate=dlg.d.date().toPyDate().isoformat(),
                                 description=dlg.fdesc.text().strip(), amount=dlg.famt.value())
            r = self.client.call("invoice_create", case_id=self.cid, **dlg.value())
        except Exception as e:
            err(self, e)
            self.refresh()
            return False
        self.refresh()
        if self.admin:
            self.tabs.setCurrentIndex(3)
        if QMessageBox.question(self, APP_NAME, f"Invoice {r['inv_no']} raised for Rs. {inr(r['total'])}.\n\n"
                                                "Send it to the client by e-mail / WhatsApp now?\n"
                                                "(No = just open the PDF)") == QMessageBox.Yes:
            self.main.send_invoice(r["id"])
        else:
            p = self.main.invoice_pdf(r["id"], ask=False)
            if p:
                open_path(p)
        return True

    def receipt(self):
        iid = selected_key(self.t_inv)
        inv = next((i for i in self.data["invoices"] if i["id"] == iid), None)
        if self.main.pages["lit"].record_receipt(inv):
            self.refresh()

    def cancel_inv(self):
        iid = selected_key(self.t_inv)
        if iid and QMessageBox.question(self, APP_NAME, "Cancel this invoice? Its fees and expenses become "
                                                        "un-billed again.") == QMessageBox.Yes:
            try:
                self.client.call("invoice_cancel", id=iid)
                self.refresh()
            except Exception as e:
                err(self, e)


class LitigationPage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Litigation Tracker", "Income Tax • GST • VAT • Service Tax — hearings, "
                                       "submissions, reminders, fees, invoices and out-of-pocket expenses."))
        cg = QHBoxLayout()
        self.k_open = Card("Open Cases", CARD_COLORS[2], "⚖")
        self.k_7 = Card("Hearings in next 7 days", CARD_COLORS[0], "📅")
        self.k_od = Card("Past date — update pending", CARD_COLORS[5], "⏰")
        self.k_dem = Card("Demand under dispute", CARD_COLORS[3], "₹")
        cards = [self.k_open, self.k_7, self.k_od, self.k_dem]
        if self.admin:
            self.k_out = Card("Fees outstanding", CARD_COLORS[1], "💵")
            cards.append(self.k_out)
        for c in cards:
            cg.addWidget(c)
        for c in (self.k_dem, getattr(self, "k_out", None)):
            if c:
                c.small()
        self.lay.addLayout(cg)
        self.tabs = QTabWidget()
        self.tabs.setMinimumHeight(560)
        self.lay.addWidget(self.tabs)
        self._build_cases()
        self._build_calendar()
        if self.admin:
            self._build_invoices()
        self.tabs.currentChanged.connect(lambda _: self.refresh())
        self.rows = []

    # ------------------------------------------------------------- cases tab
    def _build_cases(self):
        w = QWidget()
        v = QVBoxLayout(w)
        f = QHBoxLayout()
        self.f_law = QComboBox()
        self.f_law.addItems(["All laws"] + LIT_LAWS)
        self.f_grp = QComboBox()
        self.f_grp.addItems(["Open", "Closed", "All"])
        self.f_search = QLineEdit()
        self.f_search.setPlaceholderText("Search client / PAN / GSTIN / ref no / AY / section / forum ...")
        self.f_search.returnPressed.connect(self.load_cases)
        for cb in (self.f_law, self.f_grp):
            cb.currentIndexChanged.connect(self.load_cases)
        f.addWidget(QLabel("Law:"))
        f.addWidget(self.f_law)
        f.addWidget(QLabel("Status:"))
        f.addWidget(self.f_grp)
        f.addWidget(self.f_search, 1)
        f.addWidget(btn("Search", "blue", self.load_cases))
        v.addLayout(f)
        a = QHBoxLayout()
        a.addWidget(btn("➕ New", "green", self.new_case))
        a.addWidget(btn("📂 Open File", None, self.open_case))
        a.addWidget(btn("📝 Add Hearing", "orange", self.quick_proc))
        a.addWidget(btn("✎ Edit", "blue", lambda: self.edit_case(selected_key(self.table))))
        a.addWidget(btn("📨 Reminder", "green", self.share_sel))
        a.addWidget(btn("🖨 Brief", "blue", self.note_sel))
        if self.admin:
            a.addWidget(btn("🧾 Invoice", "orange", self.invoice_sel))
            a.addWidget(btn("🔒 Access", "gray", self.access_sel))
        if self.admin:
            a.addWidget(btn("Delete", "red", self.delete_case))
        a.addStretch()
        v.addLayout(a)
        heads = ["#", "Law", "Client", "AY / Period", "Forum", "Section / Issue", "Demand (Rs.)", "Handled by",
                 "Next Hearing", "Purpose", "Status"] + (["Est. Fee (Rs.)", "Invoiced (Rs.)", "Received (Rs.)",
                                                         "Balance Due (Rs.)"] if self.admin else [])
        self.table = make_table(heads, 5)
        self.table.doubleClicked.connect(self.open_case)
        v.addWidget(self.table)
        legend = QLabel("<span style='background:#ffcdd2'>&nbsp;Past date / today&nbsp;</span> &nbsp; "
                        "<span style='background:#ffe0b2'>&nbsp;Within reminder days&nbsp;</span> &nbsp; "
                        "Double-click a case to open its file.")
        lg = QHBoxLayout()
        lg.addWidget(legend, 1)
        lg.addWidget(btn("📊 MIS PDF", "orange", self.mis_pdf))
        lg.addWidget(btn("📆 Export .ics", "blue", self.export_ics))
        v.addLayout(lg)
        self.tabs.addTab(w, "⚖ Cases")

    def load_cases(self):
        law = self.f_law.currentText()
        try:
            d = self.client.call("case_list", law=law if law in LIT_LAWS else "", group=self.f_grp.currentText(),
                                 search=self.f_search.text())
        except Exception as e:
            return self.main.status(str(e))
        self.rows, s = d["rows"], d["summary"]
        self.k_open.set(s["open"], "across all laws")
        self.k_7.set(s["next7"], "firm-wide")
        self.k_od.set(s["overdue"], "record the outcome of these")
        self.k_dem.set("Rs. " + inr(s["demand"])[:-3], "open cases")
        if self.admin:
            self.k_out.set("Rs. " + inr(s.get("outstanding", 0))[:-3], "invoices not yet realised")
        days = int(float(self.main.settings.get("reminder_days") or 3))
        t = date.today()
        out = []
        for r in self.rows:
            line = [r["id"], r["law"], ("🔒 " if r.get("restricted") else "") + r["client_name"], r["period"], r["forum"],
                    (r["section"] + "  " + r["nature"]).strip(), inr(r["demand"]), r.get("assigned_name") or "",
                    dmy(r["next_hearing"]) if r["next_hearing"] else "", r["next_purpose"], r["status"]]
            if self.admin:
                line += [inr(r["est_fee"]) if r["est_fee"] else "", inr(r["billed"]) if r["billed"] else "",
                         inr(r["realised"]) if r["realised"] else "", inr(r["outstanding"]) if r["billed"] else ""]
            out.append(line)
        fill_table(self.table, out, keys=[r["id"] for r in self.rows])
        for i, r in enumerate(self.rows):
            self.table.item(i, 1).setForeground(QColor(LAW_COLORS.get(r["law"], "#333")))
            fnt = self.table.item(i, 1).font()
            fnt.setBold(True)
            self.table.item(i, 1).setFont(fnt)
            self.table.item(i, 6).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            if r["next_hearing"] and r["status"] in LIT_OPEN:
                nd = date.fromisoformat(r["next_hearing"])
                if nd <= t:
                    self.table.item(i, 8).setBackground(QColor("#ffcdd2"))
                elif nd <= t + timedelta(days=days):
                    self.table.item(i, 8).setBackground(QColor("#ffe0b2"))

    def _users(self):
        return self.main.users

    def new_case(self):
        dlg = CaseDialog(self, self._users())
        if dlg.exec_():
            try:
                r = self.client.call("case_save", case=dlg.value())
                self.main.status(f"Case #{r['id']} created")
                self.load_cases()
            except Exception as e:
                err(self, e)

    def edit_case(self, cid):
        if not cid:
            return err(self, "Please select a case")
        try:
            c = self.client.call("case_get", id=cid)["case"]
        except Exception as e:
            return err(self, e)
        dlg = CaseDialog(self, self._users(), c)
        if dlg.exec_():
            try:
                self.client.call("case_save", id=cid, case=dlg.value())
                self.load_cases()
            except Exception as e:
                err(self, e)

    def open_case(self, *a):
        cid = selected_key(self.table)
        if not cid:
            return err(self, "Please select a case")
        CaseFileDialog(self, cid).exec_()
        self.refresh()

    def add_proc(self, c):
        dlg = ProceedingDialog(self, c)
        if not dlg.exec_():
            return False
        v = dlg.value()
        payload = b""
        if dlg.file_path:
            try:
                with open(dlg.file_path, "rb") as fh:
                    payload = fh.read()
            except Exception as e:
                err(self, e)
                return False
        try:
            self.client.request("proc_add", payload, timeout=300, case_id=c["id"], **v)
            self.main.status(f"Case #{c['id']} updated")
            return True
        except Exception as e:
            err(self, e)
            return False

    def quick_proc(self):
        cid = selected_key(self.table)
        c = next((r for r in self.rows if r["id"] == cid), None)
        if not c:
            return err(self, "Please select a case")
        if self.add_proc(c):
            self.load_cases()

    def delete_case(self):
        cid = selected_key(self.table)
        if cid and QMessageBox.question(self, APP_NAME, f"Permanently delete case #{cid} with all its hearings,"
                                                        " fees and expenses?") == QMessageBox.Yes:
            try:
                self.client.call("case_delete", id=cid)
                self.load_cases()
            except Exception as e:
                err(self, e)

    def mis_pdf(self):
        if not self.rows:
            return err(self, "No cases in the current list")
        os.makedirs(REPORT_DIR, exist_ok=True)
        p, _ = QFileDialog.getSaveFileName(self, "Save MIS report", os.path.join(
            REPORT_DIR, f"Litigation_MIS_{date.today():%Y-%m-%d}.pdf"), "PDF (*.pdf)")
        if p:
            title = (f"LITIGATION MIS - {self.f_law.currentText().upper()} - {self.f_grp.currentText().upper()} "
                     f"CASES AS ON {date.today():%d-%m-%Y}")
            try:
                make_lit_mis_pdf(self.rows, self.main.settings, p, title)
                open_path(p)
            except Exception as e:
                err(self, e)

    def export_ics(self):
        t = date.today()
        try:
            ev = self.client.call("hearings", start=t.isoformat(), end=(t + timedelta(days=366)).isoformat(),
                                  history=False)
        except Exception as e:
            return err(self, e)
        if not ev:
            return err(self, "No upcoming hearings to export")
        os.makedirs(REPORT_DIR, exist_ok=True)
        p, _ = QFileDialog.getSaveFileName(self, "Save calendar file", os.path.join(
            REPORT_DIR, "Suthar_Hearings.ics"), "Calendar (*.ics)")
        if p:
            make_ics(ev, p)
            info(self, f"{len(ev)} hearing(s) exported with reminders 1 day and 2 hours before.\n\n"
                       "Double-click the .ics file to add them to Outlook, or import it in Google Calendar "
                       "(Settings › Import).")
            open_path(os.path.dirname(p))

    # ---------------------------------------------------------- calendar tab
    def _build_calendar(self):
        w = QWidget()
        h = QHBoxLayout(w)
        left = QVBoxLayout()
        self.cal = QCalendarWidget()
        self.cal.setGridVisible(True)
        self.cal.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.cal.setMinimumSize(430, 330)
        self.cal.selectionChanged.connect(self.show_day)
        self.cal.currentPageChanged.connect(lambda y, m: self.load_month())
        left.addWidget(self.cal)
        left.addWidget(QLabel("<span style='background:#ff80ab'>&nbsp;&nbsp;&nbsp;</span> Hearing scheduled "
                              "&nbsp; <span style='background:#b3e5fc'>&nbsp;&nbsp;&nbsp;</span> Hearing / "
                              "submission recorded"))
        left.addStretch()
        h.addLayout(left)
        right = QVBoxLayout()
        self.day_lbl = QLabel()
        self.day_lbl.setStyleSheet("font-size:13pt; font-weight:800; color:#4a00e0;")
        right.addWidget(self.day_lbl)
        self.t_day = make_table(["#", "Type", "Client", "Law", "Forum", "Purpose / Report"], 5)
        self.t_day.doubleClicked.connect(lambda: self._open_from(self.t_day))
        right.addWidget(self.t_day, 1)
        lb = QLabel("Upcoming hearings — next 30 days")
        lb.setStyleSheet("font-size:12pt; font-weight:800; color:#e91e63;")
        right.addWidget(lb)
        self.t_up = make_table(["#", "Date", "Day", "Client", "Law", "Forum", "Purpose"], 6)
        self.t_up.doubleClicked.connect(lambda: self._open_from(self.t_up))
        right.addWidget(self.t_up, 1)
        h.addLayout(right, 1)
        self.tabs.addTab(w, "📅 Hearing Calendar")
        self.month_ev = []

    def _open_from(self, table):
        cid = selected_key(table)
        if cid:
            CaseFileDialog(self, cid).exec_()
            self.refresh()

    def load_month(self):
        y, m = self.cal.yearShown(), self.cal.monthShown()
        first, last = month_bounds(y, m)
        try:
            self.month_ev = self.client.call("hearings", start=(first - timedelta(days=7)).isoformat(),
                                             end=(last + timedelta(days=14)).isoformat())
            t = date.today()
            up = self.client.call("hearings", start=t.isoformat(), end=(t + timedelta(days=30)).isoformat(),
                                  history=False)
        except Exception as e:
            return self.main.status(str(e))
        self.cal.setDateTextFormat(QDate(), QTextCharFormat())
        days = {}
        for e in self.month_ev:
            days.setdefault(e["date"], set()).add("H" if e["kind"] == "Hearing" else "P")
        for dstr, kinds in days.items():
            fmt = QTextCharFormat()
            fmt.setFontWeight(QFont.Bold)
            fmt.setBackground(QColor("#ff80ab" if "H" in kinds else "#b3e5fc"))
            self.cal.setDateTextFormat(QDate.fromString(dstr, "yyyy-MM-dd"), fmt)
        fill_table(self.t_up, [[e["case_id"], dmy(e["date"]), date.fromisoformat(e["date"]).strftime("%a"),
                                e["client"], e["law"], e["forum"], e["text"]] for e in up],
                   keys=[e["case_id"] for e in up])
        self.show_day()

    def show_day(self):
        d = self.cal.selectedDate().toPyDate().isoformat()
        ev = [e for e in self.month_ev if e["date"] == d]
        self.day_lbl.setText(self.cal.selectedDate().toString("dddd, dd MMMM yyyy") + f"  —  {len(ev)} item(s)")
        fill_table(self.t_day, [[e["case_id"], e["kind"], e["client"], e["law"], e["forum"], e["text"]] for e in ev],
                   keys=[e["case_id"] for e in ev])

    # ---------------------------------------------------------- invoices tab
    def _build_invoices(self):
        w = QWidget()
        v = QVBoxLayout(w)
        f = QHBoxLayout()
        f.addWidget(QLabel("Status:"))
        self.inv_f = QComboBox()
        self.inv_f.addItems(["All", "Unpaid", "Part Paid", "Paid", "Cancelled"])
        self.inv_f.currentIndexChanged.connect(self.load_invoices)
        f.addWidget(self.inv_f)
        f.addWidget(btn("📄 Invoice PDF", "orange", lambda: self.invoice_pdf(selected_key(self.t_inv))))
        f.addWidget(btn("₹ Payment Received", "green", self._receipt_sel))
        f.addWidget(btn("📨 Send Invoice", "blue", lambda: self.main.send_invoice(selected_key(self.t_inv))))
        f.addWidget(btn("Export CSV", "blue", self.export_inv))
        f.addStretch()
        v.addLayout(f)
        self.t_inv = make_table(["Invoice No.", "Date", "Client", "Law", "Fees", "OPE", "GST", "Total", "Received",
                                 "TDS", "Balance Due", "Status"], 2)
        v.addWidget(self.t_inv)
        self.inv_sum = QLabel()
        self.inv_sum.setStyleSheet("background:#f1ecff; border-radius:8px; padding:8px; color:#2a0a7a;")
        v.addWidget(self.inv_sum)
        self.tabs.addTab(w, "💵 Fees && Invoices")
        self.invs = []

    def load_invoices(self):
        st = self.inv_f.currentText()
        try:
            self.invs = self.client.call("invoice_list", status=st if st != "All" else "", source="case")
        except Exception as e:
            return self.main.status(str(e))
        fill_table(self.t_inv, [[i["inv_no"], dmy(i["inv_date"]), i["client_name"], i["law"], inr(i["fee_total"]),
                                 inr(i["ope_total"]), inr(i["cgst"] + i["sgst"] + i["igst"]), inr(i["total"]),
                                 inr(i["received"]), inr(i["tds"]), inr(i["outstanding"]), i["status"]]
                                for i in self.invs], keys=[i["id"] for i in self.invs])
        live = [i for i in self.invs if i["status"] != "Cancelled"]
        s = lambda k: sum(i[k] for i in live)
        self.inv_sum.setText(f"<b>Invoices:</b> {len(live)} &nbsp; <b>Billed:</b> Rs. {inr(s('total'))} &nbsp; "
                             f"<b>GST:</b> Rs. {inr(s('cgst') + s('sgst') + s('igst'))} &nbsp; <b>Received:</b> Rs. "
                             f"{inr(s('received'))} &nbsp; <b>TDS:</b> Rs. {inr(s('tds'))} &nbsp; "
                             f"<b>Outstanding:</b> Rs. {inr(s('outstanding'))}")

    def _receipt_sel(self):
        iid = selected_key(self.t_inv)
        if self.record_receipt(next((i for i in self.invs if i["id"] == iid), None)):
            self.load_invoices()

    def record_receipt(self, inv):
        if not inv:
            err(self, "Please select an invoice")
            return False
        if inv["status"] in ("Paid", "Cancelled"):
            err(self, f"Invoice is already {inv['status']}")
            return False
        dlg = ReceiptDialog(self, inv)
        if dlg.exec_():
            try:
                r = self.client.call("receipt_add", invoice_id=inv["id"], **dlg.value())
                self.main.status(f"Receipt {r['rcpt_no']} recorded. Invoice {inv['inv_no']} is now {r['status']}")
                self.main.after_receipt(r["id"])
                return True
            except Exception as e:
                err(self, e)
        return False

    def invoice_pdf(self, iid):
        self.main.invoice_pdf(iid)

    def note_sel(self):
        cid = selected_key(self.table)
        if not cid:
            return err(self, "Please select a case")
        self.main.hearing_note(cid)

    def access_sel(self):
        cid = selected_key(self.table)
        c = next((r for r in self.rows if r["id"] == cid), None)
        if not c:
            return err(self, "Please select a case")
        if AccessDialog(self.main, "case", cid, f"Case #{cid} — {c['client_name']} ({c['law']})").exec_():
            self.refresh()

    def invoice_sel(self):
        cid = selected_key(self.table)
        if not cid:
            return err(self, "Please select a case")
        dlg = CaseFileDialog(self, cid)
        dlg.make_invoice()
        self.refresh()

    def share_sel(self):
        cid = selected_key(self.table)
        c = next((r for r in self.rows if r["id"] == cid), None)
        if not c:
            return err(self, "Please select a case")
        self.main.share_case(c)

    def export_inv(self):
        if not self.invs:
            return err(self, "Nothing to export")
        os.makedirs(REPORT_DIR, exist_ok=True)
        p, _ = QFileDialog.getSaveFileName(self, "Save CSV", os.path.join(REPORT_DIR, "Fee_Register.csv"),
                                           "CSV (*.csv)")
        if not p:
            return
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Invoice No", "Date", "Client", "Law", "Period", "Fees", "OPE", "CGST", "SGST", "IGST",
                        "Total", "Received", "TDS", "Outstanding", "Status"])
            for i in self.invs:
                w.writerow([i["inv_no"], i["inv_date"], i["client_name"], i["law"], i["period"], i["fee_total"],
                            i["ope_total"], i["cgst"], i["sgst"], i["igst"], i["total"], i["received"], i["tds"],
                            i["outstanding"], i["status"]])
        open_path(p)

    def refresh(self):
        self.load_cases()          # also keeps the summary cards current
        idx = self.tabs.currentIndex()
        if idx == 1:
            self.load_month()
        elif idx == 2 and self.admin:
            self.load_invoices()


# =============================================================================
# 16B. GUI : TASK TRACKER & BILLING
# =============================================================================
class TaskDialog(QDialog):
    def __init__(self, parent, users, me, t=None):
        super().__init__(parent)
        t = t or {}
        self.setWindowTitle(f"{APP_NAME} - {'Edit Task' if t.get('id') else 'New Task'}")
        self.setMinimumWidth(560)
        f = QFormLayout(self)
        self.title = QComboBox()
        self.title.setEditable(True)
        self.title.addItems([""] + TASK_PRESETS)
        self.title.setCurrentText(t.get("title", ""))
        self.title.lineEdit().setPlaceholderText("Type a task or pick a common one")
        self.client = QLineEdit(t.get("client", ""))
        self.client.setPlaceholderText("Client name (optional)")
        self.desc = QTextEdit(t.get("description", ""))
        self.desc.setMaximumHeight(90)
        self.assign = QComboBox()
        for u in users:
            self.assign.addItem(f"{u['full_name']}" + ("  (Manager)" if u.get("role") == "admin" else ""),
                                u["username"])
        i = self.assign.findData(t.get("assigned_to") or me)
        self.assign.setCurrentIndex(max(0, i))
        self.due = date_edit(QDate.fromString(t["due_date"], "yyyy-MM-dd") if t.get("due_date")
                             else QDate.currentDate().addDays(7))
        self.urgent = QCheckBox("🔥 URGENT (shown at the top of everyone's list)")
        self.urgent.setChecked(bool(t.get("urgent")))
        self.repeat = QComboBox()
        self.repeat.addItems(TASK_REPEATS)
        self.hint = QLabel()
        self.hint.setWordWrap(True)
        self.hint.setStyleSheet("color:#6b6f8e; font-size:8.5pt;")
        self.due.dateChanged.connect(self._hint)
        self.repeat.currentTextChanged.connect(self._hint)
        f.addRow("Task*:", self.title)
        f.addRow("Client:", self.client)
        f.addRow("Details:", self.desc)
        f.addRow("Assign to:", self.assign)
        f.addRow("Due date*:", self.due)
        f.addRow("", self.urgent)
        self.est = QDoubleSpinBox()
        self.est.setRange(0, 1e9)
        self.est.setDecimals(2)
        self.est.setPrefix("Rs. ")
        self.est.setGroupSeparatorShown(True)
        self.est.setValue(float(t.get("est_amount") or 0))
        if t.get("fee_hidden"):
            f.addRow("Estimated billing:", QLabel("🔒 Hidden by Manager"))
        else:
            f.addRow("Estimated billing:", self.est)
        if not t.get("id"):
            f.addRow("Repeat automatically:", self.repeat)
            f.addRow("", self.hint)
        h = QHBoxLayout()
        h.addWidget(btn("Save Task", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)
        self._hint()

    def _hint(self, *a):
        r = self.repeat.currentText()
        if r == "None":
            self.hint.setText("One-time task. For GST / TDS returns choose Monthly / Quarterly — a fresh task is "
                              "created automatically at the start of each period.")
        else:
            d = self.due.date().toPyDate()
            nxt = add_months(d, REPEAT_MONTHS[r])
            self.hint.setText(f"{r}: a new task due on day {d.day} will be created automatically every "
                              f"{REPEAT_MONTHS[r]} month(s). Next one: due {nxt:%d-%m-%Y} "
                              f"(created on 01-{nxt:%m-%Y}).")

    def value(self):
        return dict(title=self.title.currentText().strip(), client=self.client.text().strip(),
                    description=self.desc.toPlainText().strip(), assigned_to=self.assign.currentData(),
                    due_date=self.due.date().toPyDate().isoformat(), urgent=self.urgent.isChecked(),
                    repeat=self.repeat.currentText(), est_amount=self.est.value())


class CompleteTaskDialog(QDialog):
    def __init__(self, parent, t):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - Complete Task")
        self.setMinimumWidth(520)
        f = QFormLayout(self)
        f.addRow(QLabel(f"<b>{html.escape(t['title'])}</b>" + (f" — {html.escape(t['client'])}" if t["client"] else "")))
        self.opt = QComboBox()
        self.opt.addItems(BILL_OPTIONS)
        self.amt = QDoubleSpinBox()
        self.amt.setRange(0, 1e9)
        self.amt.setDecimals(2)
        self.amt.setPrefix("Rs. ")
        self.amt.setGroupSeparatorShown(True)
        self.opt.currentTextChanged.connect(lambda s: self.amt.setEnabled(s != "No Billing"))
        self.ch = QTextEdit()
        self.ch.setPlaceholderText("Challenges faced (data delay from client, portal issues, notices ...)")
        self.ch.setMaximumHeight(100)
        note = QLabel("Billing Due / Yearly Billing → an entry 'To be Billed' is created in the Billing tab.  "
                      "Billing Raised → recorded as 'Billed'.  No Billing → nothing is created.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#6b6f8e; font-size:8.5pt;")
        f.addRow("Billing:", self.opt)
        if t.get("fee_hidden"):
            f.addRow("Billing amount:", QLabel("🔒 Fee is confidential — the Manager's estimate will be used"))
        else:
            f.addRow("Billing amount:", self.amt)
        self.amt.setValue(float(t.get("est_amount") or 0))
        if t.get("est_amount"):
            f.addRow("", QLabel(f"Estimated at task creation: <b>Rs. {inr(t['est_amount'])}</b> (edit if different)"))
        f.addRow(note)
        f.addRow("Remarks / challenges:", self.ch)
        h = QHBoxLayout()
        h.addWidget(btn("✔ Mark Completed", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)

    def value(self):
        return dict(billing_option=self.opt.currentText(), billing_amount=self.amt.value(),
                    challenges=self.ch.toPlainText().strip())


class TextDialog(QDialog):
    def __init__(self, parent, title, label, text=""):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - {title}")
        self.setMinimumWidth(480)
        v = QVBoxLayout(self)
        v.addWidget(QLabel(label))
        self.t = QTextEdit(text)
        v.addWidget(self.t)
        h = QHBoxLayout()
        h.addWidget(btn("Save", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        v.addLayout(h)


class BillUpdateDialog(QDialog):
    def __init__(self, parent, b):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - Update Billing #{b['id']}")
        self.setMinimumWidth(420)
        f = QFormLayout(self)
        f.addRow(QLabel(f"<b>{html.escape(b['description'] or '')}</b> — {html.escape(b['client'] or '')}"))
        self.st = QComboBox()
        self.st.addItems(BILL_STATUSES)
        self.st.setCurrentText(b["status"])
        self.ref = QLineEdit(b.get("ref") or "")
        self.ref.setPlaceholderText("Invoice no. / receipt reference")
        self.amt = QDoubleSpinBox()
        self.amt.setRange(0, 1e9)
        self.amt.setDecimals(2)
        self.amt.setPrefix("Rs. ")
        self.amt.setGroupSeparatorShown(True)
        self.amt.setValue(float(b["amount"] or 0))
        f.addRow("Status:", self.st)
        f.addRow("Invoice / Ref:", self.ref)
        f.addRow("Amount:", self.amt)
        h = QHBoxLayout()
        h.addWidget(btn("Save", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)


class TaskPage(Page):
    SCOPES = [("mine", "📥 My Tasks"), ("byme", "📤 Assigned by Me"), ("shared", "👁 Shared with me")]

    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Task Tracker", "Assign work to anyone in the team. Urgent tasks first, then "
                                       "by due date. Colour: green > 7 days, yellow 4–7 days, red ≤ 3 days / overdue."))
        cg = QHBoxLayout()
        self.k_open = Card("My Open Tasks", CARD_COLORS[2], "📋")
        self.k_red = Card("Due ≤ 3 days / Overdue", ("#e53935", "#ff6f61"), "🔴")
        self.k_yel = Card("Due in 4–7 days", ("#f7971e", "#ffd200"), "🟡")
        self.k_grn = Card("Due after 7 days", CARD_COLORS[1], "🟢")
        self.k_bill = Card("To be Billed", CARD_COLORS[3], "💰")
        for c in (self.k_open, self.k_red, self.k_yel, self.k_grn, self.k_bill):
            cg.addWidget(c)
        self.lay.addLayout(cg)
        bar = QHBoxLayout()
        bar.addWidget(btn("➕ New Task", "green", self.new_task))
        bar.addWidget(QLabel("Show:"))
        self.f_status = QComboBox()
        self.f_status.addItems(["Open", "Completed", "Cancelled", "All"])
        self.f_status.currentIndexChanged.connect(self.refresh)
        bar.addWidget(self.f_status)
        bar.addStretch()
        bar.addWidget(btn("⟳ Refresh", "gray", self.refresh))
        self.lay.addLayout(bar)
        self.tabs = QTabWidget()
        self.tabs.setMinimumHeight(500)
        self.lay.addWidget(self.tabs)
        scopes = list(self.SCOPES) + ([("all", "🗂 All Team Tasks")] if self.admin else [])
        self.task_tables, self.task_rows = {}, {}
        for key, label in scopes:
            w = QWidget()
            v = QVBoxLayout(w)
            t = make_table(["#", "Priority", "Task", "Client", "Assigned To", "Created By", "Due Date", "Days Left",
                            "Status", "Est. Billing", "Billing", "Remarks / Challenges"])
            t.doubleClicked.connect(self.open_detail)
            v.addWidget(t)
            a = QHBoxLayout()
            a.addWidget(btn("📂 Open / Files", None, self.open_detail))
            if self.admin:
                a.addWidget(btn("🔒 Access", "gray", self.access))
            a.addWidget(btn("✎ Edit", "blue", self.edit_task))
            a.addWidget(btn("▶ Start", "orange", lambda: self.set_status("In Progress")))
            a.addWidget(btn("✔ Complete", "green", self.complete))
            a.addWidget(btn("💬 Add Remark", None, self.remark))
            a.addWidget(btn("📨 Send", "blue", self.share))
            a.addWidget(btn("✖ Cancel Task", "gray", lambda: self.set_status("Cancelled")))
            a.addWidget(btn("Delete", "red", self.delete))
            a.addStretch()
            v.addLayout(a)
            self.tabs.addTab(w, label)
            self.task_tables[key] = t
            self.task_rows[key] = []
        # recurring
        w = QWidget()
        v = QVBoxLayout(w)
        self.t_series = make_table(["#", "Task", "Client", "Assigned To", "Frequency", "Due Day", "Next Due Date",
                                    "Created By"], 1)
        v.addWidget(self.t_series)
        a = QHBoxLayout()
        a.addWidget(btn("⏹ Stop Recurring", "red", self.stop_series))
        a.addWidget(QLabel("New tasks are created automatically on the 1st of the month in which they fall due."))
        a.addStretch()
        v.addLayout(a)
        self.tabs.addTab(w, "🔁 Recurring")
        # billing
        w = QWidget()
        v = QVBoxLayout(w)
        f = QHBoxLayout()
        f.addWidget(QLabel("Status:"))
        self.b_status = QComboBox()
        self.b_status.addItems(["All"] + BILL_STATUSES)
        self.b_status.currentIndexChanged.connect(self.load_billing)
        f.addWidget(self.b_status)
        f.addWidget(QLabel("Type:"))
        self.b_type = QComboBox()
        self.b_type.addItems(["All", "Billing Due", "Billing Raised", "Yearly Billing"])
        self.b_type.currentIndexChanged.connect(self.load_billing)
        f.addWidget(self.b_type)
        f.addStretch()
        if self.admin:
            f.addWidget(btn("🧾 Generate Invoice", "orange", self.gen_invoice))
            f.addWidget(btn("📨 Send Invoice", "blue", self.send_bill_invoice))
            f.addWidget(btn("₹ Payment Received", "green", self.bill_payment))
            f.addWidget(btn("✎ Update Status", "green", self.update_bill))
        f.addWidget(btn("Export CSV", "blue", self.bill_csv))
        v.addLayout(f)
        self.t_bill = make_table(["#", "Task", "Client", "Billing Type", "Estimated (Rs.)", "Fee (Rs.)", "Status",
                                  "Invoice / Ref", "Invoice Total", "Received", "Balance Due", "Completed By",
                                  "Completed On"], 1)
        v.addWidget(self.t_bill)
        self.bill_sum = QLabel()
        self.bill_sum.setStyleSheet("background:#f1ecff; border-radius:8px; padding:8px; color:#2a0a7a;")
        v.addWidget(self.bill_sum)
        self.tabs.addTab(w, "💰 Billing")
        self.series, self.bills = [], []
        self.tabs.currentChanged.connect(lambda _: self.refresh())
        legend = QLabel(f"<span style='background:{TASK_RED}'>&nbsp;≤ 3 days / overdue&nbsp;</span> "
                        f"<span style='background:{TASK_YELLOW}'>&nbsp;4–7 days&nbsp;</span> "
                        f"<span style='background:{TASK_GREEN}'>&nbsp;more than 7 days&nbsp;</span> "
                        f"<span style='background:{TASK_GREY}'>&nbsp;completed / cancelled&nbsp;</span> &nbsp; "
                        "🔥 = urgent")
        bar.insertWidget(3, legend)

    # ---------------------------------------------------------------- helpers
    def cur_scope(self):
        keys = list(self.task_tables.keys())
        i = self.tabs.currentIndex()
        return keys[i] if i < len(keys) else None

    def access(self):
        t = self.sel_task()
        if t and AccessDialog(self.main, "task", t["id"], f"Task: {t['title']}"
                                                         + (f" — {t['client']}" if t["client"] else "")).exec_():
            self.refresh()

    def open_detail(self, *a):
        t = self.sel_task()
        if t:
            TaskDetailDialog(self, t["id"]).exec_()
            self.refresh()

    def sel_task(self):
        if getattr(self, "_forced_task", None):
            return self._forced_task
        sc = self.cur_scope()
        if not sc:
            return None
        tid = selected_key(self.task_tables[sc])
        t = next((r for r in self.task_rows[sc] if r["id"] == tid), None)
        if not t:
            err(self, "Please select a task")
        return t

    def refresh(self):
        sc = self.cur_scope()
        try:
            mine = self.client.call("task_list", scope="mine", status="Open")
            bills = self.client.call("billing_list", status="To be Billed")
        except Exception as e:
            return self.main.status(str(e))
        cnt = lambda c: sum(1 for r in mine if r["colour"] == c)
        self.k_open.set(len(mine), f"{sum(r['urgent'] for r in mine)} urgent")
        self.k_red.set(cnt(TASK_RED), "act now")
        self.k_yel.set(cnt(TASK_YELLOW), "plan this week")
        self.k_grn.set(cnt(TASK_GREEN), "on track")
        self.k_bill.set(len(bills), "Rs. " + inr(sum(b["amount"] for b in bills))[:-3])
        if sc:
            self.load_tasks(sc)
        elif self.tabs.currentIndex() == len(self.task_tables):
            self.load_series()
        else:
            self.load_billing()

    def load_tasks(self, sc):
        st = self.f_status.currentText()
        try:
            rows = self.client.call("task_list", scope=sc, status=st)
        except Exception as e:
            return err(self, e)
        self.task_rows[sc] = rows
        t = self.task_tables[sc]
        out = []
        for r in rows:
            dl = r["days_left"]
            left = ("" if r["status"] in ("Completed", "Cancelled") else
                    (f"Overdue {-dl}d" if dl < 0 else ("Today" if dl == 0 else f"{dl} day(s)")))
            bill = (r["billing_option"] + (f" Rs. {inr(r['billing_amount'])}" if r["billing_amount"] else "")
                    if r["billing_option"] else "")
            last = (r["challenges"] or "").strip().split("\n")[-1] if r["challenges"] else ""
            out.append([r["id"], ("🔥 URGENT" if r["urgent"] else r["priority"]) if r["colour"] != TASK_GREY else "—",
                        ("🔒 " if r.get("hide_fees") else "") + r["title"] + (" 🔁" if r["series_id"] else ""), r["client"], r.get("assigned_name") or "",
                        r.get("creator_name") or "", dmy(r["due_date"]), left, r["status"],
                        inr(r["est_amount"]) if r.get("est_amount") else "", bill, last])
        fill_table(t, out, keys=[r["id"] for r in rows])
        for i, r in enumerate(rows):
            for j in range(t.columnCount()):
                t.item(i, j).setBackground(QColor(r["colour"]))
            if r["urgent"] and r["colour"] != TASK_GREY:
                f = t.item(i, 1).font()
                f.setBold(True)
                t.item(i, 1).setFont(f)
                t.item(i, 1).setForeground(QColor("#c62828"))

    def load_series(self):
        try:
            self.series = self.client.call("series_list")
        except Exception as e:
            return err(self, e)
        fill_table(self.t_series, [[s["id"], s["title"], s["client"], s.get("assigned_name"), s["frequency"],
                                    s["day"], dmy(s["next_due"]), s.get("creator_name")] for s in self.series],
                   keys=[s["id"] for s in self.series])

    def load_billing(self):
        st, ty = self.b_status.currentText(), self.b_type.currentText()
        try:
            self.bills = self.client.call("billing_list", status=st if st != "All" else "",
                                          btype=ty if ty != "All" else "")
        except Exception as e:
            return err(self, e)
        hid = lambda b, v: "🔒" if b.get("fee_hidden") else inr(v or 0)
        fill_table(self.t_bill, [[b["id"], b["description"], b["client"], b["btype"], hid(b, b.get("est_amount")),
                                  hid(b, b["amount"]), b["status"], b.get("inv_no") or b["ref"],
                                  inr(b["inv_total"]) if b.get("inv_total") else "",
                                  inr((b.get("inv_received") or 0) + (b.get("inv_tds") or 0)) if b.get("inv_total") else "",
                                  inr(b["inv_total"] - (b.get("inv_received") or 0) - (b.get("inv_tds") or 0))
                                  if b.get("inv_total") and b.get("inv_status") != "Cancelled" else "",
                                  b.get("done_by") or "", (b.get("completed_on") or "")[:10]]
                                 for b in self.bills], keys=[b["id"] for b in self.bills])
        for i in range(self.t_bill.rowCount()):
            for j in (4, 5, 8, 9, 10):
                self.t_bill.item(i, j).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        tot = {}
        for b in self.bills:
            tot[b["status"]] = tot.get(b["status"], 0) + b["amount"]
        self.bill_sum.setText("  |  ".join(f"<b>{k}:</b> Rs. {inr(v)}" for k, v in tot.items()) or "No billing entries")

    # ---------------------------------------------------------------- actions
    def new_task(self):
        dlg = TaskDialog(self, self.main.users, self.me)
        if dlg.exec_():
            v = dlg.value()
            try:
                r = self.client.call("task_save", task=v)
                self.main.status(f"Task #{r['id']} created" + (" (recurring)" if r.get("series_id") else ""))
                self.refresh()
            except Exception as e:
                err(self, e)

    def edit_task(self, *a):
        t = self.sel_task()
        if not t:
            return
        dlg = TaskDialog(self, self.main.users, self.me, t)
        if t["challenges"]:
            dlg.desc.setToolTip(t["challenges"])
        if dlg.exec_():
            try:
                self.client.call("task_save", id=t["id"], task=dlg.value())
                self.refresh()
            except Exception as e:
                err(self, e)

    def set_status(self, st):
        t = self.sel_task()
        if not t:
            return
        if st == "Cancelled" and QMessageBox.question(self, APP_NAME, f"Cancel task '{t['title']}'?") != QMessageBox.Yes:
            return
        try:
            self.client.call("task_status", id=t["id"], status=st)
            self.refresh()
        except Exception as e:
            err(self, e)

    def complete(self):
        t = self.sel_task()
        if not t:
            return
        dlg = CompleteTaskDialog(self, t)
        if dlg.exec_():
            try:
                r = self.client.call("task_complete", id=t["id"], **dlg.value())
                info(self, "Task completed." + ("\nA billing entry has been created in the Billing tab."
                                                if r["billing"] else ""))
                self.refresh()
            except Exception as e:
                err(self, e)

    def remark(self):
        t = self.sel_task()
        if not t:
            return
        dlg = TextDialog(self, "Remarks / Challenges", "<b>History:</b><br>" +
                         html.escape(t["challenges"] or "—").replace("\n", "<br>") +
                         "<br><br><b>Add a new remark (challenge faced, status, client follow-up):</b>")
        if dlg.exec_():
            try:
                self.client.call("task_remark", id=t["id"], text=dlg.t.toPlainText())
                self.refresh()
            except Exception as e:
                err(self, e)

    def delete(self):
        t = self.sel_task()
        if t and QMessageBox.question(self, APP_NAME, f"Delete task '{t['title']}'?") == QMessageBox.Yes:
            try:
                self.client.call("task_delete", id=t["id"])
                self.refresh()
            except Exception as e:
                err(self, e)

    def share(self):
        t = self.sel_task()
        if not t:
            return
        u = next((x for x in self.main.users if x["username"] == t["assigned_to"]), {})
        s = self.main.settings
        due = dmy(t["due_date"])
        info_txt = (f"Task: {t['title']}" + (f"\nClient: {t['client']}" if t["client"] else "") + f"\nDue date: {due}"
                    + ("\nPriority: URGENT" if t["urgent"] else "") + (f"\nDetails: {t['description']}" if t["description"] else ""))
        sig = firm_signature(s, self.main.user["full_name"])
        tpls = [("Task details", f"Task assigned: {t['title']} (due {due})",
                 f"Dear {u.get('full_name', '')},\n\nPlease take up the following task:\n\n{info_txt}\n\n{sig}"),
                ("Due date reminder", f"Reminder: {t['title']} due on {due}",
                 f"Dear {u.get('full_name', '')},\n\nThis is a reminder that the following task is due on {due} "
                 f"({t['days_left']} day(s) left):\n\n{info_txt}\n\nPlease update the status in the app.\n\n{sig}")]
        if t["client"]:
            tpls.append(("To client: documents required", f"Documents required - {t['title']}",
                         f"Dear Sir/Madam,\n\nFor '{t['title']}' due on {due}, kindly share the required "
                         f"information / documents at the earliest.\n\n{sig}"))
        ShareDialog(self.main, "Send task", u.get("full_name", ""), u.get("email", ""), u.get("phone", ""), tpls).exec_()

    def _sel_bill(self):
        bid = selected_key(self.t_bill)
        b = next((x for x in self.bills if x["id"] == bid), None)
        if not b:
            err(self, "Please select a billing entry")
        return b

    def gen_invoice(self):
        b = self._sel_bill()
        if not b:
            return
        if b.get("invoice_id") and b.get("inv_status") != "Cancelled":
            return err(self, f"Invoice {b.get('inv_no')} is already generated for this entry")
        if self.main.generate_invoice(b):
            self.load_billing()

    def bill_payment(self):
        b = self._sel_bill()
        if not b:
            return
        if not b.get("invoice_id") or b.get("inv_status") == "Cancelled":
            return err(self, "Generate the invoice first")
        try:
            inv = next(x for x in self.client.call("invoice_list") if x["id"] == b["invoice_id"])
        except Exception as e:
            return err(self, e)
        if self.main.pages["lit"].record_receipt(inv):
            self.load_billing()

    def send_bill_invoice(self):
        b = self._sel_bill()
        if not b:
            return
        if not b.get("invoice_id"):
            return err(self, "Generate the invoice first")
        self.main.send_invoice(b["invoice_id"])

    def stop_series(self):
        sid = selected_key(self.t_series)
        if not sid:
            return err(self, "Please select a recurring task")
        if QMessageBox.question(self, APP_NAME, "Stop creating this task automatically? (Existing tasks stay.)") \
                == QMessageBox.Yes:
            try:
                self.client.call("series_stop", id=sid)
                self.load_series()
            except Exception as e:
                err(self, e)

    def update_bill(self):
        bid = selected_key(self.t_bill)
        b = next((x for x in self.bills if x["id"] == bid), None)
        if not b:
            return err(self, "Please select a billing entry")
        dlg = BillUpdateDialog(self, b)
        if dlg.exec_():
            try:
                self.client.call("billing_update", id=bid, status=dlg.st.currentText(), ref=dlg.ref.text().strip(),
                                 amount=dlg.amt.value())
                self.refresh()
            except Exception as e:
                err(self, e)

    def bill_csv(self):
        if not self.bills:
            return err(self, "Nothing to export")
        os.makedirs(REPORT_DIR, exist_ok=True)
        p, _ = QFileDialog.getSaveFileName(self, "Save CSV", os.path.join(REPORT_DIR, "Task_Billing.csv"),
                                           "CSV (*.csv)")
        if not p:
            return
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["#", "Task", "Client", "Type", "Amount", "Status", "Ref", "Completed By", "Completed On"])
            for b in self.bills:
                w.writerow([b["id"], b["description"], b["client"], b["btype"], b["amount"], b["status"], b["ref"],
                            b.get("done_by"), b.get("completed_on")])
        open_path(p)


# =============================================================================
# 16C. GUI : SHARING (E-MAIL / WHATSAPP), INVOICING, CLIENT & FIRM MASTER
# =============================================================================
def norm_phone(p):
    """10-digit Indian mobile -> 91XXXXXXXXXX (format needed by wa.me links)."""
    d = re.sub(r"\D", "", str(p or ""))
    if len(d) == 11 and d.startswith("0"):
        d = d[1:]
    if len(d) == 10:
        d = "91" + d
    return d


def open_whatsapp(phone, text):
    from urllib.parse import quote
    ph = norm_phone(phone)
    QDesktopServices.openUrl(QUrl(f"https://wa.me/{ph}?text={quote(text)}" if ph
                                  else f"https://wa.me/?text={quote(text)}"))


def open_mail_app(email, subject, body):
    from urllib.parse import quote
    QDesktopServices.openUrl(QUrl(f"mailto:{email}?subject={quote(subject)}&body={quote(body)}"))


def firm_signature(s, user_name=""):
    lines = ["Regards,", user_name or s.get("signatory") or "", s.get("company_name") or FIRM_NAME]
    if s.get("firm_phone"):
        lines.append("Ph: " + s["firm_phone"])
    if s.get("firm_email"):
        lines.append(s["firm_email"])
    return "\n".join(x for x in lines if x)


def invoice_message(inv, s, reminder=False, user_name=""):
    out = round(inv["total"] - inv["received"] - inv["tds"], 2)
    age = (date.today() - date.fromisoformat(inv["inv_date"])).days
    if reminder:
        subject = f"Payment reminder - Invoice {inv['inv_no']} - Rs. {inr(out)} outstanding"
        body = (f"Dear {inv.get('client_name') or 'Sir/Madam'},\n\nGreetings from {s.get('company_name') or FIRM_NAME}."
                f"\n\nThis is a gentle reminder that our Invoice No. {inv['inv_no']} dated {dmy(inv['inv_date'])} "
                f"for Rs. {inr(inv['total'])} is outstanding for {age} days.\nBalance due: Rs. {inr(out)}"
                "\n\nKindly arrange the payment at the earliest. Please ignore if already paid.")
    else:
        subject = f"Invoice {inv['inv_no']} - {s.get('company_name') or FIRM_NAME}"
        body = (f"Dear {inv.get('client_name') or 'Sir/Madam'},\n\nGreetings from {s.get('company_name') or FIRM_NAME}."
                f"\n\nPlease find attached our Invoice No. {inv['inv_no']} dated {dmy(inv['inv_date'])} for "
                f"Rs. {inr(inv['total'])} ({amount_in_words(inv['total'])}).")
    bank = firm_bank_text(s)
    if bank:
        body += f"\n\nPayment details: {bank}"
    return subject, body + "\n\n" + firm_signature(s, user_name)


def reminder_templates(inv, s, user_name=""):
    """Escalating payment reminders: gentle -> follow-up -> final; plus acknowledgement of promise."""
    out = round(inv["total"] - inv["received"] - inv["tds"], 2)
    age = (date.today() - date.fromisoformat(inv["inv_date"])).days
    firm = s.get("company_name") or FIRM_NAME
    who = inv.get("client_name") or "Sir/Madam"
    ref = f"Invoice No. {inv['inv_no']} dated {dmy(inv['inv_date'])} for Rs. {inr(inv['total'])}"
    bank = firm_bank_text(s)
    pay = f"\n\nPayment details: {bank}" if bank else ""
    sig = "\n\n" + firm_signature(s, user_name)
    prm = inv.get("fu_promise")
    t1 = (f"Payment reminder - Invoice {inv['inv_no']}",
          f"Dear {who},\n\nGreetings from {firm}.\n\nThis is a gentle reminder that our {ref} is outstanding for "
          f"{age} days. Balance due: Rs. {inr(out)}.\n\nKindly arrange the payment at your earliest convenience. "
          "Please ignore this message if already paid." + pay + sig)
    t2 = (f"Follow-up: payment pending - Invoice {inv['inv_no']} - Rs. {inr(out)}",
          f"Dear {who},\n\nWe refer to our earlier reminder regarding our {ref}.\n\nThe balance of Rs. {inr(out)} "
          f"is still pending for {age} days"
          + (f", although payment was expected by {dmy(prm)}" if prm else "") +
          ". We request you to kindly release the payment within the next 7 days and share the payment details "
          "(UTR / cheque no.) with us." + pay + sig)
    t3 = (f"FINAL REMINDER - Outstanding Rs. {inr(out)} - Invoice {inv['inv_no']}",
          f"Dear {who},\n\nDespite our earlier reminders, our {ref} remains unpaid for {age} days. Balance due: "
          f"Rs. {inr(out)}.\n\nWe request you to clear the outstanding amount immediately. In case of any query on "
          "the invoice, please contact us at the earliest so that it can be resolved." + pay + sig)
    t4 = (f"Thank you - payment confirmation awaited - Invoice {inv['inv_no']}",
          f"Dear {who},\n\nThank you for your response. As discussed, we look forward to receiving the balance of "
          f"Rs. {inr(out)} against our {ref}" + (f" by {dmy(prm)}" if prm else "") + "." + pay + sig)
    return [("1st reminder (gentle)",) + t1, ("2nd follow-up (firm)",) + t2, ("Final reminder",) + t3,
            ("Acknowledge promise to pay",) + t4]


class ShareDialog(QDialog):
    """Send any message by E-mail (SMTP, with PDF) or WhatsApp (wa.me link) or the default mail app.
    templates: list of (label, subject, body)."""

    def __init__(self, main, title, name, email, phone, templates, attachment=None, start=0):
        super().__init__(main)
        self.main, self.attachment, self.templates = main, attachment, templates
        self.sent = None                     # (mode, template label) once something is sent
        self.setWindowTitle(f"{APP_NAME} - {title}")
        self.setMinimumWidth(620)
        f = QFormLayout(self)
        hdr = QLabel(f"<div style='font-size:13pt;font-weight:800;color:#4a00e0'>📨 {html.escape(title)}</div>"
                     f"<br>To: <b>{html.escape(name or '')}</b>")
        f.addRow(hdr)
        self.tpl = QComboBox()
        for lbl, _, _ in templates:
            self.tpl.addItem(lbl)
        self.tpl.currentIndexChanged.connect(self._load)
        if len(templates) > 1:
            f.addRow("Message type:", self.tpl)
        self.email = QLineEdit(email or "")
        self.email.setPlaceholderText("e-mail address")
        self.phone = QLineEdit(phone or "")
        self.phone.setPlaceholderText("mobile number for WhatsApp (10 digits)")
        self.subject = QLineEdit()
        self.body = QTextEdit()
        self.body.setMinimumHeight(220)
        f.addRow("E-mail:", self.email)
        f.addRow("WhatsApp no.:", self.phone)
        f.addRow("Subject:", self.subject)
        f.addRow("Message:", self.body)
        if attachment:
            al = QLabel(f"📎 <b>{html.escape(os.path.basename(attachment))}</b> — attached automatically to the "
                        "e-mail. For WhatsApp, the folder opens so you can drag the PDF into the chat.")
            al.setWordWrap(True)
            al.setStyleSheet("color:#6b6f8e;")
            al.setMinimumHeight(40)
            f.addRow(al)
        h = QHBoxLayout()
        h.addWidget(btn("✉ Send E-mail", "blue", self.send_mail))
        h.addWidget(btn("🟢 WhatsApp", "green", self.send_wa))
        h.addWidget(btn("📧 Open in Mail App", "orange", self.mail_app))
        h.addWidget(btn("📋 Copy", "gray", self.copy))
        h.addWidget(btn("Close", "gray", self.reject))
        f.addRow(h)
        self._load(0)
        if 0 < start < len(templates):
            self.tpl.setCurrentIndex(start)

    def _label(self):
        return self.templates[self.tpl.currentIndex()][0]

    def _load(self, i):
        _, sub, body = self.templates[i]
        self.subject.setText(sub)
        self.body.setPlainText(body)

    def send_mail(self):
        dlg = EmailDialog(self, self.main.cfg, self.email.text().strip(), self.subject.text(),
                          self.body.toPlainText(), self.attachment)
        if dlg.exec_():
            self.sent = ("E-mail", self._label())

    def send_wa(self):
        if not norm_phone(self.phone.text()):
            if QMessageBox.question(self, APP_NAME, "No WhatsApp number entered. Open WhatsApp and choose the "
                                                    "contact yourself?") != QMessageBox.Yes:
                return
        open_whatsapp(self.phone.text(), self.subject.text() + "\n\n" + self.body.toPlainText())
        self.sent = ("WhatsApp", self._label())
        if self.attachment:
            open_path(os.path.dirname(self.attachment))
        self.main.status("WhatsApp opened with the message" + (" — attach the PDF from the opened folder"
                                                              if self.attachment else ""))

    def mail_app(self):
        open_mail_app(self.email.text().strip(), self.subject.text(), self.body.toPlainText())
        self.sent = ("E-mail", self._label())
        if self.attachment:
            open_path(os.path.dirname(self.attachment))

    def copy(self):
        QApplication.clipboard().setText(self.subject.text() + "\n\n" + self.body.toPlainText())
        self.main.status("Message copied to clipboard")


class ClientDialog(QDialog):
    def __init__(self, parent, c=None):
        super().__init__(parent)
        c = c or {}
        self.cid = c.get("id")
        self.setWindowTitle(f"{APP_NAME} - {'Edit' if self.cid else 'New'} Client")
        self.setMinimumWidth(500)
        f = QFormLayout(self)
        self.w = {}
        for k, lbl, ph in (("name", "Client name*", ""), ("contact", "Contact person", ""),
                           ("gstin", "GSTIN", "15 characters"), ("pan", "PAN", ""), ("address", "Address", ""),
                           ("email", "E-mail", ""), ("phone", "Mobile / WhatsApp", "10-digit mobile")):
            e = QLineEdit(str(c.get(k) or ""))
            e.setPlaceholderText(ph)
            self.w[k] = e
            f.addRow(lbl + ":", e)
        self.state = QLabel(state_of(c.get("gstin")))
        self.state.setStyleSheet("color:#4a00e0;")
        self.w["gstin"].textChanged.connect(lambda t: self.state.setText(state_of(t.upper())))
        f.addRow("State (from GSTIN):", self.state)
        h = QHBoxLayout()
        h.addWidget(btn("Save", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)

    def value(self):
        d = {k: e.text().strip() for k, e in self.w.items()}
        d["gstin"], d["pan"] = d["gstin"].upper(), d["pan"].upper()
        if self.cid:
            d["id"] = self.cid
        return d


class ManualInvoiceDialog(QDialog):
    """System-generated GST invoice (task billing entry or any professional service)."""

    def __init__(self, main, clients, settings, billing=None):
        super().__init__(main)
        self.main, self.clients, self.settings, self.billing = main, clients, settings, billing
        self.setWindowTitle(f"{APP_NAME} - Generate Tax Invoice")
        self.setMinimumSize(820, 830)
        v = QVBoxLayout(self)
        if billing:
            est = billing.get("est_amount") or 0
            lb = QLabel(f"<b>Billing entry #{billing['id']}:</b> {html.escape(billing['description'] or '')} — "
                        f"{html.escape(billing['client'] or '')} &nbsp; | &nbsp; Billing type: {billing['btype']}"
                        f" &nbsp; | &nbsp; <b>Estimated: Rs. {inr(est)}</b> &nbsp; | &nbsp; "
                        f"Amount on completion: Rs. {inr(billing['amount'])}")
            lb.setWordWrap(True)
            lb.setStyleSheet("background:#f1ecff; border-radius:8px; padding:8px; color:#2a0a7a;")
            v.addWidget(lb)
        g = QGroupBox("Bill To")
        f = QFormLayout(g)
        self.cname = QComboBox()
        self.cname.setEditable(True)
        self.cname.addItem("")
        for c in clients:
            self.cname.addItem(c["name"], c["id"])
        self.cname.currentTextChanged.connect(self._fill_client)
        self.cw = {k: QLineEdit() for k in ("gstin", "pan", "address", "email", "phone")}
        f.addRow("Client*:", self.cname)
        r1 = QHBoxLayout()
        r1.addWidget(self.cw["gstin"])
        r1.addWidget(self.cw["pan"])
        self.cw["gstin"].setPlaceholderText("GSTIN")
        self.cw["pan"].setPlaceholderText("PAN")
        f.addRow("GSTIN / PAN:", r1)
        f.addRow("Address:", self.cw["address"])
        r2 = QHBoxLayout()
        r2.addWidget(self.cw["email"])
        r2.addWidget(self.cw["phone"])
        self.cw["email"].setPlaceholderText("E-mail")
        self.cw["phone"].setPlaceholderText("Mobile / WhatsApp")
        f.addRow("E-mail / Mobile:", r2)
        self.save_client = QCheckBox("Save / update this client in the Client Master")
        self.save_client.setChecked(True)
        f.addRow("", self.save_client)
        v.addWidget(g)
        g2 = QGroupBox("Invoice")
        f2 = QFormLayout(g2)
        self.d = date_edit()
        self.mode = QComboBox()
        self.mode.addItems(["CGST+SGST", "IGST"])
        self.rate = QDoubleSpinBox()
        self.rate.setRange(0, 40)
        self.rate.setSuffix(" %")
        self.rate.setValue(float(settings.get("gst_rate") or 18))
        self.gst_type = settings.get("gst_type") or "Regular"
        if self.gst_type != "Regular":
            self.mode.addItem("No GST")
            self.mode.setCurrentText("No GST")
            self.mode.setEnabled(False)
            self.rate.setValue(0)
            self.rate.setEnabled(False)
            self.mode.setToolTip(f"Firm is {self.gst_type} under GST (Firm Master) — no GST is charged")
        self.notes = QLineEdit(billing["description"] if billing else "")
        self.notes.setPlaceholderText("Particulars shown on invoice, e.g. Tax audit FY 2025-26")
        r3 = QHBoxLayout()
        r3.addWidget(QLabel("Date:"))
        r3.addWidget(self.d)
        r3.addWidget(QLabel("GST:"))
        r3.addWidget(self.mode)
        r3.addWidget(self.rate)
        r3.addStretch()
        f2.addRow(r3)
        if self.gst_type != "Regular":
            gn = QLabel(f"ℹ Firm is <b>{self.gst_type}</b> under GST — no GST will be charged "
                        f"({GST_DOC_TITLE[self.gst_type].title()}).")
            gn.setStyleSheet("background:#ecfeff; border-radius:8px; padding:6px; color:#155e75;")
            f2.addRow(gn)
        f2.addRow("Particulars:", self.notes)
        self.sdet = QTextEdit()
        self.sdet.setPlaceholderText("Additional details of service (optional) — scope of work, period covered, "
                                     "returns filed, meetings attended ...")
        self.sdet.setMaximumHeight(70)
        f2.addRow("Service details:", self.sdet)
        self.lines = QTableWidget(0, 4)
        self.lines.setHorizontalHeaderLabels(["Description", "SAC", "Amount (Rs.)", "GST applicable"])
        self.lines.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.lines.verticalHeader().setVisible(False)
        self.lines.setMinimumHeight(190)
        self.lines.setColumnWidth(1, 80)
        self.lines.setColumnWidth(2, 120)
        self.lines.setColumnWidth(3, 150)
        self.lines.itemChanged.connect(lambda *_: self._calc())
        f2.addRow(self.lines)
        lh = QHBoxLayout()
        lh.addWidget(btn("➕ Add Line", "blue", lambda: self._add_line()))
        lh.addWidget(btn("➕ Add Reimbursement (no GST)", "orange", lambda: self._add_line("", 0, False)))
        lh.addWidget(btn("Remove Line", "red", self._del_line))
        lh.addStretch()
        f2.addRow(lh)
        self.prev = QLabel()
        self.prev.setStyleSheet("color:#4a00e0; font-weight:700; font-size:11pt;")
        f2.addRow(self.prev)
        v.addWidget(g2)
        h = QHBoxLayout()
        h.addWidget(btn("🧾 Generate Invoice", "green", self._ok))
        h.addWidget(btn("Cancel", "gray", self.reject))
        v.addLayout(h)
        for w in (self.mode.currentTextChanged, self.rate.valueChanged):
            w.connect(lambda *_: self._calc())
        self.cw["gstin"].textChanged.connect(self._auto_mode)
        if billing:
            self.cname.setCurrentText(billing["client"] or "")
            amt = billing["amount"] or billing.get("est_amount") or 0
            self._add_line(f"Professional fees - {billing['description']}", amt, True)
        else:
            self._add_line()
        self._calc()

    def _fill_client(self, name):
        c = next((x for x in self.clients if x["name"].lower() == name.strip().lower()), None)
        if c:
            for k, e in self.cw.items():
                e.setText(c.get(k) or "")

    def _auto_mode(self, g):
        if getattr(self, "gst_type", "Regular") != "Regular":
            return
        fs, cs = (self.settings.get("firm_gstin") or "")[:2], (g or "")[:2]
        if fs and len(cs) == 2:
            self.mode.setCurrentText("IGST" if fs != cs else "CGST+SGST")

    def _add_line(self, desc="", amt=0.0, taxable=True):
        self.lines.blockSignals(True)
        r = self.lines.rowCount()
        self.lines.insertRow(r)
        self.lines.setItem(r, 0, QTableWidgetItem(desc))
        self.lines.setItem(r, 1, QTableWidgetItem(self.settings.get("sac_code") or "9982"))
        it = QTableWidgetItem(f"{amt:.2f}")
        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.lines.setItem(r, 2, it)
        ck = QTableWidgetItem("Yes" if taxable else "No (reimbursement)")
        ck.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
        ck.setCheckState(Qt.Checked if taxable else Qt.Unchecked)
        self.lines.setItem(r, 3, ck)
        self.lines.blockSignals(False)
        self._calc()

    def _del_line(self):
        r = self.lines.currentRow()
        if r >= 0:
            self.lines.removeRow(r)
            self._calc()

    def line_values(self):
        out = []
        for r in range(self.lines.rowCount()):
            g = lambda c: (self.lines.item(r, c).text() if self.lines.item(r, c) else "")
            try:
                amt = float(g(2).replace(",", "") or 0)
            except ValueError:
                amt = 0.0
            out.append({"description": g(0).strip(), "sac": g(1).strip(), "amount": round(amt, 2),
                        "taxable": self.lines.item(r, 3).checkState() == Qt.Checked if self.lines.item(r, 3) else True})
        return out

    def _calc(self):
        ls = self.line_values()
        tx = sum(x["amount"] for x in ls if x["taxable"])
        ot = sum(x["amount"] for x in ls if not x["taxable"])
        gst = round(tx * self.rate.value() / 100, 2)
        self.prev.setText(f"{'Taxable' if self.gst_type == 'Regular' else 'Fees'} Rs. {inr(tx)} + GST Rs. {inr(gst)}"
                          f" + Reimbursement Rs. {inr(ot)} = "
                          f"Total Rs. {inr(tx + gst + ot)}")

    def _ok(self):
        if not self.cname.currentText().strip():
            return err(self, "Select or type the client name")
        if not any(x["description"] and x["amount"] > 0 for x in self.line_values()):
            return err(self, "Add at least one line with description and amount")
        self.accept()

    def value(self):
        cl = {k: e.text().strip() for k, e in self.cw.items()}
        cl["name"] = self.cname.currentText().strip()
        return dict(client=cl, inv_date=self.d.date().toPyDate().isoformat(), gst_mode=self.mode.currentText(),
                    gst_rate=self.rate.value(), lines=self.line_values(), notes=self.notes.text().strip(),
                    service_details=self.sdet.toPlainText().strip(),
                    save_client=self.save_client.isChecked(),
                    billing_id=self.billing["id"] if self.billing else None)


class FirmMasterWidget(QWidget):
    FIELDS = [("company_name", "Firm name*"), ("firm_tagline", "Tagline (under firm name)"),
              ("firm_frn", "Firm Registration No. (FRN)"),
              ("company_address", "Address"), ("firm_city", "City / PIN"), ("firm_state", "State"),
              ("firm_phone", "Phone / Mobile"), ("firm_email", "E-mail"), ("firm_website", "Website"),
              ("firm_gstin", "GSTIN"), ("firm_pan", "PAN"), ("bank_name", "Bank name"), ("bank_acc", "Account no."),
              ("bank_ifsc", "IFSC"), ("bank_branch", "Branch"), ("upi_id", "UPI ID"),
              ("invoice_prefix", "Invoice prefix"), ("sac_code", "Default SAC code"),
              ("signatory", "Signatory (name / designation)"), ("manager_email", "Manager e-mail (claims)")]

    def __init__(self, main):
        super().__init__()
        self.main, self.client = main, main.client
        self.logo_b64 = ""
        h = QHBoxLayout(self)
        left = QGroupBox("Firm Master — used on invoices, PDFs, e-mail && WhatsApp")
        f = QFormLayout(left)
        self.w = {}
        for k, lbl in self.FIELDS:
            e = QLineEdit()
            self.w[k] = e
            f.addRow(lbl + ":", e)
        self.w["firm_state"].setPlaceholderText("e.g. Rajasthan (auto from GSTIN)")
        self.w["firm_gstin"].textChanged.connect(
            lambda t: self.w["firm_state"].setText(STATE_CODES.get(t[:2], self.w["firm_state"].text())))
        self.w["invoice_prefix"].setPlaceholderText("e.g. SS  → SS/2026-27/0001")
        self.rate = QDoubleSpinBox()
        self.rate.setRange(0, 40)
        self.rate.setSuffix(" %")
        self.days = QSpinBox()
        self.days.setRange(0, 30)
        self.days.setSuffix(" day(s) before")
        self.terms = QTextEdit()
        self.terms.setMaximumHeight(70)
        self.gst_type = QComboBox()
        self.gst_type.addItem("Regular — charge GST (CGST 9% + SGST 9% / IGST 18%)", "Regular")
        self.gst_type.addItem("Composition scheme — no GST on invoice (Bill of Supply)", "Composition")
        self.gst_type.addItem("Unregistered — no GST on invoice", "Unregistered")
        self.gst_type.currentIndexChanged.connect(lambda _: self.rate.setEnabled(self.gst_type.currentData() == "Regular"))
        f.addRow("GST registration:", self.gst_type)
        f.addRow("Default GST rate:", self.rate)
        f.addRow("Hearing reminder:", self.days)
        f.addRow("Invoice terms:", self.terms)
        f.addRow(btn("💾 Save Firm Master", "green", self.save))
        h.addWidget(left, 3)
        right = QGroupBox("Firm Logo")
        rv = QVBoxLayout(right)
        self.logo = QLabel("No logo")
        self.logo.setAlignment(Qt.AlignCenter)
        self.logo.setMinimumSize(190, 130)
        self.logo.setStyleSheet("background:#faf8ff; border:2px dashed #c9b8ff; border-radius:10px;")
        rv.addWidget(self.logo)
        rv.addWidget(btn("Upload Logo (PNG / JPG)", "blue", self.upload))
        rv.addWidget(btn("Remove Logo", "red", self.remove))
        note = QLabel("Logo is printed on the top-left of every invoice. Use a small image (under 500 KB).")
        note.setWordWrap(True)
        note.setStyleSheet("color:#6b6f8e; font-size:8.5pt;")
        rv.addWidget(note)
        sg = QGroupBox("Authorised Signature")
        sv = QVBoxLayout(sg)
        self.sign = QLabel("No signature")
        self.sign.setAlignment(Qt.AlignCenter)
        self.sign.setMinimumSize(190, 80)
        self.sign.setStyleSheet("background:white; border:2px dashed #fdba74; border-radius:10px;")
        sv.addWidget(self.sign)
        sv.addWidget(btn("Upload Signature (PNG / JPG)", "blue", self.upload_sign))
        sv.addWidget(btn("Remove Signature", "red", self.remove_sign))
        self.sign_auto = QCheckBox("Sign invoices && receipts automatically")
        self.sign_auto.setChecked(True)
        sv.addWidget(self.sign_auto)
        sn = QLabel("Tip: sign on white paper, scan / photograph it and crop tightly. A PNG with transparent "
                    "background looks best. Keep it under 500 KB.")
        sn.setWordWrap(True)
        sn.setStyleSheet("color:#6b6f8e; font-size:8.5pt;")
        sv.addWidget(sn)
        rv.addWidget(sg)
        rv.addStretch()
        h.addWidget(right, 1)
        self.sign_b64 = ""

    def load(self):
        try:
            s = self.client.call("firm_get")
        except Exception as e:
            return err(self, e)
        for k, e in self.w.items():
            e.setText(s.get(k, "") or "")
        self.rate.setValue(float(s.get("gst_rate") or 18))
        self.gst_type.setCurrentIndex(max(0, self.gst_type.findData(s.get("gst_type") or "Regular")))
        self.rate.setEnabled(self.gst_type.currentData() == "Regular")
        self.days.setValue(int(float(s.get("reminder_days") or 3)))
        self.terms.setPlainText(s.get("terms", ""))
        self.logo_b64 = s.get("firm_logo", "") or ""
        self._show_logo()
        self.sign_b64 = s.get("firm_sign", "") or ""
        self.sign_auto.setChecked(s.get("sign_auto", "1") == "1")
        self._show_sign()

    def _show_sign(self):
        import base64
        if self.sign_b64:
            pm = QPixmap()
            pm.loadFromData(base64.b64decode(self.sign_b64))
            self.sign.setPixmap(pm.scaled(220, 90, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            self.sign.setPixmap(QPixmap())
            self.sign.setText("No signature")

    def upload_sign(self):
        import base64
        p, _ = QFileDialog.getOpenFileName(self, "Select signature image", "", "Images (*.png *.jpg *.jpeg)")
        if not p:
            return
        if os.path.getsize(p) > 500 * 1024:
            return err(self, "Signature image must be under 500 KB")
        with open(p, "rb") as fh:
            self.sign_b64 = base64.b64encode(fh.read()).decode()
        self._show_sign()

    def remove_sign(self):
        self.sign_b64 = ""
        self._show_sign()

    def _show_logo(self):
        import base64
        if self.logo_b64:
            pm = QPixmap()
            pm.loadFromData(base64.b64decode(self.logo_b64))
            self.logo.setPixmap(pm.scaled(220, 140, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            self.logo.setPixmap(QPixmap())
            self.logo.setText("No logo")

    def upload(self):
        import base64
        p, _ = QFileDialog.getOpenFileName(self, "Select logo", "", "Images (*.png *.jpg *.jpeg)")
        if not p:
            return
        if os.path.getsize(p) > 500 * 1024:
            return err(self, "Logo must be under 500 KB")
        with open(p, "rb") as fh:
            self.logo_b64 = base64.b64encode(fh.read()).decode()
        self._show_logo()

    def remove(self):
        self.logo_b64 = ""
        self._show_logo()

    def save(self):
        if not self.w["company_name"].text().strip():
            return err(self, "Firm name is required")
        g = self.w["firm_gstin"].text().strip().upper()
        if g and not re.fullmatch(r"\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]", g):
            if QMessageBox.question(self, APP_NAME, "GSTIN format looks incorrect. Save anyway?") != QMessageBox.Yes:
                return
        vals = {k: e.text().strip() for k, e in self.w.items()}
        vals["firm_gstin"], vals["firm_pan"] = g, vals["firm_pan"].upper()
        if self.gst_type.currentData() == "Regular" and not g:
            if QMessageBox.question(self, APP_NAME, "GST registration is 'Regular' but GSTIN is blank. Save anyway?") \
                    != QMessageBox.Yes:
                return
        vals.update(gst_type=self.gst_type.currentData())
        vals.update(gst_rate=self.rate.value(), reminder_days=self.days.value(),
                    terms=self.terms.toPlainText().strip(), firm_logo=self.logo_b64, firm_sign=self.sign_b64,
                    sign_auto="1" if self.sign_auto.isChecked() else "0")
        try:
            self.main.settings = self.client.call("settings_set", **vals)
            PDF_BRAND["logo"] = self.logo_b64
            self.main.apply_firm_name(vals["company_name"])
            info(self, "Firm Master saved. It will be used on all new invoices and PDFs.")
        except Exception as e:
            err(self, e)


class FollowupDialog(QDialog):
    """Payment follow-up register for one invoice: send reminders and log calls / visits / promises."""

    def __init__(self, main, iid):
        super().__init__(main)
        self.main, self.client, self.iid = main, main.client, iid
        self.setWindowTitle(f"{APP_NAME} - Payment follow-up")
        self.setWindowIcon(app_icon())
        self.resize(980, 640)
        v = QVBoxLayout(self)
        self.head = QLabel()
        self.head.setWordWrap(True)
        v.addWidget(self.head)
        h = QHBoxLayout()
        h.addWidget(btn("📨 Send Reminder (E-mail / WhatsApp)", "green", self.send))
        h.addWidget(btn("₹ Update Payment", "orange", self.payment))
        h.addWidget(btn("📄 Invoice PDF", "blue", lambda: self.main.invoice_pdf(self.iid)))
        h.addStretch()
        v.addLayout(h)
        g = QGroupBox("Record a follow-up (call / visit / message)")
        gf = QGridLayout(g)
        self.fdate = date_edit()
        self.mode = QComboBox()
        self.mode.addItems(FU_MODES)
        self.remarks = QLineEdit()
        self.remarks.setPlaceholderText("Client's response, e.g. 'Accounts head will release by Friday'")
        self.pcb = QCheckBox("Payment promised on:")
        self.pdate = date_edit(QDate.currentDate().addDays(3))
        self.pdate.setEnabled(False)
        self.pcb.toggled.connect(self.pdate.setEnabled)
        self.pcb.toggled.connect(lambda on: on and self.ndate.setDate(self.pdate.date().addDays(1)))
        self.ndate = date_edit(QDate.currentDate().addDays(FU_GAP_DAYS))
        gf.addWidget(QLabel("Date:"), 0, 0)
        gf.addWidget(self.fdate, 0, 1)
        gf.addWidget(QLabel("Mode:"), 0, 2)
        gf.addWidget(self.mode, 0, 3)
        gf.addWidget(QLabel("Response / remarks:"), 1, 0)
        gf.addWidget(self.remarks, 1, 1, 1, 3)
        gf.addWidget(self.pcb, 2, 0)
        gf.addWidget(self.pdate, 2, 1)
        gf.addWidget(QLabel("Next follow-up:"), 2, 2)
        gf.addWidget(self.ndate, 2, 3)
        gf.addWidget(btn("💾 Save Follow-up", "green", self.save), 3, 3)
        v.addWidget(g)
        lb = QLabel("Follow-up history")
        lb.setStyleSheet("font-weight:800; color:#6d28d9; font-size:11pt;")
        v.addWidget(lb)
        self.t = make_table(["Date", "Mode", "Response / Remarks", "Promised Date", "Next Follow-up", "By"], 2)
        v.addWidget(self.t, 1)
        hb = QHBoxLayout()
        hb.addWidget(btn("Delete Entry", "red", self.delete))
        hb.addStretch()
        hb.addWidget(btn("Close", "gray", self.close))
        v.addLayout(hb)
        self.inv = None
        self.refresh()

    def refresh(self):
        try:
            self.inv = next(x for x in self.client.call("invoice_list") if x["id"] == self.iid)
            rows = self.client.call("followup_list", invoice_id=self.iid)
        except Exception as e:
            self.head.setText(html.escape(str(e)))
            return
        i = self.inv
        due = i["fu_next"] and i["fu_next"] <= date.today().isoformat()
        self.head.setText(
            f"<div style='font-size:14pt;font-weight:800;color:#6d28d9'>📞 {html.escape(i['client_name'])} — "
            f"{html.escape(i['inv_no'])}</div><br>Invoice date: {dmy(i['inv_date'])} &nbsp;|&nbsp; Total: Rs. {inr(i['total'])}"
            f" &nbsp;|&nbsp; Received: Rs. {inr(i['received'] + i['tds'])} &nbsp;|&nbsp; <b style='color:#c2410c'>"
            f"Balance due: Rs. {inr(i['outstanding'])}</b> &nbsp;|&nbsp; Age: {i['age_days']} days<br>"
            f"Follow-ups done: <b>{i['fu_count']}</b>"
            + (f" &nbsp;|&nbsp; Last: {dmy(i['fu_last'])}" if i["fu_last"] else "")
            + (f" &nbsp;|&nbsp; Payment promised: <b>{dmy(i['fu_promise'])}</b>" if i["fu_promise"] else "")
            + (f" &nbsp;|&nbsp; Next follow-up: <b style='color:{'#dc2626' if due else '#059669'}'>{dmy(i['fu_next'])}"
               f"{' (DUE)' if due else ''}</b>" if i["fu_next"] else "")
            + (" &nbsp;<b style='color:#059669'>✔ Fully paid</b>" if i["outstanding"] <= 0.009 else ""))
        fill_table(self.t, [[dmy(r["fdate"]), r["mode"], r["remarks"], dmy(r["promise_date"]) if r["promise_date"] else "",
                             dmy(r["next_date"]) if r["next_date"] else "", r.get("by_name") or r["by_user"]]
                            for r in rows], keys=[r["id"] for r in rows])

    def send(self):
        sent = self.main.send_payment_reminder(self.iid)
        if sent:
            self.refresh()

    def payment(self):
        if self.main.pages["lit"].record_receipt(self.inv):
            self.refresh()

    def save(self):
        try:
            self.client.call("followup_add", invoice_id=self.iid, fdate=self.fdate.date().toPyDate().isoformat(),
                             mode=self.mode.currentText(), remarks=self.remarks.text().strip(),
                             promise_date=self.pdate.date().toPyDate().isoformat() if self.pcb.isChecked() else "",
                             next_date=self.ndate.date().toPyDate().isoformat())
            self.remarks.clear()
            self.pcb.setChecked(False)
            self.refresh()
            self.main.status("Follow-up saved")
        except Exception as e:
            err(self, e)

    def delete(self):
        fid = selected_key(self.t)
        if fid and QMessageBox.question(self, APP_NAME, "Delete this follow-up entry?") == QMessageBox.Yes:
            try:
                self.client.call("followup_delete", id=fid)
                self.refresh()
            except Exception as e:
                err(self, e)


class OverdueInvoicesDialog(QDialog):
    def __init__(self, main, rows, heading=None):
        super().__init__(main)
        self.main, self.rows = main, rows
        self.setWindowTitle(f"{APP_NAME} - " + ("Payment follow-ups due" if heading else
                                               f"Invoices outstanding > {INVOICE_OVERDUE_DAYS} days"))
        self.setWindowIcon(app_icon())
        self.resize(940, 460)
        v = QVBoxLayout(self)
        tot = sum(r["outstanding"] for r in rows)
        hd = QLabel("<div style='font-size:14pt;font-weight:800;color:#c62828'>⏰ "
                    + (heading or f"{len(rows)} invoice(s) outstanding for more than {INVOICE_OVERDUE_DAYS} days")
                    + f" — Rs. {inr(tot)}</div><br>"
                    "Select an invoice and send a payment reminder by e-mail or WhatsApp, or record a follow-up.")
        v.addWidget(hd)
        self.t = make_table(["Invoice No.", "Date", "Client", "Total", "Outstanding", "Age (days)", "Follow-ups",
                             "Promised", "Next Follow-up", "Mobile"], 2)
        fill_table(self.t, [[r["inv_no"], dmy(r["inv_date"]), r["client_name"], inr(r["total"]), inr(r["outstanding"]),
                             r["age_days"], r.get("fu_count", 0), dmy(r["fu_promise"]) if r.get("fu_promise") else "",
                             dmy(r["fu_next"]) if r.get("fu_next") else "", r.get("client_phone") or ""] for r in rows],
                   keys=[r["id"] for r in rows])
        for i in range(self.t.rowCount()):
            self.t.item(i, 5).setBackground(QColor(TASK_RED))
        if rows:
            self.t.selectRow(0)
        v.addWidget(self.t)
        h = QHBoxLayout()
        h.addWidget(btn("✉ / 🟢 Send Reminder (E-mail / WhatsApp)", "green", self.send))
        h.addWidget(btn("📞 Follow-up", None, self.followup))
        h.addWidget(btn("📄 Invoice PDF", "orange", lambda: self.main.invoice_pdf(selected_key(self.t))))
        h.addWidget(btn("₹ Payment Received", "blue", self.receipt))
        h.addStretch()
        h.addWidget(btn("Close", "gray", self.close))
        v.addLayout(h)

    def send(self):
        iid = selected_key(self.t)
        if not iid:
            return err(self, "Please select an invoice")
        self.main.send_payment_reminder(iid)

    def followup(self):
        iid = selected_key(self.t)
        if not iid:
            return err(self, "Please select an invoice")
        FollowupDialog(self.main, iid).exec_()

    def receipt(self):
        iid = selected_key(self.t)
        r = next((x for x in self.rows if x["id"] == iid), None)
        if self.main.pages["lit"].record_receipt(r):
            self.t.removeRow(self.t.currentRow())


class InvoiceFolderWidget(QWidget):
    """Choose / open the dedicated folder where every invoice PDF is saved (per PC)."""

    def __init__(self, main, compact=False):
        super().__init__()
        self.main = main
        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(QLabel("📁 Folder:" if compact else "📁 Invoice folder:"))
        self.path = QLineEdit()
        self.path.setReadOnly(True)
        self.path.setMinimumWidth(160)
        h.addWidget(self.path, 1)
        h.addWidget(btn("Change", "blue", self.browse))
        h.addWidget(btn("Open", "orange", self.open))
        if not compact:
            h.addWidget(btn("Reset to default", "gray", self.reset))
        self.show_path()

    def show_path(self):
        self.path.setText(invoice_folder(self.main.cfg))
        self.path.setToolTip("File name format: Name of client_Invoice number.pdf")

    def browse(self):
        start = invoice_folder(self.main.cfg)
        d = QFileDialog.getExistingDirectory(self, "Select folder for invoices", start if os.path.isdir(start) else HOME)
        if not d:
            return
        if not os.access(d, os.W_OK):
            return err(self, "You do not have permission to save files in this folder")
        self.main.cfg["invoice_dir"] = d
        save_cfg(self.main.cfg)
        self.main.refresh_invoice_folder()
        self.main.status("Invoices will be saved in " + d)

    def reset(self):
        self.main.cfg.pop("invoice_dir", None)
        save_cfg(self.main.cfg)
        self.main.refresh_invoice_folder()

    def open(self):
        f = invoice_folder(self.main.cfg)
        try:
            os.makedirs(f, exist_ok=True)
            open_path(f)
        except OSError as e:
            err(self, e)


class InvoicingPage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Invoicing", "System-generated GST invoices for litigation, tasks and other "
                                       "services • receipts • reminders by e-mail / WhatsApp • Client & Firm Master."))
        cg = QHBoxLayout()
        self.k_bill = Card("Billed (this FY)", CARD_COLORS[2], "🧾")
        self.k_rec = Card("Received + TDS", CARD_COLORS[1], "✅")
        self.k_out = Card("Outstanding", CARD_COLORS[3], "⏳")
        self.k_od = Card(f"Outstanding > {INVOICE_OVERDUE_DAYS} days", ("#e53935", "#ff6f61"), "⏰")
        for c in (self.k_bill, self.k_rec, self.k_out, self.k_od):
            c.small()
            cg.addWidget(c)
        self.lay.addLayout(cg)
        self.tabs = QTabWidget()
        self.tabs.setMinimumHeight(560)
        self.lay.addWidget(self.tabs)
        # invoices tab
        w = QWidget()
        v = QVBoxLayout(w)
        f = QHBoxLayout()
        f.addWidget(btn("➕ New", "green", self.new_invoice))
        f.addWidget(btn("₹ Payment", "orange", self.receipt))
        f.addWidget(btn("📞 Follow-up", "blue", self.followup))
        pass
        self.f_st = QComboBox()
        self.f_st.addItems(["All statuses", "Unpaid", "Part Paid", "Paid", "Cancelled"])
        f.addWidget(self.f_st)
        pass
        self.f_src = QComboBox()
        self.f_src.addItem("All sources", "")
        self.f_src.addItem("Litigation", "case")
        self.f_src.addItem("Task billing", "task")
        self.f_src.addItem("Other services", "manual")
        f.addWidget(self.f_src)
        self.f_od = QCheckBox(f"> {INVOICE_OVERDUE_DAYS} days")
        self.f_od.setToolTip(f"Only invoices outstanding for more than {INVOICE_OVERDUE_DAYS} days")
        f.addWidget(self.f_od)
        self.f_fu = QCheckBox("Follow-up due")
        f.addWidget(self.f_fu)
        for x in (self.f_st.currentIndexChanged, self.f_src.currentIndexChanged, self.f_od.toggled, self.f_fu.toggled):
            x.connect(self.load)
        f.addStretch()
        v.addLayout(f)
        self.t = make_table(["Invoice No.", "Date", "Client", "Matter / Source", "Total", "Received", "TDS",
                             "Balance Due", "Age (days)", "Next Follow-up", "Status"])
        self.t.doubleClicked.connect(lambda: self.main.invoice_pdf(selected_key(self.t)))
        v.addWidget(self.t)
        a = QHBoxLayout()
        a.addWidget(btn("📄 PDF", "orange", lambda: self.main.invoice_pdf(selected_key(self.t))))
        a.addWidget(btn("📨 Send", "blue", lambda: self._send(False)))
        a.addWidget(btn("⏰ Reminder", "red", lambda: self._send(True)))
        a.addWidget(btn("📝 Service Details", None, self.edit_details))
        a.addWidget(btn("Cancel", "gray", self.cancel))
        a.addWidget(btn("Export CSV", "blue", self.csv))
        a.addStretch()
        v.addLayout(a)
        self.sum = QLabel()
        self.sum.setStyleSheet("background:#f1ecff; border-radius:8px; padding:8px; color:#2a0a7a;")
        v.addWidget(self.sum)
        fb = QHBoxLayout()
        self.folder = InvoiceFolderWidget(main, compact=True)
        fb.addWidget(self.folder, 1)
        fb.addWidget(btn("💾 Save All", "green", self.save_all))
        v.insertLayout(0, fb)
        self.tabs.addTab(w, "🧾 Invoices")
        self.receipts = ReceiptsWidget(main)
        self.tabs.addTab(self.receipts, "💵 Receipts")
        # clients tab
        w = QWidget()
        v = QVBoxLayout(w)
        a = QHBoxLayout()
        a.addWidget(btn("➕ New Client", "green", self.new_client))
        a.addWidget(btn("✎ Edit", "blue", self.edit_client))
        a.addWidget(btn("📨 Send Message", "orange", self.msg_client))
        a.addWidget(btn("Delete", "red", self.del_client))
        a.addStretch()
        v.addLayout(a)
        self.tc = make_table(["#", "Client", "Contact", "GSTIN", "State", "PAN", "E-mail", "Mobile"], 1)
        self.tc.doubleClicked.connect(self.edit_client)
        v.addWidget(self.tc)
        self.tabs.addTab(w, "👥 Client Master")
        # firm master
        self.firm = FirmMasterWidget(main)
        fsc = QScrollArea()                 # keep Firm Master from widening the whole page
        fsc.setWidgetResizable(True)
        fsc.setFrameShape(QFrame.NoFrame)
        fsc.setWidget(self.firm)
        self.tabs.addTab(fsc, "🏢 Firm Master")
        self.tabs.currentChanged.connect(lambda _: self.refresh())
        self.rows, self.clients = [], []

    def refresh(self):
        i = self.tabs.currentIndex()
        self.load()
        if i == 1:
            self.receipts.load()
        elif i == 2:
            self.load_clients()
        elif i == 3:
            self.firm.load()

    def load(self):
        try:
            allrows = self.client.call("invoice_list")
            self.rows = self.client.call("invoice_list", status=self.f_st.currentText() if self.f_st.currentIndex() else "",
                                         source=self.f_src.currentData(), overdue=self.f_od.isChecked(),
                                         followup_due=self.f_fu.isChecked())
        except Exception as e:
            return self.main.status(str(e))
        t = date.today()
        fy0 = date(t.year if t.month >= 4 else t.year - 1, 4, 1).isoformat()
        live = [r for r in allrows if r["status"] != "Cancelled"]
        od = [r for r in live if r["outstanding"] > 0.009 and r["age_days"] > INVOICE_OVERDUE_DAYS]
        self.k_bill.set("Rs. " + inr(sum(r["total"] for r in live if r["inv_date"] >= fy0))[:-3],
                        f"{sum(1 for r in live if r['inv_date'] >= fy0)} invoice(s) since 01-04")
        self.k_rec.set("Rs. " + inr(sum(r["received"] + r["tds"] for r in live))[:-3], "all invoices")
        self.k_out.set("Rs. " + inr(sum(r["outstanding"] for r in live))[:-3],
                       f"{sum(1 for r in live if r['outstanding'] > 0.009)} invoice(s)")
        self.k_od.set("Rs. " + inr(sum(r["outstanding"] for r in od))[:-3], f"{len(od)} invoice(s) — send reminders")
        fill_table(self.t, [[r["inv_no"], dmy(r["inv_date"]), r["client_name"], r["law"], inr(r["total"]),
                             inr(r["received"]), inr(r["tds"]), inr(r["outstanding"]),
                             r["age_days"] if r["outstanding"] > 0.009 else "",
                             (dmy(r["fu_next"]) + (" ⏰" if r["fu_due"] else "")) if r["fu_next"] and r["outstanding"] > 0.009
                             else "", r["status"]] for r in self.rows],
                   keys=[r["id"] for r in self.rows])
        for i, r in enumerate(self.rows):
            for j in (4, 5, 6, 7):
                self.t.item(i, j).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            if r["outstanding"] > 0.009 and r["age_days"] > INVOICE_OVERDUE_DAYS:
                for j in range(self.t.columnCount()):
                    self.t.item(i, j).setBackground(QColor(TASK_RED))
            if r["fu_due"]:
                self.t.item(i, 9).setBackground(QColor("#fed7aa"))
        s = lambda k: sum(r[k] for r in self.rows if r["status"] != "Cancelled")
        self.sum.setText(f"<b>Shown:</b> {len(self.rows)} &nbsp; <b>Total:</b> Rs. {inr(s('total'))} &nbsp; "
                         f"<b>Received:</b> Rs. {inr(s('received'))} &nbsp; <b>TDS:</b> Rs. {inr(s('tds'))} &nbsp; "
                         f"<b>Outstanding:</b> Rs. {inr(s('outstanding'))} &nbsp; (red rows = outstanding > "
                         f"{INVOICE_OVERDUE_DAYS} days)")

    def new_invoice(self):
        self.main.generate_invoice()
        self.load()

    def edit_details(self):
        iid = selected_key(self.t)
        if not iid:
            return err(self, "Please select an invoice")
        try:
            inv = self.client.call("invoice_get", id=iid)["invoice"]
        except Exception as e:
            return err(self, e)
        dlg = TextDialog(self, "Additional details of service",
                         f"<b>Invoice {html.escape(inv['inv_no'])}</b> — {html.escape(inv.get('client_name') or '')}"
                         "<br>These details are printed on the invoice under 'Details of services rendered':",
                         inv.get("service_details") or "")
        if dlg.exec_():
            try:
                self.client.call("invoice_details", id=iid, service_details=dlg.t.toPlainText())
                p = self.main.invoice_pdf(iid, ask=False)
                if p and QMessageBox.question(self, APP_NAME, "Details saved and invoice PDF updated.\n\nOpen the "
                                                              "invoice now?") == QMessageBox.Yes:
                    open_path(p)
            except Exception as e:
                err(self, e)

    def save_all(self):
        rows = [r for r in self.rows if r["status"] != "Cancelled"]
        if not rows:
            return err(self, "No invoices in the current list")
        if QMessageBox.question(self, APP_NAME, f"Save {len(rows)} invoice PDF(s) in:\n{invoice_folder(self.main.cfg)}"
                                                "\n\n(existing files with the same name are updated)") != QMessageBox.Yes:
            return
        ok = 0
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for r in rows:
                if self.main.invoice_pdf(r["id"], ask=False):
                    ok += 1
        finally:
            QApplication.restoreOverrideCursor()
        info(self, f"{ok} invoice(s) saved in:\n{invoice_folder(self.main.cfg)}")
        self.folder.open()

    def _send(self, reminder):
        iid = selected_key(self.t)
        if not iid:
            return err(self, "Please select an invoice")
        if reminder:
            if self.main.send_payment_reminder(iid):
                self.load()
        else:
            self.main.send_invoice(iid)

    def followup(self):
        iid = selected_key(self.t)
        if not iid:
            return err(self, "Please select an invoice")
        FollowupDialog(self.main, iid).exec_()
        self.load()

    def receipt(self):
        iid = selected_key(self.t)
        if self.main.pages["lit"].record_receipt(next((r for r in self.rows if r["id"] == iid), None)):
            self.load()

    def cancel(self):
        iid = selected_key(self.t)
        if iid and QMessageBox.question(self, APP_NAME, "Cancel this invoice?") == QMessageBox.Yes:
            try:
                self.client.call("invoice_cancel", id=iid)
                self.load()
            except Exception as e:
                err(self, e)

    def csv(self):
        if not self.rows:
            return err(self, "Nothing to export")
        os.makedirs(REPORT_DIR, exist_ok=True)
        p, _ = QFileDialog.getSaveFileName(self, "Save CSV", os.path.join(REPORT_DIR, "Invoice_Register.csv"),
                                           "CSV (*.csv)")
        if not p:
            return
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["Invoice No", "Date", "Client", "GSTIN", "Matter", "Taxable", "CGST", "SGST", "IGST",
                        "Reimbursement", "Total", "Received", "TDS", "Outstanding", "Age", "Status"])
            for r in self.rows:
                taxable = r["fee_total"] + (0 if r["pure_agent"] else r["ope_total"])
                w.writerow([r["inv_no"], r["inv_date"], r["client_name"], r.get("client_gstin"), r["law"], taxable,
                            r["cgst"], r["sgst"], r["igst"], r["ope_total"] if r["pure_agent"] else 0, r["total"],
                            r["received"], r["tds"], r["outstanding"], r["age_days"], r["status"]])
        open_path(p)

    # ---------------------------------------------------------- clients
    def load_clients(self):
        try:
            self.clients = self.client.call("client_list")
        except Exception as e:
            return err(self, e)
        fill_table(self.tc, [[c["id"], c["name"], c["contact"], c["gstin"], state_of(c["gstin"]), c["pan"],
                              c["email"], c["phone"]] for c in self.clients], keys=[c["id"] for c in self.clients])

    def _sel_client(self):
        cid = selected_key(self.tc)
        c = next((x for x in self.clients if x["id"] == cid), None)
        if not c:
            err(self, "Please select a client")
        return c

    def new_client(self):
        d = ClientDialog(self)
        if d.exec_():
            try:
                self.client.call("client_save", client=d.value())
                self.load_clients()
            except Exception as e:
                err(self, e)

    def edit_client(self, *a):
        c = self._sel_client()
        if c:
            d = ClientDialog(self, c)
            if d.exec_():
                try:
                    self.client.call("client_save", client=d.value())
                    self.load_clients()
                except Exception as e:
                    err(self, e)

    def msg_client(self):
        c = self._sel_client()
        if c:
            s = self.main.settings
            body = f"Dear {c['contact'] or c['name']},\n\n\n\n" + firm_signature(s, self.main.user["full_name"])
            ShareDialog(self.main, "Message to client", c["name"], c["email"], c["phone"],
                        [("General message", f"{s.get('company_name') or FIRM_NAME}", body)]).exec_()

    def del_client(self):
        c = self._sel_client()
        if c and QMessageBox.question(self, APP_NAME, f"Delete client '{c['name']}' from master?") == QMessageBox.Yes:
            try:
                self.client.call("client_delete", id=c["id"])
                self.load_clients()
            except Exception as e:
                err(self, e)



# =============================================================================
# 16D. GUI : DOCUMENTS (CASE / TASK FILES), TASK DETAILS, RECEIPTS
# =============================================================================
def receipt_message(rc, inv, s, user_name=""):
    subject = f"Payment receipt {rc['rcpt_no']} - Invoice {inv['inv_no']}"
    body = (f"Dear {inv.get('client_name') or 'Sir/Madam'},\n\nGreetings from {s.get('company_name') or FIRM_NAME}."
            f"\n\nWe thankfully acknowledge receipt of Rs. {inr(rc['amount'])} on {dmy(rc['rdate'])} by {rc['mode']}"
            + (f" (Ref: {rc['ref']})" if rc.get("ref") else "") +
            f" against our Invoice No. {inv['inv_no']} dated {dmy(inv['inv_date'])}."
            + (f"\nTDS deducted: Rs. {inr(rc['tds'])}." if rc["tds"] else "")
            + (f"\nBalance outstanding: Rs. {inr(rc['balance_after'])}." if rc["balance_after"] > 0.01
               else "\nThe invoice is now fully settled. Thank you!")
            + "\n\nPlease find the payment receipt attached.")
    return subject, body + "\n\n" + firm_signature(s, user_name)


class DocsWidget(QWidget):
    """Upload / download / delete files attached to a case or a task."""

    def __init__(self, main, ref_type, ref_id):
        super().__init__()
        self.main, self.client, self.ref_type, self.ref_id = main, main.client, ref_type, ref_id
        v = QVBoxLayout(self)
        h = QHBoxLayout()
        h.addWidget(btn("⬆ Upload Files", "green", self.upload))
        h.addWidget(btn("⬇ Download / Open", "blue", self.download))
        h.addWidget(btn("Delete", "red", self.delete))
        self.prog = QLabel("")
        self.prog.setStyleSheet("color:#7c3aed; font-weight:600;")
        h.addWidget(self.prog)
        h.addStretch()
        v.addLayout(h)
        self.t = make_table(["Document", "Description", "Size", "Uploaded By", "Date", "Source"], 0)
        self.t.doubleClicked.connect(self.download)
        v.addWidget(self.t)
        note = QLabel("Keep notices, replies, orders, working papers and client documents here — every team member "
                      "can open them from any PC on the LAN. Max 100 MB per file.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#8b84a8; font-size:8.5pt;")
        v.addWidget(note)
        self.rows = []
        self.load()

    def load(self):
        try:
            self.rows = self.client.call("doc_list", ref_type=self.ref_type, ref_id=self.ref_id)
        except Exception as e:
            return self.prog.setText(str(e))
        fill_table(self.t, [["📄 " + r["name"], r.get("title") or "",
                             f"{(r['size'] or 0) / 1024 / 1024:.2f} MB" if (r["size"] or 0) > 1024 * 1024
                             else f"{(r['size'] or 0) / 1024:.0f} KB", r.get("by_name") or r["by_user"],
                             dmy(r["ts"]), r["source"]] for r in self.rows], keys=[r["id"] for r in self.rows])

    def upload(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select files to upload")
        if not paths:
            return
        big = [p for p in paths if os.path.getsize(p) > 100 * 1024 * 1024]
        if big:
            return err(self, "These files are larger than 100 MB:\n" + "\n".join(os.path.basename(p) for p in big))
        self.prog.setText(f"Uploading {len(paths)} file(s)...")
        rt, rid, cl = self.ref_type, self.ref_id, self.client

        def work():
            for p in paths:
                with open(p, "rb") as f:
                    cl.request("doc_upload", f.read(), timeout=900, ref_type=rt, ref_id=rid,
                               name=os.path.basename(p))
            return len(paths)

        def done(n, e):
            self.prog.setText("")
            if e:
                err(self, e, "Upload failed")
            else:
                self.main.status(f"{n} file(s) uploaded")
            self.load()
        run_task(work, done)

    def download(self, *a):
        did = selected_key(self.t)
        r = next((x for x in self.rows if x["id"] == did), None)
        if not r:
            return err(self, "Please select a document")
        self.main.download_file(r["file_id"], self.prog)

    def delete(self):
        did = selected_key(self.t)
        if not did:
            return err(self, "Please select a document")
        if QMessageBox.question(self, APP_NAME, "Delete this document?") == QMessageBox.Yes:
            try:
                self.client.call("doc_delete", id=did)
                self.load()
            except Exception as e:
                err(self, e)


class TaskDetailDialog(QDialog):
    def __init__(self, page, tid):
        super().__init__(page.main)
        self.page, self.main, self.client, self.tid = page, page.main, page.client, tid
        self.setWindowTitle(f"{APP_NAME} - Task #{tid}")
        self.setWindowIcon(app_icon())
        self.resize(900, 640)
        v = QVBoxLayout(self)
        self.head = QLabel()
        self.head.setWordWrap(True)
        self.head.setTextInteractionFlags(Qt.TextSelectableByMouse)
        v.addWidget(self.head)
        h = QHBoxLayout()
        for label, kind, fn in (("✎ Edit", "blue", self.page.edit_task), ("▶ Start", "orange",
                                lambda: self.page.set_status("In Progress")), ("✔ Complete", "green", self.page.complete),
                                ("💬 Add Remark", None, self.page.remark), ("📨 Send", "blue", self.page.share)):
            h.addWidget(btn(label, kind, lambda _=False, f=fn: self._act(f)))
        h.addStretch()
        v.addLayout(h)
        tabs = QTabWidget()
        self.docs = DocsWidget(self.main, "task", tid)
        tabs.addTab(self.docs, "📁 Task Files")
        self.rem = QTextBrowser()
        tabs.addTab(self.rem, "💬 Remarks / Challenges")
        v.addWidget(tabs, 1)
        self.t = None
        self.refresh()

    def refresh(self):
        try:
            self.t = t = self.client.call("task_get", id=self.tid)
        except Exception as e:
            self.head.setText(html.escape(str(e)))
            return
        dl = t["days_left"]
        left = ("" if t["status"] in ("Completed", "Cancelled") else
                (f"OVERDUE by {-dl} day(s)" if dl < 0 else ("Due TODAY" if dl == 0 else f"{dl} day(s) left")))
        E = lambda x: html.escape(str(x or "—"))
        self.head.setText(
            f"<div style='font-size:15pt;font-weight:800;color:#6d28d9'>{'🔥 ' if t['urgent'] else ''}{E(t['title'])}</div>"
            f"<div style='background:{t['colour']};padding:6px;border-radius:8px'>"
            f"<b>Client:</b> {E(t['client'])} &nbsp; <b>Assigned to:</b> {E(t['assigned_name'])} &nbsp; "
            f"<b>Created by:</b> {E(t['creator_name'])}<br><b>Due:</b> {dmy(t['due_date'])} ({left}) &nbsp; "
            f"<b>Status:</b> {E(t['status'])} &nbsp; <b>Priority:</b> {E(t['priority'])} &nbsp; "
            + ("<b>Est. billing:</b> 🔒 Hidden" if t.get("fee_hidden") else f"<b>Est. billing:</b> Rs. {inr(t['est_amount'])}")
            + (f" &nbsp; <b>Billing:</b> {E(t['billing_option'])}"
               + ("" if t.get("fee_hidden") else f" Rs. {inr(t['billing_amount'])}") if t["billing_option"] else "")
            + (" &nbsp; <i>(view only — shared by Manager)</i>" if t.get("view_only") else "") + "</div>"
            + (f"<b>Details:</b> {E(t['description'])}" if t["description"] else ""))
        self.rem.setHtml(html.escape(t["challenges"] or "No remarks yet.").replace("\n", "<br>"))
        self.docs.load()

    def _act(self, fn):
        self.page._forced_task = self.t
        try:
            fn()
        finally:
            self.page._forced_task = None
        self.refresh()


class AccessDialog(QDialog):
    """Manager decides who can see a confidential case / task and who can see its fees."""

    def __init__(self, main, ref_type, ref_id, title):
        super().__init__(main)
        self.main, self.client, self.rt, self.rid = main, main.client, ref_type, ref_id
        self.setWindowTitle(f"{APP_NAME} - Access rights")
        self.setMinimumSize(600, 520)
        d = self.client.call("access_get", ref_type=ref_type, ref_id=ref_id)
        v = QVBoxLayout(self)
        hd = QLabel(f"<div style='font-size:13pt;font-weight:800;color:#6d28d9'>🔒 Access rights</div><br>"
                    f"<b>{html.escape(title)}</b>")
        hd.setWordWrap(True)
        v.addWidget(hd)
        self.restricted = QCheckBox("Confidential case — visible only to Manager, handler && members ticked below")
        self.hide_fees = QCheckBox("Hide fees / billing of this task from the team (except 'Can view fees')")
        if ref_type == "case":
            self.restricted.setChecked(bool(d["restricted"]))
            v.addWidget(self.restricted)
            note = ("Fees, estimate, invoices and receipts of every case are visible only to Manager(s). Tick "
                    "'Can view fees' to let a member see the fee summary of this case.")
        else:
            self.hide_fees.setChecked(bool(d["hide_fees"]))
            v.addWidget(self.hide_fees)
            note = ("A task is always visible to its creator, its assignee and Manager(s). Tick 'Can view' to share it "
                    "(read-only) with other members — they find it under 'Shared with me'.")
        n = QLabel(note)
        n.setWordWrap(True)
        n.setStyleSheet("color:#8b84a8; font-size:8.5pt;")
        v.addWidget(n)
        self.t = QTableWidget(len(d["users"]), 3)
        self.t.setHorizontalHeaderLabels(["Team member", "Can view", "Can view fees"])
        self.t.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.t.setColumnWidth(1, 100)
        self.t.setColumnWidth(2, 130)
        self.t.verticalHeader().setVisible(False)
        self.users = d["users"]
        special = {d.get("assigned_to"): "handler / assignee", d.get("created_by"): "creator"}
        for i, u in enumerate(self.users):
            extra = special.get(u["username"])
            it = QTableWidgetItem(u["full_name"] + (f"   ({extra} — always sees it)" if extra else ""))
            it.setFlags(Qt.ItemIsEnabled)
            self.t.setItem(i, 0, it)
            for j, k in ((1, "can_view"), (2, "can_fees")):
                c = QTableWidgetItem("")
                c.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                c.setCheckState(Qt.Checked if u[k] else Qt.Unchecked)
                self.t.setItem(i, j, c)
        v.addWidget(self.t, 1)
        h = QHBoxLayout()
        h.addWidget(btn("💾 Save Access", "green", self.save))
        h.addWidget(btn("Cancel", "gray", self.reject))
        v.addLayout(h)

    def save(self):
        users = [{"username": u["username"], "can_view": self.t.item(i, 1).checkState() == Qt.Checked,
                  "can_fees": self.t.item(i, 2).checkState() == Qt.Checked} for i, u in enumerate(self.users)]
        try:
            self.client.call("access_set", ref_type=self.rt, ref_id=self.rid, restricted=self.restricted.isChecked(),
                             hide_fees=self.hide_fees.isChecked(), users=users)
            self.accept()
        except Exception as e:
            err(self, e)


class ForceChangePwDialog(QDialog):
    def __init__(self, parent, name):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} - Set a new password")
        self.setMinimumWidth(420)
        f = QFormLayout(self)
        lb = QLabel(f"Welcome <b>{html.escape(name)}</b>.<br>Your password was reset by the Manager. Please set your "
                    "own new password to continue (minimum 6 characters).")
        lb.setWordWrap(True)
        f.addRow(lb)
        self.p1, self.p2 = QLineEdit(), QLineEdit()
        for e in (self.p1, self.p2):
            e.setEchoMode(QLineEdit.Password)
        f.addRow("New password:", self.p1)
        f.addRow("Confirm password:", self.p2)
        h = QHBoxLayout()
        h.addWidget(btn("Save & Continue".replace("&", "&&"), "green", self._ok))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)

    def _ok(self):
        if len(self.p1.text()) < 6:
            return err(self, "Password must have at least 6 characters")
        if self.p1.text() != self.p2.text():
            return err(self, "Passwords do not match")
        self.accept()


class ReceiptsWidget(QWidget):
    def __init__(self, main):
        super().__init__()
        self.main, self.client = main, main.client
        v = QVBoxLayout(self)
        h = QHBoxLayout()
        h.addWidget(btn("📄 Receipt PDF", "orange", lambda: self.main.receipt_pdf(selected_key(self.t), ask=True)))
        h.addWidget(btn("📨 Send Receipt (E-mail / WhatsApp)", "green", lambda: self.main.send_receipt(selected_key(self.t))))
        h.addWidget(btn("✎ Edit Payment / Date", "blue", self.edit))
        h.addWidget(btn("Delete Receipt", "red", self.delete))
        h.addStretch()
        v.addLayout(h)
        self.t = make_table(["Receipt No.", "Date", "Client", "Invoice No.", "Amount (Rs.)", "TDS (Rs.)", "Mode",
                             "Reference"], 2)
        self.t.doubleClicked.connect(lambda: self.main.receipt_pdf(selected_key(self.t), ask=True))
        v.addWidget(self.t)
        self.sum = QLabel()
        self.sum.setStyleSheet("background:#f1ecff; border-radius:8px; padding:8px; color:#2a0a7a;")
        v.addWidget(self.sum)

    def edit(self):
        rid = selected_key(self.t)
        if not rid:
            return err(self, "Please select a receipt")
        if self.main.edit_receipt(rid):
            self.load()

    def delete(self):
        rid = selected_key(self.t)
        if not rid:
            return err(self, "Please select a receipt")
        if QMessageBox.question(self, APP_NAME, "Delete this receipt? The invoice balance will be restored.") \
                == QMessageBox.Yes:
            try:
                self.client.call("receipt_delete", id=rid)
                self.load()
            except Exception as e:
                err(self, e)

    def load(self):
        try:
            rows = self.client.call("receipt_list")
        except Exception as e:
            return self.main.status(str(e))
        fill_table(self.t, [[r["rcpt_no"] or f"#{r['id']}", dmy(r["rdate"]), r["client_name"], r["inv_no"],
                             inr(r["amount"]), inr(r["tds"]), r["mode"], r["ref"]] for r in rows],
                   keys=[r["id"] for r in rows])
        for i in range(self.t.rowCount()):
            for j in (4, 5):
                self.t.item(i, j).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.sum.setText(f"<b>Receipts:</b> {len(rows)} &nbsp; <b>Amount received:</b> Rs. "
                         f"{inr(sum(r['amount'] for r in rows))} &nbsp; <b>TDS:</b> Rs. {inr(sum(r['tds'] for r in rows))}")

# =============================================================================
# 17. GUI : ADMIN & SETTINGS
# =============================================================================
class UserDialog(QDialog):
    def __init__(self, parent, u=None):
        super().__init__(parent)
        self.new = u is None
        u = u or {}
        self.setWindowTitle(f"{APP_NAME} - {'Add' if self.new else 'Edit'} Team Member")
        self.setMinimumWidth(460)
        f = QFormLayout(self)
        self.w = {}

        def le(key, ph=""):
            e = QLineEdit(str(u.get(key) or ""))
            e.setPlaceholderText(ph)
            self.w[key] = e
            return e
        un = le("username", "e.g. ramesh")
        un.setEnabled(self.new)
        f.addRow("Username*:", un)
        f.addRow("Full name*:", le("full_name"))
        pw = le("password", "min 6 characters" if self.new else "leave blank to keep current password")
        pw.setEchoMode(QLineEdit.Password)
        pw.setText("")
        f.addRow("Password" + ("*" if self.new else "") + ":", pw)
        self.role = QComboBox()
        self.role.addItem("Team Member", "member")
        self.role.addItem("Manager / Admin", "admin")
        self.role.setCurrentIndex(1 if u.get("role") == "admin" else 0)
        f.addRow("Role:", self.role)
        f.addRow("Designation:", le("designation", "e.g. Article Assistant"))
        f.addRow("Department:", le("department", "e.g. Audit / GST / Income Tax"))
        f.addRow("E-mail:", le("email"))
        f.addRow("Mobile / WhatsApp:", le("phone", "10-digit mobile number"))
        f.addRow("PAN:", le("pan"))
        f.addRow("Bank A/c No.:", le("bank"))
        f.addRow("Date of joining:", le("doj", "YYYY-MM-DD"))

        def money(key):
            s = QDoubleSpinBox()
            s.setRange(0, 10000000)
            s.setDecimals(2)
            s.setPrefix("Rs. ")
            s.setGroupSeparatorShown(True)
            s.setValue(float(u.get(key) or 0))
            self.w[key] = s
            return s
        f.addRow("Monthly gross salary:", money("salary"))
        f.addRow("Professional Tax / month:", money("pt"))
        f.addRow("TDS / month:", money("tds"))
        self.active = QCheckBox("Active (can log in)")
        self.active.setChecked(bool(u.get("active", 1)))
        f.addRow(self.active)
        h = QHBoxLayout()
        h.addWidget(btn("Save", "green", self.accept))
        h.addWidget(btn("Cancel", "gray", self.reject))
        f.addRow(h)

    def value(self):
        d = {k: (w.value() if isinstance(w, QDoubleSpinBox) else w.text().strip()) for k, w in self.w.items()}
        d["role"] = self.role.currentData()
        d["active"] = 1 if self.active.isChecked() else 0
        return d


class AdminPage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Admin", "Manage team members, salary details and firm settings."))
        g = QGroupBox("Team Members")
        v = QVBoxLayout(g)
        h = QHBoxLayout()
        h.addWidget(btn("➕ Add Member", "green", self.add))
        h.addWidget(btn("✎ Edit Selected", "blue", self.edit))
        h.addWidget(btn("🔑 Reset Password", "orange", self.reset_pw))
        h.addWidget(btn("👑 Make Admin", None, lambda: self.set_role("admin")))
        h.addWidget(btn("Remove Admin", "gray", lambda: self.set_role("member")))
        h.addStretch()
        v.addLayout(h)
        self.table = make_table(["Username", "Full Name", "Role", "Designation", "Department", "E-mail",
                                 "Monthly Salary", "Active"], 1)
        self.table.setMinimumHeight(280)
        self.table.doubleClicked.connect(self.edit)
        v.addWidget(self.table)
        self.lay.addWidget(g)
        g2 = QGroupBox("Firm details")
        v2 = QHBoxLayout(g2)
        v2.addWidget(QLabel("Firm name, address, GSTIN, bank, logo, invoice settings and reminder days are "
                            "maintained in <b>Invoicing › Firm Master</b>."))
        v2.addWidget(btn("Open Firm Master", "orange", lambda: self.main.open_firm_master()))
        self.lay.addWidget(g2)
        g3 = QGroupBox("Server Information")
        v3 = QVBoxLayout(g3)
        self.srv = QLabel()
        self.srv.setWordWrap(True)
        self.srv.setTextInteractionFlags(Qt.TextSelectableByMouse)
        v3.addWidget(self.srv)
        self.lay.addWidget(g3)
        self.lay.addStretch()
        self.users = []

    def refresh(self):
        try:
            self.users = self.client.call("user_list_full")
        except Exception as e:
            return err(self, e)
        fill_table(self.table, [[u["username"], u["full_name"], "👑 Admin" if u["role"] == "admin" else "Member",
                                 u["designation"], u["department"], u["email"], inr(u["salary"]),
                                 "Yes" if u["active"] else "No"] for u in self.users],
                   keys=[u["username"] for u in self.users])
        if self.main.ctrl.server:
            self.srv.setText(f"<b>This PC is hosting the server.</b><br>Team members should connect to IP: "
                             f"<b>{', '.join(local_ips())}</b> &nbsp; Port: <b>{PORT}</b><br>"
                             f"Database &amp; shared files: {html.escape(self.main.ctrl.server.data_dir)}<br>"
                             "Back up this folder regularly. Keep this app open while the team is working.")
        else:
            self.srv.setText(f"Connected to server at <b>{self.client.host}:{PORT}</b>.")

    def _sel(self):
        un = selected_key(self.table)
        u = next((x for x in self.users if x["username"] == un), None)
        if not u:
            err(self, "Please select a team member")
        return u

    def reset_pw(self):
        u = self._sel()
        if not u:
            return
        dlg = QDialog(self)
        dlg.setWindowTitle(f"{APP_NAME} - Reset password")
        f = QFormLayout(dlg)
        f.addRow(QLabel(f"Reset password of <b>{html.escape(u['full_name'])}</b> ({html.escape(u['username'])})"))
        pw = QLineEdit()
        pw.setPlaceholderText("leave blank to generate a temporary password")
        f.addRow("New password:", pw)
        f.addRow(QLabel("The member must set his / her own new password at the next login."))
        h = QHBoxLayout()
        h.addWidget(btn("🔑 Reset", "orange", dlg.accept))
        h.addWidget(btn("Cancel", "gray", dlg.reject))
        f.addRow(h)
        if not dlg.exec_():
            return
        try:
            r = self.client.call("user_reset_pw", username=u["username"], password=pw.text())
        except Exception as e:
            return err(self, e)
        QApplication.clipboard().setText(r["password"])
        info(self, f"Password of {u['full_name']} has been reset.\n\nTemporary password:  {r['password']}\n\n"
                   "(copied to clipboard — share it with the member personally; it must be changed at next login)")

    def set_role(self, role):
        u = self._sel()
        if not u:
            return
        if (u["role"] == "admin") == (role == "admin"):
            return info(self, f"{u['full_name']} is already {'an Admin' if role == 'admin' else 'a Team Member'}.")
        q = (f"Appoint {u['full_name']} as Admin?\n\nAn Admin can see salaries, fees, invoices, all cases and tasks, "
             "reset passwords and appoint other admins." if role == "admin"
             else f"Remove admin rights of {u['full_name']}?")
        if QMessageBox.question(self, APP_NAME, q) != QMessageBox.Yes:
            return
        try:
            self.client.call("user_set_role", username=u["username"], role=role)
            self.main.reload_users()
            self.refresh()
        except Exception as e:
            err(self, e)

    def add(self):
        d = UserDialog(self)
        if d.exec_():
            self._save(d.value(), True)

    def edit(self, *a):
        un = selected_key(self.table)
        u = next((x for x in self.users if x["username"] == un), None)
        if not u:
            return err(self, "Please select a member")
        d = UserDialog(self, u)
        if d.exec_():
            self._save(d.value(), False)

    def _save(self, val, new):
        try:
            self.client.call("user_save", user=val, new=new)
            info(self, f"Team member '{val['username']}' saved")
            self.main.reload_users()
            self.refresh()
        except Exception as e:
            err(self, e)


class SettingsPage(Page):
    def __init__(self, main):
        super().__init__(main)
        self.lay.addWidget(title_block("Settings & Profile"))
        g = QGroupBox("My Profile")
        self.prof = QFormLayout(g)
        self.lay.addWidget(g)
        g2 = QGroupBox("Change Password")
        f = QFormLayout(g2)
        self.old, self.new1, self.new2 = QLineEdit(), QLineEdit(), QLineEdit()
        for e in (self.old, self.new1, self.new2):
            e.setEchoMode(QLineEdit.Password)
        f.addRow("Current password:", self.old)
        f.addRow("New password:", self.new1)
        f.addRow("Confirm new password:", self.new2)
        f.addRow(btn("Update Password", "orange", self.change_pw))
        self.lay.addWidget(g2)
        g3 = QGroupBox("Folders")
        h = QHBoxLayout(g3)
        h.addWidget(btn("Open Reports Folder", "blue", lambda: (os.makedirs(REPORT_DIR, exist_ok=True),
                                                                 open_path(REPORT_DIR))))
        h.addWidget(btn("Open Downloads Folder", "blue", lambda: (os.makedirs(DOWNLOAD_DIR, exist_ok=True),
                                                                   open_path(DOWNLOAD_DIR))))
        h.addStretch()
        self.lay.addWidget(g3)
        g5 = QGroupBox("Invoice Folder (this PC)")
        v5 = QVBoxLayout(g5)
        self.inv_folder = InvoiceFolderWidget(main)
        v5.addWidget(self.inv_folder)
        n5 = QLabel("Every invoice PDF is saved here automatically as <b>Name of client_Invoice number.pdf</b> "
                    "(e.g. <i>Sarwan S Gupta_SS-2026-27-0001.pdf</i>). You can choose a shared office folder or a "
                    "network drive so the whole team finds invoices in one place.")
        n5.setWordWrap(True)
        n5.setStyleSheet("color:#8b84a8; font-size:8.5pt;")
        v5.addWidget(n5)
        self.lay.addWidget(g5)
        g4 = QGroupBox("About")
        v = QVBoxLayout(g4)
        a = QLabel(f"<b>{APP_NAME}</b> v{APP_VERSION}<br>"
                   f"Voice calling: {'available' if AUDIO_OK else 'not available (pip install sounddevice)'}<br>"
                   "Data on the LAN is not encrypted; use only on your trusted office network.")
        a.setWordWrap(True)
        v.addWidget(a)
        self.lay.addWidget(g4)
        self.lay.addStretch()

    def refresh(self):
        while self.prof.rowCount():
            self.prof.removeRow(0)
        try:
            u = self.client.call("me")["user"]
        except Exception as e:
            return err(self, e)
        for k, lbl in (("username", "Username"), ("full_name", "Full name"), ("role", "Role"),
                       ("designation", "Designation"), ("department", "Department"), ("email", "E-mail"),
                       ("pan", "PAN"), ("doj", "Date of joining")):
            self.prof.addRow(lbl + ":", QLabel(str(u.get(k) or "-")))
        self.prof.addRow("Monthly salary:", QLabel("Rs. " + inr(u.get("salary"))))

    def change_pw(self):
        if self.new1.text() != self.new2.text():
            return err(self, "New passwords do not match")
        try:
            self.client.call("change_password", old=self.old.text(), new=self.new1.text())
            for e in (self.old, self.new1, self.new2):
                e.clear()
            info(self, "Password changed successfully")
        except Exception as e:
            err(self, e)


# =============================================================================
# 18. GUI : MAIN WINDOW & CONTROLLER
# =============================================================================
class MainWindow(QMainWindow):
    def __init__(self, ctrl, client, data):
        super().__init__()
        self.ctrl, self.client, self.cfg = ctrl, client, ctrl.cfg
        self.user, self.settings = data["user"], data["settings"]
        self.users, self.combos = [], []
        self.leaving = False
        set_firm(self.settings.get("company_name"))
        self.cfg["firm_name"] = FIRM_NAME
        save_cfg(self.cfg)
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(app_icon())
        self.resize(1280, 800)
        self.reload_users(update=False)
        try:
            PDF_BRAND["logo"] = self.client.call("firm_logo")["logo"]
        except Exception:
            pass
        self._click_action = None
        self.media = MediaManager(self)

        root = QWidget()
        hl = QHBoxLayout(root)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(0)
        side = QFrame()
        side.setObjectName("Sidebar")
        side.setFixedWidth(230)
        sl = QVBoxLayout(side)
        sl.setContentsMargins(12, 18, 12, 14)
        b = QLabel(FIRM_NAME)
        b.setObjectName("Brand")
        b.setWordWrap(True)
        self.brand_lbl = b
        sl.addWidget(b)
        s = QLabel(PRODUCT_NAME)
        s.setObjectName("BrandSub")
        sl.addWidget(s)
        sl.addSpacing(10)
        badge = QLabel(f"👤 {html.escape(self.user['full_name'])}<br><span style='font-size:8pt'>"
                       f"{'Manager / Admin' if self.user['role'] == 'admin' else (self.user['designation'] or 'Team Member')}</span>")
        badge.setObjectName("UserBadge")
        sl.addWidget(badge)
        sl.addSpacing(10)
        self.stack = QStackedWidget()
        self.pages, self.nav = {}, {}
        group = QButtonGroup(self)
        group.setExclusive(True)
        items = [("dash", "🏠  Dashboard", DashboardPage), ("chat", "💬  Chat && Calls", ChatPage),
                 ("att", "🕘  Attendance", AttendancePage), ("leave", "🌴  Leave", LeavePage),
                 ("salary", "💵  Salary Slip", SalaryPage), ("reimb", "🧾  Reimbursement", ReimbPage),
                 ("lit", "⚖  Litigation Tracker", LitigationPage), ("tasks", "✅  Task Tracker", TaskPage)]
        if self.user["role"] == "admin":
            items.append(("inv", "🧾  Invoicing", InvoicingPage))
            items.append(("admin", "🛠  Admin", AdminPage))
        items.append(("settings", "⚙  Settings", SettingsPage))
        for key, label, cls in items:
            page = cls(self)
            self.pages[key] = page
            self.stack.addWidget(page)
            nb = QPushButton(label)
            nb.setCheckable(True)
            nb.setCursor(Qt.PointingHandCursor)
            nb.clicked.connect(lambda _=False, k=key: self.go(k))
            group.addButton(nb)
            self.nav[key] = nb
            sl.addWidget(nb)
        sl.addStretch()
        lo = QPushButton("⎋  Logout")
        lo.clicked.connect(self.logout)
        sl.addWidget(lo)
        hl.addWidget(side)
        hl.addWidget(self.stack, 1)
        self.setCentralWidget(root)
        self.status_lbl = QLabel()
        self.statusBar().addWidget(self.status_lbl, 1)
        conn = QLabel(("🖥 Hosting server  •  " if ctrl.server else "") + f"Connected to {client.host}:{PORT}  ")
        self.notif_btn = QPushButton("🔔 Open")
        self.notif_btn.setToolTip("Open the item of the last notification")
        self.notif_btn.setCursor(Qt.PointingHandCursor)
        self.notif_btn.setStyleSheet("QPushButton{background:#ede9fe;color:#5b21b6;border-radius:8px;padding:2px 10px;}")
        self.notif_btn.clicked.connect(self._tray_clicked)
        self.notif_btn.hide()
        self.statusBar().addPermanentWidget(self.notif_btn)
        self.statusBar().addPermanentWidget(conn)
        client.push.connect(self.on_push)
        client.disconnected.connect(self.on_disconnect)
        self.go("dash")
        if data.get("default_password"):
            QTimer.singleShot(400, lambda: QMessageBox.warning(
                self, APP_NAME, "You are using the default admin password (admin@123).\n"
                                "Please change it now from Settings › Change Password."))
        self.users_timer = QTimer(self)
        self.users_timer.setSingleShot(True)
        self.users_timer.timeout.connect(self.reload_users)
        # ---- desktop notifications (Windows notification area / Action Centre) ----
        self.tray = None
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray = QSystemTrayIcon(app_icon(), self)
            self.tray.setToolTip(f"{APP_NAME} - {self.user['full_name']}")
            menu = QMenu()
            menu.addAction("Open Suthar & Suthar", self.bring_front)
            menu.addAction("Check hearing reminders", lambda: self.check_reminders(force=True))
            menu.addAction("Open Litigation Tracker", lambda: (self.bring_front(), self.go("lit")))
            self.tray.setContextMenu(menu)
            self.tray.activated.connect(lambda r: self.bring_front() if r == QSystemTrayIcon.Trigger else None)
            self.tray.messageClicked.connect(self._tray_clicked)
            self.tray.show()
        self._reminded = set()
        self.rem_timer = QTimer(self)
        self.rem_timer.timeout.connect(self.check_reminders)
        self.rem_timer.start(30 * 60 * 1000)          # every 30 minutes
        QTimer.singleShot(2500, self.check_reminders)

    def _tray_clicked(self):
        """Notification clicked: open the case / task / page it is about."""
        self.bring_front()
        act = getattr(self, "_click_action", None)
        if act == "overdue":
            self.show_overdue()
        elif act == "followups":
            self.show_followups()
        elif isinstance(act, tuple) and act[0] == "case":
            self.open_case(act[1])
        elif isinstance(act, tuple) and act[0] == "task":
            self.open_task(act[1])
        elif isinstance(act, tuple) and act[0] == "page" and act[1] in self.pages:
            self.go(act[1])

    def open_case(self, cid):
        self.go("lit")
        CaseFileDialog(self.pages["lit"], cid).exec_()
        self.pages["lit"].refresh()

    def open_task(self, tid):
        self.go("tasks")
        TaskDetailDialog(self.pages["tasks"], tid).exec_()
        self.pages["tasks"].refresh()

    def receipt_pdf(self, rid, ask=False):
        if not rid:
            err(self, "Please select a receipt")
            return None
        try:
            d = self.client.call("receipt_get", id=rid)
        except Exception as e:
            err(self, e)
            return None
        folder = os.path.join(REPORT_DIR, "Receipts")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "Receipt_" + (d["receipt"]["rcpt_no"] or str(rid)).replace("/", "-") + ".pdf")
        if ask:
            path, _ = QFileDialog.getSaveFileName(self, "Save receipt", path, "PDF (*.pdf)")
            if not path:
                return None
        try:
            make_receipt_pdf(d, path)
        except PermissionError:
            err(self, "Cannot write the file. Is it open in a PDF viewer?")
            return None
        except Exception as e:
            err(self, e)
            return None
        if ask:
            open_path(path)
        self._last_receipt = d
        return path

    def send_receipt(self, rid):
        path = self.receipt_pdf(rid)
        if not path:
            return
        d = self._last_receipt
        inv = d["invoice"]
        if not inv.get("client_email") or not inv.get("client_phone"):
            try:
                cm = next((c for c in self.client.call("client_list")
                           if c["name"].lower() == (inv.get("client_name") or "").lower()), None)
            except Exception:
                cm = None
            if cm:
                inv["client_email"] = inv.get("client_email") or cm["email"]
                inv["client_phone"] = inv.get("client_phone") or cm["phone"]
        sub, body = receipt_message(d["receipt"], inv, d["settings"], self.user["full_name"])
        ShareDialog(self, f"Receipt {d['receipt']['rcpt_no']}", inv.get("client_name"), inv.get("client_email"),
                    inv.get("client_phone"), [("Payment receipt", sub, body)], attachment=path).exec_()

    def edit_receipt(self, rid):
        try:
            d = self.client.call("receipt_get", id=rid)
        except Exception as e:
            err(self, e)
            return False
        dlg = ReceiptDialog(self, d["invoice"], d["receipt"])
        if not dlg.exec_():
            return False
        try:
            self.client.call("receipt_update", id=rid, **dlg.value())
        except Exception as e:
            err(self, e)
            return False
        self.after_receipt(rid)
        return True

    def hearing_note(self, cid, proc_id=None):
        try:
            d = self.client.call("case_get", id=cid)
            docs = self.client.call("doc_list", ref_type="case", ref_id=cid)
        except Exception as e:
            return err(self, e)
        proc = next((p for p in d["proceedings"] if p["id"] == proc_id), None) if proc_id else None
        if not proc and not d["case"].get("next_hearing"):
            if QMessageBox.question(self, APP_NAME, "No next hearing date is set for this case. Prepare a general "
                                                    "hearing brief anyway?") != QMessageBox.Yes:
                return
        folder = os.path.join(REPORT_DIR, "Hearing Notes")
        os.makedirs(folder, exist_ok=True)
        c = d["case"]
        tag = f"{proc['pdate']}_{proc['ptype']}" if proc else f"Brief_{c.get('next_hearing') or date.today().isoformat()}"
        path = os.path.join(folder, safe_filename(f"Case{c['id']}_{c['client_name']}_{tag}.pdf").replace(" ", "_"))
        path, _ = QFileDialog.getSaveFileName(self, "Save hearing note", path, "PDF (*.pdf)")
        if not path:
            return
        try:
            make_hearing_note_pdf(d, path, proc, docs)
            open_path(path)
            self.status("Hearing note saved: " + path)
        except PermissionError:
            err(self, "Cannot write the file. Is it open in a PDF viewer?")
        except Exception as e:
            err(self, e)

    def after_receipt(self, rid):
        path = self.receipt_pdf(rid)
        if not path:
            return
        ans = QMessageBox.question(self, APP_NAME, f"Payment receipt {self._last_receipt['receipt']['rcpt_no']} "
                                                   "has been generated (PDF).\n\nSend it to the client by "
                                                   "e-mail / WhatsApp now?\n(No = just open the PDF)")
        if ans == QMessageBox.Yes:
            self.send_receipt(rid)
        else:
            open_path(path)

    def show_overdue(self, rows=None):
        if rows is None:
            try:
                rows = self.client.call("invoice_list", overdue=True)
            except Exception as e:
                return err(self, e)
        if not rows:
            return info(self, f"No invoice is outstanding for more than {INVOICE_OVERDUE_DAYS} days.")
        old = getattr(self, "_overdue_dlg", None)
        if old is not None and old.isVisible():
            old.close()
        self._overdue_dlg = OverdueInvoicesDialog(self, rows)
        self._overdue_dlg.show()

    def refresh_invoice_folder(self):
        for p in ("settings", "inv"):
            pg = self.pages.get(p)
            for attr in ("inv_folder", "folder"):
                w = getattr(pg, attr, None) if pg else None
                if isinstance(w, InvoiceFolderWidget):
                    w.show_path()

    def apply_firm_name(self, name):
        set_firm(name)
        self.cfg["firm_name"] = FIRM_NAME
        save_cfg(self.cfg)
        self.setWindowTitle(APP_NAME)
        self.brand_lbl.setText(FIRM_NAME)
        QApplication.instance().setApplicationName(APP_NAME)
        if self.tray:
            self.tray.setToolTip(f"{APP_NAME} - {self.user['full_name']}")

    def open_firm_master(self):
        self.go("inv")
        self.pages["inv"].tabs.setCurrentIndex(3)

    def invoice_pdf(self, iid, ask=True):
        """Create invoice PDF; returns path. ask=False saves straight into Reports/Invoices."""
        if not iid:
            err(self, "Please select an invoice")
            return None
        try:
            d = self.client.call("invoice_get", id=iid)
        except Exception as e:
            err(self, e)
            return None
        folder = invoice_folder(self.cfg)
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            err(self, f"Invoice folder is not available:\n{folder}\n\n{e}\n\nThe default folder will be used. "
                      "You can change the folder in Settings › Invoice Folder.")
            folder = os.path.join(REPORT_DIR, "Invoices")
            os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, invoice_filename(d["invoice"]))
        if ask:
            path, _ = QFileDialog.getSaveFileName(self, "Save invoice", path, "PDF (*.pdf)")
            if not path:
                return None
        try:
            make_invoice_pdf(d, path)
        except PermissionError:
            err(self, "Cannot write the file. Is it open in a PDF viewer?")
            return None
        except Exception as e:
            err(self, e)
            return None
        if ask:
            open_path(path)
        self._last_invoice = d
        return path

    def send_invoice(self, iid, reminder=False):
        path = self.invoice_pdf(iid, ask=False)
        if not path:
            return
        d = self._last_invoice
        inv, s = d["invoice"], d["settings"]
        if not inv.get("client_email") or not inv.get("client_phone"):     # fall back to Client Master
            try:
                cm = next((c for c in self.client.call("client_list")
                           if c["name"].lower() == (inv.get("client_name") or "").lower()), None)
            except Exception:
                cm = None
            if cm:
                inv["client_email"] = inv.get("client_email") or cm["email"]
                inv["client_phone"] = inv.get("client_phone") or cm["phone"]
        name = self.user["full_name"]
        t_new = ("Invoice",) + invoice_message(inv, s, False, name)
        t_rem = ("Payment reminder",) + invoice_message(inv, s, True, name)
        tpls = [t_rem, t_new] if reminder else [t_new, t_rem]
        ShareDialog(self, f"Invoice {inv['inv_no']}", inv.get("client_name"), inv.get("client_email"),
                    inv.get("client_phone"), tpls, attachment=path).exec_()

    def send_payment_reminder(self, iid):
        """Send gentle / follow-up / final reminder (E-mail / WhatsApp) with the PDF and log it as a follow-up."""
        try:
            row = next(x for x in self.client.call("invoice_list") if x["id"] == iid)
        except Exception as e:
            err(self, e)
            return None
        if row["outstanding"] <= 0.009:
            info(self, "This invoice is fully paid — no reminder needed.")
            return None
        path = self.invoice_pdf(iid, ask=False)
        if not path:
            return None
        inv, s = self._last_invoice["invoice"], self._last_invoice["settings"]
        inv.update(fu_promise=row["fu_promise"], client_email=row.get("client_email") or inv.get("client_email"),
                   client_phone=row.get("client_phone") or inv.get("client_phone"))
        if not inv.get("client_email") or not inv.get("client_phone"):
            try:
                cm = next((c for c in self.client.call("client_list")
                           if c["name"].lower() == (inv.get("client_name") or "").lower()), None)
            except Exception:
                cm = None
            if cm:
                inv["client_email"] = inv.get("client_email") or cm["email"]
                inv["client_phone"] = inv.get("client_phone") or cm["phone"]
        level = min(row["fu_count"], 2)          # 1st gentle -> 2nd firm -> final (and after)
        dlg = ShareDialog(self, f"Payment follow-up — {inv['inv_no']} (follow-up #{row['fu_count'] + 1})",
                          inv.get("client_name"), inv.get("client_email"), inv.get("client_phone"),
                          reminder_templates(inv, s, self.user["full_name"]), attachment=path, start=level)
        dlg.exec_()
        if dlg.sent:
            mode, label = dlg.sent
            try:
                self.client.call("followup_add", invoice_id=iid, fdate=date.today().isoformat(), mode=mode,
                                 remarks=f"{label} sent by {mode}",
                                 next_date=(date.today() + timedelta(days=FU_GAP_DAYS)).isoformat())
                self.status(f"Reminder logged — next follow-up on {(date.today() + timedelta(days=FU_GAP_DAYS)):%d-%m-%Y}")
            except Exception as e:
                err(self, e)
        return dlg.sent

    def show_followups(self, rows=None):
        if rows is None:
            try:
                rows = self.client.call("invoice_list", followup_due=True)
            except Exception as e:
                return err(self, e)
        if not rows:
            return info(self, "No payment follow-up is due today.")
        OverdueInvoicesDialog(self, rows, heading=f"{len(rows)} payment follow-up(s) due").exec_()

    def generate_invoice(self, billing=None):
        try:
            clients = self.client.call("client_list")
            st = self.client.call("firm_get")
        except Exception as e:
            err(self, e)
            return False
        st.pop("firm_logo", None)
        dlg = ManualInvoiceDialog(self, clients, st, billing)
        if not dlg.exec_():
            return False
        try:
            r = self.client.call("invoice_manual", **dlg.value())
        except Exception as e:
            err(self, e)
            return False
        ans = QMessageBox.question(self, APP_NAME, f"Invoice {r['inv_no']} generated for Rs. {inr(r['total'])}.\n\n"
                                                   "Send it to the client now by e-mail / WhatsApp?\n"
                                                   "(No = just open the PDF)")
        if ans == QMessageBox.Yes:
            self.send_invoice(r["id"])
        else:
            p = self.invoice_pdf(r["id"], ask=False)
            if p:
                open_path(p)
        return True

    def share_case(self, c):
        s = self.settings
        sig = firm_signature(s, self.user["full_name"])
        nh = dmy(c["next_hearing"]) if c.get("next_hearing") else "not yet fixed"
        details = (f"Matter: {c['law']} - {c.get('forum') or ''}\nPeriod: {c.get('period') or '-'}\n"
                   f"Section: {c.get('section') or '-'}  Ref: {c.get('ref_no') or '-'}\nIssue: {c.get('nature') or '-'}\n"
                   f"Demand: Rs. {inr(c.get('demand'))}\nStatus: {c['status']}\nNext hearing: {nh}"
                   + (f" ({c['next_purpose']})" if c.get("next_purpose") else ""))
        tpls = [("Hearing reminder", f"Hearing reminder - {c['client_name']} - {c['law']} - {nh}",
                 f"Dear {c['client_name']},\n\nThis is to remind you that the hearing in your {c['law']} matter "
                 f"({c.get('period') or ''}) before {c.get('forum') or 'the authority'} is fixed on {nh}"
                 + (f" for {c['next_purpose']}" if c.get("next_purpose") else "") +
                 ".\n\nKindly share the required documents / details at least 2 days before the hearing.\n\n" + sig),
                ("Case status update", f"Status update - {c['client_name']} - {c['law']} matter",
                 f"Dear {c['client_name']},\n\nPlease find below the current status of your matter:\n\n"
                 f"{details}\n\n{sig}")]
        u = next((x for x in self.users if x["username"] == c.get("assigned_to")), None)
        if u:
            tpls.append(("To team member (handler)", f"Hearing on {nh} - {c['client_name']} ({c['law']})",
                         f"Dear {u['full_name']},\n\nPlease prepare for the hearing:\n\nClient: {c['client_name']}\n"
                         f"{details}\n\n{sig}"))
        ShareDialog(self, f"Case #{c['id']} - {c['client_name']}", c["client_name"], c.get("client_email"),
                    c.get("client_phone"), tpls).exec_()

    def bring_front(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def desktop_notify(self, title, text, warn=False, action=None):
        self._click_action = action
        self.notif_btn.setVisible(action is not None)
        self.status("🔔 " + text.replace("\n", "  |  "))
        if self.tray:
            self.tray.showMessage(title, text, QSystemTrayIcon.Warning if warn else QSystemTrayIcon.Information,
                                  15000)
        QApplication.alert(self)

    def check_reminders(self, force=False):
        try:
            d = self.client.call("reminders")
        except Exception:
            return
        t = date.today()
        # ---- task due alerts (red: due within 3 days / overdue) ----
        tk = d.get("tasks", [])
        tkey = ("tasks", t.isoformat(), tuple((x["id"], x["due_date"]) for x in tk))
        if tk and (force or tkey not in self._reminded):
            self._reminded.add(tkey)
            tl = []
            for x in tk[:6]:
                dl = (date.fromisoformat(x["due_date"]) - t).days
                when = f"OVERDUE {-dl}d" if dl < 0 else ("TODAY" if dl == 0 else f"in {dl} day(s)")
                tl.append(("🔥 " if x["urgent"] else "") + f"{x['title']}" + (f" ({x['client']})" if x["client"] else "")
                          + f" — {when}")
            self.desktop_notify(f"{APP_NAME} - {len(tk)} task(s) due soon", "\n".join(tl), warn=True,
                                action=("task", tk[0]["id"]) if len(tk) == 1 else ("page", "tasks"))
        # ---- invoices outstanding > 30 days (Manager): notification + pop-up with e-mail / WhatsApp options
        inv = d.get("invoices", [])
        ikey = ("inv", t.isoformat(), tuple(x["id"] for x in inv))
        if inv and (force or ikey not in self._reminded):
            self._reminded.add(ikey)
            tot = sum(x["outstanding"] for x in inv)
            lines = [f"{x['client_name']}: {x['inv_no']} — Rs. {inr(x['outstanding'])} ({x['age_days']} days)"
                     for x in inv[:5]]
            self._click_action = "overdue"
            self.desktop_notify(f"{APP_NAME} - {len(inv)} invoice(s) outstanding > {INVOICE_OVERDUE_DAYS} days "
                                f"(Rs. {inr(tot)})", "\n".join(lines) + "\nClick to send reminders by e-mail / WhatsApp.",
                                warn=True, action="overdue")
            QTimer.singleShot(800, lambda r=inv: self.show_overdue(r))
        # ---- payment follow-ups due today (next follow-up / promised date reached) ----
        fus = d.get("followups", [])
        fkey = ("fu", t.isoformat(), tuple(x["id"] for x in fus))
        if fus and (force or fkey not in self._reminded):
            self._reminded.add(fkey)
            lines = [f"{x['client_name']}: {x['inv_no']} — Rs. {inr(x['outstanding'])}"
                     + (f" (promised {dmy(x['fu_promise'])})" if x["fu_promise"] else "") for x in fus[:5]]
            self.desktop_notify(f"{APP_NAME} - {len(fus)} payment follow-up(s) due today",
                                "\n".join(lines) + "\nClick to follow up by call / e-mail / WhatsApp.", warn=True,
                                action="followups")
        # ---- month-end payout reminder for the Manager (last 2 days of month / first 3 days) ----
        if self.user["role"] == "admin":
            dim = calendar.monthrange(t.year, t.month)[1]
            if t.day >= dim - 1 or t.day <= 3:
                pm = t if t.day >= dim - 1 else (t.replace(day=1) - timedelta(days=1))
                pkey = ("payout", t.isoformat())
                if force or pkey not in self._reminded:
                    self._reminded.add(pkey)
                    self.desktop_notify(APP_NAME, f"Month-end: generate the payout statement (salary + reimbursement)"
                                                  f" for {pm:%B %Y} — Salary Slip & Payout page.",
                                        action=("page", "salary"))
        key = (t.isoformat(), tuple((r["id"], r["next_hearing"]) for r in d["upcoming"] + d["overdue"]))
        if not force and key in self._reminded:
            return
        self._reminded.add(key)
        lines = []
        for r in d["upcoming"]:
            nd = date.fromisoformat(r["next_hearing"])
            when = "TODAY" if nd == t else ("Tomorrow" if nd == t + timedelta(days=1) else nd.strftime("%d-%b (%a)"))
            lines.append(f"{when}: {r['client_name']} ({r['law']}) - {r['forum'] or ''} {r['next_purpose'] or ''}")
        if d["overdue"]:
            lines.append(f"{len(d['overdue'])} case(s) have a past hearing date — please record the outcome.")
        if lines:
            self.desktop_notify(f"{APP_NAME} - Hearing reminder ({len(d['upcoming'])} in next {d['days']} days)",
                                "\n".join(lines[:6]) + ("\n..." if len(lines) > 6 else ""), warn=True,
                                action=("case", d["upcoming"][0]["id"]) if len(d["upcoming"]) == 1
                                else ("page", "lit"))
        elif force:
            self.desktop_notify(APP_NAME, f"No hearings in the next {d['days']} day(s).", action=self._click_action)

    # -------------------------------------------------------------- helpers
    def go(self, key):
        self.nav[key].setChecked(True)
        self.stack.setCurrentWidget(self.pages[key])
        try:
            self.pages[key].refresh()
        except Exception as e:
            self.status(str(e))

    def status(self, text):
        self.status_lbl.setText(f"{datetime.now().strftime('%H:%M')}  {text}")

    def name(self, username):
        for u in self.users:
            if u["username"] == username:
                return u["full_name"]
        return username

    def reload_users(self, update=True):
        try:
            self.users = self.client.call("users")
        except Exception:
            return
        if update:
            for cb, inc in self.combos:
                self.fill_user_combo(cb, inc)
            self.pages["chat"].refresh_users()

    def fill_user_combo(self, cb, include_all):
        cur = cb.currentData()
        cb.blockSignals(True)
        cb.clear()
        if self.user["role"] == "admin":
            if include_all:
                cb.addItem("All team members", "ALL")
            for u in self.users:
                cb.addItem(f"{u['full_name']} ({u['username']})", u["username"])
            i = cb.findData(cur if cur else ("ALL" if include_all else self.user["username"]))
            cb.setCurrentIndex(i if i >= 0 else 0)
        else:
            cb.addItem(self.user["full_name"], self.user["username"])
        cb.blockSignals(False)

    def download_file(self, fid, progress_label=None):
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        if progress_label:
            progress_label.setText("Downloading...")
        self.status("Downloading file...")

        def work():
            meta, data = self.client.request("file_download", timeout=900, file_id=fid)
            name = safe_filename(meta["name"])
            base, ext = os.path.splitext(name)
            path, n = os.path.join(DOWNLOAD_DIR, name), 1
            while os.path.exists(path):
                path = os.path.join(DOWNLOAD_DIR, f"{base} ({n}){ext}")
                n += 1
            with open(path, "wb") as f:
                f.write(data)
            return path

        def done(path, e):
            if progress_label:
                progress_label.setText("")
            if e:
                return err(self, e, "Download failed")
            self.status("Saved to " + path)
            if QMessageBox.question(self, APP_NAME, f"File saved to:\n{path}\n\nOpen it now?") == QMessageBox.Yes:
                open_path(path)
        run_task(work, done)

    # --------------------------------------------------------------- events
    def on_push(self, h, payload):
        ev = h.get("event")
        if ev == "chat":
            self.pages["chat"].on_message(h["msg"])
        elif ev == "presence":
            self.users_timer.start(300)
        elif ev == "relay":
            self.media.handle(h, payload)
        elif ev == "task":
            if self.stack.currentWidget() is self.pages["tasks"] and QApplication.activeModalWidget() is None:
                self.pages["tasks"].refresh()
        elif ev == "lit":
            if self.stack.currentWidget() is self.pages["lit"] and QApplication.activeModalWidget() is None:
                self.pages["lit"].refresh()
        elif ev == "notify":
            ref = h.get("ref") or {}
            act = ((ref.get("kind"), ref.get("id")) if ref.get("kind") in ("case", "task")
                   else (("page", ref.get("page")) if ref.get("kind") == "page" else None))
            self.desktop_notify(APP_NAME, h.get("text", "") + ("\nClick to open." if act else ""), action=act)
            QApplication.beep()
            cur = self.stack.currentWidget()
            if hasattr(cur, "refresh") and cur is not self.pages["chat"]:
                try:
                    cur.refresh()
                except Exception:
                    pass
        elif ev == "kicked":
            self.leaving = True
            info(self, "You have logged in from another computer. This session is closed.")
            self.ctrl.logout()

    def on_disconnect(self, why):
        if self.leaving:
            return
        self.leaving = True
        self.media.stop_all()
        QMessageBox.critical(self, APP_NAME, "Connection to the server was lost.\n"
                                             "Please make sure the host PC is running, then log in again.")
        self.ctrl.logout()

    def logout(self):
        self.leaving = True
        self.ctrl.logout()

    def closeEvent(self, ev):
        if not self.leaving:
            msg = "Exit Suthar & Suthar?"
            if self.ctrl.server:
                msg += "\n\nThis PC is hosting the server — team members will be disconnected."
            if QMessageBox.question(self, APP_NAME, msg) != QMessageBox.Yes:
                ev.ignore()
                return
            self.leaving = True
            self.media.stop_all()
            if self.tray:
                self.tray.hide()
            self.client.close()
            ev.accept()
            QTimer.singleShot(0, self.ctrl.shutdown)
            return
        ev.accept()


class Controller(QObject):
    def __init__(self, app):
        super().__init__()
        self.app, self.server, self.win = app, None, None
        self.cfg = load_cfg()

    def start(self):
        dlg = LoginDialog(self.cfg, self.server)
        ok = dlg.exec_() == QDialog.Accepted
        self.server = dlg.server or self.server
        if not ok:
            return self.shutdown()
        self.win = MainWindow(self, dlg.client, dlg.data)
        self.win.show()

    def logout(self):
        w = self.win
        self.win = None
        if w:
            w.leaving = True
            w.media.stop_all()
            w.rem_timer.stop()
            if w.tray:
                w.tray.hide()
            w.client.close()
            w.hide()
            w.deleteLater()
        QTimer.singleShot(0, self.start)

    def shutdown(self):
        if self.server:
            self.server.stop()
            self.server = None
        self.app.quit()


def run_headless_server():
    srv = Server()
    srv.start()
    print(f"{APP_NAME} server running on {', '.join(local_ips())} port {PORT}. Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        srv.stop()


def main():
    if "--server" in sys.argv:
        return run_headless_server()
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    app.setStyleSheet(QSS)
    app.setWindowIcon(app_icon())
    app.setQuitOnLastWindowClosed(False)
    ctrl = Controller(app)
    app._ctrl = ctrl   # keep reference
    QTimer.singleShot(0, ctrl.start)
    app.exec_()
    if ctrl.server:
        ctrl.server.stop()


if __name__ == "__main__":
    main()
