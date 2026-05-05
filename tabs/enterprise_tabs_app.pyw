import hashlib
import json
import locale
import os
import platform
import subprocess
import shutil
import sqlite3
import threading
import time
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest

from openpyxl import Workbook

try:
    import pyautogui
except ImportError:
    pyautogui = None

# =========================
# CONFIG
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "app.db")
INVENTORY_DEBUG_LOG = os.path.join(BASE_DIR, "inventory_debug.log")

APP_TITLE = "AikenJobs Enterprise Dashboard"
APP_VERSION = "2.0.0"

# Professional color palette
BG = "#0f172a"
SURFACE = "#111827"
CARD = "#1f2937"
BORDER = "#334155"
TEXT = "#e5e7eb"
MUTED = "#94a3b8"
PRIMARY = "#2563eb"
PRIMARY_DARK = "#1d4ed8"
SUCCESS = "#16a34a"
DANGER = "#dc2626"
WARNING = "#f59e0b"
ENTRY_BG = "#0b1220"

DEFAULT_SETTINGS = {
    "system_name": APP_TITLE,
    "default_export_folder": BASE_DIR,
    "keyboard_delay": "80",
    "auto_refresh": "1",
    "network_server_url": "http://127.0.0.1:8765",
    "workstation_name": os.environ.get("COMPUTERNAME", "CLIENT-1"),
}

AUDITOR_FIELDS = [
    "Product type",
    "Manufacturer",
    "Model",
    "Chassis type",
    "Part number",
    "Serial number",
    "Processor",
    "RAM size",
    "Disk type",
    "Disk size",
    "Optical drive",
    "Video model",
    "LCD size",
    "LCD resolution",
    "LCD type",
    "Keyboard layout",
    "OS COA",
    "Grade",
]

AUDITOR_SAMPLE_VALUES = {
    "Product type": "LAPTOP",
    "Manufacturer": "HP",
    "Model": "ELITEBOOK 8540W",
    "Chassis type": "LAPTOP",
    "Part number": "NU515AV",
    "Serial number": "CND0511LHM",
    "Processor": "Intel Core i5 M 540 2.53 GHz",
    "RAM size": "2 GB",
    "Disk type": "SATA SSD",
    "Disk size": "180 GB",
    "Optical drive": "DVD/RW",
    "Video model": "NVIDIA QUADRO FX 880M",
    "LCD size": "15.5\"",
    "LCD resolution": "1600x900",
    "LCD type": "Unknown",
    "Keyboard layout": "",
    "OS COA": "",
    "Grade": "",
}

LCD_TYPE_OPTIONS = [
    "",
    "Unknown",
    "IPS LCD",
    "TN LCD",
    "VA LCD",
    "OLED",
    "AMOLED",
    "Mini LED",
    "LCD",
]

OS_COA_OPTIONS = [
    "",
    "Windows 11 Pro",
    "Windows 11 Home",
    "Windows 10 Pro",
    "Windows 10 Home",
    "Windows 8.1 Pro",
    "Windows 8.1 Home",
    "Windows 7 Professional",
    "Windows 7 Home Premium",
    "Windows Vista Business",
    "Windows XP Professional",
    "Linux / No COA",
    "No COA",
]

FIELD_COUNT = len(AUDITOR_FIELDS)
LCD_TYPE_INDEX = AUDITOR_FIELDS.index("LCD type")
KEYBOARD_SKIP_FIELDS = {"LCD type"}
RECORD_SELECT_COLUMNS = "id, xenoglossi, timestamp, username, password, dropdown, checkbox, inputs"


# =========================
# DATABASE
# =========================
def get_conn():
    return sqlite3.connect(DB_NAME)


def hash_password(pw):
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()


def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute(
        """CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            xenoglossi TEXT NOT NULL DEFAULT '',
            timestamp TEXT,
            username TEXT,
            username TEXT,
            password TEXT,
            dropdown TEXT,
            checkbox INTEGER,
            inputs TEXT
        )"""
    )

    c.execute("PRAGMA table_info(records)")
    record_columns = [row[1] for row in c.fetchall()]
    if "xenoglossi" not in record_columns:
        c.execute("ALTER TABLE records ADD COLUMN xenoglossi TEXT NOT NULL DEFAULT ''")

    c.execute(
        "UPDATE records SET xenoglossi='LEGACY-' || id WHERE xenoglossi IS NULL OR TRIM(xenoglossi)=''"
    )

    c.execute(
        """CREATE TABLE IF NOT EXISTS app_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'User',
            status TEXT NOT NULL DEFAULT 'Active',
            created_at TEXT,
            last_login TEXT
        )"""
    )

    c.execute(
        """CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )"""
    )

    c.execute(
        """CREATE TABLE IF NOT EXISTS app_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            username TEXT,
            action TEXT,
            details TEXT
        )"""
    )

    c.execute("SELECT id FROM app_users WHERE username=?", ("admin",))
    if c.fetchone() is None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.execute(
            "INSERT INTO app_users (username, password, role, status, created_at) VALUES (?, ?, ?, ?, ?)",
            ("admin", hash_password("1234"), "Admin", "Active", now),
        )

    for key, value in DEFAULT_SETTINGS.items():
        c.execute(
            "INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )

    conn.commit()
    conn.close()


def log_action(action, details="", username=None):
    user = username or current_user or "system"
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO app_logs (timestamp, username, action, details) VALUES (?, ?, ?, ?)",
        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user, action, details),
    )
    conn.commit()
    conn.close()


def get_setting(key, default=""):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT value FROM app_settings WHERE key=?", (key,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else default


def set_setting(key, value):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)",
        (key, str(value)),
    )
    conn.commit()
    conn.close()


def selected_tree_id(tree):
    item = tree.focus()
    if not item:
        return None
    values = tree.item(item).get("values", [])
    return values[0] if values else None


# =========================
# APP ROOT
# =========================
root = tk.Tk()
root.title(APP_TITLE)
root.configure(bg=BG)


def maximize_window(window):
    try:
        window.state("zoomed")
    except tk.TclError:
        try:
            window.attributes("-zoomed", True)
        except tk.TclError:
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            window.geometry(f"{screen_width}x{screen_height}+0+0")


maximize_window(root)
SCREEN_RESOLUTION_CACHE = f"{root.winfo_screenwidth()}x{root.winfo_screenheight()}"

style = ttk.Style()
style.theme_use("clam")


# =========================
# STYLES
# =========================
style.configure("TNotebook", background=BG, borderwidth=0)
style.configure("TNotebook.Tab", padding=(18, 10), font=("Segoe UI", 10, "bold"))

style.configure("TFrame", background=BG)
style.configure("Card.TFrame", background=CARD)

style.configure("TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 10))
style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))
style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
style.configure("Header.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 20, "bold"))
style.configure("SubHeader.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
style.configure("CardHeader.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 13, "bold"))
style.configure("KpiValue.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 22, "bold"))
style.configure("KpiTitle.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9, "bold"))

style.configure(
    "TEntry",
    fieldbackground=ENTRY_BG,
    background=ENTRY_BG,
    foreground=TEXT,
    insertcolor=TEXT,
    bordercolor=BORDER,
    lightcolor=BORDER,
    darkcolor=BORDER,
    padding=7,
)

style.configure(
    "TCombobox",
    fieldbackground=ENTRY_BG,
    background=ENTRY_BG,
    foreground=TEXT,
    arrowcolor=TEXT,
    bordercolor=BORDER,
    padding=6,
)

style.configure("TCheckbutton", background=BG, foreground=TEXT, font=("Segoe UI", 10))
style.configure("Card.TCheckbutton", background=CARD, foreground=TEXT, font=("Segoe UI", 10))

style.configure(
    "TButton",
    font=("Segoe UI", 10, "bold"),
    padding=(12, 7),
    background=PRIMARY,
    foreground="white",
    borderwidth=0,
)
style.map("TButton", background=[("active", PRIMARY_DARK)])

style.configure(
    "Danger.TButton",
    font=("Segoe UI", 10, "bold"),
    padding=(12, 7),
    background=DANGER,
    foreground="white",
)
style.map("Danger.TButton", background=[("active", "#b91c1c")])

style.configure(
    "Success.TButton",
    font=("Segoe UI", 10, "bold"),
    padding=(12, 7),
    background=SUCCESS,
    foreground="white",
)
style.map("Success.TButton", background=[("active", "#15803d")])

style.configure(
    "Treeview",
    rowheight=30,
    background=ENTRY_BG,
    foreground=TEXT,
    fieldbackground=ENTRY_BG,
    bordercolor=BORDER,
)
style.configure(
    "Treeview.Heading",
    background=CARD,
    foreground=TEXT,
    font=("Segoe UI", 10, "bold"),
)
style.map("Treeview", background=[("selected", PRIMARY)], foreground=[("selected", "white")])


# =========================
# GLOBAL STATE
# =========================
is_logged_in = False
current_user = None
current_role = None
selected_id = None
selected_user_id = None
network_jobs = {}
network_clients = {}


# =========================
# HELPERS
# =========================
def make_card(parent, title=None, padx=24, pady=20):
    card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
    if title:
        ttk.Label(card, text=title, style="CardHeader.TLabel").pack(anchor="w", padx=padx, pady=(pady, 8))
    return card


def create_scrollable_area(parent, bg=BG):
    container = tk.Frame(parent, bg=bg)
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container, bg=bg, highlightthickness=0)
    scrollbar_y = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollbar_x = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)

    scrollable_frame = tk.Frame(canvas, bg=bg)
    window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    def update_scrollregion(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def resize_inner_frame(event):
        canvas.itemconfig(window_id, width=event.width)

    def on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    scrollable_frame.bind("<Configure>", update_scrollregion)
    canvas.bind("<Configure>", resize_inner_frame)
    canvas.bind_all("<MouseWheel>", on_mousewheel)

    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=1, sticky="ns")
    scrollbar_x.grid(row=1, column=0, sticky="ew")

    container.rowconfigure(0, weight=1)
    container.columnconfigure(0, weight=1)

    return scrollable_frame


def clear_tree(tree):
    tree.delete(*tree.get_children())


def clear_field(widget):
    if isinstance(widget, ttk.Combobox):
        widget.set("")
    else:
        widget.delete(0, tk.END)


def set_field(widget, value):
    clear_field(widget)
    value = "" if value is None else value
    if isinstance(widget, ttk.Combobox):
        widget.set(str(value))
    else:
        widget.insert(0, str(value))


def get_field(widget):
    return widget.get().strip()


def normalize_input_values(values):
    values = list(values)
    if len(values) == FIELD_COUNT - 1:
        values.insert(LCD_TYPE_INDEX, "")
    values += [""] * (FIELD_COUNT - len(values))
    return values[:FIELD_COUNT]


def set_protected_tabs(state):
    for tab_index in range(1, notebook.index("end")):
        notebook.tab(tab_index, state=state)


def require_login():
    if not is_logged_in:
        messagebox.showwarning("Access denied", "Please login first")
        return False
    return True


def require_admin():
    if not require_login():
        return False
    if current_role != "Admin":
        messagebox.showwarning("Access denied", "Admin role required")
        return False
    return True


def parse_iso_date(value, field_name, end_of_day=False):
    value = value.strip()
    if not value:
        return None
    try:
        parsed = datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        messagebox.showerror("Date filter", f"{field_name} must be YYYY-MM-DD")
        return False
    suffix = "23:59:59" if end_of_day else "00:00:00"
    return parsed.strftime("%Y-%m-%d") + " " + suffix


def get_default_export_folder():
    folder = get_setting("default_export_folder", BASE_DIR)
    return folder if os.path.isdir(folder) else BASE_DIR


def ask_save_path(default_ext, filetypes, title, initialfile):
    return filedialog.asksaveasfilename(
        defaultextension=default_ext,
        filetypes=filetypes,
        title=title,
        initialdir=get_default_export_folder(),
        initialfile=initialfile,
    )


def refresh_all_data():
    load_data(search_var.get() if "search_var" in globals() else "")
    refresh_analytics()
    build_report(show_message=False)
    load_users()
    load_logs()


# =========================
# MAIN LAYOUT
# =========================
header = tk.Frame(root, bg=BG)
header.pack(fill="x", padx=18, pady=(14, 8))

left_header = tk.Frame(header, bg=BG)
left_header.pack(side="left")

title_label = ttk.Label(left_header, text=APP_TITLE, style="Header.TLabel")
title_label.pack(anchor="w")
ttk.Label(left_header, text="Internal administration dashboard", style="SubHeader.TLabel").pack(anchor="w")

right_header = tk.Frame(header, bg=BG)
right_header.pack(side="right")

session_label = ttk.Label(right_header, text="Status: Logged out", style="Muted.TLabel")
session_label.pack(side="left", padx=10)

logout_btn = ttk.Button(right_header, text="Logout", command=lambda: logout(), style="Danger.TButton")

action_bar = tk.Frame(root, bg="#020617", height=58, highlightbackground=BORDER, highlightthickness=1)
action_bar.pack(fill="x", padx=18, pady=(0, 10))
action_bar.pack_propagate(False)

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=18, pady=8)


# =========================
# LOGIN LOGIC
# =========================
def do_login():
    global is_logged_in, current_user, current_role

    user = login_username_entry.get().strip()
    pw = login_password_entry.get().strip()

    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "SELECT id, username, role, status FROM app_users WHERE username=? AND password=?",
        (user, hash_password(pw)),
    )
    row = c.fetchone()

    if not row:
        conn.close()
        messagebox.showerror("Login", "Wrong credentials")
        return

    if row[3] != "Active":
        conn.close()
        messagebox.showerror("Login", "This account is disabled")
        return

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("UPDATE app_users SET last_login=? WHERE id=?", (now, row[0]))
    conn.commit()
    conn.close()

    is_logged_in = True
    current_user = row[1]
    current_role = row[2]

    set_protected_tabs("normal")
    notebook.select(form_frame)
    logout_btn.pack(side="left", padx=10)
    session_label.config(text=f"Status: {current_user} ({current_role})")
    log_action("LOGIN", "User logged in", current_user)
    refresh_all_data()
    messagebox.showinfo("Login", "Login successful")


def logout():
    global is_logged_in, current_user, current_role

    if current_user:
        log_action("LOGOUT", "User logged out", current_user)

    is_logged_in = False
    current_user = None
    current_role = None
    clear_form()
    set_protected_tabs("hidden")
    notebook.select(login_frame)
    login_username_entry.delete(0, tk.END)
    login_password_entry.delete(0, tk.END)
    logout_btn.pack_forget()
    session_label.config(text="Status: Logged out")
    messagebox.showinfo("Logout", "Logged out")


# =========================
# LOGIN TAB
# =========================
login_frame = tk.Frame(notebook, bg=BG)
notebook.add(login_frame, text="Login")

login_card = make_card(login_frame, "Secure Login")
login_card.pack(expand=True)

login_body = tk.Frame(login_card, bg=CARD)
login_body.pack(padx=34, pady=(8, 30))

ttk.Label(login_body, text="Username", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=12, pady=8)
login_username_entry = ttk.Entry(login_body, width=34)
login_username_entry.grid(row=0, column=1, pady=8)

ttk.Label(login_body, text="Password", style="Card.TLabel").grid(row=1, column=0, sticky="e", padx=12, pady=8)
login_password_entry = ttk.Entry(login_body, show="*", width=34)
login_password_entry.grid(row=1, column=1, pady=8)

ttk.Button(login_body, text="Login", command=do_login, style="Success.TButton").grid(row=2, column=0, columnspan=2, pady=18)
ttk.Label(login_body, text="Demo credentials: admin / 1234", style="Card.TLabel").grid(row=3, column=0, columnspan=2)

login_password_entry.bind("<Return>", lambda event: do_login())


# =========================
# FORM TAB
# =========================
form_frame = tk.Frame(notebook, bg=BG)
notebook.add(form_frame, text="Records", state="hidden")

form_scroll = create_scrollable_area(form_frame, bg=BG)

form_header = tk.Frame(form_scroll, bg=BG)
form_header.pack(fill="x", padx=10, pady=(12, 4))

ttk.Label(form_header, text="Record Form", style="Header.TLabel").pack(anchor="w")
ttk.Label(form_header, text="Create or update database records", style="SubHeader.TLabel").pack(anchor="w")

form_card = make_card(form_scroll, "Data Entry")
form_card.pack(fill="x", padx=10, pady=10)

identity_container = tk.Frame(form_card, bg=CARD)
identity_container.pack(padx=24, pady=(8, 8), fill="x")

ttk.Label(identity_container, text="Ξενόγλωσση", style="Card.TLabel").grid(
    row=0, column=0, sticky="e", padx=(8, 10), pady=7
)
record_xenoglossi_entry = ttk.Entry(identity_container, width=34)
record_xenoglossi_entry.grid(row=0, column=1, sticky="ew", padx=(0, 28), pady=7)

ttk.Label(identity_container, text="Record Username", style="Card.TLabel").grid(
    row=0, column=2, sticky="e", padx=(8, 10), pady=7
)
record_username_entry = ttk.Entry(identity_container, width=34)
record_username_entry.grid(row=0, column=3, sticky="ew", padx=(0, 28), pady=7)

ttk.Label(identity_container, text="Record Password", style="Card.TLabel").grid(
    row=1, column=0, sticky="e", padx=(8, 10), pady=7
)
record_password_entry = ttk.Entry(identity_container, show="*", width=34)
record_password_entry.grid(row=1, column=1, sticky="ew", padx=(0, 28), pady=7)

ttk.Label(
    identity_container,
    text="Leave password blank while editing to keep the stored password.",
    style="Card.TLabel",
).grid(row=1, column=3, sticky="w", pady=(0, 7))

identity_container.columnconfigure(1, weight=1)
identity_container.columnconfigure(3, weight=1)

form_container = tk.Frame(form_card, bg=CARD)
form_container.pack(padx=24, pady=(8, 22), fill="x")

entries = []

for i, field_name in enumerate(AUDITOR_FIELDS):
    row = i % 10
    col_block = i // 10
    label_col = col_block * 2
    entry_col = label_col + 1

    ttk.Label(form_container, text=field_name, style="Card.TLabel").grid(
        row=row, column=label_col, sticky="e", padx=(8, 10), pady=7
    )

    if field_name == "OS COA":
        e = ttk.Combobox(form_container, values=OS_COA_OPTIONS, width=32, state="readonly")
    elif field_name == "LCD type":
        e = ttk.Combobox(form_container, values=LCD_TYPE_OPTIONS, width=32, state="readonly")
    else:
        e = ttk.Entry(form_container, width=34)
    e.grid(row=row, column=entry_col, sticky="ew", padx=(0, 28), pady=7)
    entries.append(e)

form_container.columnconfigure(1, weight=1)
form_container.columnconfigure(3, weight=1)

options_row = tk.Frame(form_card, bg=CARD)
options_row.pack(fill="x", padx=24, pady=(0, 22))

ttk.Label(options_row, text="Option", style="Card.TLabel").pack(side="left", padx=(8, 10))
dropdown = ttk.Combobox(options_row, values=[f"Option {i}" for i in range(1, 11)], width=28, state="readonly")
dropdown.pack(side="left", padx=(0, 25))

checkbox_var = tk.BooleanVar()
ttk.Checkbutton(options_row, text="Accept Terms", variable=checkbox_var, style="Card.TCheckbutton").pack(side="left")

tab_empty_var = tk.BooleanVar(value=True)
ttk.Checkbutton(
    options_row,
    text="Tab empty fields",
    variable=tab_empty_var,
    style="Card.TCheckbutton",
).pack(side="left", padx=(20, 0))

ttk.Button(
    options_row,
    text="Load PC Auditor Example",
    command=lambda: load_auditor_sample(),
).pack(side="left", padx=(18, 0))

ttk.Button(
    options_row,
    text="Read This PC",
    command=lambda: start_read_current_pc_info(),
).pack(side="left", padx=(12, 0))


# =========================
# DATABASE TAB
# =========================
db_frame = tk.Frame(notebook, bg=BG)
notebook.add(db_frame, text="Database", state="hidden")

db_header = tk.Frame(db_frame, bg=BG)
db_header.pack(fill="x", padx=10, pady=(12, 4))

ttk.Label(db_header, text="Database Records", style="Header.TLabel").pack(anchor="w")
ttk.Label(db_header, text="Search, select, update, delete and export records", style="SubHeader.TLabel").pack(anchor="w")

toolbar = tk.Frame(db_frame, bg=BG)
toolbar.pack(fill="x", padx=10, pady=8)

search_var = tk.StringVar()
ttk.Label(toolbar, text="Search username / Ξενόγλωσση:").pack(side="left")
search_entry = ttk.Entry(toolbar, textvariable=search_var, width=30)
search_entry.pack(side="left", padx=8)
ttk.Button(toolbar, text="Search", command=lambda: load_data(search_var.get())).pack(side="left", padx=4)
ttk.Button(toolbar, text="Refresh", command=lambda: load_data()).pack(side="left", padx=4)

columns = ("RecordID", "Ξενόγλωσση", "Timestamp", "Username", "Dropdown", "Accepted")
tree_container = tk.Frame(db_frame, bg=BG)
tree_container.pack(fill="both", expand=True, padx=10, pady=10)

tree = ttk.Treeview(tree_container, columns=columns, show="headings")
tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
tree_scroll_x = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=160, stretch=True)
tree.configure(displaycolumns=("Ξενόγλωσση", "Timestamp", "Username", "Dropdown", "Accepted"))

tree.grid(row=0, column=0, sticky="nsew")
tree_scroll_y.grid(row=0, column=1, sticky="ns")
tree_scroll_x.grid(row=1, column=0, sticky="ew")

tree_container.rowconfigure(0, weight=1)
tree_container.columnconfigure(0, weight=1)


# =========================
# ANALYTICS TAB
# =========================
analytics_frame = tk.Frame(notebook, bg=BG)
notebook.add(analytics_frame, text="Analytics", state="hidden")

analytics_content = tk.Frame(analytics_frame, bg=BG)
analytics_content.pack(fill="both", expand=True, padx=24, pady=24)

analytics_top = tk.Frame(analytics_content, bg=BG)
analytics_top.pack(fill="x")
ttk.Label(analytics_top, text="Analytics", style="Header.TLabel").pack(side="left", anchor="w")
ttk.Button(analytics_top, text="Refresh", command=lambda: refresh_analytics()).pack(side="right")
ttk.Label(
    analytics_content,
    text="Live overview of records, users and system activity",
    style="SubHeader.TLabel",
).pack(anchor="w", pady=(0, 18))

kpi_frame = tk.Frame(analytics_content, bg=BG)
kpi_frame.pack(fill="x", pady=(0, 16))

kpi_vars = {}
for idx, key in enumerate(["Total Records", "Accepted Terms", "Active Users", "Exports"]):
    card = make_card(kpi_frame)
    card.grid(row=0, column=idx, sticky="ew", padx=(0 if idx == 0 else 8, 0))
    ttk.Label(card, text=key.upper(), style="KpiTitle.TLabel").pack(anchor="w", padx=18, pady=(14, 2))
    value_label = ttk.Label(card, text="0", style="KpiValue.TLabel")
    value_label.pack(anchor="w", padx=18, pady=(0, 14))
    kpi_vars[key] = value_label
    kpi_frame.columnconfigure(idx, weight=1)

analytics_bottom = tk.Frame(analytics_content, bg=BG)
analytics_bottom.pack(fill="both", expand=True)

option_card = make_card(analytics_bottom, "Option Distribution")
option_card.pack(side="left", fill="both", expand=True, padx=(0, 8))

option_tree = ttk.Treeview(option_card, columns=("Option", "Records"), show="headings", height=10)
for col in ("Option", "Records"):
    option_tree.heading(col, text=col)
    option_tree.column(col, anchor="center", width=160, stretch=True)
option_tree.pack(fill="both", expand=True, padx=18, pady=(8, 18))

activity_card = make_card(analytics_bottom, "Recent Activity")
activity_card.pack(side="left", fill="both", expand=True, padx=(8, 0))

activity_tree = ttk.Treeview(activity_card, columns=("Time", "User", "Action"), show="headings", height=10)
for col in ("Time", "User", "Action"):
    activity_tree.heading(col, text=col)
    activity_tree.column(col, anchor="center", width=150, stretch=True)
activity_tree.pack(fill="both", expand=True, padx=18, pady=(8, 18))


# =========================
# REPORTS TAB
# =========================
reports_frame = tk.Frame(notebook, bg=BG)
notebook.add(reports_frame, text="Reports", state="hidden")

reports_content = tk.Frame(reports_frame, bg=BG)
reports_content.pack(fill="both", expand=True, padx=24, pady=24)

ttk.Label(reports_content, text="Reports", style="Header.TLabel").pack(anchor="w")
ttk.Label(
    reports_content,
    text="Build filtered reports and export the exact result set",
    style="SubHeader.TLabel",
).pack(anchor="w", pady=(0, 18))

report_filter_card = make_card(reports_content, "Report Filters")
report_filter_card.pack(fill="x", pady=(0, 14))

report_filters = tk.Frame(report_filter_card, bg=CARD)
report_filters.pack(fill="x", padx=24, pady=(8, 20))

report_username_var = tk.StringVar()
report_option_var = tk.StringVar()
report_from_var = tk.StringVar()
report_to_var = tk.StringVar()

ttk.Label(report_filters, text="Username", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(report_filters, textvariable=report_username_var, width=24).grid(row=0, column=1, sticky="ew", padx=(0, 18), pady=6)
ttk.Label(report_filters, text="Option", style="Card.TLabel").grid(row=0, column=2, sticky="e", padx=(0, 8), pady=6)
report_option_combo = ttk.Combobox(
    report_filters,
    textvariable=report_option_var,
    values=[""] + [f"Option {i}" for i in range(1, 11)],
    width=22,
    state="readonly",
)
report_option_combo.grid(row=0, column=3, sticky="ew", padx=(0, 18), pady=6)
ttk.Label(report_filters, text="From", style="Card.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(report_filters, textvariable=report_from_var, width=24).grid(row=1, column=1, sticky="ew", padx=(0, 18), pady=6)
ttk.Label(report_filters, text="To", style="Card.TLabel").grid(row=1, column=2, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(report_filters, textvariable=report_to_var, width=24).grid(row=1, column=3, sticky="ew", padx=(0, 18), pady=6)
ttk.Button(report_filters, text="Build Report", command=lambda: build_report()).grid(row=0, column=4, padx=4, pady=6)
ttk.Button(report_filters, text="Export Excel", command=lambda: export_report_excel()).grid(row=1, column=4, padx=4, pady=6)

for idx in (1, 3):
    report_filters.columnconfigure(idx, weight=1)

report_summary_var = tk.StringVar(value="No report built yet")
ttk.Label(reports_content, textvariable=report_summary_var, style="SubHeader.TLabel").pack(anchor="w", pady=(0, 8))

report_tree_container = tk.Frame(reports_content, bg=BG)
report_tree_container.pack(fill="both", expand=True)

report_columns = ("Ξενόγλωσση", "Timestamp", "Username", "Dropdown", "Accepted", "Filled Inputs")
report_tree = ttk.Treeview(report_tree_container, columns=report_columns, show="headings")
report_y = ttk.Scrollbar(report_tree_container, orient="vertical", command=report_tree.yview)
report_x = ttk.Scrollbar(report_tree_container, orient="horizontal", command=report_tree.xview)
report_tree.configure(yscrollcommand=report_y.set, xscrollcommand=report_x.set)

for col in report_columns:
    report_tree.heading(col, text=col)
    report_tree.column(col, anchor="center", width=150, stretch=True)

report_tree.grid(row=0, column=0, sticky="nsew")
report_y.grid(row=0, column=1, sticky="ns")
report_x.grid(row=1, column=0, sticky="ew")
report_tree_container.rowconfigure(0, weight=1)
report_tree_container.columnconfigure(0, weight=1)


# =========================
# NETWORK JOBS TAB
# =========================
network_frame = tk.Frame(notebook, bg=BG)
notebook.add(network_frame, text="Network Jobs", state="hidden")

network_content = tk.Frame(network_frame, bg=BG)
network_content.pack(fill="both", expand=True, padx=24, pady=24)

ttk.Label(network_content, text="Network Jobs", style="Header.TLabel").pack(anchor="w")
ttk.Label(
    network_content,
    text="Share PC Auditor jobs through the network and execute them locally on each client",
    style="SubHeader.TLabel",
).pack(anchor="w", pady=(0, 18))

network_config_card = make_card(network_content, "Connection")
network_config_card.pack(fill="x", pady=(0, 14))

network_config = tk.Frame(network_config_card, bg=CARD)
network_config.pack(fill="x", padx=24, pady=(8, 20))

network_server_url_var = tk.StringVar()
workstation_name_var = tk.StringVar()
job_assignee_var = tk.StringVar(value="ALL")
job_title_var = tk.StringVar()
job_status_filter_var = tk.StringVar(value="Queued")

ttk.Label(network_config, text="Server URL", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(network_config, textvariable=network_server_url_var, width=36).grid(
    row=0, column=1, columnspan=3, sticky="ew", padx=(0, 18), pady=6
)
ttk.Label(network_config, text="This Client", style="Card.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(network_config, textvariable=workstation_name_var, width=24).grid(
    row=1, column=1, sticky="ew", padx=(0, 18), pady=6
)
ttk.Label(network_config, text="Status", style="Card.TLabel").grid(row=1, column=2, sticky="e", padx=(0, 8), pady=6)
ttk.Combobox(
    network_config,
    textvariable=job_status_filter_var,
    values=["Queued", "Loaded", "Completed", "Cancelled", "Failed", "All"],
    state="readonly",
    width=22,
).grid(row=1, column=3, sticky="ew", padx=(0, 18), pady=6)
ttk.Button(network_config, text="Save", command=lambda: save_network_settings()).grid(row=0, column=4, padx=4, pady=6)
ttk.Button(network_config, text="Test", command=lambda: test_network_connection()).grid(row=1, column=4, padx=4, pady=6)

for idx in (1, 3):
    network_config.columnconfigure(idx, weight=1)

network_clients_card = make_card(network_content, "Available Clients")
network_clients_card.pack(fill="x", pady=(0, 14))

network_clients_body = tk.Frame(network_clients_card, bg=CARD)
network_clients_body.pack(fill="x", padx=24, pady=(8, 20))

network_clients_toolbar = tk.Frame(network_clients_body, bg=CARD)
network_clients_toolbar.pack(fill="x", pady=(0, 8))
ttk.Button(
    network_clients_toolbar,
    text="Register This Client",
    command=lambda: register_this_client(),
).pack(side="left", padx=(0, 6))
ttk.Button(
    network_clients_toolbar,
    text="Refresh Clients",
    command=lambda: fetch_network_clients(),
).pack(side="left", padx=6)
ttk.Button(
    network_clients_toolbar,
    text="Load Client Info To Form",
    command=lambda: load_selected_client_info_to_form(),
).pack(side="left", padx=6)
ttk.Label(
    network_clients_toolbar,
    text="Select a row to use it as the publish assignee.",
    style="Card.TLabel",
).pack(side="left", padx=12)

network_client_columns = ("Name", "User", "Role", "Last Seen", "Platform", "App")
network_client_tree = ttk.Treeview(network_clients_body, columns=network_client_columns, show="headings", height=5)
client_y = ttk.Scrollbar(network_clients_body, orient="vertical", command=network_client_tree.yview)
network_client_tree.configure(yscrollcommand=client_y.set)

for col in network_client_columns:
    network_client_tree.heading(col, text=col)
    network_client_tree.column(col, anchor="center", width=140, stretch=True)
network_client_tree.column("Name", width=180)
network_client_tree.column("Platform", width=180)

network_client_tree.pack(side="left", fill="x", expand=True)
client_y.pack(side="left", fill="y")

network_publish_card = make_card(network_content, "Publish Current Form")
network_publish_card.pack(fill="x", pady=(0, 14))

network_publish = tk.Frame(network_publish_card, bg=CARD)
network_publish.pack(fill="x", padx=24, pady=(8, 20))

ttk.Label(network_publish, text="Job Title", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(network_publish, textvariable=job_title_var, width=36).grid(
    row=0, column=1, sticky="ew", padx=(0, 18), pady=6
)
ttk.Label(network_publish, text="Assignee", style="Card.TLabel").grid(row=0, column=2, sticky="e", padx=(0, 8), pady=6)
job_assignee_combo = ttk.Combobox(network_publish, textvariable=job_assignee_var, values=["ALL"], width=24)
job_assignee_combo.grid(
    row=0, column=3, sticky="ew", padx=(0, 18), pady=6
)
ttk.Button(network_publish, text="Publish Job", command=lambda: publish_current_form_job(), style="Success.TButton").grid(
    row=0, column=4, padx=4, pady=6
)
network_publish.columnconfigure(1, weight=1)
network_publish.columnconfigure(3, weight=1)

network_toolbar = tk.Frame(network_content, bg=BG)
network_toolbar.pack(fill="x", pady=(0, 8))
ttk.Button(network_toolbar, text="Fetch Jobs", command=lambda: fetch_network_jobs()).pack(side="left", padx=(0, 6))
ttk.Button(network_toolbar, text="Load Selected To Form", command=lambda: load_selected_network_job()).pack(side="left", padx=6)
ttk.Button(network_toolbar, text="Apply Selected", command=lambda: apply_selected_network_job()).pack(side="left", padx=6)
ttk.Button(network_toolbar, text="Mark Completed", command=lambda: complete_selected_network_job(), style="Success.TButton").pack(
    side="left", padx=6
)
ttk.Button(network_toolbar, text="Mark Failed", command=lambda: fail_selected_network_job(), style="Danger.TButton").pack(
    side="left", padx=6
)

network_body = tk.Frame(network_content, bg=BG)
network_body.pack(fill="both", expand=True)

network_tree_container = tk.Frame(network_body, bg=BG)
network_tree_container.pack(side="left", fill="both", expand=True, padx=(0, 8))

network_columns = ("ID", "Title", "Assignee", "Status", "Created By", "Created At", "Updated At")
network_tree = ttk.Treeview(network_tree_container, columns=network_columns, show="headings")
network_y = ttk.Scrollbar(network_tree_container, orient="vertical", command=network_tree.yview)
network_x = ttk.Scrollbar(network_tree_container, orient="horizontal", command=network_tree.xview)
network_tree.configure(yscrollcommand=network_y.set, xscrollcommand=network_x.set)

for col in network_columns:
    network_tree.heading(col, text=col)
    network_tree.column(col, anchor="center", width=130, stretch=True)
network_tree.column("Title", width=220)

network_tree.grid(row=0, column=0, sticky="nsew")
network_y.grid(row=0, column=1, sticky="ns")
network_x.grid(row=1, column=0, sticky="ew")
network_tree_container.rowconfigure(0, weight=1)
network_tree_container.columnconfigure(0, weight=1)

network_details_card = make_card(network_body, "Selected Job")
network_details_card.pack(side="left", fill="both", expand=True, padx=(8, 0))

network_details_text = tk.Text(
    network_details_card,
    bg=ENTRY_BG,
    fg=TEXT,
    insertbackground=TEXT,
    relief="flat",
    height=12,
    wrap="word",
)
network_details_text.pack(fill="both", expand=True, padx=18, pady=(8, 18))
network_details_text.configure(state="disabled")


# =========================
# DEVICE RESULTS TAB
# =========================
device_results_frame = tk.Frame(notebook, bg=BG)
notebook.add(device_results_frame, text="Device Results", state="hidden")

device_results_content = tk.Frame(device_results_frame, bg=BG)
device_results_content.pack(fill="both", expand=True, padx=24, pady=24)

ttk.Label(device_results_content, text="Device Results", style="Header.TLabel").pack(anchor="w")
ttk.Label(
    device_results_content,
    text="Remote client hardware diagnostics reported by the registration heartbeat",
    style="SubHeader.TLabel",
).pack(anchor="w", pady=(0, 18))

device_results_toolbar = tk.Frame(device_results_content, bg=BG)
device_results_toolbar.pack(fill="x", pady=(0, 10))

device_status_filter_var = tk.StringVar(value="All")
ttk.Button(device_results_toolbar, text="Refresh Results", command=lambda: refresh_device_results()).pack(side="left", padx=(0, 6))
ttk.Label(device_results_toolbar, text="Status:").pack(side="left", padx=(12, 6))
ttk.Combobox(
    device_results_toolbar,
    textvariable=device_status_filter_var,
    values=["All", "OK", "WARN", "FAIL", "UNKNOWN"],
    state="readonly",
    width=14,
).pack(side="left", padx=(0, 8))
ttk.Button(device_results_toolbar, text="Apply Filter", command=lambda: populate_device_results()).pack(side="left", padx=6)

device_summary_var = tk.StringVar(value="No diagnostics loaded")
ttk.Label(device_results_content, textvariable=device_summary_var, style="SubHeader.TLabel").pack(anchor="w", pady=(0, 8))

device_results_body = tk.Frame(device_results_content, bg=BG)
device_results_body.pack(fill="both", expand=True)

device_results_tree_container = tk.Frame(device_results_body, bg=BG)
device_results_tree_container.pack(side="left", fill="both", expand=True, padx=(0, 8))

device_result_columns = ("Client", "Device", "Status", "Details", "Last Seen")
device_results_tree = ttk.Treeview(device_results_tree_container, columns=device_result_columns, show="headings")
device_results_y = ttk.Scrollbar(device_results_tree_container, orient="vertical", command=device_results_tree.yview)
device_results_x = ttk.Scrollbar(device_results_tree_container, orient="horizontal", command=device_results_tree.xview)
device_results_tree.configure(yscrollcommand=device_results_y.set, xscrollcommand=device_results_x.set)

for col in device_result_columns:
    device_results_tree.heading(col, text=col)
    device_results_tree.column(col, anchor="center", width=150, stretch=True)
device_results_tree.column("Details", width=420)

device_results_tree.grid(row=0, column=0, sticky="nsew")
device_results_y.grid(row=0, column=1, sticky="ns")
device_results_x.grid(row=1, column=0, sticky="ew")
device_results_tree_container.rowconfigure(0, weight=1)
device_results_tree_container.columnconfigure(0, weight=1)

device_details_card = make_card(device_results_body, "Selected Result")
device_details_card.pack(side="left", fill="both", expand=True, padx=(8, 0))

device_details_text = tk.Text(
    device_details_card,
    bg=ENTRY_BG,
    fg=TEXT,
    insertbackground=TEXT,
    relief="flat",
    height=12,
    wrap="word",
)
device_details_text.pack(fill="both", expand=True, padx=18, pady=(8, 18))
device_details_text.configure(state="disabled")


# =========================
# MANUAL TESTS TAB
# =========================
manual_tests_frame = tk.Frame(notebook, bg=BG)
notebook.add(manual_tests_frame, text="Manual Tests", state="hidden")

manual_tests_content = tk.Frame(manual_tests_frame, bg=BG)
manual_tests_content.pack(fill="both", expand=True, padx=24, pady=24)

ttk.Label(manual_tests_content, text="Manual Tests", style="Header.TLabel").pack(anchor="w")
ttk.Label(
    manual_tests_content,
    text="All technician and auditor-style device/port checks in one view",
    style="SubHeader.TLabel",
).pack(anchor="w", pady=(0, 18))

manual_tests_toolbar = tk.Frame(manual_tests_content, bg=BG)
manual_tests_toolbar.pack(fill="x", pady=(0, 10))

manual_status_filter_var = tk.StringVar(value="All")
manual_required_var = tk.BooleanVar(value=False)
manual_windows_usb_expected_var = tk.StringVar()
manual_windows_optical_path_var = tk.StringVar()
manual_windows_usb_path_var = tk.StringVar()
manual_windows_test_status_var = tk.StringVar(value="Windows device tests not run")
ttk.Button(manual_tests_toolbar, text="Refresh Manual Tests", command=lambda: refresh_manual_tests()).pack(side="left", padx=(0, 6))
ttk.Label(manual_tests_toolbar, text="Status:").pack(side="left", padx=(12, 6))
ttk.Combobox(
    manual_tests_toolbar,
    textvariable=manual_status_filter_var,
    values=["All", "OK", "WARN", "FAIL", "UNKNOWN"],
    state="readonly",
    width=14,
).pack(side="left", padx=(0, 8))
ttk.Button(manual_tests_toolbar, text="Apply Filter", command=lambda: populate_manual_tests()).pack(side="left", padx=6)
ttk.Checkbutton(
    manual_tests_toolbar,
    text="Manual tests required",
    variable=manual_required_var,
).pack(side="left", padx=(18, 6))
ttk.Button(manual_tests_toolbar, text="Save Requirement", command=lambda: save_manual_requirement()).pack(side="left", padx=6)
ttk.Button(manual_tests_toolbar, text="Load Requirement", command=lambda: load_manual_requirement()).pack(side="left", padx=6)

manual_windows_card = make_card(manual_tests_content, "Windows Device Tests")
manual_windows_card.pack(fill="x", pady=(0, 10))

manual_windows_body = tk.Frame(manual_windows_card, bg=CARD)
manual_windows_body.pack(fill="x", padx=24, pady=(8, 18))

ttk.Label(manual_windows_body, text="USB ports", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(manual_windows_body, textvariable=manual_windows_usb_expected_var, width=10).grid(row=0, column=1, sticky="w", padx=(0, 18), pady=6)
ttk.Label(manual_windows_body, text="Optical path", style="Card.TLabel").grid(row=0, column=2, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(manual_windows_body, textvariable=manual_windows_optical_path_var, width=18).grid(row=0, column=3, sticky="ew", padx=(0, 8), pady=6)
ttk.Button(manual_windows_body, text="Browse", command=lambda: browse_manual_test_path(manual_windows_optical_path_var)).grid(row=0, column=4, sticky="w", padx=(0, 18), pady=6)

ttk.Label(manual_windows_body, text="USB/SD path", style="Card.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(manual_windows_body, textvariable=manual_windows_usb_path_var, width=18).grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=6)
ttk.Button(manual_windows_body, text="Browse", command=lambda: browse_manual_test_path(manual_windows_usb_path_var)).grid(row=1, column=2, sticky="w", padx=(0, 18), pady=6)
ttk.Button(manual_windows_body, text="Detect Paths", command=lambda: detect_windows_test_paths()).grid(row=1, column=3, sticky="w", padx=(0, 8), pady=6)
manual_windows_run_btn = ttk.Button(
    manual_windows_body,
    text="Run Windows Tests",
    command=lambda: run_windows_device_tests_from_tab(),
    style="Success.TButton",
)
manual_windows_run_btn.grid(row=1, column=4, sticky="w", padx=(0, 18), pady=6)

ttk.Label(manual_windows_body, textvariable=manual_windows_test_status_var, style="Card.TLabel").grid(
    row=2, column=0, columnspan=5, sticky="w", pady=(6, 0)
)
manual_windows_body.columnconfigure(1, weight=1)
manual_windows_body.columnconfigure(3, weight=1)

manual_summary_var = tk.StringVar(value="No manual tests loaded")
ttk.Label(manual_tests_content, textvariable=manual_summary_var, style="SubHeader.TLabel").pack(anchor="w", pady=(0, 8))

manual_tests_body = tk.Frame(manual_tests_content, bg=BG)
manual_tests_body.pack(fill="both", expand=True)

manual_tree_container = tk.Frame(manual_tests_body, bg=BG)
manual_tree_container.pack(side="left", fill="both", expand=True, padx=(0, 8))

manual_columns = ("Client", "Test", "Status", "Note", "Last Seen")
manual_tests_tree = ttk.Treeview(manual_tree_container, columns=manual_columns, show="headings")
manual_y = ttk.Scrollbar(manual_tree_container, orient="vertical", command=manual_tests_tree.yview)
manual_x = ttk.Scrollbar(manual_tree_container, orient="horizontal", command=manual_tests_tree.xview)
manual_tests_tree.configure(yscrollcommand=manual_y.set, xscrollcommand=manual_x.set)

for col in manual_columns:
    manual_tests_tree.heading(col, text=col)
    manual_tests_tree.column(col, anchor="center", width=150, stretch=True)
manual_tests_tree.column("Note", width=440)

manual_tests_tree.grid(row=0, column=0, sticky="nsew")
manual_y.grid(row=0, column=1, sticky="ns")
manual_x.grid(row=1, column=0, sticky="ew")
manual_tree_container.rowconfigure(0, weight=1)
manual_tree_container.columnconfigure(0, weight=1)

manual_details_card = make_card(manual_tests_body, "Selected Test")
manual_details_card.pack(side="left", fill="both", expand=True, padx=(8, 0))

manual_details_text = tk.Text(
    manual_details_card,
    bg=ENTRY_BG,
    fg=TEXT,
    insertbackground=TEXT,
    relief="flat",
    height=12,
    wrap="word",
)
manual_details_text.pack(fill="both", expand=True, padx=18, pady=(8, 18))
manual_details_text.configure(state="disabled")


# =========================
# USERS TAB
# =========================
users_frame = tk.Frame(notebook, bg=BG)
notebook.add(users_frame, text="Users", state="hidden")

users_content = tk.Frame(users_frame, bg=BG)
users_content.pack(fill="both", expand=True, padx=24, pady=24)

ttk.Label(users_content, text="Users", style="Header.TLabel").pack(anchor="w")
ttk.Label(users_content, text="Manage application users, roles and account status", style="SubHeader.TLabel").pack(
    anchor="w", pady=(0, 18)
)

user_form_card = make_card(users_content, "User Details")
user_form_card.pack(fill="x", pady=(0, 14))

user_form = tk.Frame(user_form_card, bg=CARD)
user_form.pack(fill="x", padx=24, pady=(8, 20))

user_username_var = tk.StringVar()
user_role_var = tk.StringVar(value="User")
user_active_var = tk.BooleanVar(value=True)

ttk.Label(user_form, text="Username", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Entry(user_form, textvariable=user_username_var, width=24).grid(row=0, column=1, sticky="ew", padx=(0, 18), pady=6)
ttk.Label(user_form, text="Password", style="Card.TLabel").grid(row=0, column=2, sticky="e", padx=(0, 8), pady=6)
user_password_entry = ttk.Entry(user_form, show="*", width=24)
user_password_entry.grid(row=0, column=3, sticky="ew", padx=(0, 18), pady=6)
ttk.Label(user_form, text="Role", style="Card.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=6)
ttk.Combobox(user_form, textvariable=user_role_var, values=["Admin", "User", "Viewer"], state="readonly", width=22).grid(
    row=1, column=1, sticky="ew", padx=(0, 18), pady=6
)
ttk.Checkbutton(user_form, text="Active", variable=user_active_var, style="Card.TCheckbutton").grid(
    row=1, column=2, sticky="w", padx=(0, 8), pady=6
)
ttk.Button(user_form, text="Add User", command=lambda: add_user(), style="Success.TButton").grid(row=0, column=4, padx=4)
ttk.Button(user_form, text="Update", command=lambda: update_user()).grid(row=1, column=4, padx=4)
ttk.Button(user_form, text="Clear", command=lambda: clear_user_form()).grid(row=0, column=5, padx=4)
ttk.Button(user_form, text="Delete", command=lambda: delete_user(), style="Danger.TButton").grid(row=1, column=5, padx=4)

for idx in (1, 3):
    user_form.columnconfigure(idx, weight=1)

user_tree_container = tk.Frame(users_content, bg=BG)
user_tree_container.pack(fill="both", expand=True)

user_columns = ("ID", "Username", "Role", "Status", "Created", "Last Login")
user_tree = ttk.Treeview(user_tree_container, columns=user_columns, show="headings")
user_y = ttk.Scrollbar(user_tree_container, orient="vertical", command=user_tree.yview)
user_x = ttk.Scrollbar(user_tree_container, orient="horizontal", command=user_tree.xview)
user_tree.configure(yscrollcommand=user_y.set, xscrollcommand=user_x.set)

for col in user_columns:
    user_tree.heading(col, text=col)
    user_tree.column(col, anchor="center", width=150, stretch=True)

user_tree.grid(row=0, column=0, sticky="nsew")
user_y.grid(row=0, column=1, sticky="ns")
user_x.grid(row=1, column=0, sticky="ew")
user_tree_container.rowconfigure(0, weight=1)
user_tree_container.columnconfigure(0, weight=1)


# =========================
# SETTINGS TAB
# =========================
settings_frame = tk.Frame(notebook, bg=BG)
notebook.add(settings_frame, text="Settings", state="hidden")

settings_content = tk.Frame(settings_frame, bg=BG)
settings_content.pack(fill="both", expand=True, padx=24, pady=24)

ttk.Label(settings_content, text="Settings", style="Header.TLabel").pack(anchor="w")
ttk.Label(settings_content, text="Configure application preferences and database maintenance", style="SubHeader.TLabel").pack(
    anchor="w", pady=(0, 18)
)

settings_card = make_card(settings_content, "Application Settings")
settings_card.pack(fill="x", pady=(0, 14))

settings_form = tk.Frame(settings_card, bg=CARD)
settings_form.pack(fill="x", padx=24, pady=(8, 20))

setting_system_name_var = tk.StringVar()
setting_export_folder_var = tk.StringVar()
setting_keyboard_delay_var = tk.StringVar()
setting_auto_refresh_var = tk.BooleanVar()

ttk.Label(settings_form, text="System Name", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=7)
ttk.Entry(settings_form, textvariable=setting_system_name_var, width=42).grid(
    row=0, column=1, sticky="ew", padx=(0, 18), pady=7
)
ttk.Label(settings_form, text="Export Folder", style="Card.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=7)
ttk.Entry(settings_form, textvariable=setting_export_folder_var, width=42).grid(
    row=1, column=1, sticky="ew", padx=(0, 18), pady=7
)
ttk.Button(settings_form, text="Browse", command=lambda: choose_export_folder()).grid(row=1, column=2, padx=4, pady=7)
ttk.Label(settings_form, text="Keyboard Delay (ms)", style="Card.TLabel").grid(
    row=2, column=0, sticky="e", padx=(0, 8), pady=7
)
ttk.Entry(settings_form, textvariable=setting_keyboard_delay_var, width=18).grid(
    row=2, column=1, sticky="w", padx=(0, 18), pady=7
)
ttk.Checkbutton(settings_form, text="Auto refresh dashboards after changes", variable=setting_auto_refresh_var, style="Card.TCheckbutton").grid(
    row=3, column=1, sticky="w", pady=7
)
ttk.Button(settings_form, text="Save Settings", command=lambda: save_settings(), style="Success.TButton").grid(
    row=4, column=1, sticky="w", pady=(14, 0)
)
ttk.Button(settings_form, text="Reset Defaults", command=lambda: reset_settings()).grid(row=4, column=1, sticky="e", pady=(14, 0))

settings_form.columnconfigure(1, weight=1)

maintenance_card = make_card(settings_content, "Database Maintenance")
maintenance_card.pack(fill="x")

maintenance_body = tk.Frame(maintenance_card, bg=CARD)
maintenance_body.pack(fill="x", padx=24, pady=(8, 20))

db_path_var = tk.StringVar(value=DB_NAME)
ttk.Label(maintenance_body, text="Database", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=7)
ttk.Entry(maintenance_body, textvariable=db_path_var, width=70, state="readonly").grid(
    row=0, column=1, sticky="ew", padx=(0, 18), pady=7
)
ttk.Button(maintenance_body, text="Backup Database", command=lambda: backup_database()).grid(row=0, column=2, padx=4, pady=7)
maintenance_body.columnconfigure(1, weight=1)

batch_launcher_card = make_card(settings_content, "Batch Launcher")
batch_launcher_card.pack(fill="both", expand=True, pady=(14, 0))

batch_launcher_body = tk.Frame(batch_launcher_card, bg=CARD)
batch_launcher_body.pack(fill="both", expand=True, padx=24, pady=(8, 20))

batch_toolbar = tk.Frame(batch_launcher_body, bg=CARD)
batch_toolbar.pack(fill="x", pady=(0, 8))

ttk.Button(batch_toolbar, text="Refresh Batch Files", command=lambda: load_batch_files()).pack(side="left", padx=(0, 6))
ttk.Button(batch_toolbar, text="Run Selected Background", command=lambda: run_selected_batch_files(), style="Success.TButton").pack(
    side="left", padx=6
)
ttk.Button(batch_toolbar, text="Run All Background", command=lambda: run_all_batch_files(), style="Danger.TButton").pack(side="left", padx=6)

batch_status_var = tk.StringVar(value="No batch files loaded")
ttk.Label(batch_toolbar, textvariable=batch_status_var, style="Card.TLabel").pack(side="left", padx=14)

batch_tree_container = tk.Frame(batch_launcher_body, bg=CARD)
batch_tree_container.pack(fill="both", expand=True)

batch_columns = ("File", "Modified", "Size")
batch_tree = ttk.Treeview(batch_tree_container, columns=batch_columns, show="headings", height=6, selectmode="extended")
batch_y = ttk.Scrollbar(batch_tree_container, orient="vertical", command=batch_tree.yview)
batch_tree.configure(yscrollcommand=batch_y.set)

for col in batch_columns:
    batch_tree.heading(col, text=col)
    batch_tree.column(col, anchor="center", width=160, stretch=True)
batch_tree.column("File", anchor="w", width=280)

batch_tree.pack(side="left", fill="both", expand=True)
batch_y.pack(side="left", fill="y")


# =========================
# LOGS TAB
# =========================
logs_frame = tk.Frame(notebook, bg=BG)
notebook.add(logs_frame, text="Logs", state="hidden")

logs_content = tk.Frame(logs_frame, bg=BG)
logs_content.pack(fill="both", expand=True, padx=24, pady=24)

ttk.Label(logs_content, text="Logs", style="Header.TLabel").pack(anchor="w")
ttk.Label(logs_content, text="Audit trail for sessions, records, exports and administration", style="SubHeader.TLabel").pack(
    anchor="w", pady=(0, 18)
)

log_toolbar = tk.Frame(logs_content, bg=BG)
log_toolbar.pack(fill="x", pady=(0, 10))

log_user_var = tk.StringVar()
log_action_var = tk.StringVar()
ttk.Label(log_toolbar, text="User:").pack(side="left")
ttk.Entry(log_toolbar, textvariable=log_user_var, width=22).pack(side="left", padx=8)
ttk.Label(log_toolbar, text="Action:").pack(side="left")
ttk.Entry(log_toolbar, textvariable=log_action_var, width=22).pack(side="left", padx=8)
ttk.Button(log_toolbar, text="Refresh", command=lambda: load_logs()).pack(side="left", padx=4)
ttk.Button(log_toolbar, text="Export Excel", command=lambda: export_logs_excel()).pack(side="left", padx=4)
ttk.Button(log_toolbar, text="Clear Logs", command=lambda: clear_logs(), style="Danger.TButton").pack(side="left", padx=4)

log_tree_container = tk.Frame(logs_content, bg=BG)
log_tree_container.pack(fill="both", expand=True)

log_columns = ("ID", "Timestamp", "User", "Action", "Details")
log_tree = ttk.Treeview(log_tree_container, columns=log_columns, show="headings")
log_y = ttk.Scrollbar(log_tree_container, orient="vertical", command=log_tree.yview)
log_x = ttk.Scrollbar(log_tree_container, orient="horizontal", command=log_tree.xview)
log_tree.configure(yscrollcommand=log_y.set, xscrollcommand=log_x.set)

for col in log_columns:
    log_tree.heading(col, text=col)
    log_tree.column(col, anchor="center", width=150, stretch=True)
log_tree.column("Details", width=320)

log_tree.grid(row=0, column=0, sticky="nsew")
log_y.grid(row=0, column=1, sticky="ns")
log_x.grid(row=1, column=0, sticky="ew")
log_tree_container.rowconfigure(0, weight=1)
log_tree_container.columnconfigure(0, weight=1)


# =========================
# SYSTEM INVENTORY
# =========================
def run_command(command, timeout=6):
    try:
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW") else 0
        kwargs = {
            "capture_output": True,
            "text": True,
            "errors": "replace",
            "timeout": timeout,
        }
        if os.name == "nt":
            kwargs["creationflags"] = flags
        result = subprocess.run(
            command,
            **kwargs,
        )
    except Exception:
        return ""

    if result.returncode != 0:
        return ""
    return result.stdout.strip()


def powershell_value(script):
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "$ErrorActionPreference='SilentlyContinue'; " + script,
    ]
    value = run_command(command)
    if value:
        return value
    command[0] = "pwsh"
    return run_command(command)


def read_text_file(path):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as handle:
            return handle.read().strip()
    except OSError:
        return ""


def read_binary_file(path):
    try:
        with open(path, "rb") as handle:
            return handle.read()
    except OSError:
        return b""


def first_line(value):
    for line in str(value).splitlines():
        line = line.strip()
        if line:
            return line
    return ""


def format_gb(bytes_value):
    try:
        gb = int(float(bytes_value)) / (1024 ** 3)
    except (TypeError, ValueError):
        return ""

    rounded = round(gb)
    if rounded <= 0:
        return ""
    return f"{rounded} GB"


def chassis_label(chassis_value):
    try:
        chassis = int(str(chassis_value).strip().split()[0])
    except (TypeError, ValueError, IndexError):
        return ""

    laptop_codes = {8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32}
    desktop_codes = {3, 4, 5, 6, 7, 15, 16}
    tablet_codes = {13, 17, 23, 24}

    if chassis in laptop_codes:
        return "LAPTOP"
    if chassis in desktop_codes:
        return "DESKTOP"
    if chassis in tablet_codes:
        return "TABLET"
    return ""


def screen_resolution():
    return SCREEN_RESOLUTION_CACHE


def locale_label():
    lang = locale.getlocale()[0] or os.environ.get("LANG", "")
    return lang.replace("_", "-") if lang else ""


def infer_screen_type_from_text(text):
    haystack = f" {text} ".lower()
    if "amoled" in haystack:
        return "AMOLED"
    if "oled" in haystack:
        return "OLED"
    if "mini led" in haystack or "miniled" in haystack:
        return "Mini LED"
    if "ips" in haystack:
        return "IPS LCD"
    if " va " in haystack or "mva" in haystack or "pva" in haystack:
        return "VA LCD"
    if " tn " in haystack or "twisted nematic" in haystack:
        return "TN LCD"
    if "lcd" in haystack or "led" in haystack:
        return "LCD"
    return "Unknown"


def edid_descriptor_text(edid_bytes):
    if len(edid_bytes) < 128:
        return ""

    chunks = []
    for offset in range(54, 126, 18):
        block = edid_bytes[offset:offset + 18]
        if len(block) == 18 and block[:3] == b"\x00\x00\x00" and block[3] in {0xfc, 0xfe, 0xff}:
            text = block[5:18].decode("ascii", errors="ignore").strip().replace("\n", " ")
            if text:
                chunks.append(text)
    return " ".join(chunks)


def windows_monitor_text():
    monitor_name = powershell_value(
        "Get-CimInstance -Namespace root/wmi -ClassName WmiMonitorID | "
        "ForEach-Object { "
        "$name=($_.UserFriendlyName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''; "
        "$serial=($_.SerialNumberID | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''; "
        "$name + ' ' + $serial }"
    )
    desktop_monitor = powershell_value(
        "Get-CimInstance Win32_DesktopMonitor | ForEach-Object { $_.Name + ' ' + $_.MonitorType }"
    )
    video = powershell_value(
        "Get-CimInstance Win32_VideoController | Select-Object -First 1 -ExpandProperty VideoModeDescription"
    )
    return " ".join([monitor_name, desktop_monitor, video])


def linux_monitor_text():
    parts = []
    drm_path = "/sys/class/drm"
    if os.path.isdir(drm_path):
        for name in os.listdir(drm_path):
            edid_path = os.path.join(drm_path, name, "edid")
            edid = read_binary_file(edid_path)
            text = edid_descriptor_text(edid)
            if text:
                parts.append(text)

    xrandr = run_command(["xrandr", "--verbose"], timeout=4)
    if xrandr:
        parts.append(xrandr)

    return " ".join(parts)


def windows_inventory():
    manufacturer = first_line(powershell_value("(Get-CimInstance Win32_ComputerSystem).Manufacturer"))
    model = first_line(powershell_value("(Get-CimInstance Win32_ComputerSystem).Model"))
    chassis = first_line(powershell_value("(Get-CimInstance Win32_SystemEnclosure).ChassisTypes[0]"))
    product_type = chassis_label(chassis)
    part_number = first_line(powershell_value("(Get-CimInstance Win32_ComputerSystemProduct).SKUNumber"))
    if not part_number:
        part_number = first_line(powershell_value("(Get-CimInstance Win32_ComputerSystemProduct).Name"))
    serial = first_line(powershell_value("(Get-CimInstance Win32_BIOS).SerialNumber"))
    processor = first_line(powershell_value("(Get-CimInstance Win32_Processor | Select-Object -First 1 -ExpandProperty Name)"))
    ram_size = first_line(
        powershell_value("[math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB).ToString() + ' GB'")
    )
    disk_info = first_line(
        powershell_value(
            "Get-PhysicalDisk | Select-Object -First 1 | "
            "ForEach-Object { ($_.MediaType.ToString()) + '|' + ([math]::Round($_.Size / 1GB).ToString() + ' GB') }"
        )
    )
    disk_type = ""
    disk_size = ""
    if "|" in disk_info:
        disk_type, disk_size = [part.strip() for part in disk_info.split("|", 1)]
        if disk_type.upper() == "SSD":
            disk_type = "SATA SSD"
    optical = first_line(powershell_value("(Get-CimInstance Win32_CDROMDrive | Select-Object -First 1 -ExpandProperty Name)"))
    video = first_line(powershell_value("(Get-CimInstance Win32_VideoController | Select-Object -First 1 -ExpandProperty Name)"))
    resolution = first_line(
        powershell_value(
            "Get-CimInstance Win32_VideoController | Select-Object -First 1 | "
            "ForEach-Object { if ($_.CurrentHorizontalResolution) { "
            "'{0}x{1}' -f $_.CurrentHorizontalResolution, $_.CurrentVerticalResolution } }"
        )
    ) or screen_resolution()

    return {
        "Product type": product_type,
        "Manufacturer": manufacturer,
        "Model": model,
        "Chassis type": product_type,
        "Part number": part_number,
        "Serial number": serial,
        "Processor": processor,
        "RAM size": ram_size,
        "Disk type": disk_type,
        "Disk size": disk_size,
        "Optical drive": optical,
        "Video model": video,
        "LCD resolution": resolution,
        "LCD type": infer_screen_type_from_text(windows_monitor_text()),
        "Keyboard layout": locale_label(),
    }


def linux_disk_info():
    output = run_command(["lsblk", "-bdno", "NAME,SIZE,ROTA,TYPE"])
    if output:
        for line in output.splitlines():
            parts = line.split()
            if len(parts) >= 4 and parts[3] == "disk":
                disk_type = "SATA HDD" if parts[2] == "1" else "SATA SSD"
                return disk_type, format_gb(parts[1])

    for name in os.listdir("/sys/block") if os.path.isdir("/sys/block") else []:
        if name.startswith(("loop", "ram", "fd")):
            continue
        size = read_text_file(f"/sys/block/{name}/size")
        rotational = read_text_file(f"/sys/block/{name}/queue/rotational")
        if size:
            disk_type = "SATA HDD" if rotational == "1" else "SATA SSD"
            try:
                return disk_type, format_gb(int(size) * 512)
            except ValueError:
                continue

    return "", ""


def linux_cpu_name():
    cpuinfo = read_text_file("/proc/cpuinfo")
    for line in cpuinfo.splitlines():
        if line.lower().startswith("model name"):
            return line.split(":", 1)[1].strip()
    return ""


def linux_ram_size():
    meminfo = read_text_file("/proc/meminfo")
    for line in meminfo.splitlines():
        if line.startswith("MemTotal:"):
            parts = line.split()
            if len(parts) >= 2:
                return format_gb(int(parts[1]) * 1024)
    return ""


def linux_video_model():
    output = run_command(["lspci"])
    for line in output.splitlines():
        lower = line.lower()
        if "vga compatible controller" in lower or "3d controller" in lower or "display controller" in lower:
            return line.split(":", 2)[-1].strip()
    return ""


def linux_optical_drive():
    info = read_text_file("/proc/sys/dev/cdrom/info")
    for line in info.splitlines():
        if line.startswith("drive name:"):
            value = line.split(":", 1)[1].strip()
            return value or ""
    return ""


def linux_inventory():
    manufacturer = read_text_file("/sys/class/dmi/id/sys_vendor")
    model = read_text_file("/sys/class/dmi/id/product_name")
    part_number = read_text_file("/sys/class/dmi/id/product_sku")
    serial = read_text_file("/sys/class/dmi/id/product_serial") or read_text_file("/sys/class/dmi/id/board_serial")
    chassis = read_text_file("/sys/class/dmi/id/chassis_type")
    product_type = chassis_label(chassis)
    disk_type, disk_size = linux_disk_info()

    return {
        "Product type": product_type,
        "Manufacturer": manufacturer,
        "Model": model,
        "Chassis type": product_type,
        "Part number": part_number,
        "Serial number": serial,
        "Processor": linux_cpu_name(),
        "RAM size": linux_ram_size(),
        "Disk type": disk_type,
        "Disk size": disk_size,
        "Optical drive": linux_optical_drive(),
        "Video model": linux_video_model(),
        "LCD resolution": screen_resolution(),
        "LCD type": infer_screen_type_from_text(linux_monitor_text()),
        "Keyboard layout": locale_label(),
    }


def write_inventory_debug_log(title):
    try:
        with open(INVENTORY_DEBUG_LOG, "a", encoding="utf-8") as handle:
            handle.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {title}\n")
            handle.write(traceback.format_exc())
    except OSError:
        pass


def collect_system_inventory():
    try:
        if platform.system().lower().startswith("win"):
            data = windows_inventory()
        elif platform.system().lower() == "linux":
            data = linux_inventory()
        else:
            data = {"LCD resolution": screen_resolution(), "Keyboard layout": locale_label()}
        return {key: value for key, value in data.items() if value and str(value).strip()}
    except Exception:
        write_inventory_debug_log("inventory error")
        return {}


# =========================
# RECORD FUNCTIONS
# =========================
def clear_form():
    global selected_id
    selected_id = None

    record_xenoglossi_entry.delete(0, tk.END)
    record_username_entry.delete(0, tk.END)
    record_password_entry.delete(0, tk.END)
    dropdown.set("")
    checkbox_var.set(False)

    for entry in entries:
        clear_field(entry)

    if tree.selection():
        tree.selection_remove(tree.selection())


def load_auditor_sample():
    for entry in entries:
        clear_field(entry)

    for index, field_name in enumerate(AUDITOR_FIELDS):
        set_field(entries[index], AUDITOR_SAMPLE_VALUES.get(field_name, ""))

    if not record_xenoglossi_entry.get().strip():
        record_xenoglossi_entry.insert(0, AUDITOR_SAMPLE_VALUES["Serial number"])

    if not record_username_entry.get().strip():
        record_username_entry.insert(0, AUDITOR_SAMPLE_VALUES["Serial number"])

    if not dropdown.get().strip():
        dropdown.set("Option 1")


def load_current_pc_info():
    if not require_login():
        return

    inventory = collect_system_inventory()

    if not inventory:
        messagebox.showwarning("Read This PC", "No system details could be detected on this machine")
        return

    updated = 0
    for index, field_name in enumerate(AUDITOR_FIELDS):
        value = inventory.get(field_name, "")
        if value:
            set_field(entries[index], value)
            updated += 1

    serial = inventory.get("Serial number", "")
    if serial and not record_xenoglossi_entry.get().strip():
        record_xenoglossi_entry.insert(0, serial)
    if serial and not record_username_entry.get().strip():
        record_username_entry.insert(0, serial)

    log_action("READ_THIS_PC", f"Detected {updated} field(s) from local system")
    messagebox.showinfo("Read This PC", f"Detected and filled {updated} field(s)")


def finish_read_current_pc_info(inventory, error_message):
    root.configure(cursor="")

    if error_message:
        messagebox.showerror(
            "Read This PC",
            f"Could not read this PC. Details were written to inventory_debug.log.\n\n{error_message}",
        )
        return

    if not inventory:
        messagebox.showwarning("Read This PC", "No system details could be detected on this machine")
        return

    updated = 0
    for index, field_name in enumerate(AUDITOR_FIELDS):
        value = inventory.get(field_name, "")
        if value:
            set_field(entries[index], value)
            updated += 1

    serial = inventory.get("Serial number", "")
    if serial and not record_xenoglossi_entry.get().strip():
        record_xenoglossi_entry.insert(0, serial)
    if serial and not record_username_entry.get().strip():
        record_username_entry.insert(0, serial)

    log_action("READ_THIS_PC", f"Detected {updated} field(s) from local system")
    messagebox.showinfo("Read This PC", f"Detected and filled {updated} field(s)")


def start_read_current_pc_info():
    if not require_login():
        return

    root.configure(cursor="watch")
    root.update_idletasks()

    def worker():
        try:
            inventory = collect_system_inventory()
            error_message = ""
        except Exception as exc:
            inventory = {}
            error_message = str(exc)
            write_inventory_debug_log("read this pc error")
        root.after(0, lambda: finish_read_current_pc_info(inventory, error_message))

    threading.Thread(target=worker, daemon=True).start()


def load_data(filter_text=""):
    clear_tree(tree)

    conn = get_conn()
    c = conn.cursor()

    if filter_text:
        c.execute(
            f"SELECT {RECORD_SELECT_COLUMNS} FROM records WHERE username LIKE ? OR xenoglossi LIKE ? ORDER BY id DESC",
            (f"%{filter_text}%", f"%{filter_text}%"),
        )
    else:
        c.execute(f"SELECT {RECORD_SELECT_COLUMNS} FROM records ORDER BY id DESC")

    for row in c.fetchall():
        tree.insert("", "end", values=(row[0], row[1], row[2], row[3], row[5], "Yes" if row[6] else "No"))

    conn.close()


def on_select(event):
    global selected_id

    item = tree.focus()
    if not item:
        return

    selected_id = tree.item(item)["values"][0]

    conn = get_conn()
    c = conn.cursor()
    c.execute(f"SELECT {RECORD_SELECT_COLUMNS} FROM records WHERE id=?", (selected_id,))
    row = c.fetchone()
    conn.close()

    if not row:
        return

    record_xenoglossi_entry.delete(0, tk.END)
    record_xenoglossi_entry.insert(0, row[1])

    record_username_entry.delete(0, tk.END)
    record_username_entry.insert(0, row[3])

    record_password_entry.delete(0, tk.END)

    dropdown.set(row[5])
    checkbox_var.set(bool(row[6]))

    for entry in entries:
        clear_field(entry)

    values = normalize_input_values(row[7].split("|") if row[7] else [])
    for i, val in enumerate(values):
        set_field(entries[i], val)

    notebook.select(form_frame)


tree.bind("<<TreeviewSelect>>", on_select)


def validate_form():
    if not require_login():
        return False

    xenoglossi = record_xenoglossi_entry.get().strip()
    if not xenoglossi:
        messagebox.showerror("Validation", "Ξενόγλωσση is required")
        return False

    if not record_username_entry.get().strip():
        messagebox.showerror("Validation", "Record username is required")
        return False

    if selected_id is None and not record_password_entry.get().strip():
        messagebox.showerror("Validation", "Password is required for a new record")
        return False

    if not dropdown.get().strip():
        messagebox.showerror("Validation", "Please select an option")
        return False

    conn = get_conn()
    c = conn.cursor()
    if selected_id:
        c.execute("SELECT id FROM records WHERE xenoglossi=? AND id<>?", (xenoglossi, selected_id))
    else:
        c.execute("SELECT id FROM records WHERE xenoglossi=?", (xenoglossi,))
    duplicate = c.fetchone()
    conn.close()

    if duplicate:
        messagebox.showerror("Validation", "Ξενόγλωσση must be unique")
        return False

    return True


def submit():
    global selected_id

    if not validate_form():
        return

    conn = get_conn()
    c = conn.cursor()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    xenoglossi = record_xenoglossi_entry.get().strip()
    username = record_username_entry.get().strip()
    dropdown_value = dropdown.get().strip()
    checkbox_value = int(checkbox_var.get())
    inputs_value = "|".join([get_field(entry) for entry in entries])

    if selected_id:
        if record_password_entry.get().strip():
            password_value = hash_password(record_password_entry.get().strip())
            c.execute(
                "UPDATE records SET xenoglossi=?, timestamp=?, username=?, password=?, dropdown=?, checkbox=?, inputs=? WHERE id=?",
                (xenoglossi, timestamp, username, password_value, dropdown_value, checkbox_value, inputs_value, selected_id),
            )
        else:
            c.execute(
                "UPDATE records SET xenoglossi=?, timestamp=?, username=?, dropdown=?, checkbox=?, inputs=? WHERE id=?",
                (xenoglossi, timestamp, username, dropdown_value, checkbox_value, inputs_value, selected_id),
            )
        action = "RECORD_UPDATE"
        details = f"Updated record {xenoglossi} ({username})"
    else:
        password_value = hash_password(record_password_entry.get().strip())
        c.execute(
            "INSERT INTO records (xenoglossi, timestamp, username, password, dropdown, checkbox, inputs) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (xenoglossi, timestamp, username, password_value, dropdown_value, checkbox_value, inputs_value),
        )
        selected_id = c.lastrowid
        action = "RECORD_CREATE"
        details = f"Created record {xenoglossi} ({username})"

    conn.commit()
    conn.close()

    log_action(action, details)
    load_data()
    if get_setting("auto_refresh", "1") == "1":
        refresh_analytics()
        build_report(show_message=False)
    clear_form()
    messagebox.showinfo("Success", "Record saved successfully")


def delete():
    global selected_id

    if not require_login():
        return

    if not selected_id:
        messagebox.showwarning("Delete", "Please select a record first")
        return

    confirm = messagebox.askyesno("Delete", "Are you sure you want to delete this record?")
    if not confirm:
        return

    deleted_id = selected_id
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM records WHERE id=?", (selected_id,))
    conn.commit()
    conn.close()

    log_action("RECORD_DELETE", f"Deleted record #{deleted_id}")
    load_data()
    if get_setting("auto_refresh", "1") == "1":
        refresh_analytics()
        build_report(show_message=False)
    clear_form()
    messagebox.showinfo("Delete", "Record deleted")


def export_excel():
    if not require_login():
        return

    conn = get_conn()
    c = conn.cursor()
    c.execute(f"SELECT {RECORD_SELECT_COLUMNS} FROM records ORDER BY id DESC")
    data = c.fetchall()
    conn.close()

    if not data:
        messagebox.showwarning("Export", "No data to export")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Records"

    headers = ["Ξενόγλωσση", "Timestamp", "Username", "Dropdown", "Checkbox"] + AUDITOR_FIELDS
    ws.append(headers)

    for row in data:
        input_values = normalize_input_values(row[7].split("|") if row[7] else [])
        ws.append([row[1], row[2], row[3], row[5], row[6]] + input_values)

    path = ask_save_path(
        ".xlsx",
        [("Excel files", "*.xlsx")],
        "Save Excel File",
        f"records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )

    if path:
        wb.save(path)
        log_action("EXPORT_EXCEL", f"Exported records to {path}")
        refresh_analytics()
        messagebox.showinfo("Export", "Excel file exported successfully")


def export_json():
    if not require_login():
        return

    conn = get_conn()
    c = conn.cursor()
    c.execute(f"SELECT {RECORD_SELECT_COLUMNS} FROM records ORDER BY id DESC")
    data = c.fetchall()
    conn.close()

    if not data:
        messagebox.showwarning("Export", "No data to export")
        return

    json_data = []

    for row in data:
        input_values = normalize_input_values(row[7].split("|") if row[7] else [])

        record = {
            "internal_id": row[0],
            "ξενόγλωσση": row[1],
            "timestamp": row[2],
            "username": row[3],
            "password": "hidden",
            "dropdown": row[5],
            "checkbox": bool(row[6]),
            "fields": {},
        }

        for i, field_name in enumerate(AUDITOR_FIELDS):
            record["fields"][field_name] = input_values[i]

        json_data.append(record)

    path = ask_save_path(
        ".json",
        [("JSON files", "*.json")],
        "Save JSON File",
        f"records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
    )

    if path:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, indent=4, ensure_ascii=False)

        log_action("EXPORT_JSON", f"Exported records to {path}")
        refresh_analytics()
        messagebox.showinfo("Export", "JSON file exported successfully")


# =========================
# ANALYTICS FUNCTIONS
# =========================
def refresh_analytics():
    if "kpi_vars" not in globals():
        return

    conn = get_conn()
    c = conn.cursor()

    c.execute("SELECT COUNT(*) FROM records")
    total_records = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM records WHERE checkbox=1")
    accepted_terms = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM app_users WHERE status='Active'")
    active_users = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM app_logs WHERE action LIKE 'EXPORT_%'")
    exports = c.fetchone()[0]

    kpi_vars["Total Records"].config(text=str(total_records))
    kpi_vars["Accepted Terms"].config(text=str(accepted_terms))
    kpi_vars["Active Users"].config(text=str(active_users))
    kpi_vars["Exports"].config(text=str(exports))

    clear_tree(option_tree)
    c.execute("SELECT dropdown, COUNT(*) FROM records GROUP BY dropdown ORDER BY COUNT(*) DESC, dropdown")
    option_rows = c.fetchall()
    if option_rows:
        for row in option_rows:
            option_tree.insert("", "end", values=(row[0] or "(blank)", row[1]))
    else:
        option_tree.insert("", "end", values=("No data", 0))

    clear_tree(activity_tree)
    c.execute("SELECT timestamp, username, action FROM app_logs ORDER BY id DESC LIMIT 12")
    for row in c.fetchall():
        activity_tree.insert("", "end", values=row)

    conn.close()


# =========================
# REPORT FUNCTIONS
# =========================
def get_report_rows():
    date_from = parse_iso_date(report_from_var.get(), "From")
    if date_from is False:
        return None
    date_to = parse_iso_date(report_to_var.get(), "To", end_of_day=True)
    if date_to is False:
        return None

    query = f"SELECT {RECORD_SELECT_COLUMNS} FROM records"
    clauses = []
    params = []

    username = report_username_var.get().strip()
    if username:
        clauses.append("username LIKE ?")
        params.append(f"%{username}%")

    option = report_option_var.get().strip()
    if option:
        clauses.append("dropdown = ?")
        params.append(option)

    if date_from:
        clauses.append("timestamp >= ?")
        params.append(date_from)

    if date_to:
        clauses.append("timestamp <= ?")
        params.append(date_to)

    if clauses:
        query += " WHERE " + " AND ".join(clauses)

    query += " ORDER BY id DESC"

    conn = get_conn()
    c = conn.cursor()
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()
    return rows


def build_report(show_message=True):
    if "report_tree" not in globals():
        return
    if not is_logged_in:
        return

    rows = get_report_rows()
    if rows is None:
        return

    clear_tree(report_tree)

    accepted = 0
    filled_inputs = 0
    for row in rows:
        input_values = normalize_input_values(row[7].split("|") if row[7] else [])
        filled = len([value for value in input_values if value.strip()])
        filled_inputs += filled
        accepted += 1 if row[6] else 0
        report_tree.insert(
            "",
            "end",
            values=(row[1], row[2], row[3], row[5], "Yes" if row[6] else "No", filled),
        )

    report_summary_var.set(
        f"Rows: {len(rows)} | Accepted: {accepted} | Filled inputs: {filled_inputs}"
    )

    if show_message:
        log_action("REPORT_BUILD", f"Built report with {len(rows)} rows")


def export_report_excel():
    if not require_login():
        return

    rows = get_report_rows()
    if rows is None:
        return

    if not rows:
        messagebox.showwarning("Report", "No report data to export")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.append(["Ξενόγλωσση", "Timestamp", "Username", "Dropdown", "Checkbox"] + AUDITOR_FIELDS)

    for row in rows:
        input_values = normalize_input_values(row[7].split("|") if row[7] else [])
        ws.append([row[1], row[2], row[3], row[5], row[6]] + input_values)

    path = ask_save_path(
        ".xlsx",
        [("Excel files", "*.xlsx")],
        "Save Report",
        f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )

    if path:
        wb.save(path)
        log_action("EXPORT_REPORT", f"Exported report to {path}")
        refresh_analytics()
        messagebox.showinfo("Report", "Report exported successfully")


# =========================
# NETWORK JOB FUNCTIONS
# =========================
def load_network_settings_ui():
    if "network_server_url_var" not in globals():
        return

    network_server_url_var.set(get_setting("network_server_url", "http://127.0.0.1:8765"))
    workstation_name_var.set(get_setting("workstation_name", os.environ.get("COMPUTERNAME", "CLIENT-1")))
    if not job_assignee_var.get().strip():
        job_assignee_var.set("ALL")


def save_network_settings(show_message=True):
    if not require_login():
        return False

    server_url = network_server_url_var.get().strip().rstrip("/")
    workstation = workstation_name_var.get().strip()

    if not server_url:
        messagebox.showerror("Network Jobs", "Server URL is required")
        return False

    if not server_url.startswith(("http://", "https://")):
        server_url = "http://" + server_url
        network_server_url_var.set(server_url)

    if not workstation:
        messagebox.showerror("Network Jobs", "This Client is required")
        return False

    set_setting("network_server_url", server_url)
    set_setting("workstation_name", workstation)

    if show_message:
        log_action("NETWORK_SETTINGS_SAVE", f"Saved network settings for {workstation}")
        messagebox.showinfo("Network Jobs", "Network settings saved")

    return True


def network_base_url():
    value = network_server_url_var.get().strip() or get_setting("network_server_url", "http://127.0.0.1:8765")
    value = value.rstrip("/")
    if not value.startswith(("http://", "https://")):
        value = "http://" + value
    return value


def network_request(method, path, data=None, params=None, auth=False, timeout=8):
    url = network_base_url() + path
    if params:
        url += "?" + urlparse.urlencode(params)

    body = None
    headers = {"Accept": "application/json"}
    if data is not None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json; charset=utf-8"

    req = urlrequest.Request(url, data=body, headers=headers, method=method)

    try:
        with urlrequest.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urlerror.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            payload = json.loads(raw)
            message = payload.get("error", raw)
        except json.JSONDecodeError:
            message = raw or str(exc)
        raise RuntimeError(message) from exc
    except urlerror.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc
    except TimeoutError as exc:
        raise RuntimeError("Network request timed out") from exc


def set_network_details(text):
    if "network_details_text" not in globals():
        return
    network_details_text.configure(state="normal")
    network_details_text.delete("1.0", tk.END)
    network_details_text.insert("1.0", text)
    network_details_text.configure(state="disabled")


def test_network_connection():
    if not save_network_settings(show_message=False):
        return

    try:
        result = network_request("GET", "/health", auth=False, timeout=5)
    except RuntimeError as exc:
        messagebox.showerror("Network Jobs", f"Connection failed: {exc}")
        return

    if result.get("ok"):
        register_this_client(show_message=False)
        fetch_network_clients(show_message=False)
        messagebox.showinfo("Network Jobs", "Connection successful")
    else:
        messagebox.showerror("Network Jobs", "Server did not return a healthy response")


def client_platform_label():
    return f"{platform.system()} {platform.release()} ({platform.machine()})"


def register_this_client(show_message=True):
    if not require_login():
        return False

    if not save_network_settings(show_message=False):
        return False

    client_name = workstation_name_var.get().strip()
    inventory = collect_system_inventory()
    try:
        from register_client import collect_device_diagnostics

        inventory["Diagnostics"] = collect_device_diagnostics()
    except Exception as exc:
        inventory["Diagnostics"] = {
            "items": [
                {
                    "name": "Diagnostics",
                    "status": "FAIL",
                    "details": f"Could not collect diagnostics from app registration: {exc}",
                }
            ],
            "summary": {"FAIL": 1},
        }

    try:
        result = network_request(
            "POST",
            "/clients/register",
            data={
                "name": client_name,
                "username": current_user or "",
                "role": current_role or "",
                "app_version": APP_VERSION,
                "platform": client_platform_label(),
                "inventory": inventory,
            },
            auth=True,
            timeout=8,
        )
    except RuntimeError as exc:
        if show_message:
            messagebox.showerror("Network Jobs", f"Could not register client: {exc}")
        return False

    if result.get("ok"):
        log_action("NETWORK_CLIENT_REGISTER", f"Registered client {client_name}")
        if show_message:
            messagebox.showinfo("Network Jobs", f"Client registered: {client_name}")
        return True

    if show_message:
        messagebox.showerror("Network Jobs", "Server did not accept the client registration")
    return False


def fetch_network_clients(show_message=True):
    global network_clients

    if not require_login():
        return

    if not save_network_settings(show_message=False):
        return

    try:
        result = network_request("GET", "/clients", auth=True, timeout=8)
    except RuntimeError as exc:
        if show_message:
            messagebox.showerror("Network Jobs", f"Could not fetch clients: {exc}")
        return

    clients = result.get("clients", [])
    network_clients = {client.get("name", ""): client for client in clients}

    clear_tree(network_client_tree)
    client_names = ["ALL"]

    for client in clients:
        name = client.get("name", "")
        if name:
            client_names.append(name)
        network_client_tree.insert(
            "",
            "end",
            values=(
                name,
                client.get("username", ""),
                client.get("role", ""),
                client.get("last_seen", ""),
                client.get("platform", ""),
                client.get("app_version", ""),
            ),
        )

    job_assignee_combo.configure(values=client_names)

    if job_assignee_var.get().strip() not in client_names:
        job_assignee_var.set("ALL")

    populate_device_results()
    populate_manual_tests()

    if show_message:
        messagebox.showinfo("Network Jobs", f"Fetched {len(clients)} client(s)")


def diagnostic_rows_from_clients():
    rows = []

    for client_name, client in sorted(network_clients.items()):
        inventory = client.get("inventory", {})
        diagnostics = inventory.get("Diagnostics", {}) if isinstance(inventory, dict) else {}
        items = diagnostics.get("items", []) if isinstance(diagnostics, dict) else []

        if not items:
            rows.append(
                {
                    "client": client_name,
                    "device": "Diagnostics",
                    "status": "UNKNOWN",
                    "details": "No diagnostics reported by this client",
                    "last_seen": client.get("last_seen", ""),
                    "raw": client,
                }
            )
            continue

        for item in items:
            if not isinstance(item, dict):
                continue
            rows.append(
                {
                    "client": client_name,
                    "device": item.get("name", ""),
                    "status": item.get("status", "UNKNOWN"),
                    "details": item.get("details", ""),
                    "last_seen": client.get("last_seen", ""),
                    "raw": item,
                }
            )

    return rows


def set_device_details(text):
    if "device_details_text" not in globals():
        return
    device_details_text.configure(state="normal")
    device_details_text.delete("1.0", tk.END)
    device_details_text.insert("1.0", text)
    device_details_text.configure(state="disabled")


def browse_manual_test_path(target_var):
    path = filedialog.askdirectory(title="Select drive or mounted media")
    if path:
        target_var.set(path)


def detect_windows_test_paths():
    if not platform.system().lower().startswith("win"):
        messagebox.showwarning("Manual Tests", "Windows device tests can only auto-detect Windows drive paths")
        return

    try:
        from register_client import test_path_candidates

        optical_candidates = test_path_candidates("optical")
        removable_candidates = test_path_candidates("removable")
    except Exception as exc:
        messagebox.showerror("Manual Tests", f"Could not detect paths: {exc}")
        return

    if optical_candidates and not manual_windows_optical_path_var.get().strip():
        manual_windows_optical_path_var.set(optical_candidates[0])
    if removable_candidates and not manual_windows_usb_path_var.get().strip():
        manual_windows_usb_path_var.set(removable_candidates[0])

    manual_windows_test_status_var.set(
        f"Detected optical: {optical_candidates[0] if optical_candidates else 'none'} | "
        f"USB/SD: {removable_candidates[0] if removable_candidates else 'none'}"
    )


def parse_usb_port_count():
    value = manual_windows_usb_expected_var.get().strip()
    if not value:
        messagebox.showerror("Manual Tests", "USB ports is required for one-by-one USB testing")
        return None
    try:
        count = int(value)
    except ValueError:
        messagebox.showerror("Manual Tests", "USB ports must be a number")
        return None
    if count <= 0:
        messagebox.showerror("Manual Tests", "USB ports must be at least 1")
        return None
    return count


def test_item(name, status, details=""):
    return {"name": name, "status": status, "details": details}


def usb_device_snapshot():
    try:
        from register_client import usb_device_listing

        return set(usb_device_listing())
    except Exception:
        return set()


def wait_for_usb_change(baseline, seconds=7):
    latest = set(baseline)
    for remaining in range(seconds, 0, -1):
        manual_windows_test_status_var.set(f"Waiting for Windows USB detection... {remaining}s")
        root.update_idletasks()
        root.update()
        time.sleep(1)
        latest = usb_device_snapshot()
        if latest != baseline:
            break
    return latest


def run_usb_ports_one_by_one_gui(port_count):
    messagebox.showinfo(
        "USB one-by-one test",
        "Use one known-good USB mouse or flash drive.\n\n"
        "The app will test each USB port separately. Remove the USB device before every port baseline.",
    )

    items = []
    passed = 0
    failed = 0
    skipped = 0

    for port_index in range(1, port_count + 1):
        name = f"Test USB port {port_index}"
        manual_windows_test_status_var.set(f"Testing USB port {port_index}/{port_count}")

        if not messagebox.askokcancel(
            name,
            f"Step 1/{port_count}\n\nRemove the USB test device from all USB ports, then press OK.",
        ):
            items.append(test_item(name, "UNKNOWN", "Skipped before baseline"))
            skipped += 1
            continue

        baseline = usb_device_snapshot()

        if not messagebox.askokcancel(
            name,
            f"Plug the known-good USB device into USB port {port_index}, wait for Windows to detect it, then press OK.",
        ):
            items.append(test_item(name, "UNKNOWN", "Skipped before device insertion"))
            skipped += 1
            continue

        after = wait_for_usb_change(baseline)
        added = sorted(after - baseline)
        removed = sorted(baseline - after)

        if added or len(after) > len(baseline):
            detail_parts = [f"Windows detected USB change on port {port_index}"]
            if added:
                detail_parts.append("added: " + "; ".join(added[:5]))
            if removed:
                detail_parts.append("removed: " + "; ".join(removed[:3]))
            items.append(test_item(name, "OK", " | ".join(detail_parts)))
            passed += 1
            continue

        answer = messagebox.askyesnocancel(
            name,
            "Windows did not report a new USB device automatically.\n\n"
            "Did the USB device work physically in this port?",
        )
        if answer is True:
            items.append(test_item(name, "OK", "Technician confirmed device worked; Windows USB list did not change"))
            passed += 1
        elif answer is False:
            items.append(test_item(name, "FAIL", "No USB detection and technician reported the port did not work"))
            failed += 1
        else:
            items.append(test_item(name, "UNKNOWN", "No USB detection; technician skipped manual confirmation"))
            skipped += 1

    if failed:
        summary_status = "FAIL"
    elif skipped:
        summary_status = "WARN"
    else:
        summary_status = "OK"

    items.insert(
        0,
        test_item(
            "Test USB ports one by one",
            summary_status,
            f"Ports tested: {port_count} | OK: {passed} | FAIL: {failed} | UNKNOWN: {skipped}",
        ),
    )
    return items


def post_client_registration(server_url, payload, timeout=8):
    url = server_url.rstrip("/") + "/clients/register"
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urlrequest.Request(
        url,
        data=body,
        method="POST",
        headers={"Content-Type": "application/json; charset=utf-8", "Accept": "application/json"},
    )

    try:
        with urlrequest.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urlerror.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            payload = json.loads(raw)
            message = payload.get("error", raw)
        except json.JSONDecodeError:
            message = raw or str(exc)
        raise RuntimeError(message) from exc
    except urlerror.URLError as exc:
        raise RuntimeError(str(exc.reason)) from exc
    except TimeoutError as exc:
        raise RuntimeError("Network request timed out") from exc


def run_windows_sound_test_gui():
    try:
        import tempfile
        from register_client import play_wav_file, status_item, write_stereo_channel_wav

        fd, path = tempfile.mkstemp(prefix="manual_windows_sound_", suffix=".wav")
        os.close(fd)
        try:
            write_stereo_channel_wav(path)
            played, playback_details = play_wav_file(path)
        finally:
            try:
                os.remove(path)
            except OSError:
                pass
    except Exception as exc:
        return {"name": "Test sound channels", "status": "UNKNOWN", "details": f"Could not play sound test: {exc}"}

    if not played:
        return {"name": "Test sound channels", "status": "UNKNOWN", "details": playback_details}

    answer = messagebox.askyesnocancel(
        "Sound left/right channel test",
        "Did you hear the sound correctly?\n\nFirst LEFT channel, then RIGHT channel.",
    )
    if answer is True:
        status = "OK"
        details = "User confirmed left/right sound output"
    elif answer is False:
        status = "FAIL"
        details = "User reported a left/right sound problem"
    else:
        status = "UNKNOWN"
        details = "User skipped sound confirmation"
    return status_item("Test sound channels", status, f"{details} | {playback_details}")


def summarize_test_items(items):
    summary = {"OK": 0, "WARN": 0, "FAIL": 0, "UNKNOWN": 0}
    for item in items:
        status = item.get("status", "UNKNOWN")
        summary[status] = summary.get(status, 0) + 1
    return summary


def set_windows_test_controls_running(running):
    state = "disabled" if running else "normal"
    if "manual_windows_run_btn" in globals():
        manual_windows_run_btn.configure(state=state)


def run_windows_device_tests_from_tab():
    if not require_login():
        return

    if not platform.system().lower().startswith("win"):
        messagebox.showwarning("Manual Tests", "These four required tests are Windows-only")
        return

    if not save_network_settings(show_message=False):
        return

    usb_port_count = parse_usb_port_count()
    if usb_port_count is None:
        return

    server_url = network_base_url()
    client_name = workstation_name_var.get().strip()
    optical_path = manual_windows_optical_path_var.get().strip()
    usb_path = manual_windows_usb_path_var.get().strip()

    set_windows_test_controls_running(True)
    manual_windows_test_status_var.set("Starting USB port one-by-one test...")
    usb_items = run_usb_ports_one_by_one_gui(usb_port_count)
    manual_windows_test_status_var.set("Running optical and removable media tests...")

    def worker():
        try:
            from types import SimpleNamespace
            from register_client import (
                pc_auditor_optical_read_test,
                pc_auditor_removable_read_test,
            )

            args = SimpleNamespace(
                non_interactive_tests=True,
                optical_test_path=optical_path,
                usb_test_path=usb_path,
            )
            items = list(usb_items) + [
                pc_auditor_optical_read_test(args),
                pc_auditor_removable_read_test(args),
            ]
            root.after(0, lambda: finish_windows_device_tests_from_tab(items, server_url, client_name))
        except Exception as exc:
            root.after(0, lambda exc=exc: handle_windows_device_tests_error(exc))

    threading.Thread(target=worker, daemon=True).start()


def finish_windows_device_tests_from_tab(items, server_url, client_name):
    manual_windows_test_status_var.set("Playing left/right sound test...")
    items.append(run_windows_sound_test_gui())

    diagnostics = {"items": items, "summary": summarize_test_items(items)}
    inventory = collect_system_inventory()
    inventory["Diagnostics"] = diagnostics
    payload = {
        "name": client_name,
        "username": current_user or "",
        "role": current_role or "",
        "app_version": APP_VERSION,
        "platform": client_platform_label(),
        "inventory": inventory,
    }

    manual_windows_test_status_var.set("Uploading Windows test results...")

    def upload_worker():
        try:
            result = post_client_registration(server_url, payload, timeout=8)
            root.after(0, lambda: complete_windows_device_tests_upload(payload, result))
        except Exception as exc:
            root.after(0, lambda exc=exc: handle_windows_device_tests_error(exc))

    threading.Thread(target=upload_worker, daemon=True).start()


def complete_windows_device_tests_upload(payload, result):
    global network_clients

    client_name = payload.get("name", "")
    if client_name:
        network_clients[client_name] = {
            "name": client_name,
            "username": payload.get("username", ""),
            "role": payload.get("role", ""),
            "app_version": payload.get("app_version", ""),
            "platform": payload.get("platform", ""),
            "inventory": payload.get("inventory", {}),
            "last_seen": result.get("last_seen", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        }

    populate_device_results()
    populate_manual_tests()
    set_windows_test_controls_running(False)

    diagnostics = payload.get("inventory", {}).get("Diagnostics", {})
    summary = diagnostics.get("summary", {}) if isinstance(diagnostics, dict) else {}
    manual_windows_test_status_var.set(
        f"Windows tests saved | OK: {summary.get('OK', 0)} | WARN: {summary.get('WARN', 0)} | "
        f"FAIL: {summary.get('FAIL', 0)} | UNKNOWN: {summary.get('UNKNOWN', 0)}"
    )
    log_action("WINDOWS_DEVICE_TESTS", f"Saved Windows device tests for {client_name}")
    messagebox.showinfo("Manual Tests", "Windows device tests completed and saved")


def handle_windows_device_tests_error(exc):
    set_windows_test_controls_running(False)
    manual_windows_test_status_var.set(f"Windows tests failed: {exc}")
    messagebox.showerror("Manual Tests", f"Windows device tests failed: {exc}")


def manual_rows_from_clients():
    rows = []
    for row in diagnostic_rows_from_clients():
        device = str(row.get("device", ""))
        details = str(row.get("details", ""))
        if device == "Diagnostics" and details == "No diagnostics reported by this client":
            continue
        rows.append(row)
    return rows


def set_manual_details(text):
    if "manual_details_text" not in globals():
        return
    manual_details_text.configure(state="normal")
    manual_details_text.delete("1.0", tk.END)
    manual_details_text.insert("1.0", text)
    manual_details_text.configure(state="disabled")


def populate_manual_tests():
    if "manual_tests_tree" not in globals():
        return

    clear_tree(manual_tests_tree)
    status_filter = manual_status_filter_var.get().strip()
    rows = manual_rows_from_clients()
    shown = 0
    counts = {"OK": 0, "WARN": 0, "FAIL": 0, "UNKNOWN": 0}

    for row in rows:
        status = row["status"] or "UNKNOWN"
        counts[status] = counts.get(status, 0) + 1

        if status_filter != "All" and status != status_filter:
            continue

        manual_tests_tree.insert(
            "",
            "end",
            values=(row["client"], row["device"], status, row["details"], row["last_seen"]),
        )
        shown += 1

    manual_summary_var.set(
        f"Test results: {shown}/{len(rows)} | OK: {counts.get('OK', 0)} | "
        f"WARN: {counts.get('WARN', 0)} | FAIL: {counts.get('FAIL', 0)} | "
        f"UNKNOWN: {counts.get('UNKNOWN', 0)}"
    )


def refresh_manual_tests():
    fetch_network_clients(show_message=False)
    load_manual_requirement(show_message=False)
    populate_manual_tests()


def load_manual_requirement(show_message=True):
    if not require_login():
        return

    if not save_network_settings(show_message=False):
        return

    try:
        result = network_request("GET", "/settings", timeout=8)
    except RuntimeError as exc:
        if show_message:
            messagebox.showerror("Manual Tests", f"Could not load requirement: {exc}")
        return

    settings = result.get("settings", {})
    manual_required_var.set(bool(settings.get("manual_tests_required", False)))

    if show_message:
        state = "required" if manual_required_var.get() else "optional"
        messagebox.showinfo("Manual Tests", f"Manual tests are now {state}")


def save_manual_requirement():
    if not require_admin():
        return

    if not save_network_settings(show_message=False):
        return

    required = bool(manual_required_var.get())

    try:
        network_request("POST", "/settings", data={"manual_tests_required": required}, timeout=8)
    except RuntimeError as exc:
        messagebox.showerror("Manual Tests", f"Could not save requirement: {exc}")
        return

    log_action("MANUAL_REQUIREMENT", f"Manual tests required={required}")
    state = "required" if required else "optional"
    messagebox.showinfo("Manual Tests", f"Manual tests are now {state}")


def on_manual_test_select(event=None):
    item = manual_tests_tree.focus()
    if not item:
        return
    values = manual_tests_tree.item(item).get("values", [])
    if not values:
        return

    client_name, test_name, status, note, last_seen = values
    detail_payload = {
        "client": client_name,
        "test": test_name,
        "status": status,
        "note": note,
        "last_seen": last_seen,
    }
    set_manual_details(json.dumps(detail_payload, ensure_ascii=False, indent=2))


def populate_device_results():
    if "device_results_tree" not in globals():
        return

    clear_tree(device_results_tree)
    status_filter = device_status_filter_var.get().strip()
    rows = diagnostic_rows_from_clients()
    shown = 0
    counts = {"OK": 0, "WARN": 0, "FAIL": 0, "UNKNOWN": 0}

    for row in rows:
        status = row["status"] or "UNKNOWN"
        counts[status] = counts.get(status, 0) + 1

        if status_filter != "All" and status != status_filter:
            continue

        device_results_tree.insert(
            "",
            "end",
            values=(row["client"], row["device"], status, row["details"], row["last_seen"]),
        )
        shown += 1

    total_clients = len(network_clients)
    device_summary_var.set(
        f"Clients: {total_clients} | Results: {shown}/{len(rows)} | OK: {counts.get('OK', 0)} | "
        f"WARN: {counts.get('WARN', 0)} | FAIL: {counts.get('FAIL', 0)} | UNKNOWN: {counts.get('UNKNOWN', 0)}"
    )


def refresh_device_results():
    fetch_network_clients(show_message=False)
    populate_device_results()


def on_device_result_select(event=None):
    item = device_results_tree.focus()
    if not item:
        return
    values = device_results_tree.item(item).get("values", [])
    if not values:
        return

    client_name, device, status, details, last_seen = values
    client = network_clients.get(client_name, {})
    inventory = client.get("inventory", {}) if isinstance(client, dict) else {}
    diagnostics = inventory.get("Diagnostics", {}) if isinstance(inventory, dict) else {}

    detail_payload = {
        "client": client_name,
        "device": device,
        "status": status,
        "details": details,
        "last_seen": last_seen,
        "diagnostics_summary": diagnostics.get("summary", {}) if isinstance(diagnostics, dict) else {},
    }
    set_device_details(json.dumps(detail_payload, ensure_ascii=False, indent=2))


def on_network_client_select(event=None):
    item = network_client_tree.focus()
    if not item:
        return

    values = network_client_tree.item(item).get("values", [])
    if not values:
        return

    client_name = values[0]
    if client_name:
        job_assignee_var.set(client_name)
        client = network_clients.get(client_name, {})
        set_network_details(json.dumps(client, ensure_ascii=False, indent=2))


def selected_network_client():
    item = network_client_tree.focus()
    if not item:
        return None
    values = network_client_tree.item(item).get("values", [])
    if not values:
        return None
    return network_clients.get(values[0])


def load_selected_client_info_to_form():
    client = selected_network_client()
    if not client:
        messagebox.showwarning("Network Jobs", "Select a client first")
        return

    inventory = client.get("inventory", {})
    if not inventory:
        messagebox.showwarning("Network Jobs", "Selected client has not sent inventory details yet")
        return

    clear_form()
    for index, field_name in enumerate(AUDITOR_FIELDS):
        value = inventory.get(field_name, "")
        if value:
            set_field(entries[index], value)

    serial = inventory.get("Serial number", "")
    if serial:
        record_xenoglossi_entry.insert(0, serial)
        record_username_entry.insert(0, serial)
    else:
        record_xenoglossi_entry.insert(0, client.get("name", ""))
        record_username_entry.insert(0, client.get("name", ""))

    notebook.select(form_frame)


def collect_current_form_payload():
    values = [get_field(entry) for entry in entries]
    return {
        "xenoglossi": record_xenoglossi_entry.get().strip(),
        "record_username": record_username_entry.get().strip(),
        "dropdown": dropdown.get().strip(),
        "checkbox": bool(checkbox_var.get()),
        "inputs": values,
        "fields": {field_name: values[index] for index, field_name in enumerate(AUDITOR_FIELDS)},
    }


def load_payload_to_form(payload):
    clear_form()

    xenoglossi = str(payload.get("xenoglossi", "")).strip()
    record_username = str(payload.get("record_username", "")).strip()
    dropdown_value = str(payload.get("dropdown", "")).strip()
    inputs = payload.get("inputs")

    if not isinstance(inputs, list):
        fields = payload.get("fields", {})
        inputs = [fields.get(field_name, "") for field_name in AUDITOR_FIELDS] if isinstance(fields, dict) else []
    inputs = normalize_input_values(inputs)

    if xenoglossi:
        record_xenoglossi_entry.insert(0, xenoglossi)
    if record_username:
        record_username_entry.insert(0, record_username)
    if dropdown_value:
        dropdown.set(dropdown_value)

    checkbox_var.set(bool(payload.get("checkbox", False)))

    for index, value in enumerate(inputs):
        set_field(entries[index], value)

    notebook.select(form_frame)


def selected_network_job():
    item = network_tree.focus()
    if not item:
        return None
    values = network_tree.item(item).get("values", [])
    if not values:
        return None
    return network_jobs.get(str(values[0]))


def show_selected_network_job_details(event=None):
    job = selected_network_job()
    if not job:
        set_network_details("")
        return

    set_network_details(json.dumps(job, ensure_ascii=False, indent=2))


def fetch_network_jobs(show_message=True):
    global network_jobs

    if not require_login():
        return

    if not save_network_settings(show_message=False):
        return

    register_this_client(show_message=False)

    assignee = workstation_name_var.get().strip()
    status = job_status_filter_var.get().strip() or "Queued"

    try:
        result = network_request(
            "GET",
            "/jobs",
            params={"assignee": assignee, "status": status},
            auth=True,
            timeout=8,
        )
    except RuntimeError as exc:
        if show_message:
            messagebox.showerror("Network Jobs", f"Could not fetch jobs: {exc}")
        else:
            set_network_details(f"Network unavailable: {exc}")
        return

    jobs = result.get("jobs", [])
    network_jobs = {str(job["id"]): job for job in jobs}
    clear_tree(network_tree)

    for job in jobs:
        network_tree.insert(
            "",
            "end",
            values=(
                job.get("id"),
                job.get("title", ""),
                job.get("assignee", ""),
                job.get("status", ""),
                job.get("created_by", ""),
                job.get("created_at", ""),
                job.get("updated_at", ""),
            ),
        )

    set_network_details(f"Fetched {len(jobs)} job(s) for {assignee}.")

    if show_message:
        messagebox.showinfo("Network Jobs", f"Fetched {len(jobs)} job(s)")


def publish_current_form_job():
    if not require_admin():
        return

    if not save_network_settings(show_message=False):
        return

    payload = collect_current_form_payload()
    xenoglossi = payload["xenoglossi"]
    filled_values = [value for value in payload["inputs"] if value]

    if not xenoglossi:
        messagebox.showerror("Network Jobs", "Ξενόγλωσση is required before publishing")
        return

    if not filled_values:
        messagebox.showerror("Network Jobs", "Fill at least one PC Auditor field before publishing")
        return

    title = job_title_var.get().strip() or f"PC Auditor - {xenoglossi}"
    assignee = job_assignee_var.get().strip() or "ALL"

    try:
        result = network_request(
            "POST",
            "/jobs",
            data={
                "title": title,
                "assignee": assignee,
                "created_by": current_user,
                "payload": payload,
            },
            auth=True,
            timeout=8,
        )
    except RuntimeError as exc:
        messagebox.showerror("Network Jobs", f"Could not publish job: {exc}")
        return

    job_id = result.get("job_id")
    log_action("NETWORK_JOB_CREATE", f"Published job #{job_id} for {assignee}")
    messagebox.showinfo("Network Jobs", f"Job #{job_id} published")
    fetch_network_clients(show_message=False)
    fetch_network_jobs(show_message=False)


def update_network_job_status(job_id, status, notes=""):
    result = network_request(
        "POST",
        f"/jobs/{job_id}/status",
        data={"status": status, "actor": current_user or workstation_name_var.get().strip(), "notes": notes},
        auth=True,
        timeout=8,
    )
    log_action("NETWORK_JOB_STATUS", f"Job #{job_id} -> {status}")
    return result


def load_selected_network_job(mark_loaded=True):
    job = selected_network_job()
    if not job:
        messagebox.showwarning("Network Jobs", "Select a job first")
        return None

    load_payload_to_form(job.get("payload", {}))

    if mark_loaded and job.get("status") == "Queued":
        try:
            update_network_job_status(job["id"], "Loaded", f"Loaded on {workstation_name_var.get().strip()}")
        except RuntimeError as exc:
            messagebox.showwarning("Network Jobs", f"Loaded locally, but status update failed: {exc}")

    fetch_network_jobs(show_message=False)
    return job


def apply_selected_network_job():
    job = load_selected_network_job(mark_loaded=True)
    if not job:
        return
    apply_keyboard_actions()


def complete_selected_network_job():
    job = selected_network_job()
    if not job:
        messagebox.showwarning("Network Jobs", "Select a job first")
        return

    try:
        update_network_job_status(job["id"], "Completed", f"Completed on {workstation_name_var.get().strip()}")
    except RuntimeError as exc:
        messagebox.showerror("Network Jobs", f"Could not complete job: {exc}")
        return

    messagebox.showinfo("Network Jobs", "Job marked completed")
    fetch_network_jobs(show_message=False)


def fail_selected_network_job():
    job = selected_network_job()
    if not job:
        messagebox.showwarning("Network Jobs", "Select a job first")
        return

    try:
        update_network_job_status(job["id"], "Failed", f"Failed on {workstation_name_var.get().strip()}")
    except RuntimeError as exc:
        messagebox.showerror("Network Jobs", f"Could not mark job failed: {exc}")
        return

    messagebox.showinfo("Network Jobs", "Job marked failed")
    fetch_network_jobs(show_message=False)


network_tree.bind("<<TreeviewSelect>>", show_selected_network_job_details)
network_client_tree.bind("<<TreeviewSelect>>", on_network_client_select)
device_results_tree.bind("<<TreeviewSelect>>", on_device_result_select)
manual_tests_tree.bind("<<TreeviewSelect>>", on_manual_test_select)


# =========================
# USER FUNCTIONS
# =========================
def clear_user_form():
    global selected_user_id
    selected_user_id = None
    user_username_var.set("")
    user_password_entry.delete(0, tk.END)
    user_role_var.set("User")
    user_active_var.set(True)
    if user_tree.selection():
        user_tree.selection_remove(user_tree.selection())


def load_users():
    if "user_tree" not in globals():
        return

    clear_tree(user_tree)
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, username, role, status, created_at, last_login FROM app_users ORDER BY username")
    for row in c.fetchall():
        user_tree.insert("", "end", values=(row[0], row[1], row[2], row[3], row[4] or "", row[5] or ""))
    conn.close()


def on_user_select(event):
    global selected_user_id

    item = user_tree.focus()
    if not item:
        return

    values = user_tree.item(item)["values"]
    selected_user_id = values[0]
    user_username_var.set(values[1])
    user_role_var.set(values[2])
    user_active_var.set(values[3] == "Active")
    user_password_entry.delete(0, tk.END)


user_tree.bind("<<TreeviewSelect>>", on_user_select)


def add_user():
    if not require_admin():
        return

    username = user_username_var.get().strip()
    password = user_password_entry.get().strip()
    role = user_role_var.get().strip()
    status = "Active" if user_active_var.get() else "Disabled"

    if not username or not password:
        messagebox.showerror("Users", "Username and password are required")
        return

    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            "INSERT INTO app_users (username, password, role, status, created_at) VALUES (?, ?, ?, ?, ?)",
            (username, hash_password(password), role, status, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        conn.commit()
        conn.close()
    except sqlite3.IntegrityError:
        messagebox.showerror("Users", "Username already exists")
        return

    log_action("USER_CREATE", f"Created user {username}")
    load_users()
    refresh_analytics()
    clear_user_form()
    messagebox.showinfo("Users", "User created")


def update_user():
    if not require_admin():
        return

    if not selected_user_id:
        messagebox.showwarning("Users", "Please select a user first")
        return

    username = user_username_var.get().strip()
    role = user_role_var.get().strip()
    status = "Active" if user_active_var.get() else "Disabled"
    password = user_password_entry.get().strip()

    if not username:
        messagebox.showerror("Users", "Username is required")
        return

    if username == current_user and status != "Active":
        messagebox.showwarning("Users", "You cannot disable the current account")
        return

    try:
        conn = get_conn()
        c = conn.cursor()
        if password:
            c.execute(
                "UPDATE app_users SET username=?, password=?, role=?, status=? WHERE id=?",
                (username, hash_password(password), role, status, selected_user_id),
            )
        else:
            c.execute(
                "UPDATE app_users SET username=?, role=?, status=? WHERE id=?",
                (username, role, status, selected_user_id),
            )
        conn.commit()
        conn.close()
    except sqlite3.IntegrityError:
        messagebox.showerror("Users", "Username already exists")
        return

    log_action("USER_UPDATE", f"Updated user {username}")
    load_users()
    refresh_analytics()
    clear_user_form()
    messagebox.showinfo("Users", "User updated")


def delete_user():
    if not require_admin():
        return

    if not selected_user_id:
        messagebox.showwarning("Users", "Please select a user first")
        return

    username = user_username_var.get().strip()
    if username == current_user:
        messagebox.showwarning("Users", "You cannot delete the current account")
        return

    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM app_users WHERE role='Admin' AND status='Active' AND id<>?", (selected_user_id,))
    remaining_admins = c.fetchone()[0]
    if remaining_admins == 0:
        conn.close()
        messagebox.showwarning("Users", "At least one active admin account is required")
        return

    confirm = messagebox.askyesno("Users", f"Delete user {username}?")
    if not confirm:
        conn.close()
        return

    c.execute("DELETE FROM app_users WHERE id=?", (selected_user_id,))
    conn.commit()
    conn.close()

    log_action("USER_DELETE", f"Deleted user {username}")
    load_users()
    refresh_analytics()
    clear_user_form()
    messagebox.showinfo("Users", "User deleted")


# =========================
# SETTINGS FUNCTIONS
# =========================
def load_settings_ui():
    if "setting_system_name_var" not in globals():
        return

    setting_system_name_var.set(get_setting("system_name", APP_TITLE))
    setting_export_folder_var.set(get_setting("default_export_folder", BASE_DIR))
    setting_keyboard_delay_var.set(get_setting("keyboard_delay", "80"))
    setting_auto_refresh_var.set(get_setting("auto_refresh", "1") == "1")
    load_batch_files()


def choose_export_folder():
    folder = filedialog.askdirectory(initialdir=get_default_export_folder(), title="Choose export folder")
    if folder:
        setting_export_folder_var.set(folder)


def save_settings():
    if not require_admin():
        return

    system_name = setting_system_name_var.get().strip() or APP_TITLE
    export_folder = setting_export_folder_var.get().strip() or BASE_DIR
    delay = setting_keyboard_delay_var.get().strip() or "80"

    if not os.path.isdir(export_folder):
        messagebox.showerror("Settings", "Export folder does not exist")
        return

    try:
        delay_int = int(delay)
        if delay_int < 0 or delay_int > 5000:
            raise ValueError
    except ValueError:
        messagebox.showerror("Settings", "Keyboard delay must be a number from 0 to 5000")
        return

    set_setting("system_name", system_name)
    set_setting("default_export_folder", export_folder)
    set_setting("keyboard_delay", delay_int)
    set_setting("auto_refresh", "1" if setting_auto_refresh_var.get() else "0")

    root.title(system_name)
    title_label.config(text=system_name)
    log_action("SETTINGS_SAVE", "Saved application settings")
    messagebox.showinfo("Settings", "Settings saved")


def reset_settings():
    if not require_admin():
        return

    confirm = messagebox.askyesno("Settings", "Reset settings to defaults?")
    if not confirm:
        return

    for key, value in DEFAULT_SETTINGS.items():
        set_setting(key, value)
    load_settings_ui()
    root.title(APP_TITLE)
    title_label.config(text=APP_TITLE)
    log_action("SETTINGS_RESET", "Reset application settings")
    messagebox.showinfo("Settings", "Defaults restored")


def available_batch_files():
    files = []
    try:
        names = os.listdir(BASE_DIR)
    except OSError:
        return files

    for name in sorted(names, key=str.lower):
        path = os.path.join(BASE_DIR, name)
        if os.path.isfile(path) and name.lower().endswith(".bat"):
            files.append(name)
    return files


def load_batch_files():
    if "batch_tree" not in globals():
        return

    clear_tree(batch_tree)
    files = available_batch_files()
    for name in files:
        path = os.path.join(BASE_DIR, name)
        try:
            stat = os.stat(path)
            modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            size = f"{stat.st_size} bytes"
        except OSError:
            modified = ""
            size = ""
        batch_tree.insert("", "end", values=(name, modified, size))

    batch_status_var.set(f"{len(files)} batch file(s) found")


def selected_batch_file_names():
    names = []
    if "batch_tree" not in globals():
        return names
    for item in batch_tree.selection():
        values = batch_tree.item(item).get("values", [])
        if values:
            names.append(str(values[0]))
    return names


def resolve_batch_path(name):
    base = os.path.abspath(BASE_DIR)
    path = os.path.abspath(os.path.join(base, name))
    if not path.lower().endswith(".bat"):
        raise ValueError("Only .bat files can be launched")
    if os.path.commonpath([base, path]) != base:
        raise ValueError("Batch file must be inside the application folder")
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    return path


def launch_batch_file(name):
    if not platform.system().lower().startswith("win"):
        raise RuntimeError("Batch launcher is available only on Windows")

    path = resolve_batch_path(name)
    cmd = os.environ.get("COMSPEC", "cmd.exe")
    creationflags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
    env = os.environ.copy()
    env["RUN_BACKGROUND"] = "1"
    process = subprocess.Popen(
        [cmd, "/c", path],
        cwd=BASE_DIR,
        creationflags=creationflags,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        env=env,
    )
    return path, process.pid


def run_batch_file_names(names):
    if not names:
        messagebox.showwarning("Batch Launcher", "Select at least one batch file")
        return

    launched = []
    errors = []
    for name in names:
        try:
            path, pid = launch_batch_file(name)
            launched.append(f"{os.path.basename(path)} pid={pid}")
            time.sleep(0.2)
        except Exception as exc:
            errors.append(f"{name}: {exc}")

    if launched:
        batch_status_var.set(f"Started background: {', '.join(launched)}")
        log_action("BATCH_LAUNCH_BACKGROUND", f"Started background: {', '.join(launched)}")

    if errors:
        messagebox.showerror("Batch Launcher", "\n".join(errors[:8]))
    elif launched:
        messagebox.showinfo("Batch Launcher", f"Started {len(launched)} batch file(s) in background")


def run_selected_batch_files():
    if not require_admin():
        return
    run_batch_file_names(selected_batch_file_names())


def run_all_batch_files():
    if not require_admin():
        return

    names = available_batch_files()
    if not names:
        messagebox.showwarning("Batch Launcher", "No batch files found")
        return

    confirm = messagebox.askyesno(
        "Batch Launcher",
        f"This will start {len(names)} hidden background batch process(es).\n\nContinue?",
    )
    if not confirm:
        return

    run_batch_file_names(names)


def backup_database():
    if not require_admin():
        return

    if not os.path.exists(DB_NAME):
        messagebox.showwarning("Backup", "Database file does not exist yet")
        return

    path = ask_save_path(
        ".db",
        [("SQLite database", "*.db"), ("All files", "*.*")],
        "Backup Database",
        f"app_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
    )

    if path:
        shutil.copy2(DB_NAME, path)
        set_setting("last_backup", path)
        log_action("DB_BACKUP", f"Database backup saved to {path}")
        messagebox.showinfo("Backup", "Database backup completed")


# =========================
# LOG FUNCTIONS
# =========================
def get_log_rows():
    query = "SELECT id, timestamp, username, action, details FROM app_logs"
    clauses = []
    params = []

    user = log_user_var.get().strip()
    action = log_action_var.get().strip()

    if user:
        clauses.append("username LIKE ?")
        params.append(f"%{user}%")
    if action:
        clauses.append("action LIKE ?")
        params.append(f"%{action}%")

    if clauses:
        query += " WHERE " + " AND ".join(clauses)

    query += " ORDER BY id DESC"

    conn = get_conn()
    c = conn.cursor()
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()
    return rows


def load_logs():
    if "log_tree" not in globals():
        return

    clear_tree(log_tree)
    for row in get_log_rows():
        log_tree.insert("", "end", values=row)


def export_logs_excel():
    if not require_admin():
        return

    rows = get_log_rows()
    if not rows:
        messagebox.showwarning("Logs", "No logs to export")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Logs"
    ws.append(["ID", "Timestamp", "User", "Action", "Details"])
    for row in rows:
        ws.append(list(row))

    path = ask_save_path(
        ".xlsx",
        [("Excel files", "*.xlsx")],
        "Export Logs",
        f"logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )

    if path:
        wb.save(path)
        log_action("EXPORT_LOGS", f"Exported logs to {path}")
        load_logs()
        refresh_analytics()
        messagebox.showinfo("Logs", "Logs exported successfully")


def clear_logs():
    if not require_admin():
        return

    confirm = messagebox.askyesno("Logs", "Clear all logs?")
    if not confirm:
        return

    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM app_logs")
    conn.commit()
    conn.close()

    log_action("LOGS_CLEAR", "Cleared audit log")
    load_logs()
    refresh_analytics()
    messagebox.showinfo("Logs", "Logs cleared")


# =========================
# KEYBOARD AUTOMATION
# =========================
def clipboard_text():
    try:
        return True, root.clipboard_get()
    except tk.TclError:
        return False, ""


def restore_clipboard(had_text, text):
    try:
        root.clipboard_clear()
        if had_text:
            root.clipboard_append(text)
        root.update()
    except tk.TclError:
        pass


def send_text_by_paste(value):
    try:
        root.clipboard_clear()
        root.clipboard_append(value)
        root.update()
        pyautogui.hotkey("ctrl", "v")
        time.sleep(max(0.04, pyautogui.PAUSE))
    except Exception:
        pyautogui.write(value, interval=0.03)


def apply_keyboard_actions(event=None):
    if not require_login():
        return

    if pyautogui is None:
        messagebox.showerror(
            "Missing library",
            "pyautogui is not installed. Run: pip install pyautogui",
        )
        return

    raw_values = [
        get_field(entry)
        for field_name, entry in zip(AUDITOR_FIELDS, entries)
        if field_name not in KEYBOARD_SKIP_FIELDS
    ]
    filled_values = [value for value in raw_values if value]
    values = raw_values if tab_empty_var.get() else filled_values

    if not filled_values:
        messagebox.showwarning("Apply", "There are no filled input fields")
        return

    messagebox.showinfo(
        "Apply",
        "Press OK, then click the first target field in PC Auditor within 2 seconds.",
    )

    root.iconify()
    time.sleep(2)

    pyautogui.FAILSAFE = True
    try:
        pyautogui.PAUSE = int(get_setting("keyboard_delay", "80")) / 1000
    except ValueError:
        pyautogui.PAUSE = 0.08

    had_clipboard_text, original_clipboard_text = clipboard_text()

    try:
        for value in values:
            if value:
                action = value.lower().replace(" ", "")

                if "+" in action:
                    keys = action.split("+")
                    pyautogui.hotkey(*keys)
                elif action in pyautogui.KEYBOARD_KEYS:
                    pyautogui.press(action)
                else:
                    send_text_by_paste(value)
            pyautogui.press("tab")
            time.sleep(max(0.04, pyautogui.PAUSE))
        log_action("APPLY_KEYS", f"Applied {len(values)} keyboard actions")
    except Exception as exc:
        messagebox.showerror("Apply error", str(exc))
    finally:
        restore_clipboard(had_clipboard_text, original_clipboard_text)
        root.deiconify()


root.bind("<F1>", apply_keyboard_actions)


# =========================
# NOTEBOOK EVENTS
# =========================
def on_tab_changed(event):
    tab_text = notebook.tab(notebook.select(), "text")
    if tab_text == "Analytics":
        refresh_analytics()
    elif tab_text == "Reports":
        build_report(show_message=False)
    elif tab_text == "Network Jobs":
        load_network_settings_ui()
    elif tab_text == "Device Results":
        load_network_settings_ui()
        refresh_device_results()
    elif tab_text == "Manual Tests":
        load_network_settings_ui()
        refresh_manual_tests()
    elif tab_text == "Users":
        load_users()
    elif tab_text == "Settings":
        load_settings_ui()
    elif tab_text == "Logs":
        load_logs()


notebook.bind("<<NotebookTabChanged>>", on_tab_changed)


# =========================
# ACTION BUTTONS
# =========================
tk.Button(
    action_bar,
    text="New / Clear",
    command=clear_form,
    bg="#475569",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=14,
    pady=7,
).pack(side="left", padx=6, pady=9)
tk.Button(
    action_bar,
    text="Save / Update",
    command=submit,
    bg="#16a34a",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=14,
    pady=7,
).pack(side="left", padx=6, pady=9)
tk.Button(
    action_bar,
    text="Delete",
    command=delete,
    bg="#dc2626",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=14,
    pady=7,
).pack(side="left", padx=6, pady=9)
tk.Button(
    action_bar,
    text="Export Excel",
    command=export_excel,
    bg="#2563eb",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=14,
    pady=7,
).pack(side="left", padx=6, pady=9)
tk.Button(
    action_bar,
    text="Export JSON",
    command=export_json,
    bg="#7c3aed",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=14,
    pady=7,
).pack(side="left", padx=6, pady=9)
tk.Button(
    action_bar,
    text="Apply Keys (F1)",
    command=apply_keyboard_actions,
    bg="#f59e0b",
    fg="black",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=14,
    pady=7,
).pack(side="left", padx=6, pady=9)

ttk.Label(action_bar, text=f"Version {APP_VERSION}", style="Muted.TLabel").pack(side="right", padx=5)


# =========================
# START
# =========================
init_db()
load_settings_ui()
load_network_settings_ui()
root.title(get_setting("system_name", APP_TITLE))
title_label.config(text=get_setting("system_name", APP_TITLE))
load_data()
load_users()
load_logs()
refresh_analytics()
set_protected_tabs("hidden")

root.mainloop()
