import os
import importlib.util
import sqlite3
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
PYTHON_EXE = APP_DIR / "python" / "python.exe"
LOG_FILE = Path(os.environ.get("TEMP", str(APP_DIR))) / "winpe_app_error.log"


def run_command(command, title="Command"):
    try:
        with LOG_FILE.open("a", encoding="utf-8", errors="replace") as log:
            log.write(f"\n\n[{title}]\n{command}\n")
            if isinstance(command, (list, tuple)):
                subprocess.Popen(command, cwd=str(APP_DIR), stdout=log, stderr=log)
            else:
                subprocess.Popen(["cmd.exe", "/c", str(command)], cwd=str(APP_DIR), stdout=log, stderr=log)
    except Exception as exc:
        messagebox.showerror(title, f"{exc}\n\nLog: {LOG_FILE}")


def run_python(script):
    python = PYTHON_EXE if PYTHON_EXE.exists() else Path(sys.executable)
    run_command([str(python), str(script)], script.name)


def run_db_summary(db_path, output):
    output.delete("1.0", tk.END)
    try:
        conn = sqlite3.connect(str(db_path))
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = [row[0] for row in cur.fetchall()]
        output.insert(tk.END, f"Database: {db_path.name}\n\n")
        if not tables:
            output.insert(tk.END, "No tables found.\n")
        for table in tables:
            try:
                cur.execute(f'SELECT COUNT(*) FROM "{table}"')
                count = cur.fetchone()[0]
                output.insert(tk.END, f"{table}: {count} rows\n")
            except Exception as exc:
                output.insert(tk.END, f"{table}: {exc}\n")
        conn.close()
    except Exception as exc:
        output.insert(tk.END, str(exc))


def diagnose(output):
    output.delete("1.0", tk.END)
    checks = [
        ("Python", sys.executable),
        ("App dir", str(APP_DIR)),
        ("Tkinter", "OK"),
        ("SQLite", sqlite3.sqlite_version),
        ("Server URL", os.environ.get("PC_AUDIT_SERVER_URL", "")),
        ("Client name", os.environ.get("PC_AUDIT_CLIENT_NAME", "")),
    ]
    for label, value in checks:
        output.insert(tk.END, f"{label}: {value}\n")


def load_hardware(output):
    output.delete("1.0", tk.END)
    register_client = APP_DIR / "register_client.py"
    if not register_client.exists():
        output.insert(tk.END, f"register_client.py not found in {APP_DIR}\n")
        return
    try:
        spec = importlib.util.spec_from_file_location("register_client", str(register_client))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        inventory = module.collect_inventory()
        if not inventory:
            output.insert(tk.END, "No hardware information was returned.\n")
            output.insert(tk.END, "Try rebuilding the WIM with the latest register_client.py.\n")
            return
        for key, value in inventory.items():
            output.insert(tk.END, f"{key}: {value}\n")
    except Exception as exc:
        output.insert(tk.END, f"Hardware inventory failed: {exc}\n")
        output.insert(tk.END, f"Log: {LOG_FILE}\n")


def make_button(parent, text, command):
    return ttk.Button(parent, text=text, command=command)


def build_launcher():
    root = tk.Tk()
    root.title("PC Audit WinPE Launcher")
    root.geometry("820x560")
    root.minsize(760, 500)

    style = ttk.Style(root)
    style.configure("TButton", padding=(10, 8))
    style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))

    header = ttk.Frame(root, padding=(16, 12))
    header.pack(fill=tk.X)
    ttk.Label(header, text="PC Audit WinPE Launcher", style="Title.TLabel").pack(side=tk.LEFT)
    ttk.Label(header, text=str(APP_DIR)).pack(side=tk.RIGHT)

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))

    main_tab = ttk.Frame(notebook, padding=14)
    python_tab = ttk.Frame(notebook, padding=14)
    tools_tab = ttk.Frame(notebook, padding=14)
    hardware_tab = ttk.Frame(notebook, padding=14)
    sql_tab = ttk.Frame(notebook, padding=14)
    diag_tab = ttk.Frame(notebook, padding=14)

    notebook.add(main_tab, text="Main")
    notebook.add(python_tab, text="Python")
    notebook.add(tools_tab, text="Executables")
    notebook.add(hardware_tab, text="Hardware")
    notebook.add(sql_tab, text="SQL")
    notebook.add(diag_tab, text="Diagnose")

    main_buttons = ttk.Frame(main_tab)
    main_buttons.pack(fill=tk.X)

    make_button(main_buttons, "Run PC Audit Tool", lambda: run_python(APP_DIR / "pc_audit_report.py")).grid(row=0, column=0, sticky="ew", padx=6, pady=6)
    make_button(main_buttons, "Run PC Audit App", lambda: run_python(APP_DIR / "pc_audit_app.pyw")).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
    make_button(
        main_buttons,
        "Register Client Once",
        lambda: run_command(
            [
                str(PYTHON_EXE),
                str(APP_DIR / "register_client.py"),
                "--server",
                os.environ.get("PC_AUDIT_SERVER_URL", "http://192.168.1.3:8765"),
                "--name",
                os.environ.get("PC_AUDIT_CLIENT_NAME", "%COMPUTERNAME%-WinPE"),
                "--once",
            ],
            "Register Client",
        ),
    ).grid(row=1, column=0, sticky="ew", padx=6, pady=6)
    make_button(main_buttons, "Capture WIM", lambda: run_command(str(APP_DIR / "capture_wim.cmd"), "Capture WIM")).grid(row=1, column=1, sticky="ew", padx=6, pady=6)
    make_button(main_buttons, "Apply WIM", lambda: run_command(str(APP_DIR / "apply_wim.cmd"), "Apply WIM")).grid(row=2, column=0, sticky="ew", padx=6, pady=6)
    make_button(main_buttons, "Command Prompt", lambda: run_command("cmd", "Command Prompt")).grid(row=2, column=1, sticky="ew", padx=6, pady=6)
    main_buttons.columnconfigure(0, weight=1)
    main_buttons.columnconfigure(1, weight=1)

    for script in sorted(APP_DIR.glob("*.py")) + sorted(APP_DIR.glob("*.pyw")):
        make_button(python_tab, script.name, lambda p=script: run_python(p)).pack(fill=tk.X, pady=4)

    for tool in sorted(list(APP_DIR.glob("*.exe")) + list(APP_DIR.glob("*.bat")) + list(APP_DIR.glob("*.cmd"))):
        if tool.name.lower() == "start_pc_audit_winpe.cmd":
            continue
        make_button(tools_tab, tool.name, lambda p=tool: run_command(str(p), p.name)).pack(fill=tk.X, pady=4)

    hardware_output = tk.Text(hardware_tab, height=20, wrap=tk.WORD)
    hardware_output.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
    make_button(hardware_tab, "Read This Computer", lambda: load_hardware(hardware_output)).pack(fill=tk.X)

    sql_output = tk.Text(sql_tab, height=18, wrap=tk.WORD)
    sql_output.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
    for db in sorted(list(APP_DIR.glob("*.db")) + list(APP_DIR.glob("*.sqlite")) + list(APP_DIR.glob("*.sqlite3"))):
        make_button(sql_tab, f"Inspect {db.name}", lambda p=db: run_db_summary(p, sql_output)).pack(fill=tk.X, pady=4)

    diag_output = tk.Text(diag_tab, height=20, wrap=tk.WORD)
    diag_output.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
    make_button(diag_tab, "Run Diagnosis", lambda: diagnose(diag_output)).pack(fill=tk.X)
    diagnose(diag_output)

    root.mainloop()


if __name__ == "__main__":
    build_launcher()
