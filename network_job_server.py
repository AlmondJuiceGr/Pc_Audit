import argparse
import json
import os
import sqlite3
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.environ.get("JOB_SERVER_DB", os.path.join(BASE_DIR, "network_jobs.db"))
HOST = os.environ.get("JOB_SERVER_HOST", "0.0.0.0")
PORT = int(os.environ.get("JOB_SERVER_PORT", "8765"))

VALID_STATUSES = {"Queued", "Loaded", "Completed", "Cancelled", "Failed"}


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_conn():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        """CREATE TABLE IF NOT EXISTS jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            assignee TEXT NOT NULL DEFAULT 'ALL',
            status TEXT NOT NULL DEFAULT 'Queued',
            created_by TEXT,
            created_at TEXT,
            updated_at TEXT,
            completed_at TEXT,
            payload TEXT NOT NULL,
            notes TEXT
        )"""
    )
    c.execute(
        """CREATE TABLE IF NOT EXISTS job_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id INTEGER,
            timestamp TEXT,
            actor TEXT,
            action TEXT,
            details TEXT
        )"""
    )
    c.execute(
        """CREATE TABLE IF NOT EXISTS clients (
            name TEXT PRIMARY KEY,
            username TEXT,
            role TEXT,
            app_version TEXT,
            platform TEXT,
            inventory TEXT,
            last_seen TEXT
        )"""
    )
    c.execute("PRAGMA table_info(clients)")
    client_columns = [row[1] for row in c.fetchall()]
    if "inventory" not in client_columns:
        c.execute("ALTER TABLE clients ADD COLUMN inventory TEXT")
    c.execute(
        """CREATE TABLE IF NOT EXISTS server_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )"""
    )
    c.execute(
        "INSERT OR IGNORE INTO server_settings (key, value) VALUES (?, ?)",
        ("manual_tests_required", "0"),
    )
    conn.commit()
    conn.close()


def row_to_job(row):
    job = dict(row)
    try:
        job["payload"] = json.loads(job["payload"])
    except json.JSONDecodeError:
        job["payload"] = {}
    return job


def send_json(handler, status, data):
    body = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def read_json(handler):
    length = int(handler.headers.get("Content-Length", "0"))
    if length == 0:
        return {}
    raw = handler.rfile.read(length)
    return json.loads(raw.decode("utf-8"))


def add_event(job_id, actor, action, details=""):
    conn = get_conn()
    c = conn.cursor()
    c.execute(
        "INSERT INTO job_events (job_id, timestamp, actor, action, details) VALUES (?, ?, ?, ?, ?)",
        (job_id, now(), actor or "unknown", action, details),
    )
    conn.commit()
    conn.close()


class JobHandler(BaseHTTPRequestHandler):
    server_version = "NetworkJobServer/1.0"

    def log_message(self, fmt, *args):
        print(f"{self.address_string()} - {fmt % args}")

    def do_GET(self):
        parsed = urlparse(self.path)

        if parsed.path == "/health":
            send_json(self, 200, {"ok": True, "service": "network-job-server", "port": PORT})
            return

        if parsed.path == "/jobs":
            self.handle_list_jobs(parsed)
            return

        if parsed.path == "/clients":
            self.handle_list_clients()
            return

        if parsed.path == "/settings":
            self.handle_get_settings()
            return

        if parsed.path.startswith("/jobs/"):
            self.handle_get_job(parsed)
            return

        send_json(self, 404, {"ok": False, "error": "Not found"})

    def do_POST(self):
        parsed = urlparse(self.path)

        if parsed.path == "/jobs":
            self.handle_create_job()
            return

        if parsed.path == "/clients/register":
            self.handle_register_client()
            return

        if parsed.path == "/settings":
            self.handle_update_settings()
            return

        if parsed.path.startswith("/jobs/") and parsed.path.endswith("/status"):
            self.handle_update_status(parsed)
            return

        send_json(self, 404, {"ok": False, "error": "Not found"})

    def handle_list_jobs(self, parsed):
        params = parse_qs(parsed.query)
        assignee = params.get("assignee", [""])[0].strip()
        status = params.get("status", ["Queued"])[0].strip()

        query = "SELECT * FROM jobs"
        clauses = []
        values = []

        if status and status.lower() != "all":
            clauses.append("status = ?")
            values.append(status)

        if assignee:
            clauses.append("(assignee = ? OR assignee = 'ALL' OR assignee = '')")
            values.append(assignee)

        if clauses:
            query += " WHERE " + " AND ".join(clauses)

        query += " ORDER BY id DESC"

        conn = get_conn()
        c = conn.cursor()
        c.execute(query, values)
        jobs = [row_to_job(row) for row in c.fetchall()]
        conn.close()

        send_json(self, 200, {"ok": True, "jobs": jobs})

    def handle_get_job(self, parsed):
        parts = parsed.path.strip("/").split("/")
        if len(parts) != 2 or not parts[1].isdigit():
            send_json(self, 400, {"ok": False, "error": "Invalid job id"})
            return

        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT * FROM jobs WHERE id=?", (int(parts[1]),))
        row = c.fetchone()
        conn.close()

        if not row:
            send_json(self, 404, {"ok": False, "error": "Job not found"})
            return

        send_json(self, 200, {"ok": True, "job": row_to_job(row)})

    def handle_list_clients(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            "SELECT name, username, role, app_version, platform, inventory, last_seen FROM clients ORDER BY last_seen DESC"
        )
        clients = []
        for row in c.fetchall():
            client = dict(row)
            try:
                client["inventory"] = json.loads(client.get("inventory") or "{}")
            except json.JSONDecodeError:
                client["inventory"] = {}
            clients.append(client)
        conn.close()

        send_json(self, 200, {"ok": True, "clients": clients})

    def handle_get_settings(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT key, value FROM server_settings")
        settings = {row["key"]: row["value"] for row in c.fetchall()}
        conn.close()

        send_json(
            self,
            200,
            {
                "ok": True,
                "settings": {
                    "manual_tests_required": settings.get("manual_tests_required", "0") == "1"
                },
            },
        )

    def handle_update_settings(self):
        try:
            data = read_json(self)
        except json.JSONDecodeError:
            send_json(self, 400, {"ok": False, "error": "Invalid JSON"})
            return

        manual_required = bool(data.get("manual_tests_required", False))

        conn = get_conn()
        c = conn.cursor()
        c.execute(
            "INSERT OR REPLACE INTO server_settings (key, value) VALUES (?, ?)",
            ("manual_tests_required", "1" if manual_required else "0"),
        )
        conn.commit()
        conn.close()

        send_json(self, 200, {"ok": True, "settings": {"manual_tests_required": manual_required}})

    def handle_register_client(self):
        try:
            data = read_json(self)
        except json.JSONDecodeError:
            send_json(self, 400, {"ok": False, "error": "Invalid JSON"})
            return

        name = str(data.get("name", "")).strip()
        username = str(data.get("username", "")).strip()
        role = str(data.get("role", "")).strip()
        app_version = str(data.get("app_version", "")).strip()
        client_platform = str(data.get("platform", "")).strip()
        inventory = data.get("inventory", {})

        if not name:
            send_json(self, 400, {"ok": False, "error": "Client name is required"})
            return

        if not isinstance(inventory, dict):
            inventory = {}

        timestamp = now()
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            """INSERT OR REPLACE INTO clients (name, username, role, app_version, platform, inventory, last_seen)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (name, username, role, app_version, client_platform, json.dumps(inventory, ensure_ascii=False), timestamp),
        )
        conn.commit()
        conn.close()

        send_json(self, 200, {"ok": True, "client": name, "last_seen": timestamp})

    def handle_create_job(self):
        try:
            data = read_json(self)
        except json.JSONDecodeError:
            send_json(self, 400, {"ok": False, "error": "Invalid JSON"})
            return

        title = str(data.get("title", "")).strip()
        assignee = str(data.get("assignee", "ALL")).strip() or "ALL"
        created_by = str(data.get("created_by", "")).strip()
        payload = data.get("payload")
        notes = str(data.get("notes", "")).strip()

        if not title:
            send_json(self, 400, {"ok": False, "error": "Title is required"})
            return

        if not isinstance(payload, dict):
            send_json(self, 400, {"ok": False, "error": "Payload must be an object"})
            return

        timestamp = now()
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            """INSERT INTO jobs (title, assignee, status, created_by, created_at, updated_at, payload, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (title, assignee, "Queued", created_by, timestamp, timestamp, json.dumps(payload, ensure_ascii=False), notes),
        )
        job_id = c.lastrowid
        conn.commit()
        conn.close()

        add_event(job_id, created_by, "CREATE", f"Created job for {assignee}")
        send_json(self, 201, {"ok": True, "job_id": job_id})

    def handle_update_status(self, parsed):
        parts = parsed.path.strip("/").split("/")
        if len(parts) != 3 or not parts[1].isdigit():
            send_json(self, 400, {"ok": False, "error": "Invalid job id"})
            return

        try:
            data = read_json(self)
        except json.JSONDecodeError:
            send_json(self, 400, {"ok": False, "error": "Invalid JSON"})
            return

        job_id = int(parts[1])
        status = str(data.get("status", "")).strip()
        actor = str(data.get("actor", "")).strip()
        notes = str(data.get("notes", "")).strip()

        if status not in VALID_STATUSES:
            send_json(self, 400, {"ok": False, "error": "Invalid status"})
            return

        completed_at = now() if status in {"Completed", "Cancelled", "Failed"} else None

        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT id FROM jobs WHERE id=?", (job_id,))
        if not c.fetchone():
            conn.close()
            send_json(self, 404, {"ok": False, "error": "Job not found"})
            return

        c.execute(
            "UPDATE jobs SET status=?, updated_at=?, completed_at=COALESCE(?, completed_at), notes=? WHERE id=?",
            (status, now(), completed_at, notes, job_id),
        )
        conn.commit()
        conn.close()

        add_event(job_id, actor, status.upper(), notes)
        send_json(self, 200, {"ok": True, "job_id": job_id, "status": status})


def parse_args():
    parser = argparse.ArgumentParser(description="Enterprise Desktop System network job server.")
    parser.add_argument("--host", default=HOST, help="Bind address. Use 0.0.0.0 for LAN clients.")
    parser.add_argument("--port", type=int, default=PORT, help="Server port.")
    parser.add_argument("--db", default=DB_NAME, help="SQLite database path.")
    return parser.parse_args()


def main():
    global DB_NAME, HOST, PORT

    args = parse_args()

    DB_NAME = args.db
    HOST = args.host
    PORT = args.port

    init_db()
    server = ThreadingHTTPServer((HOST, PORT), JobHandler)
    print(f"Network job server running on http://{HOST}:{PORT}")
    print("Authentication is disabled. Restrict access with your LAN/firewall if needed.")
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
