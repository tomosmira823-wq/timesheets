"""
Timesheets - timesheet management, approvals and job costing.
Desktop app for Windows (pywebview + SQLite). Single file.

Run from source:  python timesheet_app.py   (needs: pip install pywebview openpyxl fpdf2)
First run shows a setup screen: company name, your admin account, and where
the data should live (this PC, or a shared folder so the whole team uses it).
"""

import os
import sys
import csv
import re
import json
import sqlite3
import hashlib
import secrets
import datetime as dt
import time
import tempfile
import shutil
import glob
from contextlib import contextmanager
import threading

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------

APP_NAME = "Timesheets"
APP_VERSION = "1.0.0"

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

# Everything the app writes (settings, database, receipts, backups, logs) lives
# in the user's roaming AppData, never next to the installed program. The
# database can be pointed elsewhere (e.g. a shared network folder) in Settings.
DATA_DIR = os.path.join(os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA")
                        or tempfile.gettempdir(), APP_NAME)
try:
    os.makedirs(DATA_DIR, exist_ok=True)
except Exception:
    pass
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")

DEFAULT_CONFIG = {
    "db_path": os.path.join(DATA_DIR, "timesheets.db"),
    # Minutes of no mouse/keyboard before signing out. NOTE: a value in
    # config.json OVERRIDES this - check there if the timeout seems wrong.
    "inactivity_timeout_minutes": 15,
    "daily_hours_warning": 10.0,
    "company_name": "",
    # Used for the "x of y hours logged" progress on the dashboard.
    "weekly_target_hours": 37.5,
    # Open filling the screen. Set false in config.json on any PC that would
    # rather have a normal resizable window.
    "start_maximized": True,
}

DEFAULT_JOB_ROLES = [
    "Staff", "Senior Staff", "Manager", "Director",
]

# Job references seeded on a brand-new database. Empty: companies add their own.
DEFAULT_JOBS = []

# Non-project / leave categories people can book time against (seeded once).
DEFAULT_CATEGORIES = [
    ("HOLIDAY", "Holiday"),
    ("SICK", "Sick Leave"),
    ("TRAINING", "Training"),
    ("OVERHEAD", "Office / Overhead"),
]

USER_TYPES = ["Standard", "Manager", "Admin"]

# Targets shown on the Analytics page (admins can change them in the app).
DEFAULT_GOALS = {
    "utilisation_pct": 80,        # billable share of logged hours
    "monthly_invoiced": 100000,   # £ invoiced per month
    "submission_pct": 95,         # timesheets submitted on time
    "overtime_cap": 60,           # overtime hours per month, keep below
}
TS_STATUSES = ["Draft", "Submitted", "Approved", "Rejected", "Locked"]

# Overtime rates and their pay multipliers.
RATE_OPTIONS = ["Normal", "x1.5", "x2"]
RATE_MULT = {"Normal": 1.0, "x1.5": 1.5, "x2": 2.0}


def paid_hours(hours, rate):
    """Convert worked hours to paid hours using the overtime multiplier."""
    return round(float(hours) * RATE_MULT.get(rate, 1.0), 2)


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
                cfg.update(json.load(fh))
        except Exception:
            pass
    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2)


# ----------------------------------------------------------------------------
# Password hashing (PBKDF2 - standard library only)
# ----------------------------------------------------------------------------

_PBKDF2_ITERATIONS = 200_000


def hash_password(password, salt=None):
    if salt is None:
        salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), _PBKDF2_ITERATIONS
    )
    return salt, digest.hex()


def verify_password(password, salt, expected_hash):
    _, computed = hash_password(password, salt)
    return secrets.compare_digest(computed, expected_hash)


# ----------------------------------------------------------------------------
# Date helpers
# ----------------------------------------------------------------------------

def monday_of(date):
    return date - dt.timedelta(days=date.weekday())


def week_dates(week_start):
    return [week_start + dt.timedelta(days=i) for i in range(7)]


def iso(d):
    return d.isoformat()


# ----------------------------------------------------------------------------
# Database
# ----------------------------------------------------------------------------

class DB:
    def __init__(self, path):
        self.path = path
        folder = os.path.dirname(path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)
        self.init_schema()

    @contextmanager
    def conn(self):
        # Short-lived connection per operation. On a shared network drive this
        # is the safe pattern: file locks are released the instant we
        # commit+close, so multiple PCs never block each other. (A single
        # long-lived connection caused lock contention -> 'Not responding'.)
        con = sqlite3.connect(self.path, timeout=30)
        con.row_factory = sqlite3.Row
        try:
            con.execute("PRAGMA busy_timeout=15000;")
            con.execute("PRAGMA foreign_keys=ON;")
            yield con
            con.commit()
        finally:
            con.close()

    def init_schema(self):
        with self.conn() as con:
            con.executescript(
                """
                CREATE TABLE IF NOT EXISTS app_meta (
                    key   TEXT PRIMARY KEY,
                    value TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS users (
                    id            INTEGER PRIMARY KEY AUTOINCREMENT,
                    full_name     TEXT NOT NULL,
                    email         TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    password_hash TEXT NOT NULL,
                    salt          TEXT NOT NULL,
                    job_role      TEXT NOT NULL,
                    user_type     TEXT NOT NULL,
                    active        INTEGER NOT NULL DEFAULT 1
                );
                CREATE TABLE IF NOT EXISTS job_roles (
                    id      INTEGER PRIMARY KEY AUTOINCREMENT,
                    name    TEXT NOT NULL UNIQUE,
                    enabled INTEGER NOT NULL DEFAULT 1
                );
                CREATE TABLE IF NOT EXISTS job_refs (
                    id        INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_ref   TEXT NOT NULL UNIQUE,
                    job_title TEXT NOT NULL DEFAULT '',
                    client    TEXT NOT NULL DEFAULT '',
                    status    TEXT NOT NULL DEFAULT 'Active',
                    notes     TEXT NOT NULL DEFAULT '',
                    kind      TEXT NOT NULL DEFAULT 'Project',
                    quote_value REAL NOT NULL DEFAULT 0,
                    quote_attachment TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS job_rates (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id    INTEGER NOT NULL REFERENCES users(id),
                    job_ref_id INTEGER NOT NULL REFERENCES job_refs(id),
                    rate       REAL NOT NULL DEFAULT 0,
                    UNIQUE(user_id, job_ref_id)
                );
                CREATE TABLE IF NOT EXISTS timesheets (
                    id                INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id           INTEGER NOT NULL REFERENCES users(id),
                    week_start        TEXT NOT NULL,
                    status            TEXT NOT NULL DEFAULT 'Draft',
                    reject_reason     TEXT NOT NULL DEFAULT '',
                    submitted_snapshot TEXT NOT NULL DEFAULT '',
                    updated_at        TEXT NOT NULL,
                    UNIQUE(user_id, week_start)
                );
                CREATE TABLE IF NOT EXISTS timesheet_lines (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    timesheet_id INTEGER NOT NULL REFERENCES timesheets(id) ON DELETE CASCADE,
                    work_date    TEXT NOT NULL,
                    job_ref_id   INTEGER NOT NULL REFERENCES job_refs(id),
                    narrative    TEXT NOT NULL DEFAULT '',
                    hours        REAL NOT NULL DEFAULT 0,
                    job_role     TEXT NOT NULL DEFAULT '',
                    rate         TEXT NOT NULL DEFAULT 'Normal'
                );
                CREATE TABLE IF NOT EXISTS expenses (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id     INTEGER NOT NULL REFERENCES users(id),
                    exp_date    TEXT NOT NULL,
                    job_ref_id  INTEGER REFERENCES job_refs(id),
                    description TEXT NOT NULL DEFAULT '',
                    value       REAL NOT NULL DEFAULT 0,
                    receipt     TEXT NOT NULL DEFAULT '',
                    status      TEXT NOT NULL DEFAULT 'Logged',
                    payment_receipt TEXT NOT NULL DEFAULT '',
                    created_at  TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS audit (
                    id        INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_email TEXT NOT NULL,
                    action    TEXT NOT NULL,
                    timestamp TEXT NOT NULL,
                    old_value TEXT NOT NULL DEFAULT '',
                    new_value TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS job_variations (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_ref_id  INTEGER NOT NULL REFERENCES job_refs(id),
                    var_ref     TEXT NOT NULL DEFAULT '',
                    description TEXT NOT NULL DEFAULT '',
                    value       REAL NOT NULL DEFAULT 0,
                    var_date    TEXT NOT NULL DEFAULT '',
                    attachment  TEXT NOT NULL DEFAULT '',
                    created_by  TEXT NOT NULL DEFAULT '',
                    created_at  TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS job_invoices (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_ref_id  INTEGER NOT NULL REFERENCES job_refs(id),
                    invoice_no  TEXT NOT NULL DEFAULT '',
                    inv_date    TEXT NOT NULL DEFAULT '',
                    description TEXT NOT NULL DEFAULT '',
                    value       REAL NOT NULL DEFAULT 0,
                    status      TEXT NOT NULL DEFAULT 'Raised',
                    paid_date   TEXT NOT NULL DEFAULT '',
                    attachment  TEXT NOT NULL DEFAULT '',
                    created_by  TEXT NOT NULL DEFAULT '',
                    created_at  TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS job_bills (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_ref_id  INTEGER NOT NULL REFERENCES job_refs(id),
                    supplier    TEXT NOT NULL DEFAULT '',
                    bill_no     TEXT NOT NULL DEFAULT '',
                    bill_date   TEXT NOT NULL DEFAULT '',
                    description TEXT NOT NULL DEFAULT '',
                    value       REAL NOT NULL DEFAULT 0,
                    category    TEXT NOT NULL DEFAULT '',
                    status      TEXT NOT NULL DEFAULT 'Received',
                    paid_date   TEXT NOT NULL DEFAULT '',
                    attachment  TEXT NOT NULL DEFAULT '',
                    created_by  TEXT NOT NULL DEFAULT '',
                    created_at  TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS job_documents (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_ref_id  INTEGER NOT NULL REFERENCES job_refs(id),
                    title       TEXT NOT NULL DEFAULT '',
                    description TEXT NOT NULL DEFAULT '',
                    doc_date    TEXT NOT NULL DEFAULT '',
                    attachment  TEXT NOT NULL DEFAULT '',
                    created_by  TEXT NOT NULL DEFAULT '',
                    created_at  TEXT NOT NULL DEFAULT ''
                );
                CREATE INDEX IF NOT EXISTS ix_jobdoc_job ON job_documents(job_ref_id);
                CREATE INDEX IF NOT EXISTS ix_jobvar_job ON job_variations(job_ref_id);
                CREATE INDEX IF NOT EXISTS ix_jobinv_job ON job_invoices(job_ref_id);
                CREATE INDEX IF NOT EXISTS ix_jobbill_job ON job_bills(job_ref_id);
                """
            )
            # add-only migrations so existing databases gain new columns without
            # losing any data. Each tuple is (table, column, column definition).
            for table, col, ddl in [
                ("job_refs", "kind", "TEXT NOT NULL DEFAULT 'Project'"),
                ("timesheets", "reject_reason", "TEXT NOT NULL DEFAULT ''"),
                ("timesheets", "submitted_snapshot", "TEXT NOT NULL DEFAULT ''"),
                ("timesheet_lines", "job_role", "TEXT NOT NULL DEFAULT ''"),
                ("timesheet_lines", "rate", "TEXT NOT NULL DEFAULT 'Normal'"),
                ("expenses", "payment_receipt", "TEXT NOT NULL DEFAULT ''"),
                ("job_refs", "quote_value", "REAL NOT NULL DEFAULT 0"),
                ("users", "default_rate", "REAL NOT NULL DEFAULT 0"),
                ("users", "timer_start", "TEXT NOT NULL DEFAULT ''"),
                ("users", "auth_token", "TEXT NOT NULL DEFAULT ''"),
                ("job_invoices", "attachment", "TEXT NOT NULL DEFAULT ''"),
                ("job_variations", "attachment", "TEXT NOT NULL DEFAULT ''"),
                ("job_refs", "quote_attachment", "TEXT NOT NULL DEFAULT ''"),
                ("job_bills", "attachment", "TEXT NOT NULL DEFAULT ''"),
                ("job_bills", "category", "TEXT NOT NULL DEFAULT ''"),
            ]:
                have = [r["name"] for r in con.execute(
                    f"PRAGMA table_info({table})").fetchall()]
                if col not in have:
                    con.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")
            # One-time self-repair: a Draft must never carry a submitted
            # snapshot. The old "Unlock" button wrongly reset Locked/Approved
            # timesheets to Draft (it now returns them to Approved), which made
            # already-approved timesheets disappear from the Approvals list.
            # Restore any timesheet stuck in that state back to Approved.
            con.execute("UPDATE timesheets SET status='Approved' "
                        "WHERE status='Draft' AND submitted_snapshot != ''")

    def seed_defaults(self):
        with self.conn() as con:
            for r in DEFAULT_JOB_ROLES:
                con.execute("INSERT OR IGNORE INTO job_roles(name, enabled) VALUES(?,1)", (r,))
            # Seed starter jobs/categories ONLY on a brand-new database. Doing
            # it every launch made a renamed default (e.g. HOLIDAY -> 26-ATX-001)
            # reappear as a duplicate on the next start.
            if con.execute("SELECT COUNT(*) c FROM job_refs").fetchone()["c"] == 0:
                for jr, jt, cl in DEFAULT_JOBS:
                    con.execute(
                        "INSERT OR IGNORE INTO job_refs(job_ref, job_title, client, status, notes)"
                        " VALUES(?,?,?, 'Active', '')", (jr, jt, cl))
                for jr, jt in DEFAULT_CATEGORIES:
                    con.execute(
                        "INSERT OR IGNORE INTO job_refs(job_ref, job_title, client, status, notes, kind)"
                        " VALUES(?,?, '', 'Active', '', 'Non-project')", (jr, jt))
            # No default admin account: the first-run setup screen creates one.

            # Back-fill frozen submission snapshots for timesheets submitted
            # before this feature existed, so employees keep seeing their own
            # version once managers start editing the live copies. Runs once
            # per timesheet (only fills empty snapshots).
            todo = con.execute(
                "SELECT id FROM timesheets WHERE submitted_snapshot=''"
                " AND status IN ('Submitted','Approved','Locked')").fetchall()
            for tr in todo:
                tid = tr["id"]
                saved = con.execute(
                    "SELECT l.*, j.job_ref, j.job_title FROM timesheet_lines l"
                    " JOIN job_refs j ON j.id=l.job_ref_id WHERE l.timesheet_id=?"
                    " ORDER BY l.work_date, l.id", (tid,)).fetchall()
                snap = [{"work_date": s["work_date"], "job_ref_id": s["job_ref_id"],
                         "job_ref": s["job_ref"], "job_title": s["job_title"],
                         "job_role": (s["job_role"] if "job_role" in s.keys() else ""),
                         "rate": (s["rate"] if "rate" in s.keys() else "Normal"),
                         "narrative": s["narrative"], "hours": s["hours"]}
                        for s in saved]
                con.execute("UPDATE timesheets SET submitted_snapshot=? WHERE id=?",
                            (json.dumps(snap), tid))

    def get_meta(self, key, default=""):
        with self.conn() as con:
            r = con.execute("SELECT value FROM app_meta WHERE key=?", (key,)).fetchone()
            return r["value"] if r else default

    def set_meta(self, key, value):
        with self.conn() as con:
            con.execute("INSERT OR REPLACE INTO app_meta(key, value) VALUES(?,?)", (key, str(value)))

    def needs_setup(self):
        """True until the first (admin) account exists."""
        with self.conn() as con:
            return con.execute("SELECT COUNT(*) c FROM users").fetchone()["c"] == 0

    def create_first_admin(self, full_name, email, password):
        with self.conn() as con:
            if con.execute("SELECT COUNT(*) c FROM users").fetchone()["c"]:
                raise ValueError("An account already exists.")
            salt, ph = hash_password(password)
            con.execute(
                """INSERT INTO users(full_name, email, password_hash, salt,
                                     job_role, user_type, active)
                   VALUES(?,?,?,?,?,?,1)""",
                (full_name, email, ph, salt, "Director", "Admin"))
            self._audit(con, email, "SETUP", "", "First admin account created")

    def _audit(self, con, user_email, action, old="", new=""):
        con.execute(
            "INSERT INTO audit(user_email, action, timestamp, old_value, new_value)"
            " VALUES(?,?,?,?,?)",
            (user_email, action, dt.datetime.now().isoformat(timespec="seconds"),
             str(old), str(new)))

    def audit(self, user_email, action, old="", new=""):
        with self.conn() as con:
            self._audit(con, user_email, action, old, new)

    def get_audit(self, limit=500):
        with self.conn() as con:
            return [dict(r) for r in con.execute(
                "SELECT * FROM audit ORDER BY id DESC LIMIT ?", (limit,)).fetchall()]

    def authenticate(self, email, password):
        ident = (email or "").strip()
        with self.conn() as con:
            row = con.execute("SELECT * FROM users WHERE email=? COLLATE NOCASE",
                              (ident,)).fetchone()
            if not row and ident and "@" not in ident:
                # allow signing in with just the part before the @
                matches = con.execute(
                    "SELECT * FROM users WHERE "
                    "lower(substr(email, 1, instr(email,'@')-1))=lower(?)",
                    (ident,)).fetchall()
                if len(matches) == 1:
                    row = matches[0]
            if not row or not row["active"]:
                return None
            if verify_password(password, row["salt"], row["password_hash"]):
                self._audit(con, row["email"], "LOGIN", "", "Successful login")
                return dict(row)
            return None

    def get_user_by_email(self, email):
        with self.conn() as con:
            r = con.execute("SELECT * FROM users WHERE email=? COLLATE NOCASE",
                            ((email or "").strip(),)).fetchone()
            return dict(r) if r else None

    def list_users(self, include_inactive=True):
        with self.conn() as con:
            q = "SELECT * FROM users"
            if not include_inactive:
                q += " WHERE active=1"
            q += " ORDER BY full_name"
            return [dict(r) for r in con.execute(q).fetchall()]

    def get_user(self, user_id):
        with self.conn() as con:
            r = con.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
            return dict(r) if r else None

    def create_user(self, actor_email, full_name, email, password, job_role,
                    user_type, active=True, default_rate=0.0):
        salt, ph = hash_password(password)
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO users(full_name, email, password_hash, salt,
                                     job_role, user_type, active, default_rate)
                   VALUES(?,?,?,?,?,?,?,?)""",
                (full_name.strip(), email.strip(), ph, salt, job_role,
                 user_type, 1 if active else 0, float(default_rate or 0)))
            self._audit(con, actor_email, "CREATE_USER", "", email)
            return cur.lastrowid

    def update_user(self, actor_email, user_id, full_name, email, job_role,
                    user_type, active, default_rate=0.0):
        with self.conn() as con:
            old = con.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
            con.execute(
                """UPDATE users SET full_name=?, email=?, job_role=?,
                                    user_type=?, active=?, default_rate=? WHERE id=?""",
                (full_name.strip(), email.strip(), job_role, user_type,
                 1 if active else 0, float(default_rate or 0), user_id))
            self._audit(con, actor_email, "UPDATE_USER", dict(old) if old else "", email)

    def set_password(self, actor_email, user_id, new_password):
        salt, ph = hash_password(new_password)
        with self.conn() as con:
            email = con.execute("SELECT email FROM users WHERE id=?",
                                (user_id,)).fetchone()["email"]
            con.execute("UPDATE users SET password_hash=?, salt=? WHERE id=?",
                        (ph, salt, user_id))
            self._audit(con, actor_email, "RESET_PASSWORD", "", email)

    def delete_user(self, actor_email, user_id):
        with self.conn() as con:
            email = con.execute("SELECT email FROM users WHERE id=?",
                                (user_id,)).fetchone()["email"]
            has = con.execute("SELECT COUNT(*) c FROM timesheets WHERE user_id=?",
                              (user_id,)).fetchone()["c"]
            if has:
                con.execute("UPDATE users SET active=0 WHERE id=?", (user_id,))
                self._audit(con, actor_email, "DEACTIVATE_USER", "",
                            f"{email} (had timesheets)")
                return "deactivated"
            con.execute("DELETE FROM users WHERE id=?", (user_id,))
            self._audit(con, actor_email, "DELETE_USER", email, "")
            return "deleted"

    def list_roles(self, only_enabled=False):
        with self.conn() as con:
            q = "SELECT * FROM job_roles"
            if only_enabled:
                q += " WHERE enabled=1"
            q += " ORDER BY name"
            return [dict(r) for r in con.execute(q).fetchall()]

    def add_role(self, actor_email, name):
        with self.conn() as con:
            con.execute("INSERT OR IGNORE INTO job_roles(name, enabled) VALUES(?,1)",
                        (name.strip(),))
            self._audit(con, actor_email, "ADD_ROLE", "", name)

    def set_role_enabled(self, actor_email, role_id, enabled):
        with self.conn() as con:
            con.execute("UPDATE job_roles SET enabled=? WHERE id=?",
                        (1 if enabled else 0, role_id))
            self._audit(con, actor_email, "TOGGLE_ROLE", role_id,
                        "enabled" if enabled else "disabled")

    def list_jobs(self, only_active=False, search=None):
        with self.conn() as con:
            q = "SELECT * FROM job_refs"
            clauses, params = [], []
            if only_active:
                clauses.append("status='Active'")
            if search:
                clauses.append("(job_ref LIKE ? OR job_title LIKE ? OR client LIKE ?)")
                s = f"%{search}%"
                params += [s, s, s]
            if clauses:
                q += " WHERE " + " AND ".join(clauses)
            q += " ORDER BY job_ref"
            return [dict(r) for r in con.execute(q, params).fetchall()]

    def create_job(self, actor_email, job_ref, job_title, client, status, notes,
                   kind="Project", quote_value=0.0, quote_attachment=""):
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO job_refs(job_ref, job_title, client, status, notes, kind,
                                        quote_value, quote_attachment)
                   VALUES(?,?,?,?,?,?,?,?)""",
                (job_ref.strip(), job_title.strip(), client.strip(), status,
                 notes.strip(), kind, float(quote_value), quote_attachment or ""))
            self._audit(con, actor_email, "CREATE_JOB", "", job_ref)
            return cur.lastrowid

    def update_job(self, actor_email, job_id, job_ref, job_title, client, status,
                   notes, kind="Project", quote_value=0.0, quote_attachment=None):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_refs WHERE id=?", (job_id,)).fetchone()
            if quote_attachment is None:
                quote_attachment = (old["quote_attachment"] if old else "")
            con.execute(
                """UPDATE job_refs SET job_ref=?, job_title=?, client=?, status=?,
                                       notes=?, kind=?, quote_value=?,
                                       quote_attachment=? WHERE id=?""",
                (job_ref.strip(), job_title.strip(), client.strip(), status,
                 notes.strip(), kind, float(quote_value),
                 quote_attachment or "", job_id))
            self._audit(con, actor_email, "UPDATE_JOB", dict(old) if old else "", job_ref)

    def delete_job(self, actor_email, job_id):
        with self.conn() as con:
            ref = con.execute("SELECT job_ref FROM job_refs WHERE id=?",
                              (job_id,)).fetchone()["job_ref"]
            used = con.execute("SELECT COUNT(*) c FROM timesheet_lines WHERE job_ref_id=?",
                               (job_id,)).fetchone()["c"]
            if used:
                con.execute("UPDATE job_refs SET status='Closed' WHERE id=?", (job_id,))
                self._audit(con, actor_email, "CLOSE_JOB", "", f"{ref} (had time booked)")
                return "closed (it has time booked, so it was kept for history)"
            con.execute("DELETE FROM job_refs WHERE id=?", (job_id,))
            self._audit(con, actor_email, "DELETE_JOB", ref, "")
            return "deleted"

    def set_job_rate(self, actor_email, user_id, job_ref_id, rate):
        with self.conn() as con:
            con.execute("DELETE FROM job_rates WHERE user_id=? AND job_ref_id=?",
                        (user_id, job_ref_id))
            con.execute("INSERT INTO job_rates(user_id, job_ref_id, rate) VALUES(?,?,?)",
                        (user_id, job_ref_id, float(rate)))
            self._audit(con, actor_email, "SET_RATE", f"job {job_ref_id} user {user_id}",
                        str(rate))

    def get_job_rates(self, job_ref_id):
        with self.conn() as con:
            return {r["user_id"]: r["rate"] for r in con.execute(
                "SELECT user_id, rate FROM job_rates WHERE job_ref_id=?",
                (job_ref_id,)).fetchall()}

    # ---- variations against a job's quote ----

    def list_variations(self, job_ref_id=None):
        with self.conn() as con:
            q = ("SELECT v.*, j.job_ref, j.job_title FROM job_variations v "
                 "JOIN job_refs j ON j.id=v.job_ref_id")
            params = []
            if job_ref_id:
                q += " WHERE v.job_ref_id=?"
                params.append(job_ref_id)
            q += " ORDER BY v.var_date, v.id"
            return [dict(r) for r in con.execute(q, params).fetchall()]

    def add_variation(self, actor_email, job_ref_id, var_ref, description, value,
                      var_date, attachment=""):
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO job_variations(job_ref_id, var_ref, description,
                                              value, var_date, attachment,
                                              created_by, created_at)
                   VALUES(?,?,?,?,?,?,?,?)""",
                (int(job_ref_id), (var_ref or "").strip(),
                 (description or "").strip(), float(value or 0),
                 (var_date or "").strip(), attachment or "", actor_email,
                 dt.datetime.now().isoformat(timespec="seconds")))
            self._audit(con, actor_email, "ADD_VARIATION", "",
                        f"job {job_ref_id} {var_ref} {value}")
            return cur.lastrowid

    def update_variation(self, actor_email, vid, var_ref, description, value, var_date,
                         attachment=None):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_variations WHERE id=?",
                              (vid,)).fetchone()
            if attachment is None:
                attachment = (old["attachment"] if old else "")
            con.execute(
                """UPDATE job_variations SET var_ref=?, description=?, value=?,
                                             var_date=?, attachment=? WHERE id=?""",
                ((var_ref or "").strip(), (description or "").strip(),
                 float(value or 0), (var_date or "").strip(),
                 attachment or "", vid))
            self._audit(con, actor_email, "UPDATE_VARIATION",
                        (dict(old) if old else ""), f"variation {vid} -> {value}")

    def delete_variation(self, actor_email, vid):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_variations WHERE id=?",
                              (vid,)).fetchone()
            con.execute("DELETE FROM job_variations WHERE id=?", (vid,))
            self._audit(con, actor_email, "DELETE_VARIATION",
                        (dict(old) if old else ""), "")

    def variation_totals(self):
        """Total variation value per job in ONE query (network-share friendly)."""
        with self.conn() as con:
            return {r["job_ref_id"]: round(r["v"] or 0, 2) for r in con.execute(
                "SELECT job_ref_id, SUM(value) v FROM job_variations "
                "GROUP BY job_ref_id").fetchall()}

    # ---- invoices raised against a job ----

    def list_invoices(self, job_ref_id=None):
        with self.conn() as con:
            q = ("SELECT i.*, j.job_ref, j.job_title FROM job_invoices i "
                 "JOIN job_refs j ON j.id=i.job_ref_id")
            params = []
            if job_ref_id:
                q += " WHERE i.job_ref_id=?"
                params.append(job_ref_id)
            q += " ORDER BY i.inv_date, i.id"
            return [dict(r) for r in con.execute(q, params).fetchall()]

    def invoice_exists(self, job_ref_id, invoice_no):
        """Used by the importer so re-importing a file cannot double-count."""
        inv = (invoice_no or "").strip()
        if not inv:
            return False
        with self.conn() as con:
            r = con.execute(
                "SELECT 1 FROM job_invoices WHERE job_ref_id=? AND "
                "invoice_no=? COLLATE NOCASE", (int(job_ref_id), inv)).fetchone()
            return bool(r)

    def add_invoice(self, actor_email, job_ref_id, invoice_no, inv_date, description,
                    value, status="Raised", paid_date="", attachment=""):
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO job_invoices(job_ref_id, invoice_no, inv_date,
                                            description, value, status, paid_date,
                                            attachment, created_by, created_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (int(job_ref_id), (invoice_no or "").strip(),
                 (inv_date or "").strip(), (description or "").strip(),
                 float(value or 0), status or "Raised", (paid_date or "").strip(),
                 attachment or "",
                 actor_email, dt.datetime.now().isoformat(timespec="seconds")))
            self._audit(con, actor_email, "ADD_INVOICE", "",
                        f"job {job_ref_id} inv {invoice_no} {value}")
            return cur.lastrowid

    def update_invoice(self, actor_email, iid, invoice_no, inv_date, description,
                       value, status, paid_date="", attachment=None):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_invoices WHERE id=?", (iid,)).fetchone()
            if attachment is None:
                attachment = (old["attachment"] if old else "")
            con.execute(
                """UPDATE job_invoices SET invoice_no=?, inv_date=?, description=?,
                                           value=?, status=?, paid_date=?,
                                           attachment=? WHERE id=?""",
                ((invoice_no or "").strip(), (inv_date or "").strip(),
                 (description or "").strip(), float(value or 0),
                 status or "Raised", (paid_date or "").strip(),
                 attachment or "", iid))
            self._audit(con, actor_email, "UPDATE_INVOICE",
                        (dict(old) if old else ""), f"invoice {iid} -> {value} {status}")

    def set_invoice_status(self, actor_email, iid, status, paid_date=""):
        with self.conn() as con:
            old = con.execute("SELECT status FROM job_invoices WHERE id=?",
                              (iid,)).fetchone()
            con.execute("UPDATE job_invoices SET status=?, paid_date=? WHERE id=?",
                        (status, (paid_date or "").strip(), iid))
            self._audit(con, actor_email, "INVOICE_STATUS",
                        (old["status"] if old else ""), f"invoice {iid} -> {status}")

    def delete_invoice(self, actor_email, iid):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_invoices WHERE id=?", (iid,)).fetchone()
            con.execute("DELETE FROM job_invoices WHERE id=?", (iid,))
            self._audit(con, actor_email, "DELETE_INVOICE",
                        (dict(old) if old else ""), "")

    # ---- loose documents filed against a job ----

    def list_documents(self, job_ref_id):
        with self.conn() as con:
            return [dict(r) for r in con.execute(
                "SELECT * FROM job_documents WHERE job_ref_id=? ORDER BY doc_date, id",
                (job_ref_id,)).fetchall()]

    def add_document(self, actor_email, job_ref_id, title, description, doc_date,
                     attachment):
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO job_documents(job_ref_id, title, description, doc_date,
                                             attachment, created_by, created_at)
                   VALUES(?,?,?,?,?,?,?)""",
                (int(job_ref_id), (title or "").strip(), (description or "").strip(),
                 (doc_date or "").strip(), attachment or "", actor_email,
                 dt.datetime.now().isoformat(timespec="seconds")))
            self._audit(con, actor_email, "ADD_DOCUMENT", "",
                        f"job {job_ref_id}: {title}")
            return cur.lastrowid

    def delete_document(self, actor_email, did):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_documents WHERE id=?",
                              (did,)).fetchone()
            con.execute("DELETE FROM job_documents WHERE id=?", (did,))
            self._audit(con, actor_email, "DELETE_DOCUMENT",
                        (dict(old) if old else ""), "")

    # ---- bills: supplier / subcontractor costs against a job ----

    def list_bills(self, job_ref_id=None):
        with self.conn() as con:
            q = ("SELECT b.*, j.job_ref, j.job_title FROM job_bills b "
                 "JOIN job_refs j ON j.id=b.job_ref_id")
            params = []
            if job_ref_id:
                q += " WHERE b.job_ref_id=?"
                params.append(job_ref_id)
            q += " ORDER BY b.bill_date, b.id"
            return [dict(r) for r in con.execute(q, params).fetchall()]

    def bill_exists(self, job_ref_id, supplier, bill_no):
        """Same supplier + same bill number on the same job = already entered."""
        no = (bill_no or "").strip()
        if not no:
            return False
        with self.conn() as con:
            r = con.execute(
                "SELECT 1 FROM job_bills WHERE job_ref_id=? AND bill_no=? COLLATE NOCASE"
                " AND supplier=? COLLATE NOCASE",
                (int(job_ref_id), no, (supplier or "").strip())).fetchone()
            return bool(r)

    def add_bill(self, actor_email, job_ref_id, supplier, bill_no, bill_date,
                 description, value, category="", status="Received", paid_date="",
                 attachment=""):
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO job_bills(job_ref_id, supplier, bill_no, bill_date,
                                         description, value, category, status,
                                         paid_date, attachment, created_by, created_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                (int(job_ref_id), (supplier or "").strip(), (bill_no or "").strip(),
                 (bill_date or "").strip(), (description or "").strip(),
                 float(value or 0), (category or "").strip(), status or "Received",
                 (paid_date or "").strip(), attachment or "", actor_email,
                 dt.datetime.now().isoformat(timespec="seconds")))
            self._audit(con, actor_email, "ADD_BILL", "",
                        f"job {job_ref_id} {supplier} {bill_no} {value}")
            return cur.lastrowid

    def update_bill(self, actor_email, bid, supplier, bill_no, bill_date, description,
                    value, category, status, paid_date="", attachment=None):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_bills WHERE id=?", (bid,)).fetchone()
            if attachment is None:
                attachment = (old["attachment"] if old else "")
            con.execute(
                """UPDATE job_bills SET supplier=?, bill_no=?, bill_date=?,
                                        description=?, value=?, category=?, status=?,
                                        paid_date=?, attachment=? WHERE id=?""",
                ((supplier or "").strip(), (bill_no or "").strip(),
                 (bill_date or "").strip(), (description or "").strip(),
                 float(value or 0), (category or "").strip(), status or "Received",
                 (paid_date or "").strip(), attachment or "", bid))
            self._audit(con, actor_email, "UPDATE_BILL",
                        (dict(old) if old else ""), f"bill {bid} -> {value} {status}")

    def set_bill_status(self, actor_email, bid, status, paid_date=""):
        with self.conn() as con:
            old = con.execute("SELECT status FROM job_bills WHERE id=?",
                              (bid,)).fetchone()
            con.execute("UPDATE job_bills SET status=?, paid_date=? WHERE id=?",
                        (status, (paid_date or "").strip(), bid))
            self._audit(con, actor_email, "BILL_STATUS",
                        (old["status"] if old else ""), f"bill {bid} -> {status}")

    def delete_bill(self, actor_email, bid):
        with self.conn() as con:
            old = con.execute("SELECT * FROM job_bills WHERE id=?", (bid,)).fetchone()
            con.execute("DELETE FROM job_bills WHERE id=?", (bid,))
            self._audit(con, actor_email, "DELETE_BILL",
                        (dict(old) if old else ""), "")

    def bill_totals(self):
        """Per job: (total billed to us, total we've paid) in ONE query."""
        with self.conn() as con:
            out = {}
            for r in con.execute(
                    "SELECT job_ref_id, SUM(value) tot, "
                    "SUM(CASE WHEN status='Paid' THEN value ELSE 0 END) paid "
                    "FROM job_bills GROUP BY job_ref_id").fetchall():
                out[r["job_ref_id"]] = (round(r["tot"] or 0, 2),
                                        round(r["paid"] or 0, 2))
            return out

    def invoice_totals(self):
        """Per job: (total raised, total paid) in ONE query."""
        with self.conn() as con:
            out = {}
            for r in con.execute(
                    "SELECT job_ref_id, SUM(value) tot, "
                    "SUM(CASE WHEN status='Paid' THEN value ELSE 0 END) paid "
                    "FROM job_invoices GROUP BY job_ref_id").fetchall():
                out[r["job_ref_id"]] = (round(r["tot"] or 0, 2),
                                        round(r["paid"] or 0, 2))
            return out

    def job_cost_summary(self):
        """Per job: total hours, internal cost (paid hours x each person's rate on
        that job), the quote, and the invoicing position.

        The money columns are deliberately kept separate:
          - 'cost'       = LABOUR only (paid hours x rate) - unchanged meaning,
                           so anything already reading it behaves as before
          - 'bills'      = supplier/subcontractor bills against the job
          - 'total_cost' = cost + bills, i.e. what the job has really cost us.
                           This is what budget remaining / % used measure.
          - 'invoiced'   = what we have BILLED the client (revenue)
          - 'profit'     = invoiced - total_cost
        Quote figures: 'quote' is the original as first entered and never moves;
        'quote_current' is that plus any variations, and is what budget
        remaining/% used and 'to_invoice' are measured against."""
        with self.conn() as con:
            jobs = con.execute("SELECT id, job_ref, job_title, quote_value, kind, status "
                               "FROM job_refs ORDER BY job_ref").fetchall()
            rates = {(r["user_id"], r["job_ref_id"]): r["rate"]
                     for r in con.execute("SELECT user_id, job_ref_id, rate "
                                          "FROM job_rates").fetchall()}
            defaults = {r["id"]: (r["default_rate"] or 0.0) for r in con.execute(
                "SELECT id, default_rate FROM users").fetchall()}
            # LABOUR IS COSTED ON THE HOURS THE EMPLOYEE ORIGINALLY SUBMITTED
            # (the frozen submitted_snapshot), NOT on a manager's later
            # adjustments to the live lines. Cost = what the work actually took;
            # invoicing = what we billed. Only Submitted/Approved/Locked count:
            # a Draft has no snapshot (never submitted) and a Rejected one has
            # been sent back, so neither is accepted work.
            snap_rows = con.execute(
                "SELECT user_id, submitted_snapshot FROM timesheets "
                "WHERE submitted_snapshot != '' "
                "AND status IN ('Submitted','Approved','Locked')").fetchall()
            lines = []
            for sr in snap_rows:
                try:
                    for sl in json.loads(sr["submitted_snapshot"]):
                        if sl.get("job_ref_id") is None:
                            continue
                        lines.append({"job_ref_id": sl.get("job_ref_id"),
                                      "user_id": sr["user_id"],
                                      "hours": sl.get("hours", 0) or 0,
                                      "rate": sl.get("rate", "Normal"),
                                      # £/hr frozen when this was submitted
                                      "rate_at_submit": sl.get("rate_at_submit")})
                except Exception:
                    continue      # a corrupt snapshot must not break Job Costs
            varis = {r["job_ref_id"]: (r["v"] or 0.0) for r in con.execute(
                "SELECT job_ref_id, SUM(value) v FROM job_variations "
                "GROUP BY job_ref_id").fetchall()}
            invs = {}
            for r in con.execute(
                    "SELECT job_ref_id, SUM(value) tot, "
                    "SUM(CASE WHEN status='Paid' THEN value ELSE 0 END) paid "
                    "FROM job_invoices GROUP BY job_ref_id").fetchall():
                invs[r["job_ref_id"]] = (r["tot"] or 0.0, r["paid"] or 0.0)
            blls = {}
            for r in con.execute(
                    "SELECT job_ref_id, SUM(value) tot, "
                    "SUM(CASE WHEN status='Paid' THEN value ELSE 0 END) paid "
                    "FROM job_bills GROUP BY job_ref_id").fetchall():
                blls[r["job_ref_id"]] = (r["tot"] or 0.0, r["paid"] or 0.0)
            agg = {}
            for l in lines:
                jid, uid = l["job_ref_id"], l["user_id"]
                mult = RATE_MULT.get(l.get("rate") or "Normal", 1.0)
                paid = float(l["hours"]) * mult
                # Prefer the rate frozen at submission. Fall back to today's
                # rate only for work submitted before rate-freezing existed, or
                # where no rate was set up yet (frozen as 0) - so configuring a
                # rate later still costs that work properly.
                frozen = l.get("rate_at_submit")
                if frozen:
                    rate = float(frozen)
                else:
                    rate = rates[(uid, jid)] if (uid, jid) in rates \
                        else defaults.get(uid, 0.0)
                h, c = agg.get(jid, (0.0, 0.0))
                agg[jid] = (h + float(l["hours"]), c + paid * rate)
            out = []
            for j in jobs:
                h, c = agg.get(j["id"], (0.0, 0.0))
                quote = j["quote_value"] or 0.0
                var = varis.get(j["id"], 0.0)
                cur_quote = quote + var
                billed, billed_paid = invs.get(j["id"], (0.0, 0.0))
                bill_tot, bill_paid = blls.get(j["id"], (0.0, 0.0))
                total_cost = c + bill_tot
                out.append({"id": j["id"], "job_ref": j["job_ref"],
                            "job_title": j["job_title"], "kind": j["kind"],
                            # Active / Closed - lets Job Costs filter on it
                            "status": j["status"],
                            "quote": round(quote, 2),
                            "variations": round(var, 2),
                            "quote_current": round(cur_quote, 2),
                            "hours": round(h, 2),
                            "cost": round(c, 2),
                            "bills": round(bill_tot, 2),
                            "bills_paid": round(bill_paid, 2),
                            "bills_unpaid": round(bill_tot - bill_paid, 2),
                            "total_cost": round(total_cost, 2),
                            "invoiced": round(billed, 2),
                            "invoiced_paid": round(billed_paid, 2),
                            "invoiced_unpaid": round(billed - billed_paid, 2),
                            "to_invoice": round(cur_quote - billed, 2),
                            "profit": round(billed - total_cost, 2),
                            "margin": (round((billed - total_cost) / billed * 100, 1)
                                       if billed else 0.0),
                            "remaining": round(cur_quote - total_cost, 2),
                            "pct": (round(total_cost / cur_quote * 100, 1)
                                    if cur_quote else 0.0),
                            # what proportion of the money we've actually
                            # BILLED has been eaten by cost. Over 100% means
                            # the job is losing money on the work invoiced.
                            "pct_cost_of_invoiced": (
                                round(total_cost / billed * 100, 1)
                                if billed else 0.0),
                            "pct_invoiced": (round(billed / cur_quote * 100, 1)
                                             if cur_quote else 0.0)})
            return out

    def get_or_create_timesheet(self, user_id, week_start):
        with self.conn() as con:
            row = con.execute(
                "SELECT * FROM timesheets WHERE user_id=? AND week_start=?",
                (user_id, iso(week_start))).fetchone()
            if row:
                return dict(row)
            cur = con.execute(
                "INSERT INTO timesheets(user_id, week_start, status, updated_at)"
                " VALUES(?,?, 'Draft', ?)",
                (user_id, iso(week_start), dt.datetime.now().isoformat(timespec="seconds")))
            return {"id": cur.lastrowid, "user_id": user_id,
                    "week_start": iso(week_start), "status": "Draft"}

    def get_lines(self, timesheet_id):
        with self.conn() as con:
            return [dict(r) for r in con.execute(
                """SELECT l.*, j.job_ref, j.job_title
                   FROM timesheet_lines l JOIN job_refs j ON j.id=l.job_ref_id
                   WHERE l.timesheet_id=? ORDER BY l.work_date, l.id""",
                (timesheet_id,)).fetchall()]

    def hours_by_timesheet(self):
        """Total hours per timesheet in ONE query (avoids a query per row on a
        network share)."""
        with self.conn() as con:
            return {r["timesheet_id"]: round(r["h"] or 0, 2) for r in con.execute(
                "SELECT timesheet_id, SUM(hours) h FROM timesheet_lines "
                "GROUP BY timesheet_id").fetchall()}

    def user_hours_by_date(self, user_id):
        """Total hours per work-date for one user in ONE query."""
        with self.conn() as con:
            return {r["work_date"]: round(r["h"] or 0, 2) for r in con.execute(
                "SELECT l.work_date, SUM(l.hours) h FROM timesheet_lines l "
                "JOIN timesheets t ON t.id=l.timesheet_id WHERE t.user_id=? "
                "GROUP BY l.work_date", (user_id,)).fetchall()}

    def save_timesheet_lines(self, actor_email, timesheet_id, lines, new_status=None):
        with self.conn() as con:
            con.execute("DELETE FROM timesheet_lines WHERE timesheet_id=?", (timesheet_id,))
            for ln in lines:
                con.execute(
                    """INSERT INTO timesheet_lines(timesheet_id, work_date, job_ref_id,
                                                   narrative, hours, job_role, rate)
                       VALUES(?,?,?,?,?,?,?)""",
                    (timesheet_id, ln["work_date"], ln["job_ref_id"], ln["narrative"],
                     float(ln["hours"]), ln.get("job_role", ""),
                     ln.get("rate", "Normal")))
            fields = ["updated_at=?"]
            params = [dt.datetime.now().isoformat(timespec="seconds")]
            if new_status:
                fields.append("status=?")
                params.append(new_status)
            params.append(timesheet_id)
            con.execute(f"UPDATE timesheets SET {', '.join(fields)} WHERE id=?", params)
            action = "SAVE_TIMESHEET" if not new_status else f"STATUS_{new_status.upper()}"
            self._audit(con, actor_email, action, "", f"timesheet {timesheet_id}")

    def set_timesheet_status(self, actor_email, timesheet_id, status, reason=""):
        with self.conn() as con:
            old = con.execute("SELECT status FROM timesheets WHERE id=?",
                              (timesheet_id,)).fetchone()["status"]
            con.execute(
                "UPDATE timesheets SET status=?, reject_reason=?, updated_at=? WHERE id=?",
                (status, reason, dt.datetime.now().isoformat(timespec="seconds"),
                 timesheet_id))
            self._audit(con, actor_email, f"STATUS_{status.upper()}", old,
                        f"timesheet {timesheet_id}" + (f" - {reason}" if reason else ""))

    def rates_for_user(self, con, user_id):
        """The £/hour that applies to this person right now: their per-job
        overrides plus their standard rate as the fallback."""
        overrides = {r["job_ref_id"]: (r["rate"] or 0.0) for r in con.execute(
            "SELECT job_ref_id, rate FROM job_rates WHERE user_id=?",
            (user_id,)).fetchall()}
        row = con.execute("SELECT default_rate FROM users WHERE id=?",
                          (user_id,)).fetchone()
        return overrides, ((row["default_rate"] or 0.0) if row else 0.0)

    def submit_timesheet(self, actor_email, timesheet_id, lines):
        """Save the employee's lines, freeze a snapshot of exactly what they
        submitted, and mark Submitted. The snapshot is what the employee will
        always see afterwards; managers work on the live lines.

        The snapshot ALSO FREEZES THE £/HOUR in force at this moment, per line
        ('rate_at_submit'). Job costing prefers that frozen figure, so putting
        rates up later can never retro-price work that has already been
        submitted - a finished job keeps the value it had when it was done."""
        self.save_timesheet_lines(actor_email, timesheet_id, lines,
                                  new_status="Submitted")
        saved = self.get_lines(timesheet_id)
        with self.conn() as con:
            row = con.execute("SELECT user_id FROM timesheets WHERE id=?",
                              (timesheet_id,)).fetchone()
            uid = row["user_id"] if row else None
            overrides, std = self.rates_for_user(con, uid) if uid else ({}, 0.0)
        snap = [{"work_date": l["work_date"], "job_ref_id": l["job_ref_id"],
                 "job_ref": l["job_ref"], "job_title": l["job_title"],
                 "job_role": l.get("job_role", ""), "rate": l.get("rate", "Normal"),
                 "narrative": l["narrative"], "hours": l["hours"],
                 "rate_at_submit": float(overrides.get(l["job_ref_id"], std) or 0.0)}
                for l in saved]
        with self.conn() as con:
            con.execute("UPDATE timesheets SET submitted_snapshot=?, reject_reason='' "
                        "WHERE id=?", (json.dumps(snap), timesheet_id))
            self._audit(con, actor_email, "SUBMIT_RATES_FROZEN", "",
                        f"timesheet {timesheet_id}: {len(snap)} line(s) priced at "
                        f"the rates in force today")

    def reject_timesheet(self, actor_email, timesheet_id, reason):
        """Reject back to the employee. The working copy is restored to exactly
        what the employee submitted, so they fix their own version."""
        with self.conn() as con:
            r = con.execute("SELECT submitted_snapshot FROM timesheets WHERE id=?",
                            (timesheet_id,)).fetchone()
            snap = json.loads(r["submitted_snapshot"]) if r and r["submitted_snapshot"] else []
        if snap:
            restore = [{"work_date": s["work_date"], "job_ref_id": s["job_ref_id"],
                        "narrative": s["narrative"], "hours": s["hours"],
                        "job_role": s.get("job_role", ""),
                        "rate": s.get("rate", "Normal")} for s in snap]
            self.save_timesheet_lines(actor_email, timesheet_id, restore)
        self.set_timesheet_status(actor_email, timesheet_id, "Rejected", reason)

    def employee_lines(self, ts_row):
        """What the OWNER should see for this timesheet: their frozen submitted
        snapshot once submitted/approved/locked, otherwise the live lines."""
        if ts_row.get("status") in ("Submitted", "Approved", "Locked"):
            snap = ts_row.get("submitted_snapshot") or ""
            if snap:
                return json.loads(snap)
        return self.get_lines(ts_row["id"])

    def user_default_rates(self):
        with self.conn() as con:
            return {r["id"]: (r["default_rate"] or 0.0) for r in con.execute(
                "SELECT id, default_rate FROM users").fetchall()}

    def list_timesheets(self, status=None, user_id=None):
        with self.conn() as con:
            q = ("SELECT t.*, u.full_name, u.email, u.job_role"
                 " FROM timesheets t JOIN users u ON u.id=t.user_id")
            clauses, params = [], []
            if status:
                clauses.append("t.status=?")
                params.append(status)
            if user_id:
                clauses.append("t.user_id=?")
                params.append(user_id)
            if clauses:
                q += " WHERE " + " AND ".join(clauses)
            q += " ORDER BY t.week_start DESC, u.full_name"
            return [dict(r) for r in con.execute(q, params).fetchall()]

    # -- expenses ------------------------------------------------------------
    def list_expenses(self, user_id=None):
        with self.conn() as con:
            q = ("SELECT e.*, u.full_name, u.email, j.job_ref, j.job_title "
                 "FROM expenses e JOIN users u ON u.id=e.user_id "
                 "LEFT JOIN job_refs j ON j.id=e.job_ref_id")
            params = []
            if user_id:
                q += " WHERE e.user_id=?"
                params.append(user_id)
            q += " ORDER BY e.exp_date DESC, e.id DESC"
            return [dict(r) for r in con.execute(q, params).fetchall()]

    def create_expense(self, actor_email, user_id, exp_date, job_ref_id,
                       description, value, receipt):
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO expenses(user_id, exp_date, job_ref_id, description,
                                        value, receipt, status, created_at)
                   VALUES(?,?,?,?,?,?, 'Logged', ?)""",
                (user_id, exp_date, job_ref_id, description.strip(), float(value),
                 receipt, dt.datetime.now().isoformat(timespec="seconds")))
            self._audit(con, actor_email, "CREATE_EXPENSE", "",
                        f"{exp_date} {value}")
            return cur.lastrowid

    def update_expense(self, actor_email, exp_id, exp_date, job_ref_id,
                       description, value, receipt):
        with self.conn() as con:
            con.execute(
                """UPDATE expenses SET exp_date=?, job_ref_id=?, description=?,
                                       value=?, receipt=? WHERE id=?""",
                (exp_date, job_ref_id, description.strip(), float(value),
                 receipt, exp_id))
            self._audit(con, actor_email, "UPDATE_EXPENSE", "", str(exp_id))

    def delete_expense(self, actor_email, exp_id):
        with self.conn() as con:
            con.execute("DELETE FROM expenses WHERE id=?", (exp_id,))
            self._audit(con, actor_email, "DELETE_EXPENSE", str(exp_id), "")

    def set_expense_payment(self, actor_email, exp_id, payment_receipt):
        """Admin attaches a proof-of-payment receipt and marks the expense Paid;
        the employee can then see/open it."""
        with self.conn() as con:
            con.execute("UPDATE expenses SET payment_receipt=?, status='Paid' WHERE id=?",
                        (payment_receipt, exp_id))
            self._audit(con, actor_email, "EXPENSE_PAID", "", str(exp_id))


# ----------------------------------------------------------------------------
# Validation
# ----------------------------------------------------------------------------

# ----------------------------------------------------------------------------
# Document reading (offline) - used to pre-fill a bill/invoice from a scan
#
# NOTHING here is imported at module level. Every heavy import happens inside
# the function, only when someone actually presses "Scan document", so the app
# starts exactly as fast as it did before whether or not OCR is installed.
# ----------------------------------------------------------------------------

OCR_INSTALL_HINT = (
    "Reading documents needs Tesseract OCR on this PC.\n\n"
    "1. Install Tesseract (UB-Mannheim build for Windows), and\n"
    "2. Run:  pip install pytesseract pillow\n\n"
    "Only the people who enter bills need this - everyone else is unaffected. "
    "You can always type the details in by hand instead.")


def _tesseract_cmd():
    """Find tesseract.exe without requiring it on PATH: check the config, then
    the usual install locations, then alongside the app itself."""
    cands = []
    try:
        cfg = load_config()
        if cfg.get("tesseract_path"):
            cands.append(cfg["tesseract_path"])
    except Exception:
        pass
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    cands += [os.path.join(pf, "Tesseract-OCR", "tesseract.exe"),
              os.path.join(pf86, "Tesseract-OCR", "tesseract.exe"),
              os.path.join(APP_DIR, "Tesseract-OCR", "tesseract.exe"),
              os.path.join(APP_DIR, "tesseract.exe")]
    for c in cands:
        try:
            if c and os.path.exists(c):
                return c
        except Exception:
            pass
    return None


def _pdf_text(path):
    """Most emailed supplier invoices are digital PDFs carrying a real text
    layer. Reading that is far more accurate than OCR-ing a picture of it, so
    always try this first and only fall back to OCR for scans/photos."""
    try:
        from pdfminer.high_level import extract_text
        from pdfminer.layout import LAParams
        # boxes_flow=None orders the text purely by its position on the page,
        # which keeps a label next to its own amount. pdfminer's DEFAULT
        # column-detection groups all the labels together and all the values
        # together ('Net amount' / 'Tax' / 'Amount Due' then 4,250.00 /
        # 850.00 / 5,100.00), making it impossible to tell which figure
        # belongs to which label. Always try positional order first.
        t = extract_text(path, laparams=LAParams(boxes_flow=None)) or ""
        if t.strip():
            return t
    except Exception:
        pass
    try:
        from pdfminer.high_level import extract_text
        t = extract_text(path) or ""
        if t.strip():
            return t
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        return "\n".join((p.extract_text() or "") for p in PdfReader(path).pages)
    except Exception:
        return ""


def _image_text(path):
    """OCR an image with Tesseract. Raises RuntimeError with a friendly message
    if the pieces aren't installed."""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        raise RuntimeError(OCR_INSTALL_HINT)
    cmd = _tesseract_cmd()
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
    try:
        return pytesseract.image_to_string(Image.open(path))
    except Exception as e:
        if "tesseract" in str(e).lower():
            raise RuntimeError(OCR_INSTALL_HINT)
        raise


_MONEY_RE = r"[-+]?\(?\s*(?:£|GBP)?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d+\.\d{2}"


def _money(s):
    """'£1,234.56' / '(1,234.56)' -> float. Brackets mean negative (credit)."""
    if s is None:
        return None
    t = str(s).strip()
    neg = t.startswith("(") and t.endswith(")")
    t = re.sub(r"[^\d.\-]", "", t)
    if t in ("", "-", "."):
        return None
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


# A real money amount on an invoice essentially always carries either 2 decimal
# places or thousands separators. Insisting on that is what stops a payment
# term ("30 days net") or a VAT rate ("VAT @ 20%") being read as an amount.
_STRICT_MONEY_RE = r"\d{1,3}(?:,\d{3})+(?:\.\d{2})?|\d+\.\d{2}"
_ONLY_MONEY_RE = r"^\(?\s*(?:£|GBP)?\s*(" + _STRICT_MONEY_RE + r")\s*\)?\s*(CR)?$"
_DATE_RE = (r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}"
            r"|\d{4}-\d{2}-\d{2}"
            r"|\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{2,4}"
            r"|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}")
# Payment terms mention 'net' but are NOT an amount - never read a value here.
_TERMS_RE = r"payment\s*terms|days?\s*net|net\s*\d+\s*days?|net\s*monthly|terms\s*:"


def parse_document_fields(text):
    """Best-effort extraction of the fields we care about from raw document
    text. Deliberately conservative - anything it isn't reasonably sure about
    comes back empty for the user to fill in. This ALWAYS pre-fills a form for
    a human to check; it never saves anything by itself.

    Written to cope with the two things real PDFs do that break naive parsing:
      1. a label and its value often land on SEPARATE extracted lines
         ('Sub Total' then '10,000.00'), so every lookup falls through to the
         following line when its own line carries no value; and
      2. numbers that are not amounts sit right next to money-ish words
         ('30 days net', 'VAT @ 20%'), so amounts must look like money and
         payment-terms lines are excluded outright."""
    out = {"supplier": "", "number": "", "date": "", "net": None, "vat": None,
           "gross": None, "value": None, "reconciles": False,
           "value_basis": "none"}
    if not text or not text.strip():
        return out
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    def after_label(pattern, validate, lookahead=2, skip=None):
        """Value belonging to a label: the rest of the label's own line if it
        has one, otherwise the next non-empty line (unless that is itself
        another label, which would mean this label simply had no value)."""
        for i, ln in enumerate(lines):
            m = re.search(pattern, ln, re.I)
            if not m:
                continue
            if skip and re.search(skip, ln, re.I):
                continue
            rest = ln[m.end():].lstrip(" .:#-\t")
            if rest.strip():
                v = validate(rest.strip())
                if v is not None:
                    return v
            for j in range(i + 1, min(i + 1 + lookahead, len(lines))):
                nxt = lines[j].strip()
                if nxt.endswith(":"):
                    break                      # next label - this one had none
                v = validate(nxt)
                if v is not None:
                    return v
                break                          # only the immediately next line
        return None

    # --- document number ---
    def as_ref(s):
        s = s.strip().strip(".,;:")
        if not (3 <= len(s) <= 25) or not re.search(r"\d", s):
            return None
        if _norm_doc_date(s):
            return None                        # that's a date, not a reference
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-/_ ]*", s):
            return None
        return s

    out["number"] = after_label(
        r"(?:invoice|bill|document|inv)\s*(?:no\b|number|#|ref)", as_ref) or ""
    if not out["number"]:
        for pat in (r"\b(INV[-/ ]?\d{3,10})\b", r"\b([A-Z]{2,5}[-/]\d{3,10})\b"):
            m = re.search(pat, text, re.I)
            if m:
                out["number"] = m.group(1).strip(" .-/")
                break

    # --- date: prefer one explicitly labelled as the invoice/bill date ---
    def as_date(s):
        m = re.search(_DATE_RE, s)
        return _norm_doc_date(m.group(0)) if m else None

    out["date"] = (after_label(r"(?:invoice|bill|document|tax\s*point)\s*date",
                               as_date)
                   or after_label(r"\bdate\b", as_date, skip=_TERMS_RE)
                   or (as_date(text) or "")) or ""

    # --- amounts ---
    def amount(*labels, skip=None):
        for w in labels:
            for i, ln in enumerate(lines):
                if not re.search(w, ln, re.I):
                    continue
                if skip and re.search(skip, ln, re.I):
                    continue
                # a figure on the label's own line: take the right-most one
                vals = [mm.group(0) for mm in re.finditer(_STRICT_MONEY_RE, ln)]
                if vals:
                    v = _money(vals[-1])
                    if v is not None:
                        return v
                # otherwise the next line, but only if it is JUST an amount
                if i + 1 < len(lines):
                    mm = re.match(_ONLY_MONEY_RE, lines[i + 1].strip())
                    if mm:
                        v = _money(mm.group(1))
                        if v is not None:
                            return -v if mm.group(2) else v
        return None

    out["vat"] = amount(r"\bvat\b(?!\s*(?:no\b|number|reg))", r"\btax\b",
                        skip=_TERMS_RE)
    out["net"] = amount(r"sub\s*-?\s*total", r"goods\s*total", r"total\s*excl",
                        r"\bnet\s*(?:total|amount|value)\b", r"\bnet\b",
                        skip=_TERMS_RE)
    out["gross"] = amount(r"amount\s*due", r"balance\s*due", r"total\s*due",
                          r"grand\s*total", r"total\s*inc", r"\btotal\b",
                          skip=_TERMS_RE)
    # Fill in a missing figure only when the other two are known.
    if out["net"] is None and out["gross"] is not None and out["vat"] is not None:
        cand = round(out["gross"] - out["vat"], 2)
        if cand > 0:
            out["net"] = cand
    if out["gross"] is None and out["net"] is not None and out["vat"] is not None:
        out["gross"] = round(out["net"] + out["vat"], 2)
    # Last resort: the largest money-shaped figure anywhere on the page.
    if out["net"] is None and out["gross"] is None:
        vals = [v for v in (_money(x) for x in re.findall(_STRICT_MONEY_RE, text))
                if v is not None and abs(v) >= 1]
        if vals:
            out["gross"] = max(vals)
    # Does net + VAT actually add up to the gross? Shown to the user so they
    # can trust (or distrust) the reading at a glance.
    if None not in (out["net"], out["vat"], out["gross"]):
        out["reconciles"] = abs(out["net"] + out["vat"] - out["gross"]) < 0.02
    # The figure we pre-fill must be EX-VAT: VAT is reclaimable, so it must
    # never land in a job cost. 'value_basis' records how sure we are:
    #   net     - read a real net/sub-total off the document
    #   derived - worked out as gross minus a VAT figure we also read
    #   unknown - only one total was found and the document never mentions
    #             VAT, so we cannot tell whether it already includes it. The
    #             UI warns loudly in that case rather than quietly risking a
    #             VAT-inclusive figure being saved as a cost.
    if out["net"] is not None:
        out["value"] = out["net"]
        out["value_basis"] = "net" if out["vat"] is None or out["reconciles"] \
            else "derived"
        if out["gross"] is not None and out["vat"] is not None \
                and not out["reconciles"]:
            out["value_basis"] = "derived"
    elif out["gross"] is not None:
        out["value"] = out["gross"]
        out["value_basis"] = "unknown"
    else:
        out["value"] = None
        out["value_basis"] = "none"
    # A VAT-registered supplier nearly always prints the word somewhere. If it
    # never appears, a single total is most likely already VAT-free.
    if out["value_basis"] == "unknown" and not re.search(r"\bvat\b", text, re.I):
        out["value_basis"] = "no_vat_shown"

    # --- supplier: first meaningful line that isn't a heading ---
    skip_words = ("invoice", "bill", "statement", "tax", "vat", "date", "page",
                  "remittance", "purchase order", "delivery", "application for")
    for l in lines[:12]:
        ls = l.lower()
        if any(ls.startswith(s) for s in skip_words):
            continue
        if len(l) < 3 or len(l) > 60:
            continue
        if sum(ch.isdigit() for ch in l) > len(l) / 2:
            continue
        out["supplier"] = l.strip(" .,:-")
        break
    return out


def _norm_doc_date(s):
    """Dates on invoices are usually UK order (dd/mm/yyyy). _to_iso_date tries
    that first already, so reuse it, plus handle '14 Mar 2026' style."""
    iso = _to_iso_date(s)
    if iso:
        return iso
    for fmt in ("%d %b %Y", "%d %B %Y", "%b %d %Y", "%B %d %Y",
                "%d %b %y", "%d %B %y"):
        try:
            return dt.datetime.strptime(str(s).replace(",", "").strip(),
                                        fmt).date().isoformat()
        except ValueError:
            pass
    return None


def read_document(path):
    """Read a bill/invoice document and return the fields to pre-fill.
    Never raises - always returns a dict with 'ok' so the UI can explain."""
    if not path or not os.path.exists(path):
        return {"ok": False, "error": "That file could not be found."}
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            text = _pdf_text(path)
            if not (text or "").strip():
                # scanned PDF with no text layer - needs real OCR, which needs
                # page rasterising; tell the user rather than failing silently
                return {"ok": False, "error":
                        "That PDF has no readable text layer - it's a scan. "
                        "Save the page as a JPG/PNG and scan that instead, or "
                        "type the details in by hand."}
        else:
            text = _image_text(path)
    except RuntimeError as e:
        return {"ok": False, "error": str(e)}
    except Exception as e:
        return {"ok": False, "error": "Could not read that document: " + str(e)}
    f = parse_document_fields(text)
    f["ok"] = True
    f["found"] = sum(1 for k in ("supplier", "number", "date") if f.get(k)) \
        + (1 if f.get("value") is not None else 0)
    return f


def validate_bill(supplier, value):
    """A bill needs a supplier and a non-zero value. Negative IS allowed -
    that's a supplier credit note."""
    if not (supplier or "").strip():
        return False, "Supplier is required."
    try:
        v = float(value)
    except (TypeError, ValueError):
        return False, "Bill value must be a number."
    if v == 0:
        return False, "Bill value cannot be zero."
    return True, ""


def validate_invoice(invoice_no, value):
    """An invoice must have a number (so duplicates can be spotted on import)
    and a non-zero value. Negative IS allowed - that's a credit note."""
    if not (invoice_no or "").strip():
        return False, "Invoice number is required."
    try:
        v = float(value)
    except (TypeError, ValueError):
        return False, "Invoice value must be a number."
    if v == 0:
        return False, "Invoice value cannot be zero."
    return True, ""


def validate_line(narrative, hours):
    if not narrative or not narrative.strip():
        return False, "Narrative is mandatory."
    try:
        h = float(hours)
    except (TypeError, ValueError):
        return False, "Hours must be a number."
    if h < 0:
        return False, "Hours cannot be negative."
    if h == 0:
        return False, "Hours must be greater than zero."
    return True, ""


# ----------------------------------------------------------------------------
# Reports
# ----------------------------------------------------------------------------

class Reports:
    def __init__(self, db):
        self.db = db

    def _base_rows(self, date_from=None, date_to=None, user_id=None,
                   job_ref_id=None, job_role=None, status=None):
        q = ("SELECT u.id AS user_id, u.full_name, u.email, "
             "       COALESCE(NULLIF(l.job_role,''), u.job_role) AS job_role, "
             "       t.status, t.week_start, l.work_date, j.id AS job_ref_id, "
             "       j.job_ref, j.job_title, "
             "       l.narrative, l.hours, l.rate "
             "FROM timesheet_lines l "
             "JOIN timesheets t ON t.id=l.timesheet_id "
             "JOIN users u ON u.id=t.user_id "
             "JOIN job_refs j ON j.id=l.job_ref_id WHERE 1=1")
        params = []
        if date_from:
            q += " AND l.work_date>=?"; params.append(iso(date_from))
        if date_to:
            q += " AND l.work_date<=?"; params.append(iso(date_to))
        if user_id:
            q += " AND u.id=?"; params.append(user_id)
        if job_ref_id:
            q += " AND j.id=?"; params.append(job_ref_id)
        if job_role:
            q += " AND COALESCE(NULLIF(l.job_role,''), u.job_role)=?"
            params.append(job_role)
        if status:
            q += " AND t.status=?"; params.append(status)
        q += " ORDER BY u.full_name, l.work_date"
        with self.db.conn() as con:
            return [dict(r) for r in con.execute(q, params).fetchall()]

    def _snapshot_rows(self, date_from=None, date_to=None, user_id=None,
                       job_ref_id=None, job_role=None, status=None):
        """Rows built from each timesheet's frozen submitted snapshot - i.e.
        the ACTUAL hours exactly as the employee submitted them."""
        q = ("SELECT t.status, t.week_start, t.submitted_snapshot, "
             "       u.id AS uid, u.full_name, u.email, "
             "       u.job_role AS user_role "
             "FROM timesheets t JOIN users u ON u.id=t.user_id "
             "WHERE t.submitted_snapshot!=''")
        params = []
        if user_id:
            q += " AND u.id=?"; params.append(user_id)
        if status:
            q += " AND t.status=?"; params.append(status)
        with self.db.conn() as con:
            trows = [dict(r) for r in con.execute(q, params).fetchall()]
        df = iso(date_from) if date_from else None
        dtt = iso(date_to) if date_to else None
        out = []
        for t in trows:
            try:
                snap = json.loads(t["submitted_snapshot"])
            except Exception:
                snap = []
            for sline in snap:
                wd = sline.get("work_date", "")
                if df and wd < df:
                    continue
                if dtt and wd > dtt:
                    continue
                role = sline.get("job_role") or t["user_role"]
                if job_role and role != job_role:
                    continue
                if job_ref_id and sline.get("job_ref_id") != job_ref_id:
                    continue
                out.append({
                    "full_name": t["full_name"], "email": t["email"],
                    "user_id": t["uid"],
                    "job_role": role, "status": t["status"],
                    "week_start": t["week_start"], "work_date": wd,
                    "job_ref_id": sline.get("job_ref_id"),
                    "job_ref": sline.get("job_ref", ""),
                    "job_title": sline.get("job_title", ""),
                    "narrative": sline.get("narrative", ""),
                    "hours": sline.get("hours", 0),
                    "rate": sline.get("rate", "Normal")})
        out.sort(key=lambda r: (r["full_name"], r["work_date"]))
        return out

    def _rates_and_defaults(self):
        """Same lookup job_cost_summary() uses: per-job per-person rate override,
        falling back to the person's standard rate."""
        with self.db.conn() as con:
            rates = {(r["user_id"], r["job_ref_id"]): r["rate"]
                     for r in con.execute("SELECT user_id, job_ref_id, rate "
                                          "FROM job_rates").fetchall()}
            defaults = {r["id"]: (r["default_rate"] or 0.0) for r in con.execute(
                "SELECT id, default_rate FROM users").fetchall()}
        return rates, defaults

    def add_cost(self, rows):
        """Cost = paid hours (hours x overtime multiplier) x the person's rate
        for that job (falling back to their standard rate) - the same
        calculation job_cost_summary()/job_cost_breakdown() use, reused here
        so Reports agrees with Job Costs / Analytics."""
        rates, defaults = self._rates_and_defaults()
        for r in rows:
            uid, jid = r.get("user_id"), r.get("job_ref_id")
            rate_val = rates.get((uid, jid), defaults.get(uid, 0.0))
            r["cost"] = round(paid_hours(r["hours"], r.get("rate", "Normal")) * rate_val, 2)
        return rows

    def person_report(self, basis="adjusted", include_cost=False, **f):
        rows = self._snapshot_rows(**f) if basis == "submitted" else self._base_rows(**f)
        if include_cost:
            self.add_cost(rows)
        return rows, round(sum(r["hours"] for r in rows), 2)

    def summary_by(self, key, **f):
        rows = self._base_rows(**f)
        agg = {}
        for r in rows:
            agg[r[key]] = round(agg.get(r[key], 0) + r["hours"], 2)
        return agg

    def missing_timesheets(self, week_start):
        ws = iso(week_start)
        with self.db.conn() as con:
            users = con.execute("SELECT * FROM users WHERE active=1").fetchall()
            out = []
            for u in users:
                row = con.execute(
                    "SELECT status FROM timesheets WHERE user_id=? AND week_start=?",
                    (u["id"], ws)).fetchone()
                status = row["status"] if row else "None"
                if status in ("None", "Draft", "Rejected"):
                    out.append({"full_name": u["full_name"], "email": u["email"],
                                "job_role": u["job_role"], "status": status})
            return out


# ----------------------------------------------------------------------------
# Exports
# ----------------------------------------------------------------------------

REPORT_COLUMNS = [
    ("full_name", "Name"), ("job_role", "Role"), ("work_date", "Date"),
    ("job_ref", "Job Ref"), ("job_title", "Job Title"),
    ("narrative", "Narrative"), ("hours", "Hours"), ("cost", "Cost (£)"),
    ("status", "Status"),
]


def _meta_lines(title, generated_by, filters):
    return [title,
            f"Generated: {dt.datetime.now().strftime('%Y-%m-%d %H:%M')}",
            f"Generated by: {generated_by}",
            f"Filters: {filters}"]


def export_csv(path, title, generated_by, filters, rows, company="",
               columns=None):
    columns = columns or REPORT_COLUMNS
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow([company])
        for line in _meta_lines(title, generated_by, filters):
            w.writerow([line])
        w.writerow([])
        w.writerow([h for _, h in columns])
        for r in rows:
            w.writerow([r.get(k, "") for k, _ in columns])
        w.writerow([])
        total_row = []
        for k, _ in columns:
            if k == "narrative":
                total_row.append("TOTAL")
            elif k == "hours":
                total_row.append(round(sum(r.get("hours", 0) for r in rows), 2))
            elif k == "cost":
                total_row.append(round(sum(r.get("cost", 0) for r in rows), 2))
            else:
                total_row.append("")
        w.writerow(total_row)
    return path


REPORT_COL_WIDTHS = {"full_name": 22, "job_role": 16, "work_date": 12, "job_ref": 12,
                     "job_title": 22, "narrative": 40, "hours": 8, "cost": 12,
                     "status": 12}


def export_xlsx(path, title, generated_by, filters, rows, company="",
                columns=None):
    columns = columns or REPORT_COLUMNS
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    bold = Font(bold=True)
    ws["A1"] = "[COMPANY LOGO]"
    ws["A1"].font = Font(italic=True, color="888888")
    ws["A2"] = company
    ws["A2"].font = Font(bold=True, size=14)
    r = 3
    for line in _meta_lines(title, generated_by, filters):
        ws.cell(row=r, column=1, value=line)
        r += 1
    r += 1
    header_fill = PatternFill("solid", fgColor="1C1C24")
    for c, (_, h) in enumerate(columns, start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    r += 1
    for row in rows:
        for c, (k, _) in enumerate(columns, start=1):
            ws.cell(row=r, column=c, value=row.get(k, ""))
        r += 1
    for c, (k, _) in enumerate(columns, start=1):
        if k == "narrative":
            ws.cell(row=r, column=c, value="TOTAL").font = bold
        elif k == "hours":
            ws.cell(row=r, column=c,
                    value=round(sum(rw.get("hours", 0) for rw in rows), 2)).font = bold
        elif k == "cost":
            ws.cell(row=r, column=c,
                    value=round(sum(rw.get("cost", 0) for rw in rows), 2)).font = bold
    for i, (k, _) in enumerate(columns, start=1):
        wdt = REPORT_COL_WIDTHS.get(k, 14)
        ws.column_dimensions[chr(64 + i)].width = wdt
    wb.save(path)
    return path


def _to_iso_date(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, str):
        v = v.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%m/%d/%Y"):
            try:
                return dt.datetime.strptime(v, fmt).date().isoformat()
            except ValueError:
                pass
    return None


def build_timesheet_template(path, jobs, roles, company="",
                             week_start=None, person_name="", person_email=""):
    """Create a weekly Excel timesheet. If week_start is given, the 7 day rows
    are pre-filled with that week's dates so the person doesn't type them."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    GREEN = "84C341"
    thin = Side(style="thin", color="C7CDD4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = Workbook()
    ws = wb.active
    ws.title = "Timesheet"
    ws["A1"] = company + "  -  Weekly Timesheet"
    ws["A1"].font = Font(bold=True, size=16, color="2B2F36")
    ws["A2"] = "Fill in the table below (yellow cells are pre-filled), then email this file back."
    ws["A2"].font = Font(italic=True, color="777777")
    for cell, lab in [("A4", "Full name:"), ("A5", "Email:"),
                      ("A6", "Week commencing (Monday):")]:
        ws[cell] = lab
        ws[cell].font = Font(bold=True)
    if person_name:
        ws["B4"] = person_name
    if person_email:
        ws["B5"] = person_email
    if week_start:
        ws["B6"] = dt.datetime(week_start.year, week_start.month, week_start.day)
    for cell in ("B4", "B5", "B6"):
        ws[cell].fill = PatternFill("solid", fgColor="FFF6CC")
        ws[cell].border = border
    ws["B6"].number_format = "dd/mm/yyyy"
    hdr = ["Date", "Job Reference", "Role", "Description of work", "Hours", "Overtime"]
    r0 = 8
    for j, htxt in enumerate(hdr):
        c = ws.cell(row=r0, column=1 + j, value=htxt)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=GREEN)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws.row_dimensions[r0].height = 20
    for col, w in zip("ABCDEF", [16, 26, 18, 44, 9, 12]):
        ws.column_dimensions[col].width = w
    day_rows, extra = 7, 14
    nrows = day_rows + extra
    for i in range(nrows):
        r = r0 + 1 + i
        stripe = "F4F7FA" if (i % 2) else "FFFFFF"
        for cidx in range(1, 7):
            cc = ws.cell(row=r, column=cidx)
            cc.border = border
            cc.fill = PatternFill("solid", fgColor=stripe)
        ws.cell(row=r, column=1).number_format = "ddd dd/mm/yyyy"
        ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=6).alignment = Alignment(horizontal="center")
        if week_start and i < day_rows:
            d = week_start + dt.timedelta(days=i)
            ws.cell(row=r, column=1, value=dt.datetime(d.year, d.month, d.day))
    # dropdown source lists on a hidden sheet
    lst = wb.create_sheet("Lists")
    lst.sheet_state = "hidden"
    jobrefs = [j["job_ref"] for j in jobs]
    for i, jr in enumerate(jobrefs):
        lst.cell(row=1 + i, column=1, value=jr)
    for i, rl in enumerate(roles):
        lst.cell(row=1 + i, column=2, value=rl)
    for i, ot in enumerate(["Normal", "x1.5", "x2"]):
        lst.cell(row=1 + i, column=3, value=ot)
    dv_job = DataValidation(type="list",
                            formula1=f"Lists!$A$1:$A${max(1, len(jobrefs))}", allow_blank=True)
    dv_role = DataValidation(type="list",
                             formula1=f"Lists!$B$1:$B${max(1, len(roles))}", allow_blank=True)
    dv_ot = DataValidation(type="list", formula1="Lists!$C$1:$C$3", allow_blank=True)
    for dv in (dv_job, dv_role, dv_ot):
        ws.add_data_validation(dv)
    dv_job.add(f"B{r0+1}:B{r0+nrows}")
    dv_role.add(f"C{r0+1}:C{r0+nrows}")
    dv_ot.add(f"F{r0+1}:F{r0+nrows}")
    ws.freeze_panes = f"A{r0+1}"
    wb.save(path)


def parse_timesheet_xlsx(path):
    """Read a filled-in external timesheet back into a dict."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb["Timesheet"] if "Timesheet" in wb.sheetnames else wb.active

    def val(a):
        return ws[a].value

    def txt(a):
        v = val(a)
        return str(v).strip() if v is not None else ""

    out = {"name": txt("B4"), "email": txt("B5"), "week": val("B6"), "lines": []}
    r0 = 8
    for r in range(r0 + 1, ws.max_row + 1):
        jref = ws.cell(row=r, column=2).value
        hrs = ws.cell(row=r, column=5).value
        if (jref is None or str(jref).strip() == "") and \
           (hrs is None or str(hrs).strip() == ""):
            continue
        out["lines"].append({
            "date": ws.cell(row=r, column=1).value,
            "job_ref": str(ws.cell(row=r, column=2).value or "").strip(),
            "role": str(ws.cell(row=r, column=3).value or "").strip(),
            "narrative": str(ws.cell(row=r, column=4).value or "").strip(),
            "hours": ws.cell(row=r, column=5).value,
            "overtime": str(ws.cell(row=r, column=6).value or "Normal").strip(),
        })
    return out


def build_invoice_template(path, jobs, company=""):
    """Blank sheet for bulk-importing invoices, with a Job Reference dropdown
    driven by the live job list so refs always match something in the app."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    GREEN = "84C341"
    thin = Side(style="thin", color="C7CDD4")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    ws["A1"] = company + "  -  Invoice Import"
    ws["A1"].font = Font(bold=True, size=16, color="2B2F36")
    ws["A2"] = ("Fill in one row per invoice, then import this file under Job Costs. "
                "Job Reference and Invoice No are required. Credit notes: enter a "
                "negative value. Rows whose invoice number already exists on that "
                "job are skipped, so re-importing is safe.")
    ws["A2"].font = Font(italic=True, color="777777")
    hdr = ["Job Reference", "Invoice No", "Date", "Description", "Value (£)", "Status"]
    r0 = 4
    for j, htxt in enumerate(hdr):
        c = ws.cell(row=r0, column=1 + j, value=htxt)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=GREEN)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws.row_dimensions[r0].height = 20
    for col, w in zip("ABCDEF", [26, 18, 14, 46, 14, 12]):
        ws.column_dimensions[col].width = w
    nrows = 60
    for i in range(nrows):
        r = r0 + 1 + i
        stripe = "F4F7FA" if (i % 2) else "FFFFFF"
        for cidx in range(1, 7):
            cc = ws.cell(row=r, column=cidx)
            cc.border = border
            cc.fill = PatternFill("solid", fgColor=stripe)
        ws.cell(row=r, column=3).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=5).number_format = "#,##0.00"
    lst = wb.create_sheet("Lists")
    lst.sheet_state = "hidden"
    jobrefs = [j["job_ref"] for j in jobs]
    for i, jr in enumerate(jobrefs):
        lst.cell(row=1 + i, column=1, value=jr)
    for i, st in enumerate(["Raised", "Paid"]):
        lst.cell(row=1 + i, column=2, value=st)
    dv_job = DataValidation(type="list",
                            formula1=f"Lists!$A$1:$A${max(1, len(jobrefs))}",
                            allow_blank=True)
    dv_st = DataValidation(type="list", formula1="Lists!$B$1:$B$2", allow_blank=True)
    for dv in (dv_job, dv_st):
        ws.add_data_validation(dv)
    dv_job.add(f"A{r0+1}:A{r0+nrows}")
    dv_st.add(f"F{r0+1}:F{r0+nrows}")
    ws.freeze_panes = f"A{r0+1}"
    wb.save(path)
    return path


def parse_invoices_xlsx(path):
    """Read a filled-in invoice import sheet into a list of dicts. Blank rows
    are ignored; anything else is handed back for validation by the caller."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb["Invoices"] if "Invoices" in wb.sheetnames else wb.active
    out = []
    r0 = 4
    for r in range(r0 + 1, ws.max_row + 1):
        ref = ws.cell(row=r, column=1).value
        no = ws.cell(row=r, column=2).value
        val = ws.cell(row=r, column=5).value
        if all(x is None or str(x).strip() == "" for x in (ref, no, val)):
            continue
        out.append({
            "job_ref": str(ref or "").strip(),
            "invoice_no": str(no or "").strip(),
            "inv_date": _to_iso_date(ws.cell(row=r, column=3).value) or "",
            "description": str(ws.cell(row=r, column=4).value or "").strip(),
            "value": val,
            "status": str(ws.cell(row=r, column=6).value or "Raised").strip(),
        })
    return out


def export_pdf(path, title, generated_by, filters, rows, company=""):
    from fpdf import FPDF
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, company, ln=1)
    pdf.set_font("Helvetica", "", 9)
    for line in _meta_lines(title, generated_by, filters):
        pdf.cell(0, 5, line, ln=1)
    pdf.ln(2)
    headers = [h for _, h in REPORT_COLUMNS]
    widths = [40, 28, 22, 24, 45, 70, 14, 24]
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(28, 28, 36)
    pdf.set_text_color(255, 255, 255)
    for h, w in zip(headers, widths):
        pdf.cell(w, 6, h, border=1, fill=True)
    pdf.ln()
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 7)
    for row in rows:
        for (k, _), w in zip(REPORT_COLUMNS, widths):
            val = str(row.get(k, ""))
            if len(val) > 60:
                val = val[:57] + "..."
            pdf.cell(w, 5, val, border=1)
        pdf.ln()
    pdf.set_font("Helvetica", "B", 8)
    total = round(sum(r.get("hours", 0) for r in rows), 2)
    pdf.cell(sum(widths[:5]), 6, "TOTAL", border=1)
    pdf.cell(widths[5], 6, "", border=1)
    pdf.cell(widths[6], 6, str(total), border=1)
    pdf.output(path)
    return path


def export_month_xlsx(path, company, month_label, generated_by, rows):
    """Month-end summary workbook: By Person, By Job, and full Detail sheets."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    hdr = PatternFill("solid", fgColor="1C1C24")
    whitebold = Font(bold=True, color="FFFFFF")

    def head(ws, headers, r):
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = whitebold
            cell.fill = hdr

    wb = Workbook()

    # --- By Person ---
    ws1 = wb.active
    ws1.title = "By Person"
    ws1["A1"] = company
    ws1["A1"].font = Font(bold=True, size=14)
    ws1["A2"] = f"Month-end summary  -  {month_label}"
    ws1["A3"] = (f"Generated {dt.datetime.now():%Y-%m-%d %H:%M} by {generated_by}")
    by_person = {}
    for r in rows:
        k = (r["full_name"], r["job_role"])
        w, p = by_person.get(k, (0, 0))
        by_person[k] = (round(w + r["hours"], 2),
                        round(p + paid_hours(r["hours"], r.get("rate", "Normal")), 2))
    head(ws1, ["Name", "Role", "Hours worked", "Paid hours"], 5)
    rr = 6
    tot_w = tot_p = 0
    for (name, role), (w, p) in sorted(by_person.items()):
        ws1.cell(row=rr, column=1, value=name)
        ws1.cell(row=rr, column=2, value=role)
        ws1.cell(row=rr, column=3, value=w)
        ws1.cell(row=rr, column=4, value=p)
        tot_w += w
        tot_p += p
        rr += 1
    ws1.cell(row=rr, column=2, value="TOTAL").font = Font(bold=True)
    ws1.cell(row=rr, column=3, value=round(tot_w, 2)).font = Font(bold=True)
    ws1.cell(row=rr, column=4, value=round(tot_p, 2)).font = Font(bold=True)
    for col, w in zip("ABCD", [26, 18, 14, 12]):
        ws1.column_dimensions[col].width = w

    # --- By Job ---
    ws2 = wb.create_sheet("By Job")
    by_job = {}
    for r in rows:
        k = (r["job_ref"], r["job_title"])
        by_job[k] = round(by_job.get(k, 0) + r["hours"], 2)
    head(ws2, ["Job Ref", "Job Title", "Total hours"], 1)
    rr = 2
    for (ref, title), h in sorted(by_job.items(), key=lambda x: -x[1]):
        ws2.cell(row=rr, column=1, value=ref)
        ws2.cell(row=rr, column=2, value=title)
        ws2.cell(row=rr, column=3, value=h)
        rr += 1
    for col, w in zip("ABC", [18, 30, 14]):
        ws2.column_dimensions[col].width = w

    # --- Detail ---
    ws3 = wb.create_sheet("Detail")
    head(ws3, ["Name", "Role", "Date", "Job Ref", "Job Title",
               "Narrative", "Hours", "Rate", "Paid", "Status"], 1)
    rr = 2
    for r in sorted(rows, key=lambda x: (x["full_name"], x["work_date"])):
        rate = r.get("rate", "Normal")
        vals = [r.get("full_name", ""), r.get("job_role", ""), r.get("work_date", ""),
                r.get("job_ref", ""), r.get("job_title", ""), r.get("narrative", ""),
                r.get("hours", 0), rate, paid_hours(r.get("hours", 0), rate),
                r.get("status", "")]
        for c, v in enumerate(vals, start=1):
            ws3.cell(row=rr, column=c, value=v)
        rr += 1
    for col, w in zip("ABCDEFGHIJ", [22, 14, 12, 14, 24, 36, 8, 8, 8, 12]):
        ws3.column_dimensions[col].width = w

    wb.save(path)
    return path


# ===========================================================================
# Modern web UI (rendered in a desktop window via pywebview).
# Reuses ALL the data layer above. Offline, single file.
# ===========================================================================
def _prefs_path():
    base = _local_dir()
    return os.path.join(base, "prefs.json")


def _local_dir():
    """Per-PC scratch (prefs, browser profile, logs) in LocalAppData."""
    base = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(),
                        APP_NAME)
    try:
        os.makedirs(base, exist_ok=True)
    except Exception:
        pass
    return base


def _load_prefs():
    try:
        with open(_prefs_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_prefs(p):
    try:
        with open(_prefs_path(), "w", encoding="utf-8") as f:
            json.dump(p, f)
    except Exception:
        pass


def _backup_db(path, delay=90):
    """Daily safety backup into a 'backups' subfolder (keep newest 30).
    Runs in a background thread, delayed so it NEVER competes with the app
    opening/logging in for the network file, and only once per day."""
    try:
        time.sleep(delay)
        if not os.path.exists(path) or os.path.getsize(path) < 1024:
            return
        d = os.path.join(os.path.dirname(path) or DATA_DIR, "backups")
        os.makedirs(d, exist_ok=True)
        today = f"{dt.datetime.now():%Y%m%d}"
        if glob.glob(os.path.join(d, f"timesheets_{today}_*.db")):
            return          # already backed up today
        dest = os.path.join(d, f"timesheets_{today}_{dt.datetime.now():%H%M%S}.db")
        shutil.copy2(path, dest)
        for old in sorted(glob.glob(os.path.join(d, "timesheets_*.db")))[:-30]:
            try:
                os.remove(old)
            except Exception:
                pass
    except Exception:
        pass


def open_db(cfg):
    """Open the configured database. NEVER silently fall back to a blank local
    database - doing so hides the real (e.g. network) database and looks like
    total data loss. Instead we retry, and raise a clear error if it truly
    cannot be opened, so the real file is left untouched."""
    path = cfg.get("db_path") or os.path.join(DATA_DIR, "timesheets.db")
    try:
        os.makedirs(os.path.dirname(path) or DATA_DIR, exist_ok=True)
    except Exception:
        pass
    last = None
    for attempt in range(5):
        try:
            db = DB(path)
            db.seed_defaults()
            # Back up in the background - never block launch on a network copy.
            threading.Thread(target=_backup_db, args=(path,), daemon=True).start()
            return db
        except Exception as e:
            last = e
            time.sleep(1.5)
    raise RuntimeError(
        "Could not open the timesheet database:\n\n" + str(path) + "\n\n"
        + str(last) + "\n\nIf this is on a shared or network drive, make sure it "
        "is connected and try again. Your data is safe in that file - the "
        "app will not create a blank one.")


def _fmt_week(ws):
    end = ws + dt.timedelta(days=6)
    return f"{ws.strftime('%d %b %Y')} – {end.strftime('%d %b %Y')}"


def job_cost_breakdown(db, job_id):
    with db.conn() as con:
        j = con.execute("SELECT job_ref, job_title, quote_value FROM job_refs WHERE id=?",
                        (job_id,)).fetchone()
        rates = {r["user_id"]: r["rate"] for r in con.execute(
            "SELECT user_id, rate FROM job_rates WHERE job_ref_id=?", (job_id,)).fetchall()}
        defaults = {r["id"]: (r["default_rate"] or 0.0) for r in con.execute(
            "SELECT id, default_rate FROM users").fetchall()}
        # Costed on the ORIGINAL submitted hours, matching job_cost_summary()
        # so this PDF can never disagree with the Job Costs screen.
        snap_rows = con.execute(
            "SELECT t.user_id uid, u.full_name nm, t.submitted_snapshot snap "
            "FROM timesheets t JOIN users u ON u.id=t.user_id "
            "WHERE t.submitted_snapshot != '' "
            "AND t.status IN ('Submitted','Approved','Locked')").fetchall()
        lines = []
        for sr in snap_rows:
            try:
                for sl in json.loads(sr["snap"]):
                    if sl.get("job_ref_id") != job_id:
                        continue
                    lines.append({"uid": sr["uid"], "nm": sr["nm"],
                                  "hours": sl.get("hours", 0) or 0,
                                  "rate": sl.get("rate") or "Normal",
                                  "rate_at_submit": sl.get("rate_at_submit")})
            except Exception:
                continue
        vrow = con.execute("SELECT SUM(value) v FROM job_variations "
                           "WHERE job_ref_id=?", (job_id,)).fetchone()
        irow = con.execute(
            "SELECT SUM(value) tot, "
            "SUM(CASE WHEN status='Paid' THEN value ELSE 0 END) paid "
            "FROM job_invoices WHERE job_ref_id=?", (job_id,)).fetchone()
    variations = (vrow["v"] if vrow and vrow["v"] else 0.0) or 0.0
    invoiced = (irow["tot"] if irow and irow["tot"] else 0.0) or 0.0
    invoiced_paid = (irow["paid"] if irow and irow["paid"] else 0.0) or 0.0
    # Cost each LINE at the rate frozen when it was submitted (falling back to
    # today's rate for older work), then roll up per person - matching
    # job_cost_summary() exactly.
    agg = {}
    for l in lines:
        uid = l["uid"]
        d = agg.setdefault(uid, {"name": l["nm"], "hours": 0.0, "paid": 0.0,
                                 "cost": 0.0, "rates": set()})
        paid = float(l["hours"]) * RATE_MULT.get(l["rate"], 1.0)
        frozen = l.get("rate_at_submit")
        rate = float(frozen) if frozen else (
            rates[uid] if uid in rates else defaults.get(uid, 0.0))
        d["hours"] += float(l["hours"])
        d["paid"] += paid
        d["cost"] += paid * rate
        d["rates"].add(round(rate, 2))
    rows, tot_cost, tot_hours = [], 0.0, 0.0
    for uid, d in agg.items():
        cost = round(d["cost"], 2)
        tot_cost += cost
        tot_hours += d["hours"]
        # one rate -> show it; several (a rate change mid-job) -> effective rate
        rate = (list(d["rates"])[0] if len(d["rates"]) == 1
                else (round(d["cost"] / d["paid"], 2) if d["paid"] else 0.0))
        rows.append({"name": d["name"], "hours": round(d["hours"], 2),
                     "paid": round(d["paid"], 2), "rate": rate, "cost": cost})
    rows.sort(key=lambda x: -x["cost"])
    quote_orig = j["quote_value"] or 0
    quote = quote_orig + variations      # current contract value incl. variations
    return {"job_ref": j["job_ref"], "job_title": j["job_title"], "quote": quote,
            "quote_original": round(quote_orig, 2),
            "variations": round(variations, 2),
            "invoiced": round(invoiced, 2),
            "invoiced_paid": round(invoiced_paid, 2),
            "to_invoice": round(quote - invoiced, 2),
            "total_hours": round(tot_hours, 2), "total_cost": round(tot_cost, 2),
            "remaining": round(quote - tot_cost, 2),
            "pct": (round(tot_cost / quote * 100, 1) if quote else 0), "rows": rows}


def _ascii(s):
    return str(s).encode("latin-1", "replace").decode("latin-1")


def export_job_cost_pdf(path, company, d, generated_by):
    from fpdf import FPDF
    GREEN = (132, 195, 65); DARK = (28, 28, 36); MUT = (120, 128, 137)
    LINE = (224, 228, 234)
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(True, margin=16); pdf.add_page()
    W = pdf.w
    pdf.set_fill_color(*DARK); pdf.rect(0, 0, W, 26, 'F')
    pdf.set_fill_color(*GREEN); pdf.rect(0, 26, W, 1.6, 'F')
    pdf.set_xy(12, 6); pdf.set_text_color(*GREEN); pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 8, _ascii(company.upper()), ln=1)
    pdf.set_x(12); pdf.set_text_color(228, 228, 233); pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 5, _ascii(f"Job Cost Report  -  {d['job_ref']}  {d['job_title']}"), ln=1)
    pdf.set_text_color(*MUT); pdf.set_font("Helvetica", "", 9); pdf.set_xy(12, 32)
    pdf.cell(0, 5, f"Generated {dt.datetime.now():%Y-%m-%d %H:%M} by {_ascii(generated_by)}", ln=1)
    cards = [("QUOTE", f"GBP {d['quote']:.0f}" if d["quote"] else "-"),
             ("TOTAL HOURS", f"{d['total_hours']:g}"),
             ("TOTAL COST", f"GBP {d['total_cost']:.0f}"),
             ("REMAINING", f"GBP {d['remaining']:.0f}" if d["quote"] else "-"),
             ("% USED", f"{d['pct']:g}%" if d["quote"] else "-")]
    cw = 35; x = 12; y = 40
    for lab, val in cards:
        pdf.set_fill_color(246, 248, 250); pdf.set_draw_color(*LINE); pdf.rect(x, y, cw, 18, 'DF')
        pdf.set_xy(x + 3, y + 3); pdf.set_text_color(*MUT); pdf.set_font("Helvetica", "", 7)
        pdf.cell(cw - 4, 4, lab)
        pdf.set_xy(x + 3, y + 8); pdf.set_text_color(38, 38, 46); pdf.set_font("Helvetica", "B", 13)
        pdf.cell(cw - 4, 7, val)
        x += cw + 2.5
    pdf.set_y(y + 24)
    if d["quote"]:
        pct = min(100, d["pct"])
        col = (210, 58, 58) if d["total_cost"] > d["quote"] else ((230, 165, 0) if d["pct"] >= 80 else GREEN)
        pdf.set_x(12); pdf.set_font("Helvetica", "B", 10); pdf.set_text_color(38, 38, 46)
        pdf.cell(0, 6, "Budget used", ln=1)
        yy = pdf.get_y(); pdf.set_fill_color(238, 241, 245); pdf.rect(12, yy, W - 24, 7, 'F')
        pdf.set_fill_color(*col); pdf.rect(12, yy, (W - 24) * pct / 100, 7, 'F'); pdf.set_y(yy + 11)
    rows = d["rows"]
    if rows and d["total_cost"] > 0:
        pdf.set_x(12); pdf.set_font("Helvetica", "B", 11); pdf.set_text_color(38, 38, 46)
        pdf.cell(0, 7, "Cost by employee", ln=1)
        maxv = max(r["cost"] for r in rows) or 1; bar_max = W - 24 - 60 - 26
        for r in rows[:10]:
            yy = pdf.get_y(); pdf.set_x(12); pdf.set_font("Helvetica", "", 9)
            pdf.set_text_color(70, 70, 78); pdf.cell(58, 6, _ascii(str(r["name"])[:30]))
            bw = bar_max * (r["cost"] / maxv); pdf.set_fill_color(*GREEN)
            pdf.rect(72, yy + 1, max(0.6, bw), 4.6, 'F')
            pdf.set_xy(72 + bw + 2, yy); pdf.set_text_color(*MUT); pdf.set_font("Helvetica", "", 8)
            pdf.cell(26, 6, f"GBP {r['cost']:.0f}"); pdf.ln(7)
        pdf.ln(2)
    pdf.set_font("Helvetica", "B", 9); pdf.set_fill_color(*DARK); pdf.set_text_color(255, 255, 255)
    widths = [74, 28, 28, 28, 32]
    for w, htxt in zip(widths, ["Employee", "Hours", "Paid hrs", "Rate (GBP)", "Cost (GBP)"]):
        pdf.cell(w, 7, htxt, border=0, fill=True)
    pdf.ln(); pdf.set_text_color(32, 32, 38); pdf.set_font("Helvetica", "", 9); alt = False
    for r in rows:
        pdf.set_fill_color(244, 247, 250) if alt else pdf.set_fill_color(255, 255, 255)
        pdf.cell(widths[0], 6, _ascii(r["name"]), fill=True)
        pdf.cell(widths[1], 6, f"{r['hours']:g}", fill=True, align="R")
        pdf.cell(widths[2], 6, f"{r['paid']:g}", fill=True, align="R")
        pdf.cell(widths[3], 6, f"{r['rate']:.2f}", fill=True, align="R")
        pdf.cell(widths[4], 6, f"{r['cost']:.2f}", fill=True, align="R")
        pdf.ln(); alt = not alt
    if not rows:
        pdf.set_fill_color(255, 255, 255)
        pdf.cell(sum(widths), 6, "No hours booked to this job yet.", fill=True); pdf.ln()
    pdf.set_font("Helvetica", "B", 10); pdf.set_fill_color(234, 240, 230)
    pdf.cell(widths[0], 7, "TOTAL", fill=True)
    pdf.cell(widths[1], 7, f"{d['total_hours']:g}", fill=True, align="R")
    pdf.cell(widths[2], 7, "", fill=True); pdf.cell(widths[3], 7, "", fill=True)
    pdf.cell(widths[4], 7, f"{d['total_cost']:.2f}", fill=True, align="R"); pdf.ln()
    pdf.output(path)
    return path

def export_report_pdf(path, company, title, period, generated_by, rows,
                      group_key, group_label, include_cost=False):
    from fpdf import FPDF
    GREEN = (132, 195, 65); DARK = (28, 28, 36); MUT = (120, 128, 137)
    LINE = (224, 228, 234)
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(True, margin=16)
    pdf.add_page()
    W = pdf.w
    pdf.set_fill_color(*DARK); pdf.rect(0, 0, W, 26, 'F')
    pdf.set_fill_color(*GREEN); pdf.rect(0, 26, W, 1.6, 'F')
    pdf.set_xy(12, 6); pdf.set_text_color(*GREEN); pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 8, _ascii(company.upper()), ln=1)
    pdf.set_x(12); pdf.set_text_color(228, 228, 233); pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 5, _ascii(title), ln=1)
    pdf.set_text_color(*MUT); pdf.set_font("Helvetica", "", 9); pdf.set_xy(12, 32)
    pdf.cell(0, 5, _ascii(f"{period}     |     Generated "
             f"{dt.datetime.now():%Y-%m-%d %H:%M} by {generated_by}"), ln=1)
    total = round(sum(r.get("hours", 0) for r in rows), 2)
    people = len({r.get("full_name") for r in rows})
    njobs = len({r.get("job_ref") for r in rows if r.get("job_ref")})
    cards = [("TOTAL HOURS", f"{total:g}"), ("ENTRIES", str(len(rows))),
             ("PEOPLE", str(people)), ("JOBS", str(njobs))]
    x, y = 12, 40
    for lab, val in cards:
        pdf.set_fill_color(246, 248, 250); pdf.set_draw_color(*LINE)
        pdf.rect(x, y, 60, 18, 'DF')
        pdf.set_xy(x + 4, y + 3); pdf.set_text_color(*MUT)
        pdf.set_font("Helvetica", "", 7.5); pdf.cell(52, 4, lab)
        pdf.set_xy(x + 4, y + 8); pdf.set_text_color(38, 38, 46)
        pdf.set_font("Helvetica", "B", 15); pdf.cell(52, 7, val)
        x += 66
    pdf.set_y(y + 24)
    agg = {}
    for r in rows:
        k = r.get(group_key) or "(none)"
        agg[k] = round(agg.get(k, 0) + r.get("hours", 0), 2)
    items = sorted(agg.items(), key=lambda z: -z[1])[:10]
    if total > 0 and items:
        pdf.set_text_color(38, 38, 46); pdf.set_font("Helvetica", "B", 11); pdf.set_x(12)
        pdf.cell(0, 7, _ascii(f"Hours by {group_label}"), ln=1)
        maxv = max(v for _, v in items) or 1
        bar_max = W - 24 - 70 - 22
        for k, v in items:
            yy = pdf.get_y()
            pdf.set_x(12); pdf.set_font("Helvetica", "", 9); pdf.set_text_color(70, 70, 78)
            pdf.cell(68, 6, _ascii(str(k)[:34]))
            bw = bar_max * (v / maxv)
            pdf.set_fill_color(*GREEN); pdf.rect(82, yy + 1, max(0.6, bw), 4.6, 'F')
            pdf.set_xy(82 + bw + 2, yy); pdf.set_text_color(*MUT)
            pdf.set_font("Helvetica", "", 8); pdf.cell(20, 6, f"{v:g}h")
            pdf.ln(7)
        pdf.ln(2)
    if include_cost:
        cols = [("full_name", "Name", 40), ("job_role", "Role", 26),
                ("work_date", "Date", 20), ("job_ref", "Job Ref", 22),
                ("job_title", "Job Title", 38), ("narrative", "Narrative", 56),
                ("hours", "Hours", 14), ("cost", "Cost (GBP)", 20),
                ("status", "Status", 22)]
    else:
        cols = [("full_name", "Name", 40), ("job_role", "Role", 26),
                ("work_date", "Date", 22), ("job_ref", "Job Ref", 24),
                ("job_title", "Job Title", 42), ("narrative", "Narrative", 66),
                ("hours", "Hours", 14), ("status", "Status", 22)]
    pdf.set_font("Helvetica", "B", 8); pdf.set_fill_color(*DARK)
    pdf.set_text_color(255, 255, 255)
    for k, htxt, w in cols:
        pdf.cell(w, 7, _ascii(htxt), border=0, fill=True)
    pdf.ln(); pdf.set_text_color(32, 32, 38); pdf.set_font("Helvetica", "", 8); alt = False
    for r in rows:
        if pdf.get_y() > pdf.h - 18:
            pdf.add_page()
        pdf.set_fill_color(244, 247, 250) if alt else pdf.set_fill_color(255, 255, 255)
        for k, htxt, w in cols:
            if k == "cost":
                val = f"{r.get('cost', 0):,.2f}"
            else:
                val = str(r.get(k, ""))
                if k == "narrative" and len(val) > 46:
                    val = val[:44] + ".."
            pdf.cell(w, 6, _ascii(val), border=0, fill=True,
                     align=("R" if k in ("hours", "cost") else "L"))
        pdf.ln(); alt = not alt
    pdf.set_font("Helvetica", "B", 9); pdf.set_fill_color(234, 240, 230)
    hours_idx = next(i for i, c in enumerate(cols) if c[0] == "hours")
    pdf.cell(sum(w for _, _, w in cols[:hours_idx]), 7, "TOTAL", fill=True)
    pdf.cell(cols[hours_idx][2], 7, f"{total:g}", fill=True, align="R")
    if include_cost:
        cost_w = next(w for k, _, w in cols if k == "cost")
        total_cost = round(sum(r.get("cost", 0) for r in rows), 2)
        pdf.cell(cost_w, 7, f"{total_cost:,.2f}", fill=True, align="R")
    pdf.cell(cols[-1][2], 7, "", fill=True); pdf.ln()
    pdf.output(path)
    return path


class Api:
    def __init__(self):
        self.cfg = load_config()
        self.db = open_db(self.cfg)
        self.user = None

    # ---- session ----
    def login(self, email, password):
        u = self.db.authenticate(email, password)
        if not u:
            return None
        u = dict(u)
        u.pop("password_hash", None)
        u.pop("salt", None)
        self.user = u
        return u

    def startup_prefs(self):
        return {"username": _load_prefs().get("username", "")}

    def remember_login(self, remember, username=""):
        """Called after a successful login. Saves the username for pre-fill,
        and if 'remember' is ticked, stores a stay-signed-in token: the raw
        secret lives only in this PC's local prefs file, its SHA-256 lives in
        the database. Unticked clears any existing token."""
        u = self.user
        if not u:
            return {"ok": False}
        p = _load_prefs()
        if username:
            p["username"] = str(username)
        if remember:
            tok = secrets.token_hex(32)
            digest = hashlib.sha256(tok.encode()).hexdigest()
            with self.db.conn() as con:
                con.execute("UPDATE users SET auth_token=? WHERE id=?",
                            (digest, u["id"]))
            p["token"] = tok
            p["uid"] = u["id"]
        else:
            with self.db.conn() as con:
                con.execute("UPDATE users SET auth_token='' WHERE id=?",
                            (u["id"],))
            p.pop("token", None)
            p.pop("uid", None)
        _save_prefs(p)
        return {"ok": True}

    def auto_login(self):
        """Silent sign-in at startup using the stay-signed-in token."""
        p = _load_prefs()
        tok, uid = p.get("token"), p.get("uid")
        if not tok or not uid:
            return None
        row = self.db.get_user(int(uid))
        if not row or not row["active"] or not row["auth_token"]:
            return None
        if hashlib.sha256(tok.encode()).hexdigest() != row["auth_token"]:
            return None
        u = dict(row)
        u.pop("password_hash", None)
        u.pop("salt", None)
        u.pop("auth_token", None)
        self.user = u
        try:
            self.db.audit(u["email"], "LOGIN", "", "Automatic sign-in (remembered)")
        except Exception:
            pass   # a slow/locked network write must never drop the sign-in
        return u

    def forget_login(self):
        """Explicit Sign out: clear the stay-signed-in token everywhere."""
        if self.user:
            with self.db.conn() as con:
                con.execute("UPDATE users SET auth_token='' WHERE id=?",
                            (self.user["id"],))
        p = _load_prefs()
        p.pop("token", None)
        p.pop("uid", None)
        _save_prefs(p)
        return {"ok": True}

    def logo(self):
        if not hasattr(self, "_logo_cache"):
            self._logo_cache = _logo_data_uri(self.cfg)
        return self._logo_cache

    def brand(self):
        """What the login screen and sidebar show: company name + optional logo."""
        return {"company": (self.cfg.get("company_name") or "").strip(),
                "app": APP_NAME, "version": APP_VERSION,
                "logo": self.logo(), "needs_setup": self.db.needs_setup()}

    # ---- first run ----
    def needs_setup(self):
        return self.db.needs_setup()

    def pick_data_folder(self):
        """Let the person choose where the shared data folder lives (setup and
        Settings). Returns the folder, and whether it already holds a database."""
        import webview
        res = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        if not res:
            return {"ok": False}
        folder = res if isinstance(res, str) else res[0]
        has_db = os.path.exists(os.path.join(folder, "timesheets.db"))
        return {"ok": True, "folder": folder, "has_db": has_db}

    def complete_setup(self, company, full_name, email, password, data_folder=""):
        """First-run wizard. Creates the admin account and saves the company.
        If data_folder is given the database lives there (shared), otherwise in
        this PC's AppData."""
        company = (company or "").strip()
        full_name = (full_name or "").strip()
        email = (email or "").strip()
        if not full_name or not email or not password:
            return {"ok": False, "error": "Name, email and password are all needed."}
        if len(password) < 6:
            return {"ok": False, "error": "Use a password of at least 6 characters."}
        folder = (data_folder or "").strip()
        if folder:
            try:
                os.makedirs(folder, exist_ok=True)
            except Exception as e:
                return {"ok": False, "error": f"Cannot use that folder: {e}"}
            new_path = os.path.join(folder, "timesheets.db")
            if os.path.abspath(new_path) != os.path.abspath(self.cfg.get("db_path", "")):
                self.cfg["db_path"] = new_path
                save_config(self.cfg)
                self.db = open_db(self.cfg)
        if not self.db.needs_setup():
            # A shared folder that already has accounts: nothing to create.
            self.cfg["company_name"] = company or self.cfg.get("company_name", "")
            save_config(self.cfg)
            self._logo_cache = _logo_data_uri(self.cfg)
            return {"ok": True, "existing": True}
        self.cfg["company_name"] = company
        save_config(self.cfg)
        try:
            self.db.set_meta("company_name", company)
            self.db.create_first_admin(full_name, email, password)
        except Exception as e:
            return {"ok": False, "error": str(e)}
        self._logo_cache = _logo_data_uri(self.cfg)
        return {"ok": True}

    # ---- settings (admin) ----
    def get_goals(self):
        self._need()
        g = dict(DEFAULT_GOALS); g.update(self.cfg.get("goals") or {})
        return g

    def save_goals(self, vals):
        u = self._admin()
        g = dict(DEFAULT_GOALS); g.update(self.cfg.get("goals") or {})
        for k in DEFAULT_GOALS:
            if k in (vals or {}):
                try:
                    g[k] = max(0.0, float(vals[k]))
                except (TypeError, ValueError):
                    return {"ok": False, "error": "Goals must be numbers."}
        self.cfg["goals"] = g
        save_config(self.cfg)
        self.db.audit(u["email"], "GOALS", "", json.dumps(g))
        return {"ok": True, "goals": g}

    def get_settings(self):
        u = self._need()
        if u["user_type"] != "Admin":
            raise PermissionError("Admins only")
        c = self.cfg
        return {"company_name": c.get("company_name", ""),
                "weekly_target_hours": float(c.get("weekly_target_hours", 37.5)),
                "daily_hours_warning": float(c.get("daily_hours_warning", 10)),
                "inactivity_timeout_minutes": int(c.get("inactivity_timeout_minutes", 15)),
                "start_maximized": bool(c.get("start_maximized", True)),
                "db_path": c.get("db_path", ""),
                "data_folder": os.path.dirname(c.get("db_path", "")) or DATA_DIR,
                "has_logo": bool(self.logo()),
                "version": APP_VERSION}

    def save_settings(self, vals):
        u = self._need()
        if u["user_type"] != "Admin":
            raise PermissionError("Admins only")
        vals = vals or {}
        old = dict(self.cfg)
        self.cfg["company_name"] = str(vals.get("company_name", "")).strip()
        try:
            self.db.set_meta("company_name", self.cfg["company_name"])
        except Exception:
            pass
        try:
            self.cfg["weekly_target_hours"] = max(1.0, float(vals.get("weekly_target_hours", 37.5)))
            self.cfg["daily_hours_warning"] = max(1.0, float(vals.get("daily_hours_warning", 10)))
            self.cfg["inactivity_timeout_minutes"] = max(1, int(vals.get("inactivity_timeout_minutes", 15)))
        except (TypeError, ValueError):
            return {"ok": False, "error": "Hours and minutes must be numbers."}
        self.cfg["start_maximized"] = bool(vals.get("start_maximized", True))
        save_config(self.cfg)
        self.db.audit(u["email"], "SETTINGS", json.dumps({k: old.get(k) for k in
                      ("company_name", "weekly_target_hours", "daily_hours_warning",
                       "inactivity_timeout_minutes")}),
                      json.dumps({k: self.cfg.get(k) for k in
                      ("company_name", "weekly_target_hours", "daily_hours_warning",
                       "inactivity_timeout_minutes")}))
        return {"ok": True}

    def pick_logo(self):
        """Choose a PNG; it is copied into the data folder as logo.png."""
        u = self._need()
        if u["user_type"] != "Admin":
            raise PermissionError("Admins only")
        import webview
        res = self.window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=("PNG image (*.png)",))
        if not res:
            return {"ok": False}
        src = res if isinstance(res, str) else res[0]
        try:
            if os.path.getsize(src) > 700_000:
                return {"ok": False, "error": "Logo must be under 700 KB. Resize it and try again."}
            shutil.copy2(src, os.path.join(DATA_DIR, "logo.png"))
        except Exception as e:
            return {"ok": False, "error": str(e)}
        self._logo_cache = _logo_data_uri(self.cfg)
        self.db.audit(u["email"], "SETTINGS", "", "Logo changed")
        return {"ok": True, "logo": self._logo_cache}

    def remove_logo(self):
        u = self._need()
        if u["user_type"] != "Admin":
            raise PermissionError("Admins only")
        try:
            os.remove(os.path.join(DATA_DIR, "logo.png"))
        except FileNotFoundError:
            pass
        except Exception as e:
            return {"ok": False, "error": str(e)}
        self._logo_cache = ""
        return {"ok": True}

    def change_data_folder(self, folder):
        """Move the database (and Receipts) to a new folder - typically a shared
        drive so the whole team works on one database. If the folder already
        holds a timesheets.db, that one is used instead (joining an existing
        team setup) and nothing is copied."""
        u = self._need()
        if u["user_type"] != "Admin":
            raise PermissionError("Admins only")
        folder = (folder or "").strip()
        if not folder:
            return {"ok": False, "error": "Choose a folder first."}
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            return {"ok": False, "error": f"Cannot use that folder: {e}"}
        old_path = self.cfg.get("db_path", "")
        new_path = os.path.join(folder, "timesheets.db")
        if os.path.abspath(new_path) == os.path.abspath(old_path):
            return {"ok": True, "unchanged": True}
        joined = os.path.exists(new_path)
        try:
            if not joined:
                shutil.copy2(old_path, new_path)
                old_rec = os.path.join(os.path.dirname(old_path), "Receipts")
                new_rec = os.path.join(folder, "Receipts")
                if os.path.isdir(old_rec) and not os.path.exists(new_rec):
                    shutil.copytree(old_rec, new_rec)
            self.cfg["db_path"] = new_path
            save_config(self.cfg)
            self.db = open_db(self.cfg)
        except Exception as e:
            return {"ok": False, "error": str(e)}
        self.db.audit(u["email"], "SETTINGS", old_path, new_path)
        return {"ok": True, "joined": joined, "db_path": new_path,
                "needs_login": joined}

    def open_data_folder(self):
        self._need()
        folder = os.path.dirname(self.cfg.get("db_path", "")) or DATA_DIR
        try:
            os.startfile(folder)
        except Exception as e:
            return {"ok": False, "error": str(e)}
        return {"ok": True}

    def logout(self):
        if self.user:
            self.db.audit(self.user["email"], "LOGOUT")
        self.user = None
        return True

    def session(self):
        return self.user

    def _need(self):
        if not self.user:
            raise PermissionError("Not signed in")
        return self.user

    # ---- dashboard ----
    def dashboard(self):
        u = self._need()
        today = dt.date.today()
        wk = monday_of(today)
        wk_end = wk + dt.timedelta(days=6)
        m_start = today.replace(day=1)
        m_end = (dt.date(today.year, 12, 31) if today.month == 12
                 else dt.date(today.year, today.month + 1, 1) - dt.timedelta(days=1))

        dayh = self.db.user_hours_by_date(u["id"])   # 1 query

        def my_hours(a, b):
            tot = 0.0
            for ds, h in dayh.items():
                try:
                    d = dt.date.fromisoformat(ds)
                except Exception:
                    continue
                if a <= d <= b:
                    tot += h
            return round(tot, 2)

        ts_list = self.db.list_timesheets(user_id=u["id"])   # 1 query
        hbt = self.db.hours_by_timesheet()                   # 1 query

        def ts_hours(t):
            if t.get("status") in ("Submitted", "Approved", "Locked"):
                snap = t.get("submitted_snapshot") or ""
                if snap:
                    try:
                        return round(sum(float(l.get("hours", 0))
                                         for l in json.loads(snap)), 2)
                    except Exception:
                        pass
            return hbt.get(t["id"], 0.0)

        prev = [{"week": parse_iso(t["week_start"]), "status": t["status"],
                 "hours": ts_hours(t)} for t in ts_list
                if hbt.get(t["id"], 0.0) > 0 or t["status"] != "Draft"]

        weeks = []
        for i in range(7, -1, -1):
            w = wk - dt.timedelta(days=7 * i)
            weeks.append({"label": w.strftime("%d %b"),
                          "hours": my_hours(w, w + dt.timedelta(days=6))})

        ts = self.db.get_or_create_timesheet(u["id"], wk)
        last_wk = wk - dt.timedelta(days=7)
        last_ts = next((t for t in ts_list if t["week_start"] == last_wk.isoformat()),
                       None)
        out = {"user": u, "week": _fmt_week(wk), "status": ts["status"],
               "first_name": (u["full_name"] or "").split(" ")[0],
               "my_week": my_hours(wk, wk_end), "my_month": my_hours(m_start, m_end),
               "target": float(self.cfg.get("weekly_target_hours", 37.5)),
               "prev": prev[:14], "weeks": weeks,
               "is_mgr": u["user_type"] in ("Manager", "Admin"),
               "is_admin": u["user_type"] == "Admin",
               # things that need doing, for the attention cards
               "rejected": ts["status"] == "Rejected",
               "reject_reason": ts.get("reject_reason", ""),
               "last_week_status": (last_ts["status"] if last_ts else "None"),
               "last_week_label": last_wk.strftime("%d %b")}
        if out["is_mgr"]:
            rep = Reports(self.db)
            out["co_week"] = rep.person_report(date_from=wk, date_to=wk_end)[1]
            out["co_month"] = rep.person_report(date_from=m_start, date_to=m_end)[1]
            out["active_staff"] = len(self.db.list_users(include_inactive=False))
            waiting = self.db.list_timesheets(status="Submitted")
            out["awaiting"] = len(waiting)
            # how long the oldest one has been sitting there
            old = 0
            for t in waiting:
                try:
                    d = (today - dt.date.fromisoformat(t["week_start"])).days
                    old = max(old, d)
                except Exception:
                    pass
            out["awaiting_oldest_days"] = old
            miss = rep.missing_timesheets(last_wk)
            out["missing"] = len(miss)
            out["missing_names"] = [m["full_name"] for m in miss][:6]
            br = rep.summary_by("job_role", date_from=m_start, date_to=m_end)
            out["by_role"] = [{"label": k, "value": v} for k, v in
                              sorted(br.items(), key=lambda x: -x[1])]
        if out["is_admin"]:
            # Money: cheap aggregates only - no timesheet snapshots parsed here,
            # so signing in stays fast. Job cost/profit is loaded separately.
            with self.db.conn() as con:
                q = con.execute("SELECT COALESCE(SUM(quote_value),0) v FROM job_refs "
                                "WHERE status='Active'").fetchone()["v"] or 0.0
                var = con.execute(
                    "SELECT COALESCE(SUM(v.value),0) s FROM job_variations v "
                    "JOIN job_refs j ON j.id=v.job_ref_id "
                    "WHERE j.status='Active'").fetchone()["s"] or 0.0
                inv = con.execute(
                    "SELECT COALESCE(SUM(i.value),0) s FROM job_invoices i "
                    "JOIN job_refs j ON j.id=i.job_ref_id "
                    "WHERE j.status='Active'").fetchone()["s"] or 0.0
                inv_month = con.execute(
                    "SELECT COALESCE(SUM(value),0) s FROM job_invoices "
                    "WHERE inv_date>=? AND inv_date<=?",
                    (m_start.isoformat(), m_end.isoformat())).fetchone()["s"] or 0.0
                unpaid = con.execute(
                    "SELECT COALESCE(SUM(value),0) s FROM job_invoices "
                    "WHERE status!='Paid'").fetchone()["s"] or 0.0
                bills_unpaid = con.execute(
                    "SELECT COALESCE(SUM(value),0) s FROM job_bills "
                    "WHERE status!='Paid'").fetchone()["s"] or 0.0
            out["money"] = {"invoiced_month": round(inv_month, 2),
                            "to_invoice": round(q + var - inv, 2),
                            "invoiced_unpaid": round(unpaid, 2),
                            "bills_unpaid": round(bills_unpaid, 2)}
        return out

    def dash_jobs_at_risk(self, limit=6):
        """Jobs closest to (or past) their quote. Loaded AFTER the dashboard
        renders because it has to price every submitted timesheet, which is the
        one genuinely slow query here - signing in must not wait for it."""
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"rows": [], "over_budget": 0, "near": 0}
        rows = [d for d in self.db.job_cost_summary()
                if d["kind"] == "Project" and d["quote_current"] > 0
                and d["total_cost"] > 0]
        rows.sort(key=lambda d: -d["pct"])
        return {"rows": [{"id": d["id"], "job_ref": d["job_ref"],
                          "job_title": d["job_title"], "pct": d["pct"],
                          "over": d["pct"] > 100} for d in rows[:int(limit)]],
                "over_budget": sum(1 for d in rows if d["pct"] > 100),
                "near": sum(1 for d in rows if 85 <= d["pct"] <= 100)}

    # ---- my timesheet (interactive) ----
    def _own_ts(self, ts_id):
        return any(t["id"] == ts_id for t in
                   self.db.list_timesheets(user_id=self.user["id"]))

    def timesheet(self, week_start=None):
        u = self._need()
        wk = (dt.date.fromisoformat(week_start) if week_start
              else monday_of(dt.date.today()))
        ts = self.db.get_or_create_timesheet(u["id"], wk)
        editable = ts["status"] in ("Draft", "Rejected")
        lines = self.db.get_lines(ts["id"]) if editable else self.db.employee_lines(ts)
        worked = round(sum(l["hours"] for l in lines), 2)
        paid = round(sum(paid_hours(l["hours"], l.get("rate", "Normal")) for l in lines), 2)
        return {"id": ts["id"], "week_start": wk.isoformat(), "week": _fmt_week(wk),
                "status": ts["status"], "editable": editable,
                "reject_reason": ts.get("reject_reason", ""),
                "days": [(wk + dt.timedelta(days=i)).isoformat() for i in range(7)],
                "day_labels": [(wk + dt.timedelta(days=i)).strftime("%A %d %b")
                               for i in range(7)],
                "lines": [{"id": l.get("id"), "work_date": l["work_date"],
                           "job_ref_id": l["job_ref_id"], "job_ref": l["job_ref"],
                           "job_role": l.get("job_role", ""), "rate": l.get("rate", "Normal"),
                           "narrative": l["narrative"], "hours": l["hours"]} for l in lines],
                "worked": worked, "paid": paid}

    def timesheet_jobs(self):
        self._need()
        # Sorted by job reference, exactly like the Job References screen.
        # Non-project entries (Holiday, Sick Leave etc) used to be forced to
        # the bottom of the list and shown by title alone, so someone looking
        # for "26-ATX-001" could not find it and assumed it was missing. They
        # now sit in reference order with everything else, still tagged
        # "(non-project)" so they are easy to tell apart.
        js = sorted(self.db.list_jobs(only_active=True),
                    key=lambda j: str(j["job_ref"]))
        return [{"id": j["id"],
                 "label": (f"{j['job_ref']} - {j['job_title']}"
                           if j.get("kind", "Project") == "Project"
                           else f"{j['job_ref']} - {j['job_title']} (non-project)")}
                for j in js]

    def timesheet_roles(self):
        self._need()
        return [r["name"] for r in self.db.list_roles(only_enabled=True)]

    def save_timesheet(self, ts_id, lines):
        u = self._need()
        if not self._own_ts(ts_id):
            return {"ok": False, "error": "Not your timesheet."}
        # Draft save: hours must be valid, but narrative may be left blank so
        # copied/pasted lines can be saved and their narratives filled in
        # later. Narratives are enforced on submit.
        for l in lines:
            try:
                h = float(l["hours"])
            except (TypeError, ValueError):
                return {"ok": False, "error": "Hours must be a number."}
            if h <= 0:
                return {"ok": False, "error": "Hours must be greater than zero."}
        self.db.save_timesheet_lines(u["email"], ts_id, lines)
        return {"ok": True}

    def timer_get(self):
        u = self._need()
        with self.db.conn() as con:
            r = con.execute("SELECT timer_start FROM users WHERE id=?",
                            (u["id"],)).fetchone()
        v = (r["timer_start"] if r else "") or ""
        try:
            return {"start_ms": int(v)} if v else {"start_ms": None}
        except ValueError:
            return {"start_ms": None}

    def timer_start(self, start_ms):
        u = self._need()
        with self.db.conn() as con:
            con.execute("UPDATE users SET timer_start=? WHERE id=?",
                        (str(int(start_ms)), u["id"]))
        return {"ok": True}

    def timer_clear(self):
        u = self._need()
        with self.db.conn() as con:
            con.execute("UPDATE users SET timer_start='' WHERE id=?", (u["id"],))
        return {"ok": True}

    def submit_my_timesheet(self, ts_id, lines):
        u = self._need()
        if not self._own_ts(ts_id):
            return {"ok": False, "error": "Not your timesheet."}
        if not lines:
            return {"ok": False, "error": "Cannot submit an empty timesheet."}
        for l in lines:
            ok, msg = validate_line(l["narrative"], l["hours"])
            if not ok:
                return {"ok": False, "error": msg}
        self.db.submit_timesheet(u["email"], ts_id, lines)
        return {"ok": True}

    # ---- read views ----
    def submissions(self):
        u = self._need()
        hbt = self.db.hours_by_timesheet()   # 1 query
        rows = []
        for t in self.db.list_timesheets(user_id=u["id"]):   # 1 query
            if hbt.get(t["id"], 0.0) == 0 and t["status"] == "Draft":
                continue   # hide only empty Draft placeholders
            snap = t.get("submitted_snapshot") or ""
            if t.get("status") in ("Submitted", "Approved", "Locked") and snap:
                try:
                    tot = round(sum(float(l.get("hours", 0))
                                    for l in json.loads(snap)), 2)
                except Exception:
                    tot = hbt.get(t["id"], 0.0)
            else:
                tot = hbt.get(t["id"], 0.0)
            rows.append({"week": parse_iso(t["week_start"]),
                         "week_start": t["week_start"], "status": t["status"],
                         "hours": tot, "reason": t.get("reject_reason", "")})
        return rows

    def _basis_lines(self, t, basis):
        """Lines for a timesheet row in the requested basis: the employee's
        frozen submission, or the live (manager-adjusted) lines."""
        if basis == "submitted":
            snap = t.get("submitted_snapshot") or ""
            if snap:
                return json.loads(snap)
        return self.db.get_lines(t["id"])

    def approvals(self, status="Submitted", basis="adjusted"):
        self._need()
        st = None if status == "All" else status
        hbt = self.db.hours_by_timesheet()   # 1 query
        out = []
        for t in self.db.list_timesheets(status=st):   # 1 query
            if hbt.get(t["id"], 0.0) == 0 and t["status"] == "Draft":
                continue   # hide only empty Draft placeholders; keep anything
                           # that's been through the workflow (Approved etc.)
            if basis == "submitted":
                snap = t.get("submitted_snapshot") or ""
                try:
                    tot = round(sum(float(l.get("hours", 0))
                                    for l in json.loads(snap)), 2) if snap else 0.0
                except Exception:
                    tot = 0.0
            else:
                tot = hbt.get(t["id"], 0.0)
            out.append({"id": t["id"], "name": t["full_name"], "role": t["job_role"],
                        "week": parse_iso(t["week_start"]),
                        # raw yyyy-mm-dd so the screen can sort by real date -
                        # sorting the display string would put '13 Jul' after
                        # '27 Jul' because it compares alphabetically.
                        "week_iso": t["week_start"],
                        "status": t["status"],
                        "hours": tot, "reason": t.get("reject_reason", "")})
        return out

    def approve(self, tid, status, reason=""):
        u = self._need()
        if status == "Rejected":
            self.db.reject_timesheet(u["email"], tid, reason or "")
        else:
            self.db.set_timesheet_status(u["email"], tid, status)
        return True

    def approval_lines(self, tid, basis="adjusted"):
        self._need()
        row = next((t for t in self.db.list_timesheets() if t["id"] == tid), None)
        if not row:
            return None
        wk = dt.date.fromisoformat(row["week_start"])
        snap_raw = row.get("submitted_snapshot") or ""
        has_snap = bool(snap_raw)
        if basis == "submitted" and has_snap:
            raw = json.loads(snap_raw)
            lines = [{"id": None, "work_date": l.get("work_date", ""),
                      "job_ref_id": l.get("job_ref_id"), "job_ref": l.get("job_ref", ""),
                      "job_role": l.get("job_role", ""), "rate": l.get("rate", "Normal"),
                      "narrative": l.get("narrative", ""), "hours": l.get("hours", 0)}
                     for l in raw]
            used = "submitted"
        else:
            lines = [{"id": l["id"], "work_date": l["work_date"],
                      "job_ref_id": l["job_ref_id"], "job_ref": l["job_ref"],
                      "job_role": l.get("job_role", ""),
                      "rate": l.get("rate", "Normal"), "narrative": l["narrative"],
                      "hours": l["hours"]} for l in self.db.get_lines(tid)]
            used = "adjusted"
        return {"id": tid, "name": row["full_name"], "owner_role": row["job_role"],
                "week": _fmt_week(wk), "status": row["status"],
                "basis": used, "has_snapshot": has_snap,
                "days": [(wk + dt.timedelta(days=i)).isoformat() for i in range(7)],
                "day_labels": [(wk + dt.timedelta(days=i)).strftime("%A %d %b")
                               for i in range(7)],
                "lines": lines}

    def save_admin_lines(self, tid, lines):
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"ok": False, "error": "Not allowed."}
        for l in lines:
            ok, msg = validate_line(l["narrative"], l["hours"])
            if not ok:
                return {"ok": False, "error": msg}
        self.db.save_timesheet_lines(u["email"], tid, lines)
        self.db.audit(u["email"], "ADMIN_EDIT", "", f"timesheet {tid}")
        return {"ok": True}

    def job_costs(self):
        self._need()
        # Hide empty non-project categories (HOLIDAY etc), but never hide one
        # that carries a quote, variations, hours or invoices.
        return [d for d in self.db.job_cost_summary()
                if not (d["kind"] != "Project" and d["quote_current"] == 0
                        and d["hours"] == 0 and d["invoiced"] == 0
                        and d["bills"] == 0)]

    def expenses(self):
        u = self._need()
        is_mgr = u["user_type"] in ("Manager", "Admin")
        out = []
        for e in self.db.list_expenses(user_id=None if is_mgr else u["id"]):
            out.append({"id": e["id"], "date": e["exp_date"], "who": e["full_name"],
                        "project": (f"{e['job_ref']}" if e.get("job_ref") else ""),
                        "desc": e["description"], "value": e["value"],
                        "status": e.get("status", "Logged"),
                        "receipt": bool(e["receipt"]), "receipt_path": e["receipt"],
                        "payment": bool(e.get("payment_receipt")),
                        "payment_path": e.get("payment_receipt", ""),
                        "value_n": e["value"], "job_ref_id": e.get("job_ref_id")})
        return out

    def users(self):
        self._need()
        return [{"id": u["id"], "name": u["full_name"], "email": u["email"],
                 "role": u["job_role"], "type": u["user_type"],
                 "active": bool(u["active"]),
                 "rate": (u["default_rate"] if "default_rate" in u.keys() else 0) or 0}
                for u in self.db.list_users()]

    def jobs(self):
        self._need()
        return [{"id": j["id"], "ref": j["job_ref"], "title": j["job_title"],
                 "client": j["client"], "type": j.get("kind", "Project"),
                 "status": j["status"], "notes": j["notes"],
                 "quote": j.get("quote_value", 0),
                 "quote_att": bool(j.get("quote_attachment")),
                 "quote_att_path": j.get("quote_attachment", "")}
                for j in self.db.list_jobs()]

    def roles(self):
        self._need()
        return [{"id": r["id"], "name": r["name"], "enabled": bool(r["enabled"])}
                for r in self.db.list_roles()]

    def audit(self):
        self._need()
        return [{"time": a["timestamp"], "user": a["user_email"], "action": a["action"],
                 "old": a["old_value"], "new": a["new_value"]}
                for a in self.db.get_audit(300)]

    # ---- option lists for dropdowns ----
    def options(self):
        self._need()
        return {"roles": [r["name"] for r in self.db.list_roles(only_enabled=True)],
                "user_types": USER_TYPES, "statuses": TS_STATUSES,
                "jobs": [{"id": j["id"], "label": f"{j['job_ref']} - {j['job_title']}"}
                         for j in self.db.list_jobs()],
                "active_jobs": self.timesheet_jobs(),
                "users": [{"id": u["id"], "name": u["full_name"]}
                          for u in self.db.list_users()]}

    def _admin(self):
        u = self._need()
        if u["user_type"] != "Admin":
            raise PermissionError("Admins only")
        return u

    # ---- users ----
    def add_user(self, name, email, password, role, utype, active, rate=0):
        u = self._admin()
        try:
            self.db.create_user(u["email"], name, email, password, role, utype,
                                bool(active), float(rate or 0))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def update_user(self, uid, name, email, role, utype, active, rate=0):
        u = self._admin()
        try:
            self.db.update_user(u["email"], uid, name, email, role, utype,
                                bool(active), float(rate or 0))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def reset_password(self, uid, newpw):
        u = self._admin()
        self.db.set_password(u["email"], uid, newpw)
        return {"ok": True}

    def delete_user(self, uid):
        u = self._admin()
        if uid == u["id"]:
            return {"ok": False, "error": "You cannot delete yourself."}
        res = self.db.delete_user(u["email"], uid)
        return {"ok": True, "result": res}

    def change_my_password(self, newpw):
        u = self._need()
        self.db.set_password(u["email"], u["id"], newpw)
        return {"ok": True}

    # ---- jobs ----
    def add_job(self, ref, title, client, kind, status, notes, quote, attach_src=""):
        u = self._admin()
        try:
            stored = self._store(attach_src, f"QUOTE_{(ref or 'x').replace('/', '-')}_",
                                 job_ref=ref) \
                if attach_src else ""
            self.db.create_job(u["email"], ref, title, client, status, notes, kind,
                               float(quote or 0), stored)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def update_job(self, jid, ref, title, client, kind, status, notes, quote,
                   attach_src="", keep_attach=""):
        u = self._admin()
        try:
            stored = self._store(attach_src, f"QUOTE_{(ref or 'x').replace('/', '-')}_",
                                 job_ref=ref) \
                if attach_src else keep_attach
            self.db.update_job(u["email"], jid, ref, title, client, status, notes, kind,
                               float(quote or 0), stored)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def delete_job(self, jid):
        u = self._admin()
        return {"ok": True, "result": self.db.delete_job(u["email"], jid)}

    # ---- roles ----
    def add_role(self, name):
        u = self._admin()
        self.db.add_role(u["email"], name)
        return {"ok": True}

    def toggle_role(self, rid, enabled):
        u = self._admin()
        self.db.set_role_enabled(u["email"], rid, bool(enabled))
        return {"ok": True}

    # ---- files / receipts ----
    def pick_file(self):
        self._need()
        import webview
        res = self.window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=("Receipts (*.pdf;*.png;*.jpg;*.jpeg;*.gif;*.heic)",
                        "All files (*.*)"))
        if not res:
            return None
        return res[0] if isinstance(res, (list, tuple)) else res

    def open_file(self, path):
        try:
            os.startfile(path)
            return True
        except Exception:
            return False

    # ---- external timesheet import/export ----
    def export_timesheet_template(self, week_start=None, user_id=None):
        self._need()
        import webview
        wk = (monday_of(dt.date.fromisoformat(week_start)) if week_start
              else monday_of(dt.date.today()))
        name = email = ""
        if user_id:
            us = self.db.get_user(int(user_id))
            if us:
                name, email = us["full_name"], us["email"]
        fn = f"Timesheet_{(email.split('@')[0] if email else 'blank')}_{wk.isoformat()}.xlsx"
        res = self.window.create_file_dialog(webview.SAVE_DIALOG, save_filename=fn)
        if not res:
            return {"ok": False}
        path = res if isinstance(res, str) else res[0]
        try:
            build_timesheet_template(
                path, self.db.list_jobs(only_active=True),
                [r["name"] for r in self.db.list_roles(only_enabled=True)],
                self.cfg.get("company_name", ""),
                week_start=wk, person_name=name, person_email=email)
            return {"ok": True, "path": path}
        except ImportError:
            return {"ok": False, "error": "Template needs openpyxl (pip install openpyxl)."}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def import_external_timesheet(self):
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"ok": False, "error": "Managers/Admins only."}
        import webview
        res = self.window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=("Excel timesheet (*.xlsx)", "All files (*.*)"))
        if not res:
            return {"ok": False}
        path = res[0] if isinstance(res, (list, tuple)) else res
        try:
            data = parse_timesheet_xlsx(path)
        except Exception as e:
            return {"ok": False, "error": "Could not read that file: " + str(e)}
        email = data.get("email", "")
        if not email:
            return {"ok": False, "error": "No email found in the sheet (cell B5)."}
        user = self.db.get_user_by_email(email)
        if not user:
            return {"ok": False, "error": f"No user with email '{email}'. Add them "
                    "under Users first - only people already set up can be imported."}
        wkraw = _to_iso_date(data.get("week"))
        if not wkraw:
            return {"ok": False, "error": "No valid 'Week commencing' date in cell B6."}
        wk = monday_of(dt.date.fromisoformat(wkraw))
        jobmap = {j["job_ref"].strip().lower(): j["id"] for j in self.db.list_jobs()}
        lines, errors = [], []
        for i, l in enumerate(data.get("lines", []), 1):
            jref = (l.get("job_ref") or "").strip()
            if not jref:
                errors.append(f"Row {i}: no job reference."); continue
            jid = jobmap.get(jref.lower())
            if not jid:
                errors.append(f"Row {i}: job reference '{jref}' not recognised."); continue
            wd = _to_iso_date(l.get("date"))
            if not wd:
                errors.append(f"Row {i}: missing/invalid date."); continue
            try:
                hrs = float(l.get("hours"))
            except (TypeError, ValueError):
                errors.append(f"Row {i}: hours '{l.get('hours')}' is not a number."); continue
            if hrs <= 0:
                errors.append(f"Row {i}: hours must be greater than zero."); continue
            narr = (l.get("narrative") or "").strip()
            if not narr:
                errors.append(f"Row {i}: description is required."); continue
            ot = l.get("overtime") if l.get("overtime") in RATE_OPTIONS else "Normal"
            lines.append({"work_date": wd, "job_ref_id": jid, "narrative": narr,
                          "hours": hrs, "job_role": (l.get("role") or user["job_role"]),
                          "rate": ot})
        if not lines:
            return {"ok": False, "error": "No valid rows found to import.",
                    "details": errors}
        ts = self.db.get_or_create_timesheet(user["id"], wk)
        self.db.submit_timesheet(email, ts["id"], lines)
        self.db.audit(u["email"], "IMPORT_TIMESHEET", "",
                      f"{email} week {wk.isoformat()} ({len(lines)} lines)")
        return {"ok": True, "name": user["full_name"], "week": _fmt_week(wk),
                "count": len(lines),
                "total": round(sum(x["hours"] for x in lines), 2),
                "errors": errors}

    def _receipts_dir(self):
        base = os.path.dirname(self.cfg.get("db_path", "")) or DATA_DIR
        folder = os.path.join(base, "Receipts")
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception:
            pass
        return folder

    def _job_dir(self, job_ref, create=True):
        """Receipts\\<JOB REF>\\ - so Explorer shows one job's documents rather
        than everybody's. Job references can contain characters Windows won't
        allow in a folder name (Telford 400kv/132kv), so they get sanitised."""
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", str(job_ref or "Unfiled"))
        safe = safe.strip().rstrip(". ")[:80] or "Unfiled"
        folder = os.path.join(self._receipts_dir(), safe)
        if create:
            try:
                os.makedirs(folder, exist_ok=True)
            except Exception:
                return self._receipts_dir()
        return folder

    def _job_ref_for(self, table, rec_id):
        """The job a variation / invoice / bill belongs to - the update_*
        methods are given the record's own id, not the job's."""
        try:
            with self.db.conn() as con:
                r = con.execute(
                    f"SELECT j.job_ref FROM {table} t "
                    f"JOIN job_refs j ON j.id=t.job_ref_id WHERE t.id=?",
                    (int(rec_id),)).fetchone()
                return r["job_ref"] if r else ""
        except Exception:
            return ""

    def _job_ref_of(self, job_id):
        try:
            j = next((x for x in self.db.list_jobs() if x["id"] == int(job_id)), None)
            return j["job_ref"] if j else ""
        except Exception:
            return ""

    @staticmethod
    def _unique(path):
        """Never silently overwrite a document that is already there."""
        if not os.path.exists(path):
            return path
        stem, ext = os.path.splitext(path)
        for n in range(2, 999):
            cand = f"{stem} ({n}){ext}"
            if not os.path.exists(cand):
                return cand
        return path

    def _store(self, src, prefix="", job_ref=None):
        """Put an attachment where it belongs. Given a job_ref it goes into
        that job's own folder; without one it stays in the Receipts root
        (staff expense receipts, which aren't tied to a job)."""
        if not src or not os.path.exists(src):
            return src or ""
        root = os.path.abspath(self._receipts_dir())
        dest_dir = self._job_dir(job_ref) if job_ref else self._receipts_dir()
        here = os.path.dirname(os.path.abspath(src))
        if here == os.path.abspath(dest_dir):
            return src                      # already in the right place
        dest = self._unique(os.path.join(dest_dir, prefix + os.path.basename(src)))
        try:
            # A dropped file is written into the Receipts root as a staging
            # step - move that rather than leaving a duplicate behind.
            if here == root:
                shutil.move(src, dest)
            else:
                shutil.copy2(src, dest)
            return dest
        except Exception:
            return src

    # ---- expenses (write) ----
    def add_expense(self, date, job_ref_id, description, value, receipt_src):
        u = self._need()
        stored = self._store(receipt_src,
                             f"{date}_{u['full_name'].replace(' ', '')}_")
        self.db.create_expense(u["email"], u["id"], date,
                               int(job_ref_id) if job_ref_id else None,
                               description, float(value or 0), stored)
        return {"ok": True}

    def update_expense(self, eid, date, job_ref_id, description, value,
                       receipt_src, keep_receipt):
        u = self._need()
        stored = self._store(receipt_src,
                             f"{date}_{u['full_name'].replace(' ', '')}_") if receipt_src else keep_receipt
        self.db.update_expense(u["email"], eid, date,
                               int(job_ref_id) if job_ref_id else None,
                               description, float(value or 0), stored)
        return {"ok": True}

    def delete_expense(self, eid):
        u = self._need()
        self.db.delete_expense(u["email"], eid)
        return {"ok": True}

    def attach_payment(self, eid, receipt_src):
        u = self._admin()
        stored = self._store(receipt_src, "PAYMENT_")
        self.db.set_expense_payment(u["email"], eid, stored)
        return {"ok": True}

    # ---- job cost rates ----
    def rates_for_job(self, job_id):
        self._need()
        rates = self.db.get_job_rates(job_id)
        defs = self.db.user_default_rates()
        return [{"id": x["id"], "name": x["full_name"], "role": x["job_role"],
                 "rate": rates.get(x["id"], 0), "default": defs.get(x["id"], 0)}
                for x in self.db.list_users(include_inactive=False)]

    def save_rates(self, job_id, rates):
        u = self._need()
        for uid, r in rates.items():
            self.db.set_job_rate(u["email"], int(uid), int(job_id), float(r or 0))
        return {"ok": True}

    # ---- variations & invoices (viewing: Manager/Admin, editing: Admin only) ----

    def _need_admin(self):
        u = self._need()
        if u["user_type"] != "Admin":
            return None
        return u

    def job_billing(self, job_id):
        """Everything the invoicing popup needs for one job."""
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"error": "Not allowed."}
        jid = int(job_id)
        summary = next((d for d in self.db.job_cost_summary() if d["id"] == jid), None)
        if not summary:
            return {"error": "Job not found."}
        jrow = next((j for j in self.db.list_jobs() if j["id"] == jid), None)
        quote_doc = (jrow.get("quote_attachment", "") if jrow else "") or ""
        return {
            "job": {"id": jid, "job_ref": summary["job_ref"],
                    "job_title": summary["job_title"],
                    "quote_doc": bool(quote_doc), "quote_doc_path": quote_doc,
                    "quote": summary["quote"], "variations": summary["variations"],
                    "quote_current": summary["quote_current"],
                    "invoiced": summary["invoiced"],
                    "invoiced_paid": summary["invoiced_paid"],
                    "invoiced_unpaid": summary["invoiced_unpaid"],
                    "to_invoice": summary["to_invoice"],
                    "hours": summary["hours"], "cost": summary["cost"],
                    "bills": summary["bills"], "bills_unpaid": summary["bills_unpaid"],
                    "total_cost": summary["total_cost"],
                    "profit": summary["profit"], "margin": summary["margin"]},
            "can_edit": u["user_type"] == "Admin",
            "variations": [{"id": v["id"], "ref": v["var_ref"],
                            "desc": v["description"], "value": v["value"],
                            "date": v["var_date"],
                            "att": bool(v.get("attachment")),
                            "att_path": v.get("attachment", ""),
                            "by": v["created_by"]}
                           for v in self.db.list_variations(jid)],
            "invoices": [{"id": i["id"], "no": i["invoice_no"], "date": i["inv_date"],
                          "desc": i["description"], "value": i["value"],
                          "status": i["status"], "paid_date": i["paid_date"],
                          "att": bool(i.get("attachment")),
                          "att_path": i.get("attachment", ""),
                          "by": i["created_by"]}
                         for i in self.db.list_invoices(jid)],
            "bills": [{"id": b["id"], "supplier": b["supplier"], "no": b["bill_no"],
                       "date": b["bill_date"], "desc": b["description"],
                       "value": b["value"], "category": b.get("category", ""),
                       "status": b["status"], "paid_date": b["paid_date"],
                       "att": bool(b.get("attachment")),
                       "att_path": b.get("attachment", ""),
                       "by": b["created_by"]}
                      for b in self.db.list_bills(jid)],
        }

    def add_variation(self, job_id, ref, desc, value, vdate, attach_src=""):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can add variations."}
        try:
            val = float(value)
        except (TypeError, ValueError):
            return {"ok": False, "error": "Variation value must be a number."}
        if not (desc or "").strip():
            return {"ok": False, "error": "Please describe the variation."}
        stored = self._store(attach_src, f"VAR_{(ref or 'x').replace('/', '-')}_",
                             job_ref=self._job_ref_of(job_id)) \
            if attach_src else ""
        self.db.add_variation(u["email"], int(job_id), ref, desc, val, vdate, stored)
        return {"ok": True}

    def update_variation(self, vid, ref, desc, value, vdate, attach_src="",
                         keep_attach=""):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can edit variations."}
        try:
            val = float(value)
        except (TypeError, ValueError):
            return {"ok": False, "error": "Variation value must be a number."}
        if not (desc or "").strip():
            return {"ok": False, "error": "Please describe the variation."}
        stored = self._store(attach_src, f"VAR_{(ref or 'x').replace('/', '-')}_",
                             job_ref=self._job_ref_for("job_variations", vid)) \
            if attach_src else keep_attach
        self.db.update_variation(u["email"], int(vid), ref, desc, val, vdate, stored)
        return {"ok": True}

    # ---- the job's document folder ----

    def _doc_info(self, path):
        """Filename, size and date from disk - and whether it is still there.
        A document can go missing if someone tidies the Receipts folder, and
        it is far better to say so than to show a button that does nothing."""
        out = {"name": os.path.basename(path or ""), "exists": False,
               "size": 0, "modified": ""}
        try:
            if path and os.path.exists(path):
                st = os.stat(path)
                out["exists"] = True
                out["size"] = st.st_size
                out["modified"] = dt.datetime.fromtimestamp(
                    st.st_mtime).strftime("%Y-%m-%d")
        except Exception:
            pass
        return out

    def job_documents(self, job_id):
        """Every document attached anywhere on this job, in one list: the
        quote, anything on a variation, invoice or bill, plus loose documents
        filed straight against the job."""
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"error": "Not allowed."}
        jid = int(job_id)
        docs = []

        def add(kind, ref, date, desc, path, value=None, did=None, by=""):
            if not path:
                return
            d = {"kind": kind, "ref": ref or "", "date": date or "",
                 "desc": desc or "", "path": path, "value": value,
                 "doc_id": did, "by": by}
            d.update(self._doc_info(path))
            docs.append(d)

        jrow = next((j for j in self.db.list_jobs() if j["id"] == jid), None)
        if jrow:
            add("Quote", jrow["job_ref"], "", jrow.get("job_title", ""),
                jrow.get("quote_attachment", ""), jrow.get("quote_value"))
        for v in self.db.list_variations(jid):
            add("Variation", v["var_ref"], v["var_date"], v["description"],
                v.get("attachment", ""), v["value"], by=v.get("created_by", ""))
        for i in self.db.list_invoices(jid):
            add("Invoice", i["invoice_no"], i["inv_date"], i["description"],
                i.get("attachment", ""), i["value"], by=i.get("created_by", ""))
        for b in self.db.list_bills(jid):
            # Don't repeat the supplier in both columns: it's the reference
            # only when there's no bill number, and the description otherwise.
            if b["bill_no"]:
                ref, desc = b["bill_no"], b["supplier"]
                if b["description"]:
                    desc = (desc + " - " + b["description"]) if desc \
                        else b["description"]
            else:
                ref, desc = b["supplier"], b["description"]
            add("Bill", ref, b["bill_date"], desc,
                b.get("attachment", ""), b["value"], by=b.get("created_by", ""))
        for d in self.db.list_documents(jid):
            add("Document", d["title"], d["doc_date"], d["description"],
                d.get("attachment", ""), None, did=d["id"],
                by=d.get("created_by", ""))
        docs.sort(key=lambda x: (x["date"] or "9999", x["kind"]))
        jref = (jrow or {}).get("job_ref", "")
        return {"docs": docs, "can_edit": u["user_type"] == "Admin",
                # this job's own folder, so Explorer shows what the app shows
                "folder": self._job_dir(jref, create=False) if jref
                          else self._receipts_dir(),
                "loose": sum(1 for d in docs if d["path"] and
                             os.path.dirname(os.path.abspath(d["path"])) ==
                             os.path.abspath(self._receipts_dir())),
                "missing": sum(1 for d in docs if not d["exists"])}

    def document_folders(self):
        """One folder per job, with how many documents it holds. Counted with
        five grouped queries rather than one per job, so it stays quick even
        with a lot of jobs sitting on a network share."""
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"folders": [], "total": 0}
        counts, latest = {}, {}

        def bump(jid, n=1, when=""):
            if jid is None:
                return
            counts[jid] = counts.get(jid, 0) + n
            if when and when > latest.get(jid, ""):
                latest[jid] = when

        with self.db.conn() as con:
            for r in con.execute("SELECT id FROM job_refs "
                                 "WHERE quote_attachment!=''").fetchall():
                bump(r["id"])
            for tbl, dcol in (("job_variations", "var_date"),
                              ("job_invoices", "inv_date"),
                              ("job_bills", "bill_date"),
                              ("job_documents", "doc_date")):
                for r in con.execute(
                        f"SELECT job_ref_id jid, COUNT(*) c, MAX({dcol}) d "
                        f"FROM {tbl} WHERE attachment!='' "
                        f"GROUP BY job_ref_id").fetchall():
                    bump(r["jid"], r["c"], r["d"] or "")
        jobs = {j["id"]: j for j in self.db.list_jobs()}
        out = []
        for jid, n in counts.items():
            j = jobs.get(jid)
            if not j:
                continue
            out.append({"id": jid, "job_ref": j["job_ref"],
                        "job_title": j.get("job_title", ""),
                        "client": j.get("client", ""),
                        "status": j.get("status", ""),
                        "count": n, "latest": latest.get(jid, "")})
        # most recent activity first - the jobs actually being worked on
        out.sort(key=lambda x: (x["latest"] or "", x["count"]), reverse=True)
        return {"folders": out, "total": sum(counts.values())}

    def tidy_documents(self):
        """Move every attachment still sitting loose in the Receipts root into
        its job's folder, and update the stored path to match.

        Deliberately cautious: files are MOVED one at a time, and the database
        is only updated once a move has actually succeeded. If a file is open,
        locked or missing it is skipped and reported - never lost, and never
        left with the database pointing somewhere the file isn't."""
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can tidy documents."}
        root = os.path.abspath(self._receipts_dir())
        moved, skipped, errors = 0, 0, []
        # (table, id column, path column) - plus the quote on job_refs itself
        targets = [("job_variations", "attachment"), ("job_invoices", "attachment"),
                   ("job_bills", "attachment"), ("job_documents", "attachment")]
        with self.db.conn() as con:
            rows = []
            for tbl, col in targets:
                for r in con.execute(
                        f"SELECT t.id, t.{col} p, j.job_ref ref FROM {tbl} t "
                        f"JOIN job_refs j ON j.id=t.job_ref_id "
                        f"WHERE t.{col}!=''").fetchall():
                    rows.append((tbl, col, r["id"], r["p"], r["ref"]))
            for r in con.execute("SELECT id, quote_attachment p, job_ref ref "
                                 "FROM job_refs WHERE quote_attachment!=''").fetchall():
                rows.append(("job_refs", "quote_attachment", r["id"], r["p"], r["ref"]))
        for tbl, col, rid, path, ref in rows:
            try:
                if not path or not os.path.exists(path):
                    skipped += 1
                    continue
                if os.path.dirname(os.path.abspath(path)) != root:
                    continue                      # already in a job folder
                dest = self._unique(os.path.join(self._job_dir(ref),
                                                 os.path.basename(path)))
                shutil.move(path, dest)           # move first...
                with self.db.conn() as con:       # ...then record it
                    con.execute(f"UPDATE {tbl} SET {col}=? WHERE id=?", (dest, rid))
                moved += 1
            except Exception as e:
                errors.append(f"{os.path.basename(path or '?')}: {e}")
        if moved:
            self.db.audit(u["email"], "TIDY_DOCUMENTS", "",
                          f"{moved} file(s) moved into job folders")
        return {"ok": True, "moved": moved, "skipped": skipped, "errors": errors[:12]}

    def add_document(self, job_id, title, desc, ddate, attach_src):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can add documents."}
        if not attach_src:
            return {"ok": False, "error": "Pick or drop a file first."}
        stored = self._store(attach_src, "DOC_",
                             job_ref=self._job_ref_of(job_id))
        self.db.add_document(u["email"], int(job_id),
                             title or os.path.basename(stored), desc, ddate, stored)
        return {"ok": True}

    def delete_document(self, did):
        """Takes it off the job's list. The file itself stays in the Receipts
        folder - deleting someone's only copy of a document from a button in a
        timesheet app would be a nasty surprise."""
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can remove documents."}
        self.db.delete_document(u["email"], int(did))
        return {"ok": True}

    def open_folder(self, path):
        """Open the folder a document lives in, selecting the file if we can."""
        self._need()
        try:
            if path and os.path.isfile(path):
                import subprocess
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
                return True
            target = path if (path and os.path.isdir(path)) \
                else os.path.dirname(path or "")
            if target and os.path.isdir(target):
                os.startfile(target)
                return True
        except Exception:
            pass
        return False

    def save_dropped_file(self, filename, b64):
        """Take a file DROPPED onto the window and put it on disk.

        A dropped file only ever reaches the page as its contents - browsers
        deliberately withhold the real path - so the bytes come across as
        base64 and get written into the Receipts folder here. Everything
        downstream (attaching, scanning) then works on a normal path, exactly
        as if the file had been chosen through the picker."""
        self._need()
        import base64
        try:
            raw = base64.b64decode(b64 or "")
        except Exception:
            return {"ok": False, "error": "That file could not be read."}
        if not raw:
            return {"ok": False, "error": "That file appears to be empty."}
        if len(raw) > 25 * 1024 * 1024:
            return {"ok": False, "error": "That file is too big (25 MB limit). "
                    "Use the Attach button to link it instead."}
        # strip any directory part (both separators) before sanitising, so a
        # crafted name can never write outside the Receipts folder
        base = (filename or "document").replace("\\", "/").rsplit("/", 1)[-1]
        safe = re.sub(r"[^A-Za-z0-9._ -]", "_", base)[:80].strip() or "document"
        dest = os.path.join(self._receipts_dir(),
                            f"{dt.datetime.now():%Y%m%d_%H%M%S}_{safe}")
        try:
            with open(dest, "wb") as fh:
                fh.write(raw)
        except Exception as e:
            return {"ok": False, "error": "Could not save that file: " + str(e)}
        return {"ok": True, "path": dest}

    def delete_variation(self, vid):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can delete variations."}
        self.db.delete_variation(u["email"], int(vid))
        return {"ok": True}

    def add_invoice(self, job_id, no, idate, desc, value, status="Raised",
                    paid_date="", attach_src=""):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can add invoices."}
        ok, msg = validate_invoice(no, value)
        if not ok:
            return {"ok": False, "error": msg}
        if self.db.invoice_exists(int(job_id), no):
            return {"ok": False,
                    "error": f"Invoice {no} already exists on this job."}
        stored = self._store(attach_src, f"INV_{(no or 'x').replace('/', '-')}_",
                             job_ref=self._job_ref_of(job_id)) \
            if attach_src else ""
        self.db.add_invoice(u["email"], int(job_id), no, idate, desc, float(value),
                            status, paid_date, stored)
        return {"ok": True}

    def update_invoice(self, iid, no, idate, desc, value, status, paid_date="",
                       attach_src="", keep_attach=""):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can edit invoices."}
        ok, msg = validate_invoice(no, value)
        if not ok:
            return {"ok": False, "error": msg}
        stored = self._store(attach_src, f"INV_{(no or 'x').replace('/', '-')}_",
                             job_ref=self._job_ref_for("job_invoices", iid)) \
            if attach_src else keep_attach
        self.db.update_invoice(u["email"], int(iid), no, idate, desc, float(value),
                               status, paid_date, stored)
        return {"ok": True}

    def set_invoice_paid(self, iid, paid):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can update invoices."}
        self.db.set_invoice_status(
            u["email"], int(iid), "Paid" if paid else "Raised",
            dt.date.today().isoformat() if paid else "")
        return {"ok": True}

    def delete_invoice(self, iid):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can delete invoices."}
        self.db.delete_invoice(u["email"], int(iid))
        return {"ok": True}

    # ---- bills (money out) ----

    def add_bill(self, job_id, supplier, no, bdate, desc, value, category="",
                 status="Received", paid_date="", attach_src=""):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can add bills."}
        ok, msg = validate_bill(supplier, value)
        if not ok:
            return {"ok": False, "error": msg}
        if no and self.db.bill_exists(int(job_id), supplier, no):
            return {"ok": False, "error":
                    f"Bill {no} from {supplier} is already on this job."}
        stored = self._store(attach_src, f"BILL_{(no or 'x').replace('/', '-')}_",
                             job_ref=self._job_ref_of(job_id)) \
            if attach_src else ""
        self.db.add_bill(u["email"], int(job_id), supplier, no, bdate, desc,
                         float(value), category, status, paid_date, stored)
        return {"ok": True}

    def update_bill(self, bid, supplier, no, bdate, desc, value, category, status,
                    paid_date="", attach_src="", keep_attach=""):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can edit bills."}
        ok, msg = validate_bill(supplier, value)
        if not ok:
            return {"ok": False, "error": msg}
        stored = self._store(attach_src, f"BILL_{(no or 'x').replace('/', '-')}_",
                             job_ref=self._job_ref_for("job_bills", bid)) \
            if attach_src else keep_attach
        self.db.update_bill(u["email"], int(bid), supplier, no, bdate, desc,
                            float(value), category, status, paid_date, stored)
        return {"ok": True}

    def set_bill_paid(self, bid, paid):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can update bills."}
        self.db.set_bill_status(
            u["email"], int(bid), "Paid" if paid else "Received",
            dt.date.today().isoformat() if paid else "")
        return {"ok": True}

    def delete_bill(self, bid):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can delete bills."}
        self.db.delete_bill(u["email"], int(bid))
        return {"ok": True}

    # ---- document scanning (offline) ----

    def pick_document(self):
        """File picker for a bill/invoice document, then read it in one step."""
        self._need()
        import webview
        # NOTE: pywebview validates this string against ^([\w ]+)\(...\)$ - the
        # description may contain ONLY letters, digits and spaces. An '&' (or
        # any punctuation) makes it raise ValueError, which surfaces as the
        # button silently doing nothing. Keep it plain.
        try:
            res = self.window.create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=False,
                file_types=("Bills and invoices (*.pdf;*.png;*.jpg;*.jpeg;*.tif;*.tiff)",
                            "All files (*.*)"))
        except Exception:
            # never let a file-filter problem leave the user staring at a
            # button that does nothing - fall back to an unfiltered picker
            res = self.window.create_file_dialog(webview.OPEN_DIALOG,
                                                 allow_multiple=False)
        if not res:
            return None
        return res[0] if isinstance(res, (list, tuple)) else res

    def scan_document(self, path):
        """Read a document and hand back suggested field values. This ONLY
        pre-fills the form - nothing is saved until the user checks it and
        presses Save."""
        self._need()
        return read_document(path)

    def ocr_status(self):
        """Tell the UI whether this PC can read documents, without importing
        anything heavy unless asked."""
        self._need()
        import importlib.util
        has_lib = all(importlib.util.find_spec(m) is not None
                      for m in ("pytesseract", "PIL"))
        return {"ready": bool(has_lib and _tesseract_cmd()),
                "lib": has_lib, "exe": bool(_tesseract_cmd()),
                "hint": OCR_INSTALL_HINT}

    def export_invoice_template(self):
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can import invoices."}
        import webview
        try:
            res = self.window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=f"Invoice_Import_{dt.datetime.now():%Y%m%d}.xlsx")
            if not res:
                return {"ok": False}
            path = res if isinstance(res, str) else res[0]
            build_invoice_template(
                path, self.db.list_jobs(only_active=True),
                self.cfg.get("company_name", ""))
            return {"ok": True, "path": path}
        except ImportError:
            return {"ok": False, "error": "Excel export needs openpyxl. "
                    "Run: pip install openpyxl"}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def import_invoices(self):
        """Bulk-import invoices from a filled-in spreadsheet. Rows whose invoice
        number already exists on that job are SKIPPED, not added again, so
        importing the same file twice cannot double-count invoiced value."""
        u = self._need_admin()
        if not u:
            return {"ok": False, "error": "Only an Admin can import invoices."}
        import webview
        res = self.window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=("Excel files (*.xlsx;*.xlsm)", "All files (*.*)"))
        if not res:
            return None
        path = res[0] if isinstance(res, (list, tuple)) else res
        try:
            rows = parse_invoices_xlsx(path)
        except ImportError:
            return {"ok": False, "error": "Excel import needs openpyxl. "
                    "Run: pip install openpyxl"}
        except Exception as e:
            return {"ok": False, "error": "Could not read that file.",
                    "details": [str(e)]}
        jobs = {(j["job_ref"] or "").strip().lower(): j["id"]
                for j in self.db.list_jobs()}
        added = skipped_dupe = 0
        errors, total = [], 0.0
        for n, r in enumerate(rows, start=2):
            ref = (r.get("job_ref") or "").strip()
            jid = jobs.get(ref.lower())
            if not jid:
                errors.append(f"Row {n}: job reference '{ref}' not recognised.")
                continue
            ok, msg = validate_invoice(r.get("invoice_no"), r.get("value"))
            if not ok:
                errors.append(f"Row {n}: {msg}")
                continue
            if self.db.invoice_exists(jid, r.get("invoice_no")):
                skipped_dupe += 1
                continue
            status = "Paid" if str(r.get("status", "")).strip().lower() == "paid" \
                else "Raised"
            self.db.add_invoice(u["email"], jid, r.get("invoice_no"),
                                r.get("inv_date") or "", r.get("description") or "",
                                float(r.get("value")), status,
                                (r.get("inv_date") or "") if status == "Paid" else "")
            added += 1
            total += float(r.get("value"))
        return {"ok": True, "added": added, "duplicates": skipped_dupe,
                "total": round(total, 2), "errors": errors}

    def job_cost_pdf(self, job_id):
        u = self._need()
        import webview, re
        try:
            d = job_cost_breakdown(self.db, int(job_id))
            safe = re.sub(r"[^A-Za-z0-9_-]", "_", str(d["job_ref"]))[:40]
            res = self.window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=f"JobCost_{safe}_{dt.datetime.now():%H%M%S}.pdf")
            if not res:
                return {"ok": False}
            path = res if isinstance(res, str) else res[0]
            export_job_cost_pdf(path, self.cfg.get("company_name", ""),
                                d, u["full_name"])
            return {"ok": True, "path": path}
        except ImportError:
            return {"ok": False, "error": "PDF export needs fpdf2. Run: pip install fpdf2"}
        except Exception as e:
            import traceback
            return {"ok": False, "error": "Job Cost PDF failed:\n" + repr(e)
                    + "\n\n" + traceback.format_exc().splitlines()[-2]}

    def _report_fd(self, f):
        fd = {}
        if f.get("user_id"):
            fd["user_id"] = int(f["user_id"])
        if f.get("job_role"):
            fd["job_role"] = f["job_role"]
        if f.get("job_ref_id"):
            fd["job_ref_id"] = int(f["job_ref_id"])
        if f.get("status"):
            fd["status"] = f["status"]
        fd["basis"] = "submitted" if f.get("basis") == "submitted" else "adjusted"
        try:
            fd["date_from"] = dt.date.fromisoformat(f["date_from"])
            fd["date_to"] = dt.date.fromisoformat(f["date_to"])
        except Exception:
            pass
        return fd

    def run_report(self, f):
        self._need()
        rep = Reports(self.db)
        if f.get("type") == "Missing Timesheets":
            try:
                wk = monday_of(dt.date.fromisoformat(f["date_to"]))
            except Exception:
                wk = monday_of(dt.date.today())
            rows = rep.missing_timesheets(wk)
            return {"missing": True,
                    "rows": [{"name": m["full_name"], "role": m["job_role"],
                              "status": m["status"]} for m in rows],
                    "total": len(rows)}
        include_cost = bool(f.get("include_cost"))
        rows, total = rep.person_report(include_cost=include_cost, **self._report_fd(f))
        out_rows = [{"name": r["full_name"], "role": r["job_role"],
                     "date": r["work_date"], "job_ref": r["job_ref"],
                     "title": r["job_title"], "narrative": r["narrative"],
                     "hours": r["hours"], "status": r["status"]} for r in rows]
        resp = {"rows": out_rows, "total": total, "cost_included": include_cost}
        if include_cost:
            for i, r in enumerate(rows):
                out_rows[i]["cost"] = r.get("cost", 0)
            resp["total_cost"] = round(sum(r.get("cost", 0) for r in rows), 2)
        return resp

    def export_report(self, kind, f):
        u = self._need()
        import webview
        rep = Reports(self.db)
        include_cost = bool(f.get("include_cost"))
        if f.get("type") == "Missing Timesheets":
            include_cost = False  # no cost concept for a missing-timesheet list
            title = "Missing Timesheet Report"
            rows = [{"full_name": m["full_name"], "job_role": m["job_role"],
                     "work_date": "", "job_ref": "", "job_title": "",
                     "narrative": "Missing/incomplete", "hours": 0,
                     "status": m["status"]}
                    for m in rep.missing_timesheets(
                        monday_of(dt.date.fromisoformat(f["date_to"])))]
        else:
            title = f.get("type", "Report") + " Report"
            rows = [dict(r) for r in
                    rep.person_report(include_cost=include_cost, **self._report_fd(f))[0]]
        # The hours basis is deliberately NOT put in the title - these reports
        # go to clients, and "(Adjusted)" invites questions we don't want asked.
        ext = {"csv": ".csv", "xlsx": ".xlsx", "pdf": ".pdf"}[kind]
        res = self.window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=f"{title.replace(' ', '_')}_{dt.datetime.now():%H%M%S}{ext}")
        if not res:
            return {"ok": False}
        path = res if isinstance(res, str) else res[0]
        company = self.cfg.get("company_name", "")
        # 'basis' and 'include_cost' are internal switches - keep them out of
        # the Filters line as well, or the exported file still says "adjusted".
        flt = ", ".join(f"{k}={v}" for k, v in f.items()
                        if v and k not in ("basis", "include_cost"))
        columns = REPORT_COLUMNS if include_cost else [c for c in REPORT_COLUMNS
                                                        if c[0] != "cost"]
        try:
            if kind == "csv":
                export_csv(path, title, u["full_name"], flt, rows, company, columns=columns)
            elif kind == "xlsx":
                export_xlsx(path, title, u["full_name"], flt, rows, company, columns=columns)
            else:
                export_report_pdf(path, company, title,
                                  dt.date.today().isoformat(), u["full_name"], rows,
                                  ("job_ref" if f.get("type") == "Person" else "full_name"),
                                  ("Job" if f.get("type") == "Person" else "Person"),
                                  include_cost=include_cost)
            return {"ok": True, "path": path}
        except ImportError:
            lib = "openpyxl" if kind == "xlsx" else "fpdf2"
            return {"ok": False, "error": f"{kind.upper()} export needs '{lib}'. "
                    f"Run: pip install {lib}"}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def analytics(self, date_from=None, date_to=None):
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"error": "Not allowed."}
        today = dt.date.today()
        df = dt.date.fromisoformat(date_from) if date_from else today - dt.timedelta(days=90)
        dtt = dt.date.fromisoformat(date_to) if date_to else today
        dfi, dti = df.isoformat(), dtt.isoformat()
        with self.db.conn() as con:
            rates = {(r["user_id"], r["job_ref_id"]): r["rate"]
                     for r in con.execute("SELECT user_id, job_ref_id, rate "
                                          "FROM job_rates").fetchall()}
            defaults = {r["id"]: (r["default_rate"] or 0.0) for r in con.execute(
                "SELECT id, default_rate FROM users").fetchall()}
            jobs_meta = {r["id"]: (r["job_ref"], r["job_title"])
                         for r in con.execute("SELECT id, job_ref, job_title "
                                              "FROM job_refs").fetchall()}
            adj = con.execute(
                "SELECT u.id uid, u.full_name, j.id jid, j.job_ref, j.job_title, "
                "j.kind, l.hours, l.rate, l.work_date FROM timesheet_lines l "
                "JOIN timesheets t ON t.id=l.timesheet_id "
                "JOIN users u ON u.id=t.user_id "
                "JOIN job_refs j ON j.id=l.job_ref_id "
                "WHERE l.work_date BETWEEN ? AND ?", (dfi, dti)).fetchall()
            snaps = con.execute(
                "SELECT t.user_id uid, u.full_name, t.submitted_snapshot snap "
                "FROM timesheets t JOIN users u ON u.id=t.user_id "
                "WHERE t.submitted_snapshot!=''").fetchall()
        name_map = {}
        util = {}            # uid -> billable / non-billable
        ot = {}              # uid -> overtime hours & cost
        adj_person = {}      # uid -> adjusted hours
        adj_job = {}         # jid -> adjusted hours
        tot_bill = tot_non = 0.0
        weekly = {}          # monday iso -> billable / nonbill / overtime
        for r in adj:
            uid = r["uid"]; h = float(r["hours"]); rt = r["rate"] or "Normal"
            kind = r["kind"] or "Project"
            name_map[uid] = r["full_name"]
            try:
                wk = monday_of(dt.date.fromisoformat(r["work_date"])).isoformat()
            except Exception:
                wk = dfi
            wrow = weekly.setdefault(wk, {"billable": 0.0, "nonbill": 0.0, "overtime": 0.0})
            wrow["billable" if kind == "Project" else "nonbill"] += h
            if RATE_MULT.get(rt, 1.0) > 1.0:
                wrow["overtime"] += h
            d = util.setdefault(uid, {"billable": 0.0, "nonbill": 0.0})
            if kind == "Project":
                d["billable"] += h; tot_bill += h
            else:
                d["nonbill"] += h; tot_non += h
            extra = h * RATE_MULT.get(rt, 1.0) - h
            if extra > 0:
                o = ot.setdefault(uid, {"hours": 0.0, "cost": 0.0})
                o["hours"] += h
                _rt = rates[(uid, r["jid"])] if (uid, r["jid"]) in rates else defaults.get(uid, 0.0)
                o["cost"] += extra * _rt
            adj_person[uid] = adj_person.get(uid, 0.0) + h
            adj_job[r["jid"]] = adj_job.get(r["jid"], 0.0) + h
        sub_person = {}; sub_job = {}
        for srow in snaps:
            try:
                lines = json.loads(srow["snap"])
            except Exception:
                lines = []
            name_map.setdefault(srow["uid"], srow["full_name"])
            for l in lines:
                wd = l.get("work_date", "")
                if wd < dfi or wd > dti:
                    continue
                h = float(l.get("hours", 0) or 0)
                sub_person[srow["uid"]] = sub_person.get(srow["uid"], 0.0) + h
                jid = l.get("job_ref_id")
                if jid is not None:
                    sub_job[jid] = sub_job.get(jid, 0.0) + h
        # utilisation
        util_rows = []
        for uid, d in util.items():
            b = round(d["billable"], 2); n = round(d["nonbill"], 2); t = b + n
            util_rows.append({"name": name_map.get(uid, "?"), "billable": b,
                              "nonbill": n,
                              "pct": (round(b / t * 100, 1) if t else 0.0)})
        util_rows.sort(key=lambda x: -(x["billable"] + x["nonbill"]))
        comp_util = (round(tot_bill / (tot_bill + tot_non) * 100, 1)
                     if (tot_bill + tot_non) else 0.0)
        # profit leakage per person
        leak_person = []
        for uid in set(list(adj_person) + list(sub_person)):
            a = round(adj_person.get(uid, 0.0), 2)
            sb = round(sub_person.get(uid, 0.0), 2)
            leak_person.append({"name": name_map.get(uid, "?"), "submitted": sb,
                                "adjusted": a, "diff": round(a - sb, 2)})
        leak_person.sort(key=lambda x: -abs(x["diff"]))
        # profit leakage per job
        leak_job = []
        for jid in set(list(adj_job) + list(sub_job)):
            a = round(adj_job.get(jid, 0.0), 2)
            sb = round(sub_job.get(jid, 0.0), 2)
            ref, title = jobs_meta.get(jid, ("?", ""))
            leak_job.append({"job_ref": ref, "job_title": title, "submitted": sb,
                             "adjusted": a, "diff": round(a - sb, 2)})
        leak_job.sort(key=lambda x: -abs(x["diff"]))
        # overtime rows
        ot_rows = []
        for uid, o in ot.items():
            ot_rows.append({"name": name_map.get(uid, "?"),
                            "hours": round(o["hours"], 2),
                            "cost": round(o["cost"], 2)})
        ot_rows.sort(key=lambda x: -x["hours"])
        ot_tot_h = round(sum(x["hours"] for x in ot_rows), 2)
        ot_tot_c = round(sum(x["cost"] for x in ot_rows), 2)
        # portfolio health (cumulative budget vs quote)
        jcs = self.db.job_cost_summary()
        portfolio = []
        for d in jcs:
            # measured against the CURRENT contract value (quote + variations)
            if d["kind"] != "Project" or d["quote_current"] <= 0:
                continue
            pct = d["pct"]
            flag = "over" if pct > 100 else ("warn" if pct >= 80 else "ok")
            portfolio.append({"job_ref": d["job_ref"], "job_title": d["job_title"],
                              "quote": d["quote_current"], "cost": d["total_cost"],
                              "remaining": d["remaining"], "pct": pct, "flag": flag})
        portfolio.sort(key=lambda x: -x["pct"])
        # top jobs ranking by billable value logged
        costed = [d for d in jcs if d["cost"] > 0]
        if costed:
            top_metric = "cost"
            top_src = sorted(costed, key=lambda x: -x["cost"])[:8]
        else:
            top_metric = "hours"   # no charge rates set yet - rank by hours
            top_src = sorted([d for d in jcs if d["hours"] > 0],
                             key=lambda x: -x["hours"])[:8]
        top_jobs = [{"job_ref": d["job_ref"], "job_title": d["job_title"],
                     "cost": d["cost"], "hours": d["hours"], "quote": d["quote_current"],
                     "pct": d["pct"]} for d in top_src]
        # expenses
        outstanding = 0.0; period_total = 0.0; exp_by_job = {}
        for e in self.db.list_expenses():
            val = float(e.get("value", 0) or 0)
            if (e.get("status") or "Logged") != "Paid":
                outstanding += val
            ed = e.get("exp_date", "")
            if dfi <= ed <= dti:
                period_total += val
                k = e.get("job_ref") or "(no job)"
                exp_by_job[k] = round(exp_by_job.get(k, 0.0) + val, 2)
        exp_rows = sorted([{"job": k, "value": v} for k, v in exp_by_job.items()],
                          key=lambda x: -x["value"])
        # weekly series (every week in the period, zero-filled)
        wk_series = []
        wk0 = monday_of(df)
        while wk0 <= dtt:
            k = wk0.isoformat(); w = weekly.get(k, {"billable": 0.0, "nonbill": 0.0, "overtime": 0.0})
            wk_series.append({"week": k, "label": wk0.strftime("%d %b"),
                              "billable": round(w["billable"], 1), "nonbill": round(w["nonbill"], 1),
                              "overtime": round(w["overtime"], 1)})
            wk0 += dt.timedelta(days=7)
        # money in the period
        inv_period = inv_unpaid = bills_period = bills_unpaid = 0.0
        for i in self.db.list_invoices():
            v = float(i.get("value", 0) or 0)
            if dfi <= (i.get("inv_date") or "") <= dti:
                inv_period += v
            if (i.get("status") or "") != "Paid":
                inv_unpaid += v
        for b in self.db.list_bills():
            v = float(b.get("value", 0) or 0)
            if dfi <= (b.get("bill_date") or "") <= dti:
                bills_period += v
            if (b.get("status") or "") != "Paid":
                bills_unpaid += v
        # on-time submission rate: share of (active person x completed week) with a submitted sheet
        with self.db.conn() as con:
            active_ids = [r["id"] for r in con.execute("SELECT id FROM users WHERE active=1").fetchall()]
            done = {(r["user_id"], r["week_start"]) for r in con.execute(
                "SELECT user_id, week_start FROM timesheets WHERE status IN ('Submitted','Approved','Locked')").fetchall()}
        with self.db.conn() as con:
            first_wk = (con.execute("SELECT MIN(week_start) m FROM timesheets").fetchone() or {"m": None})["m"] or dfi
        weeks_done = [w["week"] for w in wk_series
                      if w["week"] >= first_wk and dt.date.fromisoformat(w["week"]) + dt.timedelta(days=6) < today]
        expected = len(active_ids) * len(weeks_done)
        got = sum(1 for uid in active_ids for w in weeks_done if (uid, w) in done)
        sub_rate = round(got / expected * 100, 1) if expected else 100.0
        months = max(1.0, ((dtt - df).days + 1) / 30.4)
        goals = dict(DEFAULT_GOALS); goals.update(self.cfg.get("goals") or {})
        return {
            "period": {"from": dfi, "to": dti, "months": round(months, 2),
                       "label": f"{df.strftime('%d %b %Y')} - {dtt.strftime('%d %b %Y')}"},
            "weekly": wk_series, "goals": goals,
            "money": {"invoiced": round(inv_period, 2), "invoiced_unpaid": round(inv_unpaid, 2),
                      "bills": round(bills_period, 2), "bills_unpaid": round(bills_unpaid, 2)},
            "submission_rate": sub_rate, "people": len(active_ids),
            "util": {"company_pct": comp_util, "billable": round(tot_bill, 2),
                     "nonbill": round(tot_non, 2), "rows": util_rows},
            "leak_person": leak_person, "leak_job": leak_job,
            "overtime": {"rows": ot_rows, "total_hours": ot_tot_h,
                         "total_cost": ot_tot_c},
            "portfolio": portfolio, "top_jobs": top_jobs,
            "top_metric": top_metric,
            "expenses": {"outstanding": round(outstanding, 2),
                         "period_total": round(period_total, 2), "by_job": exp_rows},
        }

    def job_timeline(self, weeks=12):
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"error": "Not allowed."}
        weeks = max(4, min(int(weeks or 12), 52))
        wk0 = monday_of(dt.date.today())
        wk_iso = [(wk0 - dt.timedelta(days=7 * i)).isoformat()
                  for i in range(weeks - 1, -1, -1)]
        wkset = set(wk_iso)
        with self.db.conn() as con:
            rows = con.execute(
                "SELECT j.id jid, j.job_ref, j.job_title, t.week_start, SUM(l.hours) h "
                "FROM timesheet_lines l JOIN timesheets t ON t.id=l.timesheet_id "
                "JOIN job_refs j ON j.id=l.job_ref_id "
                "WHERE t.week_start >= ? GROUP BY j.id, t.week_start",
                (wk_iso[0],)).fetchall()
        jobs = {}
        for r in rows:
            if r["week_start"] not in wkset:
                continue
            d = jobs.setdefault(r["job_ref"], {"id": r["jid"], "title": r["job_title"],
                                               "cells": {}, "total": 0.0})
            d["cells"][r["week_start"]] = round(r["h"], 2)
            d["total"] += r["h"]
        out = []
        for ref, d in sorted(jobs.items(), key=lambda kv: -kv[1]["total"]):
            out.append({"id": d["id"], "job_ref": ref, "job_title": d["title"],
                        "total": round(d["total"], 2),
                        "cells": [d["cells"].get(w, 0) for w in wk_iso]})
        maxcell = max([c for j in out for c in j["cells"]], default=1) or 1
        labels = [w[8:10] + "/" + w[5:7] for w in wk_iso]
        return {"weeks": labels, "jobs": out, "maxcell": maxcell,
                "grand": round(sum(j["total"] for j in out), 2)}

    def job_timeline_detail(self, job_id, weeks=12):
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"error": "Not allowed."}
        weeks = max(4, min(int(weeks or 12), 52))
        wk0 = monday_of(dt.date.today())
        wk_iso = [(wk0 - dt.timedelta(days=7 * i)).isoformat()
                  for i in range(weeks - 1, -1, -1)]
        wkset = set(wk_iso)
        with self.db.conn() as con:
            job = con.execute("SELECT job_ref, job_title FROM job_refs WHERE id=?",
                              (int(job_id),)).fetchone()
            rows = con.execute(
                "SELECT u.id uid, u.full_name nm, t.week_start, SUM(l.hours) h "
                "FROM timesheet_lines l JOIN timesheets t ON t.id=l.timesheet_id "
                "JOIN users u ON u.id=t.user_id "
                "WHERE l.job_ref_id=? AND t.week_start >= ? "
                "GROUP BY u.id, t.week_start", (int(job_id), wk_iso[0])).fetchall()
        people = {}
        for r in rows:
            if r["week_start"] not in wkset:
                continue
            d = people.setdefault(r["uid"], {"name": r["nm"], "cells": {}, "total": 0.0})
            d["cells"][r["week_start"]] = round(r["h"], 2)
            d["total"] += r["h"]
        out = [{"uid": uid, "name": d["name"], "total": round(d["total"], 2),
                "cells": [d["cells"].get(w, 0) for w in wk_iso]}
               for uid, d in sorted(people.items(), key=lambda kv: -kv[1]["total"])]
        maxcell = max([c for p in out for c in p["cells"]], default=1) or 1
        labels = [w[8:10] + "/" + w[5:7] for w in wk_iso]
        return {"job_ref": job["job_ref"] if job else "?",
                "job_title": job["job_title"] if job else "",
                "weeks": labels, "week_iso": wk_iso, "people": out,
                "maxcell": maxcell,
                "grand": round(sum(p["total"] for p in out), 2)}

    def timeline_lines(self, job_id, user_id, week_start):
        u = self._need()
        if u["user_type"] not in ("Manager", "Admin"):
            return {"error": "Not allowed."}
        wk = monday_of(dt.date.fromisoformat(week_start))
        with self.db.conn() as con:
            job = con.execute("SELECT job_ref, job_title FROM job_refs WHERE id=?",
                              (int(job_id),)).fetchone()
            person = con.execute("SELECT full_name FROM users WHERE id=?",
                                 (int(user_id),)).fetchone()
            rows = con.execute(
                "SELECT l.work_date, l.narrative, l.hours, l.rate, l.job_role "
                "FROM timesheet_lines l JOIN timesheets t ON t.id=l.timesheet_id "
                "WHERE l.job_ref_id=? AND t.user_id=? AND t.week_start=? "
                "ORDER BY l.work_date, l.id",
                (int(job_id), int(user_id), wk.isoformat())).fetchall()
        lines = []
        for r in rows:
            try:
                lab = dt.date.fromisoformat(r["work_date"]).strftime("%a %d %b")
            except Exception:
                lab = r["work_date"]
            lines.append({"date": lab, "narrative": r["narrative"],
                          "hours": r["hours"], "rate": r["rate"],
                          "role": r["job_role"]})
        return {"job_ref": job["job_ref"] if job else "?",
                "job_title": job["job_title"] if job else "",
                "name": person["full_name"] if person else "?",
                "week": _fmt_week(wk), "lines": lines,
                "total": round(sum(x["hours"] for x in lines), 2)}

    def app_config(self):
        self._need()
        return {"timeout_minutes": int(self.cfg.get("inactivity_timeout_minutes", 15)),
                "daily_warning": float(self.cfg.get("daily_hours_warning", 10)),
                "db_path": self.cfg.get("db_path", "")}

    def month_end_export(self, ym):
        u = self._need()
        import webview
        try:
            y, m = [int(x) for x in ym.split("-")]
            m_start = dt.date(y, m, 1)
            m_end = (dt.date(y, 12, 31) if m == 12
                     else dt.date(y, m + 1, 1) - dt.timedelta(days=1))
        except Exception:
            return {"ok": False, "error": "Use YYYY-MM, e.g. 2026-06."}
        rows = [dict(r) for r in Reports(self.db)._base_rows(date_from=m_start,
                                                             date_to=m_end)]
        if not rows:
            return {"ok": False, "error": "No hours recorded for that month."}
        res = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=f"Month_End_{ym}_{dt.datetime.now():%H%M%S}.xlsx")
        if not res:
            return {"ok": False}
        path = res if isinstance(res, str) else res[0]
        try:
            export_month_xlsx(path, self.cfg.get("company_name", ""),
                              m_start.strftime("%B %Y"), u["full_name"], rows)
            return {"ok": True, "path": path}
        except ImportError:
            return {"ok": False, "error": "Excel export needs openpyxl. Run: pip install openpyxl"}
        except Exception as e:
            return {"ok": False, "error": str(e)}


def parse_iso(s):
    return dt.date.fromisoformat(s).strftime("%d %b %Y")


INDEX_HTML = r"""<!DOCTYPE html><html><head><meta charset="utf-8">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap">
<style>
/* Timesheets - design tokens. Dark ink sidebar, cool neutral workspace, green as an accent only. */
:root{--green:#6cc04a;--green-d:#3f8a23;--green-soft:#e9f6df;--ink:#15171c;--side:#0f1117;--side-2:#181b23;
--bg:#f3f4f7;--card:#ffffff;--mut:#6b7280;--mut-2:#9aa0ab;--line:#e6e8ec;--line-2:#eef0f3;
--amber:#d98c00;--amber-soft:#fff3dc;--red:#d4383a;--red-soft:#fde7e7;--blue:#2f6fd6;--blue-soft:#e4edfb;
--font:"Plus Jakarta Sans",-apple-system,"Segoe UI",Roboto,Arial,sans-serif;--mono:var(--font);
--r:12px;--r-sm:9px;--shadow:0 1px 2px rgba(16,20,30,.05),0 1px 1px rgba(16,20,30,.03);--shadow-lg:0 12px 40px rgba(16,20,30,.14);}
*{box-sizing:border-box;font-family:var(--font);}
html{font-size:14px}
body{margin:0;background:var(--bg);color:var(--ink);-webkit-font-smoothing:antialiased;line-height:1.45}
button{cursor:pointer;border:0;font-family:inherit;}
a{color:var(--green-d)}
.hidden{display:none!important;}
::selection{background:#d6efc5}
*::-webkit-scrollbar{width:10px;height:10px}*::-webkit-scrollbar-thumb{background:#c9ced6;border-radius:8px;border:3px solid transparent;background-clip:padding-box}*::-webkit-scrollbar-track{background:transparent}
.side ::-webkit-scrollbar-thumb{background:#2c3140;border:3px solid transparent;background-clip:padding-box}

/* ---------- sign in / sign up ---------- */
#login{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:var(--bg);background-image:radial-gradient(ellipse at 20% 0%,#e7f3dd 0,transparent 55%),radial-gradient(ellipse at 90% 100%,#e4e9f2 0,transparent 50%)}
.lcard{background:#fff;border-radius:18px;box-shadow:var(--shadow-lg);width:380px;overflow:hidden;border:1px solid var(--line)}
.lcard .bar{height:4px;background:linear-gradient(90deg,var(--green),#a8dc8a)}
.lcard .in{padding:36px 38px 32px}
.brand{font-weight:800;font-size:24px;color:var(--ink);letter-spacing:-.02em;line-height:1.1}
.brand-sub{font-size:10.5px;font-weight:700;color:var(--mut-2);letter-spacing:.18em;margin:6px 0 22px}
.brand-logo{display:none;max-width:160px;height:auto;margin-bottom:14px}
#setup{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:var(--bg);overflow:auto;padding:20px}
.lcard.setup{width:470px}
.lcard.setup .in{padding:30px 36px}
.s-lab{display:block;font-size:12px;font-weight:600;color:#3b4150;margin:12px 0 4px}
.s-choice label{display:flex;gap:8px;align-items:flex-start;font-size:13px;margin:4px 0;cursor:pointer}
.s-choice input{width:auto;margin:2px 0 0}
.s-hint{font-size:12px;color:var(--mut);margin-top:10px;line-height:1.5}
.set-grid{display:grid;grid-template-columns:230px 1fr;gap:14px 18px;align-items:center;max-width:680px}
.set-grid input[type=text],.set-grid input[type=number]{width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:var(--r-sm);font-size:13.5px;outline:none;margin:0}
.set-grid input:focus{border-color:var(--green);box-shadow:0 0 0 3px rgba(108,192,74,.18)}
.set-grid .k{font-size:13px;font-weight:600;color:#2b3040}
.set-grid .k small{display:block;font-weight:400;color:var(--mut);font-size:11.5px;margin-top:2px}
.brand-co{white-space:normal;line-height:1.1}
.rem{display:flex;align-items:center;gap:8px;font-size:12.5px;color:var(--mut);margin:4px 0 6px}
.rem input{width:auto;margin:0}
.err{color:var(--red);font-size:12.5px;min-height:18px;margin-top:6px}
input{width:100%;padding:11px 13px;border:1px solid var(--line);border-radius:var(--r-sm);font-size:14px;margin:6px 0;outline:none;background:#fff;color:var(--ink);transition:border-color .12s,box-shadow .12s}
input::placeholder{color:var(--mut-2)}
input:focus{border-color:var(--green);box-shadow:0 0 0 3px rgba(108,192,74,.18)}
select,textarea{font-family:inherit;color:var(--ink)}

/* ---------- buttons ---------- */
.btn{background:var(--green);color:#0f2a07;font-weight:700;padding:11px 16px;border-radius:var(--r-sm);width:100%;font-size:13.5px;letter-spacing:.01em;transition:background .15s,color .15s,box-shadow .15s,transform .1s}
.btn:hover{background:#5db13c;color:#0f2a07}
.btn:active{transform:translateY(1px)}
.btn.sm{width:auto;padding:8px 13px;font-size:12.5px}
.avatar{display:inline-flex;align-items:center;justify-content:center;border-radius:50%;font-weight:700;letter-spacing:.02em;flex:0 0 auto;line-height:1}
.person{display:inline-flex;align-items:center;gap:9px;white-space:nowrap}
.btn.xs{font-size:11px;padding:3px 9px;border-radius:7px;width:auto}
.btn.ghost{background:#fff;color:#2b3040;border:1px solid var(--line);box-shadow:var(--shadow)}
.btn.ghost:hover{background:#f7f8fa;border-color:#d5d9e0;color:var(--ink)}
.btn:focus-visible{outline:2px solid var(--green);outline-offset:2px}

/* ---------- shell ---------- */
#app{display:flex;height:100vh}
.side{width:248px;background:var(--side);color:#dfe3ea;display:flex;flex-direction:column;flex:0 0 248px;border-right:1px solid #1d212b}
.side .top{padding:22px 22px 10px}
.side .top .brand{color:#fff;font-size:19px}
.side .top .brand-sub{color:#6d7484;margin:4px 0 0;letter-spacing:.16em}
.side .brand-logo{max-width:150px;margin-bottom:6px}
.side .who{font-size:12px;color:#8e95a3;padding:8px 22px 14px;border-bottom:1px solid #1d212b;margin-bottom:6px}
.nav{flex:1;overflow:auto;padding:4px 12px 10px}
.nav .grp{font-size:10px;font-weight:700;letter-spacing:.14em;color:#5f6676;padding:14px 12px 5px;text-transform:uppercase}
.nav a{display:flex;align-items:center;gap:11px;padding:8px 12px;margin:1px 0;border-radius:8px;color:#b3bac6;font-size:13.5px;font-weight:500;text-decoration:none;position:relative;transition:background .12s,color .12s}
.nav a svg{width:17px;height:17px;flex:0 0 auto;opacity:.8}
.nav a:hover{background:var(--side-2);color:#fff}
.nav a.active{background:#1c2029;color:#fff;font-weight:600}
.nav a.active::before{content:"";position:absolute;left:-12px;top:7px;bottom:7px;width:3px;border-radius:0 3px 3px 0;background:var(--green)}
.nav a.active svg{opacity:1;color:var(--green)}
.dbinfo{display:none}
.signout{padding:11px 14px;margin:10px 12px 14px;border-radius:8px;background:transparent;color:#aab1be;text-align:left;font-size:13px;border:1px solid #262b36;font-weight:600}
.signout:hover{background:var(--side-2);color:#fff;border-color:#343a47}
.main{flex:1;overflow:auto;padding:30px 36px 48px;min-width:0}
h1{font-size:24px;font-weight:800;letter-spacing:-.02em;margin:0 0 4px;color:var(--ink)}
.h-accent,.hbar{display:none}
.sub{color:var(--mut);margin:2px 0 22px;font-size:13px}
.sec-h{font-weight:700;font-size:11px;letter-spacing:.12em;color:var(--mut-2);margin:0 0 10px;text-transform:uppercase}
.sec-t{font-weight:700;font-size:14px;margin:0 0 12px;letter-spacing:-.01em}

/* ---------- cards / kpis ---------- */
.grid{display:grid;gap:14px}
.kpis{grid-template-columns:repeat(auto-fit,minmax(190px,1fr))}
.kpis-4{grid-template-columns:repeat(auto-fit,minmax(158px,1fr))}
.card{background:var(--card);border-radius:var(--r);border:1px solid var(--line);box-shadow:var(--shadow);padding:18px 20px}
.kpi .lab{font-size:11px;color:var(--mut);font-weight:700;text-transform:uppercase;letter-spacing:.1em}
.kpi .val{font-size:27px;font-weight:800;margin-top:8px;color:var(--ink);font-variant-numeric:tabular-nums;letter-spacing:-.02em}
.kpi .val.green{color:var(--green-d)}
.two{display:grid;grid-template-columns:1.4fr 1fr;gap:14px;margin-top:14px}
.chart-empty{color:var(--mut);font-size:13px;padding:20px 0}

/* ---------- tables ---------- */
table{width:100%;border-collapse:collapse}
th{text-align:left;font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--mut);padding:9px 12px;border-bottom:1px solid var(--line);font-weight:700;white-space:nowrap}
td{padding:10px 12px;border-bottom:1px solid var(--line-2);font-size:13.5px;vertical-align:middle}
tbody tr:last-child td{border-bottom:0}
tr:hover td{background:#f8fafc}
tbody tr{transition:background .12s ease}
.tnum{text-align:right;font-variant-numeric:tabular-nums}
th.tnum{font-size:10.5px}
th.sortable{cursor:pointer;user-select:none}
th.sortable:hover{color:var(--ink)}
.sortarrow{font-size:9px;margin-left:4px;color:var(--green-d)}
.dayrow td{background:#f6f7f9;font-size:12.5px;font-weight:600;color:#3b4150}
tr.dayrow{cursor:default}
.lrow{cursor:pointer}
tr.lrow.sel td{background:var(--green-soft)}
tr.lrow.sel td:first-child{box-shadow:inset 3px 0 0 var(--green)}
tr.jobrow.dragover td{background:#f1f9ea;box-shadow:inset 0 0 0 1px var(--green)}
.xdel{background:none;color:var(--red);font-size:13px;padding:0 4px;border:0;cursor:pointer;opacity:.7}.xdel:hover{opacity:1}

/* ---------- badges / pills ---------- */
.badge{display:inline-flex;align-items:center;gap:6px;padding:3px 9px 3px 8px;border-radius:999px;font-size:11.5px;font-weight:700;letter-spacing:.01em;white-space:nowrap}
.badge::before{content:"";width:6px;height:6px;border-radius:50%;background:currentColor;opacity:.85}
.b-Approved,.b-Paid,.b-Active{background:var(--green-soft);color:#2e6a16}
.b-Rejected{background:var(--red-soft);color:#a62628}
.b-Submitted,.b-Raised,.b-Received{background:var(--amber-soft);color:#8a5a00}
.b-Draft,.b-Logged,.b-Closed{background:#eef0f3;color:#4b5160}
.b-Locked{background:var(--blue-soft);color:#24559e}
.badge.mini{font-size:10.5px;padding:1px 7px;vertical-align:1px}
.pill{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;letter-spacing:.04em}.pill-adj{background:var(--green-soft);color:#2e6a16}.pill-sub{background:var(--blue-soft);color:#24559e}
.muted{color:var(--mut)}
.toolbar{display:flex;gap:8px;margin:16px 0 0;align-items:center;flex-wrap:wrap}
.toolbar select,.rep-grid select,.rep-grid input,.modal-b select,.modal-b textarea,.modal-b input{padding:9px 11px;border:1px solid var(--line);border-radius:var(--r-sm);font-size:13.5px;outline:none;background:#fff}
.toolbar select:focus,.rep-grid select:focus,.rep-grid input:focus,.modal-b select:focus,.modal-b textarea:focus,.modal-b input:focus{border-color:var(--green);box-shadow:0 0 0 3px rgba(108,192,74,.18)}
.rep-grid{display:grid;grid-template-columns:150px 1fr;gap:10px 14px;align-items:center;max-width:600px}
.bar-wrap{background:#eceef2;border-radius:999px;height:8px;overflow:hidden;width:120px;display:inline-block;vertical-align:middle}
.bar-fill{height:100%;background:var(--green)}
.bar-fill.warn{background:var(--amber)}.bar-fill.over{background:var(--red)}

/* ---------- dashboard ---------- */
.att-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(258px,1fr));gap:12px;margin-bottom:22px;align-items:stretch}
.att{background:var(--card);border-radius:var(--r);padding:16px 18px;border:1px solid var(--line);box-shadow:var(--shadow);cursor:pointer;display:flex;gap:13px;align-items:flex-start;transition:border-color .12s,box-shadow .15s}
.att:hover{border-color:#cfd4dc;box-shadow:0 6px 18px rgba(16,20,30,.08)}
.att.warn,.att.bad,.att.ok{border-left:4px solid var(--amber)}.att.bad{border-left-color:var(--red)}.att.ok{border-left-color:var(--green)}
.att-ico{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex:0 0 auto;background:#eef0f3;color:#4b5160}
.att.warn .att-ico{background:var(--amber-soft);color:#8a5a00}
.att.bad .att-ico{background:var(--red-soft);color:#a62628}
.att.ok .att-ico{background:var(--green-soft);color:#2e6a16}
.att-ico svg{width:18px;height:18px}
.att-big{font-size:19px;font-weight:800;color:var(--ink);line-height:1.15;letter-spacing:-.01em}
.att-lab{font-size:12.5px;color:#4b5160;margin-top:3px;line-height:1.4;overflow-wrap:anywhere}
.att-go{font-size:11.5px;font-weight:700;color:var(--green-d);margin-top:7px}
.jrow{display:flex;align-items:center;gap:10px;padding:9px 2px;border-bottom:1px solid var(--line-2);font-size:13px;cursor:pointer;border-radius:6px}
.jrow:last-of-type{border-bottom:none}
.jrow:hover{background:#f8fafc}
.jref{flex:0 0 148px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:600}
.jtitle{flex:0 0 175px;color:var(--mut);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.jbar{flex:1 1 auto;height:7px;border-radius:5px;background:#eceef2;overflow:hidden;min-width:80px}
.jpct{flex:0 0 52px;text-align:right;font-weight:700;font-variant-numeric:tabular-nums;font-size:12.5px}
.jbar i{display:block;height:100%;border-radius:5px}
.kpi-doc{cursor:pointer;position:relative;transition:border-color .12s,box-shadow .15s}
.kpi-doc:hover{border-color:#cfd4dc;box-shadow:0 6px 18px rgba(16,20,30,.08)}
.kpi-clip{font-size:12px}
.kpi-open{font-size:11px;color:var(--green-d);font-weight:600;margin-top:2px}
.rankrow{display:flex;align-items:center;gap:12px;padding:9px 4px;border-bottom:1px solid var(--line-2)}
.rankrow:last-child{border-bottom:none}
.rankno{width:26px;height:26px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:700;color:#2e6a16;background:var(--green-soft);font-size:12px;flex:0 0 auto}
.rankbar{height:7px;border-radius:5px;background:#eceef2;margin-top:5px;overflow:hidden}
.rankbar-f{height:100%;background:var(--green)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:4px}@media(max-width:900px){.grid2{grid-template-columns:1fr}}
.stk{height:10px;border-radius:6px;background:#f0c9c9;overflow:hidden}.stk-b{height:100%;background:var(--green)}

/* ---------- folders / documents ---------- */
.foldsearch{width:100%;max-width:320px;padding:8px 12px;border:1px solid var(--line);border-radius:var(--r-sm);font-size:13px;font-family:inherit;outline:none;margin-bottom:12px}
.foldsearch:focus{border-color:var(--green);box-shadow:0 0 0 3px rgba(108,192,74,.18)}
.fold-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(216px,1fr));gap:10px}
.fold{display:flex;gap:11px;align-items:center;padding:11px 12px;border:1px solid var(--line);border-radius:10px;cursor:pointer;background:#fff;transition:border-color .12s,box-shadow .15s}
.fold:hover{border-color:#cfd4dc;box-shadow:0 6px 18px rgba(16,20,30,.08)}
.fold.over{border-color:var(--green);border-style:dashed;background:#f4faec;box-shadow:0 0 0 3px rgba(108,192,74,.2)}
.fold-ico{position:relative;flex:0 0 auto;width:38px;height:38px;display:flex;align-items:center;justify-content:center}
.fold-ico svg{width:36px;height:36px}
.fold-ico .fbody{color:#f0a92b}.fold-ico .ftab{color:#f8cf7a}
.fold:hover .fold-ico .fbody,.fold.over .fold-ico .fbody{color:var(--green)}
.fold:hover .fold-ico .ftab,.fold.over .fold-ico .ftab{color:#b7dd8b}
.fold-n{position:absolute;right:-6px;bottom:-3px;background:var(--ink);color:#fff;font-size:9.5px;font-weight:700;line-height:1;border-radius:999px;min-width:17px;height:17px;display:flex;align-items:center;justify-content:center;padding:0 4px;border:2px solid #fff}
.fold-ref{font-weight:700;font-size:13px;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.fold-title{font-size:11.5px;color:var(--mut);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.fold-meta{font-size:10.5px;color:var(--mut);margin-top:2px}
.crumbs{font-size:12px;color:var(--mut);margin-bottom:8px;display:flex;gap:7px;align-items:center}
.crumbs a{color:var(--green-d);font-weight:600;cursor:pointer}.crumbs a:hover{text-decoration:underline}
.doc-filters{display:flex;gap:6px;flex-wrap:wrap;margin:2px 0 14px}
.doc-chip{border:1px solid var(--line);background:#fff;color:#4b5160;border-radius:999px;padding:5px 13px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;white-space:nowrap;width:auto;transition:border-color .12s,color .12s,background .12s}
.doc-chip:hover{border-color:#cfd4dc;color:var(--ink)}
.doc-chip.on{background:var(--ink);border-color:var(--ink);color:#fff}
.doc-list{display:flex;flex-direction:column}
.doc-row{display:flex;align-items:center;gap:12px;padding:9px 10px;border-radius:10px}
.doc-row+.doc-row{border-top:1px solid var(--line-2)}
.doc-row:hover{background:#f8fafc}
.doc-ico{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;flex:0 0 auto}
.doc-ico svg{width:18px;height:18px}
.di-Quote,.dt-Quote{background:var(--green-soft);color:#2e6a16}
.di-Variation,.dt-Variation{background:var(--amber-soft);color:#8a5a00}
.di-Invoice,.dt-Invoice{background:var(--blue-soft);color:#24559e}
.di-Bill,.dt-Bill{background:#f3e4f1;color:#7d2a66}
.di-Document,.dt-Document{background:#eef0f3;color:#4b5160}
.di-gone{background:var(--red-soft);color:#a62628}
.doc-main{min-width:0;flex:1}
.doc-name{font-size:13.5px;font-weight:600;color:var(--ink);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.doc-meta{font-size:11.5px;color:var(--mut);margin-top:3px;display:flex;gap:7px;flex-wrap:wrap;align-items:center}
.doc-tag{padding:1px 8px;border-radius:999px;font-size:10.5px;font-weight:700}
.doc-act{display:flex;gap:6px;flex:0 0 auto}
.doc-row.gone .doc-name{color:var(--red)}
.doc-empty{color:var(--mut);font-size:13px;padding:18px 4px;text-align:center}
.dropzone{border:2px dashed var(--line);border-radius:10px;padding:12px;text-align:center;color:var(--mut);font-size:12.5px;margin:8px 0 4px;transition:border-color .15s,background .15s,color .15s}
.dropzone.over{border-color:var(--green);background:#f4faec;color:var(--green-d);font-weight:600}
.page-drop{position:fixed;inset:14px;border:3px dashed var(--green);border-radius:16px;background:rgba(108,192,74,.12);z-index:60;display:flex;align-items:center;justify-content:center;font-size:19px;font-weight:700;color:#2e6a16;pointer-events:none}

/* ---------- context menu / modal / toast / idle ---------- */
.ctxmenu{position:fixed;z-index:90;background:#fff;border:1px solid var(--line);border-radius:11px;box-shadow:var(--shadow-lg);padding:6px;min-width:232px}
.ctxmenu button{display:block;width:100%;text-align:left;background:none;border:0;padding:8px 12px;border-radius:7px;cursor:pointer;font-family:inherit;font-size:13px;color:#2b3040;white-space:nowrap}
.ctxmenu button:hover:not(:disabled){background:#f3f4f7;color:var(--ink)}
.ctxmenu button:disabled{opacity:.38;cursor:default}
.ctxmenu .k{float:right;opacity:.45;font-size:11px;margin-left:18px}
.ctxmenu .sep{height:1px;background:var(--line);margin:5px 4px}
.ctxmenu .hint{font-size:10.5px;font-weight:800;letter-spacing:.05em;color:var(--mut);padding:5px 12px 6px}
.modal-bg{position:fixed;inset:0;background:rgba(15,17,23,.5);display:flex;align-items:center;justify-content:center;z-index:50;backdrop-filter:blur(2px)}
.modal{background:#fff;border-radius:14px;width:400px;max-width:92vw;box-shadow:var(--shadow-lg);overflow:hidden;border:1px solid var(--line)}
.modal-h{padding:16px 22px;font-weight:700;font-size:15px;border-bottom:1px solid var(--line)}
.modal-b{padding:14px 22px 8px;max-height:70vh;overflow:auto}
.modal-b label{display:block;font-size:12px;color:#3b4150;font-weight:600;margin:10px 0 4px}
.modal-b select,.modal-b textarea,.modal-b input{width:100%;margin:0}
.modal-f{display:flex;justify-content:flex-end;gap:8px;padding:14px 22px;border-top:1px solid var(--line);background:#fafbfc}
.toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%) translateY(20px);background:var(--ink);color:#fff;padding:10px 18px;border-radius:10px;font-size:13px;opacity:0;pointer-events:none;transition:opacity .2s,transform .2s;z-index:9999;box-shadow:var(--shadow-lg)}.toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
.idle-warn{position:fixed;bottom:20px;right:20px;z-index:80;background:#fff;border:1px solid var(--line);border-left:4px solid var(--amber);border-radius:12px;padding:14px 16px;box-shadow:var(--shadow-lg);display:flex;align-items:center;gap:14px;font-size:13.5px;max-width:420px}

/* ---------- timesheet / timer / timeline ---------- */
.ts-top{display:flex;align-items:center;gap:10px;margin-bottom:10px}
.wk-strip{display:flex;gap:6px;margin:0 0 14px;flex-wrap:wrap}
.wk{display:inline-flex;align-items:center;gap:7px;padding:6px 11px;border-radius:999px;border:1px solid var(--line);background:#fff;font-size:12.5px;font-weight:600;color:#4b5160;cursor:pointer;transition:border-color .15s,background .15s,color .15s,transform .15s cubic-bezier(.22,1,.36,1)}
.wk:hover{border-color:#cfd4dc;transform:translateY(-1px)}
.wk.on{background:var(--ink);border-color:var(--ink);color:#fff}
.wk.now:not(.on){border-color:var(--green)}
.wk-d{width:7px;height:7px;border-radius:50%;background:#d5d9e0}
.wk.s-Approved .wk-d,.wk.s-Locked .wk-d{background:var(--green)}
.wk.s-Submitted .wk-d{background:var(--amber)}
.wk.s-Rejected .wk-d{background:var(--red)}
.wk.s-Draft .wk-d{background:#9aa0ab}
.wk.on .wk-d{box-shadow:0 0 0 2px rgba(255,255,255,.35)}
.ts-week{font-weight:700;color:var(--ink);font-size:15px}
.ts-foot{display:flex;align-items:center;margin-top:14px;gap:8px}
.timer-bar{display:flex;align-items:center;gap:12px;background:#fff;border:1px solid var(--line);border-radius:10px;padding:8px 14px;margin-bottom:12px;box-shadow:var(--shadow)}
.timer-lbl{font-size:10.5px;font-weight:800;letter-spacing:.12em;color:var(--green-d)}
.timer-disp{font-variant-numeric:tabular-nums;font-size:20px;font-weight:700;color:var(--ink);min-width:98px;letter-spacing:.02em}
.day-actions{float:right;opacity:1;display:inline-flex;gap:6px}
.tl{border-collapse:separate;border-spacing:2px;font-size:12.5px}
.tl th{font-weight:700;color:var(--mut);padding:4px 6px;text-align:center;font-size:10.5px;border:0}
.tl th.tl-job{text-align:left;min-width:150px}
.tl-job{text-align:left;white-space:nowrap;padding-right:12px}
.tl-sub{font-size:10.5px;color:var(--mut-2);font-weight:400}
.tlc{width:40px;text-align:center;border-radius:6px;padding:6px 2px;font-variant-numeric:tabular-nums;font-size:12.5px}
.tl-tot{background:#eef0f3;font-weight:700}
.tlc-click{cursor:pointer}.tlc-click:hover{outline:2px solid var(--green)}
.tl-row{cursor:pointer}.tl-row:hover td.tl-job{color:var(--green-d)}.tl-row:hover .tlc{outline:1px solid rgba(108,192,74,.5)}

/* ---------- date range picker ---------- */
.dateinput{cursor:pointer;background:#fff url("data:image/svg+xml;charset=utf8,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%236b7280' stroke-width='2' stroke-linecap='round'%3E%3Crect x='3' y='4' width='18' height='18' rx='2'/%3E%3Cpath d='M16 2v4M8 2v4M3 10h18'/%3E%3C/svg%3E") no-repeat right 10px center}
.drp-bg{position:fixed;inset:0;background:rgba(15,17,23,.5);display:flex;align-items:center;justify-content:center;z-index:70}
.drp{background:#fff;border-radius:14px;box-shadow:var(--shadow-lg);display:flex;overflow:hidden;max-width:94vw;border:1px solid var(--line)}
.drp-side{background:#f7f8fa;border-right:1px solid var(--line);padding:14px 10px;display:flex;flex-direction:column;gap:4px;min-width:158px}
.drp-side .lab{font-size:10.5px;font-weight:800;letter-spacing:.06em;color:var(--mut);padding:2px 10px 6px}
.drp-side button{background:none;border:0;text-align:left;padding:7px 10px;border-radius:8px;font-size:13px;color:#2b3040;cursor:pointer;font-family:inherit}
.drp-side button:hover{background:#e9edf2;color:var(--ink)}
.drp-main{padding:14px 16px}
.drp-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;gap:10px}
.drp-nav{background:none;border:1px solid var(--line);border-radius:8px;width:30px;height:28px;cursor:pointer;font-size:14px;color:#2b3040;font-family:inherit}
.drp-nav:hover{border-color:var(--green);color:var(--green-d)}
.drp-months{display:flex;gap:20px}
.drp-cal{width:228px}
.drp-cal .mname{text-align:center;font-weight:700;font-size:13px;margin-bottom:6px;color:#2b3040}
.drp-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px}
.drp-grid .wd{font-size:10.5px;color:var(--mut);text-align:center;font-weight:700;padding:2px 0}
.drp-d{height:29px;display:flex;align-items:center;justify-content:center;font-size:12.5px;border-radius:7px;cursor:pointer;color:#2b3040;user-select:none}
.drp-d:hover{background:#e8f3d6}
.drp-d.out{color:#ccd0d6;cursor:default}.drp-d.out:hover{background:none}
.drp-d.in{background:var(--green-soft);border-radius:2px}
.drp-d.sel{background:var(--green);color:#0f2a07;font-weight:700}
.drp-d.today{box-shadow:inset 0 0 0 1.5px var(--green-d)}
.drp-foot{display:flex;align-items:center;gap:10px;margin-top:12px;padding-top:12px;border-top:1px solid var(--line)}
.drp-sel{font-size:13px;color:#2b3040;font-weight:600}

/* ---------- analytics ---------- */
.an-top{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin:0 0 18px}
.seg{display:inline-flex;background:#fff;border:1px solid var(--line);border-radius:10px;padding:3px;gap:2px}
.seg button{background:transparent;border:0;padding:6px 12px;border-radius:7px;font-size:12.5px;font-weight:600;color:#4b5160;cursor:pointer;transition:background .15s,color .15s}
.seg button:hover{background:#f3f4f7;color:var(--ink)}
.seg button.on{background:var(--ink);color:#fff}
.an-range{display:flex;align-items:center;gap:8px;margin-left:auto}
.an-range input{width:132px;margin:0;padding:7px 30px 7px 10px;font-size:12.5px}
.goal-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:14px}
.goal{display:flex;gap:14px;align-items:center;padding:16px 18px;border-left:4px solid var(--green)}
.goal.warn{border-left-color:var(--amber)}.goal.bad{border-left-color:var(--red)}
.goal .ring{flex:0 0 64px}
.goal-b{min-width:0}
.goal .lab{font-size:11px;color:var(--mut);font-weight:700;text-transform:uppercase;letter-spacing:.1em}
.goal .val{font-size:24px;font-weight:800;letter-spacing:-.02em;line-height:1.15;margin-top:3px;font-variant-numeric:tabular-nums}
.goal-t{font-size:12.5px;color:#3b4150;margin-top:3px}
.goal-s{font-size:11.5px;color:var(--mut);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ring circle:nth-child(2){stroke-dashoffset:var(--c);animation:ringIn 1s .15s cubic-bezier(.22,1,.36,1) forwards}
@keyframes ringIn{to{stroke-dashoffset:var(--o)}}
.legend{margin-left:auto;display:inline-flex;align-items:center;gap:6px;font-size:11.5px;font-weight:500;color:var(--mut)}
.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-left:10px}
.hbars{display:flex;flex-direction:column;gap:7px}
.hrow{display:grid;grid-template-columns:180px 1fr 52px 64px;gap:10px;align-items:center;font-size:13px}
.hrow .hname{min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hrow .hb{height:9px;background:#eceef2;border-radius:5px;overflow:hidden}
.hrow .hb i{display:block;height:100%;border-radius:5px;transform-origin:left center;animation:growx .7s cubic-bezier(.22,1,.36,1) both}
.hrow .hv{text-align:right;font-weight:700;font-variant-numeric:tabular-nums}
.hrow .hh{text-align:right;font-size:11.5px;font-variant-numeric:tabular-nums;white-space:nowrap}
.an-details summary{display:flex;align-items:center;gap:12px;cursor:pointer;list-style:none;flex-wrap:wrap}
.an-details summary::-webkit-details-marker{display:none}
.an-details summary::before{content:"▸";color:var(--mut);font-size:12px;transition:transform .2s}
.an-details[open] summary::before{transform:rotate(90deg)}
/* ---------- charts ---------- */
svg.bars .bar{transform-origin:bottom;animation:growy .6s cubic-bezier(.22,1,.36,1) both}
svg.bars .bar rect{transition:fill .15s ease}
svg.bars .bar:hover rect:first-child{fill:var(--green)}
svg.bars .trend{stroke-dasharray:1000;stroke-dashoffset:1000;animation:draw 1.1s .3s ease-out forwards}
@keyframes draw{to{stroke-dashoffset:0}}
/* ---------- motion ---------- */
.nowrap{white-space:nowrap}
.main{transition:opacity .16s ease}
.main.leaving{opacity:.45}
.main.enter>*{animation:fadeUp .38s cubic-bezier(.22,1,.36,1) both}
.main.enter>*:nth-child(2){animation-delay:.04s}.main.enter>*:nth-child(3){animation-delay:.08s}.main.enter>*:nth-child(4){animation-delay:.12s}.main.enter>*:nth-child(5){animation-delay:.16s}.main.enter>*:nth-child(6){animation-delay:.2s}.main.enter>*:nth-child(7){animation-delay:.24s}.main.enter>*:nth-child(n+8){animation-delay:.28s}
.main.enter .att-grid>*,.main.enter .kpis>*,.main.enter .fold-grid>*{animation:fadeUp .42s cubic-bezier(.22,1,.36,1) both}
.main.enter .att-grid>*:nth-child(2),.main.enter .kpis>*:nth-child(2),.main.enter .fold-grid>*:nth-child(2){animation-delay:.06s}
.main.enter .att-grid>*:nth-child(3),.main.enter .kpis>*:nth-child(3),.main.enter .fold-grid>*:nth-child(3){animation-delay:.12s}
.main.enter .att-grid>*:nth-child(4),.main.enter .kpis>*:nth-child(4),.main.enter .fold-grid>*:nth-child(4){animation-delay:.18s}
.main.enter .att-grid>*:nth-child(n+5),.main.enter .kpis>*:nth-child(n+5),.main.enter .fold-grid>*:nth-child(n+5){animation-delay:.24s}
.main.enter tbody tr{animation:fadeIn .35s ease both}
.main.enter tbody tr:nth-child(2){animation-delay:.03s}.main.enter tbody tr:nth-child(3){animation-delay:.06s}.main.enter tbody tr:nth-child(4){animation-delay:.09s}.main.enter tbody tr:nth-child(5){animation-delay:.12s}.main.enter tbody tr:nth-child(6){animation-delay:.15s}.main.enter tbody tr:nth-child(n+7){animation-delay:.18s}
.nav a{transition:background .18s ease,color .18s ease,transform .18s cubic-bezier(.22,1,.36,1)}
.nav a:hover{transform:translateX(2px)}
.nav a.active::before{animation:growy .28s cubic-bezier(.22,1,.36,1) both;transform-origin:center}
@keyframes growy{from{transform:scaleY(0)}to{transform:scaleY(1)}}
.nav{scrollbar-width:none}.nav::-webkit-scrollbar{width:0;height:0}
.btn{transition:background .18s ease,color .18s ease,box-shadow .22s ease,transform .18s cubic-bezier(.34,1.56,.64,1)}
.btn:hover{transform:translateY(-1px);box-shadow:0 6px 16px rgba(108,192,74,.32)}
.btn.ghost:hover{box-shadow:0 4px 12px rgba(16,20,30,.08)}
.btn:active{transform:translateY(0) scale(.98);box-shadow:none;transition-duration:.08s}
.card{transition:box-shadow .28s ease,border-color .28s ease}
.card:hover{box-shadow:0 6px 22px rgba(16,20,30,.07)}
.att{transition:transform .24s cubic-bezier(.22,1,.36,1),box-shadow .24s ease,border-color .24s ease}
.att:hover{transform:translateY(-2px)}
tr:hover td{transition:background .15s ease}
.badge{transition:transform .15s ease}.badge:hover{transform:scale(1.04)}
.jbar i,.rankbar-f,.bar-fill,.stk-b{transform-origin:left center;animation:growx .75s cubic-bezier(.22,1,.36,1) both}
.jrow{transition:background .15s ease,padding-left .18s ease}.jrow:hover{padding-left:6px}
.fold{transition:transform .2s cubic-bezier(.22,1,.36,1),border-color .2s ease,box-shadow .2s ease}.fold:hover{transform:translateY(-2px)}
input,select,textarea{transition:border-color .15s ease,box-shadow .15s ease}
.signout{transition:background .18s ease,color .18s ease,border-color .18s ease}
.toast{transition:opacity .25s cubic-bezier(.22,1,.36,1),transform .25s cubic-bezier(.22,1,.36,1)}
.modal{animation:popIn .3s cubic-bezier(.34,1.3,.64,1)}
#login .lcard{animation:popIn .5s cubic-bezier(.22,1,.36,1)}
.doc-chip{transition:border-color .15s,color .15s,background .15s,transform .15s}.doc-chip:hover{transform:translateY(-1px)}
.drp-d{transition:background .12s ease,transform .12s ease}.drp-d:hover{transform:scale(1.08)}
/* (base keyframes) */
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes fadeUp{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
@keyframes growx{from{transform:scaleX(0)}to{transform:scaleX(1)}}
@keyframes popIn{from{opacity:0;transform:scale(.98)}to{opacity:1;transform:scale(1)}}
.modal-bg{animation:fadeIn .16s ease}.modal{animation:popIn .2s ease}.lcard{animation:popIn .3s ease}
@keyframes popIn{from{opacity:0;transform:scale(.97) translateY(4px)}to{opacity:1;transform:none}}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style></head>
<body class="{{LOGOCLS}}">
<div id="login"><div class="lcard"><div class="bar"></div><div class="in">
  <img class="brand-logo lg" src="{{LOGO}}" alt="">
  <div class="brand brand-fallback brand-co">Timesheets</div><div class="brand-sub brand-fallback brand-app">TIMESHEETS</div>
  <input id="le" placeholder="Username (the part before @) or email" autofocus>
  <input id="lp" type="password" placeholder="Password">
  <label class="rem"><input type="checkbox" id="lrem"> Keep me signed in on this PC</label>
  <div class="err" id="lerr" style="color:#7a8089">Starting up…</div>
  <button class="btn" id="lbtn">Sign in</button>
</div></div></div>

<div id="setup" class="hidden"><div class="lcard setup"><div class="bar"></div><div class="in">
  <div class="brand">Welcome</div><div class="brand-sub">SET UP TIMESHEETS</div>
  <div class="s-step" id="s1">
    <label class="muted s-lab">Company name</label><input id="su_co" placeholder="e.g. Northgate Electrical Ltd">
    <label class="muted s-lab">Your name</label><input id="su_name" placeholder="Full name">
    <label class="muted s-lab">Your email (this is your login)</label><input id="su_email" placeholder="you@company.com">
    <label class="muted s-lab">Password</label><input id="su_pw" type="password" placeholder="At least 6 characters">
    <label class="muted s-lab">Where should the data live?</label>
    <div class="s-choice"><label><input type="radio" name="su_where" value="local" checked> On this PC only</label>
      <label><input type="radio" name="su_where" value="shared"> In a shared folder, so the whole team uses one database</label></div>
    <div id="su_folderrow" class="hidden" style="display:flex;gap:8px;align-items:center;margin-top:6px"><input id="su_folder" placeholder="No folder chosen" readonly style="flex:1"><button class="btn sm ghost" type="button" onclick="setupPickFolder()">Choose…</button></div>
    <div class="s-hint" id="su_hint">Settings, database and receipts are kept in your AppData folder. You can move them to a shared folder later in Settings.</div>
    <div class="err" id="serr"></div>
    <button class="btn" id="sbtn" onclick="doSetup()">Create company &amp; sign in</button>
  </div>
</div></div></div>

<div id="app" class="hidden">
  <div class="side">
    <div class="top"><img class="brand-logo" src="{{LOGO}}" alt="">
      <div class="brand brand-fallback brand-co" style="font-size:20px">Timesheets</div>
      <div class="brand-sub brand-fallback brand-app" style="margin:0">TIMESHEETS</div></div>
    <div class="who" id="who"></div>
    <div class="nav" id="nav"></div>
    <div class="dbinfo" id="dbinfo"></div>
    <button class="signout" onclick="doLogout()">Sign out</button>
  </div>
  <div class="main" id="main"></div>
</div>

<div id="modal" class="hidden modal-bg"><div class="modal">
  <div class="modal-h" id="modalTitle"></div>
  <div class="modal-b" id="modalBody"></div>
  <div class="modal-f"><button class="btn ghost sm" id="modalCancel" onclick="closeModal()">Cancel</button>
    <button class="btn sm" id="modalOk">Save</button></div>
</div></div>

<script>
function _errBar(msg){try{var d=document.createElement('div');d.style.cssText='position:fixed;top:0;left:0;right:0;background:#b22;color:#fff;padding:8px 14px;font:13px sans-serif;z-index:99999;cursor:pointer';d.textContent=msg;d.title='Click to dismiss';d.onclick=function(){d.remove();};document.body.appendChild(d);}catch(e){}}
window.onerror=function(m,src,line){_errBar('App error: '+m+' (line '+line+')');};
// Anything awaited inside an async function (i.e. every API call) rejects into
// here, NOT into window.onerror. Without this a failing API call makes a button
// look completely dead instead of reporting the problem.
// A file dropped anywhere we're NOT expecting one would otherwise make the
// webview navigate to it, blanking the whole app. Swallow those.
['dragover','drop'].forEach(function(ev){
  window.addEventListener(ev,function(e){e.preventDefault();},false);
});
window.addEventListener('unhandledrejection',function(ev){
  var r=ev.reason; var m=(r&&(r.message||r.toString&&r.toString()))||'unknown error';
  _errBar('Action failed: '+m);
});
let API=null, USER=null;
const E=(id)=>document.getElementById(id);
function esc(s){return (s==null?'':String(s)).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));}
const _AVC=['#6cc04a','#2f6fd6','#d98c00','#9b59b6','#e2575b','#15a3a3','#7a8a2b','#d2477a','#4f6bed','#b36b00'];
function avatar(name,size){const n=String(name||'').trim();const parts=n.split(/\s+/);const ini=((parts[0]||'')[0]||'')+((parts[1]||'')[0]||'');let h=0;for(const ch of n)h=(h*31+ch.charCodeAt(0))>>>0;const c=_AVC[h%_AVC.length];const sz=size||26;
  return `<span class="avatar" style="width:${sz}px;height:${sz}px;font-size:${Math.round(sz*0.4)}px;background:${c}1f;color:${c}" title="${esc(n)}">${esc(ini.toUpperCase())}</span>`;}
function person(name,size){return `<span class="person">${avatar(name,size)}<span>${esc(name)}</span></span>`;}
function badge(s){return `<span class="badge b-${esc(s)}">${esc(s)}</span>`;}
function _isoUTC(iso){const p=String(iso).slice(0,10).split('-').map(Number);return Date.UTC(p[0],p[1]-1,p[2]);}
function shiftISO(iso,days){const d=new Date(_isoUTC(iso));d.setUTCDate(d.getUTCDate()+days);return d.toISOString().slice(0,10);}
function diffDaysISO(a,b){return Math.round((_isoUTC(b)-_isoUTC(a))/86400000);}
function toast(msg){let t=document.getElementById('_toast');if(!t){t=document.createElement('div');t.id='_toast';t.className='toast';document.body.appendChild(t);}t.textContent=msg;t.classList.add('show');clearTimeout(window._tt);window._tt=setTimeout(()=>t.classList.remove('show'),1800);}

function ensureApi(){
  const a = window.pywebview && window.pywebview.api;
  // Only latch on once the Python methods are actually attached - the api
  // object can appear a moment before its methods are populated.
  if(a && typeof a.login === 'function'){ API = a; }
  return API;
}
async function prefillLogin(){
  if(window._prefilled)return;
  // Bridge methods attach one by one - wait until this one actually exists,
  // and only mark done once it has succeeded (the timer below retries).
  if(!API||typeof API.startup_prefs!=='function')return;
  window._prefilled=true;
  try{const p=await API.startup_prefs();
    setBootMsg('');
    if(p&&p.username&&E('le')&&!E('le').value){E('le').value=p.username;const r=E('lrem');if(r)r.checked=true;const pw=E('lp');if(pw)pw.focus();}
  }catch(e){}
}
function setBootMsg(t){const e=E('lerr'); if(e){e.style.color=t?'#7a8089':''; e.textContent=t||'';}}
async function applyBrand(){
  try{const b=await API.brand(); window._brand=b;
    const co=(b.company||'').trim();
    document.querySelectorAll('.brand-co').forEach(e=>e.textContent=co||b.app);
    document.querySelectorAll('.brand-app').forEach(e=>e.textContent=co?b.app.toUpperCase():'');
    document.querySelectorAll('.brand-logo').forEach(e=>{e.src=b.logo||'';});
    document.body.classList.toggle('haslogo',!!b.logo);
    return b;}catch(e){return null;}
}
async function bootAuth(){
  if(window._booted)return;
  if(!API||typeof API.auto_login!=='function'||typeof API.brand!=='function')return;
  window._booted=true;
  const b=await applyBrand();
  if(b&&b.needs_setup){ E('login').classList.add('hidden'); E('setup').classList.remove('hidden'); setTimeout(()=>{const c=E('su_co');if(c)c.focus();},50); return; }
  setBootMsg('Signing you in\u2026');
  let u=null; try{u=await API.auto_login();}catch(e){}
  if(u){enterApp(u);return;}
  setBootMsg('');
  prefillLogin();
}
window.addEventListener('pywebviewready', ()=>{ ensureApi(); bootAuth(); });
ensureApi(); bootAuth();
let _apiTries=0; const _apiTimer=setInterval(()=>{ ensureApi(); bootAuth();
  if((API&&window._booted)||_apiTries++>600){ clearInterval(_apiTimer); if(!API) setBootMsg('Could not start the app engine. Close this window and open Timesheets again.'); } }, 100);
E('lbtn').onclick=doLogin;
document.addEventListener('keydown',e=>{ if(e.key==='Enter' && !E('login').classList.contains('hidden')) doLogin(); });

function waitForApi(ms){
  return new Promise(resolve=>{
    if(ensureApi()){resolve(true);return;}
    let waited=0;
    const t=setInterval(()=>{
      if(ensureApi()){clearInterval(t);resolve(true);}
      else if((waited+=100)>=ms){clearInterval(t);resolve(false);}
    },100);
  });
}
async function doLogin(){
  const btn=E('lbtn'); if(btn&&btn.dataset.busy==='1')return;
  const email=E('le').value.trim(), pw=E('lp').value;
  E('lerr').textContent='';
  // Wait (up to ~12s) for the database bridge to attach, then log in
  // automatically - no more "click again".
  if(!ensureApi()){
    if(btn){btn.dataset.busy='1';btn.textContent='Connecting\u2026';btn.disabled=true;}
    const ready=await waitForApi(45000);
    if(btn){btn.dataset.busy='';btn.textContent='Sign in';btn.disabled=false;}
    if(!ready){E('lerr').textContent='Still starting up - this can take a little longer the first time it is opened. Wait a few seconds and click Sign in again.';return;}
  }
  let u=null;
  try{ u=await API.login(email, pw); }
  catch(e){ E('lerr').textContent='Login error: '+((e&&e.message)?e.message:e); return; }
  if(!u){ E('lerr').textContent='Invalid email or password, or account disabled.'; return; }
  try{API.remember_login(!!(E('lrem')&&E('lrem').checked), email);}catch(e){}
  enterApp(u);
}
async function enterApp(u){
  USER=u; E('login').classList.add('hidden'); E('app').classList.remove('hidden');
  E('who').textContent=`${u.full_name} (${u.user_type})`;
  try{const c=await API.app_config(); window._dailyWarn=c.daily_warning; _idleMs=(c.timeout_minutes||20)*60000; const dbi=E('dbinfo'); if(dbi&&c.db_path) dbi.textContent='Data file: '+c.db_path;}catch(e){}
  resetIdle(); buildNav(); nav('dashboard');
}
async function doLogout(){ try{await API.forget_login();}catch(e){} await API.logout(); USER=null; window._booted=true; E('app').classList.add('hidden'); E('login').classList.remove('hidden'); E('lp').value=''; prefillLogin(); }

// ---- first-run setup ----
document.querySelectorAll('input[name=su_where]').forEach(r=>r.addEventListener('change',()=>{
  const shared=document.querySelector('input[name=su_where]:checked').value==='shared';
  E('su_folderrow').classList.toggle('hidden',!shared);
  E('su_hint').textContent=shared?'Pick a folder everyone can reach (a network drive or shared folder). If it already holds a Timesheets database, this PC joins it and no new company is created.':'Settings, database and receipts are kept in your AppData folder. You can move them to a shared folder later in Settings.';
}));
async function setupPickFolder(){ const r=await API.pick_data_folder(); if(r&&r.ok){E('su_folder').value=r.folder; if(r.has_db){E('su_hint').textContent='That folder already has a Timesheets database - this PC will join it. Sign in with an account that exists there.';}} }
async function doSetup(){
  const btn=E('sbtn'); if(btn.dataset.busy==='1')return;
  const shared=document.querySelector('input[name=su_where]:checked').value==='shared';
  const folder=shared?E('su_folder').value.trim():'';
  if(shared&&!folder){E('serr').textContent='Choose the shared folder first.';return;}
  E('serr').textContent=''; btn.dataset.busy='1'; btn.textContent='Setting up\u2026';
  let r=null; try{ r=await API.complete_setup(E('su_co').value,E('su_name').value,E('su_email').value,E('su_pw').value,folder); }catch(e){ r={ok:false,error:String(e&&e.message||e)}; }
  btn.dataset.busy=''; btn.textContent='Create company & sign in';
  if(!r||!r.ok){E('serr').textContent=(r&&r.error)||'Setup failed.';return;}
  await applyBrand();
  E('setup').classList.add('hidden'); E('login').classList.remove('hidden');
  if(r.existing){ setBootMsg(''); E('lerr').textContent='Joined the existing database. Sign in with your account.'; E('lerr').style.color=''; return; }
  E('le').value=E('su_email').value.trim(); E('lp').value=E('su_pw').value;
  await doLogin();
}

const NAV=[
  ['dashboard','Dashboard',['Standard','Manager','Admin']],
  ['timesheet','My Timesheet',['Standard','Manager','Admin']],
  ['submissions','My Submissions',['Standard','Manager','Admin']],
  ['expenses','Expenses',['Standard','Manager','Admin']],
  ['approvals','Approvals',['Manager','Admin']],
  ['allts','All Timesheets',['Manager','Admin']],
  ['import','Import (external)',['Manager','Admin']],
  ['reports','Reports',['Manager','Admin']],
  ['analytics','Analytics',['Manager','Admin']],
  ['timeline','Timeline',['Manager','Admin']],
  ['jobcosts','Job Costs',['Manager','Admin']],
  ['users','Users',['Admin']],
  ['jobs','Job References',['Admin']],
  ['roles','Job Roles',['Admin']],
  ['audit','Audit Trail',['Admin']],
  ['settings','Settings',['Admin']],
  ['changepw','Change Password',['Standard','Manager','Admin']],
];
const NAV_ICON={
  dashboard:'<rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/>',
  timesheet:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  submissions:'<path d="M4 4h11l5 5v11H4z"/><path d="M15 4v5h5"/><path d="M8 13h8M8 17h5"/>',
  expenses:'<rect x="3" y="6" width="18" height="12" rx="2"/><path d="M3 10h18"/><path d="M7 15h3"/>',
  approvals:'<path d="M9 12l2 2 4-4"/><circle cx="12" cy="12" r="9"/>',
  allts:'<path d="M4 6h16M4 12h16M4 18h10"/>',
  import:'<path d="M12 3v12"/><path d="M8 11l4 4 4-4"/><path d="M4 17v3h16v-3"/>',
  reports:'<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
  analytics:'<path d="M3 17l5-6 4 3 5-7 4 4"/>',
  timeline:'<path d="M3 6h18"/><rect x="5" y="10" width="8" height="3" rx="1"/><rect x="9" y="15" width="10" height="3" rx="1"/>',
  jobcosts:'<path d="M12 2v20"/><path d="M16 6H10a3 3 0 0 0 0 6h4a3 3 0 0 1 0 6H8"/>',
  users:'<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0"/><circle cx="17" cy="9" r="2.5"/><path d="M15 15.5a5 5 0 0 1 6.5 4.5"/>',
  jobs:'<path d="M3 7h6l2 2h10v11H3z"/>',
  roles:'<rect x="4" y="7" width="16" height="13" rx="2"/><path d="M9 7V5a3 3 0 0 1 6 0v2"/>',
  audit:'<path d="M5 3h10l4 4v14H5z"/><path d="M9 12h6M9 16h6M9 8h2"/>',
  settings:'<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
  changepw:'<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>'
};
const NAV_GROUPS=[['',['dashboard','timesheet','submissions','expenses']],['Team',['approvals','allts','import','reports','analytics','timeline','jobcosts']],['Admin',['users','jobs','roles','audit','settings']],['Account',['changepw']]];
function buildNav(){
  const allowed=NAV.filter(n=>n[2].includes(USER.user_type));
  let h='';
  NAV_GROUPS.forEach(([label,keys])=>{
    const items=keys.map(k=>allowed.find(n=>n[0]===k)).filter(Boolean);
    if(!items.length)return;
    if(label) h+=`<div class="grp">${label}</div>`;
    h+=items.map(n=>`<a href="#" data-v="${n[0]}" onclick="nav('${n[0]}');return false;"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${NAV_ICON[n[0]]||'<circle cx="12" cy="12" r="4"/>'}</svg>${n[1]}</a>`).join('');
  });
  E('nav').innerHTML=h;
}
function setActive(v){[...document.querySelectorAll('.nav a')].forEach(a=>a.classList.toggle('active',a.dataset.v===v));}
function head(t,sub){return `<h1>${esc(t)}</h1><div class="h-accent"></div>`+(sub?`<div class="sub">${esc(sub)}</div>`:'');}

async function nav(v){
  setActive(v);
  const m=E('main'); m.classList.remove('enter'); m.classList.add('leaving');
  try{ await _navInner(v,m); }
  finally{ m.classList.remove('leaving'); void m.offsetWidth; m.classList.add('enter'); m.scrollTop=0; setTimeout(countUp,30); clearTimeout(window._enterT); window._enterT=setTimeout(()=>m.classList.remove('enter'),800); }
}
function fmtTs(v){ const d=new Date(v); if(isNaN(d)) return esc(v); const dd=d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}); const t=d.toTimeString().slice(0,5); return `<span class="nowrap">${dd} <span class="muted">${t}</span></span>`; }
const nowrap=v=>`<span class="nowrap">${esc(v)}</span>`;
function countUp(){
  document.querySelectorAll('.kpi .val, .att-big, .timer-disp').forEach(el=>{
    if(el.dataset.counted) return; el.dataset.counted='1';
    const txt=el.textContent.trim(); const mm=txt.match(/^(£?)(-?[\d,]+)(\.\d+)?(h|%| h)?$/); if(!mm) return;
    const target=parseFloat((mm[2]+(mm[3]||'')).replace(/,/g,'')); if(!isFinite(target)||Math.abs(target)<2) return;
    const dec=mm[3]?mm[3].length-1:0, pre=mm[1]||'', suf=mm[4]||''; const t0=performance.now(), dur=650;
    const fmt=n=>pre+(dec?n.toFixed(dec):Math.round(n).toLocaleString('en-GB'))+suf;
    const step=now=>{ const p=Math.min(1,(now-t0)/dur); const e=1-Math.pow(1-p,3); el.textContent=fmt(target*e); if(p<1) requestAnimationFrame(step); else el.textContent=txt; };
    requestAnimationFrame(step);
  });
}
async function _navInner(v,m){
  try{
    if(v==='dashboard') return renderDash(await API.dashboard());
    if(v==='timesheet') return renderTimesheet();
    if(v==='submissions') return renderSubmissions();
    if(v==='approvals') return renderApprovals();
    if(v==='allts') return renderAllTs();
    if(v==='import') return renderImport();
    if(v==='jobcosts') return renderJobCosts();
    if(v==='expenses') return renderExpenses();
    if(v==='users') return renderUsers();
    if(v==='jobs') return renderJobs();
    if(v==='roles') return renderRoles();
    if(v==='reports') return renderReports();
    if(v==='analytics') return renderAnalytics();
    if(v==='timeline') return renderTimeline();
    if(v==='changepw') return changePw();
    if(v==='settings') return renderSettings();
    if(v==='audit') return renderTable(m,'Audit Trail','', await API.audit(),
        [['time','When',fmtTs],['user','Who',nowrap],['action','Action',nowrap],['old','Before'],['new','After']]);
    // not-yet-ported interactive screens
    m.innerHTML=head(NAV.find(n=>n[0]===v)[1])+
      `<div class="card"><div class="muted">This screen's interactive version is coming in the next update. The data and all rules are already live underneath.</div></div>`;
  }catch(e){ m.innerHTML=head('Error')+`<div class="card">${esc(e)}</div>`; }
}

// ---- cell formatters ----
const badgeCell=(v)=>badge(v);
const yn=(v)=>v?'Yes':'No';
function renderTable(m,title,sub,rows,cols){
  let h=head(title,sub)+'<div class="card" style="padding:6px 8px"><table><thead><tr>';
  cols.forEach(c=>h+=`<th>${esc(c[1])}</th>`); h+='</tr></thead><tbody>';
  if(!rows.length) h+=`<tr><td colspan="${cols.length}" class="muted">Nothing to show.</td></tr>`;
  rows.forEach(r=>{ h+='<tr>'; cols.forEach(c=>{
    let v=r[c[0]], f=c[2];
    if(f==='num') h+=`<td class="tnum">${esc(v)}</td>`;
    else if(f==='money') h+=`<td class="tnum">${money(v)}</td>`;
    else if(typeof f==='function') h+=`<td>${f(v)}</td>`;
    else h+=`<td>${esc(v)}</td>`;
  }); h+='</tr>'; });
  h+='</tbody></table></div>'; m.innerHTML=h;
}

// ---- dashboard ----
function kpi(lab,val,green){return `<div class="card kpi"><div class="lab">${esc(lab)}</div><div class="val ${green?'green':''}">${esc(val)}</div></div>`;}
// A KPI card that opens an attached document when clicked.
function kpiDoc(lab,val,path,green){
  if(!path)return kpi(lab,val,green);
  return `<div class="card kpi kpi-doc" title="Open the quote document" onclick="openAtt('${jsstr(path)}')">
    <div class="lab">${esc(lab)} <span class="kpi-clip">📎</span></div>
    <div class="val ${green?'green':''}">${esc(val)}</div>
    <div class="kpi-open">Click to open the quote</div></div>`;}
const ATT_SVG={
  clock:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  inbox:'<path d="M3 13h5l2 3h4l2-3h5"/><path d="M5 5h14l2 8v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4z"/>',
  alert:'<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
  people:'<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/>',
  tick:'<path d="M20 6 9 17l-5-5"/>'
};
function attCard(cls,icon,big,lab,go,act){
  return `<div class="att ${cls}" onclick="${act}">
    <div class="att-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${ATT_SVG[icon]}</svg></div>
    <div><div class="att-big">${big}</div><div class="att-lab">${lab}</div>
    ${go?`<div class="att-go">${go} &rarr;</div>`:''}</div></div>`;
}
function renderDash(d){
  const hr=new Date().getHours();
  const greet=hr<12?'Good morning':(hr<18?'Good afternoon':'Good evening');
  let h=`<h1>${greet}, ${esc(d.first_name||'')}</h1><div class="hb"></div>`;

  // ---- things that actually need doing (only shown when true) ----
  const cards=[];
  if(d.rejected)
    cards.push(attCard('bad','alert','Timesheet rejected',
      esc(d.reject_reason||'Your manager sent this week back to you.'),'Put it right',"nav('timesheet')"));
  if(['None','Draft','Rejected'].includes(d.last_week_status))
    cards.push(attCard('bad','clock','Last week',
      `(w/c ${esc(d.last_week_label)}) still isn't submitted`,'Submit it',"nav('timesheet')"));
  else if(d.status==='Draft'){
    const pct=d.target?Math.round(d.my_week/d.target*100):0;
    cards.push(attCard(d.my_week>=d.target?'ok':'warn','clock','This week',
      `${d.my_week} of ${d.target} hours logged${pct?` (${pct}%)`:''}`,'Open my timesheet',"nav('timesheet')"));
  }
  if(d.is_mgr&&d.awaiting)
    cards.push(attCard(d.awaiting_oldest_days>10?'bad':'warn','inbox',
      d.awaiting+' timesheet'+(d.awaiting===1?'':'s'),
      'waiting for your approval'+(d.awaiting_oldest_days>10?', oldest is '+d.awaiting_oldest_days+' days old':''),
      'Go to Approvals',"nav('approvals')"));
  if(d.is_mgr&&d.missing){
    // Keep to one short line. Listing every name made this card three times
    // taller than the ones beside it and threw the whole row out.
    const shown=(d.missing_names||[]).slice(0,2).map(esc).join(', ');
    const rest=d.missing-Math.min(2,(d.missing_names||[]).length);
    cards.push(attCard('warn','people',d.missing+' '+(d.missing===1?'person':'people'),
      "haven't submitted last week"+(shown?' — '+shown+(rest>0?' +'+rest+' more':''):''),
      'See who',"nav('reports')"));
  }
  if(cards.length) h+='<div class="sec-h">NEEDS YOUR ATTENTION</div><div class="att-grid" id="attgrid">'+cards.join('')+'</div>';
  else h+=`<div class="att-grid"><div class="att ok" style="cursor:default">
      <div class="att-ico"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${ATT_SVG.tick}</svg></div>
      <div><div class="att-big">All clear</div><div class="att-lab">Nothing needs your attention right now.</div></div></div></div>`;

  // ---- money (Admin only) ----
  if(d.is_admin&&d.money){
    const m=d.money;
    h+='<div class="sec-h">MONEY</div><div class="grid kpis kpis-4" style="margin-bottom:20px">'
      +kpi('Invoiced this month',money0(m.invoiced_month))
      +kpi('Left to invoice',money0(m.to_invoice),1)
      +kpi('Invoiced, unpaid',money0(m.invoiced_unpaid))
      +kpi('Bills unpaid',money0(m.bills_unpaid))+'</div>';
  }
  h+='<div id="riskcard"></div><div id="folderwrap"></div>';

  // ---- charts ----
  h+='<div class="two">';
  h+=`<div class="card"><div class="sec-t">My hours – last 8 weeks</div>${barChart(d.weeks,d.target)}</div>`;
  if(d.is_mgr && d.by_role && d.by_role.length)
    h+=`<div class="card"><div class="sec-t">Company hours by role (this month)</div>${donut(d.by_role)}</div>`;
  else
    h+=`<div class="card"><div class="sec-t">Recent submissions</div>${miniSubs(d.prev)}</div>`;
  h+='</div>';

  if(d.is_mgr){
    h+='<div class="sec-h" style="margin-top:22px">COMPANY</div><div class="grid kpis">'+
      kpi('Hours this week',d.co_week)+kpi('Hours this month',d.co_month)+
      kpi('Active staff',d.active_staff)+kpi('Awaiting approval',d.awaiting)+'</div>';
  }
  h+=`<div class="sec-h" style="margin-top:22px">MY PREVIOUS TIMESHEETS</div><div class="card" style="padding:6px 8px"><table><thead><tr><th>Week starting</th><th>Status</th><th class="tnum">Total hours</th></tr></thead><tbody>`;
  if(!d.prev.length)h+='<tr><td colspan="3" class="muted">Nothing yet.</td></tr>';
  d.prev.forEach(p=>h+=`<tr><td>${esc(p.week)}</td><td>${badge(p.status)}</td><td class="tnum">${p.hours}</td></tr>`);
  h+='</tbody></table></div>';
  E('main').innerHTML=h;
  if(d.is_mgr){loadRisk(); loadFolders();}   // both fetched after the page is up
}
// ===== Folders: one folder per job, holding all of its documents ==========
let _folders=null, _folderQ='';
async function loadFolders(){
  let r=null; try{r=await API.document_folders();}catch(e){return;}
  if(!r||!r.folders)return;
  _folders=r; drawFolders();
}
function folderSearch(v){_folderQ=(v||'').toLowerCase();drawFolders(true);}
function drawFolders(keepFocus){
  const w=E('folderwrap'); if(!w||!_folders)return;
  const all=_folders.folders;
  const canDrop=USER&&USER.user_type==='Admin';
  const list=_folderQ?all.filter(f=>(f.job_ref+' '+f.job_title+' '+(f.client||'')).toLowerCase().includes(_folderQ)):all;
  let h='<div class="card" style="margin-bottom:20px">'
    +`<div class="sec-t" style="margin-bottom:8px">Folders <span class="muted" style="font-weight:400;font-size:12px">&nbsp;${all.length} job${all.length===1?'':'s'} &middot; ${_folders.total} document${_folders.total===1?'':'s'}</span></div>`;
  if(!all.length){
    w.innerHTML=h+'<div class="doc-empty">No documents filed against any job yet.<br><span style="font-size:12px">Attach a quote, invoice, bill or variation and the folder for that job appears here.</span></div></div>';
    return;
  }
  h+=`<input id="foldq" class="foldsearch" placeholder="Search jobs…" value="${esc(_folderQ)}" oninput="folderSearch(this.value)">`;
  if(!list.length)h+='<div class="doc-empty">No jobs match that.</div>';
  else{
    h+='<div class="fold-grid">';
    list.forEach(f=>{
      h+=`<div class="fold" data-job="${f.id}" onclick="openFolder(${f.id})" title="${esc(f.job_ref)} — ${esc(f.job_title)}${canDrop?' · drop files here to file them against this job':''}">
        <div class="fold-ico"><svg viewBox="0 0 24 24" fill="currentColor">
          <path class="ftab" d="M2.5 6.6A2.6 2.6 0 0 1 5.1 4h4.2c.66 0 1.3.25 1.78.7l1.4 1.3h6.4a2.6 2.6 0 0 1 2.6 2.6v8.8a2.6 2.6 0 0 1-2.6 2.6H5.1a2.6 2.6 0 0 1-2.6-2.6z"/>
          <path class="fbody" d="M2.5 10.2h19v7.2a2.6 2.6 0 0 1-2.6 2.6H5.1a2.6 2.6 0 0 1-2.6-2.6z"/></svg>
          <span class="fold-n">${f.count}</span></div>
        <div style="min-width:0">
          <div class="fold-ref">${esc(f.job_ref)}</div>
          <div class="fold-title">${esc(f.job_title||f.client||'')}</div>
          <div class="fold-meta">${f.count} file${f.count===1?'':'s'}${f.latest?' · '+dPretty(f.latest):''}</div>
        </div></div>`;});
    h+='</div>';
  }
  if(canDrop)h+='<div id="folddrop" class="dropzone" style="margin-top:12px">Drag files onto a job above to file them there — or drop here to pick the job</div>'
    +`<div style="display:flex;gap:8px;align-items:center;margin-top:8px;flex-wrap:wrap">
       <button class="btn sm ghost" onclick="tidyDocs()">Tidy files into job folders</button>
       <span class="muted" style="font-size:11.5px">Moves anything still loose in the shared Receipts folder into its own job folder.</span></div>`;
  w.innerHTML=h+'</div>';
  if(keepFocus){const i=E('foldq'); if(i){i.focus();i.setSelectionRange(i.value.length,i.value.length);}}
  if(canDrop){
    // each tile files straight onto that job
    w.querySelectorAll('.fold').forEach(el=>wireDrop(el,async fs=>{
      const jid=Number(el.getAttribute('data-job'));
      if(await fileFilesToJob(jid,fs))loadFolders();
    }));
    // dropping on the strip below asks which job it belongs to
    wireDrop(E('folddrop'),async fs=>{
      const arr=[]; for(let i=0;i<fs.length;i++)arr.push(fs[i]);
      pickJobForFiles(arr);
    },true);
  }
}
function pickJobForFiles(files){
  if(!_folders||!_folders.folders.length){alert('No job folders yet — attach a document to a job first.');return;}
  const opts=_folders.folders.map(f=>`<option value="${f.id}">${esc(f.job_ref)} — ${esc(f.job_title||'')}</option>`).join('');
  openModal('File '+files.length+' document(s)',
    `<p class="sub">Which job do these belong to?</p>
     <label>Job</label><select id="fj_job">${opts}</select>`,
    async ()=>{
      const jid=Number(E('fj_job').value); closeModal();
      if(await fileFilesToJob(jid,files))loadFolders();
    });
}
async function tidyDocs(){
  if(!confirm('Move every document that is still loose in the shared Receipts folder into its own job folder?\n\nFiles are moved, never deleted, and the app is updated to point at the new location.'))return;
  const r=await API.tidy_documents();
  if(!r||!r.ok){alert((r&&r.error)||'Could not tidy the documents.');return;}
  let msg=r.moved+' file(s) moved into job folders.';
  if(r.skipped)msg+='\n'+r.skipped+' skipped (file not found on disk).';
  if(r.errors&&r.errors.length)msg+='\n\nCould not move:\n'+r.errors.join('\n');
  alert(msg);
  loadFolders();
}
// open one job's folder as its own screen, with a way back
async function openFolder(jobId){
  const d=await API.job_documents(jobId);
  if(d&&d.error){alert(d.error);return;}
  DOCS=d; _docFilter='All';
  _docCtx={jobId:jobId,mode:'folder'};
  const f=(_folders&&_folders.folders.find(x=>x.id===jobId))||{job_ref:'',job_title:''};
  let h=`<div class="crumbs"><a onclick="nav('dashboard')">Folders</a><span>›</span><b>${esc(f.job_ref||'')}</b></div>`;
  h+=`<h1 style="font-size:22px">${esc(f.job_ref||'')} <span class="muted" style="font-weight:400;font-size:15px">${esc(f.job_title||'')}</span></h1><div class="hb"></div>`;
  h+=`<div class="toolbar" style="margin:0 0 14px"><button class="btn sm ghost" onclick="nav('dashboard')">‹ Back to folders</button>
      <button class="btn sm ghost" onclick="renderBilling(${jobId})">Open this job's money</button>
      <button class="btn sm ghost" onclick="API.open_folder('${jsstr(d.folder)}')">Open in Explorer</button></div>`;
  h+='<div class="card" id="foldcard"></div>';
  if(d.can_edit)
    h+=`<div id="folddrop2" class="dropzone" style="margin-top:14px">Drag files here to file them against this job</div>
        <div style="display:flex;gap:8px;align-items:center;margin-top:8px;flex-wrap:wrap">
          <button class="btn sm" onclick="pickDocFor(${jobId})">Add document…</button>
          <span class="muted" style="font-size:11.5px">Stored in ${esc(d.folder)}</span></div>`;
  E('main').innerHTML=h;
  folderFilter(jobId,'All');
  if(d.can_edit)wireDrop(E('folddrop2'),async fs=>{
    if(await fileFilesToJob(jobId,fs)){await loadFolders(); openFolder(jobId);}
  },true);
}
function folderFilter(jobId,k){
  _docFilter=k;
  const d=DOCS, card=E('foldcard'); if(!d||!card)return;
  const kinds=['All','Quote','Variation','Invoice','Bill','Document'];
  const counts={}; d.docs.forEach(x=>counts[x.kind]=(counts[x.kind]||0)+1);
  const rows=d.docs.filter(x=>k==='All'||x.kind===k);
  card.innerHTML='<div class="doc-filters">'+kinds.filter(x=>x==='All'||counts[x])
      .map(x=>{const c=x==='All'?d.docs.length:counts[x];
        return `<button class="doc-chip ${_docFilter===x?'on':''}" onclick="folderFilter(${jobId},'${x}')">${x}<span style="opacity:.7"> ${c}</span></button>`;}).join('')+'</div>'
    +(d.docs.length?(rows.length?docRowsHtml(rows,d.can_edit):'<div class="doc-empty">Nothing of that type.</div>')
      :'<div class="doc-empty">This folder is empty.</div>');
}
async function loadRisk(){
  let r=null; try{r=await API.dash_jobs_at_risk(6);}catch(e){return;}
  if(!r||!r.rows||!r.rows.length)return;
  // add an attention card if anything is actually over budget
  if(r.over_budget){
    const g=E('attgrid');
    if(g)g.insertAdjacentHTML('beforeend',attCard('bad','alert',
      r.over_budget+' job'+(r.over_budget===1?'':'s'),'over budget on cost','Review job costs',"nav('jobcosts')"));
  }
  const w=E('riskcard'); if(!w)return;
  let h='<div class="card" style="margin:0 0 20px"><div class="sec-t">Jobs by cost against quote</div>';
  r.rows.forEach(j=>{const col=j.pct>100?'#d23a3a':(j.pct>=85?'#e6a100':'#84c341');
    h+=`<div class="jrow" onclick="renderBilling(${j.id})" title="Open this job's money">
      <b class="jref" title="${esc(j.job_ref)}">${esc(j.job_ref)}</b>
      <span class="jtitle" title="${esc(j.job_title)}">${esc(j.job_title)}</span>
      <div class="jbar"><i style="width:${Math.min(j.pct,100)}%;background:${col}"></i></div>
      <span class="jpct" style="color:${col}">${Number(j.pct).toFixed(1)}%</span></div>`;});
  h+='<div class="sub" style="margin-top:8px">Amber 85%+ &middot; red over budget. Click a job to open its money.</div></div>';
  w.innerHTML=h;
}
function miniSubs(prev){
  if(!prev.length) return '<div class="chart-empty">No submissions yet.</div>';
  return prev.slice(0,6).map(p=>`<div style="display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid var(--line)"><span>${esc(p.week)}</span>${badge(p.status)}</div>`).join('');
}
function barChart(data,target){
  const vals=data.map(d=>Number(d.hours)||0); const top=Math.max(1,...vals,target||0); const max=Math.ceil(top/10)*10||10;
  const W=640,H=210,padL=34,padR=12,padT=18,padB=30,iw=W-padL-padR,ih=H-padT-padB,bw=iw/data.length;
  const y=v=>padT+ih-(ih*v/max);
  let s=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}" class="bars" role="img" aria-label="Hours per week">`;
  [0,max/2,max].forEach(g=>{ s+=`<line x1="${padL}" x2="${W-padR}" y1="${y(g)}" y2="${y(g)}" stroke="var(--line)" stroke-dasharray="${g?'3 4':'0'}"/><text x="${padL-6}" y="${y(g)+3.5}" font-size="10" fill="var(--mut-2)" text-anchor="end">${g}</text>`; });
  if(target){ s+=`<line x1="${padL}" x2="${W-padR}" y1="${y(target)}" y2="${y(target)}" stroke="var(--amber)" stroke-width="1.2" stroke-dasharray="5 4"/><text x="${padL+4}" y="${y(target)-5}" font-size="9.5" font-weight="700" fill="var(--amber)">target ${target}h</text>`; }
  const pts=[];
  data.forEach((d,i)=>{ const v=vals[i]; const x=padL+i*bw+bw*0.22, w=bw*0.56, yy=y(v), last=i===data.length-1;
    pts.push([x+w/2,yy]);
    s+=`<g class="bar" style="animation-delay:${i*45}ms"><rect x="${x}" y="${yy}" width="${w}" height="${Math.max(0,padT+ih-yy)}" rx="4" fill="${last?'var(--green)':'#b9e2a3'}"></rect>`;
    s+=`<rect x="${padL+i*bw}" y="${padT}" width="${bw}" height="${ih}" fill="transparent"><title>${esc(d.label)}: ${v}h</title></rect>`;
    s+=`<text x="${x+w/2}" y="${H-padB+15}" font-size="10" fill="var(--mut)" text-anchor="middle">${esc(d.label)}</text>`;
    if(v) s+=`<text x="${x+w/2}" y="${yy-5}" font-size="10.5" font-weight="700" fill="${last?'var(--green-d)':'var(--mut)'}" text-anchor="middle">${v}</text>`;
    s+='</g>'; });
  if(pts.length>1){ s+=`<polyline points="${pts.map(p=>p.join(',')).join(' ')}" fill="none" stroke="var(--green-d)" stroke-width="1.5" stroke-opacity=".55" stroke-linejoin="round" class="trend"/>`; }
  s+='</svg>'; return s;
}
function donut(data){
  const tot=data.reduce((a,b)=>a+b.value,0)||1; const cx=90,cy=90,r=66,sw=22; let a=-Math.PI/2;
  const cols=['#6cc04a','#2f6fd6','#d98c00','#9b59b6','#15a3a3','#e2575b','#7a8a2b','#4f6bed'];
  let s=`<svg viewBox="0 0 360 180" width="100%" height="180">`;
  data.forEach((d,i)=>{ const frac=d.value/tot, a2=a+frac*2*Math.PI;
    const x1=cx+r*Math.cos(a),y1=cy+r*Math.sin(a),x2=cx+r*Math.cos(a2),y2=cy+r*Math.sin(a2);
    const large=frac>0.5?1:0;
    s+=`<path d="M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}" fill="none" stroke="${cols[i%cols.length]}" stroke-width="${sw}" stroke-linecap="butt"><title>${esc(d.label)}: ${d.value}h (${Math.round(frac*100)}%)</title></path>`;
    a=a2; });
  s+=`<text x="${cx}" y="${cy+6}" font-size="21" font-weight="800" text-anchor="middle" fill="var(--ink)">${Math.round(tot)}h</text>`;
  data.slice(0,6).forEach((d,i)=>{ s+=`<rect x="195" y="${26+i*22}" width="10" height="10" rx="3" fill="${cols[i%cols.length]}"></rect><text x="212" y="${35+i*22}" font-size="11.5" fill="var(--ink)">${esc(String(d.label).length>17?String(d.label).slice(0,16)+'…':d.label)}</text><text x="352" y="${35+i*22}" font-size="11.5" font-weight="700" fill="var(--mut)" text-anchor="end">${d.value}h</text>`; });
  s+='</svg>'; return s;
}
async function doApprove(id,st){
  if(!ensureApi()){ alert('Still connecting - wait a second and try again.'); return; }
  try{ await API.approve(id,st,''); }
  catch(e){ alert('Could not update timesheet: '+((e&&e.message)?e.message:e)); return; }
  toast(st==='Locked'?'Locked':(st==='Approved'?'Approved ✓':'Updated'));
  nav('approvals');
}
async function doReject(id){
  const r=prompt('Reason for rejection (the person will see this):'); if(r===null)return;
  if(!ensureApi()){ alert('Still connecting - wait a second and try again.'); return; }
  try{ await API.approve(id,'Rejected',r); }
  catch(e){ alert('Could not reject timesheet: '+((e&&e.message)?e.message:e)); return; }
  toast('Sent back to employee');
  nav('approvals');
}
function openModal(title,bodyHtml,onOk){E('modalTitle').textContent=title;E('modalBody').innerHTML=bodyHtml;E('modalOk').textContent='Save';const mc=E('modalCancel');if(mc)mc.style.display='';E('modal').classList.remove('hidden');window._mok=onOk;}
function closeModal(){E('modal').classList.add('hidden');}
E('modalOk').onclick=()=>{if(window._mok)window._mok();};

let TS=null;
async function renderTimesheet(weekStart){
  TS=await API.timesheet(weekStart||null);
  try{ window._mySubs=await API.submissions(); }catch(e){ window._mySubs=[]; }
  window._tjobs=await API.timesheet_jobs(); window._troles=await API.timesheet_roles();
  try{const tm=await API.timer_get(); window._timerStart=(tm&&tm.start_ms)?tm.start_ms:null;}catch(e){}
  drawTimesheet();
}
function weekStrip(cur){
  const st={}; (window._mySubs||[]).forEach(x=>st[x.week_start]=x.status);
  const today=new Date(); const dow=(today.getDay()+6)%7; const thisMon=new Date(today); thisMon.setDate(today.getDate()-dow);
  const curD=new Date(cur); const weeks=[]; const start=new Date(Math.min(curD,thisMon)); const endW=new Date(Math.max(curD,thisMon));
  const first=new Date(endW); first.setDate(first.getDate()-7*7); const from=new Date(Math.min(first,start));
  for(let w=new Date(from); w<=endW; w.setDate(w.getDate()+7)) weeks.push(new Date(w));
  const iso=x=>{const z=new Date(x); z.setMinutes(z.getMinutes()-z.getTimezoneOffset()); return z.toISOString().slice(0,10);};
  const lab=x=>x.toLocaleDateString('en-GB',{day:'numeric',month:'short'});
  return `<div class="wk-strip">${weeks.slice(-8).map(w=>{const k=iso(w); const s=st[k]||''; const cls=(k===cur?' on':'')+(k===iso(thisMon)?' now':'')+(s?' s-'+s:' s-none');
    return `<button class="wk${cls}" onclick="renderTimesheet('${k}')" title="w/c ${lab(w)}${s?' · '+s:''}"><span class="wk-d"></span>${lab(w)}</button>`;}).join('')}</div>`;
}
function weekShift(n){const w=new Date(TS.week_start);w.setDate(w.getDate()+n);renderTimesheet(w.toISOString().slice(0,10));}
function drawTimesheet(){
  const d=TS; let h=head('My Timesheet');
  h+=`<div class="ts-top"><button class="btn sm ghost" onclick="weekShift(-7)" title="Previous week">‹</button>
    <span class="ts-week">${esc(d.week)}</span>
    <button class="btn sm ghost" onclick="weekShift(7)" title="Next week">›</button>
    <button class="btn sm ghost" onclick="renderTimesheet()">This week</button>
    <span style="margin-left:auto">${badge(d.status)}</span></div>`+weekStrip(d.week_start);
  if(d.status==='Rejected'&&d.reject_reason) h+=`<div class="card" style="border-left:4px solid #d23a3a;color:#b21d36;margin-bottom:12px">Rejected: ${esc(d.reject_reason)}</div>`;
  if(d.editable){const running=!!window._timerStart;
    h+=`<div class="timer-bar"><span class="timer-lbl">TIMER</span>
      <span id="timerDisp" class="timer-disp">${running?fmtHMS(Date.now()-window._timerStart):'00:00:00'}</span>
      ${running?`<button class="btn sm" onclick="timerStopAdd()">Stop &amp; add line</button> <button class="btn sm ghost" onclick="timerReset()">Reset</button>`:`<button class="btn sm" onclick="timerStart()">Start timer</button>`}
      <span class="muted" style="font-size:12px;margin-left:auto">Times what you're on now &middot; rounds to nearest 15 min</span></div>`;}
  h+='<div class="card" style="padding:6px 8px"><table><thead><tr><th>Day</th><th>Job Ref</th><th>Role</th><th>Narrative</th><th class="tnum">Hours</th><th>Rate</th></tr></thead><tbody>';
  d.days.forEach((day,i)=>{
    const dl=d.lines.filter(l=>l.work_date===day); const tot=dl.reduce((a,b)=>a+b.hours,0);
    const wflag=(tot>(window._dailyWarn||24))?' <span style="color:#d23a3a">(!)</span>':'';
    const acts=d.editable?`<span class="day-actions">
        <button class="btn xs" onclick="addLineForDay('${day}')">+ line</button>
      </span>`:'';
    h+=`<tr class="dayrow" ${d.editable?`oncontextmenu="dayMenu(event,'${day}')" onclick="_tsDay='${day}'"`:''}>
        <td colspan="6"><b>${esc(d.day_labels[i])}</b>${tot?` <span class="muted">(${tot}h)${wflag}</span>`:''}${acts}</td></tr>`;
    dl.forEach(l=>{ const idx=d.lines.indexOf(l);
      const narr = l.narrative && l.narrative.trim() ? esc(l.narrative) : '<span style="color:#c0392b">(add narrative)</span>';
      h+=`<tr class="${d.editable?'lrow':''}${_tsSel.has(idx)?' sel':''}" data-idx="${idx}"
        ${d.editable?`onclick="selLine(event,${idx},'${day}')" ondblclick="editLine(${idx})" oncontextmenu="lineMenu(event,${idx},'${day}')"`:''}>
        <td></td><td>${esc(l.job_ref)}</td><td>${esc(l.job_role)}</td><td>${narr}</td>
        <td class="tnum">${l.hours}</td><td>${l.rate==='Normal'?'':esc(l.rate)} ${d.editable?`<button class="xdel" title="Delete" onclick="event.stopPropagation();delLine(${idx})">✕</button>`:''}</td></tr>`; });
  });
  h+='</tbody></table></div>';
  h+=`<div class="ts-foot"><span class="muted">Week total: <b>${d.worked} h</b>${d.paid!==d.worked?` &nbsp;(paid ${d.paid} h)`:''}</span>`;
  if(d.editable) h+=`<span style="margin-left:auto;display:flex;gap:8px">
      <span class="muted" style="align-self:center;font-size:11.5px;margin-right:4px">Click a line to select &middot; right-click to copy &amp; paste</span>
      <button class="btn sm ghost" onclick="pasteLastWeek()">Paste last week</button>
      <button class="btn sm ghost" onclick="addLine()">Add line</button>
      <button class="btn sm" onclick="submitTs()">Submit for approval</button></span>`;
  else h+=`<span style="margin-left:auto" class="muted">This is the version you submitted (${esc(d.status)}).</span>`;
  h+='</div>'; E('main').innerHTML=h; timerEnsure();
}
function lineModal(line,idx,presetDay){
  const days=TS.days.map((d,i)=>`<option value="${d}">${esc(TS.day_labels[i])}</option>`).join('');
  const jobOpts=(line?'':'<option value="" disabled selected>— select a job —</option>')+window._tjobs.map(j=>`<option value="${j.id}">${esc(j.label)}</option>`).join('');
  const roleOpts=window._troles.map(r=>`<option value="${esc(r)}">${esc(r)}</option>`).join('');
  const rateOpts=['Normal','x1.5','x2'].map(r=>`<option value="${r}">${r}</option>`).join('');
  openModal('Timesheet line',
    `<label>Day</label><select id="m_day">${days}</select>
     <label>Job reference</label><select id="m_job">${jobOpts}</select>
     <label>Role</label><select id="m_role">${roleOpts}</select>
     <label>Overtime rate</label><select id="m_rate">${rateOpts}</select>
     <label>Narrative</label><textarea id="m_narr" rows="3"></textarea>
     <label>Hours</label><input id="m_hours" type="number" step="0.25" min="0">`,
    ()=>saveLineModal(idx));
  if(line){E('m_day').value=line.work_date;E('m_job').value=line.job_ref_id;E('m_narr').value=line.narrative;E('m_hours').value=line.hours;E('m_rate').value=line.rate||'Normal';}
  else if(presetDay){E('m_day').value=presetDay;}
  E('m_role').value=(line&&line.job_role)||USER.job_role;
}
function addLine(){lineModal(null,-1);}
function addLineForDay(day){lineModal(null,-1,day);}

// ===== select / copy / paste on the timesheet ==============================
// An in-app clipboard rather than the Windows one: we're moving timesheet
// lines, not text, and this leaves Ctrl+C working normally inside text boxes.
let _tsSel=new Set(), _tsClip=[], _tsDay='', _tsAnchor=null;
function selLine(ev,idx,day){
  if(ev&&ev.target&&ev.target.classList&&ev.target.classList.contains('xdel'))return;
  _tsDay=day;
  if(ev&&ev.shiftKey&&_tsAnchor!==null){
    const a=Math.min(_tsAnchor,idx),b=Math.max(_tsAnchor,idx);
    _tsSel=new Set(); for(let i=a;i<=b;i++)_tsSel.add(i);
  }else if(ev&&(ev.ctrlKey||ev.metaKey)){
    if(_tsSel.has(idx))_tsSel.delete(idx); else _tsSel.add(idx);
    _tsAnchor=idx;
  }else{
    if(_tsSel.size===1&&_tsSel.has(idx))_tsSel=new Set();   // click again to deselect
    else _tsSel=new Set([idx]);
    _tsAnchor=idx;
  }
  drawTimesheet();
}
function tsCopy(idxs){
  const list=(idxs&&idxs.length?idxs:Array.from(_tsSel)).sort((a,b)=>a-b);
  if(!list.length)return 0;
  _tsClip=list.map(i=>TS.lines[i]).filter(Boolean).map(l=>({
    job_ref_id:l.job_ref_id,job_role:l.job_role,rate:l.rate,
    narrative:l.narrative||'',hours:l.hours}));
  toast(_tsClip.length+' line(s) copied');
  return _tsClip.length;
}
function tsCopyDay(day){
  const idxs=TS.lines.map((l,i)=>l.work_date===day?i:-1).filter(i=>i>=0);
  if(!idxs.length){toast('Nothing on that day');return;}
  tsCopy(idxs);
}
async function tsPaste(day,keepNarr){
  if(!_tsClip.length){toast('Nothing copied yet');return;}
  if(!day){toast('Click the day you want to paste into first');return;}
  _tsClip.forEach(l=>TS.lines.push({work_date:day,job_ref_id:l.job_ref_id,job_role:l.job_role,
    rate:l.rate,narrative:keepNarr?(l.narrative||''):'',hours:l.hours,job_ref:''}));
  _tsSel=new Set();
  await autosaveTs();
  toast(_tsClip.length+' line(s) pasted'+(keepNarr?'':' without narratives'));
}
async function tsDeleteSel(){
  const list=Array.from(_tsSel).sort((a,b)=>b-a);
  if(!list.length)return;
  if(!confirm('Delete '+list.length+' selected line(s)?'))return;
  list.forEach(i=>TS.lines.splice(i,1));
  _tsSel=new Set();
  await autosaveTs();
}
async function tsClearDay(day){
  const n=TS.lines.filter(l=>l.work_date===day).length;
  if(!n)return;
  if(!confirm('Delete all '+n+' line(s) on this day?'))return;
  TS.lines=TS.lines.filter(l=>l.work_date!==day);
  _tsSel=new Set();
  await autosaveTs();
}
// ---- right-click menu ----
function closeCtx(){const m=E('ctxmenu'); if(m)m.remove();}
function showCtx(ev,items){
  ev.preventDefault(); ev.stopPropagation();
  closeCtx();
  const m=document.createElement('div'); m.id='ctxmenu'; m.className='ctxmenu';
  m.innerHTML=items.map(it=>it.sep?'<div class="sep"></div>'
    :(it.hint?`<div class="hint">${esc(it.hint)}</div>`
    :`<button ${it.disabled?'disabled':''} onclick="closeCtx();${it.act}">${it.label}${it.key?`<span class="k">${it.key}</span>`:''}</button>`)).join('');
  document.body.appendChild(m);
  const w=m.offsetWidth,h=m.offsetHeight;      // keep it fully on screen
  m.style.left=Math.max(4,Math.min(ev.clientX,window.innerWidth-w-8))+'px';
  m.style.top=Math.max(4,Math.min(ev.clientY,window.innerHeight-h-8))+'px';
}
document.addEventListener('mousedown',e=>{
  if(!e.target.closest||!e.target.closest('#ctxmenu'))closeCtx();},true);
window.addEventListener('resize',closeCtx);
function lineMenu(ev,idx,day){
  if(!TS||!TS.editable)return;
  if(!_tsSel.has(idx)){_tsSel=new Set([idx]);_tsAnchor=idx;drawTimesheet();}
  _tsDay=day;
  const n=_tsSel.size,c=_tsClip.length;
  showCtx(ev,[
    {label:`Copy ${n} line${n===1?'':'s'}`,act:'tsCopy()',key:'Ctrl+C'},
    {label:'Duplicate on this day',act:`tsCopy();tsPaste('${day}',true)`},
    {sep:1},
    {label:c?`Paste ${c} line${c===1?'':'s'} here`:'Paste',disabled:!c,act:`tsPaste('${day}',true)`,key:'Ctrl+V'},
    {label:'Paste without narratives',disabled:!c,act:`tsPaste('${day}',false)`},
    {sep:1},
    {label:'Edit line…',act:`editLine(${idx})`},
    {label:`Delete ${n} line${n===1?'':'s'}`,act:'tsDeleteSel()',key:'Del'}
  ]);
}
function dayMenu(ev,day){
  if(!TS||!TS.editable)return;
  _tsDay=day;
  const dl=TS.lines.filter(l=>l.work_date===day).length,c=_tsClip.length;
  const lbl=TS.day_labels[TS.days.indexOf(day)]||'this day';
  showCtx(ev,[
    {hint:lbl},
    {label:`Copy this day (${dl} line${dl===1?'':'s'})`,disabled:!dl,act:`tsCopyDay('${day}')`},
    {sep:1},
    {label:c?`Paste ${c} line${c===1?'':'s'} here`:'Paste',disabled:!c,act:`tsPaste('${day}',true)`,key:'Ctrl+V'},
    {label:'Paste without narratives',disabled:!c,act:`tsPaste('${day}',false)`},
    {sep:1},
    {label:'Add a line to this day…',act:`addLineForDay('${day}')`},
    {label:'Clear this day',disabled:!dl,act:`tsClearDay('${day}')`}
  ]);
}
// ---- keyboard shortcuts, ignored while typing in a field ----
document.addEventListener('keydown',e=>{
  if(!TS||!TS.editable)return;
  if(!E('main')||!E('main').querySelector('.ts-top'))return;      // not on My Timesheet
  if(!E('modal').classList.contains('hidden'))return;             // a form is open
  const t=e.target&&e.target.tagName;
  if(t==='INPUT'||t==='TEXTAREA'||t==='SELECT')return;            // let text copy normally
  const k=(e.key||'').toLowerCase();
  if((e.ctrlKey||e.metaKey)&&k==='c'&&_tsSel.size){e.preventDefault();tsCopy();}
  else if((e.ctrlKey||e.metaKey)&&k==='v'&&_tsClip.length){e.preventDefault();tsPaste(_tsDay,true);}
  else if((e.key==='Delete')&&_tsSel.size){e.preventDefault();tsDeleteSel();}
  else if(e.key==='Escape'){_tsSel=new Set();closeCtx();drawTimesheet();}
});
function fmtHMS(ms){const s=Math.max(0,Math.floor(ms/1000));const p=n=>String(n).padStart(2,'0');return p(Math.floor(s/3600))+':'+p(Math.floor(s%3600/60))+':'+p(s%60);}
function timerTick(){const el=E('timerDisp');if(el&&window._timerStart)el.textContent=fmtHMS(Date.now()-window._timerStart);}
function timerEnsure(){if(window._timerInt){clearInterval(window._timerInt);window._timerInt=null;}if(window._timerStart){window._timerInt=setInterval(timerTick,1000);timerTick();}}
function timerStart(){const ms=Date.now();window._timerStart=ms;try{API.timer_start(ms);}catch(e){}drawTimesheet();}
function timerReset(){window._timerStart=null;if(window._timerInt){clearInterval(window._timerInt);window._timerInt=null;}try{API.timer_clear();}catch(e){}drawTimesheet();}
async function timerStopAdd(){
  if(!window._timerStart)return;
  let hrs=Math.round(((Date.now()-window._timerStart)/3600000)/0.25)*0.25; if(hrs<0.25)hrs=0.25;
  window._timerStart=null; if(window._timerInt){clearInterval(window._timerInt);window._timerInt=null;}
  try{await API.timer_clear();}catch(e){}
  const t=new Date(); t.setHours(12,0,0,0);
  const todayISO=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,'0')}-${String(t.getDate()).padStart(2,'0')}`;
  if(!TS.days.includes(todayISO)){ await renderTimesheet(); }  // jump to this week
  const day=TS.days.includes(todayISO)?todayISO:TS.days[0];
  drawTimesheet(); lineModal(null,-1,day); const hb=E('m_hours'); if(hb)hb.value=hrs;
}
function editLine(idx){lineModal(TS.lines[idx],idx);}
async function delLine(idx){TS.lines.splice(idx,1);await autosaveTs();}
function strip(l){return {work_date:l.work_date,job_ref_id:Number(l.job_ref_id),narrative:l.narrative,hours:Number(l.hours),job_role:l.job_role,rate:l.rate};}
async function saveLineModal(idx){
  const hrs=parseFloat(E('m_hours').value); const narr=E('m_narr').value.trim(); const jobVal=E('m_job').value;
  if(!jobVal){alert('Please select a job.');return;}
  if(!narr){alert('Narrative is required.');return;}
  if(!(hrs>0)){alert('Hours must be greater than zero.');return;}
  const ln={work_date:E('m_day').value,job_ref_id:Number(jobVal),job_role:E('m_role').value,rate:E('m_rate').value,narrative:narr,hours:hrs,job_ref:''};
  if(idx>=0)TS.lines[idx]=Object.assign({},TS.lines[idx],ln); else TS.lines.push(ln);
  closeModal(); await autosaveTs();
}
async function autosaveTs(){const r=await API.save_timesheet(TS.id,TS.lines.map(strip));if(!r.ok)alert(r.error);await renderTimesheet(TS.week_start);}
async function pasteLastWeek(){
  const prevISO=shiftISO(TS.week_start,-7);
  const last=await API.timesheet(prevISO);
  const src=(last&&last.lines)?last.lines:[];
  if(!src.length){alert('Last week has no lines to copy.');return;}
  if(!confirm(`Copy ${src.length} line(s) from last week onto this week? Job, role, hours and overtime are copied \u2014 narratives are left blank for you to fill in.`))return;
  src.forEach(l=>{const idx=diffDaysISO(prevISO,l.work_date);
    if(idx>=0&&idx<=6)TS.lines.push({work_date:TS.days[idx],job_ref_id:l.job_ref_id,job_role:l.job_role,rate:l.rate,narrative:'',hours:l.hours,job_ref:''});});
  await autosaveTs();
  toast(src.length+' line(s) pasted from last week');
}
async function submitTs(){
  const blank=TS.lines.filter(l=>!l.narrative||!l.narrative.trim()).length;
  if(blank){alert(`Please fill in the narrative for ${blank} line(s) before submitting \u2014 they are shown in red as "(add narrative)".`);return;}
  if(!confirm('Submit this timesheet for approval? Afterwards you can only view what you submitted (unless it is rejected back to you).'))return;
  const r=await API.submit_my_timesheet(TS.id,TS.lines.map(strip));
  if(!r.ok){alert(r.error);return;}
  renderTimesheet(TS.week_start);
}

function jsstr(s){return String(s==null?'':s).replace(/\\/g,'\\\\').replace(/'/g,"\\'");}
async function opts(){if(!window._opts)window._opts=await API.options();return window._opts;}
function invOpts(){window._opts=null;}

// ===== Users =====
let _users=[];
async function renderUsers(){
  _users=await API.users();
  let h=head('Users','Set each person\'s standard charge rate here - it applies to every job unless you override it per job under Job Costs.')+'<div class="card" style="padding:6px 8px"><table><thead><tr><th>Name</th><th>Email</th><th>Job Role</th><th>User Type</th><th class="tnum">Rate (£/hr)</th><th>Active</th><th></th></tr></thead><tbody>';
  _users.forEach(u=>h+=`<tr><td>${person(u.name)}</td><td class="muted">${esc(u.email)}</td><td>${esc(u.role)}</td><td>${esc(u.type)}</td><td class="tnum">${u.rate?money(u.rate):'—'}</td><td>${u.active?'Yes':'No'}</td>
    <td style="text-align:right"><button class="btn sm ghost" onclick="userModal(${u.id})">Edit</button>
      <button class="btn sm ghost" onclick="resetPw(${u.id})">Reset pw</button>
      <button class="btn sm ghost" onclick="delUser(${u.id})">Delete</button></td></tr>`);
  h+='</tbody></table></div><div class="toolbar"><button class="btn sm" onclick="userModal(0)">Add user</button></div>';
  E('main').innerHTML=h;
}
function userModal(id){
  const u=id?_users.find(x=>x.id===id):null;
  opts().then(o=>{
    const roles=o.roles.map(r=>`<option>${esc(r)}</option>`).join('');
    const types=o.user_types.map(t=>`<option>${esc(t)}</option>`).join('');
    openModal(u?'Edit user':'Add user',
     `<label>Full name</label><input id="u_name">
      <label>Email</label><input id="u_email">
      <label>Job role</label><select id="u_role">${roles}</select>
      <label>User type</label><select id="u_type">${types}</select>
      <label>Standard rate (£/hr)</label><input id="u_rate" type="number" step="0.01" min="0" placeholder="0.00">
      <label style="display:flex;gap:8px;align-items:center;margin-top:12px"><input type="checkbox" id="u_active" style="width:auto"> Active</label>
      ${u?'':'<label>Password</label><input id="u_pw" type="password">'}`,()=>saveUser(u));
    if(u){E('u_name').value=u.name;E('u_email').value=u.email;E('u_role').value=u.role;E('u_type').value=u.type;E('u_rate').value=u.rate||'';E('u_active').checked=u.active;}
    else E('u_active').checked=true;
  });
}
async function saveUser(u){
  const name=E('u_name').value.trim(),email=E('u_email').value.trim();
  if(!name||!email){alert('Name and email are required.');return;}
  const rate=parseFloat(E('u_rate').value)||0;
  let r;
  if(u)r=await API.update_user(u.id,name,email,E('u_role').value,E('u_type').value,E('u_active').checked,rate);
  else{const pw=E('u_pw').value;if(!pw){alert('Password required.');return;}r=await API.add_user(name,email,pw,E('u_role').value,E('u_type').value,E('u_active').checked,rate);}
  if(!r.ok){alert(r.error);return;}
  closeModal();invOpts();renderUsers();
}
async function resetPw(id){const p=prompt('New password for this user:');if(!p)return;await API.reset_password(id,p);alert('Password reset.');}
async function delUser(id){if(!confirm('Delete (or deactivate) this user?'))return;const r=await API.delete_user(id);if(!r.ok){alert(r.error);return;}renderUsers();}

// ===== Job References =====
let _jobs=[];
// Which jobs the two job screens show. Defaults to Active on every load so
// closed work stays out of the way until someone deliberately asks for it.
let _jobStatus='Active';
function statusOpts(sel,act,cls,all){
  return `<option value="Active" ${sel==='Active'?'selected':''}>Active (${act})</option>`
       + `<option value="Closed" ${sel==='Closed'?'selected':''}>Closed (${cls})</option>`
       + `<option value="All" ${sel==='All'?'selected':''}>All jobs (${all})</option>`;
}
async function renderJobs(){
  _jobs=await API.jobs();
  drawJobs();
}
function drawJobs(){
  const nA=_jobs.filter(j=>j.status==='Active').length, nC=_jobs.length-nA;
  const list=_jobStatus==='All'?_jobs.slice():_jobs.filter(j=>j.status===_jobStatus);
  let h=head('Job References')
    +`<div class="toolbar" style="margin:0 0 14px"><label class="muted" style="align-self:center">Show:</label>
      <select id="jr_st" onchange="_jobStatus=this.value;drawJobs()">${statusOpts(_jobStatus,nA,nC,_jobs.length)}</select>
      <span class="muted" style="align-self:center;margin-left:auto">${list.length} job${list.length===1?'':'s'}</span></div>`
    +'<div class="card" style="padding:6px 8px"><table><thead><tr><th>Job Ref</th><th>Title</th><th>Client</th><th>Type</th><th>Status</th><th class="tnum">Quote (£)</th><th></th></tr></thead><tbody>';
  if(!list.length)h+=`<tr><td colspan="7" class="muted">No ${_jobStatus==='All'?'':_jobStatus.toLowerCase()+' '}job references${_jobStatus==='Closed'?' — nothing has been closed yet.':'.'}</td></tr>`;
  list.forEach(j=>{const doc=j.quote_att?` <button class="btn sm ghost" title="Open the quote document" onclick="event.stopPropagation();openAtt('${jsstr(j.quote_att_path)}')">📎</button>`:'';
    h+=`<tr class="jobrow" data-job="${j.id}" title="Drop a quote document here to attach it"><td>${esc(j.ref)}</td><td>${esc(j.title)}</td><td>${esc(j.client)}</td><td>${esc(j.type)}</td><td>${badge(j.status)}</td><td class="tnum">${j.quote?moneyS(j.quote):'—'}${doc}</td>
    <td style="text-align:right"><button class="btn sm ghost" onclick="jobModal(${j.id})">Edit</button><button class="btn sm ghost" onclick="delJob(${j.id})">Delete</button></td></tr>`;});
  h+='</tbody></table></div><div class="toolbar"><button class="btn sm" onclick="jobModal(0)">Add job</button><span class="muted" style="align-self:center;margin-left:8px">Tip: drag a quote PDF onto a job row to attach it.</span></div>';
  E('main').innerHTML=h;
  wireJobRowDrop();
}
// Drop a quote document straight onto a job row - it attaches to that job and
// opens the form so the quote value can be checked before saving.
function wireJobRowDrop(){
  document.querySelectorAll('tr.jobrow').forEach(tr=>{
    const stop=e=>{e.preventDefault();e.stopPropagation();};
    ['dragenter','dragover'].forEach(ev=>tr.addEventListener(ev,e=>{stop(e);tr.classList.add('dragover');}));
    ['dragleave','dragend'].forEach(ev=>tr.addEventListener(ev,e=>{stop(e);tr.classList.remove('dragover');}));
    tr.addEventListener('drop',async e=>{
      stop(e); tr.classList.remove('dragover');
      const f=e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0];
      if(!f)return;
      const p=await dropToPath(f); if(!p)return;
      jobModal(Number(tr.getAttribute('data-job')));
      await scanPath(p,quoteFill);
    });
  });
}
function jobModal(id){
  const j=id?_jobs.find(x=>x.id===id):null;
  const types=['Project','Non-project'].map(t=>`<option>${t}</option>`).join('');
  const stat=['Active','Closed'].map(t=>`<option>${t}</option>`).join('');
  window._attachSrc='';
  openModal(j?'Edit job reference':'Add job reference',
    `<label>Job reference</label><input id="j_ref">
     <label>Job title</label><input id="j_title">
     <label>Client</label><input id="j_client">
     <label>Quote (£)</label><input id="j_quote" type="number" step="0.01" min="0">
     <div style="display:flex;gap:6px;margin-top:8px"><button class="btn sm" onclick="scanQuote()">Scan quote…</button><button class="btn sm ghost" onclick="attachOnly()">Attach only</button>${j&&j.quote_att?` <button class="btn sm ghost" onclick="openAtt('${jsstr(j.quote_att_path)}')">📎 View</button>`:''}</div>
     ${dzHtml()}
     <div id="scanmsg" class="muted" style="font-size:12px;min-height:16px"></div>
     <div class="muted" style="font-size:12px">Quote document: <span id="attname">${j&&j.quote_att?'(already attached)':'(none)'}</span></div>
     <label>Type</label><select id="j_kind">${types}</select>
     <label>Status</label><select id="j_status">${stat}</select>
     <label>Notes</label><input id="j_notes">`,()=>saveJob(j));
  if(j){E('j_ref').value=j.ref;E('j_title').value=j.title;E('j_client').value=j.client;E('j_quote').value=j.quote||'';E('j_kind').value=j.type;E('j_status').value=j.status;E('j_notes').value=j.notes;}
  dzWire(quoteFill);
}
function quoteFill(r){
  // scanning the quote document fills the quote value if it isn't set yet
  if(r.value!=null&&!E('j_quote').value)E('j_quote').value=r.value;
  if(r.supplier&&!E('j_client').value)E('j_client').value=r.supplier;
}
function scanQuote(){scanInto(quoteFill);}
async function saveJob(j){
  const ref=E('j_ref').value.trim();if(!ref){alert('Job reference is required.');return;}
  const keep=j?(j.quote_att_path||''):'';
  const a=[ref,E('j_title').value,E('j_client').value,E('j_kind').value,E('j_status').value,E('j_notes').value,E('j_quote').value,window._attachSrc||''];
  const r=j?await API.update_job(j.id,...a,keep):await API.add_job(...a);
  if(!r.ok){alert(r.error);return;}
  closeModal();invOpts();renderJobs();
}
async function delJob(id){if(!confirm('Delete this job reference? If it has time booked, it will be Closed instead.'))return;const r=await API.delete_job(id);if(!r.ok){alert(r.error);return;}renderJobs();}

// ===== Job Roles =====
async function renderRoles(){
  const rows=await API.roles();
  let h=head('Job Roles')+'<div class="card" style="padding:6px 8px"><table><thead><tr><th>Role</th><th>Enabled</th><th></th></tr></thead><tbody>';
  rows.forEach(r=>h+=`<tr><td>${esc(r.name)}</td><td>${r.enabled?'Yes':'No'}</td><td style="text-align:right"><button class="btn sm ghost" onclick="toggleRole(${r.id},${r.enabled?0:1})">${r.enabled?'Disable':'Enable'}</button></td></tr>`);
  h+='</tbody></table></div><div class="toolbar"><button class="btn sm" onclick="addRole()">Add role</button></div>';
  E('main').innerHTML=h;
}
async function addRole(){const n=prompt('New role name:');if(!n)return;await API.add_role(n);invOpts();renderRoles();}
async function toggleRole(id,en){await API.toggle_role(id,!!en);renderRoles();}

// ===== Expenses =====
let _exp=[];
async function renderExpenses(){
  _exp=await API.expenses();
  const mgr=['Manager','Admin'].includes(USER.user_type);
  let h=head('Expenses',mgr?'All staff expenses.':'');
  h+='<div class="card" style="padding:6px 8px"><table><thead><tr><th>Date</th>'+(mgr?'<th>Person</th>':'')+'<th>Project</th><th>Description</th><th class="tnum">Value</th><th>Status</th><th>Receipt</th><th>Payment</th><th></th></tr></thead><tbody>';
  if(!_exp.length)h+='<tr><td colspan="9" class="muted">No expenses yet.</td></tr>';
  _exp.forEach(e=>h+=`<tr><td>${esc(e.date)}</td>${mgr?`<td>${esc(e.who)}</td>`:''}<td>${esc(e.project)}</td><td>${esc(e.desc)}</td><td class="tnum">£${Number(e.value).toFixed(2)}</td><td>${badge(e.status)}</td>
    <td>${e.receipt?`<button class="btn sm ghost" onclick="API.open_file('${jsstr(e.receipt_path)}')">open</button>`:'—'}</td>
    <td>${e.payment?`<button class="btn sm ghost" onclick="API.open_file('${jsstr(e.payment_path)}')">open</button>`:'—'}</td>
    <td style="text-align:right;white-space:nowrap"><button class="btn sm ghost" onclick="expModal(${e.id})">Edit</button><button class="btn sm ghost" onclick="delExp(${e.id})">Del</button>${USER.user_type==='Admin'?`<button class="btn sm" onclick="attachPay(${e.id})">Pay</button>`:''}</td></tr>`);
  h+='</tbody></table></div><div class="toolbar"><button class="btn sm" onclick="expModal(0)">Add expense</button></div>';
  E('main').innerHTML=h;
}
let _expRc=null;
function expModal(id){
  const e=id?_exp.find(x=>x.id===id):null;_expRc=null;
  opts().then(o=>{
    const jobs='<option value="">(none)</option>'+o.active_jobs.map(j=>`<option value="${j.id}">${esc(j.label)}</option>`).join('');
    openModal(e?'Edit expense':'Add expense',
     `<label>Date (YYYY-MM-DD)</label><input id="e_date">
      <label>Project</label><select id="e_job">${jobs}</select>
      <label>Description</label><input id="e_desc">
      <label>Value (£)</label><input id="e_val" type="number" step="0.01" min="0">
      <label>Receipt</label><div><button class="btn sm ghost" onclick="pickReceipt()">Browse…</button> <span id="e_rc" class="muted">${e&&e.receipt?'(current kept)':'(none)'}</span></div>`,()=>saveExp(e));
    if(e){E('e_date').value=e.date;E('e_job').value=e.job_ref_id||'';E('e_desc').value=e.desc;E('e_val').value=e.value;}
    else E('e_date').value=new Date().toISOString().slice(0,10);
  });
}
async function pickReceipt(){const p=await API.pick_file();if(p){_expRc=p;E('e_rc').textContent=p.split(/[\\/]/).pop();}}
async function saveExp(e){
  const v=parseFloat(E('e_val').value);if(isNaN(v)){alert('Value must be a number.');return;}
  let r;
  if(e)r=await API.update_expense(e.id,E('e_date').value,E('e_job').value||null,E('e_desc').value,v,_expRc||'',e.receipt_path||'');
  else r=await API.add_expense(E('e_date').value,E('e_job').value||null,E('e_desc').value,v,_expRc||'');
  if(!r.ok){alert(r.error);return;}closeModal();renderExpenses();
}
async function delExp(id){if(!confirm('Delete this expense?'))return;await API.delete_expense(id);renderExpenses();}
async function attachPay(id){const p=await API.pick_file();if(!p)return;const r=await API.attach_payment(id,p);if(!r.ok){alert(r.error);return;}alert('Payment receipt attached — the employee can now see it.');renderExpenses();}

// ===== Job Costs (with Set rates) =====
// ===== date range picker: click the start day, then the end day ============
// Dates are handled as local y-m-d throughout - never via Date parsing of the
// ISO string, which would shift by a day for anyone west of UTC.
function dISO(d){return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}
function dParse(s){const p=String(s||'').trim().split('-').map(Number);
  return (p.length===3&&p[0]&&p[1]&&p[2])?new Date(p[0],p[1]-1,p[2]):null;}
function dPretty(s){const d=dParse(s); if(!d)return '—';
  return d.toLocaleDateString('en-GB',{day:'numeric',month:'short',year:'numeric'});}
let _drp=null;
function openRange(fromId,toId,onApply){
  const f=dParse(E(fromId)?E(fromId).value:''), t=dParse(E(toId)?E(toId).value:'');
  const view=f?new Date(f.getFullYear(),f.getMonth(),1)
              :new Date(new Date().getFullYear(),new Date().getMonth(),1);
  _drp={fromId:fromId,toId:toId,from:f?dISO(f):'',to:t?dISO(t):'',view:view,onApply:onApply};
  let o=E('drpick');
  if(!o){o=document.createElement('div');o.id='drpick';o.className='drp-bg';document.body.appendChild(o);
    o.addEventListener('mousedown',e=>{if(e.target===o)drpClose();});}
  drpDraw();
}
function drpClose(){const o=E('drpick'); if(o)o.remove(); _drp=null;}
function drpNav(n){_drp.view=new Date(_drp.view.getFullYear(),_drp.view.getMonth()+n,1);drpDraw();}
function drpPick(iso){
  if(!_drp.from||(_drp.from&&_drp.to)){_drp.from=iso;_drp.to='';}    // start again
  else if(iso<_drp.from){_drp.to=_drp.from;_drp.from=iso;}           // picked earlier
  else _drp.to=iso;
  drpDraw();
}
function drpPreset(k){
  const n=new Date(), y=n.getFullYear(), m=n.getMonth();
  let a,b;
  if(k==='tw'){a=new Date(n);a.setDate(n.getDate()-((n.getDay()+6)%7));b=new Date(a);b.setDate(a.getDate()+6);}
  else if(k==='lw'){a=new Date(n);a.setDate(n.getDate()-((n.getDay()+6)%7)-7);b=new Date(a);b.setDate(a.getDate()+6);}
  else if(k==='tm'){a=new Date(y,m,1);b=new Date(y,m+1,0);}
  else if(k==='lm'){a=new Date(y,m-1,1);b=new Date(y,m,0);}
  else if(k==='l30'){b=new Date(n);a=new Date(n);a.setDate(n.getDate()-29);}
  else if(k==='l90'){b=new Date(n);a=new Date(n);a.setDate(n.getDate()-89);}
  else if(k==='ty'){a=new Date(y,0,1);b=new Date(y,11,31);}
  else if(k==='fy'){const s=(m>=3)?y:y-1;a=new Date(s,3,1);b=new Date(s+1,2,31);}  // UK tax year
  else return;
  _drp.from=dISO(a);_drp.to=dISO(b);_drp.view=new Date(a.getFullYear(),a.getMonth(),1);
  drpDraw();
}
function drpMonth(view){
  const y=view.getFullYear(), m=view.getMonth();
  const first=new Date(y,m,1), start=((first.getDay()+6)%7);   // Monday-first
  const days=new Date(y,m+1,0).getDate(), today=dISO(new Date());
  let s=`<div class="drp-cal"><div class="mname">${first.toLocaleDateString('en-GB',{month:'long',year:'numeric'})}</div><div class="drp-grid">`;
  ['M','T','W','T','F','S','S'].forEach(w=>s+=`<div class="wd">${w}</div>`);
  for(let i=0;i<start;i++)s+='<div class="drp-d out"></div>';
  for(let d=1;d<=days;d++){
    const iso=dISO(new Date(y,m,d));
    const isFrom=iso===_drp.from, isTo=iso===_drp.to;
    const inRange=_drp.from&&_drp.to&&iso>_drp.from&&iso<_drp.to;
    const cls=['drp-d',(isFrom||isTo)?'sel':'',inRange?'in':'',iso===today?'today':''].filter(Boolean).join(' ');
    s+=`<div class="${cls}" onclick="drpPick('${iso}')">${d}</div>`;
  }
  return s+'</div></div>';
}
function drpDraw(){
  const v=_drp.view, nxt=new Date(v.getFullYear(),v.getMonth()+1,1);
  const both=_drp.from&&_drp.to;
  const days=both?(diffDaysISO(_drp.from,_drp.to)+1):0;
  E('drpick').innerHTML=`<div class="drp" onmousedown="event.stopPropagation()">
    <div class="drp-side">
      <div class="lab">QUICK PICK</div>
      <button onclick="drpPreset('tw')">This week</button>
      <button onclick="drpPreset('lw')">Last week</button>
      <button onclick="drpPreset('tm')">This month</button>
      <button onclick="drpPreset('lm')">Last month</button>
      <button onclick="drpPreset('l30')">Last 30 days</button>
      <button onclick="drpPreset('l90')">Last 90 days</button>
      <button onclick="drpPreset('ty')">This year</button>
      <button onclick="drpPreset('fy')">This tax year</button>
    </div>
    <div class="drp-main">
      <div class="drp-head">
        <button class="drp-nav" onclick="drpNav(-1)">‹</button>
        <span class="muted" style="font-size:12px">${_drp.from&&!_drp.to?'Now click the end date':'Click the start date, then the end date'}</span>
        <button class="drp-nav" onclick="drpNav(1)">›</button>
      </div>
      <div class="drp-months">${drpMonth(v)}${drpMonth(nxt)}</div>
      <div class="drp-foot">
        <span class="drp-sel">${dPretty(_drp.from)} &rarr; ${dPretty(_drp.to)}${both?` <span class="muted" style="font-weight:400">(${days} day${days===1?'':'s'})</span>`:''}</span>
        <span style="margin-left:auto;display:flex;gap:8px">
          <button class="btn sm ghost" onclick="drpClose()">Cancel</button>
          <button class="btn sm" onclick="drpApply()" ${both?'':'disabled style="opacity:.5;cursor:default"'}>Apply</button>
        </span>
      </div>
    </div></div>`;
}
function drpApply(){
  if(!_drp.from||!_drp.to)return;
  const f=E(_drp.fromId), t=E(_drp.toId), cb=_drp.onApply;
  if(f)f.value=_drp.from; if(t)t.value=_drp.to;
  drpClose();
  if(cb)cb();
}
function money(v){const n=Number(v||0);return (n<0?'−£':'£')+Math.abs(n).toLocaleString('en-GB',{minimumFractionDigits:2,maximumFractionDigits:2});}
function moneyS(v){const n=Number(v||0);const whole=Math.abs(n)%1===0;return (n<0?'−£':'£')+Math.abs(n).toLocaleString('en-GB',{minimumFractionDigits:whole?0:2,maximumFractionDigits:2});}
function money0(v){return '£'+Math.round(Number(v||0)).toLocaleString('en-GB');}
let _jcRows=[],_jcStatus='Active';
async function renderJobCosts(){
  _jcRows=await API.job_costs();
  drawJobCosts();
}
function drawJobCosts(){
  const nA=_jcRows.filter(r=>r.status==='Active').length, nC=_jcRows.length-nA;
  const rows=_jcStatus==='All'?_jcRows.slice():_jcRows.filter(r=>r.status===_jcStatus);
  const adm=USER.user_type==='Admin';
  let h=head('Job Costs','Labour = the hours staff originally submitted × each person\'s rate — not a manager\'s later adjustments. Only submitted/approved timesheets count. Invoiced = what we have billed. Left to invoice = quote + variations − invoiced.');
  h+=`<div class="toolbar" style="margin:0 0 14px"><label class="muted" style="align-self:center">Show:</label>
      <select id="jc_st" onchange="_jcStatus=this.value;drawJobCosts()">${statusOpts(_jcStatus,nA,nC,_jcRows.length)}</select>
      <span class="muted" style="align-self:center;margin-left:auto">Totals below cover the ${_jcStatus==='All'?'':_jcStatus.toLowerCase()+' '}jobs shown (${rows.length})</span></div>`;
  // invoicing position across the jobs currently shown
  const tQ=rows.reduce((a,b)=>a+(b.quote_current||0),0), tI=rows.reduce((a,b)=>a+(b.invoiced||0),0);
  const tT=rows.reduce((a,b)=>a+(b.to_invoice||0),0), tU=rows.reduce((a,b)=>a+(b.invoiced_unpaid||0),0);
  const tB=rows.reduce((a,b)=>a+(b.bills||0),0), tBU=rows.reduce((a,b)=>a+(b.bills_unpaid||0),0);
  h+='<div class="grid kpis" style="margin-bottom:16px">'+kpi('Quoted (incl. variations)',money0(tQ))+kpi('Invoiced',money0(tI))+kpi('Left to invoice',money0(tT),1)+kpi('Bills received',money0(tB))+kpi('Invoiced, unpaid',money0(tU))+kpi('Bills unpaid',money0(tBU))+'</div>';
  if(adm)h+=`<div class="toolbar" style="margin:0 0 14px"><button class="btn sm ghost" onclick="invTemplate()">Blank invoice sheet</button>
    <button class="btn sm ghost" onclick="importInvoices()">Import invoices…</button><span id="invimp" class="muted" style="align-self:center;margin-left:8px"></span></div>`;
  h+='<div class="card" style="padding:6px 8px;overflow-x:auto"><table><thead><tr><th>Job Ref</th><th>Title</th><th class="tnum">Quote</th><th class="tnum">Hours</th><th class="tnum">Labour</th><th class="tnum">Bills</th><th class="tnum">Total cost</th><th class="tnum">Invoiced</th><th class="tnum">To invoice</th><th class="tnum">Profit</th><th>Cost vs invoiced</th><th></th></tr></thead><tbody>';
  if(!rows.length)h+=`<tr><td colspan="12" class="muted">No ${_jcStatus==='All'?'':_jcStatus.toLowerCase()+' '}jobs to cost${_jcStatus==='Closed'?' — nothing has been closed yet.':'.'}</td></tr>`;
  rows.forEach(d=>{const pct=d.pct||0,q=d.quote_current||0;
    // The bar measures TOTAL COST against what has actually been INVOICED, so
    // it answers "are we making money on the work we've billed?" - over 100%
    // means cost has overtaken revenue. The small grey figure underneath keeps
    // the old view: how much of the quote the cost has consumed.
    const cpi=d.pct_cost_of_invoiced||0;
    const ccls=cpi>100?'over':(cpi>=85?'warn':'');
    const bar=d.invoiced
      ? `<div class="bar-wrap" title="Cost ${money(d.total_cost||0)} against invoiced ${money(d.invoiced||0)}"><div class="bar-fill ${ccls}" style="width:${Math.min(100,cpi)}%"></div></div> <span class="muted">${cpi}%</span>${q?`<div class="muted" style="font-size:11px">${pct}% of quote spent</div>`:''}`
      : (q?`<span class="muted">not billed yet</span><div class="muted" style="font-size:11px">${pct}% of quote spent</div>`:'<span class="muted">—</span>');
    const qc=q?money(q):'—';
    const qsub=d.variations?`<div class="muted" style="font-size:11px">${money0(d.quote)} + ${money0(d.variations)} var</div>`:'';
    const ti=d.to_invoice||0;
    const tic=q?`<span style="color:${ti>0.005?'#b8860b':(ti<-0.005?'#c0392b':'#1d8f3b')};font-weight:600">${money(ti)}</span>`:'<span class="muted">—</span>';
    const pf=d.profit||0;
    const pfc=d.invoiced?`<span style="color:${pf<0?'#c0392b':'#1d8f3b'};font-weight:600">${money(pf)}</span>${d.margin?`<div class="muted" style="font-size:11px">${d.margin}%</div>`:''}`:'<span class="muted">—</span>';
    // when closed jobs are on screen, say which ones they are
    const cl=(d.status==='Closed'&&_jcStatus!=='Closed')?' <span class="badge mini b-Closed">Closed</span>':'';
    h+=`<tr><td>${esc(d.job_ref)}${cl}</td><td>${esc(d.job_title)}</td><td class="tnum">${qc}${qsub}</td><td class="tnum">${d.hours}</td><td class="tnum">${money(d.cost)}</td><td class="tnum">${d.bills?money(d.bills):'<span class="muted">—</span>'}</td><td class="tnum"><b>${money(d.total_cost)}</b></td><td class="tnum">${d.invoiced?money(d.invoiced):'<span class="muted">—</span>'}</td><td class="tnum">${tic}</td><td class="tnum">${pfc}</td><td>${bar}</td><td style="text-align:right;white-space:nowrap"><button class="btn sm ghost" onclick="renderBilling(${d.id})">Money</button> <button class="btn sm ghost" onclick="jobCostPdf(${d.id})">PDF</button> <button class="btn sm ghost" onclick="ratesModal(${d.id},'${jsstr(d.job_ref)}')">Set rates</button></td></tr>`;});
  h+='</tbody></table></div><div class="sub" style="margin-top:8px">Labour is priced at the rate frozen when each timesheet was submitted, so putting rates up later never re-prices work already done. Total cost = labour + bills. Profit = invoiced − total cost. The bar shows cost against what you have actually invoiced: <b>green</b> = comfortably profitable, <b>amber</b> = 85%+ of the billed money spent, <b>red</b> = cost has overtaken what you have billed. The small figure underneath is how much of the quote the cost has consumed.</div>';
  E('main').innerHTML=h;
}
async function invTemplate(){
  const r=await API.export_invoice_template();
  if(r&&r.ok){await API.open_file(r.path);}else if(r&&r.error)alert(r.error);
}
async function importInvoices(){
  const o=E('invimp'); if(o)o.textContent='Reading file…';
  const r=await API.import_invoices();
  if(!r){if(o)o.textContent='';return;}
  if(!r.ok){if(o)o.textContent='';alert((r.error||'Import failed')+(r.details?'\n\n'+r.details.join('\n'):''));return;}
  let msg=`Imported ${r.added} invoice(s), ${money(r.total)}.`;
  if(r.duplicates)msg+=` ${r.duplicates} already existed and were skipped.`;
  if(r.errors&&r.errors.length)msg+=`\n\nSkipped ${r.errors.length} row(s):\n`+r.errors.join('\n');
  alert(msg); toast(r.added+' invoice(s) imported');
  renderJobCosts();
}
// ===== Invoicing detail for one job =====
let BILL=null;
async function renderBilling(jobId){
  const d=await API.job_billing(jobId);
  if(d&&d.error){alert(d.error);return;}
  BILL=d; drawBilling();
}
function drawBilling(){
  const d=BILL, j=d.job, ed=d.can_edit;
  let h=head('Money – '+esc(j.job_ref), esc(j.job_title));
  h+=`<div class="toolbar" style="margin:0 0 14px"><button class="btn sm ghost" onclick="renderJobCosts()">‹ Back to job costs</button></div>`;
  h+='<div class="grid kpis" style="margin-bottom:18px">'+
     kpiDoc('Original quote',money0(j.quote),j.quote_doc?j.quote_doc_path:'')+
     kpi('Variations',(j.variations<0?'-':'')+money0(Math.abs(j.variations)))+
     kpi('Current quote',money0(j.quote_current))+
     kpi('Invoiced',money0(j.invoiced))+
     kpi('Left to invoice',money0(j.to_invoice),1)+'</div>';
  h+='<div class="grid kpis" style="margin-bottom:18px">'+
     kpi('Labour cost',money0(j.cost))+
     kpi('Bills',money0(j.bills))+
     kpi('Total cost',money0(j.total_cost))+
     kpi('Profit so far',(j.profit<0?'-':'')+money0(Math.abs(j.profit)))+
     kpi('Margin',(j.invoiced?j.margin+'%':'—'))+'</div>';
  // variations
  h+='<div class="card" style="margin-bottom:16px"><div class="sec-t">Variations <span class="muted" style="font-weight:400;font-size:12px">— added to the original quote. Use a negative value for an omission.</span></div>';
  h+='<table><thead><tr><th>Ref</th><th>Date</th><th>Description</th><th class="tnum">Value</th><th></th></tr></thead><tbody>';
  if(!d.variations.length)h+='<tr><td colspan="5" class="muted">No variations — the quote is still as first entered.</td></tr>';
  d.variations.forEach(v=>{const doc=v.att?` <button class="btn sm ghost" title="Open the attached document" onclick="openAtt('${jsstr(v.att_path)}')">📎</button>`:'';
    h+=`<tr><td>${esc(v.ref)||'<span class="muted">—</span>'}</td><td>${esc(v.date)}</td><td>${esc(v.desc)}</td><td class="tnum" style="color:${v.value<0?'#c0392b':'#1d8f3b'}">${money(v.value)}</td><td style="text-align:right;white-space:nowrap">${doc}${ed?` <button class="btn sm ghost" onclick="varModal(${v.id})">Edit</button> <button class="xdel" onclick="delVariation(${v.id})">✕</button>`:''}</td></tr>`;});
  h+=`</tbody><tfoot><tr><td colspan="3"><b>Total variations</b></td><td class="tnum"><b>${money(j.variations)}</b></td><td></td></tr></tfoot></table>`;
  if(ed)h+='<div style="margin-top:10px"><button class="btn sm" onclick="varModal(0)">Add variation</button></div>';
  h+='</div>';
  // invoices
  h+='<div class="card"><div class="sec-t">Invoices <span class="muted" style="font-weight:400;font-size:12px">— what has been billed to the client. A credit note is a negative value.</span></div>';
  h+='<table><thead><tr><th>Invoice No</th><th>Date</th><th>Description</th><th class="tnum">Value</th><th>Status</th><th></th></tr></thead><tbody>';
  if(!d.invoices.length)h+='<tr><td colspan="6" class="muted">Nothing invoiced against this job yet.</td></tr>';
  d.invoices.forEach(i=>{const doc=i.att?` <button class="btn sm ghost" title="Open the attached document" onclick="openAtt('${jsstr(i.att_path)}')">📎</button>`:'';
    h+=`<tr><td>${esc(i.no)}</td><td>${esc(i.date)}</td><td>${esc(i.desc)}</td><td class="tnum">${money(i.value)}</td><td>${badge(i.status)}${i.status==='Paid'&&i.paid_date?` <span class="muted" style="font-size:11px">${esc(i.paid_date)}</span>`:''}</td><td style="text-align:right;white-space:nowrap">${doc}${ed?` <button class="btn sm ghost" onclick="togglePaid(${i.id},${i.status==='Paid'?'false':'true'})">${i.status==='Paid'?'Mark unpaid':'Mark paid'}</button> <button class="btn sm ghost" onclick="invModal(${i.id})">Edit</button> <button class="xdel" onclick="delInvoice(${i.id})">✕</button>`:''}</td></tr>`;});
  h+=`</tbody><tfoot><tr><td colspan="3"><b>Total invoiced</b></td><td class="tnum"><b>${money(j.invoiced)}</b></td><td class="muted">${money(j.invoiced_unpaid)} unpaid</td><td></td></tr></tfoot></table>`;
  if(ed)h+='<div style="margin-top:10px"><button class="btn sm" onclick="invModal(0)">Add invoice</button></div>';
  h+='</div>';
  // bills (money out)
  h+='<div class="card" style="margin-top:16px"><div class="sec-t">Bills <span class="muted" style="font-weight:400;font-size:12px">— supplier and subcontractor costs. These count towards the job\'s total cost.</span></div>';
  h+='<table><thead><tr><th>Supplier</th><th>Bill No</th><th>Date</th><th>Description</th><th class="tnum">Value</th><th>Status</th><th></th></tr></thead><tbody>';
  if(!d.bills.length)h+='<tr><td colspan="7" class="muted">No bills against this job yet.</td></tr>';
  d.bills.forEach(b=>{const doc=b.att?` <button class="btn sm ghost" title="Open the attached document" onclick="openAtt('${jsstr(b.att_path)}')">📎</button>`:'';
    h+=`<tr><td>${esc(b.supplier)}</td><td>${esc(b.no)||'<span class="muted">—</span>'}</td><td>${esc(b.date)}</td><td>${esc(b.desc)}${b.category?` <span class="muted">(${esc(b.category)})</span>`:''}</td><td class="tnum">${money(b.value)}</td><td>${badge(b.status)}${b.status==='Paid'&&b.paid_date?` <span class="muted" style="font-size:11px">${esc(b.paid_date)}</span>`:''}</td><td style="text-align:right;white-space:nowrap">${doc}${ed?` <button class="btn sm ghost" onclick="toggleBillPaid(${b.id},${b.status==='Paid'?'false':'true'})">${b.status==='Paid'?'Mark unpaid':'Mark paid'}</button> <button class="btn sm ghost" onclick="billModal(${b.id})">Edit</button> <button class="xdel" onclick="delBill(${b.id})">✕</button>`:''}</td></tr>`;});
  h+=`</tbody><tfoot><tr><td colspan="4"><b>Total bills</b></td><td class="tnum"><b>${money(j.bills)}</b></td><td class="muted">${money(j.bills_unpaid)} unpaid</td><td></td></tr></tfoot></table>`;
  if(ed)h+='<div style="margin-top:10px"><button class="btn sm" onclick="billModal(0)">Add bill</button> <span class="muted" style="font-size:12px;margin-left:6px">Tip: attach the document and press Scan to fill the form in for you.</span></div>';
  h+='</div>';
  h+=`<div class="sub" style="margin-top:12px">Left to invoice = current quote (${money(j.quote_current)}) − invoiced (${money(j.invoiced)}) = <b>${money(j.to_invoice)}</b><br>
      Total cost = labour (${money(j.cost)}) + bills (${money(j.bills)}) = <b>${money(j.total_cost)}</b> &nbsp;·&nbsp; Profit = invoiced − total cost = <b>${money(j.profit)}</b></div>`;
  h+='<div id="docwrap"></div>';
  E('main').innerHTML=h;
  if(ed)wirePageDrop();
  loadDocs();
}
// ===== the job's document folder =====
let DOCS=null, _docFilter='All';
// Which screen is showing the document list right now: the Money page or a
// job folder. Anything that changes documents refreshes whichever is open.
let _docCtx={jobId:null,mode:null};
async function loadDocs(){
  if(!BILL||!BILL.job)return;
  _docCtx={jobId:BILL.job.id,mode:'money'};
  const d=await API.job_documents(BILL.job.id);
  if(d&&d.error)return;
  DOCS=d; drawDocs();
}
async function refreshDocs(){
  if(_docCtx.mode==='folder'&&_docCtx.jobId){
    await loadFolders(); openFolder(_docCtx.jobId); return;
  }
  if(BILL&&BILL.job)loadDocs();
}
function fsize(n){ if(!n)return ''; const u=['B','KB','MB','GB']; let i=0,v=n;
  while(v>=1024&&i<u.length-1){v/=1024;i++;}
  return (i?v.toFixed(v<10?1:0):v)+' '+u[i]; }
// Icons say what the document IS, not what file type it happens to be -
// nearly everything here is a PDF, so a PDF icon told you nothing. Invoices
// (money in) and bills (money out) get mirrored arrows so they read apart at
// a glance.
const DOC_SVG={
  Quote:'<path d="M20.6 13.4 12 22l-9-9V4h9l8.6 8.6a2 2 0 0 1 0 2.8z"/><circle cx="7.5" cy="7.5" r="1.4"/>',
  Variation:'<path d="M17 3a2.85 2.85 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5z"/><path d="M15 5l4 4"/>',
  Invoice:'<path d="M3 15v4a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-4"/><path d="M12 3v11"/><path d="m8 10 4 4 4-4"/>',
  Bill:'<path d="M3 15v4a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-4"/><path d="M12 14V3"/><path d="m8 7 4-4 4 4"/>',
  Document:'<path d="M21.4 11.05 12.25 20.2a5.5 5.5 0 0 1-7.78-7.78l9.2-9.19a3.67 3.67 0 1 1 5.18 5.18l-9.2 9.2a1.83 1.83 0 1 1-2.59-2.59l8.49-8.48"/>',
  Missing:'<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/>'
};
function docIcon(kind,ok){
  const k=ok?(DOC_SVG[kind]?kind:'Document'):'Missing';
  return `<div class="doc-ico ${ok?'di-'+(DOC_SVG[kind]?kind:'Document'):'di-gone'}" title="${esc(ok?kind:'File not found')}">`
    +`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${DOC_SVG[k]}</svg></div>`;
}
function docFilter(k){_docFilter=k;drawDocs();}
// Shared by the Money screen and the Folders browser so the two can never
// drift apart.
function docRowsHtml(rows,canEdit){
  let h='<div class="doc-list">';
  rows.forEach(d=>{
    const meta=[`<span class="doc-tag dt-${d.kind}">${esc(d.kind)}</span>`];
    if(d.ref)meta.push(esc(d.ref));
    if(d.date)meta.push(dPretty(d.date));
    if(d.value!=null&&d.value!=='')meta.push(money(d.value));
    if(d.exists&&d.size)meta.push(fsize(d.size));
    if(d.desc)meta.push('<span style="opacity:.85">'+esc(d.desc.slice(0,70))+(d.desc.length>70?'…':'')+'</span>');
    h+=`<div class="doc-row ${d.exists?'':'gone'}">
      ${docIcon(d.kind,d.exists)}
      <div class="doc-main">
        <div class="doc-name" title="${esc(d.name)}">${esc(d.name)}</div>
        <div class="doc-meta">${meta.join('<span style="opacity:.4">·</span>')}</div>
      </div>
      <div class="doc-act">
        ${d.exists?`<button class="btn xs" onclick="openAtt('${jsstr(d.path)}')">Open</button>
           <button class="btn xs ghost" title="Show this file in Windows Explorer" onclick="API.open_folder('${jsstr(d.path)}')">Folder</button>`
          :'<span class="muted" style="font-size:11.5px">file not found</span>'}
        ${(canEdit&&d.doc_id)?`<button class="xdel" title="Remove from this list" onclick="delDoc(${d.doc_id})">✕</button>`:''}
      </div></div>`;});
  return h+'</div>';
}
function drawDocs(){
  const w=E('docwrap'); if(!w||!DOCS)return;
  const kinds=['All','Quote','Variation','Invoice','Bill','Document'];
  const counts={}; DOCS.docs.forEach(d=>counts[d.kind]=(counts[d.kind]||0)+1);
  const rows=DOCS.docs.filter(d=>_docFilter==='All'||d.kind===_docFilter);
  const n=DOCS.docs.length;
  let h='<div class="card" style="margin-top:16px">'
    +`<div class="sec-t" style="margin-bottom:6px">Documents <span class="muted" style="font-weight:400;font-size:12px">&nbsp;${n} file${n===1?'':'s'} on this job`
    +(DOCS.missing?` &middot; <span style="color:#c0392b;font-weight:600">${DOCS.missing} missing</span>`:'')+'</span></div>';
  // filter pills - only show a type that actually has something
  h+='<div class="doc-filters">'+kinds.filter(k=>k==='All'||counts[k])
      .map(k=>{const c=k==='All'?n:counts[k];
        return `<button class="doc-chip ${_docFilter===k?'on':''}" onclick="docFilter('${k}')">${k}<span style="opacity:.7"> ${c}</span></button>`;}).join('')+'</div>';
  if(!n){
    h+='<div class="doc-empty">Nothing attached to this job yet.<br><span style="font-size:12px">Attach files to the quote, a variation, an invoice or a bill — or drop them below to file them straight against the job.</span></div>';
  }else if(!rows.length){
    h+='<div class="doc-empty">No '+esc(_docFilter.toLowerCase())+' documents.</div>';
  }else{
    h+=docRowsHtml(rows,DOCS.can_edit);
  }
  if(DOCS.can_edit){
    h+='<div id="docdrop" class="dropzone" style="margin-top:14px">Drag files here to file them against this job</div>'
      +`<div style="display:flex;gap:8px;align-items:center;margin-top:8px;flex-wrap:wrap">
         <button class="btn sm" onclick="pickDoc()">Add document…</button>
         <button class="btn sm ghost" onclick="API.open_folder('${jsstr(DOCS.folder)}')">Open documents folder</button>
         <span class="muted" style="font-size:11.5px">Stored in ${esc(DOCS.folder)}</span></div>`;
  }
  h+='</div>';
  w.innerHTML=h;
  if(DOCS.can_edit)wireDocDrop();
}
// Shared by every place you can drop a document: the Money screen, a folder
// tile on the dashboard, and inside an open folder.
async function fileFilesToJob(jobId,fs){
  let n=0;
  for(let i=0;i<fs.length;i++){                       // whole batches at once
    const p=await dropToPath(fs[i]); if(!p)continue;
    const r=await API.add_document(jobId,fs[i].name||'','','',p);
    if(r&&r.ok)n++;
    else if(r&&r.error){alert(r.error);break;}
  }
  if(n)toast(n+' document(s) filed');
  return n;
}
function wireDrop(el,onFiles,busyText){
  if(!el)return;
  const stop=e=>{e.preventDefault();e.stopPropagation();};
  ['dragenter','dragover'].forEach(ev=>el.addEventListener(ev,e=>{stop(e);el.classList.add('over');}));
  ['dragleave','dragend'].forEach(ev=>el.addEventListener(ev,e=>{stop(e);el.classList.remove('over');}));
  el.addEventListener('drop',async e=>{
    stop(e); el.classList.remove('over');
    const fs=e.dataTransfer&&e.dataTransfer.files; if(!fs||!fs.length)return;
    const was=el.textContent;
    if(busyText)el.textContent='Filing '+fs.length+' file(s)…';
    try{ await onFiles(fs); }finally{ if(busyText)el.textContent=was; }
  });
}
function wireDocDrop(){
  wireDrop(E('docdrop'),async fs=>{
    if(!BILL||!BILL.job)return;
    await fileFilesToJob(BILL.job.id,fs);
    loadDocs();
  },true);
}
async function pickDocFor(jobId){
  const p=await API.pick_document(); if(!p)return;
  const r=await API.add_document(jobId,p.split(/[\\/]/).pop(),'','',p);
  if(!r.ok){alert(r.error);return;}
  toast('Document filed'); await loadFolders(); openFolder(jobId);
}
async function pickDoc(){
  if(!BILL||!BILL.job)return;
  const p=await API.pick_document(); if(!p)return;
  const name=p.split(/[\\/]/).pop();
  const r=await API.add_document(BILL.job.id,name,'','',p);
  if(!r.ok){alert(r.error);return;}
  toast('Document filed'); loadDocs();
}
async function delDoc(did){
  if(!confirm('Remove this document from the job?\n\nThe file itself stays in the documents folder.'))return;
  const r=await API.delete_document(did);
  if(!r||!r.ok){alert((r&&r.error)||'Could not remove that document.');return;}
  toast('Removed'); refreshDocs();
}
// ---- drop a document anywhere on this screen ----
function wirePageDrop(){
  const main=E('main'); if(!main)return;
  const stop=e=>{e.preventDefault();e.stopPropagation();};
  const show=on=>{let o=E('pagedrop');
    if(on&&!o){o=document.createElement('div');o.id='pagedrop';o.className='page-drop';
      o.textContent='Drop the document here to add it to this job';document.body.appendChild(o);}
    else if(!on&&o)o.remove();};
  main.addEventListener('dragenter',e=>{stop(e);if(E('modal').classList.contains('hidden'))show(true);});
  main.addEventListener('dragover',e=>{stop(e);});
  main.addEventListener('dragleave',e=>{stop(e);
    if(e.relatedTarget&&main.contains(e.relatedTarget))return;   // still inside
    show(false);});
  main.addEventListener('drop',async e=>{
    stop(e); show(false);
    if(!E('modal').classList.contains('hidden'))return;          // a form is open
    const f=e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0];
    if(!f)return;
    const p=await dropToPath(f);
    if(p)dropChoose(p,f.name||'document');
  });
}
function dropChoose(path,name){
  window._dropPath=path;
  openModal('What is this document?',
    `<p class="sub" style="margin-bottom:12px">${esc(name)}</p>
     <div style="display:flex;flex-direction:column;gap:8px">
       <button class="btn" onclick="dropAs('bill')">Bill — a cost coming in</button>
       <button class="btn" onclick="dropAs('invoice')">Invoice — something we have billed out</button>
       <button class="btn ghost" onclick="dropAs('variation')">Variation — a change to the quote</button>
     </div>`, ()=>closeModal());
  E('modalOk').style.display='none';
}
async function dropAs(kind){
  const p=window._dropPath; window._dropPath='';
  E('modalOk').style.display='';
  closeModal();
  if(kind==='bill'){billModal(0); await scanPath(p,billFill);}
  else if(kind==='invoice'){invModal(0); await scanPath(p,invFill);}
  else {varModal(0); await scanPath(p,varFill);}
}
async function openAtt(p){ if(p) await API.open_file(p); }
// ---- shared document scan helper: pick a file, read it, hand back fields ----
// ---- dropped files: the page only gets the CONTENTS, never the path, so the
// bytes go to Python which writes them to Receipts and hands back a real path.
function _readAsB64(file){
  return new Promise((res,rej)=>{
    const fr=new FileReader();
    fr.onload=()=>{const s=String(fr.result||''); const i=s.indexOf(',');res(i>=0?s.slice(i+1):s);};
    fr.onerror=()=>rej(new Error('Could not read that file.'));
    fr.readAsDataURL(file);
  });
}
async function dropToPath(file){
  if(!file)return null;
  if(file.size>25*1024*1024){alert('That file is too big (25 MB limit).\nUse "Attach only" to link it instead.');return null;}
  try{
    const b64=await _readAsB64(file);
    const r=await API.save_dropped_file(file.name||'document',b64);
    if(!r||!r.ok){alert((r&&r.error)||'Could not save that file.');return null;}
    return r.path;
  }catch(e){alert('Could not read that file:\n\n'+((e&&e.message)?e.message:e));return null;}
}
// A dashed target inside a form. Dropping on it attaches + scans in one go.
function dzHtml(){return '<div id="dropzone" class="dropzone">Drag a PDF or photo here to attach and read it</div>';}
function dzWire(setter){
  const z=E('dropzone'); if(!z)return;
  const stop=e=>{e.preventDefault();e.stopPropagation();};
  ['dragenter','dragover'].forEach(ev=>z.addEventListener(ev,e=>{stop(e);z.classList.add('over');}));
  ['dragleave','dragend'].forEach(ev=>z.addEventListener(ev,e=>{stop(e);z.classList.remove('over');}));
  z.addEventListener('drop',async e=>{
    stop(e); z.classList.remove('over');
    const f=e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0];
    if(!f)return;
    const p=await dropToPath(f);
    if(p)await scanPath(p,setter);
  });
}
async function scanPath(p,setter){
  const st=E('scanmsg');
  window._attachSrc=p;                  // attach it whether or not it reads
  const an=E('attname'); if(an)an.textContent=p.split(/[\\/]/).pop();
  try{
    if(st){st.textContent='Reading document…';st.style.color='#7a8089';}
    const r=await API.scan_document(p);
    if(!r||!r.ok){ if(st){st.style.color='#9a6300';st.textContent='Attached, but could not read it automatically — please type the details in.';}
      if(r&&r.error)alert(r.error);
      return; }
    if(setter)setter(r);
    if(st){
      // Be explicit about VAT: the value filled in must be EX-VAT, because VAT
      // is reclaimable and must never end up inside a job cost.
      let msg,col;
      if(r.value_basis==='net'||r.value_basis==='derived'){
        col='#4f7d12';
        msg=`Read ${r.found} field(s). Filled in ${money(r.value)} EX-VAT`;
        if(r.vat!=null)msg+=` (VAT ${money(r.vat)} excluded`+(r.gross!=null?`, gross ${money(r.gross)}`:'')+')';
        msg+=' — please check before saving.';
      }else if(r.value_basis==='no_vat_shown'){
        col='#4f7d12';
        msg=`Read ${r.found} field(s). Filled in ${money(r.value)} — the document shows no VAT, so this is treated as ex-VAT. Please check.`;
      }else if(r.value_basis==='unknown'){
        col='#c0392b';
        msg=`⚠ Only one total (${money(r.value)}) was found and it is not clear whether it includes VAT. Check the document and enter the EX-VAT figure.`;
      }else{
        col='#9a6300';
        msg=`Read ${r.found} field(s), but no amount could be read — please type the value in.`;
      }
      st.style.color=col; st.textContent=msg;}
  }catch(e){
    if(st){st.textContent='';}
    alert('Could not open that document:\n\n'+((e&&e.message)?e.message:e));
  }
}
async function scanInto(setter){
  try{
    if(!API||typeof API.pick_document!=='function'){alert('This build of the app cannot open documents. Close and reopen the app, and if it persists contact whoever supports the app.');return;}
    const p=await API.pick_document();
    if(!p)return;                       // user cancelled the picker
    await scanPath(p,setter);
  }catch(e){
    alert('Could not open that document:\n\n'+((e&&e.message)?e.message:e));
  }
}
async function attachOnly(){
  try{
    if(!API||typeof API.pick_document!=='function'){alert('This build of the app cannot open documents. Close and reopen the app, and if it persists contact whoever supports the app.');return;}
    const p=await API.pick_document();
    if(!p)return;
    window._attachSrc=p; const an=E('attname'); if(an)an.textContent=p.split(/[\\/]/).pop();
    const st=E('scanmsg'); if(st){st.style.color='#4f7d12';st.textContent='Document attached.';}
  }catch(e){
    alert('Could not attach that document:\n\n'+((e&&e.message)?e.message:e));
  }
}
function billModal(bid){
  const b=bid?BILL.bills.find(x=>x.id===bid):null;
  window._attachSrc='';
  openModal(b?'Edit bill':'Add bill',
    `<div style="display:flex;gap:6px;margin-bottom:4px"><button class="btn sm" onclick="scanBill()">Scan document…</button><button class="btn sm ghost" onclick="attachOnly()">Attach only</button></div>
     ${dzHtml()}
     <div id="scanmsg" class="muted" style="font-size:12px;min-height:16px"></div>
     <div class="muted" style="font-size:12px">Document: <span id="attname">${b&&b.att?'(already attached)':'(none)'}</span></div>
     <label>Supplier</label><input id="b_sup" placeholder="e.g. Acme Electrical Supplies Ltd">
     <label>Bill / invoice number</label><input id="b_no" placeholder="e.g. INV-20418">
     <label>Date</label><input id="b_date" type="date">
     <label>Category</label><input id="b_cat" placeholder="e.g. Materials, Subcontractor, Plant hire">
     <label>Description</label><textarea id="b_desc" rows="2"></textarea>
     <label>Value (£) — net of VAT. Negative for a credit note</label><input id="b_val" type="number" step="0.01">
     <label>Status</label><select id="b_status"><option value="Received">Received</option><option value="Paid">Paid</option></select>`,
    ()=>saveBill(bid));
  if(b){E('b_sup').value=b.supplier;E('b_no').value=b.no;E('b_date').value=b.date;E('b_cat').value=b.category||'';E('b_desc').value=b.desc;E('b_val').value=b.value;E('b_status').value=b.status;}
  else E('b_date').value=new Date().toISOString().slice(0,10);
  dzWire(billFill);
}
function billFill(r){
  if(r.supplier&&!E('b_sup').value)E('b_sup').value=r.supplier;
  if(r.number)E('b_no').value=r.number;
  if(r.date)E('b_date').value=r.date;
  if(r.value!=null)E('b_val').value=r.value;
}
function scanBill(){scanInto(billFill);}
async function saveBill(bid){
  const val=parseFloat(E('b_val').value);
  if(!E('b_sup').value.trim()){alert('Supplier is required.');return;}
  if(isNaN(val)||val===0){alert('Enter a non-zero bill value.');return;}
  const st=E('b_status').value, pd=(st==='Paid')?(E('b_date').value||''):'';
  const keep=bid?(BILL.bills.find(x=>x.id===bid)||{}).att_path||'':'';
  const a=[E('b_sup').value,E('b_no').value,E('b_date').value,E('b_desc').value,val,E('b_cat').value,st,pd,window._attachSrc||''];
  const r=bid?await API.update_bill(bid,...a,keep):await API.add_bill(BILL.job.id,E('b_sup').value,E('b_no').value,E('b_date').value,E('b_desc').value,val,E('b_cat').value,st,pd,window._attachSrc||'');
  if(!r.ok){alert(r.error);return;}
  closeModal(); toast('Bill saved'); renderBilling(BILL.job.id);
}
async function toggleBillPaid(bid,paid){
  const r=await API.set_bill_paid(bid,paid);
  if(!r.ok){alert(r.error);return;}
  renderBilling(BILL.job.id);
}
async function delBill(bid){
  if(!confirm('Delete this bill? It will no longer count towards the job cost.'))return;
  const r=await API.delete_bill(bid);
  if(!r.ok){alert(r.error);return;}
  toast('Bill deleted'); renderBilling(BILL.job.id);
}
function varModal(vid){
  const v=vid?BILL.variations.find(x=>x.id===vid):null;
  window._attachSrc='';
  openModal(v?'Edit variation':'Add variation',
    `<div style="display:flex;gap:6px;margin-bottom:4px"><button class="btn sm" onclick="scanVar()">Scan document…</button><button class="btn sm ghost" onclick="attachOnly()">Attach only</button></div>
     ${dzHtml()}
     <div id="scanmsg" class="muted" style="font-size:12px;min-height:16px"></div>
     <div class="muted" style="font-size:12px">Document: <span id="attname">${v&&v.att?'(already attached)':'(none)'}</span></div>
     <label>Variation ref</label><input id="v_ref" placeholder="e.g. VO-001">
     <label>Date</label><input id="v_date" type="date">
     <label>Description</label><textarea id="v_desc" rows="3" placeholder="What changed and why"></textarea>
     <label>Value (£) — negative for an omission</label><input id="v_val" type="number" step="0.01">`,
    ()=>saveVariation(vid));
  if(v){E('v_ref').value=v.ref;E('v_date').value=v.date;E('v_desc').value=v.desc;E('v_val').value=v.value;}
  else E('v_date').value=new Date().toISOString().slice(0,10);
  dzWire(varFill);
}
function varFill(r){
  if(r.number&&!E('v_ref').value)E('v_ref').value=r.number;
  if(r.date)E('v_date').value=r.date;
  if(r.value!=null)E('v_val').value=r.value;
  if(!E('v_desc').value){
    if(r.supplier)E('v_desc').value=r.supplier;
  }
}
function scanVar(){scanInto(varFill);}
async function saveVariation(vid){
  const val=parseFloat(E('v_val').value);
  if(isNaN(val)){alert('Enter a variation value.');return;}
  if(!E('v_desc').value.trim()){alert('Please describe the variation.');return;}
  const keep=vid?(BILL.variations.find(x=>x.id===vid)||{}).att_path||'':'';
  const a=[E('v_ref').value,E('v_desc').value,val,E('v_date').value,window._attachSrc||''];
  const r=vid?await API.update_variation(vid,...a,keep):await API.add_variation(BILL.job.id,...a);
  if(!r.ok){alert(r.error);return;}
  closeModal(); toast('Variation saved'); renderBilling(BILL.job.id);
}
async function delVariation(vid){
  if(!confirm('Delete this variation? The quote will drop back by its value.'))return;
  const r=await API.delete_variation(vid);
  if(!r.ok){alert(r.error);return;}
  toast('Variation deleted'); renderBilling(BILL.job.id);
}
function invModal(iid){
  const i=iid?BILL.invoices.find(x=>x.id===iid):null;
  window._attachSrc='';
  openModal(i?'Edit invoice':'Add invoice',
    `<div style="display:flex;gap:6px;margin-bottom:4px"><button class="btn sm" onclick="scanInv()">Scan document…</button><button class="btn sm ghost" onclick="attachOnly()">Attach only</button></div>
     ${dzHtml()}
     <div id="scanmsg" class="muted" style="font-size:12px;min-height:16px"></div>
     <div class="muted" style="font-size:12px">Document: <span id="attname">${i&&i.att?'(already attached)':'(none)'}</span></div>
     <label>Invoice number</label><input id="i_no" placeholder="e.g. INV-1042">
     <label>Date</label><input id="i_date" type="date">
     <label>Description</label><textarea id="i_desc" rows="2" placeholder="What this invoice covers"></textarea>
     <label>Value (£) — negative for a credit note</label><input id="i_val" type="number" step="0.01">
     <label>Status</label><select id="i_status"><option value="Raised">Raised</option><option value="Paid">Paid</option></select>`,
    ()=>saveInvoice(iid));
  if(i){E('i_no').value=i.no;E('i_date').value=i.date;E('i_desc').value=i.desc;E('i_val').value=i.value;E('i_status').value=i.status;}
  else E('i_date').value=new Date().toISOString().slice(0,10);
  dzWire(invFill);
}
function invFill(r){
  if(r.number)E('i_no').value=r.number;
  if(r.date)E('i_date').value=r.date;
  if(r.value!=null)E('i_val').value=r.value;
  if(r.supplier&&!E('i_desc').value)E('i_desc').value=r.supplier;
}
function scanInv(){scanInto(invFill);}
async function saveInvoice(iid){
  const val=parseFloat(E('i_val').value);
  if(!E('i_no').value.trim()){alert('Invoice number is required.');return;}
  if(isNaN(val)||val===0){alert('Enter a non-zero invoice value.');return;}
  const st=E('i_status').value, pd=(st==='Paid')?(E('i_date').value||''):'';
  const keep=iid?(BILL.invoices.find(x=>x.id===iid)||{}).att_path||'':'';
  const a=[E('i_no').value,E('i_date').value,E('i_desc').value,val,st,pd,window._attachSrc||''];
  const r=iid?await API.update_invoice(iid,...a,keep):await API.add_invoice(BILL.job.id,...a);
  if(!r.ok){alert(r.error);return;}
  closeModal(); toast('Invoice saved'); renderBilling(BILL.job.id);
}
async function togglePaid(iid,paid){
  const r=await API.set_invoice_paid(iid,paid);
  if(!r.ok){alert(r.error);return;}
  renderBilling(BILL.job.id);
}
async function delInvoice(iid){
  if(!confirm('Delete this invoice? It will no longer count towards invoiced value.'))return;
  const r=await API.delete_invoice(iid);
  if(!r.ok){alert(r.error);return;}
  toast('Invoice deleted'); renderBilling(BILL.job.id);
}
async function ratesModal(jobId,ref){
  const list=await API.rates_for_job(jobId);
  let b=`<div class="muted" style="margin-bottom:4px">£/hour for each person on ${esc(ref)}. Leave blank to use the person's standard rate.</div>`;
  list.forEach(u=>{const hint=u.default?`standard ${money(u.default)}`:'no standard rate set';
    b+=`<label>${esc(u.name)} <span class="muted">(${esc(u.role)})</span></label><input id="rate_${u.id}" type="number" step="0.01" min="0" value="${u.rate||''}" placeholder="${hint}">`;});
  openModal('Set rates',b,()=>saveRates(jobId,list));
}
async function saveRates(jobId,list){
  const rates={};list.forEach(u=>rates[u.id]=parseFloat(E('rate_'+u.id).value)||0);
  await API.save_rates(jobId,rates);closeModal();renderJobCosts();
}

// ===== Reports =====
async function renderReports(){
  const o=await opts();
  const opt=(arr,first)=>`<option value="">${first}</option>`+arr.map(a=>`<option value="${a.v!==undefined?a.v:esc(a)}">${esc(a.t!==undefined?a.t:a)}</option>`).join('');
  const today=new Date().toISOString().slice(0,10);
  const ago=new Date(Date.now()-30*864e5).toISOString().slice(0,10);
  E('main').innerHTML=head('Reports')+`<div class="card"><div class="rep-grid">
    <label>Report type</label><select id="r_type"><option>Person</option><option>Job Reference</option><option>Job Role</option><option>Missing Timesheets</option></select>
    <label>Person</label><select id="r_user">${opt(o.users.map(u=>({v:u.id,t:u.name})),'(All)')}</select>
    <label>Job role</label><select id="r_role">${opt(o.roles,'(All)')}</select>
    <label>Job reference</label><select id="r_job">${opt(o.jobs.map(j=>({v:j.id,t:j.label})),'(All)')}</select>
    <label>Status</label><select id="r_status">${opt(o.statuses,'(All)')}</select>
    <label>Date range</label><div style="display:flex;gap:8px;align-items:center">
      <input id="r_from" class="dateinput" value="${ago}" readonly onclick="openRange('r_from','r_to')" title="Click to pick a date range">
      <span class="muted">to</span>
      <input id="r_to" class="dateinput" value="${today}" readonly onclick="openRange('r_from','r_to')" title="Click to pick a date range"></div>
    <label>Hours basis</label><select id="r_basis"><option value="adjusted">Adjusted (for billing / profit)</option><option value="submitted">As submitted (actual hours)</option></select>
    <label>Include cost</label><input type="checkbox" id="r_cost" style="width:16px;height:16px;padding:0;justify-self:start;accent-color:#84c341">
    </div><div class="toolbar"><button class="btn sm" onclick="runReport()">Run report</button>
    <button class="btn sm ghost" onclick="expReport('csv')">Export CSV</button>
    <button class="btn sm ghost" onclick="expReport('xlsx')">Export Excel</button>
    <button class="btn sm ghost" onclick="expReport('pdf')">Export PDF</button>
    <button class="btn sm ghost" onclick="monthEnd()">Month-end Excel (all staff)</button></div></div>
    <div id="repout" style="margin-top:16px"></div>`;
}
function repF(){return{type:E('r_type').value,user_id:E('r_user').value,job_role:E('r_role').value,job_ref_id:E('r_job').value,status:E('r_status').value,date_from:E('r_from').value,date_to:E('r_to').value,basis:E('r_basis')?E('r_basis').value:'adjusted',include_cost:E('r_cost')?E('r_cost').checked:false};}
function hbar(items){
  if(!items.length)return '';
  const max=Math.max(1,...items.map(i=>i[1])); const W=620,rh=26,H=items.length*rh+10;
  let s=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}">`;
  items.forEach((it,i)=>{const y=i*rh+8; const bw=(W-210)*(it[1]/max);
    s+=`<text x="0" y="${y+13}" font-size="12.5" fill="#33373e">${esc(String(it[0]).slice(0,26))}</text>`;
    s+=`<rect x="170" y="${y+3}" width="${Math.max(2,bw)}" height="14" rx="4" fill="#84c341"></rect>`;
    s+=`<text x="${170+bw+6}" y="${y+15}" font-size="11.5" fill="#5e8e1f">${it[1]}h</text>`;});
  return s+'</svg>';
}
async function runReport(){
  const r=await API.run_report(repF());const out=E('repout');
  if(r.missing){let h='<div class="grid kpis" style="margin-bottom:16px">'+kpi('Users missing',r.total)+'</div><div class="card" style="padding:6px 8px"><table><thead><tr><th>Name</th><th>Role</th><th>Status</th></tr></thead><tbody>';
    r.rows.forEach(x=>h+=`<tr><td>${esc(x.name)}</td><td>${esc(x.role)}</td><td>${badge(x.status)}</td></tr>`);h+='</tbody></table></div>';out.innerHTML=h;return;}
  const rows=r.rows, type=E('r_type').value;
  const gk=type==='Person'?'job_ref':'name', gl=type==='Person'?'job':'person';
  const agg={}; rows.forEach(x=>{const k=x[gk]||'(none)';agg[k]=(agg[k]||0)+x.hours;});
  const items=Object.entries(agg).map(([k,v])=>[k,Math.round(v*100)/100]).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const people=new Set(rows.map(x=>x.name)).size, njobs=new Set(rows.map(x=>x.job_ref).filter(Boolean)).size;
  const _bw=(E('r_basis')&&E('r_basis').value==='submitted');
  const _bl=_bw?'As submitted — actual hours logged by staff':'Adjusted — manager-edited hours (billing / profit)';
  const withCost=!!r.cost_included;
  let h='<div class="sub" style="margin-bottom:10px;display:flex;align-items:center;gap:8px"><span class="pill '+(_bw?'pill-sub':'pill-adj')+'">'+(_bw?'ACTUAL':'ADJUSTED')+'</span><span>'+_bl+'</span></div>'+
        '<div class="grid kpis" style="margin-bottom:16px">'+kpi('Total hours',r.total,1)+(withCost?kpi('Total cost','£'+(r.total_cost||0).toFixed(2)):'')+kpi('Entries',rows.length)+kpi('People',people)+kpi('Jobs',njobs)+'</div>';
  if(items.length)h+=`<div class="card" style="margin-bottom:16px"><div class="sec-t">Hours by ${gl}</div>${hbar(items)}</div>`;
  h+='<div class="card" style="padding:6px 8px"><table><thead><tr><th>Name</th><th>Role</th><th>Date</th><th>Job Ref</th><th>Job Title</th><th>Narrative</th><th class="tnum">Hours</th>'+(withCost?'<th class="tnum">Cost</th>':'')+'<th>Status</th></tr></thead><tbody>';
  if(!rows.length)h+=`<tr><td colspan="${withCost?9:8}" class="muted">No data.</td></tr>`;
  rows.forEach(x=>h+=`<tr><td>${esc(x.name)}</td><td>${esc(x.role)}</td><td>${esc(x.date)}</td><td>${esc(x.job_ref)}</td><td>${esc(x.title)}</td><td>${esc(x.narrative)}</td><td class="tnum">${x.hours}</td>${withCost?`<td class="tnum">£${(x.cost||0).toFixed(2)}</td>`:''}<td>${badge(x.status)}</td></tr>`);
  h+='</tbody></table></div>';out.innerHTML=h;
}
async function expReport(kind){
  const r=await API.export_report(kind,repF());
  if(!r)return;
  if(r.error){alert(r.error);return;}
  if(!r.ok)return;                       // save dialog cancelled - say nothing
  // Open it straight away (Excel / PDF reader), rather than just telling the
  // user a path they then have to go and find.
  let opened=false;
  try{opened=await API.open_file(r.path);}catch(e){}
  if(opened)toast((kind==='xlsx'?'Excel file':kind.toUpperCase())+' saved — opening…');
  else alert('Saved to:\n'+r.path);      // couldn't open it: show where it went
}

// ===== Analytics =====
let _anFrom=null,_anTo=null;
function _stk(b,n){const t=b+n||1;const bw=Math.round(b/t*100);return `<div class="stk"><div class="stk-b" style="width:${bw}%"></div></div>`;}
async function renderAnalytics(){
  const today=new Date();
  const first=new Date(Date.now()-90*864e5).toISOString().slice(0,10);
  const tod=today.toISOString().slice(0,10);
  _anFrom=_anFrom||first; _anTo=_anTo||tod;
  const a=await API.analytics(_anFrom,_anTo);
  if(a&&a.error){E('main').innerHTML=head('Analytics',a.error);return;}
  const g=a.goals||{}; const isAdmin=USER.user_type==='Admin';
  const preset=(k)=>{const t=new Date(); let f; if(k==='30'){f=new Date(t-30*864e5);} else if(k==='90'){f=new Date(t-90*864e5);} else if(k==='m'){f=new Date(t.getFullYear(),t.getMonth(),1);} else if(k==='q'){f=new Date(t.getFullYear(),Math.floor(t.getMonth()/3)*3,1);} else {f=new Date(t.getFullYear(),0,1);} return f.toISOString().slice(0,10);};
  const curKey=(()=>{for(const k of ['30','90','m','q','y']) if(preset(k)===_anFrom&&_anTo===tod) return k; return '';})();
  let h=head('Analytics','How the company is doing against its goals: utilisation, money, overtime and discipline on timesheets.');
  h+=`<div class="an-top">
    <div class="seg">${[['30','30 days'],['90','90 days'],['m','This month'],['q','This quarter'],['y','This year']].map(([k,l])=>`<button class="${curKey===k?'on':''}" onclick="_anFrom='${preset(k)}';_anTo='${tod}';renderAnalytics()">${l}</button>`).join('')}</div>
    <div class="an-range"><input id="an_from" class="dateinput" value="${_anFrom}" readonly onclick="openRange('an_from','an_to')"><span class="muted">to</span><input id="an_to" class="dateinput" value="${_anTo}" readonly onclick="openRange('an_from','an_to')"><button class="btn sm ghost" onclick="_anFrom=E('an_from').value;_anTo=E('an_to').value;renderAnalytics()">Apply</button></div>
  </div>`;

  // ---- goals ----
  const months=a.period.months||1;
  const invPerMonth=a.money.invoiced/months, otPerMonth=a.overtime.total_hours/months;
  const goals=[
    {k:'utilisation_pct',lab:'Utilisation',val:a.util.company_pct,goal:g.utilisation_pct,fmt:v=>v.toFixed(1)+'%',sub:`${a.util.billable}h billable of ${(a.util.billable+a.util.nonbill).toFixed(0)}h`},
    {k:'monthly_invoiced',lab:'Invoiced per month',val:invPerMonth,goal:g.monthly_invoiced,fmt:v=>moneyS(Math.round(v)),sub:`${moneyS(Math.round(a.money.invoiced))} in period · ${moneyS(Math.round(a.money.invoiced_unpaid))} unpaid`},
    {k:'submission_pct',lab:'Timesheets on time',val:a.submission_rate,goal:g.submission_pct,fmt:v=>v.toFixed(0)+'%',sub:`${a.people} active people`},
    {k:'overtime_cap',lab:'Overtime per month',val:otPerMonth,goal:g.overtime_cap,fmt:v=>v.toFixed(0)+'h',sub:`${moneyS(Math.round(a.overtime.total_cost))} overtime cost in period`,lower:true},
  ];
  h+=`<div class="sec-h" style="display:flex;align-items:center;gap:10px">Goals${isAdmin?`<button class="btn xs ghost" style="margin-left:auto" onclick="goalsModal()">Set goals</button>`:''}</div><div class="goal-grid">`;
  goals.forEach(x=>{ const goal=Number(x.goal)||0; let p=goal?(x.lower?Math.min(1,x.val/goal):Math.min(1,x.val/goal)):0; const hit=x.lower?(x.val<=goal):(x.val>=goal);
    const tone=x.lower?(x.val>goal?'bad':(p>0.85?'warn':'ok')):(hit?'ok':(p>=0.7?'warn':'bad')); const togo=x.lower?(hit?`${x.fmt(goal-x.val)} under the cap`:`${x.fmt(x.val-goal)} over the cap`):(hit?'Goal reached':`${x.fmt(goal-x.val)} to go`);
    h+=`<div class="card goal ${tone}">${ring(x.lower?Math.min(1,goal?x.val/goal:0):p,tone)}<div class="goal-b"><div class="lab">${esc(x.lab)}</div><div class="val">${x.fmt(x.val)}</div><div class="goal-t">${x.lower?'Cap':'Goal'} ${x.fmt(goal)} · <b>${togo}</b></div><div class="goal-s">${esc(x.sub)}</div></div></div>`; });
  h+='</div>';

  // ---- headline numbers ----
  h+='<div class="grid kpis" style="margin:18px 0;grid-template-columns:repeat(auto-fit,minmax(150px,1fr))">'+
     kpi('Billable hours',a.util.billable,1)+kpi('Non-billable',a.util.nonbill)+
     kpi('Invoiced',moneyS(Math.round(a.money.invoiced)))+kpi('Supplier bills',moneyS(Math.round(a.money.bills)))+
     kpi('Overtime cost',moneyS(Math.round(a.overtime.total_cost)))+kpi('Expenses',moneyS(Math.round(a.expenses.period_total)))+'</div>';

  // ---- weekly trend ----
  h+=`<div class="card" style="margin-bottom:14px"><div class="sec-t" style="display:flex;align-items:center">Hours per week<span class="legend"><i style="background:var(--green)"></i>Billable<i style="background:#cfd4dc"></i>Non-billable<i style="background:var(--amber);height:2px;width:14px;border-radius:1px"></i>Overtime</span></div>${stackChart(a.weekly)}</div>`;

  // ---- utilisation by person + portfolio ----
  const rows=a.util.rows.slice().sort((x,y)=>y.pct-x.pct||y.billable-x.billable);
  h+='<div class="grid2"><div class="card"><div class="sec-t">Utilisation by person</div><div class="hbars">';
  if(!rows.length)h+='<div class="muted">No hours in this period.</div>';
  rows.forEach((r,i)=>{const t=r.billable+r.nonbill; h+=`<div class="hrow${i>=10?' more hidden':''}"><span class="hname">${person(r.name,22)}</span><span class="hb"><i style="width:${r.pct}%;background:${r.pct>=g.utilisation_pct?'var(--green)':r.pct>=g.utilisation_pct*0.85?'var(--amber)':'var(--red)'}"></i></span><span class="hv">${r.pct}%</span><span class="hh muted">${t}h</span></div>`;});
  if(rows.length>10)h+=`<button class="btn xs ghost" style="margin-top:8px" onclick="[...this.parentNode.querySelectorAll('.more')].forEach(e=>e.classList.toggle('hidden'));this.textContent=this.textContent.startsWith('Show')?'Show fewer':'Show all ${rows.length}'">Show all ${rows.length}</button>`;
  h+=`</div><div class="sub" style="margin:10px 0 0">Billable = project jobs. Goal line ${g.utilisation_pct}%.</div></div>`;
  h+='<div class="card"><div class="sec-t">Job budgets — cost against quote</div><div class="hbars">';
  if(!a.portfolio.length)h+='<div class="muted">No quoted projects yet.</div>';
  a.portfolio.slice(0,10).forEach(d=>{const col=d.flag==='over'?'var(--red)':(d.flag==='warn'?'var(--amber)':'var(--green)');
    h+=`<div class="hrow" title="${esc(d.job_title)}"><span class="hname"><b>${esc(d.job_ref)}</b><span class="muted" style="margin-left:6px;font-size:11.5px">${esc(d.job_title.slice(0,26))}${d.job_title.length>26?'…':''}</span></span><span class="hb"><i style="width:${Math.min(100,d.pct)}%;background:${col}"></i></span><span class="hv" style="color:${col}">${d.pct}%</span><span class="hh muted">${moneyS(Math.round(d.remaining))} left</span></div>`;});
  h+='</div><div class="sub" style="margin:10px 0 0">Green under 80% · amber 80–100% · red over the quote.</div></div></div>';

  // ---- top jobs + expenses ----
  h+='<div class="grid2" style="margin-top:14px"><div class="card"><div class="sec-t">Biggest jobs by '+(a.top_metric==='cost'?'labour cost':'hours')+'</div>';
  const maxv=Math.max(1,...a.top_jobs.map(j=>a.top_metric==='cost'?j.cost:j.hours));
  a.top_jobs.forEach((j,i)=>{const v=a.top_metric==='cost'?j.cost:j.hours; h+=`<div class="rankrow"><div class="rankno">${i+1}</div><div style="flex:1;min-width:0"><div style="display:flex;justify-content:space-between;gap:10px"><b>${esc(j.job_ref)}</b><span class="tnum">${a.top_metric==='cost'?moneyS(Math.round(v)):v+'h'}</span></div><div class="muted" style="font-size:11.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(j.job_title)}</div><div class="rankbar"><div class="rankbar-f" style="width:${Math.round(v/maxv*100)}%"></div></div></div></div>`;});
  if(!a.top_jobs.length)h+='<div class="muted">Nothing logged yet.</div>';
  h+='</div><div class="card"><div class="sec-t">Expenses by job</div>';
  const emax=Math.max(1,...a.expenses.by_job.map(r=>r.value));
  if(!a.expenses.by_job.length)h+='<div class="muted">No expenses in this period.</div>';
  h+='<div class="hbars">'+a.expenses.by_job.slice(0,10).map(r=>`<div class="hrow"><span class="hname">${esc(r.job)}</span><span class="hb"><i style="width:${Math.round(r.value/emax*100)}%;background:var(--blue)"></i></span><span class="hv">${money(r.value)}</span></div>`).join('')+'</div>';
  h+=`<div class="sub" style="margin:10px 0 0">${money(a.expenses.period_total)} in period · ${money(a.expenses.outstanding)} still to reimburse.</div></div></div>`;

  // ---- leakage (detail) ----
  h+='<details class="card an-details" style="margin-top:14px"><summary><span class="sec-t" style="margin:0">Profit leakage — submitted vs adjusted hours</span><span class="muted" style="font-size:12px">Hours changed by managers after submission. Positive = added, negative = cut.</span></summary><div class="grid2" style="margin-top:14px">';
  h+='<div><table><thead><tr><th>Person</th><th class="tnum">Submitted</th><th class="tnum">Adjusted</th><th class="tnum">Diff</th></tr></thead><tbody>';
  if(!a.leak_person.length)h+='<tr><td colspan="4" class="muted">No data.</td></tr>';
  a.leak_person.forEach(r=>h+=`<tr><td>${person(r.name,22)}</td><td class="tnum">${r.submitted}</td><td class="tnum">${r.adjusted}</td><td class="tnum">${diffc(r.diff)}</td></tr>`);
  h+='</tbody></table></div><div><table><thead><tr><th>Job</th><th class="tnum">Sub.</th><th class="tnum">Adj.</th><th class="tnum">Diff</th></tr></thead><tbody>';
  if(!a.leak_job.length)h+='<tr><td colspan="4" class="muted">No data.</td></tr>';
  a.leak_job.slice(0,12).forEach(r=>h+=`<tr><td title="${esc(r.job_title)}">${esc(r.job_ref)}</td><td class="tnum">${r.submitted}</td><td class="tnum">${r.adjusted}</td><td class="tnum">${diffc(r.diff)}</td></tr>`);
  h+='</tbody></table></div></div></details>';
  // ---- overtime (detail) ----
  h+='<details class="card an-details" style="margin-top:14px"><summary><span class="sec-t" style="margin:0">Overtime by person</span><span class="muted" style="font-size:12px">'+a.overtime.total_hours+'h · '+moneyS(Math.round(a.overtime.total_cost))+' premium</span></summary>';
  h+='<table style="margin-top:12px"><thead><tr><th>Person</th><th class="tnum">OT hours</th><th class="tnum">Premium cost</th></tr></thead><tbody>';
  if(!a.overtime.rows.length)h+='<tr><td colspan="3" class="muted">No overtime in this period.</td></tr>';
  a.overtime.rows.forEach(r=>h+=`<tr><td>${person(r.name,22)}</td><td class="tnum">${r.hours}</td><td class="tnum">${money(r.cost)}</td></tr>`);
  h+='</tbody></table></details>';
  E('main').innerHTML=h;
}
function ring(p,tone){ const r=26,c=2*Math.PI*r; p=Math.max(0,Math.min(1,p||0)); const col=tone==='ok'?'var(--green)':tone==='warn'?'var(--amber)':'var(--red)';
  return `<svg class="ring" viewBox="0 0 64 64" width="64" height="64"><circle cx="32" cy="32" r="${r}" fill="none" stroke="var(--line)" stroke-width="7"/><circle cx="32" cy="32" r="${r}" fill="none" stroke="${col}" stroke-width="7" stroke-linecap="round" stroke-dasharray="${c}" stroke-dashoffset="${c*(1-p)}" transform="rotate(-90 32 32)" style="--c:${c};--o:${c*(1-p)}"/><text x="32" y="36" text-anchor="middle" font-size="12" font-weight="800" fill="var(--ink)">${Math.round(p*100)}%</text></svg>`; }
function stackChart(weeks){
  if(!weeks||!weeks.length) return '<div class="chart-empty">No hours in this period.</div>';
  const W=1080,H=230,padL=36,padR=10,padT=16,padB=28,iw=W-padL-padR,ih=H-padT-padB,bw=iw/weeks.length;
  const tops=weeks.map(w=>w.billable+w.nonbill); const max=Math.max(10,Math.ceil(Math.max(...tops)/50)*50);
  const y=v=>padT+ih-ih*v/max; const step=Math.max(1,Math.ceil(weeks.length/10));
  let s=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}" class="bars" role="img" aria-label="Hours per week">`;
  [0,max/2,max].forEach(gv=>{ s+=`<line x1="${padL}" x2="${W-padR}" y1="${y(gv)}" y2="${y(gv)}" stroke="var(--line)" stroke-dasharray="${gv?'3 4':'0'}"/><text x="${padL-6}" y="${y(gv)+3.5}" font-size="10" fill="var(--mut-2)" text-anchor="end">${gv}</text>`; });
  const pts=[];
  weeks.forEach((w,i)=>{ const x=padL+i*bw+bw*0.18, bwid=bw*0.64; const yb=y(w.billable), yt=y(w.billable+w.nonbill);
    s+=`<g class="bar" style="animation-delay:${i*30}ms"><rect x="${x}" y="${yb}" width="${bwid}" height="${Math.max(0,padT+ih-yb)}" rx="3" fill="var(--green)"/><rect x="${x}" y="${yt}" width="${bwid}" height="${Math.max(0,yb-yt)}" rx="3" fill="#cfd4dc"/><rect x="${padL+i*bw}" y="${padT}" width="${bw}" height="${ih}" fill="transparent"><title>w/c ${w.label}: ${w.billable}h billable, ${w.nonbill}h non-billable, ${w.overtime}h overtime</title></rect>`;
    if(i%step===0) s+=`<text x="${x+bwid/2}" y="${H-padB+15}" font-size="10" fill="var(--mut)" text-anchor="middle">${esc(w.label)}</text>`;
    s+='</g>'; pts.push([x+bwid/2,y(w.overtime)]); });
  s+=`<polyline class="trend" points="${pts.map(p=>p.join(',')).join(' ')}" fill="none" stroke="var(--amber)" stroke-width="2" stroke-linejoin="round"/>`;
  pts.forEach(p=>{ s+=`<circle cx="${p[0]}" cy="${p[1]}" r="2.5" fill="#fff" stroke="var(--amber)" stroke-width="1.5"/>`; });
  return s+'</svg>';
}
async function goalsModal(){
  const g=await API.get_goals();
  openModal('Company goals',`<label>Utilisation goal (% of hours billable)</label><input id="g_u" type="number" step="1" min="0" max="100" value="${g.utilisation_pct}">
    <label>Invoiced per month (£)</label><input id="g_i" type="number" step="1000" min="0" value="${g.monthly_invoiced}">
    <label>Timesheets submitted on time (%)</label><input id="g_s" type="number" step="1" min="0" max="100" value="${g.submission_pct}">
    <label>Overtime cap per month (hours)</label><input id="g_o" type="number" step="5" min="0" value="${g.overtime_cap}">
    <div class="muted" style="font-size:12px;margin-top:10px">Shown to managers on the Analytics page. Change them any time.</div>`,
    async()=>{ const r=await API.save_goals({utilisation_pct:E('g_u').value,monthly_invoiced:E('g_i').value,submission_pct:E('g_s').value,overtime_cap:E('g_o').value}); if(!r||!r.ok){alert((r&&r.error)||'Could not save');return;} closeModal(); toast('Goals saved'); renderAnalytics(); });
}
function diffc(v){const c=v>0?'#1d8f3b':(v<0?'#c0392b':'#888');const s=v>0?'+':'';return `<span style="color:${c};font-weight:600">${s}${v}</span>`;}

// ===== Change password =====
async function renderSettings(){
  const s=await API.get_settings();
  const m=E('main');
  m.innerHTML=head('Settings','Company details, targets and where the data is kept. Admins only.')+
  `<div class="card"><div class="set-grid">
    <div class="k">Company name<small>Shown on the sign-in screen, sidebar and every export.</small></div><input type="text" id="st_co" value="${esc(s.company_name)}">
    <div class="k">Logo<small>PNG, under 700 KB. Optional.</small></div><div style="display:flex;gap:8px;align-items:center"><button class="btn sm ghost" onclick="pickLogo()">${s.has_logo?'Change logo…':'Add logo…'}</button>${s.has_logo?'<button class="btn sm ghost" onclick="removeLogo()">Remove</button>':''}</div>
    <div class="k">Weekly target hours<small>Drives the "hours logged" progress on the dashboard.</small></div><input type="number" id="st_wk" step="0.5" min="1" value="${s.weekly_target_hours}">
    <div class="k">Daily hours warning<small>Days over this get flagged on timesheets.</small></div><input type="number" id="st_day" step="0.5" min="1" value="${s.daily_hours_warning}">
    <div class="k">Sign out after (minutes idle)</div><input type="number" id="st_idle" step="1" min="1" value="${s.inactivity_timeout_minutes}">
    <div class="k">Open maximised</div><label style="font-size:13px"><input type="checkbox" id="st_max" ${s.start_maximized?'checked':''}> Fill the screen when the app opens</label>
  </div><div style="margin-top:16px"><button class="btn sm" onclick="saveSettings()">Save settings</button></div></div>
  <div class="card" style="margin-top:16px"><div style="font-weight:700;margin-bottom:6px">Data</div>
    <div class="muted" style="font-size:13px;line-height:1.5">Database, receipts and daily backups are in:<br><code style="font-size:12.5px">${esc(s.data_folder)}</code></div>
    <div style="display:flex;gap:8px;margin-top:12px;flex-wrap:wrap"><button class="btn sm ghost" onclick="API.open_data_folder()">Open folder</button><button class="btn sm ghost" onclick="moveData()">Move to a shared folder…</button></div>
    <div class="muted" style="font-size:12px;margin-top:10px;line-height:1.5">To let a team use one database, move it to a folder everyone can reach (a network drive). Other PCs then pick the same folder on their first run and join it. If the folder you pick already has a Timesheets database, this PC joins that one instead and you'll be asked to sign in again.</div>
    <div class="muted" style="font-size:11.5px;margin-top:12px">${esc(window._brand?window._brand.app:'Timesheets')} ${esc(s.version)}</div></div>`;
}
async function saveSettings(){
  const r=await API.save_settings({company_name:E('st_co').value,weekly_target_hours:E('st_wk').value,daily_hours_warning:E('st_day').value,inactivity_timeout_minutes:E('st_idle').value,start_maximized:E('st_max').checked});
  if(!r||!r.ok){alert((r&&r.error)||'Could not save.');return;}
  await applyBrand(); try{const c=await API.app_config(); window._dailyWarn=c.daily_warning; _idleMs=(c.timeout_minutes||20)*60000;}catch(e){}
  toast('Settings saved');
}
async function pickLogo(){ const r=await API.pick_logo(); if(r&&r.error)alert(r.error); if(r&&r.ok){await applyBrand(); renderSettings(); toast('Logo updated');} }
async function removeLogo(){ if(!confirm('Remove the logo?'))return; await API.remove_logo(); await applyBrand(); renderSettings(); }
async function moveData(){
  const p=await API.pick_data_folder(); if(!p||!p.ok)return;
  const msg=p.has_db?'That folder already has a Timesheets database.\n\nThis PC will switch to it (your current data stays where it is, untouched) and you will need to sign in again. Continue?':'Copy the database and receipts to:\n'+p.folder+'\n\nand use that from now on? Your current copy stays where it is as a backup.';
  if(!confirm(msg))return;
  const r=await API.change_data_folder(p.folder);
  if(!r||!r.ok){alert((r&&r.error)||'Could not move the data.');return;}
  if(r.needs_login){ await API.logout(); USER=null; E('app').classList.add('hidden'); E('login').classList.remove('hidden'); E('lp').value=''; await applyBrand(); E('lerr').textContent='Now using the shared database. Sign in with an account that exists there.'; return; }
  toast('Data folder changed'); renderSettings();
}
function changePw(){E('main').innerHTML=head('Change Password')+'<div class="card" style="max-width:420px"><label class="muted" style="font-size:13px">New password</label><input id="cp_pw" type="password" style="width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:9px;margin-top:6px"><div style="margin-top:14px"><button class="btn sm" onclick="doChangePw()">Update password</button></div></div>';}
async function doChangePw(){const p=E('cp_pw').value;if(!p){alert('Enter a new password.');return;}await API.change_my_password(p);alert('Your password has been changed.');E('cp_pw').value='';}


let _apStatus='Submitted', _apWho='', _apRows=[];
async function renderApprovals(){
  _apRows=(await API.approvals(_apStatus)).filter(r=>r.status!=='Draft');
  drawApprovals();
}
function drawApprovals(){
  const mgr=['Manager','Admin'].includes(USER.user_type);
  const sts=['Submitted','Approved','Rejected','Locked','All'];
  // Keep whoever is currently selected in the list even if they have nothing
  // under this status, so changing status doesn't silently clear the filter.
  const people=Array.from(new Set(_apRows.map(r=>r.name).concat(_apWho?[_apWho]:[])))
                    .sort((a,b)=>a.localeCompare(b));
  let rows=_apRows.slice();
  if(_apWho) rows=rows.filter(r=>r.name===_apWho);
  let h=head('Approvals','Timesheets people have submitted for your decision. Drafts must be submitted first.');
  h+=`<div class="toolbar" style="margin:0 0 14px"><label class="muted" style="align-self:center">Person:</label>
      <select id="ap_who" onchange="_apWho=this.value;drawApprovals()">
        <option value="">(Everyone)</option>
        ${people.map(p=>`<option ${p===_apWho?'selected':''}>${esc(p)}</option>`).join('')}</select>
      <label class="muted" style="align-self:center;margin-left:12px">Show:</label>
      <select id="ap_st" onchange="_apStatus=this.value;renderApprovals()">${sts.map(x=>`<option ${x===_apStatus?'selected':''}>${x}</option>`).join('')}</select>
      <span class="muted" style="align-self:center;margin-left:auto">${rows.length} timesheet(s)${_apWho?' · '+esc(_apWho):''}</span></div>`;
  h+='<div class="card" style="padding:6px 8px"><table><thead><tr><th>Name</th><th>Role</th><th>Week</th><th>Status</th><th class="tnum">Hours</th><th>Reason</th><th>Action</th></tr></thead><tbody>';
  if(!rows.length)h+=`<tr><td colspan="7" class="muted">${_apWho?'Nothing here for '+esc(_apWho)+'.':'Nothing here.'}</td></tr>`;
  rows.forEach(r=>{const s=r.status;
    let acts=`<button class="btn sm ghost" onclick="viewDetail(${r.id})">View</button>`;
    if(mgr && (s==='Submitted'||s==='Approved')) acts+=` <button class="btn sm ghost" onclick="adminEdit(${r.id})">Edit</button>`;
    if(s==='Submitted'){ acts+=` <button class="btn sm" onclick="doApprove(${r.id},'Approved')">Approve</button>`;
                         acts+=` <button class="btn sm ghost" onclick="doReject(${r.id})">Reject</button>`; }
    if(s==='Approved'){ acts+=` <button class="btn sm ghost" onclick="doReject(${r.id})">Send back</button>`;
                        acts+=` <button class="btn sm ghost" onclick="doApprove(${r.id},'Locked')">Lock</button>`; }
    if(s==='Locked') acts+=` <button class="btn sm ghost" onclick="doApprove(${r.id},'Approved')">Unlock</button>`;
    h+=`<tr><td>${person(r.name)}</td><td class="muted">${esc(r.role)}</td><td>${esc(r.week)}</td><td>${badge(r.status)}</td><td class="tnum">${r.hours}</td><td class="muted">${esc(r.reason||'')}</td><td style="white-space:nowrap">${acts}</td></tr>`;});
  h+='</tbody></table></div>';
  E('main').innerHTML=h;
}
function renderApprovals2(){renderApprovals();}
async function viewDetail(tid,basis){
  basis=basis||'adjusted';
  const d=await API.approval_lines(tid,basis);
  let toggle='';
  if(d.has_snapshot){
    toggle=`<div style="margin-bottom:10px;display:flex;gap:6px;align-items:center">
      <button class="btn sm ${d.basis==='adjusted'?'':'ghost'}" onclick="viewDetail(${tid},'adjusted')">Adjusted (billing)</button>
      <button class="btn sm ${d.basis==='submitted'?'':'ghost'}" onclick="viewDetail(${tid},'submitted')">Employee submission</button>
      <span class="pill ${d.basis==='submitted'?'pill-sub':'pill-adj'}" style="margin-left:auto">${d.basis==='submitted'?'AS SUBMITTED':'ADJUSTED'}</span></div>`;
  }
  let b=toggle+'<div style="max-height:55vh;overflow:auto"><table style="width:100%;border-collapse:collapse">';
  d.days.forEach((day,i)=>{const dl=d.lines.filter(l=>l.work_date===day);if(!dl.length)return;
    b+=`<tr><td colspan="3" style="background:#eef1f5;padding:6px 8px"><b>${esc(d.day_labels[i])}</b></td></tr>`;
    dl.forEach(l=>b+=`<tr><td style="padding:5px 8px">${esc(l.job_ref)}</td><td style="padding:5px 8px">${esc(l.narrative)}</td><td style="padding:5px 8px;text-align:right">${l.hours}h${l.rate==='Normal'?'':' '+esc(l.rate)}</td></tr>`);});
  const tot=d.lines.reduce((a,b)=>a+b.hours,0);
  b+=`</table></div><div style="margin-top:10px"><b>Total: ${tot} h</b></div>`;
  openModal('Timesheet detail – '+esc(d.name),b,()=>closeModal());
  E('modalOk').textContent='Close'; { const mc=E('modalCancel'); if(mc)mc.style.display='none'; }
}

async function renderTimeline(){
  const n=window._tlWeeks||12;
  const d=await API.job_timeline(n);
  if(d&&d.error){E('main').innerHTML=head('Timeline',d.error);return;}
  let h=head('Timeline','Hours logged per job across recent weeks. Darker green = more hours that week.');
  h+=`<div class="toolbar" style="margin:0 0 12px"><label class="muted" style="align-self:center">Weeks:</label>
     <select onchange="window._tlWeeks=Number(this.value);renderTimeline()">${[8,12,26].map(x=>`<option ${x===n?'selected':''}>${x}</option>`).join('')}</select>
     <span class="muted" style="align-self:center;margin-left:10px">${d.jobs.length} job(s) &middot; ${d.grand} h total</span></div>`;
  if(!d.jobs.length){h+='<div class="card"><span class="muted">No hours logged in this period.</span></div>';E('main').innerHTML=h;return;}
  h+='<div class="card" style="overflow-x:auto;padding:8px 10px"><table class="tl"><thead><tr><th class="tl-job">Job</th>'+d.weeks.map(w=>`<th>${w}</th>`).join('')+'<th>Total</th></tr></thead><tbody>';
  d.jobs.forEach(j=>{
    h+=`<tr class="tl-row" onclick="renderTimelineDetail(${j.id})" title="Click to see who worked on ${esc(j.job_ref)}"><td class="tl-job"><b>${esc(j.job_ref)}</b><div class="tl-sub">${esc(j.job_title)}</div></td>`;
    j.cells.forEach(c=>{const op=c>0?(0.14+0.86*Math.min(1,c/d.maxcell)):0;const col=op>0.55?'#fff':'#33373e';
      h+=`<td class="tlc" style="background:rgba(132,195,65,${op.toFixed(2)});color:${col}">${c>0?c:''}</td>`;});
    h+=`<td class="tlc tl-tot">${j.total}</td></tr>`;
  });
  h+='</tbody></table></div>';
  E('main').innerHTML=h;
}
async function renderTimelineDetail(jobId){
  const n=window._tlWeeks||12;
  const d=await API.job_timeline_detail(jobId,n);
  if(d&&d.error){E('main').innerHTML=head('Timeline',d.error);return;}
  let h=head('Timeline \u2013 '+esc(d.job_ref), esc(d.job_title)+' \u00b7 hours each person logged to this job, by week.');
  h+=`<div class="toolbar" style="margin:0 0 12px"><button class="btn sm ghost" onclick="renderTimeline()">\u2039 Back to all jobs</button>
     <span class="muted" style="align-self:center;margin-left:8px">${d.people.length} person(s) &middot; ${d.grand} h total</span></div>`;
  if(!d.people.length){h+='<div class="card"><span class="muted">No hours logged to this job in this period.</span></div>';E('main').innerHTML=h;return;}
  h+='<div class="card" style="overflow-x:auto;padding:8px 10px"><table class="tl"><thead><tr><th class="tl-job">Person</th>'+d.weeks.map(w=>`<th>${w}</th>`).join('')+'<th>Total</th></tr></thead><tbody>';
  d.people.forEach(p=>{
    h+=`<tr><td class="tl-job"><b>${esc(p.name)}</b></td>`;
    p.cells.forEach((c,ci)=>{const op=c>0?(0.14+0.86*Math.min(1,c/d.maxcell)):0;const col=op>0.55?'#fff':'#33373e';
      const click=c>0?`onclick="tlLines(${jobId},${p.uid},'${d.week_iso[ci]}')" title="View what ${esc(p.name)} did that week"`:'';
      h+=`<td class="${c>0?'tlc tlc-click':'tlc'}" ${click} style="background:rgba(132,195,65,${op.toFixed(2)});color:${col}">${c>0?c:''}</td>`;});
    h+=`<td class="tlc tl-tot">${p.total}</td></tr>`;
  });
  h+='</tbody></table></div>';
  E('main').innerHTML=h;
}
async function tlLines(jobId,uid,weekISO){
  const d=await API.timeline_lines(jobId,uid,weekISO);
  if(d&&d.error){alert(d.error);return;}
  let b=`<div class="sub" style="margin-bottom:10px">${esc(d.name)} &middot; ${esc(d.job_ref)} &middot; ${esc(d.week)}</div>`;
  if(!d.lines.length){b+='<div class="muted">No detail found.</div>';}
  else{b+='<div style="max-height:55vh;overflow:auto"><table style="width:100%;border-collapse:collapse">';
    d.lines.forEach(l=>b+=`<tr style="border-bottom:1px solid var(--line)"><td style="padding:6px 8px;white-space:nowrap;vertical-align:top"><b>${esc(l.date)}</b></td><td style="padding:6px 8px">${esc(l.narrative)}${l.role?` <span class="muted">(${esc(l.role)})</span>`:''}</td><td style="padding:6px 8px;text-align:right;white-space:nowrap;vertical-align:top">${l.hours}h${l.rate==='Normal'?'':' '+esc(l.rate)}</td></tr>`);
    b+=`</table></div><div style="margin-top:10px"><b>Total: ${d.total} h</b></div>`;}
  openModal('What they were doing',b,()=>closeModal());
  E('modalOk').textContent='Close'; { const mc=E('modalCancel'); if(mc)mc.style.display='none'; }
}
async function renderImport(){
  const o=await opts();
  const _t=new Date(); _t.setHours(12,0,0,0); _t.setDate(_t.getDate()-((_t.getDay()+6)%7));
  const monISO=`${_t.getFullYear()}-${String(_t.getMonth()+1).padStart(2,'0')}-${String(_t.getDate()).padStart(2,'0')}`;
  let h=head('Import (external timesheets)','For people who do not use the app. Generate a sheet for a specific week (dates pre-filled), email it out, then import the file they send back \u2014 it goes into Approvals like any other.');
  h+=`<div class="card" style="max-width:680px">
    <div class="sec-t">Step 1 &middot; Add the person under Users</div>
    <p class="sub">Add the contractor under <b>Users</b> first (matched by their email). Only people already set up can be imported.</p>
    <div class="sec-t" style="margin-top:16px">Step 2 &middot; Generate a timesheet for a week</div>
    <p class="sub">The 7 day rows are pre-filled with that week's dates. Optionally pick the person to pre-fill their name &amp; email.</p>
    <div class="rep-grid" style="max-width:520px">
      <label>Person (optional)</label><select id="imp_user"><option value="">\u2014 blank \u2014</option>${o.users.map(u=>`<option value="${u.id}">${esc(u.name)}</option>`).join('')}</select>
      <label>Week commencing</label><input id="imp_week" type="date" value="${monISO}">
    </div>
    <div style="margin-top:10px"><button class="btn sm" onclick="dlTemplate()">Generate timesheet for this week</button></div>
    <div class="sec-t" style="margin-top:18px">Step 3 &middot; Import their filled-in file</div>
    <p class="sub">When they email it back, import it here. It appears in <b>Approvals</b> as Submitted.</p>
    <button class="btn sm" onclick="doImport()">Import filled timesheet\u2026</button>
    <div id="impout" style="margin-top:16px"></div>
  </div>`;
  E('main').innerHTML=h;
}
async function dlTemplate(){
  const wk=E('imp_week').value||null, uid=E('imp_user').value||null;
  const r=await API.export_timesheet_template(wk,uid);
  if(r&&r.ok){await API.open_file(r.path);}else if(r&&r.error)alert(r.error);
}
async function doImport(){
  const o=E('impout'); o.innerHTML='<span class="muted">Reading file\u2026</span>';
  const r=await API.import_external_timesheet();
  if(!r){o.innerHTML='';return;}
  const li=(a)=>(a&&a.length)?'<ul style="margin:6px 0 0 0">'+a.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul>':'';
  if(!r.ok){ o.innerHTML=`<div class="card" style="border:1px solid #e6bcbc;background:#fdf3f3"><b style="color:#b22">Import failed</b><div>${esc(r.error||'')}</div>${li(r.details)}</div>`; return; }
  const warn=(r.errors&&r.errors.length)?`<div style="margin-top:8px;color:#9a6a00">Skipped ${r.errors.length} row(s):${li(r.errors)}</div>`:'';
  o.innerHTML=`<div class="card" style="border:1px solid #bcd89a;background:#f6faef"><b style="color:#4f7d12">Imported \u2713</b><div>${esc(r.name)} &middot; ${esc(r.week)} &middot; ${r.count} lines &middot; ${r.total} h. Now in Approvals as Submitted.</div>${warn}</div>`;
  toast('Timesheet imported');
}

let _atStatus='All', _atBasis='adjusted';
let _atRows=[], _atSort='week', _atDesc=true, _atWho='';
function atSort(col){
  if(_atSort===col){_atDesc=!_atDesc;} else {_atSort=col;_atDesc=(col==='week'||col==='hours');}
  drawAllTs();
}
async function renderAllTs(){
  _atRows=await API.approvals(_atStatus,_atBasis);
  drawAllTs();
}
function drawAllTs(){
  const sts=['All','Submitted','Approved','Rejected','Locked','Draft'];
  // people currently in the list, so the filter only offers real options
  const people=Array.from(new Set(_atRows.map(r=>r.name))).sort((a,b)=>a.localeCompare(b));
  let rows=_atRows.slice();
  if(_atWho) rows=rows.filter(r=>r.name===_atWho);
  const key=r=>{
    if(_atSort==='name')return (r.name||'').toLowerCase();
    if(_atSort==='role')return (r.role||'').toLowerCase();
    if(_atSort==='status')return (r.status||'').toLowerCase();
    if(_atSort==='hours')return Number(r.hours)||0;
    return r.week_iso||r.week||'';        // week: sort on the real date
  };
  rows.sort((a,b)=>{const x=key(a),y=key(b);
    if(x<y)return _atDesc?1:-1; if(x>y)return _atDesc?-1:1;
    return (a.name||'').localeCompare(b.name||'');});   // stable tie-break by name
  let h=head('All Timesheets','Browse every timesheet. Click any column heading to sort by it, or pick a person to see just theirs.');
  h+=`<div class="toolbar" style="margin:0 0 14px;flex-wrap:wrap">
    <label class="muted" style="align-self:center">Person:</label>
    <select id="at_who" onchange="_atWho=this.value;drawAllTs()">
      <option value="">(Everyone)</option>
      ${people.map(p=>`<option ${p===_atWho?'selected':''}>${esc(p)}</option>`).join('')}</select>
    <label class="muted" style="align-self:center;margin-left:12px">Status:</label>
    <select id="at_st" onchange="_atStatus=this.value;renderAllTs()">${sts.map(x=>`<option ${x===_atStatus?'selected':''}>${x}</option>`).join('')}</select>
    <label class="muted" style="align-self:center;margin-left:12px">Hours basis:</label>
    <select id="at_basis" onchange="_atBasis=this.value;renderAllTs()">
      <option value="adjusted" ${_atBasis==='adjusted'?'selected':''}>Adjusted (billing / profit)</option>
      <option value="submitted" ${_atBasis==='submitted'?'selected':''}>Employee submission (actual)</option>
    </select>
    <span class="pill ${_atBasis==='submitted'?'pill-sub':'pill-adj'}" style="align-self:center;margin-left:6px">${_atBasis==='submitted'?'AS SUBMITTED':'ADJUSTED'}</span>
    <span class="muted" style="align-self:center;margin-left:auto">${rows.length} timesheet(s)${_atWho?' · '+esc(_atWho):''}</span></div>`;
  const th=(col,label,cls)=>`<th class="sortable ${cls||''}" onclick="atSort('${col}')" title="Sort by ${label.toLowerCase()}">${label}${_atSort===col?`<span class="sortarrow">${_atDesc?'▼':'▲'}</span>`:''}</th>`;
  h+='<div class="card" style="padding:6px 8px"><table><thead><tr>'
    +th('name','Name')+th('role','Role')+th('week','Week')+th('status','Status')
    +th('hours','Hours','tnum')+'<th>Action</th></tr></thead><tbody>';
  if(!rows.length)h+='<tr><td colspan="6" class="muted">No timesheets found.</td></tr>';
  rows.forEach(r=>h+=`<tr><td>${person(r.name)}</td><td class="muted">${esc(r.role)}</td><td>${esc(r.week)}</td><td>${badge(r.status)}</td><td class="tnum">${r.hours}</td><td><button class="btn sm ghost" onclick="viewDetail(${r.id},'${_atBasis}')">View</button></td></tr>`);
  h+='</tbody></table></div>';
  E('main').innerHTML=h;
}
// ---- generic line modal (used by admin edit) ----
function lineModalG(ctx,line,onOk){
  const days=ctx.days.map((d,i)=>`<option value="${d}">${esc(ctx.day_labels[i])}</option>`).join('');
  const jobOpts=(line?'':'<option value="" disabled selected>— select a job —</option>')+window._tjobs.map(j=>`<option value="${j.id}">${esc(j.label)}</option>`).join('');
  const roleOpts=window._troles.map(r=>`<option value="${esc(r)}">${esc(r)}</option>`).join('');
  const rateOpts=['Normal','x1.5','x2'].map(r=>`<option value="${r}">${r}</option>`).join('');
  openModal('Timesheet line',
    `<label>Day</label><select id="m_day">${days}</select><label>Job reference</label><select id="m_job">${jobOpts}</select><label>Role</label><select id="m_role">${roleOpts}</select><label>Overtime rate</label><select id="m_rate">${rateOpts}</select><label>Narrative</label><textarea id="m_narr" rows="3"></textarea><label>Hours</label><input id="m_hours" type="number" step="0.25" min="0">`,
    ()=>{const hrs=parseFloat(E('m_hours').value),narr=E('m_narr').value.trim(),jobVal=E('m_job').value;
      if(!jobVal){alert('Please select a job.');return;}
      if(!narr){alert('Narrative is required.');return;} if(!(hrs>0)){alert('Hours must be greater than zero.');return;}
      closeModal();onOk({work_date:E('m_day').value,job_ref_id:Number(jobVal),job_role:E('m_role').value,rate:E('m_rate').value,narrative:narr,hours:hrs,job_ref:''});});
  if(line){E('m_day').value=line.work_date;E('m_job').value=line.job_ref_id;E('m_narr').value=line.narrative;E('m_hours').value=line.hours;E('m_rate').value=line.rate||'Normal';}
  E('m_role').value=(line&&line.job_role)||ctx.defaultRole;
}
// ---- admin edit (override) ----
let AE=null;
async function adminEdit(tid){
  AE=await API.approval_lines(tid); if(!AE){alert('Not found');return;}
  window._tjobs=await API.timesheet_jobs(); window._troles=await API.timesheet_roles();
  aeDraw();
}
function aeCtx(){return {days:AE.days,day_labels:AE.day_labels,defaultRole:AE.owner_role};}
function aeDraw(){
  const d=AE; let h=head('Edit hours – '+esc(d.name), d.week+'  ·  status '+d.status);
  h+='<div class="card" style="padding:6px 8px"><table><thead><tr><th>Day</th><th>Job Ref</th><th>Role</th><th>Narrative</th><th class="tnum">Hours</th><th>Rate</th></tr></thead><tbody>';
  d.days.forEach((day,i)=>{const dl=d.lines.filter(l=>l.work_date===day);const tot=dl.reduce((a,b)=>a+b.hours,0);
    h+=`<tr class="dayrow"><td colspan="6"><b>${esc(d.day_labels[i])}</b>${tot?` <span class="muted">(${tot}h)</span>`:''}</td></tr>`;
    dl.forEach(l=>{const idx=d.lines.indexOf(l);
      h+=`<tr class="lrow" ondblclick="aeEdit(${idx})"><td></td><td>${esc(l.job_ref)}</td><td>${esc(l.job_role)}</td><td>${esc(l.narrative)}</td><td class="tnum">${l.hours}</td><td>${l.rate==='Normal'?'':esc(l.rate)} <button class="xdel" onclick="event.stopPropagation();aeDel(${idx})">✕</button></td></tr>`;});});
  h+='</tbody></table></div><div class="ts-foot"><button class="btn sm ghost" onclick="renderApprovals2()">‹ Back to approvals</button><span style="margin-left:auto;display:flex;gap:8px"><button class="btn sm ghost" onclick="aeAdd()">Add line</button><button class="btn sm" onclick="aeSave()">Save changes</button></span></div>';
  E('main').innerHTML=h;
}
function aeAdd(){lineModalG(aeCtx(),null,(ln)=>{AE.lines.push(ln);aeDraw();});}
function aeEdit(i){lineModalG(aeCtx(),AE.lines[i],(ln)=>{AE.lines[i]=Object.assign({},AE.lines[i],ln);aeDraw();});}
function aeDel(i){AE.lines.splice(i,1);aeDraw();}
async function aeSave(){const r=await API.save_admin_lines(AE.id,AE.lines.map(strip));if(!r.ok){alert(r.error);return;}alert('Hours updated.');renderApprovals2();}


async function jobCostPdf(id){const r=await API.job_cost_pdf(id);if(r&&r.ok){await API.open_file(r.path);}else if(r&&r.error)alert(r.error);}

let _idleMs=15*60000, _idleT=null, _idleWarnT=null, _lastAct=0;
function idleClearWarn(){const w=E('idlewarn'); if(w)w.remove();}
function resetIdle(){
  if(_idleT)clearTimeout(_idleT);
  if(_idleWarnT)clearTimeout(_idleWarnT);
  idleClearWarn();
  if(!USER)return;
  // Warn a minute before, so being signed out is never a surprise.
  const warnIn=Math.max(_idleMs-60000, Math.floor(_idleMs*0.5));
  _idleWarnT=setTimeout(idleWarn,warnIn);
  _idleT=setTimeout(idleOut,_idleMs);
}
function idleWarn(){
  if(!USER||E('idlewarn'))return;
  const secs=Math.max(1,Math.round((_idleMs-Math.max(_idleMs-60000,Math.floor(_idleMs*0.5)))/1000));
  const w=document.createElement('div');
  w.id='idlewarn'; w.className='idle-warn';
  w.innerHTML=`<span><b>Still there?</b> You'll be signed out in about ${secs} seconds.</span>
    <button class="btn sm" onclick="resetIdle()">Stay signed in</button>`;
  document.body.appendChild(w);
}
async function idleOut(){ if(!USER)return;
  idleClearWarn();
  // Name the actual timeout, so a wrong value in config.json is obvious.
  alert('Signed out after '+Math.round(_idleMs/60000)+' minutes with no mouse or keyboard activity.\n\n'
        +'(An administrator can change this: "inactivity_timeout_minutes" in config.json.)');
  try{await API.logout();}catch(e){}
  USER=null; E('app').classList.add('hidden'); E('login').classList.remove('hidden'); E('lp').value='';
  // Stay-signed-in token is kept: closing/reopening the app signs you back in.
}
// Any activity resets the clock. Throttled so a mousemove storm doesn't churn
// timers, but always responsive while the warning is showing.
function onActivity(){
  const n=Date.now();
  if(E('idlewarn')||n-_lastAct>2000){_lastAct=n;resetIdle();}
}
['mousemove','keydown','click','wheel','touchstart','focus'].forEach(
  ev=>document.addEventListener(ev,onActivity,{passive:true,capture:true}));

async function renderSubmissions(){
  const rows=await API.submissions();
  const by=(st)=>rows.filter(r=>r.status===st);
  const decided=rows.filter(r=>['Approved','Rejected'].includes(r.status)).slice(0,5);
  let h=head('My Submissions','Double-click a Draft or Rejected week to open and fix it.');
  h+=subSec('Latest activity (approved & rejected)',decided,true);
  h+=subSec('Rejected – needs your attention',by('Rejected'),true,'#b21d36');
  h+=subSec('Approved',by('Approved'),false,'#3e6b13');
  h+=subSec('Awaiting approval',by('Submitted'),false);
  h+=subSec('Drafts (not yet submitted)',by('Draft'),false,'#5a616b');
  if(by('Locked').length)h+=subSec('Locked',by('Locked'),false);
  E('main').innerHTML=h;
}
function subSec(title,rows,showReason,colour){
  let h=`<div class="sec-t" style="margin-top:18px${colour?';color:'+colour:''}">${esc(title)}</div>`;
  if(!rows.length)return h+'<div class="muted" style="margin-bottom:2px">— none —</div>';
  h+='<div class="card" style="padding:6px 8px"><table><thead><tr><th>Week starting</th><th>Status</th><th class="tnum">Hours</th>'+(showReason?'<th>Rejection reason</th>':'')+'</tr></thead><tbody>';
  rows.forEach(r=>{ const ed=['Draft','Rejected'].includes(r.status);
    h+=`<tr class="${ed?'lrow':''}" ${ed?`ondblclick="openWeek('${r.week_start}')"`:''}><td>${esc(r.week)}</td><td>${badge(r.status)}</td><td class="tnum">${r.hours}</td>${showReason?`<td class="muted">${esc(r.reason||'')}</td>`:''}</tr>`; });
  return h+'</tbody></table></div>';
}
function openWeek(ws){ setActive('timesheet'); renderTimesheet(ws); }
async function monthEnd(){ const ym=prompt('Which month? (YYYY-MM)', new Date().toISOString().slice(0,7)); if(!ym)return; const r=await API.month_end_export(ym); if(r&&r.ok){await API.open_file(r.path);} else if(r&&r.error)alert(r.error); }
</script></body></html>"""


def _logo_data_uri(cfg):
    """Find logo.png next to the exe / db and return it as a data URI so the
    real company logo shows in the UI. Empty string if none found."""
    import base64
    cands = [os.path.join(DATA_DIR, "logo.png")]
    dbp = cfg.get("db_path", "")
    if dbp:
        cands.append(os.path.join(os.path.dirname(dbp), "logo.png"))
    if getattr(sys, "frozen", False):
        cands.append(os.path.join(os.path.dirname(sys.executable), "logo.png"))
    for p in cands:
        try:
            if p and os.path.exists(p):
                if os.path.getsize(p) > 700_000:   # too big to embed safely
                    continue
                with open(p, "rb") as f:
                    return "data:image/png;base64," + base64.b64encode(
                        f.read()).decode("ascii")
        except Exception:
            pass
    return ""


def _trace(msg, reset=False):
    """Append a timestamped startup step to startup_log.txt. If the app
    freezes, the LAST line written tells us exactly where it hung."""
    try:
        base = _local_dir()
        with open(os.path.join(base, "startup_log.txt"),
                  "w" if reset else "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%H:%M:%S}  {msg}\n")
    except Exception:
        pass


def _log_crash(e):
    try:
        import traceback
        with open(os.path.join(DATA_DIR, "crash_log.txt"), "a", encoding="utf-8") as f:
            f.write(f"\n===== {dt.datetime.now().isoformat(timespec='seconds')} =====\n")
            f.write(traceback.format_exc())
    except Exception:
        pass


def main():
    import webview
    _trace(f"LAUNCH  {APP_NAME} {APP_VERSION}", reset=True)
    _trace("1. main() started, webview imported")
    try:
        _trace("2. opening database / building Api ...")
        api = Api()
        _trace("3. Api ready (database open)")
    except Exception as e:
        _trace("!! Api FAILED: " + repr(e))
        _log_crash(e)
        msg = (str(e).replace("&", "&amp;").replace("<", "&lt;")
               .replace(">", "&gt;").replace("\n", "<br>"))
        err = ("<body style='font-family:Segoe UI,sans-serif;padding:30px;color:#222'>"
               "<h2 style='color:#b22'>Cannot open the timesheet database</h2>"
               "<p>" + msg + "</p><p style='color:#555'>The app has <b>not</b> made "
               "any changes and has <b>not</b> created a blank database. Reconnect "
               "the drive and reopen.</p></body>")
        webview.create_window(APP_NAME + " - database error", html=err,
                              width=680, height=360)
        webview.start()
        return
    # Logo is loaded lazily via API.logo() AFTER render - inlining a large
    # PNG here makes the html payload too big for the WebView to load.
    _trace("4. building page HTML")
    html = INDEX_HTML.replace("{{LOGO}}", "").replace("{{LOGOCLS}}", "")
    # WebView2 browser profile on LOCAL disk, a FRESH folder per launch. A
    # reused profile can be locked by a lingering engine process and hang the
    # next launch forever; a unique folder each time avoids that. This matches
    # the build that opened reliably. Old folders are removed on next launch.
    base = _local_dir()
    store = None
    try:
        os.makedirs(base, exist_ok=True)
        store = os.path.join(base, f"run_{os.getpid()}_{dt.datetime.now():%H%M%S}")
        os.makedirs(store, exist_ok=True)
        shutil.rmtree(os.path.join(base, "profile"), ignore_errors=True)
        for p in glob.glob(os.path.join(base, "run_*")):
            if p != store:
                shutil.rmtree(p, ignore_errors=True)
    except Exception:
        store = None
    _trace("5. profile folder = " + str(store))
    try:
        _trace("6. creating window ...")
        co = (api.cfg.get("company_name") or "").strip()
        title = f"{co}  -  {APP_NAME}" if co else APP_NAME
        wkw = dict(html=html, js_api=api, width=1240, height=800,
                   min_size=(1040, 700))
        # Open maximised so the app fills the screen instead of a small window.
        # 'maximized' only exists on newer pywebview, so fall back cleanly on an
        # older one rather than failing to open at all.
        if api.cfg.get("start_maximized", True):
            try:
                api.window = webview.create_window(title, maximized=True, **wkw)
                _trace("6a. window created maximised")
            except TypeError:
                api.window = webview.create_window(title, **wkw)
                _trace("6a. this pywebview has no 'maximized' - normal window")
        else:
            api.window = webview.create_window(title, **wkw)
        _trace("7. window created; starting the window engine "
               "(if this is the LAST line, it froze HERE - in WebView2)")
        if store:
            webview.start(private_mode=False, storage_path=store)
        else:
            webview.start()
        _trace("8. window engine returned (app closed normally)")
    except Exception as e:
        _trace("!! start FAILED: " + repr(e))
        _log_crash(e)
        raise


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        _log_crash(e)
        raise
