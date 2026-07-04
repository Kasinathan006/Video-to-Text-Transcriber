#!/usr/bin/env python3
"""
database.py - VoxDoc AI persistence, accounts, quotas & licensing layer

SQLite-backed (zero external services) storage for:
  * users        - accounts with PBKDF2-hashed passwords and subscription tiers
  * sessions     - bearer tokens for API/dashboard auth
  * jobs         - transcription jobs (survive server restarts)
  * license_keys - sellable upgrade keys (mint with manage.py, redeem in-app)
  * usage        - per-job minutes ledger; monthly quotas are computed from it

Thread-safe: a single connection guarded by a lock (SQLite serialized mode).
"""

import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import threading
import time
import uuid
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
STORAGE_DIR = APP_DIR / "storage"
DB_PATH = STORAGE_DIR / "voxdoc.db"

SESSION_TTL_SECONDS = 30 * 24 * 3600      # 30-day login sessions
PBKDF2_ITERATIONS = 200_000

# Subscription tiers (Build Guide §6 — Monetization & Business Models)
TIERS = {
    "free": {
        "name": "Starter (Free)",
        "price": "$0 / mo",
        "monthly_minutes": 60,
        "max_upload_mb": 500,
        "concurrent_jobs": 1,
    },
    "pro": {
        "name": "Creator Pro",
        "price": "$19 / mo",
        "monthly_minutes": 900,
        "max_upload_mb": 10 * 1024,
        "concurrent_jobs": 2,
    },
    "agency": {
        "name": "Agency / Team",
        "price": "$49 / mo",
        "monthly_minutes": 3000,
        "max_upload_mb": 50 * 1024,
        "concurrent_jobs": 4,
    },
}

_LOCK = threading.Lock()
_CONN: sqlite3.Connection | None = None


# ----------------------------------------------------------------------------
# Config (.env loader — no external dependency)
# ----------------------------------------------------------------------------

def load_env(path: Path = APP_DIR / ".env"):
    """Load KEY=VALUE lines from .env into os.environ (existing vars win)."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


# ----------------------------------------------------------------------------
# Connection & schema
# ----------------------------------------------------------------------------

def get_conn() -> sqlite3.Connection:
    global _CONN
    if _CONN is None:
        STORAGE_DIR.mkdir(parents=True, exist_ok=True)
        _CONN = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        _CONN.row_factory = sqlite3.Row
        _CONN.execute("PRAGMA journal_mode=WAL")
        _CONN.execute("PRAGMA foreign_keys=ON")
    return _CONN


def init_db():
    conn = get_conn()
    with _LOCK:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id              TEXT PRIMARY KEY,
            email           TEXT UNIQUE NOT NULL,
            full_name       TEXT DEFAULT '',
            pw_salt         TEXT NOT NULL,
            pw_hash         TEXT NOT NULL,
            tier            TEXT NOT NULL DEFAULT 'free',
            tier_expires_at REAL,               -- NULL = never expires
            is_admin        INTEGER NOT NULL DEFAULT 0,
            created_at      REAL NOT NULL
        );
        CREATE TABLE IF NOT EXISTS sessions (
            token      TEXT PRIMARY KEY,
            user_id    TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            created_at REAL NOT NULL,
            expires_at REAL NOT NULL
        );
        CREATE TABLE IF NOT EXISTS jobs (
            id           TEXT PRIMARY KEY,
            user_id      TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            filename     TEXT NOT NULL,
            size_mb      REAL NOT NULL DEFAULT 0,
            engine       TEXT NOT NULL,
            params       TEXT NOT NULL DEFAULT '{}',
            status       TEXT NOT NULL DEFAULT 'queued',
            stage        TEXT NOT NULL DEFAULT 'queued',
            progress     INTEGER NOT NULL DEFAULT 0,
            detail       TEXT DEFAULT '',
            duration_sec REAL,
            words        INTEGER,
            docx_name    TEXT,
            model_label  TEXT,
            created_at   REAL NOT NULL,
            started_at   REAL,
            finished_at  REAL
        );
        CREATE INDEX IF NOT EXISTS idx_jobs_user ON jobs(user_id, created_at DESC);
        CREATE TABLE IF NOT EXISTS license_keys (
            key           TEXT PRIMARY KEY,
            tier          TEXT NOT NULL,
            duration_days INTEGER NOT NULL DEFAULT 30,
            created_at    REAL NOT NULL,
            redeemed_by   TEXT REFERENCES users(id),
            redeemed_at   REAL
        );
        CREATE TABLE IF NOT EXISTS usage (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            job_id     TEXT,
            minutes    REAL NOT NULL,
            month      TEXT NOT NULL,            -- 'YYYY-MM'
            created_at REAL NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_usage_month ON usage(user_id, month);
        """)
        conn.commit()


# ----------------------------------------------------------------------------
# Password & session helpers
# ----------------------------------------------------------------------------

def _hash_password(password: str, salt: bytes) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                               salt, PBKDF2_ITERATIONS).hex()


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def validate_credentials(email: str, password: str):
    if not EMAIL_RE.match(email or ""):
        raise ValueError("Please enter a valid email address.")
    if len(password or "") < 8:
        raise ValueError("Password must be at least 8 characters long.")


def create_user(email: str, password: str, full_name: str = "",
                tier: str = "free", is_admin: bool = False) -> dict:
    validate_credentials(email, password)
    if tier not in TIERS:
        raise ValueError(f"Unknown tier '{tier}'. Valid: {', '.join(TIERS)}")
    salt = secrets.token_bytes(16)
    user_id = uuid.uuid4().hex
    conn = get_conn()
    with _LOCK:
        try:
            conn.execute(
                "INSERT INTO users (id, email, full_name, pw_salt, pw_hash, tier, "
                "is_admin, created_at) VALUES (?,?,?,?,?,?,?,?)",
                (user_id, email.lower().strip(), full_name.strip(), salt.hex(),
                 _hash_password(password, salt), tier, int(is_admin), time.time()),
            )
            conn.commit()
        except sqlite3.IntegrityError:
            raise ValueError("An account with this email already exists.")
    return get_user(user_id)


def verify_login(email: str, password: str) -> dict | None:
    conn = get_conn()
    with _LOCK:
        row = conn.execute("SELECT * FROM users WHERE email = ?",
                           (email.lower().strip(),)).fetchone()
    if row is None:
        return None
    expected = row["pw_hash"]
    actual = _hash_password(password, bytes.fromhex(row["pw_salt"]))
    if not hmac.compare_digest(expected, actual):
        return None
    return dict(row)


def create_session(user_id: str) -> str:
    token = secrets.token_urlsafe(32)
    now = time.time()
    conn = get_conn()
    with _LOCK:
        conn.execute("INSERT INTO sessions (token, user_id, created_at, expires_at) "
                     "VALUES (?,?,?,?)", (token, user_id, now, now + SESSION_TTL_SECONDS))
        conn.execute("DELETE FROM sessions WHERE expires_at < ?", (now,))
        conn.commit()
    return token


def get_user_by_token(token: str) -> dict | None:
    if not token:
        return None
    conn = get_conn()
    with _LOCK:
        row = conn.execute(
            "SELECT u.* FROM sessions s JOIN users u ON u.id = s.user_id "
            "WHERE s.token = ? AND s.expires_at > ?", (token, time.time())).fetchone()
    return dict(row) if row else None


def delete_session(token: str):
    conn = get_conn()
    with _LOCK:
        conn.execute("DELETE FROM sessions WHERE token = ?", (token,))
        conn.commit()


def get_user(user_id: str) -> dict | None:
    conn = get_conn()
    with _LOCK:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    return dict(row) if row else None


# ----------------------------------------------------------------------------
# Tiers, quotas & usage
# ----------------------------------------------------------------------------

def effective_tier(user: dict) -> str:
    """Paid tiers fall back to 'free' after their expiry timestamp passes."""
    tier = user.get("tier", "free")
    if tier == "free" or tier not in TIERS:
        return "free" if tier not in TIERS else tier
    expires = user.get("tier_expires_at")
    if expires is not None and expires < time.time():
        return "free"
    return tier


def current_month() -> str:
    return time.strftime("%Y-%m")


def minutes_used_this_month(user_id: str) -> float:
    conn = get_conn()
    with _LOCK:
        row = conn.execute(
            "SELECT COALESCE(SUM(minutes), 0) AS m FROM usage WHERE user_id = ? AND month = ?",
            (user_id, current_month())).fetchone()
    return float(row["m"])


def quota_summary(user: dict) -> dict:
    tier_key = effective_tier(user)
    tier = TIERS[tier_key]
    used = minutes_used_this_month(user["id"])
    return {
        "tier": tier_key,
        "tier_name": tier["name"],
        "tier_price": tier["price"],
        "tier_expires_at": user.get("tier_expires_at") if tier_key != "free" else None,
        "monthly_minutes": tier["monthly_minutes"],
        "minutes_used": round(used, 1),
        "minutes_remaining": round(max(0.0, tier["monthly_minutes"] - used), 1),
        "max_upload_mb": tier["max_upload_mb"],
        "concurrent_jobs": tier["concurrent_jobs"],
        "month": current_month(),
    }


def record_usage(user_id: str, job_id: str, minutes: float):
    conn = get_conn()
    with _LOCK:
        conn.execute("INSERT INTO usage (user_id, job_id, minutes, month, created_at) "
                     "VALUES (?,?,?,?,?)",
                     (user_id, job_id, round(minutes, 2), current_month(), time.time()))
        conn.commit()


def set_user_tier(user_id: str, tier: str, duration_days: int | None = 30):
    if tier not in TIERS:
        raise ValueError(f"Unknown tier '{tier}'. Valid: {', '.join(TIERS)}")
    expires = None
    if tier != "free" and duration_days is not None:
        expires = time.time() + duration_days * 86400
    conn = get_conn()
    with _LOCK:
        cur = conn.execute("UPDATE users SET tier = ?, tier_expires_at = ? WHERE id = ?",
                           (tier, expires, user_id))
        conn.commit()
    if cur.rowcount == 0:
        raise ValueError("User not found.")


def list_users() -> list[dict]:
    conn = get_conn()
    with _LOCK:
        rows = conn.execute("SELECT * FROM users ORDER BY created_at").fetchall()
    return [dict(r) for r in rows]


# ----------------------------------------------------------------------------
# License keys (sell via Gumroad / Lemon Squeezy / invoice — redeem in-app)
# ----------------------------------------------------------------------------

def generate_license_keys(tier: str, duration_days: int = 30, count: int = 1) -> list[str]:
    if tier not in TIERS or tier == "free":
        raise ValueError("License keys can only be minted for paid tiers: pro, agency.")
    if not (1 <= count <= 100):
        raise ValueError("count must be between 1 and 100.")
    keys = []
    conn = get_conn()
    with _LOCK:
        for _ in range(count):
            raw = secrets.token_hex(8).upper()
            key = f"VOX-{tier.upper()}-{raw[:4]}-{raw[4:8]}-{raw[8:12]}-{raw[12:16]}"
            conn.execute("INSERT INTO license_keys (key, tier, duration_days, created_at) "
                         "VALUES (?,?,?,?)", (key, tier, duration_days, time.time()))
            keys.append(key)
        conn.commit()
    return keys


def redeem_license_key(user_id: str, key: str) -> dict:
    key = (key or "").strip().upper()
    conn = get_conn()
    with _LOCK:
        row = conn.execute("SELECT * FROM license_keys WHERE key = ?", (key,)).fetchone()
        if row is None:
            raise ValueError("Invalid license key.")
        if row["redeemed_by"]:
            raise ValueError("This license key has already been redeemed.")
        expires = time.time() + row["duration_days"] * 86400
        conn.execute("UPDATE license_keys SET redeemed_by = ?, redeemed_at = ? WHERE key = ?",
                     (user_id, time.time(), key))
        conn.execute("UPDATE users SET tier = ?, tier_expires_at = ? WHERE id = ?",
                     (row["tier"], expires, user_id))
        conn.commit()
    return {"tier": row["tier"], "duration_days": row["duration_days"]}


def list_license_keys(include_redeemed: bool = True) -> list[dict]:
    conn = get_conn()
    with _LOCK:
        q = "SELECT * FROM license_keys ORDER BY created_at DESC"
        if not include_redeemed:
            q = "SELECT * FROM license_keys WHERE redeemed_by IS NULL ORDER BY created_at DESC"
        rows = conn.execute(q).fetchall()
    return [dict(r) for r in rows]


# ----------------------------------------------------------------------------
# Jobs (persistent — survive server restarts)
# ----------------------------------------------------------------------------

def create_job(user_id: str, filename: str, size_mb: float, engine: str, params: dict,
               job_id: str | None = None) -> dict:
    job_id = job_id or uuid.uuid4().hex[:12]
    conn = get_conn()
    with _LOCK:
        conn.execute(
            "INSERT INTO jobs (id, user_id, filename, size_mb, engine, params, status, "
            "stage, progress, detail, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (job_id, user_id, filename, round(size_mb, 2), engine,
             json.dumps(params), "queued", "queued", 0, "Waiting for a worker...", time.time()))
        conn.commit()
    return get_job(job_id)


def get_job(job_id: str) -> dict | None:
    conn = get_conn()
    with _LOCK:
        row = conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
    return dict(row) if row else None


def update_job(job_id: str, **fields):
    if not fields:
        return
    cols = ", ".join(f"{k} = ?" for k in fields)
    conn = get_conn()
    with _LOCK:
        conn.execute(f"UPDATE jobs SET {cols} WHERE id = ?", (*fields.values(), job_id))
        conn.commit()


def delete_job(job_id: str):
    conn = get_conn()
    with _LOCK:
        conn.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
        conn.commit()


def list_jobs_for_user(user_id: str, limit: int = 100) -> list[dict]:
    conn = get_conn()
    with _LOCK:
        rows = conn.execute("SELECT * FROM jobs WHERE user_id = ? "
                            "ORDER BY created_at DESC LIMIT ?", (user_id, limit)).fetchall()
    return [dict(r) for r in rows]


def count_active_jobs(user_id: str) -> int:
    conn = get_conn()
    with _LOCK:
        row = conn.execute("SELECT COUNT(*) AS c FROM jobs WHERE user_id = ? "
                           "AND status IN ('queued', 'processing')", (user_id,)).fetchone()
    return int(row["c"])


def fail_interrupted_jobs() -> int:
    """Called at server startup: jobs left mid-flight by a previous run are failed
    honestly instead of appearing stuck forever."""
    conn = get_conn()
    with _LOCK:
        cur = conn.execute(
            "UPDATE jobs SET status='failed', stage='error', progress=100, "
            "detail='Interrupted by a server restart — please submit the file again.', "
            "finished_at=? WHERE status IN ('queued', 'processing')", (time.time(),))
        conn.commit()
    return cur.rowcount
