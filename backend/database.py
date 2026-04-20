"""
database.py — SQLite database for Payless Automation Hub
Handles users and per-user settings.
"""

import sqlite3
import json
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / 'data' / 'users.db'


def get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create tables if they don't exist."""
    conn = get_db()
    c = conn.cursor()

    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            username        TEXT    UNIQUE NOT NULL,
            password_hash   TEXT    NOT NULL,
            is_admin        INTEGER DEFAULT 0,
            allowed_modules TEXT    DEFAULT '["*"]',
            created_at      TEXT    DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS user_settings (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id     INTEGER NOT NULL,
            module_id   TEXT    NOT NULL,
            key         TEXT    NOT NULL,
            value       TEXT    NOT NULL,
            updated_at  TEXT    DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, module_id, key),
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        )
    ''')

    conn.commit()
    conn.close()


# ── User CRUD ──────────────────────────────────────────────────────────────────

def get_user_by_username(username: str):
    conn = get_db()
    row = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_by_id(user_id: int):
    conn = get_db()
    row = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_users():
    conn = get_db()
    rows = conn.execute(
        'SELECT id, username, is_admin, allowed_modules, created_at FROM users ORDER BY created_at'
    ).fetchall()
    conn.close()
    result = []
    for r in rows:
        u = dict(r)
        u['allowed_modules'] = json.loads(u['allowed_modules'])
        result.append(u)
    return result


def create_user(username: str, password_hash: str, is_admin: bool = False,
                allowed_modules: list = None):
    modules = json.dumps(allowed_modules if allowed_modules is not None else ['*'])
    conn = get_db()
    conn.execute(
        'INSERT INTO users (username, password_hash, is_admin, allowed_modules) VALUES (?, ?, ?, ?)',
        (username, password_hash, 1 if is_admin else 0, modules)
    )
    conn.commit()
    conn.close()


def update_user(user_id: int, allowed_modules: list = None, password_hash: str = None):
    conn = get_db()
    if allowed_modules is not None:
        conn.execute(
            'UPDATE users SET allowed_modules = ? WHERE id = ?',
            (json.dumps(allowed_modules), user_id)
        )
    if password_hash is not None:
        conn.execute(
            'UPDATE users SET password_hash = ? WHERE id = ?',
            (password_hash, user_id)
        )
    conn.commit()
    conn.close()


def delete_user(user_id: int):
    conn = get_db()
    conn.execute('DELETE FROM user_settings WHERE user_id = ?', (user_id,))
    conn.execute('DELETE FROM users WHERE id = ?', (user_id,))
    conn.commit()
    conn.close()


# ── Per-user settings ──────────────────────────────────────────────────────────

def get_user_setting(user_id: int, module_id: str, key: str):
    conn = get_db()
    row = conn.execute(
        'SELECT value FROM user_settings WHERE user_id = ? AND module_id = ? AND key = ?',
        (user_id, module_id, key)
    ).fetchone()
    conn.close()
    return json.loads(row['value']) if row else None


def set_user_setting(user_id: int, module_id: str, key: str, value):
    conn = get_db()
    conn.execute('''
        INSERT INTO user_settings (user_id, module_id, key, value, updated_at)
        VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(user_id, module_id, key) DO UPDATE SET
            value      = excluded.value,
            updated_at = excluded.updated_at
    ''', (user_id, module_id, key, json.dumps(value)))
    conn.commit()
    conn.close()


def get_module_settings(user_id: int, module_id: str) -> dict:
    conn = get_db()
    rows = conn.execute(
        'SELECT key, value FROM user_settings WHERE user_id = ? AND module_id = ?',
        (user_id, module_id)
    ).fetchall()
    conn.close()
    return {row['key']: json.loads(row['value']) for row in rows}


# ── Commission config migration ─────────────────────────────────────────────────

def migrate_commission_config(config_path, admin_user_id: int):
    """
    One-time migration: if commission_config.json exists and the admin user
    has no rep mapping saved yet, import the reps from the file into the DB.
    """
    if not config_path.exists():
        return
    existing = get_user_setting(admin_user_id, 'commission', 'reps')
    if existing:
        return  # already migrated
    try:
        data = json.loads(config_path.read_text())
        reps = data.get('reps', [])
        if reps:
            set_user_setting(admin_user_id, 'commission', 'reps', reps)
            print(f'  ✅  Migrated {len(reps)} commission rep(s) from config file to database.')
    except Exception as e:
        print(f'  ⚠️   Could not migrate commission_config.json: {e}')
