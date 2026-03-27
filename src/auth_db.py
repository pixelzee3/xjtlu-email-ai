"""
本地用户库：SQLite user.db，邮箱唯一，密码 bcrypt 哈希。
默认种子用户仅写入密码哈希（明文不出现在仓库中）。
"""

from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

import bcrypt

DB_PATH = Path(__file__).parent / "user.db"

# 预置账号：夏众希 / 2196785278@qq.com（密码哈希由 bcrypt 生成，非明文）
SEED_EMAIL = "2196785278@qq.com"
SEED_USERNAME = "夏众希"
SEED_PASSWORD_HASH = (
    "$2b$12$mTOR83xu88jk/STeELYkUOU.7h.FBzMZ/n.6YfvUbU.c7bbqaxO.e"
)


def _conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with _conn() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS user_configs (
                user_id INTEGER PRIMARY KEY,
                config_json TEXT NOT NULL,
                FOREIGN KEY (user_id) REFERENCES users(id)
            );
            """
        )


def _hash_password(plain: str) -> str:
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt()).decode("ascii")


def verify_password(plain: str, password_hash: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode("utf-8"), password_hash.encode("ascii"))
    except Exception:
        return False


def _default_config_dict() -> dict[str, Any]:
    """与 config.example.json 结构一致的空配置。"""
    return {
        "ai": {
            "base_url": "https://api.openai.com/v1",
            "api_key": "",
            "model": "gpt-4o-mini",
        },
        "email": {
            "url": "https://mail.xjtlu.edu.cn/owa",
            "login_type": "cookie",
            "username": "",
            "password": "",
            "cookies": [],
            "cookie_file": "cookies.txt",
        },
        "selectors": {
            "search_box": "input[aria-label='Search']",
            "email_list": "table[role='presentation'] tr.zA",
            "email_date": "td.xY span",
            "email_subject": "div.y6 span",
            "email_body": "div.gs div.ii",
        },
        "browser": {"prelaunch": False},
    }


def _read_legacy_config_file() -> Optional[dict[str, Any]]:
    p = Path(__file__).parent / "config.json"
    if not p.exists():
        return None
    try:
        raw = p.read_text(encoding="utf-8").strip()
        if not raw:
            return None
        data = json.loads(raw)
        return data if isinstance(data, dict) else None
    except Exception:
        return None


def get_user_by_email(email: str) -> Optional[dict[str, Any]]:
    email = (email or "").strip().lower()
    if not email:
        return None
    with _conn() as c:
        row = c.execute(
            "SELECT id, username, email, password_hash, created_at FROM users WHERE email = ?",
            (email,),
        ).fetchone()
    if not row:
        return None
    return dict(row)


def get_user_by_id(user_id: int) -> Optional[dict[str, Any]]:
    with _conn() as c:
        row = c.execute(
            "SELECT id, username, email, password_hash, created_at FROM users WHERE id = ?",
            (user_id,),
        ).fetchone()
    if not row:
        return None
    return dict(row)


def load_user_config(user_id: int) -> dict[str, Any]:
    with _conn() as c:
        row = c.execute(
            "SELECT config_json FROM user_configs WHERE user_id = ?",
            (user_id,),
        ).fetchone()
    if not row:
        return _default_config_dict()
    try:
        data = json.loads(row["config_json"])
        return data if isinstance(data, dict) else _default_config_dict()
    except Exception:
        return _default_config_dict()


def save_user_config(user_id: int, config: dict[str, Any]) -> None:
    payload = json.dumps(config, ensure_ascii=False, indent=4)
    with _conn() as c:
        c.execute(
            """
            INSERT INTO user_configs (user_id, config_json) VALUES (?, ?)
            ON CONFLICT(user_id) DO UPDATE SET config_json = excluded.config_json
            """,
            (user_id, payload),
        )


def ensure_seed_user_and_migrate_legacy() -> None:
    """创建种子用户；若尚无配置行，则从旧版 config.json 导入（视为夏众希的数据）。"""
    init_db()
    legacy = _read_legacy_config_file()

    with _conn() as c:
        row = c.execute(
            "SELECT id FROM users WHERE email = ?", (SEED_EMAIL.lower(),)
        ).fetchone()

        if not row:
            now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            cur = c.execute(
                """
                INSERT INTO users (username, email, password_hash, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (SEED_USERNAME, SEED_EMAIL.lower(), SEED_PASSWORD_HASH, now),
            )
            user_id = cur.lastrowid
        else:
            user_id = row["id"]

        has_cfg = c.execute(
            "SELECT 1 FROM user_configs WHERE user_id = ?", (user_id,)
        ).fetchone()

        if not has_cfg:
            cfg = legacy if legacy else _default_config_dict()
            payload = json.dumps(cfg, ensure_ascii=False, indent=4)
            c.execute(
                """
                INSERT INTO user_configs (user_id, config_json) VALUES (?, ?)
                ON CONFLICT(user_id) DO UPDATE SET config_json = excluded.config_json
                """,
                (user_id, payload),
            )


def create_user(username: str, email: str, password: str) -> tuple[bool, str]:
    email = email.strip().lower()
    username = username.strip()
    if not email or "@" not in email:
        return False, "邮箱格式无效"
    if not username:
        return False, "用户名不能为空"
    if len(password) < 6:
        return False, "密码至少 6 位"
    ph = _hash_password(password)
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    try:
        with _conn() as c:
            cur = c.execute(
                """
                INSERT INTO users (username, email, password_hash, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (username, email, ph, now),
            )
            uid = cur.lastrowid
        save_user_config(uid, _default_config_dict())
        return True, "ok"
    except sqlite3.IntegrityError:
        return False, "该邮箱已注册"
    except Exception as e:
        return False, str(e)


def verify_login(email: str, password: str) -> Optional[dict[str, Any]]:
    u = get_user_by_email(email)
    if not u:
        return None
    if not verify_password(password, u["password_hash"]):
        return None
    return {"id": u["id"], "username": u["username"], "email": u["email"]}

def update_username(user_id: int, new_username: str) -> tuple[bool, str]:
    new_username = new_username.strip()
    if not new_username:
        return False, "用户名不能为空"
    try:
        with _conn() as c:
            c.execute("UPDATE users SET username = ? WHERE id = ?", (new_username, user_id))
        return True, "ok"
    except Exception as e:
        return False, str(e)
