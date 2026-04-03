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
            CREATE TABLE IF NOT EXISTS digest_jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                period_label TEXT NOT NULL,
                run_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'pending',
                payload_json TEXT NOT NULL,
                error_message TEXT,
                artifact_id INTEGER,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT,
                FOREIGN KEY (user_id) REFERENCES users(id)
            );
            CREATE INDEX IF NOT EXISTS idx_digest_jobs_pending_run
                ON digest_jobs(status, run_at);
            CREATE INDEX IF NOT EXISTS idx_digest_jobs_user_period
                ON digest_jobs(user_id, period_label, status);
            CREATE TABLE IF NOT EXISTS digest_artifacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                job_id INTEGER,
                period_label TEXT NOT NULL,
                cadence TEXT NOT NULL,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                status TEXT NOT NULL,
                summary_text TEXT,
                result_json TEXT,
                error_message TEXT,
                FOREIGN KEY (user_id) REFERENCES users(id)
            );
            CREATE INDEX IF NOT EXISTS idx_digest_artifacts_user_time
                ON digest_artifacts(user_id, finished_at DESC);
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
        "digest": {
            "enabled": False,
            "cadence": "daily",
            "local_time": "08:00",
            "timezone": "",
            "weekday": 0,
            "keyword": "",
            "instruction": "请用自然语气总结这些邮件，重点关注活动、课程作业和重要事项。",
            "mode": "daily",
            "email_count": 10,
            "last_success_at": None,
            "last_success_period": None,
        },
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
        cfg = _default_config_dict()
    else:
        try:
            data = json.loads(row["config_json"])
            cfg = data if isinstance(data, dict) else _default_config_dict()
        except Exception:
            cfg = _default_config_dict()
    d_raw = cfg.get("digest")
    if not isinstance(d_raw, dict):
        cfg["digest"] = _default_config_dict()["digest"].copy()
    else:
        base = _default_config_dict()["digest"]
        merged = dict(base)
        merged.update(d_raw)
        cfg["digest"] = merged
    return cfg


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


def list_user_ids() -> list[int]:
    with _conn() as c:
        rows = c.execute("SELECT id FROM users ORDER BY id").fetchall()
    return [int(r["id"]) for r in rows]


def digest_has_success_artifact(user_id: int, period_label: str) -> bool:
    with _conn() as c:
        row = c.execute(
            """
            SELECT 1 FROM digest_artifacts
            WHERE user_id = ? AND period_label = ? AND status = 'success'
            LIMIT 1
            """,
            (user_id, period_label),
        ).fetchone()
    return row is not None


def digest_has_active_job_for_period(user_id: int, period_label: str) -> bool:
    with _conn() as c:
        row = c.execute(
            """
            SELECT 1 FROM digest_jobs
            WHERE user_id = ? AND period_label = ?
              AND status IN ('pending', 'running')
            LIMIT 1
            """,
            (user_id, period_label),
        ).fetchone()
    return row is not None


def digest_has_terminal_job_for_period(user_id: int, period_label: str) -> bool:
    """本周期已有终态任务（成功/失败/跳过），避免失败后每轮重复入队。"""
    with _conn() as c:
        row = c.execute(
            """
            SELECT 1 FROM digest_jobs
            WHERE user_id = ? AND period_label = ?
              AND status IN ('completed', 'failed', 'skipped')
            LIMIT 1
            """,
            (user_id, period_label),
        ).fetchone()
    return row is not None


def digest_enqueue_job(
    user_id: int, period_label: str, run_at_iso: str, payload: dict[str, Any]
) -> Optional[int]:
    if digest_has_success_artifact(user_id, period_label):
        return None
    if digest_has_active_job_for_period(user_id, period_label):
        return None
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    payload_json = json.dumps(payload, ensure_ascii=False)
    with _conn() as c:
        cur = c.execute(
            """
            INSERT INTO digest_jobs
            (user_id, period_label, run_at, status, payload_json, created_at)
            VALUES (?, ?, ?, 'pending', ?, ?)
            """,
            (user_id, period_label, run_at_iso, payload_json, now),
        )
        return int(cur.lastrowid)


def digest_claim_next_job(now_iso: str) -> Optional[dict[str, Any]]:
    with _conn() as c:
        row = c.execute(
            """
            SELECT * FROM digest_jobs
            WHERE status = 'pending' AND run_at <= ?
            ORDER BY run_at ASC, id ASC
            LIMIT 1
            """,
            (now_iso,),
        ).fetchone()
        if not row:
            return None
        jid = int(row["id"])
        c.execute(
            """
            UPDATE digest_jobs
            SET status = 'running', started_at = ?
            WHERE id = ? AND status = 'pending'
            """,
            (now_iso, jid),
        )
        if c.total_changes != 1:
            return None
        row2 = c.execute("SELECT * FROM digest_jobs WHERE id = ?", (jid,)).fetchone()
    return dict(row2) if row2 else None


def digest_finish_job(
    job_id: int,
    status: str,
    *,
    error_message: Optional[str] = None,
    artifact_id: Optional[int] = None,
) -> None:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    with _conn() as c:
        c.execute(
            """
            UPDATE digest_jobs
            SET status = ?, finished_at = ?, error_message = ?, artifact_id = ?
            WHERE id = ?
            """,
            (status, now, error_message, artifact_id, job_id),
        )


def digest_insert_artifact(
    user_id: int,
    job_id: Optional[int],
    period_label: str,
    cadence: str,
    started_at: str,
    finished_at: str,
    status: str,
    summary_text: Optional[str],
    result_json: Optional[str],
    error_message: Optional[str],
) -> int:
    with _conn() as c:
        cur = c.execute(
            """
            INSERT INTO digest_artifacts
            (user_id, job_id, period_label, cadence, started_at, finished_at,
             status, summary_text, result_json, error_message)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user_id,
                job_id,
                period_label,
                cadence,
                started_at,
                finished_at,
                status,
                summary_text,
                result_json,
                error_message,
            ),
        )
        return int(cur.lastrowid)


def digest_list_artifacts(user_id: int, limit: int = 30) -> list[dict[str, Any]]:
    lim = max(1, min(int(limit), 100))
    with _conn() as c:
        rows = c.execute(
            """
            SELECT id, user_id, job_id, period_label, cadence, started_at, finished_at,
                   status, summary_text, result_json, error_message
            FROM digest_artifacts
            WHERE user_id = ?
            ORDER BY COALESCE(finished_at, started_at) DESC, id DESC
            LIMIT ?
            """,
            (user_id, lim),
        ).fetchall()
    return [dict(r) for r in rows]


def digest_update_user_success_meta(
    user_id: int, period_label: str, success_at_iso: str
) -> None:
    cfg = load_user_config(user_id)
    if "digest" not in cfg or not isinstance(cfg["digest"], dict):
        cfg["digest"] = _default_config_dict()["digest"].copy()
    cfg["digest"]["last_success_at"] = success_at_iso
    cfg["digest"]["last_success_period"] = period_label
    save_user_config(user_id, cfg)
