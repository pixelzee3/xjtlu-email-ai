"""
定时 Digest：周期标签、下次运行提示、是否到点入队（纯函数，便于测试）。
"""
from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any, Optional


def _parse_hhmm(s: str) -> tuple[int, int]:
    parts = (s or "08:00").strip().split(":")
    try:
        h = int(parts[0]) if parts else 8
    except ValueError:
        h = 8
    try:
        m = int(parts[1]) if len(parts) > 1 else 0
    except ValueError:
        m = 0
    return max(0, min(23, h)), max(0, min(59, m))


def default_digest_dict() -> dict[str, Any]:
    return {
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
    }


def merge_digest_defaults(raw: Any) -> dict[str, Any]:
    base = default_digest_dict()
    if not isinstance(raw, dict):
        return dict(base)
    out = dict(base)
    for k, v in raw.items():
        if k in out or k in (
            "enabled",
            "cadence",
            "local_time",
            "timezone",
            "weekday",
            "keyword",
            "instruction",
            "mode",
            "email_count",
            "last_success_at",
            "last_success_period",
        ):
            out[k] = v
    return out


def compute_period_label(now: datetime, cadence: str) -> str:
    c = (cadence or "daily").strip().lower()
    if c == "weekly":
        y, w, _ = now.isocalendar()
        return f"{y}-W{w:02d}"
    return now.strftime("%Y-%m-%d")


def period_slot_start(
    period_label: str, cadence: str, digest: dict[str, Any], now: datetime
) -> datetime:
    """本周期内计划运行的时间点（本地时间，与 now 无时区一致）。"""
    h, m = _parse_hhmm(digest.get("local_time", "08:00"))
    c = (cadence or "daily").strip().lower()
    if c == "weekly":
        if "-W" not in period_label:
            y, w, _ = now.isocalendar()
            period_label = f"{y}-W{w:02d}"
        year_s, w_s = period_label.split("-W", 1)
        year = int(year_s)
        week = int(w_s)
        wd = int(digest.get("weekday", 0))
        wd = max(0, min(6, wd))
        iso_day = wd + 1
        d = datetime.fromisocalendar(year, week, iso_day)
        return d.replace(hour=h, minute=m, second=0, microsecond=0)
    try:
        d = datetime.strptime(period_label, "%Y-%m-%d")
    except ValueError:
        d = now.replace(hour=0, minute=0, second=0, microsecond=0)
    return d.replace(hour=h, minute=m, second=0, microsecond=0)


def is_digest_due(
    digest: dict[str, Any], now: datetime, period_label: str, cadence: str
) -> bool:
    if not digest.get("enabled"):
        return False
    slot = period_slot_start(period_label, cadence, digest, now)
    return now >= slot


def compute_next_run_hint(digest: dict[str, Any], now: Optional[datetime] = None) -> Optional[str]:
    if not digest.get("enabled"):
        return None
    now = now or datetime.now()
    cadence = (digest.get("cadence") or "daily").strip().lower()
    h, m = _parse_hhmm(digest.get("local_time", "08:00"))
    if cadence == "daily":
        candidate = now.replace(hour=h, minute=m, second=0, microsecond=0)
        if candidate <= now:
            candidate += timedelta(days=1)
        return candidate.isoformat(timespec="seconds")
    wd = max(0, min(6, int(digest.get("weekday", 0))))
    iso_day = wd + 1
    y, w, _ = now.isocalendar()
    anchor = datetime.fromisocalendar(int(y), int(w), iso_day).replace(
        hour=h, minute=m, second=0, microsecond=0
    )
    while anchor <= now:
        anchor += timedelta(days=7)
    return anchor.isoformat(timespec="seconds")


def build_execute_request_payload(digest: dict[str, Any]) -> dict[str, Any]:
    """生成存入 digest_jobs 的 ExecuteRequest 字段 dict。"""
    mode = (digest.get("mode") or "daily").strip().lower()
    if mode != "daily":
        mode = "daily"
    ec = digest.get("email_count", 10)
    try:
        ec = int(ec)
    except (TypeError, ValueError):
        ec = 10
    return {
        "keyword": (digest.get("keyword") or "").strip(),
        "instruction": (digest.get("instruction") or "").strip()
        or "请用自然语气总结这些邮件，重点关注活动、课程作业和重要事项。",
        "mode": mode,
        "email_count": ec,
        "indices": None,
        "date_from": None,
        "date_to": None,
    }
