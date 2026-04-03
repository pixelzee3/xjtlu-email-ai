"""
深度扫描：基于已缓存的正文与元数据做确定性优先级评分（不访问邮箱）。
"""
from __future__ import annotations

import re
from datetime import datetime
from typing import Any, Optional

from main import parse_email_date_for_filter

# 与 main.EMAIL_CATEGORY_LABELS 顺序无关，按类别名给基础分
_CATEGORY_BASE: dict[str, float] = {
    "系统/安全": 26.0,
    "课程/学习": 19.0,
    "招聘/实习": 17.0,
    "活动": 11.0,
    "校友通知": 8.0,
    "体育/场馆": 7.0,
    "其他": 5.0,
}

# (正则, 加分, 人类可读原因)
_URGENCY_RULES: list[tuple[re.Pattern, float, str]] = [
    (re.compile(r"\burgent\b|urgency|asap|immediately", re.I), 14.0, "含紧急措辞"),
    (re.compile(r"action\s*required|需要.*行动|请.*处理|务必|尽快", re.I), 12.0, "可能需行动"),
    (re.compile(r"deadline|due\s+date|截止|截止日期|最迟|before\s+\d", re.I), 16.0, "含截止/期限相关"),
    (re.compile(r"reminder|温馨提示|gentle\s+reminder|勿忘", re.I), 8.0, "提醒类"),
    (re.compile(r"验证码|verification\s*code|otp|mfa|password\s*reset|重置密码", re.I), 15.0, "安全/验证码相关"),
    (re.compile(r"survey\s+closes|问卷.*关闭|last\s+chance|final\s+notice", re.I), 11.0, "即将截止的问卷/通知"),
    (re.compile(r"\btoday\b|今日|明天|tomorrow|within\s+24|24\s*小时内", re.I), 9.0, "时间紧迫表述"),
    (re.compile(r"exam|考试|assignment\s*due|作业.*提交|ddl", re.I), 10.0, "学业/作业相关"),
]

_DATE_HINT = re.compile(
    r"(20\d{2}[-/年]\d{1,2}[-/月]\d{1,2})|"
    r"(\d{1,2}\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*,?\s*20\d{2})",
    re.I,
)


def _normalize_subject(s: str) -> str:
    t = (s or "").strip().lower()
    t = re.sub(r"^(re|fw|fwd|回复|转发)\s*:\s*", "", t, flags=re.I)
    return t[:120]


def deadline_hint_from_text(text: str) -> Optional[str]:
    if not text or len(text) < 8:
        return None
    m = _DATE_HINT.search(text[:4000])
    if m:
        return (m.group(0) or "").strip()[:80]
    return None


def _recency_bonus(parsed: Optional[datetime], now: datetime) -> tuple[float, Optional[str]]:
    if parsed is None:
        return 0.0, None
    try:
        d = parsed.date() if hasattr(parsed, "date") else parsed
        n = now.date() if hasattr(now, "date") else now
        days_ago = (n - d).days
    except Exception:
        return 0.0, None
    if days_ago < 0:
        return 4.0, "日期在未来/解析异常"
    if days_ago <= 1:
        return 18.0, "最近 1 天内"
    if days_ago <= 3:
        return 14.0, "最近 3 天内"
    if days_ago <= 7:
        return 10.0, "最近一周内"
    if days_ago <= 14:
        return 5.0, "最近两周内"
    if days_ago <= 30:
        return 2.0, "最近一月内"
    return 0.0, None


def compute_priority_for_sample(
    sample: dict[str, Any],
    *,
    now: Optional[datetime] = None,
) -> dict[str, Any]:
    """
    返回要合并进 sample 的字段（不修改 sample）。
    """
    now = now or datetime.now()
    sender = str(sample.get("sender") or "")
    subj = str(sample.get("subject") or "")
    body = str(sample.get("body") or "")
    cat = str(sample.get("category") or "其他")
    blob_head = f"{sender}\n{subj}\n{body[:2500]}".lower()

    base = _CATEGORY_BASE.get(cat, _CATEGORY_BASE["其他"])
    reasons: list[str] = [f"类别「{cat}」基础分"]

    score = base

    ok = sample.get("ok", True)
    if ok is False or (body.startswith("[正文提取失败") if body else False):
        score -= 35.0
        reasons.append("正文提取失败或标记失败，已降权")

    bc = int(sample.get("body_chars") or len(body))
    if bc < 60:
        score -= 18.0
        reasons.append("正文过短，信息可能不全")
    elif bc < 200:
        score -= 6.0
        reasons.append("正文较短")

    # 日期（与列表 date 字段）
    date_str = str(sample.get("date") or "")
    parsed = parse_email_date_for_filter(date_str)
    rb, rreason = _recency_bonus(parsed, now)
    score += rb
    if rreason:
        reasons.append(f"时效：{rreason}")

    urgency_add = 0.0
    matched_labels: list[str] = []
    for pat, add, label in _URGENCY_RULES:
        if pat.search(blob_head):
            urgency_add += add
            matched_labels.append(label)
    urgency_add = min(urgency_add, 28.0)
    score += urgency_add
    for lb in matched_labels[:5]:
        reasons.append(lb)
    if len(matched_labels) > 3:
        reasons.append(f"共 {len(matched_labels)} 条紧迫相关信号")

    dh = deadline_hint_from_text(subj + "\n" + body[:2000])
    is_actionable = bool(
        re.search(r"请|务必|需要您|请尽快|register|sign\s*up|submit|回复|填写", blob_head, re.I)
    )

    # 分数范围整理
    score = max(0.0, min(120.0, score))

    if score >= 72:
        level = "高"
    elif score >= 42:
        level = "中"
    else:
        level = "低"

    # 去重、截断
    seen_r = set()
    uniq_reasons = []
    for r in reasons:
        if r and r not in seen_r:
            seen_r.add(r)
            uniq_reasons.append(r)
        if len(uniq_reasons) >= 8:
            break

    return {
        "priority_score": round(score, 2),
        "priority_level": level,
        "priority_reasons": uniq_reasons,
        "deadline_hint": dh,
        "is_actionable": is_actionable,
    }


def apply_priority_to_samples(samples: list[dict], *, now: Optional[datetime] = None) -> None:
    """原地为每条 sample 写入 priority_* 等字段。"""
    now = now or datetime.now()
    for s in samples:
        extra = compute_priority_for_sample(s, now=now)
        s.update(extra)


def sort_indices_by_priority(samples: list[dict]) -> list[int]:
    """按 priority_score 降序的 index 列表。"""
    indexed = []
    for s in samples:
        idx = int(s.get("index") or 0)
        sc = float(s.get("priority_score") or 0)
        indexed.append((sc, idx))
    indexed.sort(key=lambda x: (-x[0], x[1]))
    return [i for _, i in indexed]


def dedupe_top_indices(
    samples: list[dict],
    *,
    top_n: int = 12,
) -> list[dict]:
    """
    按分数排序后，对同 convid、同规范化主题去重，保留分最高的一封。
    返回用于 LLM digest 的精简条目列表。
    """
    by_id = {int(s.get("index") or 0): s for s in samples}
    order = sort_indices_by_priority(samples)
    seen_conv: set[str] = set()
    seen_subj: set[str] = set()
    out: list[dict] = []
    for idx in order:
        s = by_id.get(idx)
        if not s:
            continue
        cv = str(s.get("convid") or "").strip()[:120]
        ns = _normalize_subject(str(s.get("subject") or ""))
        if cv and cv in seen_conv:
            continue
        if ns and ns in seen_subj:
            continue
        if cv:
            seen_conv.add(cv)
        if ns:
            seen_subj.add(ns)
        out.append(s)
        if len(out) >= top_n:
            break
    return out


def build_priority_digest_prompt(
    *,
    today: str,
    keyword: str,
    top_items: list[dict],
) -> str:
    """单次 LLM：仅根据 Top 条目的元数据与短摘要生成优先级说明。"""
    lines = []
    for i, s in enumerate(top_items, 1):
        subj = (s.get("subject") or "")[:200]
        snd = (s.get("sender") or "")[:120]
        dt = (s.get("date") or "")[:40]
        cat = (s.get("category") or "")[:40]
        pr = s.get("priority_reasons") or []
        pr_s = "；".join(pr[:4]) if isinstance(pr, list) else str(pr)
        preview = (s.get("body") or "")[:320]
        lines.append(
            f"{i}. [#{s.get('index')}] {subj}\n"
            f"   发件人：{snd}  日期：{dt}  分类：{cat}\n"
            f"   规则命中：{pr_s}\n"
            f"   正文摘录：{preview}\n"
        )
    block = "\n".join(lines)
    kw_line = f"用户搜索关键词：{keyword or '（空，表示最新邮件）'}\n"
    return f"""当前日期：{today}

{kw_line}
下面是按规则排序后、去重后的优先邮件条目（最多 12 封）。请严格基于下列文字作答，禁止编造不存在的邮件。

{block}

请用中文输出，且不要使用 Markdown。按下面结构输出（每部分用简短条目即可）：

【优先处理摘要】
用 2-4 句话说明：这批邮件里用户最值得先处理什么。

【紧急主题】
列出 2-4 个主题标签（每个不超过 12 字）。

【建议动作】
列出 3-6 条可执行建议（如「先读 #3」「处理验证码邮件」），序号对应上方 #index。

【Top 邮件一句话】
对上面每一封用一行：#序号 — 一句话说明为何重要或需注意什么。
"""


__all__ = [
    "apply_priority_to_samples",
    "sort_indices_by_priority",
    "dedupe_top_indices",
    "build_priority_digest_prompt",
    "deadline_hint_from_text",
]
