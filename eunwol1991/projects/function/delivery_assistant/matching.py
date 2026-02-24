import re
import unicodedata
from datetime import date
from difflib import SequenceMatcher

try:
    from rapidfuzz import fuzz

    _HAS_RAPIDFUZZ = True
except Exception:
    fuzz = None
    _HAS_RAPIDFUZZ = False


def normalize_text(value: str) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold().strip()
    text = re.sub(r"[^\w\s]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def confidence_level(top_score: float, second_score: float) -> str:
    gap = top_score - second_score
    if top_score >= 92 and gap >= 8:
        return "high"
    if top_score >= 82 and gap >= 5:
        return "medium"
    if top_score >= 75:
        return "low"
    return "none"


def _field_score(query: str, value: str) -> float:
    nq = normalize_text(query)
    nv = normalize_text(value)
    if not nq or not nv:
        return 0.0
    if _HAS_RAPIDFUZZ:
        return float(fuzz.WRatio(nq, nv))
    return SequenceMatcher(None, nq, nv).ratio() * 100.0


def _recency_key(record: dict) -> tuple[date, int]:
    rec_date = record.get("record_date") or date.min
    row_idx = int(record.get("row_idx") or 0)
    return rec_date, row_idx


def rank_candidates(query: dict, records: list[dict], limit: int = 5) -> list[dict]:
    scored: list[dict] = []
    for rec in records:
        customer = _field_score(query.get("customer", ""), rec.get("customer", ""))
        outlet = _field_score(query.get("outlet", ""), rec.get("outlet", ""))
        desc = _field_score(query.get("description", ""), rec.get("description", ""))
        code = _field_score(query.get("product_code", ""), rec.get("product_code", ""))
        base_score = (customer * 0.4) + (outlet * 0.25) + (desc * 0.15) + (code * 0.2)
        rec_date, row_idx = _recency_key(rec)
        recency_bonus = 1.0 if rec_date != date.min else 0.0
        final = base_score + recency_bonus
        scored.append(
            {
                "record": rec,
                "score": final,
                "breakdown": {
                    "customer": customer,
                    "outlet": outlet,
                    "description": desc,
                    "product_code": code,
                },
                "recency": {"date": rec_date, "row_idx": row_idx},
            }
        )
    scored.sort(
        key=lambda x: (x["score"], x["recency"]["date"], x["recency"]["row_idx"]),
        reverse=True,
    )
    return scored[:limit]
