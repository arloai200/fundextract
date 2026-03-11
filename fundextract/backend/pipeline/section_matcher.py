"""
section_matcher.py — Map TOC titles to canonical section codes.

Known section codes (matching the template):
  SOA   Statement of Assets and Liabilities
  SOO   Statement of Operations
  SOCS  Statement of Changes in Net Assets
  SOCF  Statement of Cash Flows
  FNS   Financial Highlights (per-share data)
  SOI   Schedule of Investments (not extracted to template, but logged)
"""

import logging
from typing import Any

from rapidfuzz import fuzz

logger = logging.getLogger(__name__)

# Canonical patterns for each section code
_SECTION_PATTERNS: dict[str, list[str]] = {
    "SOA": [
        "statement of assets and liabilities",
        "statement of assets & liabilities",
        "balance sheet",
        "assets and liabilities",
    ],
    "SOO": [
        "statement of operations",
        "statement of income",
        "income statement",
        "operations",
    ],
    "SOCS": [
        "statement of changes in net assets",
        "changes in net assets",
        "statement of net assets",
    ],
    "SOCF": [
        "statement of cash flows",
        "cash flow statement",
        "cash flows",
    ],
    "FNS": [
        "financial highlights",
        "per share data",
        "per-share operating performance",
        "selected per share data",
    ],
    "SOI": [
        "schedule of investments",
        "portfolio of investments",
        "investment schedule",
    ],
}

_SCORE_THRESHOLD = 65  # rapidfuzz partial_ratio threshold


def match_sections(toc_entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Given TOC entries [{title, page}, ...], return section meta dicts:
      [{"code": "SOA", "title": "...", "start_page": N, "end_page": N+1}, ...]
    end_page is inferred as the start of the next section.
    """
    matched: list[dict[str, Any]] = []
    seen_codes: set[str] = set()

    for entry in toc_entries:
        title_lower = entry["title"].lower()
        best_code: str | None = None
        best_score = 0

        for code, patterns in _SECTION_PATTERNS.items():
            if code in seen_codes:
                continue
            for pat in patterns:
                score = fuzz.partial_ratio(title_lower, pat)
                if score > best_score:
                    best_score = score
                    best_code = code

        if best_code and best_score >= _SCORE_THRESHOLD:
            matched.append({
                "code": best_code,
                "title": entry["title"],
                "start_page": entry["page"],
                "end_page": None,  # filled in below
                "match_score": best_score,
            })
            seen_codes.add(best_code)
            logger.debug("TOC '%s' → %s (score=%d)", entry["title"], best_code, best_score)

    # ── Infer end pages ──────────────────────────────────────────────────────
    for i, sec in enumerate(matched):
        if i + 1 < len(matched):
            sec["end_page"] = matched[i + 1]["start_page"] - 1
        else:
            sec["end_page"] = sec["start_page"] + 10  # generous last-section estimate

    if not matched:
        logger.warning("section_matcher: no sections matched from TOC — "
                       "returning default section stubs for vision fallback.")
        matched = _default_stubs()

    return matched


def _default_stubs() -> list[dict[str, Any]]:
    """Return placeholder stubs when no TOC is available."""
    return [
        {"code": c, "title": c, "start_page": None, "end_page": None, "match_score": 0}
        for c in ("SOA", "SOO", "SOCS", "SOCF", "FNS")
    ]
