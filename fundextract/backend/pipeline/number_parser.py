"""
number_parser.py — Normalise raw string values extracted from PDF tables.

Handles:
  - Comma separators: "1,234,567" → 1234567
  - Parenthetical negatives: "(1,234)" → -1234
  - Dash placeholders: "—", "-", "–" → None
  - Unit scaling: if PDF uses "$000s" or "in thousands", multiply by 1000
  - Percent strings: "12.5%" → 0.125 (for Financial Highlights)
"""

import re
import logging
from typing import Any

logger = logging.getLogger(__name__)

_UNITS_PATTERNS = {
    "thousands": [r"in\s+thousands", r"\$000", r"\(000\)", r"thousands\s+of\s+dollars"],
    "millions":  [r"in\s+millions",  r"\$000,000", r"\(millions\)", r"millions\s+of\s+dollars"],
}

_STRIP_RE = re.compile(r"[,$%\s]")
_NEG_RE   = re.compile(r"^\((.+)\)$")
_DASH_RE  = re.compile(r"^[—\-–]+$")


def normalise_numbers(
    raw_tables: dict[str, list[dict[str, Any]]],
    hint: str = "auto",
) -> tuple[str, dict[str, list[dict[str, Any]]]]:
    """
    Detect units from raw table text and convert all cy_raw / py_raw strings
    to float | None.

    Returns (detected_units, normalised_tables).
    detected_units is one of "actual", "thousands", "millions".
    """
    # ── Detect units ─────────────────────────────────────────────────────────
    detected = hint if hint != "auto" else "actual"

    if hint == "auto":
        sample_text = " ".join(
            row.get("cy_raw", "") or ""
            for rows in raw_tables.values()
            for row in rows[:20]
        ).lower()
        for unit, patterns in _UNITS_PATTERNS.items():
            if any(re.search(p, sample_text) for p in patterns):
                detected = unit
                break
        logger.info("Unit detection: '%s'", detected)

    multiplier = {"actual": 1, "thousands": 1_000, "millions": 1_000_000}.get(detected, 1)

    # ── Normalise ─────────────────────────────────────────────────────────────
    normalised: dict[str, list[dict[str, Any]]] = {}
    for code, rows in raw_tables.items():
        norm_rows = []
        for row in rows:
            norm_rows.append({
                "label":  row["label"],
                "cy":     _parse(row.get("cy_raw"), multiplier),
                "py":     _parse(row.get("py_raw"), multiplier),
                "page":   row.get("page"),
            })
        normalised[code] = norm_rows

    return detected, normalised


def _parse(raw: str | None, multiplier: int) -> float | None:
    if raw is None:
        return None
    raw = raw.strip()
    if not raw or _DASH_RE.match(raw):
        return None
    # Strip percent sign — keep as decimal
    is_pct = raw.endswith("%")
    raw = _STRIP_RE.sub("", raw)
    # Handle parenthetical negatives
    m = _NEG_RE.match(raw)
    sign = -1 if m else 1
    if m:
        raw = m.group(1)
    try:
        val = float(raw) * sign
        if is_pct:
            return round(val / 100, 6)
        return val * multiplier
    except ValueError:
        return None
