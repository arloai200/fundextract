"""
table_extractor.py — Extract tables from PDF pages using pdfplumber (primary)
and Camelot (fallback for borderless/lattice tables).

Returns a dict keyed by section code:
  {
    "SOA": [{"label": str, "cy": float|None, "py": float|None, "page": int}, ...],
    ...
  }
"""

import logging
from typing import Any

import pdfplumber

logger = logging.getLogger(__name__)


def extract_tables(
    pdf_path: str,
    sections_meta: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    """
    Extract raw (un-normalised) rows from the page ranges defined in sections_meta.
    Returns a mapping of code → list of raw row dicts.
    """
    results: dict[str, list[dict[str, Any]]] = {}

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)

        for sec in sections_meta:
            code = sec["code"]
            start = sec.get("start_page")
            end = sec.get("end_page")

            if not start:
                logger.debug("%s: no page range, skipping pdfplumber", code)
                continue

            # Convert to 0-indexed
            p_start = max(0, start - 1)
            p_end = min(total_pages - 1, (end or start + 5) - 1)

            rows: list[dict[str, Any]] = []
            for page_idx in range(p_start, p_end + 1):
                page = pdf.pages[page_idx]
                page_rows = _extract_page(page, page_idx + 1)
                rows.extend(page_rows)

            if rows:
                results[code] = rows
                logger.info("%s: extracted %d raw rows (pages %d–%d)",
                            code, len(rows), start, end or start)
            else:
                logger.warning("%s: pdfplumber found 0 rows on pages %d–%d — "
                               "will try Camelot or Vision", code, start, end or start)
                camelot_rows = _try_camelot(pdf_path, p_start + 1, p_end + 1, code)
                if camelot_rows:
                    results[code] = camelot_rows

    return results


# ── pdfplumber extraction ────────────────────────────────────────────────────

def _extract_page(page, page_num: int) -> list[dict[str, Any]]:
    """Extract row dicts from a single pdfplumber page object."""
    rows: list[dict[str, Any]] = []
    tables = page.extract_tables()

    if not tables:
        # Fall back to raw text lines
        text = page.extract_text() or ""
        for line in text.splitlines():
            row = _parse_text_line(line, page_num)
            if row:
                rows.append(row)
    else:
        for table in tables:
            for row in table:
                if not row:
                    continue
                label = (row[0] or "").strip()
                if not label:
                    continue
                cy = _safe_val(row[1] if len(row) > 1 else None)
                py = _safe_val(row[2] if len(row) > 2 else None)
                rows.append({"label": label, "cy_raw": cy, "py_raw": py, "page": page_num})

    return rows


def _parse_text_line(line: str, page_num: int) -> dict[str, Any] | None:
    """Parse a raw text line into a label + up to two numeric values."""
    import re
    # Match: label then 1–2 numbers (possibly in parentheses for negatives)
    m = re.match(
        r"^(.+?)\s+([\(\-]?[\d,]+\.?\d*\)?)\s+([\(\-]?[\d,]+\.?\d*\)?)?\s*$",
        line.strip(),
    )
    if not m:
        return None
    label = m.group(1).strip()
    if len(label) < 3:
        return None
    return {
        "label": label,
        "cy_raw": m.group(2),
        "py_raw": m.group(3) or None,
        "page": page_num,
    }


def _safe_val(v) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


# ── Camelot fallback ─────────────────────────────────────────────────────────

def _try_camelot(
    pdf_path: str,
    start_page: int,
    end_page: int,
    code: str,
) -> list[dict[str, Any]]:
    """Attempt table extraction using Camelot (lattice then stream mode)."""
    try:
        import camelot  # type: ignore
    except ImportError:
        logger.warning("Camelot not installed — skipping Camelot fallback for %s", code)
        return []

    pages_str = f"{start_page}-{end_page}" if start_page != end_page else str(start_page)
    rows: list[dict[str, Any]] = []

    for flavor in ("lattice", "stream"):
        try:
            tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor=flavor)
            for table in tables:
                df = table.df
                for _, row in df.iterrows():
                    label = str(row.iloc[0]).strip()
                    if not label or label.isdigit():
                        continue
                    cy = str(row.iloc[1]).strip() if len(row) > 1 else None
                    py = str(row.iloc[2]).strip() if len(row) > 2 else None
                    rows.append({
                        "label": label,
                        "cy_raw": cy or None,
                        "py_raw": py or None,
                        "page": table.parsing_report.get("page", start_page),
                    })
            if rows:
                logger.info("%s: Camelot (%s) extracted %d rows", code, flavor, len(rows))
                return rows
        except Exception as exc:
            logger.debug("Camelot %s failed for %s: %s", flavor, code, exc)

    return []
