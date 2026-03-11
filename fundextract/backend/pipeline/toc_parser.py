"""
toc_parser.py — Extract the table of contents from a mutual fund annual report PDF.

Strategy:
  1. Search the first 20 pages for a page that looks like a TOC (many lines
     containing a number at the end, preceded by dots or whitespace).
  2. Parse each TOC line into (title, page_number).
  3. Return as a list of dicts for section_matcher to consume.
"""

import re
import logging
from typing import Any

import pdfplumber

logger = logging.getLogger(__name__)

# Heuristic: a TOC line has at least one word followed eventually by a number.
_TOC_LINE_RE = re.compile(r"^(.+?)\s*\.{2,}?\s*(\d{1,4})\s*$|^(.+?)\s{3,}(\d{1,4})\s*$")
_MIN_TOC_LINES = 5  # page must have this many matching lines to qualify


def parse_toc(pdf_path: str) -> list[dict[str, Any]]:
    """
    Return a list of TOC entries:
      [{"title": str, "page": int}, ...]
    Page numbers are as printed in the PDF (not 0-indexed).
    """
    entries: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        scan_pages = min(25, len(pdf.pages))

        for i in range(scan_pages):
            page = pdf.pages[i]
            text = page.extract_text() or ""
            lines = text.splitlines()

            hits = []
            for line in lines:
                m = _TOC_LINE_RE.match(line.strip())
                if m:
                    title = (m.group(1) or m.group(3) or "").strip()
                    page_num = int(m.group(2) or m.group(4) or 0)
                    if title and page_num:
                        hits.append({"title": title, "page": page_num})

            if len(hits) >= _MIN_TOC_LINES:
                logger.info("TOC found on PDF page %d (%d entries)", i + 1, len(hits))
                entries = hits
                break  # take the first qualifying page

    if not entries:
        logger.warning("No TOC detected — will attempt to scan all pages for statement headers.")

    return entries
