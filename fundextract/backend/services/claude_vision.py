"""
claude_vision.py — Vision-based fallback extraction using Claude Sonnet.

Used when pdfplumber/Camelot cannot parse a section (e.g. scanned images,
complex multi-column layouts). Renders the relevant PDF pages as images and
sends them to the Claude API with a structured extraction prompt.
"""

import asyncio
import base64
import logging
import os
from typing import Any

import anthropic
import pdfplumber
from PIL import Image
import io

logger = logging.getLogger(__name__)

_CLIENT: anthropic.AsyncAnthropic | None = None

_EXTRACTION_PROMPT = """
You are an expert financial data extractor. The image shows a page from a
mutual fund annual report.

Extract every line item from the financial statement visible in this image.
Return ONLY a JSON array. Each element must be:
  {
    "label":  "<line item description>",
    "cy_raw": "<current year value as shown, or null>",
    "py_raw": "<prior year value as shown, or null>"
  }

Rules:
- Include section headers (e.g. "ASSETS", "LIABILITIES") with null values.
- Include total rows (e.g. "TOTAL ASSETS").
- Preserve parentheses for negatives exactly as printed.
- Do NOT include page numbers, footnotes, or column headers.
- Output valid JSON only — no explanation, no markdown.
"""


def _get_client() -> anthropic.AsyncAnthropic:
    global _CLIENT
    if _CLIENT is None:
        _CLIENT = anthropic.AsyncAnthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    return _CLIENT


async def extract_with_vision(
    pdf_path: str,
    sections: list[dict[str, Any]],
) -> dict[str, list[dict[str, Any]]]:
    """
    Render pages for each failed section and call Claude Vision.
    Returns a dict of code → raw row list (same schema as table_extractor output).
    """
    client = _get_client()
    tasks = [_extract_section(client, pdf_path, sec) for sec in sections]
    results_list = await asyncio.gather(*tasks, return_exceptions=True)

    output: dict[str, list[dict[str, Any]]] = {}
    for sec, result in zip(sections, results_list):
        if isinstance(result, Exception):
            logger.error("Vision failed for %s: %s", sec["code"], result)
        else:
            output[sec["code"]] = result
    return output


async def _extract_section(
    client: anthropic.AsyncAnthropic,
    pdf_path: str,
    sec: dict[str, Any],
) -> list[dict[str, Any]]:
    """Render section pages → base64 images → Claude API → parsed rows."""
    images = _render_pages(pdf_path, sec.get("start_page"), sec.get("end_page"))
    if not images:
        logger.warning("Vision: no images rendered for %s", sec["code"])
        return []

    all_rows: list[dict[str, Any]] = []
    for page_num, img_b64 in images:
        try:
            response = await client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4096,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": img_b64,
                            },
                        },
                        {"type": "text", "text": _EXTRACTION_PROMPT},
                    ],
                }],
            )
            import json
            raw_text = response.content[0].text.strip()
            # Strip possible markdown code fences
            if raw_text.startswith("```"):
                raw_text = raw_text.split("\n", 1)[1].rsplit("```", 1)[0]
            rows = json.loads(raw_text)
            for r in rows:
                r["page"] = page_num
            all_rows.extend(rows)
            logger.info("Vision: %d rows extracted from page %d for %s",
                        len(rows), page_num, sec["code"])
        except Exception as exc:
            logger.error("Vision API error on page %d: %s", page_num, exc)

    return all_rows


def _render_pages(
    pdf_path: str,
    start_page: int | None,
    end_page: int | None,
) -> list[tuple[int, str]]:
    """Render PDF page range to base64 PNG images."""
    result: list[tuple[int, str]] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
            p_start = max(0, (start_page or 1) - 1)
            p_end   = min(total - 1, (end_page or p_start + 5) - 1)

            for i in range(p_start, p_end + 1):
                page = pdf.pages[i]
                img: Image.Image = page.to_image(resolution=150).original
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                b64 = base64.standard_b64encode(buf.getvalue()).decode()
                result.append((i + 1, b64))
    except Exception as exc:
        logger.error("PDF render error: %s", exc)
    return result
