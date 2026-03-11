"""
POST /api/v1/extract — Core extraction endpoint.
Accepts a PDF upload, runs the pipeline, returns structured JSON.
"""

import os
import tempfile
import logging
from typing import Literal

from fastapi import APIRouter, File, Form, UploadFile, HTTPException
from fastapi.responses import JSONResponse

from backend.pipeline.toc_parser import parse_toc
from backend.pipeline.section_matcher import match_sections
from backend.pipeline.table_extractor import extract_tables
from backend.pipeline.number_parser import normalise_numbers
from backend.pipeline.line_matcher import match_lines
from backend.services.cleanup import delete_temp_file
from backend.services.claude_vision import extract_with_vision

logger = logging.getLogger(__name__)

router = APIRouter()


@router.post("/extract")
async def extract(
    file: UploadFile = File(..., description="Annual report PDF"),
    extract_py: bool = Form(True, description="Also extract prior-year column"),
    units: Literal["actual", "thousands", "millions", "auto"] = Form("auto"),
    mode: Literal["extract", "trace"] = Form("extract"),
):
    """
    Main extraction endpoint shared by FundExtract (mode=extract) and
    FundTrace (mode=trace).  Returns structured JSON with all financial
    statement line items, confidence scores, source page numbers, and
    cross-statement footing checks.
    """
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF files are accepted.")

    # ── Write upload to /tmp (RAM-backed on Render) ─────────────────────────
    tmp_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False, dir="/tmp") as tmp:
            tmp.write(await file.read())
            tmp_path = tmp.name

        logger.info("Received PDF: %s → %s (%d bytes)", file.filename, tmp_path,
                    os.path.getsize(tmp_path))

        # ── Step 1: Parse table-of-contents to identify statement pages ──────
        toc = parse_toc(tmp_path)

        # ── Step 2: Match TOC entries to known section codes ─────────────────
        sections_meta = match_sections(toc)

        # ── Step 3: Extract tables from identified page ranges ───────────────
        raw_tables = extract_tables(tmp_path, sections_meta)

        # ── Step 4: If pdfplumber/Camelot failed any section, fall back to
        #            Claude Vision (only if ANTHROPIC_API_KEY is configured) ────
        failed = [s for s in sections_meta if s["code"] not in raw_tables]
        if failed:
            if os.getenv("ANTHROPIC_API_KEY"):
                logger.info("Vision fallback for %d section(s): %s",
                            len(failed), [s["code"] for s in failed])
                vision_results = await extract_with_vision(tmp_path, failed)
                raw_tables.update(vision_results)
            else:
                logger.info(
                    "Vision fallback skipped (ANTHROPIC_API_KEY not set) "
                    "for section(s): %s", [s["code"] for s in failed]
                )

        # ── Step 5: Normalise numbers (detect units, strip commas/parens) ────
        detected_units, normalised = normalise_numbers(raw_tables, hint=units)

        # ── Step 6: Fuzzy-match line labels to template rows ─────────────────
        result = match_lines(normalised, extract_py=extract_py)

        # ── Step 7: Build response payload ───────────────────────────────────
        payload = {
            "fund_name": result.get("fund_name", ""),
            "fiscal_year_end": result.get("fiscal_year_end", ""),
            "units": detected_units,
            "extraction_mode": mode,
            "sections": result["sections"],
            "cross_checks": result.get("cross_checks", []),
            "extraction_log": result.get("log", []),
        }
        return JSONResponse(content=payload)

    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Extraction failed for %s", file.filename)
        raise HTTPException(status_code=500, detail=f"Extraction error: {exc}") from exc
    finally:
        # ── Always delete the temp PDF immediately ────────────────────────────
        if tmp_path:
            delete_temp_file(tmp_path)
