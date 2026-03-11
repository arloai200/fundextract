"""
test_line_matching.py — Unit tests for the fuzzy line-item matcher.
Run with: pytest tests/test_line_matching.py -v
"""

import pytest
from backend.pipeline.line_matcher import match_lines


SAMPLE_INPUT = {
    "SOA": [
        {"label": "Investments at fair value", "cy": 1_234_567, "py": 1_100_234, "page": 12},
        {"label": "Cash & cash equivalents",   "cy":    45_678, "py":    39_000, "page": 12},
        {"label": "TOTAL ASSETS",              "cy": 1_280_245, "py": 1_139_234, "page": 12},
        {"label": "Management fees payable",   "cy":     1_234, "py":     1_100, "page": 12},
        {"label": "TOTAL LIABILITIES",         "cy":    22_345, "py":    19_876, "page": 12},
        {"label": "Paid-in capital",           "cy": 1_100_000, "py":   980_000, "page": 12},
        {"label": "TOTAL NET ASSETS",          "cy": 1_257_900, "py": 1_119_358, "page": 12},
    ],
}


def test_match_lines_returns_sections():
    result = match_lines(SAMPLE_INPUT)
    assert "sections" in result
    assert len(result["sections"]) == 1
    assert result["sections"][0]["code"] == "SOA"


def test_match_lines_high_confidence_for_exact():
    result = match_lines(SAMPLE_INPUT)
    soa = result["sections"][0]
    labels = {item["label"]: item["confidence"] for item in soa["line_items"]}
    assert labels.get("TOTAL ASSETS") == "high"
    assert labels.get("Investments at fair value") == "high"


def test_match_lines_fuzzy_cash():
    """'Cash & cash equivalents' should fuzzy-match to 'Cash and cash equivalents'."""
    result = match_lines(SAMPLE_INPUT)
    soa = result["sections"][0]
    matched = [i for i in soa["line_items"] if "cash" in i["label"].lower()]
    assert len(matched) > 0


def test_match_lines_cy_values_preserved():
    result = match_lines(SAMPLE_INPUT)
    soa = result["sections"][0]
    for item in soa["line_items"]:
        if item["label"] == "TOTAL ASSETS":
            assert item["cy_value"] == 1_280_245


def test_match_lines_no_py():
    result = match_lines(SAMPLE_INPUT, extract_py=False)
    soa = result["sections"][0]
    for item in soa["line_items"]:
        assert item["py_value"] is None
