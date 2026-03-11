"""
test_extract_endpoint.py — Integration tests for POST /api/v1/extract.
Run with: pytest tests/test_extract_endpoint.py -v
Requires: a running backend (uvicorn backend.main:app) or use TestClient.
"""

import os
import pytest
from fastapi.testclient import TestClient

# Adjust import path as needed when running from repo root.
from backend.main import app

client = TestClient(app)

SAMPLE_PDF = os.path.join(os.path.dirname(__file__), "sample_reports", "sample_fund.pdf")


def test_health():
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["status"] == "ok"


@pytest.mark.skipif(not os.path.exists(SAMPLE_PDF), reason="Sample PDF not present")
def test_extract_returns_sections():
    with open(SAMPLE_PDF, "rb") as f:
        response = client.post(
            "/api/v1/extract",
            data={"extract_py": "true", "units": "auto", "mode": "extract"},
            files={"file": ("sample_fund.pdf", f, "application/pdf")},
        )
    assert response.status_code == 200
    data = response.json()
    assert "sections" in data
    assert len(data["sections"]) > 0
    # Every section must have a code and line_items
    for sec in data["sections"]:
        assert "code" in sec
        assert "line_items" in sec


def test_extract_rejects_non_pdf():
    response = client.post(
        "/api/v1/extract",
        data={"extract_py": "true", "units": "auto", "mode": "extract"},
        files={"file": ("report.xlsx", b"fake content", "application/octet-stream")},
    )
    assert response.status_code == 400


@pytest.mark.skipif(not os.path.exists(SAMPLE_PDF), reason="Sample PDF not present")
def test_extract_confidence_scores():
    with open(SAMPLE_PDF, "rb") as f:
        response = client.post(
            "/api/v1/extract",
            data={"extract_py": "true", "units": "auto", "mode": "extract"},
            files={"file": ("sample_fund.pdf", f, "application/pdf")},
        )
    data = response.json()
    for sec in data["sections"]:
        for item in sec["line_items"]:
            assert item["confidence"] in ("high", "medium", "low")
