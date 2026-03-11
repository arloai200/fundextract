"""
FundExtract — FastAPI application entry point.
Serves the Office.js add-in static files and exposes the extraction API.
Shared backend is compatible with FundTrace (same pipeline modules).
"""

import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from dotenv import load_dotenv

from backend.api.extract import router as extract_router

load_dotenv()

# ── Allowed origins: Office app domains + localhost for dev ─────────────────
_raw_origins = os.getenv(
    "ALLOWED_ORIGINS",
    "https://appsforoffice.microsoft.com,https://excel.officeapps.live.com,https://localhost:3000",
)
ALLOWED_ORIGINS = [o.strip() for o in _raw_origins.split(",")]

app = FastAPI(
    title="FundExtract API",
    description="Automated mutual fund financial data extraction for Excel Add-In.",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── API routes ───────────────────────────────────────────────────────────────
app.include_router(extract_router, prefix="/api/v1")

# ── Static files: serve the addin/ folder ───────────────────────────────────
_addin_dir = os.path.join(os.path.dirname(__file__), "..", "addin")
if os.path.isdir(_addin_dir):
    app.mount("/", StaticFiles(directory=_addin_dir, html=True), name="addin")


@app.get("/health")
async def health():
    return {"status": "ok", "service": "FundExtract"}
