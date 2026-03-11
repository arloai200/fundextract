# FundExtract

**Excel Add-In for Automated Mutual Fund Financial Data Extraction**

FundExtract is an Office.js task pane add-in that lets auditors upload a mutual fund annual report PDF and have every financial statement automatically extracted into a single unified worksheet — directly inside Excel.

## Architecture

| Component | Technology |
|-----------|-----------|
| Excel Add-In (frontend) | Office.js, HTML/CSS/JS |
| Backend API | Python FastAPI (hosted on Render) |
| PDF Extraction | pdfplumber + Camelot |
| AI Fallback | Claude Sonnet Vision API |
| Fuzzy Matching | rapidfuzz |

## Quick Start (Development)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set your Anthropic API key
export ANTHROPIC_API_KEY=sk-ant-...

# 3. Start the backend
uvicorn backend.main:app --reload --port 8000

# 4. Serve the add-in frontend (separate terminal)
cd addin
npx office-addin-dev-certs install   # installs localhost HTTPS certs
npx http-server -S -C ~/.office-addin-dev-certs/localhost.crt \
                   -K ~/.office-addin-dev-certs/localhost.key \
                   -p 3000

# 5. Sideload in Excel
#    Insert → My Add-ins → Upload My Add-in → select manifest-dev.xml
```

## Deployment (Render)

1. Push this repo to GitHub.
2. On [render.com](https://render.com): New → Web Service → connect repo.
3. Render auto-detects `render.yaml` — just add the `ANTHROPIC_API_KEY` secret.
4. Update `manifest.xml` with your Render URL.
5. Distribute `manifest.xml` via Microsoft 365 Admin Center → Integrated Apps.

## API

```
POST /api/v1/extract
  file        (PDF, multipart)
  extract_py  (bool, default true)
  units       (auto|actual|thousands|millions)
  mode        (extract|trace)
```

## Template Structure

Single sheet "FundExtract" with 5 stacked statement sections:

- **SOA** — Statement of Assets & Liabilities (rows 6–52)
- **SOO** — Statement of Operations (rows 53–95)
- **SOCS** — Statement of Changes in Net Assets (rows 96–120)
- **SOCF** — Statement of Cash Flows (rows 121–150)
- **FNS** — Financial Highlights (rows 151–170)
- **CHK** — Cross-statement footing checks

Columns A–H are extracted data (locked structure). Column I+ is free workspace.

## Companion Tool

FundTrace — browser-based PDF annotation tool for footing and tracing. Shares the same FastAPI backend.
