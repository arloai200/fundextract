/**
 * taskpane.js — FundExtract Office.js task pane entry point.
 *
 * Manages the three-screen flow:
 *   UPLOAD → PROCESSING → RESULTS
 *
 * Calls the FastAPI backend, then delegates to excel-writer.js.
 */

import { buildTemplate } from "./template-builder.js";
import { writeExtractedData } from "./excel-writer.js";

// ── Config ───────────────────────────────────────────────────────────────────
const API_BASE = window.location.hostname === "localhost"
  ? "http://localhost:8000"
  : "https://fundextract.onrender.com";

const STEP_NAMES = ["toc", "sections", "soa", "soo", "socs", "socf", "fns", "write"];

// ── State ────────────────────────────────────────────────────────────────────
let selectedFile = null;
let extractionLog = [];

// ── Office initialisation ─────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    showError("FundExtract requires Microsoft Excel.");
    return;
  }
  initUI();
});

// ── UI setup ─────────────────────────────────────────────────────────────────
function initUI() {
  const dropZone    = document.getElementById("drop-zone");
  const fileInput   = document.getElementById("file-input");
  const browseLink  = document.getElementById("browse-link");
  const btnExtract  = document.getElementById("btn-extract");
  const btnAnother  = document.getElementById("btn-extract-another");
  const btnViewLog  = document.getElementById("btn-view-log");
  const btnCloseLog = document.getElementById("btn-close-log");

  // Drag-and-drop
  dropZone.addEventListener("dragover",  (e) => { e.preventDefault(); dropZone.classList.add("drag-over"); });
  dropZone.addEventListener("dragleave", ()  => dropZone.classList.remove("drag-over"));
  dropZone.addEventListener("drop",      (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    handleFileSelect(e.dataTransfer.files[0]);
  });
  dropZone.addEventListener("click",  () => fileInput.click());
  dropZone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") fileInput.click(); });
  browseLink.addEventListener("click", (e) => { e.stopPropagation(); fileInput.click(); });
  fileInput.addEventListener("change", () => handleFileSelect(fileInput.files[0]));

  // Extract button
  btnExtract.addEventListener("click", runExtraction);

  // Results screen
  btnAnother.addEventListener("click", resetToUpload);
  btnViewLog.addEventListener("click", () => {
    document.getElementById("log-content").textContent = extractionLog.join("\n");
    document.getElementById("modal-log").classList.remove("fe-hidden");
  });
  btnCloseLog.addEventListener("click", () => {
    document.getElementById("modal-log").classList.add("fe-hidden");
  });
}

// ── File selection ────────────────────────────────────────────────────────────
function handleFileSelect(file) {
  clearError();
  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".pdf")) {
    showError("Please select a PDF file.");
    return;
  }

  selectedFile = file;
  document.getElementById("file-name").textContent = `📄 ${file.name}`;
  document.getElementById("btn-extract").disabled = false;
}

// ── Main extraction flow ──────────────────────────────────────────────────────
async function runExtraction() {
  clearError();
  showScreen("processing");
  resetSteps();

  const extractPY = document.getElementById("opt-extract-py").checked;
  const units     = document.getElementById("opt-units").value;

  try {
    // Simulate step progression while the API call is in flight
    const progressPromise = animateSteps();

    // ── POST to backend ──────────────────────────────────────────────────────
    const formData = new FormData();
    formData.append("file", selectedFile);
    formData.append("extract_py", extractPY);
    formData.append("units", units);
    formData.append("mode", "extract");

    const response = await fetch(`${API_BASE}/api/v1/extract`, {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({ detail: response.statusText }));
      throw new Error(err.detail || `HTTP ${response.status}`);
    }

    const data = await response.json();
    extractionLog = data.extraction_log || [];

    // ── Write to Excel ───────────────────────────────────────────────────────
    markStep("write", "active");
    await Excel.run(async (context) => {
      const sheet = await buildTemplate(context);
      await writeExtractedData(context, sheet, data);
      await context.sync();
    });
    markStep("write", "done");
    setProgress(100, "Complete");

    await progressPromise;
    showResults(data);

  } catch (err) {
    console.error("FundExtract error:", err);
    showScreen("upload");
    showError(`Extraction failed: ${err.message}`);
  }
}

// ── Step animation helpers ────────────────────────────────────────────────────
function resetSteps() {
  STEP_NAMES.forEach((s) => {
    const el = document.querySelector(`[data-step="${s}"]`);
    if (el) el.className = "fe-step";
  });
  setProgress(0, "");
}

function markStep(stepId, state) {
  const el = document.querySelector(`[data-step="${stepId}"]`);
  if (!el) return;
  el.className = `fe-step ${state}`;
}

function setProgress(pct, label) {
  document.getElementById("progress-bar").style.width = `${pct}%`;
  document.getElementById("progress-label").textContent = label;
}

// Animate steps at roughly even intervals to match perceived backend progress.
// The actual result overwrites state when it arrives.
async function animateSteps() {
  const steps = STEP_NAMES.slice(0, -1); // all except "write" (handled separately)
  for (let i = 0; i < steps.length; i++) {
    if (i > 0) markStep(steps[i - 1], "done");
    markStep(steps[i], "active");
    const pct = Math.round(((i + 1) / (steps.length + 1)) * 90);
    setProgress(pct, `Step ${i + 1} of ${steps.length + 1}…`);
    await sleep(900);
  }
  markStep(steps[steps.length - 1], "done");
}

// ── Results screen ────────────────────────────────────────────────────────────
function showResults(data) {
  // Count totals across all sections
  let hi = 0, me = 0, lo = 0, total = 0;
  (data.sections || []).forEach((sec) => {
    (sec.line_items || []).forEach((item) => {
      total++;
      if (item.confidence === "high")   hi++;
      else if (item.confidence === "medium") me++;
      else lo++;
    });
  });

  document.getElementById("results-headline").textContent =
    `${total} line items extracted across ${(data.sections || []).length} statements`;

  document.getElementById("conf-high").textContent   = `${hi} High`;
  document.getElementById("conf-medium").textContent = `${me} Medium`;
  document.getElementById("conf-low").textContent    = `${lo} Low`;

  // Cross-checks
  const list = document.getElementById("cross-checks-list");
  list.innerHTML = "";
  (data.cross_checks || []).forEach((chk) => {
    const li = document.createElement("li");
    const pass = !chk.status || chk.status === "pass";
    li.innerHTML = `<span class="${pass ? "fe-check-pass" : "fe-check-review"}">${pass ? "✓" : "⚠"}</span>
                    <span>${chk.name}</span>`;
    list.appendChild(li);
  });

  document.getElementById("units-label").textContent =
    `Values in ${data.units || "auto-detected units"}`;

  showScreen("results");
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function showScreen(name) {
  ["upload", "processing", "results"].forEach((id) => {
    document.getElementById(`screen-${id}`).classList.toggle("fe-hidden", id !== name);
  });
}

function resetToUpload() {
  selectedFile = null;
  document.getElementById("file-input").value = "";
  document.getElementById("file-name").textContent = "";
  document.getElementById("btn-extract").disabled = true;
  extractionLog = [];
  clearError();
  showScreen("upload");
}

function showError(msg) {
  document.getElementById("upload-error").textContent = msg;
}

function clearError() {
  document.getElementById("upload-error").textContent = "";
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
