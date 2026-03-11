/**
 * excel-writer.js — Receives structured JSON from the backend and writes
 * all extracted line items into the FundExtract sheet using batch operations.
 *
 * Writing strategy:
 *   • One range.values = [...] call per section (fast — avoids cell-by-cell).
 *   • Formulas written separately in a second pass (Δ$ and Δ%).
 *   • Formatting applied last (section headers, totals, confidence colours).
 *   • Cross-check rows appended at the bottom.
 *   • Named ranges created for key totals.
 */

// ── Formatting constants (must match template-builder.js) ────────────────────
const NAVY   = "#1b3a6b";
const BLUE   = "#deecf9";
const GREEN  = "#dff6dd";
const RED    = "#fde7e9";
const ORANGE = "#fed9cc";
const BLUE_FONT = "#0078d4";

// Data starts on row 6 (rows 1–5 are header block).
const DATA_START_ROW = 6;

/**
 * Write all extracted data into the sheet.
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet}     sheet
 * @param {object}              data   — backend JSON response
 */
export async function writeExtractedData(context, sheet, data) {
  let currentRow = DATA_START_ROW;

  // Write fund name + fiscal year end into header cells
  if (data.fund_name) {
    sheet.getRange("B2").values = [[data.fund_name]];
  }
  if (data.fiscal_year_end) {
    sheet.getRange("E2").values = [[data.fiscal_year_end]];
  }
  sheet.getRange("E3").values = [[new Date().toLocaleDateString()]];
  sheet.getRange("B3").values = [[data.units || "auto"]];

  // Section totals rows tracking (for named ranges)
  const namedRanges = {};

  for (const section of (data.sections || [])) {
    const code = section.code;

    // ── Section header row ─────────────────────────────────────────────────
    const headerLabel = section.title || code;
    const headerRowRange = sheet.getRange(`A${currentRow}:I${currentRow}`);
    headerRowRange.values = [[code, headerLabel, "", "", "", "", "", "", ""]];
    headerRowRange.format.fill.color = NAVY;
    headerRowRange.format.font.color = "#ffffff";
    headerRowRange.format.font.bold  = true;
    currentRow++;

    const sectionDataStart = currentRow;

    // ── Batch-write all line items ─────────────────────────────────────────
    const items = section.line_items || [];
    if (items.length === 0) continue;

    const valueGrid  = [];
    const formulaGrid = [];

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const absRow = currentRow + i;
      const cy = item.cy_value ?? "";
      const py = item.py_value ?? "";

      // Values array (A–I): section, label, CY, PY, blank E, blank F, confidence, page, notes
      valueGrid.push([
        code,
        item.label,
        cy !== "" ? cy : "",
        py !== "" ? py : "",
        "",  // E — formula inserted below
        "",  // F — formula inserted below
        item.confidence || "",
        item.source_page || "",
        "",  // I — user workspace
      ]);

      // Formula array — only E and F matter; rest can be empty
      const hasValues = cy !== "" && py !== "";
      formulaGrid.push([
        "", // A
        "", // B
        "", // C
        "", // D
        hasValues ? `=C${absRow}-D${absRow}` : "",
        hasValues ? `=IFERROR((C${absRow}-D${absRow})/D${absRow},"-")` : "",
        "", // G
        "", // H
        "", // I
      ]);
    }

    // Batch write values
    const dataRange = sheet.getRange(`A${currentRow}:I${currentRow + items.length - 1}`);
    dataRange.values = valueGrid;

    // Batch write formulas
    const formulaRange = sheet.getRange(`A${currentRow}:I${currentRow + items.length - 1}`);
    formulaRange.formulas = formulaGrid;

    // ── Formatting pass ───────────────────────────────────────────────────
    for (let i = 0; i < items.length; i++) {
      const item  = items[i];
      const absRow = currentRow + i;
      const rowRange = sheet.getRange(`A${absRow}:I${absRow}`);

      if (item.is_total) {
        rowRange.format.fill.color = BLUE;
        rowRange.format.font.bold  = true;
        // Track named range candidates
        _maybeRegisterNamedRange(namedRanges, code, item.label, absRow);
      }

      if (item.unmatched) {
        sheet.getRange(`G${absRow}`).values = [["UNMATCHED"]];
        sheet.getRange(`G${absRow}`).format.fill.color = ORANGE;
      }

      // Colour CY/PY cells blue to signal extracted (non-formula) data
      const cy = item.cy_value ?? "";
      const py = item.py_value ?? "";
      if (cy !== "") sheet.getRange(`C${absRow}`).format.font.color = BLUE_FONT;
      if (py !== "") sheet.getRange(`D${absRow}`).format.font.color = BLUE_FONT;

      // Low confidence → orange background on G
      if (item.confidence === "low" && !item.is_total) {
        sheet.getRange(`G${absRow}`).format.fill.color = ORANGE;
      }
    }

    // Number format for CY/PY/Δ$ columns
    const valueColRange = sheet.getRange(`C${sectionDataStart}:E${currentRow + items.length - 1}`);
    valueColRange.numberFormat = [["#,##0;(#,##0);-"]];

    const pctRange = sheet.getRange(`F${sectionDataStart}:F${currentRow + items.length - 1}`);
    pctRange.numberFormat = [["0.0%;(0.0%);-"]];

    currentRow += items.length;

    // Blank spacer row between sections
    currentRow++;
  }

  // ── Cross-check rows ───────────────────────────────────────────────────────
  const checkHeaderRange = sheet.getRange(`A${currentRow}:I${currentRow}`);
  checkHeaderRange.values = [["CHK", "CROSS-STATEMENT FOOTING & TRACE CHECKS", "", "", "", "", "", "", ""]];
  checkHeaderRange.format.fill.color = NAVY;
  checkHeaderRange.format.font.color = "#ffffff";
  checkHeaderRange.format.font.bold  = true;
  currentRow++;

  const crossChecks = _buildCrossCheckFormulas(namedRanges);
  for (const chk of crossChecks) {
    const rowRange = sheet.getRange(`A${currentRow}:I${currentRow}`);
    rowRange.values = [["CHECK", chk.label, "", "", "", "", "", "", ""]];

    // Formula for CY difference in C, result (PASS/REVIEW) in E
    if (chk.cyFormula) {
      sheet.getRange(`C${currentRow}`).formulas = [[chk.cyFormula]];
      sheet.getRange(`D${currentRow}`).formulas = [[chk.dyFormula || ""]];
    }
    sheet.getRange(`E${currentRow}`).formulas = [[
      `=IF(AND(ISNUMBER(C${currentRow}),C${currentRow}=0),"✓ PASS","⚠ REVIEW")`
    ]];

    // Conditional: green if PASS, orange if REVIEW
    const passRange  = sheet.getRange(`E${currentRow}`);
    passRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText).textComparison.format.fill.color = GREEN;
    passRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText).textComparison.format.fill.color = ORANGE;

    currentRow++;
  }

  // ── Named ranges ──────────────────────────────────────────────────────────
  const workbook = context.workbook;
  for (const [name, row] of Object.entries(namedRanges)) {
    try {
      workbook.names.add(name, sheet.getRange(`C${row}`));
    } catch (_) {
      // Named range may already exist — ignore
    }
  }

  await context.sync();
}


// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Register key totals as named ranges (Excel names cannot contain spaces).
 */
const _NAMED_RANGE_MAP = {
  "SOA":  { "TOTAL ASSETS": "SOA_TotalAssets", "TOTAL LIABILITIES": "SOA_TotalLiabilities", "TOTAL NET ASSETS": "SOA_TotalNetAssets" },
  "SOO":  { "NET INCREASE (DECREASE) IN NET ASSETS FROM OPERATIONS": "SOO_NetOps" },
  "SOCS": { "NET ASSETS — END OF YEAR": "SOCS_EndingNetAssets", "NET INCREASE (DECREASE) FROM OPERATIONS": "SOCS_OpsTotal" },
  "SOCF": { "CASH AND EQUIVALENTS — END OF YEAR": "SOCF_EndingCash" },
};

function _maybeRegisterNamedRange(registry, code, label, row) {
  const sectionMap = _NAMED_RANGE_MAP[code];
  if (!sectionMap) return;
  for (const [fragment, name] of Object.entries(sectionMap)) {
    if (label.toUpperCase().includes(fragment.toUpperCase())) {
      registry[name] = row;
    }
  }
}

/**
 * Build cross-check formula objects using named range row numbers.
 */
function _buildCrossCheckFormulas(namedRanges) {
  const r = namedRanges;

  const checks = [
    {
      label:     "SOA: Total Assets = Total Liab + Net Assets (CY)",
      cyFormula: r.SOA_TotalAssets && r.SOA_TotalLiabilities && r.SOA_TotalNetAssets
        ? `=C${r.SOA_TotalAssets}-C${r.SOA_TotalLiabilities}-C${r.SOA_TotalNetAssets}`
        : "",
      dyFormula: r.SOA_TotalAssets && r.SOA_TotalLiabilities && r.SOA_TotalNetAssets
        ? `=D${r.SOA_TotalAssets}-D${r.SOA_TotalLiabilities}-D${r.SOA_TotalNetAssets}`
        : "",
    },
    {
      label:     "SOCS: Ending Net Assets ties to SOA Total Net Assets (CY)",
      cyFormula: r.SOCS_EndingNetAssets && r.SOA_TotalNetAssets
        ? `=C${r.SOCS_EndingNetAssets}-C${r.SOA_TotalNetAssets}`
        : "",
      dyFormula: r.SOCS_EndingNetAssets && r.SOA_TotalNetAssets
        ? `=D${r.SOCS_EndingNetAssets}-D${r.SOA_TotalNetAssets}`
        : "",
    },
    {
      label:     "SOO: Net Ops ties to SOCS Operations total (CY)",
      cyFormula: r.SOO_NetOps && r.SOCS_OpsTotal
        ? `=C${r.SOO_NetOps}-C${r.SOCS_OpsTotal}`
        : "",
      dyFormula: r.SOO_NetOps && r.SOCS_OpsTotal
        ? `=D${r.SOO_NetOps}-D${r.SOCS_OpsTotal}`
        : "",
    },
    {
      label:     "SOCF: Ending Cash ties to SOA Cash line (CY)",
      cyFormula: r.SOCF_EndingCash
        ? `=C${r.SOCF_EndingCash}`   // compared against SOA cash row (user verifies)
        : "",
      dyFormula: "",
    },
  ];

  return checks;
}
