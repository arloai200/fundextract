/**
 * template-builder.js — Creates (or resets) the "FundExtract" sheet
 * in the active workbook with the full unified template structure.
 *
 * Column layout (A–I):
 *   A  Section code
 *   B  Line Item
 *   C  Current Year ($)
 *   D  Prior Year ($)
 *   E  Change ($)   [formula]
 *   F  Change (%)   [formula]
 *   G  Confidence
 *   H  PDF Page #
 *   I  Notes / Workpaper Ref (user workspace starts here)
 */

const SHEET_NAME = "FundExtract";

// ── Formatting constants ──────────────────────────────────────────────────────
const NAVY   = "#1b3a6b";
const BLUE   = "#deecf9";
const GREEN  = "#dff6dd";
const RED    = "#fde7e9";
const ORANGE = "#fed9cc";

/**
 * Build or reset the FundExtract sheet.
 * Returns the worksheet object (already synced for column widths).
 */
export async function buildTemplate(context) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  let sheet = null;
  for (const s of sheets.items) {
    if (s.name === SHEET_NAME) { sheet = s; break; }
  }

  if (!sheet) {
    sheet = sheets.add(SHEET_NAME);
  } else {
    // Ask user before overwriting — handled in taskpane.js;
    // here we just clear contents (preserve sheet for speed).
    sheet.getUsedRange().clear();
  }

  // ── Column widths ──────────────────────────────────────────────────────────
  const colWidths = [
    { col: "A", width: 7  },   // Section code (narrow)
    { col: "B", width: 46 },   // Line Item
    { col: "C", width: 18 },   // CY
    { col: "D", width: 18 },   // PY
    { col: "E", width: 16 },   // Δ$
    { col: "F", width: 10 },   // Δ%
    { col: "G", width: 10 },   // Confidence
    { col: "H", width: 8  },   // Page
    { col: "I", width: 30 },   // Notes
  ];

  colWidths.forEach(({ col, width }) => {
    sheet.getRange(`${col}:${col}`).columnWidth = width;
  });

  // ── Row 1 — Title ──────────────────────────────────────────────────────────
  _writeRow(sheet, 1, "A", [
    ["FundExtract — Unified Financial Statement Data", "", "", "", "", "", "", "", ""],
  ]);
  const titleRange = sheet.getRange("A1:I1");
  titleRange.merge();
  titleRange.format.fill.color = NAVY;
  titleRange.format.font.color = "#ffffff";
  titleRange.format.font.size  = 14;
  titleRange.format.font.bold  = true;

  // ── Rows 2–3 — Header metadata fields ─────────────────────────────────────
  _writeRow(sheet, 2, "A", [["Fund:", "", "", "FYE:", "", "", "", "", ""]]);
  _writeRow(sheet, 3, "A", [["Extracted:", "", "", "Units:", "", "", "", "", ""]]);

  // ── Row 4 — blank spacer ──────────────────────────────────────────────────
  // (left empty intentionally)

  // ── Row 5 — Column headers ─────────────────────────────────────────────────
  const headers = [
    "Section", "Line Item", "Current Year ($)", "Prior Year ($)",
    "Change ($)", "Change (%)", "Confidence", "Pg #", "Notes / Workpaper Ref",
  ];
  _writeRow(sheet, 5, "A", [headers]);
  const headerRange = sheet.getRange("A5:I5");
  headerRange.format.fill.color = NAVY;
  headerRange.format.font.color = "#ffffff";
  headerRange.format.font.bold  = true;
  headerRange.format.horizontalAlignment = "Center";

  // ── Freeze panes: top 5 rows + columns A–B ────────────────────────────────
  sheet.freezePanes.freezeAt(sheet.getRange("C6"));

  // ── Named ranges (key totals — row numbers are fixed) ─────────────────────
  // These match the exact row layout from the xlsx template.
  // Rows 6–52: SOA; rows 53–95: SOO; etc.
  // Named ranges point to CY column (C) — add after data is written.

  // ── Print area ────────────────────────────────────────────────────────────
  sheet.pageLayout.setPrintArea(sheet.getRange("A1:I200"));

  await context.sync();
  return sheet;
}


// ── Internal helper ────────────────────────────────────────────────────────────
function _writeRow(sheet, startRow, startCol, values) {
  const range = sheet.getRange(`${startCol}${startRow}`);
  range.getResizedRange(values.length - 1, values[0].length - 1).values = values;
}
