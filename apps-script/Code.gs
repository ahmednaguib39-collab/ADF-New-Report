// ============================================================
// MV Daily Shift Report – Main Controller
// ALFANAR PROJECTS | O&M Medium Voltage Distribution Network
// Author: Ahmed Naguib | Version: 1.0
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
// Replace these IDs with your actual Google Sheet / Doc IDs
const CONFIG = {
  SHEET_ID:         "/1mO_IgsMeG_pCjEZShfzyEvkdaa20hpOkMNYFbG0_QoI/",   // The ADF_New_Report sheet
  TEMPLATE_DOC_ID:  "/1wBMvCMH6ssavVvQuftYfkcpFCY-RhQ3atZ-VI-UMeA4/", // MV_Daily_Shift_Report_Template
  OUTPUT_FOLDER_ID: "/152c62QHVZg_i3qqYfPRyTheLg-xDROCR",         // Where generated reports are saved
  TIMEZONE:         "Asia/Riyadh",

  // Sheet tab names (must match exactly)
  SHEETS: {
    SHIFT:      "Shift_Reports",
    ATTENDANCE: "Shift_Attendance",
    FAULTS:     "Fault_Logs",
    ALARMS:     "Major_Alarms",
    READINGS:   "System_Readings",
    USERS:      "Users_Directory",
  },

  // Status column in Shift_Reports that prevents re-generation
  STATUS_COL: "Report_Status",  // values: "" | "Generated" | "Approved"
};
// ────────────────────────────────────────────────────────────


/**
 * Entry point called by the daily trigger OR manually from the
 * "MV Reports" menu.  Finds every Shift_Report row whose
 * Report_Status is blank (pending) and generates a Google Doc
 * for each one.
 */
function generatePendingReports() {
  const ss       = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const reports  = getSheetData(ss, CONFIG.SHEETS.SHIFT);

  let generated = 0;

  reports.forEach(row => {
    if (row["Report_Status"] && row["Report_Status"].toString().trim() !== "") return; // already done

    try {
      const docUrl = buildReport(ss, row);
      markReportGenerated(ss, row["Report_ID"], docUrl);
      generated++;
      Logger.log(`✅  Generated report for ${row["Report_ID"]} → ${docUrl}`);
    } catch (e) {
      Logger.log(`❌  Failed for ${row["Report_ID"]}: ${e.message}`);
    }
  });

  showToast(`${generated} report(s) generated.`);
}


/**
 * Build one Google Doc from the template, filling every
 * placeholder with live data from the Sheet.
 *
 * @param {Spreadsheet} ss   - The open spreadsheet object
 * @param {Object}      row  - One row from Shift_Reports as a plain object
 * @returns {string}         - URL of the newly created Doc
 */
function buildReport(ss, row) {
  const reportId = row["Report_ID"];

  // ── 1. Fetch related sub-tables ──────────────────────────
  const allAttendance = getSheetData(ss, CONFIG.SHEETS.ATTENDANCE);
  const allFaults     = getSheetData(ss, CONFIG.SHEETS.FAULTS);
  const allAlarms     = getSheetData(ss, CONFIG.SHEETS.ALARMS);
  const allReadings   = getSheetData(ss, CONFIG.SHEETS.READINGS);

  const attendance = allAttendance.filter(r => r["Report_ID"] === reportId);
  const faults     = allFaults.filter(r => r["Report_ID"] === reportId);
  const alarms     = allAlarms.filter(r => r["Report_ID"] === reportId);
  const readings   = allReadings.filter(r => r["Report_ID"] === reportId);

  // ── 2. Copy the template Doc ─────────────────────────────
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  const dateLabel    = formatDate(row["Date"]);
  const shiftLabel   = (row["Shift_Type"] || "Shift").replace(" ", "_");
  const newFileName  = `MV_Shift_Report_${dateLabel}_${shiftLabel}`;

  const folder  = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const newFile = templateFile.makeCopy(newFileName, folder);
  const doc     = DocumentApp.openById(newFile.getId());
  const body    = doc.getBody();

  // ── 3. Replace scalar placeholders ───────────────────────
  replacePlaceholder(body, "<<[Date]>>",                 dateLabel);
  replacePlaceholder(body, "<<[Type]>>",                 (row["Shift_Type"] || "").toUpperCase());
  replacePlaceholder(body, "<<[Overall_System_Status]>>", row["Overall_System_Status"] || "");
  replacePlaceholder(body, "<<[System_Remarks]>>",        row["System_Remarks"] || "");
  replacePlaceholder(body, "<<[TBT_Conducted]>>",         row["TBT_Conducted"] || "");
  replacePlaceholder(body, "<<[TBT_Topic]>>",             row["TBT_Topic"] || "");
  replacePlaceholder(body, "<<[Incident_Reported]>>",     row["Incident_Reported"] || "");
  replacePlaceholder(body, "<<[Near_Miss]>>",             row["Near_Miss"] || "");
  replacePlaceholder(body, "<<[Safe_Man_Hours]>>",        row["Safe_Man_Hours"] || "");
  replacePlaceholder(body, "<<[Active_PTW]>>",            row["Active_PTW"] || "");
  replacePlaceholder(body, "<<[PPE_Compliance]>>",        row["PPE_Compliance"] || "");
  replacePlaceholder(body, "<<[HSE_Remarks]>>",           row["HSE_Remarks"] || "");
  replacePlaceholder(body, "<<[Pending_Items]>>",         row["Pending_Items"] || "");
  replacePlaceholder(body, "<<[Handover_Given_By]>>",     row["Handover_Given_By"] || "");
  replacePlaceholder(body, "<<[Handover_Received_By]>>",  row["Handover_Received_By"] || "");

  // ── 4. Expand repeating blocks ────────────────────────────
  expandTableBlock(body, "Related Shift_Attendances", attendance, [
    "Name", "File_No", "Designation", "Work_Location", "Shift", "Type"
  ]);

  expandTableBlock(body, "Related Fault_Logs", faults, [
    "Location_Feeder_Bay", "Fault_Type", "Time", "Restoration_Actions"
  ]);

  expandTableBlock(body, "Related Major_Alarms", alarms, [
    "Equipment_Location", "Alarm_Description", "Severity", "Time", "Reported_By"
  ]);

  expandReadingsBlock(body, "Related System_Readings", readings);

  // ── 5. Save and return URL ────────────────────────────────
  doc.saveAndClose();
  return newFile.getUrl();
}


// ============================================================
// PLACEHOLDER / TABLE HELPERS
// ============================================================

/**
 * Simple text find-and-replace throughout the document body.
 */
function replacePlaceholder(body, placeholder, value) {
  body.replaceText(escapeRegex(placeholder), value !== null && value !== undefined ? String(value) : "");
}

/**
 * Expand a <<Start:[tag]>> … <<End>> block inside a table row
 * into one row per data item, then remove the template row.
 *
 * Strategy:
 *   1. Find the table containing <<Start:[tag]>>
 *   2. Identify the template row
 *   3. For each data item, duplicate the row and fill values
 *   4. Delete the original template row
 */
function expandTableBlock(body, tag, dataRows, columns) {
  const startTag = `<<Start:[${tag}]>>`;
  const endTag   = "<<End>>";

  // Find which table contains the start tag
  const tables = body.getTables();
  for (let t = 0; t < tables.length; t++) {
    const table    = tables[t];
    const numRows  = table.getNumRows();

    for (let r = 0; r < numRows; r++) {
      const rowText = table.getRow(r).getText();
      if (!rowText.includes(startTag)) continue;

      const templateRow = table.getRow(r);

      if (dataRows.length === 0) {
        // No data – just clear the template markers and leave one blank row
        clearTagsInRow(templateRow, startTag, endTag, columns);
        return;
      }

      // Insert filled rows above the template row
      dataRows.forEach((dataItem, idx) => {
        const newRow = table.insertTableRow(r + idx); // inserts above template
        // Copy style from template row, cell by cell
        const numCells = templateRow.getNumCells();
        for (let c = 0; c < numCells; c++) {
          const templateCell = templateRow.getCell(c);
          const newCell      = newRow.getCell(c) || newRow.appendTableCell();
          let cellText = templateCell.getText();

          // Replace Start/End tags and column placeholders
          cellText = cellText.replace(startTag, "").replace(endTag, "");
          columns.forEach(col => {
            cellText = cellText.replace(new RegExp(escapeRegex(`<<[${col}]>>`), "g"),
                                        dataItem[col] !== undefined ? String(dataItem[col]) : "");
          });

          newCell.clear();
          newCell.appendParagraph(cellText)
                 .setAttributes(templateCell.getAttributes());
        }
      });

      // Delete the original template row (now shifted down by dataRows.length)
      table.removeRow(r + dataRows.length);
      return;
    }
  }
}

/**
 * Specialised expander for System_Readings which has two separate
 * repeating blocks (Incomer readings and Loop readings) in the same
 * table region, keyed by the same tag.
 *
 * Each reading row in the sheet becomes one table row in every
 * <<Start:[Related System_Readings]>> block found.
 */
function expandReadingsBlock(body, tag, readings) {
  const startTag = `<<Start:[${tag}]>>`;
  const endTag   = "<<End>>";

  const readingCols = [
    "Time_Interval",
    "Incomer1_kV","Incomer1_A","Incomer1_MW","Incomer1_Hz",
    "Incomer2_kV","Incomer2_A","Incomer2_MW","Incomer2_Hz",
    "Loop1_A_SWGRA","Loop1_A_SWGRB","Loop1_kW_SWGRA","Loop1_kW_SWGRB",
    "Loop2_A_SWGRA","Loop2_A_SWGRB","Loop2_kW_SWGRA","Loop2_kW_SWGRB",
    "Loop3_A_SWGRA","Loop3_A_SWGRB","Loop3_kW_SWGRA","Loop3_kW_SWGRB",
    "Loop1_NOP_SWGRA","Loop1_NOP_SWGRB",
    "Loop2_NOP_SWGRA","Loop2_NOP_SWGRB",
    "Loop3_NOP_SWGRA","Loop3_NOP_SWGRB",
  ];

  expandTableBlock(body, tag, readings, readingCols);

  // Also replace the NOP standalone placeholders if they exist outside tables
  if (readings.length > 0) {
    const r = readings[0]; // Use first reading for NOP display (usually static per shift)
    replacePlaceholder(body, "<<[Loop1_NOP_SWGRA]>>", r["Loop1_NOP_SWGRA"] || "");
    replacePlaceholder(body, "<<[Loop1_NOP_SWGRB]>>", r["Loop1_NOP_SWGRB"] || "");
    replacePlaceholder(body, "<<[Loop2_NOP_SWGRA]>>", r["Loop2_NOP_SWGRA"] || "");
    replacePlaceholder(body, "<<[Loop2_NOP_SWGRB]>>", r["Loop2_NOP_SWGRB"] || "");
    replacePlaceholder(body, "<<[Loop3_NOP_SWGRA]>>", r["Loop3_NOP_SWGRA"] || "");
    replacePlaceholder(body, "<<[Loop3_NOP_SWGRB]>>", r["Loop3_NOP_SWGRB"] || "");
  }
}

/**
 * Clear template marker tags from a row without adding data rows.
 */
function clearTagsInRow(row, startTag, endTag, columns) {
  for (let c = 0; c < row.getNumCells(); c++) {
    let txt = row.getCell(c).getText();
    txt = txt.replace(startTag, "").replace(endTag, "");
    columns.forEach(col => {
      txt = txt.replace(new RegExp(escapeRegex(`<<[${col}]>>`), "g"), "");
    });
    row.getCell(c).clear();
    row.getCell(c).appendParagraph(txt);
  }
}


// ============================================================
// SHEET UTILITIES
// ============================================================

/**
 * Read a sheet tab and return an array of plain objects,
 * one per data row, keyed by the header row.
 */
function getSheetData(ss, sheetName) {
  const sheet  = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

/**
 * Write "Generated" + Doc URL back to the Report_Status column
 * for the given Report_ID.
 */
function markReportGenerated(ss, reportId, docUrl) {
  const sheet   = ss.getSheetByName(CONFIG.SHEETS.SHIFT);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf("Report_ID");
  const stCol   = headers.indexOf("Report_Status");
  if (idCol < 0 || stCol < 0) return;

  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === reportId) {
      sheet.getRange(r + 1, stCol + 1).setValue(`Generated: ${docUrl}`);
      break;
    }
  }
}


// ============================================================
// FORMAT / MISC UTILITIES
// ============================================================

function formatDate(raw) {
  if (!raw) return "Unknown_Date";
  const d = (raw instanceof Date) ? raw : new Date((raw - 25569) * 86400000); // Excel serial
  return Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy-MM-dd");
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function showToast(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, "MV Report Generator", 5); }
  catch(e) { Logger.log(msg); }
}


// ============================================================
// MENU & TRIGGERS
// ============================================================

/** Adds a custom menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 MV Reports")
    .addItem("Generate Pending Reports", "generatePendingReports")
    .addSeparator()
    .addItem("Test: Log Sheet Names",    "logSheetNames")
    .addToUi();
}

function logSheetNames() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  Logger.log(ss.getSheets().map(s => s.getName()).join(", "));
}
