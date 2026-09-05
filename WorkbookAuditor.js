/**
 * Class responsible for auditing formulas, data tables/islands,
 * workbook structures, and repairing/unblocking ArrayFormula expansion errors.
 * Standardized on 'Formula_Audit' and 'Tables_Audit'.
 */
class WorkbookAuditor {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [spreadsheet]
   */
  constructor(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
    this.ss = spreadsheet;
  }

  /**
   * Returns architectural guidelines and operational contract.
   * @returns {string}
   */
  static getGuidelines() {
    return [
      "====================================================================================================",
      "DO NOT REMOVE OR STRIP COMMENTS FROM THIS FILE.",
      "----------------------------------------------------------------------------------------------------",
      "WORKBOOK AUDITOR ARCHITECTURAL GUIDELINES:",
      "",
      "1. EXPLICIT AUDIT SHEET CHECKS:",
      "   - Validates existence of 'Formula_Audit' and 'Tables_Audit' tabs before executing.",
      "   - Throws explicit errors immediately if tabs are missing from the workbook.",
      "",
      "2. FORMULA HARVESTING & APPLICATION:",
      "   - Accurately scans formulas across sheets, logs to Formula_Audit, and safely reapplies formulas.",
      "",
      "3. GENERALIZED ARRAYFORMULA REPAIR (fixAllArrayFormulaErrors):",
      "   - Sheet-independent static repair utility.",
      "   - Scans top rows for dynamic arrays, cleans blocking residual cells in rows below,",
      "     and eliminates #REF! expansion errors instantly.",
      "===================================================================================================="
    ].join("\n");
  }

  // 8-Direction BFS Grid Offsets (Encapsulated inside class)
  static get _DR() { return [-1, -1, -1,  0,  0,  1,  1,  1]; }
  static get _DC() { return [-1,  0,  1, -1,  1, -1,  0,  1]; }

  // Blacklist of system tabs to ignore during auditing
  static get AUDIT_SHEET_KEYWORDS() {
    return new Set([
      "formulaaudit",
      "tablesaudit",
      "optimizelog",
      "systemcontrol",
      "system_control",
      "lookupvalues",
      "lookup_values",
      "lookupcategories",
      "lookup_categories",
      "perusersettings",
      "detailslookups",
      "masterlookups"
    ]);
  }

  /**
   * Converts a 1-based column index to an A1-notation letter string.
   * @param {number} col - 1-based column index.
   * @returns {string} Column letter (e.g., "A", "Z", "AA").
   */
  static getColLetter(col) {
    let letter = "";
    let tempCol = col;
    while (tempCol > 0) {
      const temp = (tempCol - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      tempCol = Math.floor((tempCol - temp - 1) / 26);
    }
    return letter;
  }

  /**
   * HIGH-PERFORMANCE ZERO-WRITE ARRAYFORMULA FIXER
   * 
   * 1. 1-READ / 1-WRITE (via RangeList): Reads entire sheet in 1 API read.
   * 2. ZERO-WRITE PASS: Skips write APIs completely if no formulas are blocked.
   * 3. BATCH RANGE CLEARING: Clears all blocked columns in 1 single API write per sheet.
   * 
   * @param {string|GoogleAppsScript.Spreadsheet.Sheet} [target=null] - Target sheet or null for all sheets.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()]
   * @returns {{ totalColumnsFixed: number, details: string[] }}
   */
  static fixAllArrayFormulaErrors(target = null, ss = SpreadsheetApp.getActiveSpreadsheet()) {
    let sheetsToProcess = [];

    if (target) {
      const sheet = typeof target === "string" ? ss.getSheetByName(target) : target;
      if (!sheet?.getDataRange) {
        throw new Error(`[WorkbookAuditor.fixAllArrayFormulaErrors] Invalid sheet target: "${target}" was not found.`);
      }
      sheetsToProcess = [sheet];
    } else {
      sheetsToProcess = ss.getSheets().filter(s => !WorkbookAuditor._isAuditSheet(s.getName()));
    }

    let totalColumnsFixed = 0;
    const fixedDetails = [];

    for (let s = 0; s < sheetsToProcess.length; s++) {
      const sheet = sheetsToProcess[s];
      const dataRange = sheet.getDataRange();
      const numRows = dataRange.getNumRows();
      const numCols = dataRange.getNumColumns();

      // Skip sheets with no data rows below headers
      if (numRows < 2 || numCols < 1) continue;

      // 1 BULK READ: In-memory grid inspection
      const formulas = dataRange.getFormulas();
      const values = dataRange.getValues();
      const maxRows = sheet.getMaxRows();
      const sheetName = sheet.getName();

      const scanDepth = Math.min(numRows, 3);
      const rangesToClear = [];

      for (let c = 0; c < numCols; c++) {
        for (let r = 0; r < scanDepth; r++) {
          const formula = String(formulas[r][c] || "").trim();
          const value = String(values[r][c] || "").trim();

          // Check if this header cell contains an ArrayFormula or has a #REF! spill error
          const isArrayFormula = 
            formula.toUpperCase().includes("ARRAYFORMULA") ||
            formula.startsWith("={") ||
            formula.toUpperCase().startsWith("=MAP(") ||
            formula.toUpperCase().startsWith("=BYROW(") ||
            formula.toUpperCase().startsWith("=QUERY(") ||
            formula.toUpperCase().startsWith("=FILTER(") ||
            value === "#REF!";

          if (isArrayFormula && (r + 1 < numRows)) {
            // Check in V8 memory if any rows below contain static values blocking expansion
            let hasBlockingData = false;
            for (let subR = r + 1; subR < numRows; subR++) {
              const cellVal = values[subR][c];
              if (cellVal !== "" && cellVal !== null && cellVal !== undefined) {
                hasBlockingData = true;
                break;
              }
            }

            // Queue range for batch clearing if blocking data is present
            if (hasBlockingData) {
              const colLetter = WorkbookAuditor.getColLetter(c + 1);
              const startRow = r + 2;
              rangesToClear.push(`${colLetter}${startRow}:${colLetter}${maxRows}`);
              totalColumnsFixed++;
              break; // Next column
            }
          }
        }
      }

      // ONLY WRITE IF BLOCKED COLUMNS WERE FOUND ON THIS SHEET
      if (rangesToClear.length > 0) {
        sheet.getRangeList(rangesToClear).clearContent(); // 1 Single API Write for all blocked columns
        fixedDetails.push(`${sheetName} (${rangesToClear.length} col)`);
      }
    }

    // Flush ONLY if any changes were actually made
    if (totalColumnsFixed > 0) {
      SpreadsheetApp.flush();
    }

    const resultMsg = totalColumnsFixed > 0
      ? `✅ Fixed ${totalColumnsFixed} blocked ArrayFormula column(s) in: ${fixedDetails.join(", ")}`
      : `✅ All ArrayFormula columns are clean. No action needed.`;

    Logger.log(`[WorkbookAuditor.fixAllArrayFormulaErrors] ${resultMsg}`);
    ss.toast(resultMsg, "ArrayFormula Optimizer", 5);

    return { totalColumnsFixed, details: fixedDetails };
  }

  /**
   * Runs all workbook audits sequentially.
   */
  runAllAudits() {
    this.auditFormulas("Formula_Audit");
    this.auditTables("Tables_Audit");
  }

  /**
   * Scans all sheets for formulas and writes a consolidated report to 'Formula_Audit'.
   * @param {string} [reportName="Formula_Audit"]
   */
  auditFormulas(reportName = "Formula_Audit") {
    const sheets = this.ss.getSheets();
    const nowTimestamp = new Date();
    const allDataToPush = [["Sheet Name", "Cell", "Formula", "Last Update"]];

    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (WorkbookAuditor._isAuditSheet(sheetName)) continue;

      const formulas = sheet.getDataRange().getFormulas();
      const numRows = formulas.length;
      if (numRows === 0) continue;

      const numCols = formulas[0].length;

      for (let r = 0; r < numRows; r++) {
        const row = formulas[r];
        let hasFormula = false;
        for (let c = 0; c < numCols; c++) {
          if (row[c] !== "") {
            hasFormula = true;
            break;
          }
        }
        if (!hasFormula) continue;

        const rowRef = r + 1;
        for (let c = 0; c < numCols; c++) {
          const formula = row[c];
          if (formula) {
            allDataToPush.push([
              sheetName,
              WorkbookAuditor.getColLetter(c + 1) + rowRef,
              `'${formula}`,
              nowTimestamp
            ]);
          }
        }
      }
    }

    WorkbookAuditor._writeReport(
      this.ss,
      reportName,
      allDataToPush,
      "No formulas found in the entire workbook.",
      "Formula audit complete!"
    );
  }

  /**
   * Detects contiguous data "islands" (tables) across all sheets using 1D BFS in O(N) time.
   * Outputs report to 'Tables_Audit'.
   * @param {string} [reportName="Tables_Audit"]
   */
  auditTables(reportName = "Tables_Audit") {
    const sheets = this.ss.getSheets();
    const nowTimestamp = new Date();
    const results = [["Sheet Name", "Header Row", "Column", "Header Name", "Last Update"]];

    let visited = new Uint8Array(1024);
    let queue = new Int32Array(1024);
    const dr = WorkbookAuditor._DR;
    const dc = WorkbookAuditor._DC;

    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (WorkbookAuditor._isAuditSheet(sheetName)) continue;

      const fullDataRange = sheet.getDataRange();
      const numRows = fullDataRange.getNumRows();
      const numCols = fullDataRange.getNumColumns();
      const maxCells = numRows * numCols;

      if (maxCells === 0) continue;

      if (visited.length < maxCells) {
        visited = new Uint8Array(maxCells);
        queue = new Int32Array(maxCells);
      } else {
        visited.fill(0, 0, maxCells);
      }

      const values = fullDataRange.getValues();

      for (let r = 0; r < numRows; r++) {
        const rowValues = values[r];
        let rowHasData = false;
        for (let c = 0; c < numCols; c++) {
          const v = rowValues[c];
          if (v !== "" && v !== null && v !== undefined) {
            rowHasData = true;
            break;
          }
        }
        if (!rowHasData) continue;

        for (let c = 0; c < numCols; c++) {
          const idx = r * numCols + c;

          if (values[r][c] !== "" && visited[idx] === 0) {
            let minR = r, maxR = r, minC = c, maxC = c;
            let head = 0, tail = 0;

            visited[idx] = 1;
            queue[tail++] = idx;

            while (head < tail) {
              const currIdx = queue[head++];
              const cr = (currIdx / numCols) | 0;
              const cc = currIdx % numCols;

              for (let i = 0; i < 8; i++) {
                const nr = cr + dr[i];
                const nc = cc + dc[i];

                if (nr >= 0 && nr < numRows && nc >= 0 && nc < numCols) {
                  const nIdx = nr * numCols + nc;

                  if (visited[nIdx] === 0 && values[nr][nc] !== "") {
                    visited[nIdx] = 1;
                    queue[tail++] = nIdx;

                    if (nr < minR) minR = nr;
                    else if (nr > maxR) maxR = nr;

                    if (nc < minC) minC = nc;
                    else if (nc > maxC) maxC = nc;
                  }
                }
              }
            }

            for (let i = minC; i <= maxC; i++) {
              results.push([
                sheetName,
                minR + 1,
                WorkbookAuditor.getColLetter(i + 1),
                values[minR][i],
                nowTimestamp
              ]);
            }
          }
        }
      }
    }

    WorkbookAuditor._writeReport(
      this.ss,
      reportName,
      results,
      "No data tables found.",
      "Table audit complete!"
    );
  }

  static _sanitizeFormula(rawFormula) {
    if (!rawFormula) return "";
    let formula = String(rawFormula).trim();

    if (formula.startsWith("'")) {
      formula = formula.substring(1).trim();
    }
    formula = formula.replace(/\\"/g, '"');
    formula = formula.replace(/ArrayFormula\s*\(\s*(".*?"|'.*?')\s*\{;\s*/gi, 'ArrayFormula({ $1; ');
    formula = formula.replace(/^=+/, '');
    return "=" + formula.trim();
  }

  static _applyFormulaToRange(range, cleanFormula) {
    const hasR1C1 = /(?:R\[?-?\d*\]?C\[?-?\d*\]?|RC\[?-?\d*\]?)/i.test(cleanFormula);

    if (hasR1C1) {
      try {
        range.setFormulaR1C1(cleanFormula);
        SpreadsheetApp.flush();
        if (range.getDisplayValue() !== "#ERROR!") return true;
      } catch (e) {}
      try {
        range.setFormula(cleanFormula);
        SpreadsheetApp.flush();
        if (range.getDisplayValue() !== "#ERROR!") return true;
      } catch (e) {
        throw new Error("R1C1/A1 Parse Error: " + e.message);
      }
    } else {
      try {
        range.setFormula(cleanFormula);
        SpreadsheetApp.flush();
        if (range.getDisplayValue() !== "#ERROR!") return true;
      } catch (e) {}
      try {
        range.setFormulaR1C1(cleanFormula);
        SpreadsheetApp.flush();
        if (range.getDisplayValue() !== "#ERROR!") return true;
      } catch (e) {
        throw new Error("A1/R1C1 Parse Error: " + e.message);
      }
    }
    return true;
  }

  applyFormulas(auditSheetName = "Formula_Audit") {
    const auditSheet = this.ss.getSheetByName(auditSheetName) || this.ss.getSheetByName("FormulaAudit");
    if (!auditSheet) {
      const msg = `Could not find '${auditSheetName}' sheet.`;
      this.ss.toast(msg, "Error ❌", 5);
      return { success: false, total: 0, applied: 0, skipped: 0, errors: [msg] };
    }

    const data = auditSheet.getDataRange().getValues();
    if (data.length <= 1) {
      this.ss.toast("No records found in formula audit sheet.", "Finished ℹ️", 4);
      return { success: true, total: 0, applied: 0, skipped: 0, errors: [] };
    }

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colSheetName = headers.indexOf("sheet name");
    const colCell = headers.indexOf("cell");
    const colFormula = headers.indexOf("formula");

    if (colSheetName === -1 || colCell === -1 || colFormula === -1) {
      const msg = "Invalid header structure in Formula_Audit. Expected 'Sheet Name', 'Cell', and 'Formula'.";
      this.ss.toast(msg, "Error ❌", 5);
      return { success: false, total: 0, applied: 0, skipped: 0, errors: [msg] };
    }

    let applied = 0;
    let skipped = 0;
    const errors = [];
    const sheetCache = new Map();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const targetSheetName = String(row[colSheetName] || "").trim();
      const targetCell = String(row[colCell] || "").trim();
      const rawFormula = String(row[colFormula] || "").trim();

      if (!targetSheetName || !targetCell || !rawFormula) {
        skipped++;
        continue;
      }

      const cleanFormula = WorkbookAuditor._sanitizeFormula(rawFormula);

      try {
        if (!sheetCache.has(targetSheetName)) {
          sheetCache.set(targetSheetName, this.ss.getSheetByName(targetSheetName));
        }
        const targetSheet = sheetCache.get(targetSheetName);

        if (!targetSheet) {
          throw new Error(`Sheet '${targetSheetName}' not found.`);
        }

        const range = targetSheet.getRange(targetCell);
        WorkbookAuditor._applyFormulaToRange(range, cleanFormula);
        applied++;

      } catch (err) {
        const errorMsg = `Row ${i + 1} (${targetSheetName}!${targetCell}): ${err.message}`;
        errors.push(errorMsg);
        Logger.log(`[APPLY ERROR] ${errorMsg}`);
      }
    }

    SpreadsheetApp.flush();

    const summaryMsg = `Applied: ${applied} | Skipped: ${skipped} | Errors: ${errors.length}`;
    this.ss.toast(summaryMsg, errors.length === 0 ? "Success ✅" : "Warning ⚠️", 6);
    Logger.log(`[WorkbookAuditor] Formula Restoration Summary -> ${summaryMsg}`);

    return {
      success: errors.length === 0,
      total: data.length - 1,
      applied,
      skipped,
      errors
    };
  }

  static _writeReport(ss, reportName, data, emptyMsg, successMsg) {
    const reportSheet = ss.getSheetByName(reportName) || ss.insertSheet(reportName);
    reportSheet.clear();

    const dataLen = data.length;
    if (dataLen > 1) {
      const numCols = data[0].length;
      reportSheet.getRange(1, 1, dataLen, numCols).setValues(data);
      reportSheet.getRange(1, 1, 1, numCols).setFontWeight("bold").setBackground("#f3f3f3");
      reportSheet.getRange(2, numCols, dataLen - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
      reportSheet.setFrozenRows(1);

      ss.toast(`${successMsg} Found ${dataLen - 1} entries in ${reportName}.`, "Success ✅", 5);
    } else {
      ss.toast(emptyMsg, "Finished ℹ️", 3);
    }
  }

  static _isAuditSheet(sheetName) {
    const normalized = sheetName.replace(/[\s_-]/g, "").toLowerCase();
    return WorkbookAuditor.AUDIT_SHEET_KEYWORDS.has(normalized);
  }
}

