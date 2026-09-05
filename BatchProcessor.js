/**
 * Fully self-contained, high-performance batch data processor for Google Sheets.
 * Operates on 1 API Read + 1 API Write + (1 optional API Delete / Insert for reduced or expanded ranges).
 */
class BatchProcessor {


  /* ====================================================================================================
 * DO NOT REMOVE OR STRIP COMMENTS FROM THIS FILE.
 * ----------------------------------------------------------------------------------------------------
 * BATCHPROCESSOR ARCHITECTURAL CONTRACT:
 * 
 * 1. HIGH-PERFORMANCE 1-READ / 1-WRITE BATCH EXECUTION:
 *    Eliminates spreadsheet API latency by fetching data in 1 bulk read (`getDataRange().getValues()`), 
 *    transforming data purely in V8 memory, and writing back in 1 bulk write (`getRange().setValues()`).
 * 
 * 2. SHALLOW ROW CLONING FOR ACCURATE CHANGE DETECTION (`.slice()`):
 *    Before passing a row to `processFn`, `BatchProcessor` creates a shallow copy (`rawData[index].slice()`).
 *    This preserves `rawData[index]` with the original values read directly from Google Sheets.
 *    When `processFn` mutates the row in memory, `rawData[index]` remains untouched, allowing the Zero-Write
 *    pass (`pRow[c] !== oRow[c]`) to accurately detect true changes without in-memory reference equality bugs!
 * 
 * 3. ZERO-WRITE PASS FOR UNCHANGED DATA:
 *    Compares processed output against original raw sheet data (ignoring ArrayFormula columns).
 *    If no non-formula values changed and no rows were inserted/deleted, `BatchProcessor` skips `setValues()`
 *    write API calls entirely, making no-op runs finish in <0.5s!
 * 
 * 4. SMART FILTER PROTECTION (`_clearFilterCriteria`):
 *    Before performing batch writes, checks if columns have active filter criteria (`getColumnFilterCriteria`).
 *    Clears criteria ONLY on columns that currently have active filters hiding rows, preventing row misalignment.
 * 
 * 5. ARRAYFORMULA PROTECTION:
 *    Columns registered in `arrayFormulaHeaders` are automatically nullified (`null`) on write-back.
 *    This prevents Google Apps Script from overwriting live Google Sheets ArrayFormulas.
 * 
 * 6. GRID OVERFLOW & EXCESS ROW BOUNDARY PROTECTION:
 *    - Automatically expands sheet rows (`insertRowsAfter`) if in-memory insertions exceed current max rows.
 *    - Safely deletes excess trailing rows (`deleteRows`) when output rows are fewer than input rows, 
 *      guarded strictly against grid boundaries.
 * ==================================================================================================== */


  /**
   * Normalizes header strings by trimming whitespace and lowercasing.
   * @private
   */
  static _normalizeHeader(header) {
    return String(header ?? "").trim().toLowerCase();
  }

  /**
   * Clears filter criteria ONLY from columns that currently have active filter criteria hiding rows.
   * Preserves the Filter UI dropdown buttons while unhiding all rows.
   * @private
   */
  static _clearFilterCriteria(sheet) {
    if (!sheet) return;
    const filter = sheet.getFilter();
    if (!filter) return;

    const range = filter.getRange();
    const startCol = range.getColumn();
    const endCol = startCol + range.getNumColumns() - 1;

    for (let col = startCol; col <= endCol; col++) {
      if (filter.getColumnFilterCriteria(col) !== null) {
        filter.removeColumnFilterCriteria(col);
      }
    }
  }

  /**
   * Resolves a sheet input (string name or Sheet object) into a valid Sheet instance.
   * @private
   */
  static _resolveSheet(sheetTarget) {
    const sheet = typeof sheetTarget === "string"
      ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTarget)
      : sheetTarget;

    if (!sheet || typeof sheet.getDataRange !== "function") {
      throw new Error(`Invalid sheet target: "${sheetTarget}" was not found.`);
    }
    return sheet;
  }

  /**
   * PURE READ: Reads the actual data range of a sheet and returns a 2D array 
   * with normalized headers as Row 0 followed by raw data rows.
   * 
   * @param {string|GoogleAppsScript.Spreadsheet.Sheet} sheetTarget - Sheet name string or Sheet object.
   * @param {number} [headerRowNumber=1] - 1-based row number where headers reside.
   * @returns {Array<Array<any>>} [ [normalizedHeader1, normalizedHeader2, ...], [row1...], [row2...] ]
   */
  static read(sheetTarget, headerRowNumber = 1) {
    const sheet = BatchProcessor._resolveSheet(sheetTarget);

    // 1. SINGLE API READ
    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < headerRowNumber) {
      return [];
    }

    // 2. EXTRACT & NORMALIZE HEADERS (SINGLE PASS)
    const headerRowIndex = headerRowNumber - 1;
    const rawHeaderRow = allValues[headerRowIndex];
    const allNormalized = rawHeaderRow.map(h => BatchProcessor._normalizeHeader(h));

    let lastHeaderColIndex = -1;
    for (let c = allNormalized.length - 1; c >= 0; c--) {
      if (allNormalized[c] !== "") {
        lastHeaderColIndex = c;
        break;
      }
    }

    if (lastHeaderColIndex === -1) return []; // No valid headers found

    const numCols = lastHeaderColIndex + 1;
    const normalizedHeaders = allNormalized.slice(0, numCols);

    // 3. REVERSE SCAN FOR LAST DATA ROW (INLINED FAST-PATH LOOP)
    let lastDataRowIndex = -1;
    for (let r = allValues.length - 1; r >= headerRowNumber; r--) {
      const row = allValues[r];
      let hasData = false;

      for (let c = 0; c < numCols; c++) {
        const val = row[c];
        if (val !== "" && val !== null && val !== undefined) {
          hasData = true;
          break;
        }
      }

      if (hasData) {
        lastDataRowIndex = r;
        break;
      }
    }

    if (lastDataRowIndex < headerRowNumber) {
      return [normalizedHeaders]; // Return only headers if no data rows exist
    }

    // 4. SLICE RAW DATA ROWS (TRIMMED TO ACTUAL RANGE)
    const rawData = new Array(lastDataRowIndex - headerRowNumber + 1);
    for (let i = 0, r = headerRowNumber; r <= lastDataRowIndex; r++, i++) {
      const row = allValues[r];
      rawData[i] = row.length >= numCols ? row.slice(0, numCols) : row;
    }

    // 5. RETURN 2D ARRAY
    return [normalizedHeaders].concat(rawData);
  }

  /**
   * BATCH PROCESS & WRITE: Transforms data in-memory and writes changes back to sheet.
   * Automatically unhides filtered rows before processing to prevent row misalignment.
   * Features a Zero-Write pass using .slice() row cloning to skip setValues() if no data changed.
   * 
   * @param {string|GoogleAppsScript.Spreadsheet.Sheet} sheetTarget - Sheet name string or Sheet object.
   * @param {number} headerRowNumber - 1-based row number where headers reside.
   * @param {Function} processFn - Callback: (row, col, index, rawData) => Array<Array<any>>
   * @param {string[]} [arrayFormulaHeaders=[]] - Header names containing ARRAYFORMULAs.
   */
  static process(sheetTarget, headerRowNumber, processFn, arrayFormulaHeaders = []) {
    const targetSheet = BatchProcessor._resolveSheet(sheetTarget);

    // AUTOMATIC FILTER PROTECTION: Clear active filter criteria before processing
    BatchProcessor._clearFilterCriteria(targetSheet);

    // 1. FETCH RAW RANGE WITH NORMALIZED HEADERS
    const table = BatchProcessor.read(targetSheet, headerRowNumber);
    if (!table || table.length <= 1) {
      return; // No data rows to process
    }

    const normalizedHeaders = table[0];
    const rawData = table.slice(1);
    const oldNumRows = rawData.length;
    const numCols = normalizedHeaders.length;

    // 2. BUILD HEADER MAP & MEMOIZED COL HELPER
    const headerMap = new Map();
    normalizedHeaders.forEach((h, colIdx) => {
      if (h !== "") headerMap.set(h, colIdx);
    });

    const colCache = new Map();
    const col = (headerName) => {
      let idx = colCache.get(headerName);
      if (idx === undefined) {
        const normalized = BatchProcessor._normalizeHeader(headerName);
        idx = headerMap.has(normalized) ? headerMap.get(normalized) : -1;
        colCache.set(headerName, idx);
      }
      return idx;
    };

    const formulaColIndices = arrayFormulaHeaders
      .map(col)
      .filter(idx => idx !== -1);
    const formulaColSet = new Set(formulaColIndices);
    const numFormulaCols = formulaColIndices.length;

    // 3. TRANSFORM DATA IN-MEMORY (FAST NO-ALLOCATION PUSH LOOP)
    const processedData = [];
    for (let index = 0; index < oldNumRows; index++) {

      // SHALLOW ROW COPY (.slice()):
      // Creates a shallow copy of the row before passing it to processFn.
      // This prevents processFn from mutating rawData[index] in place.
      // Keeping rawData[index] pristine with original sheet values allows the Zero-Write
      // pass below (pRow[c] !== oRow[c]) to detect actual value changes accurately!
      const rowInput = rawData[index].slice();

      const res = processFn(rowInput, col, index, rawData);
      if (!Array.isArray(res) || res.length === 0) continue;

      for (let r = 0; r < res.length; r++) {
        const targetRow = res[r];
        if (Array.isArray(targetRow)) {
          while (targetRow.length < numCols) targetRow.push("");
          if (targetRow.length > numCols) targetRow.length = numCols;

          // Nullify ArrayFormula columns
          for (let f = 0; f < numFormulaCols; f++) {
            targetRow[formulaColIndices[f]] = null;
          }
          processedData.push(targetRow);
        }
      }
    }

    const startRow = headerRowNumber + 1;
    const newNumRows = processedData.length;

    // ZERO-WRITE PASS CHECK: Compare processed data against pristine original raw sheet data
    let dataHasChanged = (newNumRows !== oldNumRows);
    if (!dataHasChanged) {
      for (let r = 0; r < newNumRows; r++) {
        const pRow = processedData[r];
        const oRow = rawData[r]; // Unmutated original sheet values because rowInput was sliced
        for (let c = 0; c < numCols; c++) {
          if (formulaColSet.has(c)) continue; // Skip ArrayFormula columns from change detection
          if (pRow[c] !== oRow[c]) {
            dataHasChanged = true;
            break;
          }
        }
        if (dataHasChanged) break;
      }
    }

    // 4. WRITE BACK ONLY IF DATA CHANGED
    if (dataHasChanged && newNumRows > 0) {
      const requiredMaxRow = startRow + newNumRows - 1;
      const currentMaxRows = targetSheet.getMaxRows();
      if (requiredMaxRow > currentMaxRows) {
        targetSheet.insertRowsAfter(currentMaxRows, requiredMaxRow - currentMaxRows);
      }

      // Single 1-Write pass
      targetSheet.getRange(startRow, 1, newNumRows, numCols).setValues(processedData);
    }

    // Delete extra reduced rows if row count shrank
    if (newNumRows < oldNumRows) {
      const deleteStartRow = startRow + newNumRows;
      const rowsToDelete = oldNumRows - newNumRows;
      const currentMaxRows = targetSheet.getMaxRows();

      if (deleteStartRow <= currentMaxRows) {
        const safeRowsToDelete = Math.min(rowsToDelete, currentMaxRows - deleteStartRow + 1);
        if (safeRowsToDelete > 0) {
          targetSheet.deleteRows(deleteStartRow, safeRowsToDelete);
        }
      }
    }
  }
}

