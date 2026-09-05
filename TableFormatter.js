/**
 * Class responsible for dynamically formatting, updating data validation,
 * adjusting bandings/borders, and managing filters across spreadsheet tables.
 * Optimized with zero-write early exit when table dimensions are unchanged.
 */
class TableFormatter {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] - Target sheet. Defaults to active sheet.
   */
  constructor(sheet = SpreadsheetApp.getActiveSpreadsheet()?.getActiveSheet()) {
    this.sheet = sheet;
  }

  /**
   * Refreshes formatting, data validation, and filters for the target sheet.
   * Features a smart early-exit path when table dimensions are unchanged.
   */
  format() {
    if (!this.sheet) return;

    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const numRows = values?.length || 0;
    if (numRows === 0) return;
    
    const numCols = values[0]?.length || 0;
    if (numCols === 0) return;

    const dataStartRow = dataRange.getRow();
    const dataStartCol = dataRange.getColumn();
    const dataLastCol  = dataStartCol + numCols - 1;

    // 1. Dynamic Header Row & Last Data Row Detection (Fast Inlined Loops)
    let headerIndex = -1;
    for (let r = 0; r < numRows; r++) {
      const row = values[r];
      let hasData = false;
      for (let c = 0; c < numCols; c++) {
        const val = row[c];
        if (val !== "" && val !== null && val !== undefined) {
          hasData = true;
          break;
        }
      }
      if (hasData) {
        headerIndex = r;
        break;
      }
    }

    if (headerIndex === -1 || headerIndex >= numRows - 1) return;

    let lastIndex = -1;
    for (let r = numRows - 1; r >= headerIndex; r--) {
      const row = values[r];
      let hasData = false;
      for (let c = 0; c < numCols; c++) {
        const val = row[c];
        if (val !== "" && val !== null && val !== undefined) {
          hasData = true;
          break;
        }
      }
      if (hasData) {
        lastIndex = r;
        break;
      }
    }

    const headerRow = dataStartRow + headerIndex;
    const lastRow   = dataStartRow + lastIndex;
    const startRow  = headerRow + 1;

    if (lastRow <= headerRow) return;

    // 2. Filter Evaluation
    const filter = this.sheet.getFilter();
    let filterNeedsUpdate = false;
    let templateRow = startRow;

    if (filter) {
      const fRange = filter.getRange();
      const isExactMatch = 
        fRange.getColumn() === dataStartCol &&
        fRange.getLastColumn() === dataLastCol &&
        fRange.getRow() === headerRow &&
        fRange.getLastRow() === lastRow;

      if (isExactMatch) {
        templateRow = lastRow;
      } else {
        filterNeedsUpdate = true;
        const fLastRow = fRange.getLastRow();
        templateRow = fLastRow >= startRow ? Math.min(fLastRow, lastRow) : startRow;
      }
    } else {
      filterNeedsUpdate = true;
      templateRow = Math.max(startRow, lastRow - 1);
    }

    // SMART EARLY EXIT: If row count didn't grow and filter is exact match, skip expensive formatting writes!
    if (lastRow <= templateRow && !filterNeedsUpdate) {
      return; 
    }

    // 3. Incremental Formatting & Data Validation Copying (Only if new rows exist)
    if (lastRow > templateRow) {
      const numRowsToCopy = lastRow - templateRow;
      const sourceRange   = this.sheet.getRange(templateRow, dataStartCol, 1, numCols);
      const pasteTarget   = sourceRange.offset(1, 0, numRowsToCopy);

      sourceRange.copyTo(pasteTarget, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      const rules = sourceRange.getDataValidations()[0];
      let hasValidation = false;
      for (let i = 0; i < rules.length; i++) {
        if (rules[i] !== null) {
          hasValidation = true;
          break;
        }
      }

      if (hasValidation) {
        sourceRange.copyTo(pasteTarget, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      }
    }

    // 4. Handle Bandings and Borders (Only if dimensions changed)
    const fullTableRange = this.sheet.getRange(headerRow, dataStartCol, lastRow - headerRow + 1, numCols);

    const bandings = this.sheet.getBandings();
    let tableBanding = null;
    for (let i = 0; i < bandings.length; i++) {
      const bRange = bandings[i].getRange();
      if (
        bRange.getRow() <= lastRow &&
        bRange.getLastRow() >= headerRow &&
        bRange.getColumn() <= dataLastCol &&
        bRange.getLastColumn() >= dataStartCol
      ) {
        tableBanding = bandings[i];
        break;
      }
    }

    if (tableBanding) {
      if (tableBanding.getRange().getLastRow() !== lastRow) {
        tableBanding.setRange(fullTableRange);
      }
    } else {
      const bodyRange = this.sheet.getRange(startRow, dataStartCol, lastRow - startRow + 1, numCols);
      bodyRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    }

    // 5. Recreate Filter Only If Dimensions Changed
    if (filterNeedsUpdate) {
      if (filter) filter.remove();
      fullTableRange.createFilter();
    }
  }

  /**
   * Static helper method for quick execution without explicit instantiation.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet]
   */
  static extend(sheet) {
    new TableFormatter(sheet).format();
  }
}

