/**
 * Unified class managing Pivot Table control setup, fast event handling, and sorting operations.
 */
class PivotSortManager {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet] - Target active sheet
   */
  constructor(sheet = SpreadsheetApp.getActiveSheet()) {
    this.sheet = sheet;
  }

  /**
   * Entry point for onEdit trigger. 
   * Includes multi-stage fast exits so non-control edits exit immediately.
   * @param {Object} e - Trigger event object
   */
  static handleEdit(e) {
    // Stage 1 Fast Exits: Check event properties in memory before API calls
    if (!e || !e.value || !e.range) return;

    const range = e.range;

    // Fast Exit: Ensure edit was a single cell change
    if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

    const row = range.getRow();
    const col = range.getColumn();

    // Fast Exit: Controls are placed at col >= 2 and row >= 2
    if (col < 2 || row < 2) return;

    const sheet = range.getSheet();

    // Stage 2: Read single 3x2 matrix surrounding edit area (1 API call)
    const block = sheet.getRange(row - 1, col - 1, 3, 2).getValues();
    const label = block[1][0];

    // Fast Exit: Strict label check - exits instantly for ANY edit outside pivot sort dropdowns
    if (label !== "Sort By:" && label !== "Direction:") return;

    // Locate target pivot anchor
    const pivotAnchorRow = (label === "Sort By:") ? row + 3 : row + 2;
    const pivotAnchorCol = col - 1;

    const pivotTable = sheet.getPivotTables().find(p => {
      const anchor = p.getAnchorCell();
      return anchor.getColumn() === pivotAnchorCol && Math.abs(anchor.getRow() - pivotAnchorRow) <= 2;
    });

    if (!pivotTable) return;

    // Extract values from memory matrix block
    const sortBy = (label === "Sort By:") ? e.value : block[0][1];
    const direction = (label === "Direction:") ? e.value : block[2][1];

    // Apply sort using class instance
    new PivotSortManager(sheet).applySort(pivotTable, sortBy, direction);
  }

  /**
   * Sets up control dropdowns above all pivot tables on the sheet.
   */
  setupControls() {
    const pivots = this.sheet.getPivotTables();

    if (pivots.length === 0) {
      SpreadsheetApp.getUi().alert("No pivot tables found.");
      return;
    }

    // Keep direct reference to pivot object to avoid secondary API calls
    const pivotList = pivots.map(p => ({
      pivot: p,
      r: p.getAnchorCell().getRow(),
      c: p.getAnchorCell().getColumn()
    }));

    // Top-down sort for proper offset calculations
    pivotList.sort((a, b) => a.r - b.r);

    const initialData = this.sheet.getDataRange().getValues();
    let offset = 0;

    pivotList.forEach(item => {
      // 0-based array index for row 3 positions above anchor
      const checkRowIdx = item.r - 4;
      const checkColIdx = item.c - 1;

      const exists = checkRowIdx >= 0 && 
                     initialData[checkRowIdx] && 
                     initialData[checkRowIdx][checkColIdx] === "Sort By:";

      if (!exists) {
        this.sheet.insertRowsBefore(item.r + offset, 3);
        offset += 3;
      }
    });

    if (offset > 0) {
      SpreadsheetApp.flush();
    }

    // Re-read data once into memory after all insertions
    const newData = this.sheet.getDataRange().getValues();
    const dirRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Ascending", "Descending"], true)
      .build();

    pivotList.forEach(item => {
      // Pivot anchor cells update dynamically after insertRowsBefore
      const anchorRow = item.pivot.getAnchorCell().getRow();
      const anchorCol = item.pivot.getAnchorCell().getColumn();

      const numRowGroups = item.pivot.getRowGroups().length;

      // Header row is 1 row above anchor
      const headerRowIdx = anchorRow - 2;
      const rawRowData = newData[headerRowIdx] || [];
      const headersRaw = rawRowData.slice(anchorCol - 1);
      const headers = headersRaw.filter(String).slice(numRowGroups);

      if (headers.length === 0) headers.push("No Aggregates");

      const sortByRule = SpreadsheetApp.newDataValidation().requireValueInList(headers, true).build();

      // Read current values directly from memory
      const sortByRowIdx = anchorRow - 4;
      const dirRowIdx = anchorRow - 3;
      const valColIdx = anchorCol; // Column B relative to label cell in A

      const currentSortBy = newData[sortByRowIdx] ? newData[sortByRowIdx][valColIdx] : null;
      const currentDir = newData[dirRowIdx] ? newData[dirRowIdx][valColIdx] : null;

      const sortByVal = headers.includes(currentSortBy) ? currentSortBy : headers[0];
      const dirVal = (currentDir === "Ascending" || currentDir === "Descending") ? currentDir : "Descending";

      // Batch update formatting, values, and validations in a single range call
      const controlRange = this.sheet.getRange(anchorRow - 3, anchorCol, 2, 2);
      controlRange.setBackground("#f3f3f3")
                  .setBorder(true, true, true, true, null, null)
                  .setValues([
                    ["Sort By:", sortByVal],
                    ["Direction:", dirVal]
                  ])
                  .setDataValidations([
                    [null, sortByRule],
                    [null, dirRule]
                  ])
                  .setFontWeights([
                    ["bold", "normal"],
                    ["bold", "normal"]
                  ]);
    });

    this.sheet.getParent().toast("Controls Setup Complete");
  }

  /**
   * Core logic to update row group sorting on a pivot table.
   * @param {GoogleAppsScript.Spreadsheet.PivotTable} pivotTable 
   * @param {string} colName 
   * @param {string} direction 
   */
  applySort(pivotTable, colName, direction) {
    try {
      if (!colName || !direction) return;

      const anchor = pivotTable.getAnchorCell();
      const anchorRow = anchor.getRow();
      const anchorCol = anchor.getColumn();

      const maxSearchCols = Math.min(100, this.sheet.getLastColumn() - anchorCol + 1);
      if (maxSearchCols <= 0) return;

      const headers = this.sheet.getRange(anchorRow, anchorCol, 1, maxSearchCols).getValues()[0];
      const colIdx = headers.indexOf(colName);
      const rowGroups = pivotTable.getRowGroups();
      const pivotValues = pivotTable.getPivotValues();

      const valIdx = colIdx - rowGroups.length;

      if (rowGroups.length === 0 || valIdx < 0 || valIdx >= pivotValues.length) {
        console.warn("Could not match column to aggregate value.");
        return;
      }

      const targetValue = pivotValues[valIdx];
      const isAscending = (direction === "Ascending");

      rowGroups.forEach(group => {
        group.sortBy(targetValue, []);
        if (isAscending) {
          group.sortAscending();
        } else {
          group.sortDescending();
        }
      });

      this.sheet.getParent().toast(`Pivot sorted by ${colName} (${direction})`);
    } catch (err) {
      console.error("Sort failed: " + err.message);
    }
  }
}
