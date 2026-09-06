/**
 * ====================================================================
 * DYNAMIC COMMAND ENGINE & TRIGGER HANDLER (AppSheet & UI Dropdown Safe)
 * ====================================================================
 */

const DynamicCommandEngine = {
  CATEGORY_ID: 5,
  LOOKUP_SHEET: "Lookup_Values",
  CONTROL_SHEET: "System_Control",
  HEADER_ROW: 2,

  getGuidelines() {
    return [
      "====================================================================================================",
      "DYNAMIC COMMAND ENGINE & TRIGGER ARCHITECTURAL GUIDELINES:",
      "1. Reads 'Lookup_Values' strictly starting on Header Row 2 via BatchProcessor.read().",
      "2. Multi-way matching: Value_ID, Item, Description, or normalized string.",
      "3. Retains action text in System_Control!B2 during run; completes with atomic lock release.",
      "4. Safe for both UI Dropdown edits (onEdit) and AppSheet API edits (onChange).",
      "===================================================================================================="
    ].join("\n");
  },

  cleanText(str) {
    return String(str || "")
      .replace(/[\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]/gu, "")
      .replace(/[\s_\-]/g, "")
      .trim()
      .toLowerCase();
  },

  loadCommands(ss = SpreadsheetApp.getActiveSpreadsheet()) {
    const table = BatchProcessor.read(this.LOOKUP_SHEET, this.HEADER_ROW);
    if (!table || table.length <= 1) {
      throw new Error(`[DynamicCommandEngine] No data found in "${this.LOOKUP_SHEET}" at Header Row ${this.HEADER_ROW}.`);
    }

    const headers = table[0];
    const getRequiredCol = (name) => {
      const idx = headers.indexOf(name.toLowerCase());
      if (idx === -1) throw new Error(`Missing column "${name}" in "${this.LOOKUP_SHEET}".`);
      return idx;
    };

    const getOptionalCol = (name, fallback = null) => {
      let idx = headers.indexOf(name.toLowerCase());
      if (idx === -1 && fallback) idx = headers.indexOf(fallback.toLowerCase());
      return idx;
    };

    const idIdx     = getOptionalCol("value_id", "id");
    const catIdx    = getRequiredCol("category_id");
    const itemIdx   = getRequiredCol("item");
    const descIdx   = getRequiredCol("description");
    const orderIdx  = getOptionalCol("sort_order", "order");
    const activeIdx = getOptionalCol("is_active", "active");
    const extraIdx  = getOptionalCol("extra_value_1");

    const rawRows = table.slice(1);
    const commands = [];

    for (let i = 0; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (Number(row[catIdx]) !== this.CATEGORY_ID) continue;

      if (activeIdx !== -1) {
        const act = String(row[activeIdx] ?? "").trim().toUpperCase();
        if (act === "N" || act === "FALSE" || row[activeIdx] === false) continue;
      }

      const rawItem = String(row[itemIdx] || "").trim();
      if (!rawItem || rawItem.startsWith("-")) continue;

      const rawDesc = String(row[descIdx] || "").trim();
      const desc = rawDesc || rawItem;
      const valueId = idIdx !== -1 ? String(row[idIdx] || "").trim() : "";
      const order = orderIdx !== -1 ? (Number(row[orderIdx]) || (i + 1)) : (i + 1);
      const extraVal = extraIdx !== -1 ? String(row[extraIdx] || "").trim().toUpperCase() : "";

      commands.push({
        valueId: valueId,
        item: rawItem,
        description: desc,
        order: order,
        addSeparator: (extraVal === "Y" || extraVal === "YES" || extraVal === "TRUE"),
        rowIndex: this.HEADER_ROW + 1 + i
      });
    }

    return commands.sort((a, b) => a.order - b.order);
  },

  syncDropdownValidation(ss = SpreadsheetApp.getActiveSpreadsheet()) {
    const controlSheet = ss.getSheetByName(this.CONTROL_SHEET);
    if (!controlSheet) return;

    const commands = this.loadCommands(ss);
    const dropdownValues = ["- Select Action -"];

    commands.forEach(cmd => {
      if (cmd.description && !dropdownValues.includes(cmd.description)) {
        dropdownValues.push(cmd.description);
      }
    });

    const b2Cell = controlSheet.getRange("B2");
    b2Cell.clearDataValidations();

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdownValues, true)
      .setAllowInvalid(true)
      .build();

    b2Cell.setDataValidation(rule);

    const cur = String(b2Cell.getValue() || "").trim();
    if (!cur || !dropdownValues.includes(cur)) {
      b2Cell.setValue("- Select Action -");
    }
  },

  dispatch(actionInput) {
    if (!actionInput) {
      throw new Error("[DynamicCommandEngine] No action input provided.");
    }

    const rawInput = String(actionInput).trim();
    const cleanInput = this.cleanText(rawInput);
    const commands = this.loadCommands();
    let targetFn = null;

    for (let i = 0; i < commands.length; i++) {
      const cmd = commands[i];
      if (
        rawInput === cmd.valueId ||
        rawInput === cmd.item ||
        rawInput === cmd.description ||
        cleanInput === this.cleanText(cmd.valueId) ||
        cleanInput === this.cleanText(cmd.item) ||
        cleanInput === this.cleanText(cmd.description)
      ) {
        targetFn = cmd.item;
        break;
      }
    }

    if (!targetFn) targetFn = rawInput;

    // Resolve camelCase canonical aliases
    const aliases = {
      "optimizeall": "optimizeAll",
      "optimize all": "optimizeAll",
      "resetandrecalculateall": "resetAndRecalculateAll",
      "reset & recalculate all stocks": "resetAndRecalculateAll",
      "sorttransactions": "sortTransactions",
      "sort transactions": "sortTransactions",
      "sortandresequencetransactions": "sortTransactions",
      "sortgooglefinance": "sortGoogleFinance",
      "sort google finance": "sortGoogleFinance",
      "sortandresequencegooglefinance": "sortGoogleFinance",
      "addsortfacilitytopivot": "setupPivotSortControls",
      "add sort facility to pivot": "setupPivotSortControls",
      "setuppivotsortcontrols": "setupPivotSortControls",
      "extendformat": "extendFormat",
      "extend format": "extendFormat",
      "listallformulas": "listAllFormulas",
      "list all formulas": "listAllFormulas",
      "listalltables": "listAllTables",
      "list all tables": "listAllTables",
      "runallaudits": "runAllAudits",
      "run all audits": "runAllAudits",
      "fixallarrayformulaerrors": "fixAllArrayFormulaErrors"
    };

    const resolvedName = aliases[this.cleanText(targetFn)] || targetFn;

    if (typeof globalThis[resolvedName] !== "function") {
      throw new Error(`[Execution Error] Function "${targetFn}" (resolved: "${resolvedName}") is not defined.`);
    }

    globalThis[resolvedName]();
    return true;
  }
};

/**
 * TRIGGER ENTRY POINTS
 */
function onOpen() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("🚀 System Control");

    const commands = DynamicCommandEngine.loadCommands(ss);
    commands.forEach(cmd => {
      const cleanTarget = DynamicCommandEngine.cleanText(cmd.item);
      if (typeof globalThis[cmd.item] === "function" || cleanTarget) {
        menu.addItem(cmd.description, cmd.item);
        if (cmd.addSeparator) menu.addSeparator();
      }
    });

    menu.addToUi();
    DynamicCommandEngine.syncDropdownValidation(ss);
  } catch (err) {
    Logger.log(`[onOpen Error]: ${err.stack || err.message}`);
  }
}

/** Handles direct user edits in the Google Sheet UI */
function onEdit(e) {
  if (typeof PivotSortManager !== "undefined") {
    PivotSortManager.handleEdit(e);
  }
  handleExternalChange(e);
}


function handleExternalChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlSheet = ss.getSheetByName("System_Control");
  if (!controlSheet) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return;

  try {
    const runCell = controlSheet.getRange("B2");
    const rawAction = String(runCell.getValue() || "").trim();

    if (
      !rawAction || 
      rawAction === "- Select Action -" || 
      rawAction === "Not Running" || 
      rawAction === "IDLE" || 
      rawAction === "Completed"
    ) {
      return;
    }

    Logger.log(`[handleExternalChange] Dispatching: "${rawAction}"`);
    DynamicCommandEngine.dispatch(rawAction);

  } catch (err) {
    Logger.log(`[handleExternalChange Error]: ${err.stack || err.message}`);
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    controlSheet.getRange("C2").setValue("Error");
    controlSheet.getRange("D2").setValue(timestamp);
    controlSheet.getRange("E2").setValue(err.message || String(err));
    controlSheet.getRange("B2").setValue("- Select Action -");
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// Global API & Command Aliases
function optimizeAll() { new PortfolioManager().optimizeAll(); }
function resetAndRecalculateAll() { new PortfolioManager().resetAndRecalculateAll(); }
function sortTransactions() { const pm = new PortfolioManager(); pm.runWithStatus(() => pm.sortTransactions()); }
function sortAndResequenceTransactions() { sortTransactions(); }
function sortGoogleFinance() { const pm = new PortfolioManager(); pm.runWithStatus(() => pm.sortGoogleFinance()); }
function sortAndResequenceGoogleFinance() { sortGoogleFinance(); }
function setupPivotSortControls() { new PivotSortManager().setupControls(); }
function addSortFacilityToPivot() { setupPivotSortControls(); }
function extendFormat(sheet) { TableFormatter.extend(sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()); }
function listAllFormulas() { new WorkbookAuditor().auditFormulas("Formula_Audit"); }
function listAllTables() { new WorkbookAuditor().auditTables("Tables_Audit"); }
function applyAllFormulaAudit() { new WorkbookAuditor().applyFormulas("Formula_Audit"); }
function runAllAudits() { new WorkbookAuditor().runAllAudits(); }
function fixAllArrayFormulaErrors(target = null) { return WorkbookAuditor.fixAllArrayFormulaErrors(target); }