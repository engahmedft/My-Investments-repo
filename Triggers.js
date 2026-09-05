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

  /**
   * Returns the architectural guidelines and operational contract.
   * @returns {string}
   */
  getGuidelines() {
    return [
      "====================================================================================================",
      "DO NOT REMOVE OR STRIP COMMENTS FROM THIS FILE.",
      "----------------------------------------------------------------------------------------------------",
      "DYNAMIC COMMAND ENGINE & TRIGGER ARCHITECTURAL GUIDELINES:",
      "",
      "1. STRICT HEADER RESOLUTION (ROW 2):",
      "   - Reads 'Lookup_Values' strictly starting on Header Row 2 via BatchProcessor.read().",
      "   - Missing mandatory columns ('Category_ID', 'Item', 'Description') throw immediate errors.",
      "",
      "2. MULTI-WAY RESOLUTION (FAIL-FAST):",
      "   - Resolves commands by Value_ID, Item, Description, or normalized string without spaces/emojis.",
      "   - Zero silent failures: errors are recorded in System_Control!E2 and logged to Apps Script Logger.",
      "",
      "3. RUNTIME RETENTION & STATE TRACKING:",
      "   - Retains chosen action text in System_Control!B2 during execution.",
      "   - Sets Status (C2) to 'Running', updates timestamp in D2, and clears error in E2.",
      "   - On completion, sets Status (C2) to 'Completed' and resets B2 to '- Select Action -'.",
      "",
      "4. ATOMIC CONCURRENCY & BACKGROUND SAFETY:",
      "   - Safe for background triggers (no getUi calls in trigger flow).",
      "   - Uses LockService with 10s timeout to safely debounce onEdit and onChange events.",
      "===================================================================================================="
    ].join("\n");
  },

  getGuideline() {
    return this.getGuidelines();
  },

  /**
   * Universal Cleaner: Strips emojis, spaces, dashes, and underscores for 100% matching.
   */
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
      if (idx === -1) {
        throw new Error(`[Header Error] Missing mandatory column "${name}" in "${this.LOOKUP_SHEET}" (Row ${this.HEADER_ROW}).`);
      }
      return idx;
    };

    const getOptionalCol = (name, fallbackName = null) => {
      let idx = headers.indexOf(name.toLowerCase());
      if (idx === -1 && fallbackName) {
        idx = headers.indexOf(fallbackName.toLowerCase());
      }
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

  /**
   * Resolves and invokes action matching Value_ID, Item, Description, or clean string.
   */
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

      // Multi-way matching: Value_ID, exact Item, exact Description, or normalized clean string
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

    if (!targetFn) {
      targetFn = rawInput;
    }

    if (typeof globalThis[targetFn] !== "function") {
      throw new Error(`[Execution Error] Function "${targetFn}" is not defined in any script file.`);
    }

    // Direct invocation
    globalThis[targetFn]();
    return true;
  }
};

/**
 * ====================================================================
 * SPREADSHEET INITIALIZATION & TRIGGERS
 * ====================================================================
 */

function onOpen() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("🚀 System Control");

    const commands = DynamicCommandEngine.loadCommands(ss);

    commands.forEach(cmd => {
      if (typeof globalThis[cmd.item] === "function") {
        menu.addItem(cmd.description, cmd.item);
        if (cmd.addSeparator) {
          menu.addSeparator();
        }
      }
    });

    menu.addToUi();
    DynamicCommandEngine.syncDropdownValidation(ss);

  } catch (err) {
    Logger.log(`[onOpen Fatal Error]: ${err.stack || err.message}`);
  }
}

function refreshDropdown() {
  DynamicCommandEngine.syncDropdownValidation();
}

/**
 * Centralized Execution Router: Safe for both UI Dropdown edits and AppSheet API edits.
 */
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

// Global Entry Points
function optimizeAll() { new PortfolioManager().optimizeAll(); }
function resetAndRecalculateAll() { new PortfolioManager().resetAndRecalculateAll(); }
function sortAndResequenceGoogleFinance() { const pm = new PortfolioManager(); pm.runWithStatus(() => pm.sortGoogleFinance()); }
function sortAndResequenceTransactions() { const pm = new PortfolioManager(); pm.runWithStatus(() => pm.sortTransactions()); }
function setupPivotSortControls() { new PivotSortManager().setupControls(); }
function extendFormat(sheet) { TableFormatter.extend(sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()); }
function listAllFormulas() { new WorkbookAuditor().auditFormulas("Formula_Audit"); }
function listAllTables() { new WorkbookAuditor().auditTables("Tables_Audit"); }
function applyAllFormulaAudit() { new WorkbookAuditor().applyFormulas("Formula_Audit"); }
function runAllAudits() { new WorkbookAuditor().runAllAudits(); }
function fixAllArrayFormulaErrors(target = null) {
  return WorkbookAuditor.fixAllArrayFormulaErrors(target);
}