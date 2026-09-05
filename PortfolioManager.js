/**
 * Core orchestrator and calculation engine for investment portfolio operations.
 */
class PortfolioManager {


  /* ====================================================================================================
 * DO NOT REMOVE OR STRIP COMMENTS FROM THIS FILE.
 * ----------------------------------------------------------------------------------------------------
 * MASTER SYSTEM ARCHITECTURE & OPERATIONAL SPECIFICATION:
 * 
 * ----------------------------------------------------------------------------------------------------
 * PART I: INTEGRATION & COMMAND CONTRACTS
 * ----------------------------------------------------------------------------------------------------
 * 1. UNIFIED COMMAND & MENU NAMES (EXACT MATCHING REQUIRED):
 *    - Commands triggered via AppSheet, the Google Sheets UI menu (`onOpen`), or the `System_Control` sheet 
 *      dropdown (Row 2, Column B) MUST match the exact canonical action names handled in `handleExternalChange`:
 *      * "Optimize All"
 *      * "Reset & Recalculate All Stocks"
 *      * "Sort Transactions"
 *      * "Sort Google Finance"
 *      * "Add Sort facility to pivot"
 *      * "Extend Format"
 *      * "List All Formulas"
 *      * "List All Tables"
 * 
 * 2. STRICT CANONICAL TABLE HEADERS (NO ALIASES):
 *    - All header strings in `ARRAY_FORMULAS` and `col()` lookups MUST match sheet table headers character-for-character.
 *    - Distinct Canonical Header Names:
 *      * `Stocks` sheet:     "Avg Cost", "Fut.Dividend %", "Tot. Liquidity Value", "Threeshold Quantity", "Total Protofolio P/L"
 *      * `platforms` sheet:    "Fut.Div %", "Tot. Liq. Value", "Div %"
 *      * `Transactions`:     "Avg. Cost"
 *    - Exact naming ensures `BatchProcessor` writes `null` to live `ARRAYFORMULA` columns, preventing `#REF!` errors.
 * 
 * 3. SELF-CONTAINED INTEGRATION & CASCADE LOOP GUARD:
 *    - Bound Script: Runs 100% inside the Google Spreadsheet instance without external servers or web services.
 *    - Cascade Loop Guard: `handleExternalChange` checks `getActiveSheet().getName() === "System_Control"`. Edits on any
 *      other sheet (`Transactions`, `Stocks`) exit immediately to prevent trigger recursion.
 *    - Control Panel State (`System_Control` Row 2): `runWithStatus()` sets Col C ("Status") to "Running" during execution,
 *      and atomically resets Col B ("Run") to "- Select Action -" and Col C to "Not Running" upon completion.
 * 
 * ----------------------------------------------------------------------------------------------------
 * PART II: FINANCIAL CALCULATION & EXECUTION RULES
 * ----------------------------------------------------------------------------------------------------
 * 4. POSITION & AVERAGE COST RULES:
 *    - Running Shares (`p.qty`): Total executed shares owned (`Buys - Sells`).
 *    - Unit Average Cost (`p.avg`): `Total Invested Cost / Total Shares`.
 *      * BUY / GIFT: Increases shares and invested capital, recalculating unit `p.avg`.
 *      * SELL: Proportionally reduces shares and invested capital. Selling NEVER alters unit `p.avg`.
 *      * SOLD OUT (`qty = 0`): Capital drops to 0, but the last unit `p.avg` is retained in memory & on `Stocks`.
 *    - Cost Selection Rule (`Stocks` sheet): Uses `normalAvgCost` (executed on/before today) unless null
 *      (when 100% of buys are pending future limit orders), in which case it falls back to `futureAvgCost`.
 * 
 * 5. EXECUTION DATE & DIVIDEND RULES:
 *    - Execution Date = Distribution / Payment Date (تاريخ التوزيع).
 *    - Transaction Date = Deserved / Eligibility Date (تاريخ الاستحقاق).
 *    - Execution Guard (`tExec <= todayLimit`): Pending future orders (`tExec > todayLimit`) do NOT affect
 *      active share balances (`p.qty`) or current average costs today.
 *    - Undeserved Dividends (Transaction Date > Today): Quantities auto-sync with active share balance.
 *    - Deserved / Past Dividends (Transaction Date <= Today): Immutable historical records (quantities locked).
 *    - Multi-platform Dividend Replication: Auto-creates dividend rows for all platforms holding shares on the
 *      Deserved Date (Transaction Date), assigning each new row a sequential primary key (`maxSerial + 1`).
 * 
 * ----------------------------------------------------------------------------------------------------
 * PART III: SURGICAL BATCHING & AUDIT LOGGING
 * ----------------------------------------------------------------------------------------------------
 * 6. SURGICAL BATCHING & RECALCULATION:
 *    - Normal Runs: Only recalculates stocks where `Last Update > Optimize Date` (or `Optimize Date` is empty).
 *    - Full Recalculation: `resetAndRecalculateAll()` clears `Optimize Date` across `Transactions` to force a 100% pass.
 *    - Consolidated Audit Logging: All operations write to `Optimize_Log` in a single pass (legacy `StockAudit` is unused).
 * ==================================================================================================== */

  
  /**
   * Static Pre-compiled Date Regex for fast O(1) date key parsing.
   * @private
   */
  static get _DATE_REGEX() {
    if (!PortfolioManager._DATE_REGEX_CACHE) {
      PortfolioManager._DATE_REGEX_CACHE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/;
    }
    return PortfolioManager._DATE_REGEX_CACHE;
  }

  /**
   * Canonical list of ArrayFormulas per sheet (Strict exact character matches, no aliases).
   * Prevents writing over Google Sheets formula columns during batch write operations.
   */
  static get ARRAY_FORMULAS() {
    if (!PortfolioManager._ARRAY_FORMULAS_CACHE) {
      PortfolioManager._ARRAY_FORMULAS_CACHE = {
        "Google_Finance_Data": [
          "Combined Name", "Google Name", "Google Min", "Google Price", "Year Min", "Current Price"
        ],
        "Transactions": [
          "Exchange", "Company", "Value + Fee", "Total Value", "Total Value + Fee", 
          "P/L", "Total P/L", "Percentage"
        ],
        "Stocks": [
          "Exchange", "Company", "Combined Name", "Field", "Quantity", "Year Min", 
          "Market Value", "Liquidity Value", "Tot.Mkt.Value", "Tot. Liquidity Value", 
          "Total Cost", "Min + 0.5%", "Market +0.5%", "Mkt - Min", "Mkt - Min %", 
          "Cost - Min", "Cost - Min%", "Threeshold Quantity", "Quantity To Sell", 
          "Safe Sell Price", "Quantity Of Min+0.5%", "Cash Needed", "System_Control Quantity", 
          "Holding", "Selling P/L", "Dividend", "Dividend %", "Fut.Dividend", 
          "Fut.Dividend %", "Share P/L", "Tot. Shares P/L", "Shares P/L %", 
          "Total Protofolio P/L", "Total P/L", "Total P/L %", "10% Profit", "10% Loss"
        ],
        "platforms": [
          "Active Stocks", "Inactive Stocks", "Cost Portfolio", "Market Portfolio", 
          "Liquidity Portfolio", "Holding", "Total Cash", "Total Cost", "Tot.Mkt.Value", 
          "Tot. Liq. Value", "Selling P/L", "Dividend", "Div %", "Fut.Dividend", 
          "Fut.Div %", "Tot. Shares P/L", "Portfolio P/L", "Total P/L", "Total P/L %"
        ],
        "System_Control": [],
        "Optimize_Log": []
      };
    }
    return PortfolioManager._ARRAY_FORMULAS_CACHE;
  }

  /**
   * Safely retrieves array formula headers for a given sheet.
   */
  static getArrayFormulas(sheetName) {
    if (typeof SHEET_ARRAY_FORMULAS !== "undefined" && SHEET_ARRAY_FORMULAS[sheetName]) {
      return SHEET_ARRAY_FORMULAS[sheetName];
    }
    return PortfolioManager.ARRAY_FORMULAS[sheetName] || [];
  }

  /**
   * Fast 2-digit padding helper (e.g., 5 -> "05").
   */
  static _pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  /**
   * Formats Date objects into clean YYYY-MM-DD string format.
   */
  static _fmtDate(d) {
    return `${d.getFullYear()}-${PortfolioManager._pad2(d.getMonth() + 1)}-${PortfolioManager._pad2(d.getDate())}`;
  }

  /**
   * Converts any date input into millisecond timestamp for fast numeric comparisons.
   */
  static getTime(val) {
    if (!val) return 0;
    if (typeof val === "number") return val;
    if (val instanceof Date) return val.getTime();
    const d = new Date(val);
    return isNaN(d.getTime()) ? 0 : d.getTime();
  }

  /**
   * Timezone-safe local date key formatter (YYYY-MM-DD).
   * Prevents UTC date shifting when reading Google Sheets dates.
   */
  static formatDateKey(val) {
    if (!val) return "";
    if (val instanceof Date) {
      return isNaN(val.getTime()) ? "" : PortfolioManager._fmtDate(val);
    }
    const str = String(val).trim();
    const match = str.match(PortfolioManager._DATE_REGEX);
    if (match) {
      return `${match[1]}-${PortfolioManager._pad2(match[2])}-${PortfolioManager._pad2(match[3])}`;
    }
    const d = new Date(str);
    return isNaN(d.getTime()) ? "" : PortfolioManager._fmtDate(d);
  }

  /**
   * Helper: Selects effective average cost for a stock position.
   */
  static getEffectiveCost(item) {
    if (!item) return 0;
    return item.normalAvgCost !== null ? item.normalAvgCost : item.futureAvgCost;
  }

  /**
   * Dynamic Transaction Fee Engine for Saudi Arabia and US markets.
   */
  static calculateFee(market, platform, orderType, quantity, price) {
    const q = typeof quantity === "number" ? Math.abs(quantity) : (Math.abs(parseFloat(quantity)) || 0);
    const p = typeof price === "number" ? Math.abs(price) : (Math.abs(parseFloat(price)) || 0);

    if (q === 0 || p === 0) return 0;

    const tv = q * p;
    const o = String(orderType || "").trim().toLowerCase();
    const m = String(market || "").trim().toLowerCase();

    const isSaudi = (m === "saudi") || (m === "ksa") || (m === "sa");
    const isUs = (m === "us") || (m === "usa") || (m === "american");
    const isSell = (o === "sell");
    const isDiv = (o === "dividend");
    const sign = (isSell || isDiv) ? -1 : 1;
    const round2 = (val) => Math.round(val * 100) / 100;

    // Dividend Tax: US 30% withholding tax, Saudi 0%
    if (isDiv) {
      if (isSaudi) return 0;
      if (isUs) return round2(tv * 0.3) * sign;
      return 0;
    }

    const b = String(platform || "").toLowerCase().replace(/\s/g, "");
    const isSahm = b.includes("sahm");
    const isAlrajhi = b.includes("alrajhi");

    // Saudi Market: Tadawul + CMA + platform commission + 15% VAT
    if (isSaudi) {
      const cmaFee = round2(tv * 0.0003);
      const tadFees = round2(tv * 0.00009) + 2 * round2(tv * 0.00005) + round2(tv * 0.00001);
      const brkRate = isSahm ? 0 : (isAlrajhi ? 0.00105 : 0.001009);
      const vatableFees = tadFees + round2(tv * brkRate);
      const vat = round2(vatableFees * 0.15);
      return round2(cmaFee + vatableFees + vat) * sign;
    }

    // US Market: SEC + FINRA TAF + CAT + platform commission + VAT
    if (isUs) {
      const isAbyan = b.includes("abyan");
      let usplatform = 0;
      if (isSahm) {
        const baseFee = p <= 5 ? 0.49 : 1.99;
        usplatform = Math.max(baseFee, Math.min(q * 0.015, tv * 0.015));
      } else if (isAbyan) {
        usplatform = Math.max(0.95, q * 0.0075);
      } else if (isAlrajhi) {
        usplatform = Math.max(1.99, q * 0.0199);
      } else {
        return 0;
      }

      let regFees = 0;
      if (isSahm) {
        const settlement = Math.max(q * 0.003, 0.01);
        const catFee = Math.max(q * 0.000027, 0.01);
        const taf = isSell ? Math.min(Math.max(q * 0.000166, 0.01), 8.30) : 0;
        regFees = settlement + catFee + taf;
      } else {
        regFees = (isSell && !isAbyan) ? (tv * 0.0000278 + Math.min(Math.max(q * 0.000195, 0.01), 5.95)) : 0;
      }

      return round2(usplatform * 1.15 + regFees) * sign;
    }

    return 0;
  }

  constructor(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
    this.ss = spreadsheet;
  }

  getSheet(name) {
    return this.ss.getSheetByName(name);
  }

  _formatSheets(sheetNames, benchmark) {
    for (const name of sheetNames) {
      const sh = this.getSheet(name);
      if (sh) {
        if (benchmark) {
          benchmark.time(`Extend Format [${name}]`, () => TableFormatter.extend(sh));
        } else {
          TableFormatter.extend(sh);
        }
      }
    }
  }

  _validateHeaders(sheet, headerRowIndex) {
    if (!sheet || sheet.getLastColumn() === 0) return false;
    const headerValues = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nonEntries = headerValues.filter(h => String(h || "").trim() !== "").length;
    if (nonEntries < 3) {
      Logger.log(`⚠️ Header validation failed for sheet "${sheet.getName()}": Header row appears empty.`);
      this.ss.toast(`Header row in sheet "${sheet.getName()}" is empty. Aborting to protect sheet.`, "Header Protection", 8);
      return false;
    }
    return true;
  }

  _writeLogSheet(logEntries, timestamp) {
    const headers = ["Timestamp", "Sheet Name", "Action", "Order", "platform", "Symbol", "Stock Name", "Order Type", "Details"];

    let logSheet = this.getSheet("Optimize_Log");
    if (!logSheet) {
      logSheet = this.ss.insertSheet("Optimize_Log");
    }

    logSheet.clear();

    if (logEntries.length === 0) {
      logEntries.push([
        timestamp, "Portfolio", "UPDATE", "N/A", "N/A", "N/A", "N/A", "N/A",
        "No changes were required. Portfolio is up to date."
      ]);
    }

    const fullData = [headers, ...logEntries];
    logSheet.getRange(1, 1, fullData.length, headers.length).setValues(fullData);

    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#3c78d8")
      .setFontColor("#ffffff");

    logSheet.getRange(2, 1, logEntries.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidths(1, 8, 125);
    logSheet.setColumnWidth(9, 420);
  }

/**
   * Unified execution wrapper with live status and time tracking.
   * (Self-blocking guard removed to allow trigger execution).
   */
  runWithStatus(actionFn) {
    const sheet = this.getSheet("System_Control");
    if (!sheet || !this._validateHeaders(sheet, 1)) {
      actionFn();
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toLowerCase());
    const runCol = headers.indexOf("run") + 1 || 2;
    const statusCol = headers.indexOf("status") + 1 || 3;
    const dateCol = headers.indexOf("last update") + 1 || 4;
    const errorCol = headers.indexOf("error") + 1 || 5;

    const tz = this.ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const getTimestamp = () => Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    // 1. Mark status as Running
    sheet.getRange(2, statusCol).setValue("Running");
    sheet.getRange(2, dateCol).setValue(getTimestamp());
    sheet.getRange(2, errorCol).clearContent();
    SpreadsheetApp.flush();

    let errorOccurred = null;

    try {
      actionFn();
    } catch (e) {
      errorOccurred = e.message || String(e);
      Logger.log(`❌ Execution Error: ${e.stack || e.message}`);
      throw e;
    } finally {
      // 2. Mark Completed/Error and reset B2
      sheet.getRange(2, runCol).setValue("- Select Action -");
      sheet.getRange(2, statusCol).setValue(errorOccurred ? "Error" : "Completed");
      sheet.getRange(2, dateCol).setValue(getTimestamp());

      if (errorOccurred) {
        sheet.getRange(2, errorCol).setValue(errorOccurred);
      } else {
        sheet.getRange(2, errorCol).clearContent();
      }

      SpreadsheetApp.flush();
    }
  }

  /**
   * Sorts `Google_Finance_Data` sheet.
   * Order: Exchange Category -> Active Status -> Symbol Name.
   */
  sortGoogleFinance(logEntries = null) {
    const sheet = this.getSheet("Google_Finance_Data");
    if (!sheet || !this._validateHeaders(sheet, 2)) return;

    let sortedData = null;
    const nowTimestamp = new Date();

    try {
      BatchProcessor.process(
        sheet,
        2,
        (row, col, index, rawData) => {
          if (index === 0) {
            const googleMinIdx = col("Google Min");
            const exchangeIdx = col("Exchange");
            const combinedNameIdx = col("Combined Name");
            const symbolIdx = col("Symbol"); // 👈 Symbol lookup
            const orderIdx = col("Order");

            const numRows = rawData.length;
            const decorated = new Array(numRows);

            for (let i = 0; i < numRows; i++) {
              const r = rawData[i].slice();
              const ex = String(r[exchangeIdx] || "").trim().toUpperCase();
              
              const rank = (ex === "SA" || ex === "KSA" || ex === "SAUDI") ? 0 :
                           (ex === "US" || ex === "USA" || ex === "AMERICAN") ? 1 :
                           (ex === "GOLD" || ex === "XAU") ? 2 : 3;
              
              const minVal = r[googleMinIdx];
              const active = (minVal !== "" && minVal !== null && minVal !== undefined) ? 1 : 0;
              const name = String(r[combinedNameIdx] || "").toLowerCase();
              
              decorated[i] = { row: r, rank, active, name };
            }

            decorated.sort((a, b) => {
              if (a.rank !== b.rank) return a.rank - b.rank;
              if (a.active !== b.active) return a.active - b.active;
              return a.name.localeCompare(b.name);
            });

            sortedData = new Array(numRows);
            const standaloneLogs = logEntries || [];

            for (let i = 0; i < numRows; i++) {
              const r = decorated[i].row;
              const newOrder = i + 1;
              const oldOrder = orderIdx !== -1 ? (Number(r[orderIdx]) || newOrder) : newOrder;

              if (orderIdx !== -1) r[orderIdx] = newOrder;
              sortedData[i] = r;

              if (oldOrder !== newOrder) {
                const combinedName = combinedNameIdx !== -1 ? String(r[combinedNameIdx] || "").trim() : "";
                const symbol = symbolIdx !== -1 ? String(r[symbolIdx] || "").trim() : "N/A"; // 👈 System_Controlct Symbol

                standaloneLogs.push([
                  nowTimestamp,
                  "Google_Finance_Data",
                  "UPDATE",
                  newOrder,
                  "N/A",
                  symbol, // 👈 Put Symbol in log
                  combinedName,
                  "N/A",
                  `Order changed from ${oldOrder} to ${newOrder}`
                ]);
              }
            }

            if (!logEntries && standaloneLogs.length > 0) {
              this._writeLogSheet(standaloneLogs, nowTimestamp);
            }
          }

          return [sortedData[index]];
        },
        PortfolioManager.getArrayFormulas("Google_Finance_Data")
      );

      this.ss.toast("Google_Finance_Data sorted and re-sequenced.", "Success", 3);
    } catch (e) {
      this.ss.toast(e.message, "Sort Error", 10);
    }
  }

  /**
   * Sorts `Transactions` sheet chronologically.
   * Multi-tier Sort: Execution Date -> Transaction Date -> Last Update -> Company -> Order Type -> Avg Cost
   */
  sortTransactions(logEntries = null) {
    const sheet = this.getSheet("Transactions");
    if (!sheet || !this._validateHeaders(sheet, 2)) return;

    let sortedData = null;
    const nowTimestamp = new Date();

    try {
      BatchProcessor.process(
        sheet,
        2,
        (row, col, index, rawData) => {
          if (index === 0) {
            const execDateIdx = col("Execution Date");
            const transDateIdx = col("Transaction Date");
            const companyIdx = col("Company");
            const orderTypeIdx = col("Order Type");
            const avgCostIdx = col("Avg. Cost");
            const orderIdx = col("Order");
            const platformIdx = col("platform");
            const symbolIdx = col("Symbol");
            const lastUpdateIdx = col("Last Update");

            const numRows = rawData.length;
            const decorated = new Array(numRows);

            for (let i = 0; i < numRows; i++) {
              const r = rawData[i].slice();
              decorated[i] = {
                row: r,
                tExec: PortfolioManager.getTime(r[execDateIdx]),
                tTrans: PortfolioManager.getTime(r[transDateIdx]),
                tUpdate: lastUpdateIdx !== -1 ? PortfolioManager.getTime(r[lastUpdateIdx]) : 0,
                comp: String(r[companyIdx] || "").toLowerCase(),
                type: String(r[orderTypeIdx] || "").toLowerCase(),
                cost: Number(r[avgCostIdx]) || 0
              };
            }

            decorated.sort((a, b) => {
              if (a.tExec !== b.tExec) return a.tExec - b.tExec;
              if (a.tTrans !== b.tTrans) return a.tTrans - b.tTrans;
              if (a.tUpdate !== b.tUpdate) return a.tUpdate - b.tUpdate;
              const cComp = a.comp.localeCompare(b.comp);
              if (cComp !== 0) return cComp;
              const cType = a.type.localeCompare(b.type);
              if (cType !== 0) return cType;
              return b.cost - a.cost;
            });

            sortedData = new Array(numRows);
            const standaloneLogs = logEntries || [];

            for (let i = 0; i < numRows; i++) {
              const r = decorated[i].row;
              const newOrder = i + 1;
              const oldOrder = orderIdx !== -1 ? (Number(r[orderIdx]) || newOrder) : newOrder;

              if (orderIdx !== -1) r[orderIdx] = newOrder;
              sortedData[i] = r;

              if (oldOrder !== newOrder) {
                const b = platformIdx !== -1 ? String(r[platformIdx] || "").trim() : "N/A";
                const s = symbolIdx !== -1 ? String(r[symbolIdx] || "").trim() : "N/A";
                const c = companyIdx !== -1 ? String(r[companyIdx] || "").trim() : "N/A";
                const t = orderTypeIdx !== -1 ? String(r[orderTypeIdx] || "").trim() : "N/A";

                standaloneLogs.push([
                  nowTimestamp,
                  "Transactions",
                  "UPDATE",
                  newOrder,
                  b,
                  s,
                  c,
                  t,
                  `Order changed from ${oldOrder} to ${newOrder}`
                ]);
              }
            }

            if (!logEntries && standaloneLogs.length > 0) {
              this._writeLogSheet(standaloneLogs, nowTimestamp);
            }
          }

          return [sortedData[index]];
        },
        PortfolioManager.getArrayFormulas("Transactions")
      );

      this.ss.toast("Transactions sorted and re-sequenced.", "Success", 3);
    } catch (e) {
      this.ss.toast(e.message, "Sort Error", 10);
    }
  }

  _optimizeAllCore() {
    const benchmark = new Benchmark("Optimize All Pipeline");

    const dirtyKeys = new Set();
    const costDataMap = new Map();
    const nowTimestamp = new Date();
    const logEntries = [];

    // =========================================================================
    // STAGE 1: SORT GOOGLE FINANCE DATA SHEET
    // =========================================================================
    benchmark.time("Sort Google Finance Data", () => this.sortGoogleFinance(logEntries));

    const transSheet = this.getSheet("Transactions");
    if (!transSheet || !this._validateHeaders(transSheet, 2)) return;

    // =========================================================================
    // STAGE 2: PROCESS TRANSACTIONS PASS
    // =========================================================================
    benchmark.time("Process & Write Transactions Pass", () => {
      let sortedData = null;
      const newDividendRows = [];

      BatchProcessor.process(
        transSheet,
        2,
        (row, col, index, rawData) => {
          if (index === 0) {
            const serialIdx = col("Serial");
            const execDateIdx = col("Execution Date");
            const transDateIdx = col("Transaction Date");
            const companyIdx = col("Company");
            const orderTypeIdx = col("Order Type");
            const avgCostIdx = col("Avg. Cost");
            const orderIdx = col("Order");
            const platformIdx = col("platform");
            const symbolIdx = col("Symbol");
            const qtyIdx = col("Quantity");
            const priceIdx = col("Value");
            const marketIdx = col("Exchange");
            const feeIdx = col("Fee");
            const totalFeeIdx = col("Total Fee");
            const lastUpdateIdx = col("Last Update");
            const optDateIdx = col("Optimize Date");

            const numRows = rawData.length;
            const decorated = new Array(numRows);
            const useSurgicalBatch = (lastUpdateIdx !== -1 && optDateIdx !== -1);
            const existingDivKeys = new Set();
            let maxSerial = 0;

            for (let i = 0; i < numRows; i++) {
              const r = rawData[i].slice();
              
              if (serialIdx !== -1) {
                const sNum = Number(r[serialIdx]) || 0;
                if (sNum > maxSerial) maxSerial = sNum;
              }

              const b = String(r[platformIdx] || "").trim();
              const s = String(r[symbolIdx] || "").trim();
              const cName = companyIdx !== -1 ? String(r[companyIdx] || "").trim() : "";
              const bLower = b.toLowerCase();
              const sLower = s.toLowerCase();
              const rawTypeStr = String(r[orderTypeIdx] || "").trim();
              const typeLower = rawTypeStr.toLowerCase();
              const keyLower = (bLower && sLower) ? `${bLower}|${sLower}` : "";
              const transDateStr = PortfolioManager.formatDateKey(r[transDateIdx]);

              if (keyLower && typeLower === "dividend" && transDateStr) {
                existingDivKeys.add(`${keyLower}|${transDateStr}`);
              }

              if (keyLower) {
                if (useSurgicalBatch) {
                  if (!dirtyKeys.has(keyLower)) {
                    const modVal = r[lastUpdateIdx];
                    const optVal = r[optDateIdx];
                    const modTime = modVal ? PortfolioManager.getTime(modVal) : 0;
                    const optTime = optVal ? PortfolioManager.getTime(optVal) : 0;
                    if (modTime > optTime || !optVal) dirtyKeys.add(keyLower);
                  }
                } else {
                  dirtyKeys.add(keyLower);
                }
              }

              const q = Math.abs(parseFloat(r[qtyIdx])) || 0;
              const p = Math.abs(parseFloat(r[priceIdx])) || 0;

              decorated[i] = {
                row: r,
                platform: b,
                symbol: s,
                companyName: cName,
                bLower: bLower,
                sLower: sLower,
                keyLower: keyLower,
                market: r[marketIdx],
                type: typeLower,
                rawType: rawTypeStr || r[orderTypeIdx],
                qty: q,
                price: p,
                tExec: PortfolioManager.getTime(r[execDateIdx]),
                tTrans: PortfolioManager.getTime(r[transDateIdx]),
                tUpdate: lastUpdateIdx !== -1 ? PortfolioManager.getTime(r[lastUpdateIdx]) : 0,
                transDateStr: transDateStr,
                comp: cName.toLowerCase(),
                cost: Number(r[avgCostIdx]) || 0
              };
            }

            decorated.sort((a, b) => {
              if (a.tExec !== b.tExec) return a.tExec - b.tExec;
              if (a.tTrans !== b.tTrans) return a.tTrans - b.tTrans;             
              const cComp = a.comp.localeCompare(b.comp);
              if (cComp !== 0) return cComp;
              const cType = a.type.localeCompare(b.type);
              if (cType !== 0) return cType;
              if (a.tUpdate !== b.tUpdate) return a.tUpdate - b.tUpdate;
              return b.cost - a.cost;
            });

            const portfolioMap = {};
            const symbolplatformsMap = new Map();
            const todayLimit = new Date().setHours(23, 59, 59, 999);
            const todayStartMs = new Date().setHours(0, 0, 0, 0);

            sortedData = new Array(numRows);

            for (let i = 0; i < numRows; i++) {
              const item = decorated[i];
              const rowItem = item.row;

              const newOrder = i + 1;
              const oldOrder = orderIdx !== -1 ? (Number(rowItem[orderIdx]) || newOrder) : newOrder;

              if (orderIdx !== -1) rowItem[orderIdx] = newOrder;

              if (oldOrder !== newOrder) {
                logEntries.push([
                  nowTimestamp,
                  "Transactions",
                  "UPDATE",
                  newOrder,
                  item.platform || "N/A",
                  item.symbol || "N/A",
                  item.companyName || "N/A",
                  item.rawType || "N/A",
                  `Order changed from ${oldOrder} to ${newOrder}`
                ]);
              }

              const platform = item.platform;
              const symbol = item.symbol;
              const keyLower = item.keyLower;

              if (!keyLower) {
                sortedData[i] = rowItem;
                continue;
              }

              const type = item.type;
              const qtyVal = item.qty;
              const priceVal = item.price;

              let p = portfolioMap[keyLower];
              if (!p) {
                p = { qty: 0, cost: 0, avg: 0 };
                portfolioMap[keyLower] = p;
              }

              const totalFee = PortfolioManager.calculateFee(item.market, platform, item.rawType, qtyVal, priceVal);
              const singleFee = qtyVal > 0 ? Math.round((totalFee / qtyVal) * 10000) / 10000 : 0;
              const totalValFee = Math.round(((qtyVal * priceVal) + totalFee) * 100) / 100;

              const isExecuted = (item.tExec <= todayLimit);

              if (type === "buy" || type === "gift") {
                if (isExecuted) {
                  if (p.qty <= 0.000001) {
                    p.qty = 0;
                    p.cost = 0;
                  }
                  p.qty += qtyVal;
                  p.cost += Math.abs(totalValFee);
                  if (p.qty > 0) p.avg = p.cost / p.qty;
                }

                const cost = Math.round(p.avg * 10000) / 10000;
                let mapItem = costDataMap.get(keyLower);
                if (!mapItem) {
                  mapItem = { normalAvgCost: null, futureAvgCost: null, origplatform: platform, origSymbol: symbol, origCompany: item.companyName, qty: p.qty };
                  costDataMap.set(keyLower, mapItem);
                } else if (!mapItem.origCompany && item.companyName) {
                  mapItem.origCompany = item.companyName;
                }
                
                mapItem.futureAvgCost = cost;
                if (isExecuted) mapItem.normalAvgCost = cost; 
                mapItem.qty = p.qty;

              } else if (type === "sell") {
                if (isExecuted) {
                  if (p.qty > 0) {
                    const sellQty = Math.min(qtyVal, p.qty);
                    p.cost -= (sellQty * p.avg);
                    p.qty -= qtyVal;
                  }
                  if (p.qty <= 0.000001) {
                    p.qty = 0;
                    p.cost = 0;
                  }
                  let mapItem = costDataMap.get(keyLower);
                  if (mapItem) {
                    mapItem.qty = p.qty;
                  }
                }

              } else if (type === "dividend") {
                // UNDESERVED DIVIDENDS: Deserved Date (Transaction Date) is in the future
                if (item.tTrans > todayLimit) {
                  const targetQty = Math.max(0, p.qty);
                  const oldQty = Number(rowItem[qtyIdx]) || 0;
                  
                  if (oldQty !== targetQty) {
                    rowItem[qtyIdx] = targetQty;
                    if (lastUpdateIdx !== -1) rowItem[lastUpdateIdx] = nowTimestamp;
                    
                    logEntries.push([
                      nowTimestamp,
                      "Transactions",
                      "UPDATE",
                      newOrder,
                      platform,
                      symbol,
                      item.companyName,
                      item.rawType || "Dividend",
                      `Updated Dividend Quantity from ${oldQty} to ${targetQty} (Deserved Date: ${item.transDateStr})`
                    ]);
                  }
                }

                // MULTI-platform DIVIDEND REPLICATION (Check holdings on Deserved Date)
                if (item.tTrans >= todayStartMs && symbol && item.transDateStr) {
                  const targetplatforms = symbolplatformsMap.get(item.sLower);
                  if (targetplatforms) {
                    for (const targetplatformName of targetplatforms) {
                      const targetBLower = targetplatformName.toLowerCase();
                      if (targetBLower === item.bLower) continue;

                      const targetKeyLower = `${targetBLower}|${item.sLower}`;
                      const targetP = portfolioMap[targetKeyLower];

                      if (targetP && targetP.qty > 0) {
                        const targetDivKey = `${targetKeyLower}|${item.transDateStr}`;

                        if (!existingDivKeys.has(targetDivKey)) {
                          existingDivKeys.add(targetDivKey);

                          maxSerial++;
                          const cloneRow = rowItem.slice();
                          const repOrderNum = numRows + newDividendRows.length + 1;

                          if (serialIdx !== -1) cloneRow[serialIdx] = maxSerial;
                          if (orderIdx !== -1) cloneRow[orderIdx] = repOrderNum; 
                          if (platformIdx !== -1) cloneRow[platformIdx] = targetplatformName;
                          if (qtyIdx !== -1) cloneRow[qtyIdx] = targetP.qty;

                          const targetTotalFee = PortfolioManager.calculateFee(item.market, targetplatformName, item.rawType, targetP.qty, priceVal);
                          const targetSingleFee = targetP.qty > 0 ? Math.round((targetTotalFee / targetP.qty) * 10000) / 10000 : 0;

                          if (feeIdx !== -1) cloneRow[feeIdx] = targetSingleFee;
                          if (totalFeeIdx !== -1) cloneRow[totalFeeIdx] = targetTotalFee;
                          if (avgCostIdx !== -1) cloneRow[avgCostIdx] = Math.round(targetP.avg * 10000) / 10000;
                          if (lastUpdateIdx !== -1) cloneRow[lastUpdateIdx] = nowTimestamp;
                          if (optDateIdx !== -1) cloneRow[optDateIdx] = nowTimestamp;

                          newDividendRows.push(cloneRow);
                          dirtyKeys.add(targetKeyLower);

                          logEntries.push([
                            nowTimestamp,
                            "Transactions",
                            "INSERT",
                            repOrderNum,
                            targetplatformName,
                            symbol,
                            item.companyName,
                            item.rawType || "Dividend",
                            `Replicated Dividend created (Serial: ${maxSerial}, Qty: ${targetP.qty}, Deserved Date: ${item.transDateStr}, Fee: ${targetTotalFee})`
                          ]);
                        }
                      }
                    }
                  }
                }
              }

              let bSet = symbolplatformsMap.get(item.sLower);
              if (!bSet) {
                bSet = new Set();
                symbolplatformsMap.set(item.sLower, bSet);
              }
              if (p.qty > 0) {
                bSet.add(platform);
              } else {
                bSet.delete(platform);
              }

              const oldFee = Number(rowItem[feeIdx]) || 0;
              const oldTotalFee = Number(rowItem[totalFeeIdx]) || 0;
              const oldAvg = item.cost;
              const oldOptDate = optDateIdx !== -1 ? rowItem[optDateIdx] : null;

              if (feeIdx !== -1) rowItem[feeIdx] = singleFee;
              if (totalFeeIdx !== -1) rowItem[totalFeeIdx] = totalFee;

              const newAvg = Math.round(p.avg * 10000) / 10000;
              if (avgCostIdx !== -1) rowItem[avgCostIdx] = newAvg;

              const feeChanged = Math.abs(oldFee - singleFee) > 0.0001;
              const totalFeeChanged = Math.abs(oldTotalFee - totalFee) > 0.01;
              const avgChanged = Math.abs(oldAvg - newAvg) > 0.0001;
              const optDateMissing = !oldOptDate;

              if (dirtyKeys.has(keyLower) || feeChanged || totalFeeChanged || avgChanged || optDateMissing) {
                dirtyKeys.add(keyLower);

                if (optDateIdx !== -1) {
                  rowItem[optDateIdx] = nowTimestamp;
                }

                if (feeChanged || totalFeeChanged || avgChanged) {
                  if (lastUpdateIdx !== -1) rowItem[lastUpdateIdx] = nowTimestamp;

                  const changes = [];
                  if (feeChanged || totalFeeChanged) changes.push(`Fee: ${oldFee} -> ${singleFee}`);
                  if (avgChanged) changes.push(`Avg Cost: ${oldAvg} -> ${newAvg}`);

                  logEntries.push([
                    nowTimestamp,
                    "Transactions",
                    "UPDATE",
                    newOrder,
                    platform,
                    symbol,
                    item.companyName,
                    item.rawType || "N/A",
                    `Transaction updated (${changes.join(", ")})`
                  ]);
                }
              }

              sortedData[i] = rowItem;
            }
          }

          if (index === rawData.length - 1 && newDividendRows.length > 0) {
            return [sortedData[index], ...newDividendRows];
          }

          return [sortedData[index]];
        },
        PortfolioManager.getArrayFormulas("Transactions")
      );
    });

    // =========================================================================
    // STAGE 3: PROCESS STOCKS SHEET PASS
    // =========================================================================
    const stocksSheet = this.getSheet("Stocks");
    if (stocksSheet && this._validateHeaders(stocksSheet, 2)) {
      benchmark.time("Process & Write Stocks Pass", () => {
        let matchedKeys = new Set();
        let sLastCol = 0;
        let idxplatform = -1, idxSymbol = -1, idxCompany = -1, idxCost = -1, lastUpdateIdx = -1;
        const formulaIndices = [];

        BatchProcessor.process(
          stocksSheet,
          2,
          (row, col, index, rawData) => {
            if (index === 0) {
              idxplatform = col("platform");
              idxSymbol = col("Symbol");
              idxCompany = col("Company");
              idxCost = col("Avg Cost");
              sLastCol = row.length;
              lastUpdateIdx = col("Last Update");

              const formulaHeaders = PortfolioManager.getArrayFormulas("Stocks");
              formulaHeaders.forEach(fh => {
                const fIdx = col(fh);
                if (fIdx !== -1) formulaIndices.push(fIdx);
              });
            }

            const platform = String(row[idxplatform] || "").trim();
            const symbol = String(row[idxSymbol] || "").trim();
            const stockName = idxCompany !== -1 ? String(row[idxCompany] || "").trim() : "";
            const keyLower = (platform && symbol) ? `${platform.toLowerCase()}|${symbol.toLowerCase()}` : null;

            if (keyLower) {
              matchedKeys.add(keyLower);

              if (dirtyKeys.has(keyLower)) {
                const item = costDataMap.get(keyLower);
                if (item) {
                  const targetRaw = PortfolioManager.getEffectiveCost(item);
                  const target2pt = Math.round(targetRaw * 100) / 100;
                  const currentCost = Math.round((Number(row[idxCost]) || 0) * 100) / 100;

                  if (idxplatform !== -1) row[idxplatform] = item.origplatform || platform;
                  if (idxSymbol !== -1) row[idxSymbol] = item.origSymbol || symbol;

                  const costChanged = Math.abs(currentCost - target2pt) > 0.011;

                  if (costChanged) {
                    row[idxCost] = targetRaw;
                    if (lastUpdateIdx !== -1) row[lastUpdateIdx] = nowTimestamp;
                    
                    const qtyNote = item.qty === 0 ? " (Quantity = 0)" : "";
                    logEntries.push([
                      nowTimestamp,
                      "Stocks",
                      "UPDATE",
                      "N/A",
                      item.origplatform,
                      item.origSymbol,
                      item.origCompany || stockName,
                      "N/A",
                      `Avg Cost updated from ${currentCost} to ${target2pt}${qtyNote}`
                    ]);
                  }
                }
              }
            }

            if (index === rawData.length - 1) {
              const newStockRows = [];
              for (const [k, item] of costDataMap.entries()) {
                if (!matchedKeys.has(k)) {
                  const finalCost = PortfolioManager.getEffectiveCost(item);
                  const newRow = new Array(sLastCol).fill("");

                  if (idxplatform !== -1) newRow[idxplatform] = item.origplatform;
                  if (idxSymbol !== -1) newRow[idxSymbol] = item.origSymbol;
                  if (idxCompany !== -1) newRow[idxCompany] = item.origCompany || "";
                  if (idxCost !== -1) newRow[idxCost] = Math.round(finalCost * 10000) / 10000;
                  if (lastUpdateIdx !== -1) newRow[lastUpdateIdx] = nowTimestamp;

                  formulaIndices.forEach(idx => {
                    if (idx >= 0 && idx < sLastCol) newRow[idx] = null;
                  });

                  newStockRows.push(newRow);

                  logEntries.push([
                    nowTimestamp,
                    "Stocks",
                    "INSERT",
                    "N/A",
                    item.origplatform,
                    item.origSymbol,
                    item.origCompany || "",
                    "N/A",
                    `New stock position added with Avg Cost: ${Math.round(finalCost * 10000) / 10000}`
                  ]);
                }
              }

              return [row, ...newStockRows];
            }

            return [row];
          },
          PortfolioManager.getArrayFormulas("Stocks")
        );
      });
    }


    // =======================================================================
    // STAGE 4: SYNCHRONIZE AUDIT TABLES (Tables_Audit & Formula_Audit)
    // =======================================================================
    benchmark.time("Synchronize Audits", () => {
      if (typeof runAllAudits === "function") {
        runAllAudits();
      } else {
        if (typeof listAllTables === "function") listAllTables();
        if (typeof listAllFormulas === "function") listAllFormulas();
      }
    });


    // =========================================================================
    // STAGE 5: WRITE AUDIT LOG SHEET
    // =========================================================================
    benchmark.time("Write Optimize Log Pass", () => {
      this._writeLogSheet(logEntries, nowTimestamp);
    });

    // =========================================================================
    // STAGE 6: SELECTIVE FORMATTING PASS
    // =========================================================================
    this._formatSheets(["Transactions", "Stocks"], benchmark);

    const totalSec = benchmark.report();
    this.ss.toast(`Optimizations complete in ${totalSec}s. Log written to Optimize_Log.`, "Optimize All", 5);
  }

  /**
   * Internal Reset & Recalculate Core.
   */
  _resetAndRecalculateAllCore() {
    const transSheet = this.getSheet("Transactions");
    if (!transSheet) {
      this.ss.toast("Transactions sheet not found.", "Error", 5);
      return;
    }

    const headers = transSheet.getRange(2, 1, 1, transSheet.getLastColumn()).getValues()[0];
    const optDateIdx = headers.indexOf("Optimize Date");

    if (optDateIdx === -1) {
      this.ss.toast("⚠️ 'Optimize Date' column not found in Transactions sheet.", "Error", 5);
      return;
    }

    const lastRow = transSheet.getLastRow();
    if (lastRow > 2) {
      transSheet.getRange(3, optDateIdx + 1, lastRow - 2, 1).clearContent();
      this.ss.toast("Cleared all Optimize Dates. Recalculating full portfolio...", "Reset Complete", 5);
    }

    this._optimizeAllCore();
  }

  optimizeAll() {
    this.runWithStatus(() => this._optimizeAllCore());
  }

  resetAndRecalculateAll() {
    this.runWithStatus(() => this._resetAndRecalculateAllCore());
  }

}
