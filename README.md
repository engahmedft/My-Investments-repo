# My-Investment-repo
## 🌐 System Components & Live Resources

### 1. Database & Schema Layer (Google Sheets)
* **Live Published Sheet:** [View Public Web Spreadsheet (`pubhtml`)](https://docs.google.com/spreadsheets/d/e/2PACX-1vSekkQGc4fEyvY_FM-a_F0JjGi2Q6PD-Z4K6Zl4PTdehPd-agtChOzxXr85zajLKaw9qEiAHLjmnD2j/pubhtml)
* **Table Structures (`Tables_Audit` tab):** Auto-generated inventory capturing all table dimensions, header rows, and column positions via 1D BFS island scanning.
* **Formulas & Dependencies (`Formula_Audit` tab):** Live catalog documenting every static and `ARRAYFORMULA` cell definition across all workbook sheets.

---

### 2. Automation & Execution Engine (Google Apps Script)
* **Codebase Repository:** [GitHub: `engahmedft/My-Investments-repo`](https://github.com/engahmedft/My-Investments-repo)
* **Script Project Editor:** [Google Apps Script Project](https://script.google.com/home/projects/1vDCqZl8wwwBoTFWBT_3BwXUEafaBCXqEFRULKSaY1ovIE7QQ8cY-Ga50/edit)
* **Core Modules:**
  * `PortfolioManager.js` – WAC financial calculations, fee engines (Saudi 15% VAT / US 30% tax), and dividend replication.
  * `BatchProcessor.js` – High-performance 1-Read / 1-Write batch I/O and Zero-Write change detection.
  * `Triggers.js` – Event routers (`onEdit`, `onChange`), `LockService` concurrency control, and `DynamicCommandEngine`.
  * `WorkbookAuditor.js` – ArrayFormula spill unblocker (`fixAllArrayFormulaErrors`) and audit generator.
  * `TableFormatter.js` & `PivotSort.js` – Dynamic grid formatting and matrix pivot-sorting controls.
  * `Benchmark.js` & `DoGet.js` – Monotonic step profiler and zero-permission web viewer.

---

### 3. Frontend & UX Layer (AppSheet)
* **Local Specification File:** `/appsheet/` directory *(Full PDF schema export & Markdown spec)*
