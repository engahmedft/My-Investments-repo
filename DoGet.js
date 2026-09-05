/**
 * Master Codebase Viewer (Light Mode):
 * - Native .toString() extraction (Zero GCP permissions/APIs required)
 * - 1-Click "Copy Entire Project" for AI sharing/backups
 * - Accordion / Collapse toggles + Expand All / Collapse All
 * - ScrollSpy automatic sidebar tracking
 * - Keyboard shortcut ('/' or 'Ctrl+K') for instant search focus
 * - High-readability light theme & large typography
 */
function doGet(e) {
  
/* ====================================================================================================
 * 🤖 CRITICAL DIRECTIVE FOR AI ASSISTANTS & DEVELOPERS:
 * ----------------------------------------------------------------------------------------------------
 * MANIFEST SYNCHRONIZATION CONTRACT:
 * Whenever you modify the project's root `appsscript.json` manifest file (e.g., adding/removing `oauthScopes`, 
 * updating `timeZone`, `dependencies`, or `webapp` deployment settings), you MUST SIMULTANEOUSLY update 
 * the in-memory `manifestJson` object inside this `doGet()` function.
 * 
 * WHY:
 * Google Apps Script does not expose a permission-free local API to read `appsscript.json` dynamically. 
 * To eliminate Google Cloud REST API quotas, 403 permission errors, and GCP project linking requirements, 
 * `doGet()` mirrors `appsscript.json` via the in-memory `manifestJson` object below. 
 * ALWAYS keep `manifestJson` 100% identical to the actual `appsscript.json` file.
 * ==================================================================================================== */

  try {
    // 1. Exact Mirror of appsscript.json (Keep in sync with root manifest)
    const manifestJson = {
      "timeZone": "Asia/Riyadh",
      "dependencies": {},
      "exceptionLogging": "STACKDRIVER",
      "runtimeVersion": "V8",
      "webapp": {
        "executeAs": "USER_DEPLOYING",
        "access": "ANYONE_ANONYMOUS"
      },
      "oauthScopes": [
        "https://www.googleapis.com/auth/script.projects",
        "https://www.googleapis.com/auth/script.projects.readonly",
        "https://www.googleapis.com/auth/script.external_request",
        "https://www.googleapis.com/auth/spreadsheets"
      ]
    };

    // 2. Project Modules & Associated Objects / Functions
    const modules = [
      {
        fileName: "appsscript.json",
        lang: "json",
        description: "Google Apps Script manifest and deployment configuration",
        content: JSON.stringify(manifestJson, null, 2)
      },
      {
        fileName: "Triggers.gs",
        lang: "javascript",
        description: "Dynamic command engine, trigger hooks (onOpen, onEdit, onChange, handleExternalChange) & all public API/UI entry points",
        items: [
          typeof DynamicCommandEngine !== "undefined" ? { name: "DynamicCommandEngine", target: DynamicCommandEngine } : null,
          typeof onOpen !== "undefined" ? onOpen : null,
          typeof handleExternalChange !== "undefined" ? handleExternalChange : null,
          typeof onEdit !== "undefined" ? onEdit : null,
          typeof onChange !== "undefined" ? onChange : null,
          // Batch & Audit Wrappers
          typeof batchRead !== "undefined" ? batchRead : null,
          typeof batchProcessing !== "undefined" ? batchProcessing : null,
          typeof applyAllFormulaAudit !== "undefined" ? applyAllFormulaAudit : null,
          typeof applyAllFormulas !== "undefined" ? applyAllFormulas : null,
          typeof listAllFormulas !== "undefined" ? listAllFormulas : null,
          typeof listAllTables !== "undefined" ? listAllTables : null,
          typeof runAllAudits !== "undefined" ? runAllAudits : null,
          typeof getColLetter !== "undefined" ? getColLetter : null,
          typeof extendFormat !== "undefined" ? extendFormat : null,
          // Portfolio & UI Action Wrappers
          typeof optimizeAll !== "undefined" ? optimizeAll : null,
          typeof resetAndRecalculateAll !== "undefined" ? resetAndRecalculateAll : null,
          typeof sortAndResequenceGoogleFinance !== "undefined" ? sortAndResequenceGoogleFinance : null,
          typeof sortAndResequenceTransactions !== "undefined" ? sortAndResequenceTransactions : null,
          typeof setupPivotSortControls !== "undefined" ? setupPivotSortControls : null,
          typeof addSortFacilityToPivot !== "undefined" ? addSortFacilityToPivot : null
        ]
      },
      {
        fileName: "PortfolioManager.gs",
        lang: "javascript",
        description: "Pure calculation engine, transaction sequencer, and fee engine class",
        items: [
          typeof PortfolioManager !== "undefined" ? PortfolioManager : null
        ]
      },
      {
        fileName: "BatchProcessor.gs",
        lang: "javascript",
        description: "High-performance 1-Read / 1-Write batch engine class with Zero-Write detection",
        items: [
          typeof BatchProcessor !== "undefined" ? BatchProcessor : null
        ]
      },
      {
        fileName: "TableFormatter.gs",
        lang: "javascript",
        description: "Dynamic formatting, bandings, borders & data validation class",
        items: [
          typeof TableFormatter !== "undefined" ? TableFormatter : null
        ]
      },
      {
        fileName: "WorkbookAuditor.gs",
        lang: "javascript",
        description: "Formula audit scanner and 1D BFS table island detection class",
        items: [
          typeof WorkbookAuditor !== "undefined" ? WorkbookAuditor : null
        ]
      },
      {
        fileName: "PivotSort.gs",
        lang: "javascript",
        description: "Pivot table dynamic dropdown controls & sorting manager class",
        items: [
          typeof PivotSortManager !== "undefined" ? PivotSortManager : null
        ]
      },
      {
        fileName: "Benchmark.gs",
        lang: "javascript",
        description: "High-precision monotonic pipeline step timer and performance reporter",
        items: [
          typeof Benchmark !== "undefined" ? Benchmark : null
        ]
      },
      {
        fileName: "DoGet.gs",
        lang: "javascript",
        description: "Self-contained architecture & source code viewer (renders live via .toString())",
        items: [
          typeof doGet !== "undefined" ? doGet : null,
          typeof escapeHtml !== "undefined" ? escapeHtml : null
        ]
      }
    ];

    // 3. Process Code & Pre-compile Whole Project Markdown Bundle
    let navHtml = '';
    let cardsHtml = '';
    let wholeProjectMarkdown = '';

    modules.forEach((mod, idx) => {
      const modId = 'mod_' + idx;
      navHtml += `
        <button class="nav-link" id="btn_${modId}" onclick="scrollToSection('${modId}', this)">
          📄 ${escapeHtml(mod.fileName)}
        </button>
      `;

      let codeText = '';
      if (mod.content) {
        codeText = mod.content;
      } else if (mod.items) {
        mod.items.forEach(item => {
          if (!item) return;

          if (typeof item === 'function') {
            codeText += item.toString() + '\n\n';
          } else if (typeof item === 'object') {
            // Serializes object literals cleanly
            const objName = item.name || 'ConfigObject';
            const target = item.target || item;
            const props = Object.entries(target).map(([k, v]) => {
              if (typeof v === 'function') {
                return `  ${v.toString()}`;
              }
              return `  ${k}: ${JSON.stringify(v, null, 2).replace(/\n/g, '\n  ')}`;
            }).join(',\n\n');
            codeText += `const ${objName} = {\n${props}\n};\n\n`;
          }
        });
      }

      if (!codeText.trim()) {
        codeText = '// Component not loaded or not found in memory.';
      }

      codeText = codeText.trim();
      wholeProjectMarkdown += `### 📄 ${mod.fileName}\n\`\`\`${mod.lang}\n${codeText}\n\`\`\`\n\n`;

      cardsHtml += `
        <div class="file-card" id="${modId}">
          <div class="file-header" onclick="toggleCard('${modId}')">
            <div class="header-left">
              <span class="toggle-icon" id="icon_${modId}">▼</span>
              <div>
                <div class="file-title">📄 ${escapeHtml(mod.fileName)}</div>
                <div class="file-desc">${escapeHtml(mod.description)}</div>
              </div>
            </div>
            <div class="header-actions" onclick="event.stopPropagation()">
              <button class="btn btn-secondary" onclick="copySnippet('${modId}_code', this)">Copy</button>
            </div>
          </div>
          <div class="card-body" id="body_${modId}">
            <pre><code id="${modId}_code" class="language-${mod.lang}">${escapeHtml(codeText)}</code></pre>
          </div>
        </div>
      `;
    });

    // 4. Assemble HTML Output
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>My-Investments Architecture Viewer</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/github.min.css">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js"></script>
        <style>
          * { box-sizing: border-box; }
          html, body { height: 100%; margin: 0; padding: 0; }
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: #f6f8fa;
            color: #1f2328;
            display: flex;
            height: 100vh;
            overflow: hidden;
            font-size: 15px;
          }

          /* Sidebar Styling */
          .sidebar {
            width: 310px;
            background-color: #ffffff;
            border-right: 1px solid #d0d7de;
            display: flex;
            flex-direction: column;
            flex-shrink: 0;
          }
          .sidebar-header {
            padding: 16px 14px 12px 14px;
            border-bottom: 1px solid #eaeef2;
          }
          .search-box {
            width: 100%;
            padding: 9px 12px;
            font-size: 13.5px;
            border: 1px solid #d0d7de;
            border-radius: 6px;
            outline: none;
            background: #f6f8fa;
            transition: all 0.2s;
          }
          .search-box:focus {
            background: #ffffff;
            border-color: #0969da;
            box-shadow: 0 0 0 3px rgba(9, 105, 218, 0.15);
          }
          .sidebar-list {
            padding: 12px 10px;
            overflow-y: auto;
            flex-grow: 1;
          }
          .sidebar-title {
            font-size: 11.5px;
            text-transform: uppercase;
            color: #656d76;
            margin: 4px 0 10px 8px;
            letter-spacing: 0.8px;
            font-weight: 700;
          }
          .nav-link {
            display: block;
            width: 100%;
            text-align: left;
            padding: 9px 12px;
            color: #24292f;
            background: none;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            margin-bottom: 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: pointer;
            transition: all 0.15s ease;
          }
          .nav-link:hover {
            background-color: #f3f4f6;
            color: #0969da;
          }
          .nav-link.active {
            background-color: #ddf4ff;
            color: #0969da;
            font-weight: 600;
          }

          /* Main Content Area */
          .main-content {
            flex-grow: 1;
            padding: 28px 40px;
            overflow-y: auto;
            scroll-behavior: smooth;
            background-color: #f6f8fa;
          }
          .main-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            border-bottom: 1px solid #d0d7de;
            padding-bottom: 16px;
            flex-wrap: wrap;
            gap: 12px;
          }
          .main-header h1 {
            margin: 0 0 4px 0;
            font-size: 24px;
            color: #1f2328;
          }
          .global-actions {
            display: flex;
            gap: 10px;
            align-items: center;
          }

          /* Button Styles */
          .btn {
            border: 1px solid #d0d7de;
            padding: 7px 14px;
            border-radius: 6px;
            font-size: 13.5px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            display: inline-flex;
            align-items: center;
            gap: 6px;
          }
          .btn-primary {
            background: #2da44e;
            color: #ffffff;
            border-color: rgba(27, 31, 36, 0.15);
          }
          .btn-primary:hover {
            background: #2c974b;
          }
          .btn-secondary {
            background: #ffffff;
            color: #24292f;
          }
          .btn-secondary:hover {
            background: #f3f4f6;
            border-color: #8c959f;
          }

          /* File Cards */
          .file-card {
            background: #ffffff;
            border: 1px solid #d0d7de;
            border-radius: 8px;
            margin-bottom: 28px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(31, 35, 40, 0.08);
          }
          .file-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 18px;
            background: #f6f8fa;
            border-bottom: 1px solid #d0d7de;
            cursor: pointer;
            user-select: none;
          }
          .file-header:hover {
            background: #eaeef2;
          }
          .header-left {
            display: flex;
            align-items: center;
            gap: 12px;
          }
          .toggle-icon {
            font-size: 12px;
            color: #656d76;
            transition: transform 0.2s ease;
          }
          .file-title {
            font-weight: 700;
            color: #0969da;
            font-size: 16px;
          }
          .file-desc {
            font-size: 13px;
            color: #656d76;
            margin-top: 2px;
          }

          /* Code Typography */
          pre {
            margin: 0;
            padding: 20px;
            background: #ffffff !important;
            font-size: 15px !important;
            line-height: 1.65 !important;
            overflow-x: auto;
          }
          code {
            font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace !important;
            font-size: 15px !important;
          }
        </style>
      </head>
      <body>
        <div class="sidebar">
          <div class="sidebar-header">
            <input type="text" id="filterInput" class="search-box" placeholder="🔍 Filter files... (Press '/')" oninput="filterLinks()">
          </div>
          <div class="sidebar-list" id="sidebarList">
            <div class="sidebar-title">Modules (${modules.length})</div>
            ${navHtml}
          </div>
        </div>

        <div class="main-content" id="mainContainer">
          <div class="main-header">
            <div>
              <h1>🚀 My-Investments Codebase</h1>
              <span style="font-size: 13px; color: #656d76;">Clean Light Theme &bull; Live in-memory .toString() rendering</span>
            </div>
            <div class="global-actions">
              <button class="btn btn-secondary" onclick="toggleAllCards(true)">Expand All</button>
              <button class="btn btn-secondary" onclick="toggleAllCards(false)">Collapse All</button>
              <button class="btn btn-primary" id="btnCopyAll" onclick="copyEntireProject(this)">📋 Copy Entire Project</button>
            </div>
          </div>
          ${cardsHtml}
        </div>

        <textarea id="entireProjectStore" style="display:none;">${escapeHtml(wholeProjectMarkdown)}</textarea>

        <script>
          hljs.highlightAll();

          function scrollToSection(id, btn) {
            const target = document.getElementById(id);
            if (target) {
              const body = document.getElementById('body_' + id);
              if (body.style.display === 'none') {
                toggleCard(id);
              }
              target.scrollIntoView({ behavior: 'smooth', block: 'start' });
              
              document.querySelectorAll('.nav-link').forEach(el => el.classList.remove('active'));
              if (btn) btn.classList.add('active');
            }
          }

          function filterLinks() {
            const q = document.getElementById('filterInput').value.toLowerCase();
            document.querySelectorAll('.nav-link').forEach(btn => {
              const txt = btn.innerText.toLowerCase();
              btn.style.display = txt.includes(q) ? 'block' : 'none';
            });
          }

          function toggleCard(id) {
            const body = document.getElementById('body_' + id);
            const icon = document.getElementById('icon_' + id);
            if (body.style.display === 'none') {
              body.style.display = 'block';
              icon.innerText = '▼';
            } else {
              body.style.display = 'none';
              icon.innerText = '▶';
            }
          }

          function toggleAllCards(expand) {
            document.querySelectorAll('.card-body').forEach(b => b.style.display = expand ? 'block' : 'none');
            document.querySelectorAll('.toggle-icon').forEach(i => i.innerText = expand ? '▼' : '▶');
          }

          function copySnippet(elementId, btn) {
            const code = document.getElementById(elementId).innerText;
            navigator.clipboard.writeText(code).then(() => {
              const prev = btn.innerText;
              btn.innerText = '✅ Copied!';
              btn.style.color = '#1a7f37';
              setTimeout(() => {
                btn.innerText = prev;
                btn.style.color = '#24292f';
              }, 2000);
            });
          }

          function copyEntireProject(btn) {
            const allCode = document.getElementById('entireProjectStore').value;
            navigator.clipboard.writeText(allCode).then(() => {
              const prev = btn.innerHTML;
              btn.innerHTML = '✅ Entire Project Copied!';
              btn.style.background = '#1f883d';
              setTimeout(() => {
                btn.innerHTML = prev;
                btn.style.background = '#2da44e';
              }, 2500);
            });
          }

          const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
              if (entry.isIntersecting) {
                const id = entry.target.id;
                document.querySelectorAll('.nav-link').forEach(link => {
                  link.classList.toggle('active', link.id === 'btn_' + id);
                });
              }
            });
          }, { root: document.getElementById('mainContainer'), threshold: 0.2 });

          document.querySelectorAll('.file-card').forEach(card => observer.observe(card));

          window.addEventListener('keydown', (e) => {
            if ((e.key === '/' || (e.ctrlKey && e.key === 'k')) && document.activeElement !== document.getElementById('filterInput')) {
              e.preventDefault();
              const input = document.getElementById('filterInput');
              input.focus();
              input.select();
            }
          });
        </script>
      </body>
      </html>
    `;

    return HtmlService.createHtmlOutput(html)
      .setTitle("My-Investments Architecture Viewer")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return HtmlService.createHtmlOutput(`<h3>Render Error:</h3><pre>${escapeHtml(err.stack || err.message)}</pre>`);
  }
}

/**
 * Escapes characters to prevent HTML injection bugs in code blocks.
 */
function escapeHtml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}