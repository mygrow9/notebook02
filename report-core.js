import {
  CANONICAL_HEADERS,
  buildReport,
  monthSheetName,
  normalizeHeader,
  parseNumber,
} from "./report-core.js";

const XLSX = window.XLSX;

const els = {
  fileInput: document.querySelector("#fileInput"),
  dropZone: document.querySelector("#dropZone"),
  downloadBtn: document.querySelector("#downloadBtn"),
  downloadFinalBtn: document.querySelector("#downloadFinalBtn"),
  tabs: document.querySelector("#tabs"),
  previewTable: document.querySelector("#previewTable"),
  tableTitle: document.querySelector("#tableTitle"),
  tableMeta: document.querySelector("#tableMeta"),
  chartPanel: document.querySelector("#chartPanel"),
  messagePanel: document.querySelector("#messagePanel"),
  monthMetric: document.querySelector("#monthMetric"),
  unitMetric: document.querySelector("#unitMetric"),
  totalMetric: document.querySelector("#totalMetric"),
  statusMetric: document.querySelector("#statusMetric"),
};

const state = {
  report: null,
  activeTab: null,
};

if (els.fileInput) {
  els.fileInput.addEventListener("change", async (event) => {
    await handleFiles(Array.from(event.target.files ?? []));
    event.target.value = "";
  });
}

els.downloadBtn.addEventListener("click", () => {
  if (!state.report) return;
  downloadWorkbook(state.report);
});

if (els.downloadFinalBtn) {
  els.downloadFinalBtn.addEventListener("click", () => {
    if (!state.report) return;
    downloadFinalHtml(state.report);
  });
}

if (els.dropZone) {
  for (const eventName of ["dragenter", "dragover"]) {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.add("dragging");
    });
  }

  for (const eventName of ["dragleave", "drop"]) {
    els.dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropZone.classList.remove("dragging");
    });
  }

  els.dropZone.addEventListener("drop", async (event) => {
    await handleFiles(Array.from(event.dataTransfer?.files ?? []));
  });
}

async function handleFiles(files) {
  clearMessage();

  const excelFiles = files.filter((file) => /\.(xlsx|xls)$/i.test(file.name));
  if (excelFiles.length === 0) {
    showMessage(["請選擇 Excel 檔案。"]);
    return;
  }

  try {
    const parsedReports = [];
    const errors = [];

    for (const file of excelFiles) {
      try {
        parsedReports.push(await parseWorkbookFile(file));
      } catch (error) {
        errors.push(`${file.name}：${error.message}`);
      }
    }

    if (errors.length) {
      showMessage(errors);
      resetResults("部分檔案無法解析");
      return;
    }

    const dedupedReports = dedupeByMonth(parsedReports);
    state.report = buildReport(dedupedReports);
    state.activeTab = "加總";
    renderAll();
  } catch (error) {
    showMessage([error.message]);
    resetResults("產生失敗");
  }
}

async function parseWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("找不到工作表。");

  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false,
  });

  const headerIndex = rawRows.findIndex((row) =>
    row.some((cell) => normalizeHeader(cell) === "單位"),
  );
  if (headerIndex < 0) throw new Error("找不到「單位」標題列。");

  const headers = rawRows[headerIndex].map((cell) => String(cell ?? "").trim());
  const columnIndexes = mapColumns(headers);
  const month = detectMonth({ fileName: file.name, sheetName, rows: rawRows });

  const rows = rawRows.slice(headerIndex + 1)
    .map((rawRow) => rowToRecord(rawRow, columnIndexes))
    .filter((row) => row["單位"]);

  return {
    month,
    fileName: file.name,
    sheetName: monthSheetName(month),
    headers: CANONICAL_HEADERS,
    rows,
    rawRows: trimRows(rawRows),
  };
}

function mapColumns(headers) {
  const normalizedHeaders = headers.map(normalizeHeader);
  const indexes = {};

  for (const header of CANONICAL_HEADERS) {
    const index = normalizedHeaders.indexOf(normalizeHeader(header));
    if (index < 0) throw new Error(`缺少欄位「${header}」。`);
    indexes[header] = index;
  }

  return indexes;
}

function rowToRecord(rawRow, columnIndexes) {
  const record = {};
  for (const header of CANONICAL_HEADERS) {
    const value = rawRow[columnIndexes[header]];
    record[header] = header === "單位" ? String(value ?? "").trim() : parseNumber(value);
  }
  return record;
}

function detectMonth({ fileName, sheetName, rows }) {
  const textSources = [
    fileName,
    sheetName,
    ...rows.slice(0, 5).flat().map((value) => String(value ?? "")),
  ];

  for (const text of textSources) {
    const periodMatch = text.match(/統計期間[:：]\s*(?:\d{4}\s*\/\s*)?(\d{1,2})/);
    if (periodMatch) return parseMonth(periodMatch[1]);
  }

  for (const text of [fileName, sheetName]) {
    const simpleMatch = text.match(/(?:^|[^\d])([01]?\d)(?:[^\d]|$)/);
    if (simpleMatch) return parseMonth(simpleMatch[1]);
  }

  throw new Error("無法從檔名、工作表或統計期間辨識月份。");
}

function parseMonth(value) {
  const month = parseInt(value, 10);
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error(`月份不在 1-12 範圍：${value}`);
  }
  return month;
}

function trimRows(rows) {
  return rows
    .map((row) => {
      let last = row.length - 1;
      while (last >= 0 && String(row[last] ?? "").trim() === "") last -= 1;
      return row.slice(0, last + 1);
    })
    .filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""));
}

function dedupeByMonth(reports) {
  const byMonth = new Map();
  for (const report of reports) {
    byMonth.set(report.month, report);
  }
  return Array.from(byMonth.values()).sort((a, b) => a.month - b.month);
}

function renderAll() {
  const { report } = state;
  const totalRow = report.parentRows.at(-1);
  els.monthMetric.textContent = report.months.length;
  els.unitMetric.textContent = Math.max(0, report.totalRows.length - 1).toLocaleString("zh-TW");
  els.totalMetric.textContent = Number(totalRow.at(-1) ?? 0).toLocaleString("zh-TW");
  els.statusMetric.textContent = "已產生";
  els.downloadBtn.disabled = false;
  if (els.downloadFinalBtn) els.downloadFinalBtn.disabled = false;

  renderTabs();
  renderActiveTab();
}

function renderTabs() {
  const tabNames = getTabs().map((tab) => tab.name);
  els.tabs.replaceChildren(...tabNames.map((name) => {
    const button = document.createElement("button");
    button.className = `tab-button${state.activeTab === name ? " active" : ""}`;
    button.type = "button";
    button.textContent = name;
    button.addEventListener("click", () => {
      state.activeTab = name;
      renderTabs();
      renderActiveTab();
    });
    return button;
  }));
}

function getTabs() {
  if (!state.report) return [];
  return [
    { name: "加總", rows: state.report.totalRows },
    { name: "母單位月彙總", rows: state.report.parentRows },
    { name: "直條圖", rows: state.report.chartRows },
    ...state.report.monthlyReports.map((report) => ({
      name: report.sheetName,
      rows: rowsForMonthlyPreview(report),
    })),
  ];
}

function renderActiveTab() {
  const tab = getTabs().find((item) => item.name === state.activeTab);
  if (!tab) return;

  els.tableTitle.textContent = tab.name;
  els.tableMeta.textContent = `${Math.max(0, tab.rows.length - 1).toLocaleString("zh-TW")} 筆資料`;
  renderTable(tab.rows);
  renderChart(tab.name === "直條圖" ? tab.rows : null);
}

function rowsForMonthlyPreview(report) {
  return [
    CANONICAL_HEADERS,
    ...report.rows.map((row) => CANONICAL_HEADERS.map((header) => row[header])),
  ];
}

function renderTable(rows) {
  if (!rows.length) {
    els.previewTable.replaceChildren();
    return;
  }

  const [headers, ...bodyRows] = rows;
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headers.forEach((header, index) => {
    const th = document.createElement("th");
    th.textContent = header;
    if (index > 0) th.classList.add("number");
    headRow.append(th);
  });
  thead.append(headRow);

  const tbody = document.createElement("tbody");
  bodyRows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell, index) => {
      const td = document.createElement("td");
      td.textContent = typeof cell === "number" ? cell.toLocaleString("zh-TW") : String(cell ?? "");
      if (index > 0 && typeof cell === "number") td.classList.add("number");
      tr.append(td);
    });
    tbody.append(tr);
  });

  els.previewTable.replaceChildren(thead, tbody);
}

function renderChart(rows) {
  if (!rows) {
    els.chartPanel.hidden = true;
    els.chartPanel.replaceChildren();
    return;
  }

  const data = rows.slice(1);
  const max = Math.max(...data.map((row) => Number(row[1]) || 0), 1);
  els.chartPanel.replaceChildren(...data.map(([label, value]) => {
    const row = document.createElement("div");
    row.className = "bar-row";
    const name = document.createElement("span");
    name.textContent = label;
    const track = document.createElement("div");
    track.className = "bar-track";
    const fill = document.createElement("div");
    fill.className = "bar-fill";
    const numericValue = Number(value) || 0;
    const percent = numericValue > 0 ? (numericValue / max) * 100 : 0;
    fill.style.width = `${percent}%`;
    fill.style.minWidth = numericValue > 0 ? "2px" : "0";
    track.append(fill);
    const total = document.createElement("strong");
    total.textContent = Number(value).toLocaleString("zh-TW");
    row.append(name, track, total);
    return row;
  }));
  els.chartPanel.hidden = false;
}

function downloadWorkbook(report) {
  const workbook = XLSX.utils.book_new();

  for (const monthlyReport of report.monthlyReports) {
    const worksheet = XLSX.utils.aoa_to_sheet(monthlyReport.rawRows);
    styleWorksheetColumns(worksheet, monthlyReport.rawRows);
    XLSX.utils.book_append_sheet(workbook, worksheet, monthlyReport.sheetName);
  }

  appendSheet(workbook, "加總", report.totalRows);
  appendSheet(workbook, "母單位月彙總", report.parentRows);
  appendSheet(workbook, "直條圖", report.chartRows);

  XLSX.writeFile(workbook, "總表.xlsx", { compression: true });
}

function downloadFinalHtml(report) {
  const html = buildFinalHtml(report);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "final.html";
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildFinalHtml(report) {
  const reportJson = JSON.stringify(report).replace(/</g, "\\u003c");
  return `<!doctype html>
<html lang="zh-Hant">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>車牌辨識張數月報表</title>
    <style>${FINAL_HTML_CSS}</style>
  </head>
  <body>
    <main class="app-shell">
      <section class="workspace">
        <header class="topbar">
          <div>
            <p class="eyebrow">Excel 月報彙整</p>
            <h1>車牌辨識張數月報表</h1>
          </div>
          <div class="button-group">
            <button id="downloadBtn" class="primary-button" type="button">下載總表.xlsx</button>
          </div>
        </header>

        <section class="status-grid" aria-live="polite">
          <div><span class="metric-label">月份</span><strong id="monthMetric">0</strong></div>
          <div><span class="metric-label">單位列數</span><strong id="unitMetric">0</strong></div>
          <div><span class="metric-label">總數</span><strong id="totalMetric">0</strong></div>
          <div><span class="metric-label">狀態</span><strong id="statusMetric">已載入</strong></div>
        </section>

        <section class="results">
          <nav id="tabs" class="tabs" aria-label="結果表"></nav>
          <div class="table-toolbar">
            <h2 id="tableTitle">加總</h2>
            <span id="tableMeta"></span>
          </div>
          <div id="chartPanel" class="chart-panel" hidden></div>
          <div class="table-wrap"><table id="previewTable"></table></div>
        </section>
      </section>
    </main>
    <script id="reportData" type="application/json">${reportJson}</script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"><\/script>
    <script>${FINAL_HTML_SCRIPT}<\/script>
  </body>
</html>`;
}

const FINAL_HTML_CSS = `:root{--bg:#f6f7f9;--panel:#ffffff;--panel-strong:#eef3f8;--ink:#17202a;--muted:#667085;--line:#d8dee8;--brand:#0f766e;--brand-dark:#115e59;--bar:#2563eb;font-family:"Microsoft JhengHei","PingFang TC","Noto Sans TC",Arial,sans-serif}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink)}button{font:inherit}.app-shell{min-height:100vh;padding:24px}.workspace{max-width:1280px;margin:0 auto}.topbar{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:18px}.button-group{display:flex;align-items:center;justify-content:flex-end;gap:10px;flex-wrap:wrap}.eyebrow{margin:0 0 4px;color:var(--brand-dark);font-size:13px;font-weight:700}h1,h2{margin:0;letter-spacing:0}h1{font-size:28px}h2{font-size:18px}.primary-button{border:1px solid var(--brand-dark);background:var(--brand);color:#fff;border-radius:6px;min-height:42px;padding:0 16px;display:inline-flex;align-items:center;justify-content:center;cursor:pointer;white-space:nowrap}.status-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-bottom:14px}.status-grid>div{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:12px 14px}.metric-label{display:block;color:var(--muted);font-size:12px;margin-bottom:4px}.status-grid strong{font-size:20px}.results{background:var(--panel);border:1px solid var(--line);border-radius:8px;overflow:hidden}.tabs{display:flex;gap:4px;overflow-x:auto;padding:10px;background:var(--panel-strong);border-bottom:1px solid var(--line)}.tab-button{border:1px solid transparent;background:transparent;color:var(--muted);border-radius:6px;min-height:34px;padding:0 12px;cursor:pointer;white-space:nowrap}.tab-button.active{background:#fff;border-color:var(--line);color:var(--ink);font-weight:700}.table-toolbar{display:flex;align-items:baseline;justify-content:space-between;gap:12px;padding:14px 16px;border-bottom:1px solid var(--line)}.table-toolbar span{color:var(--muted);font-size:14px}.table-wrap{overflow:auto;max-height:62vh}table{width:100%;min-width:840px;border-collapse:separate;border-spacing:0}th,td{border-right:1px solid var(--line);border-bottom:1px solid var(--line);padding:8px 10px;font-size:13px;line-height:1.35;background:#fff;white-space:nowrap}th{position:sticky;top:0;z-index:1;background:#dbeafe;text-align:left;font-weight:700}td.number,th.number{text-align:right}tr:nth-child(even) td{background:#fbfcfe}.chart-panel{padding:14px 16px;border-bottom:1px solid var(--line);display:grid;gap:8px}.bar-row{display:grid;grid-template-columns:minmax(120px,180px) 1fr 76px;align-items:center;gap:10px;font-size:13px}.bar-track{height:18px;background:#eef2ff;border-radius:4px;overflow:hidden}.bar-fill{height:100%;min-width:2px;background:var(--bar)}@media(max-width:760px){.app-shell{padding:14px}.topbar,.button-group,.table-toolbar{align-items:stretch;flex-direction:column}.status-grid{grid-template-columns:repeat(2,minmax(0,1fr))}h1{font-size:22px}}`;

const FINAL_HTML_SCRIPT = `(function(){const XLSX=window.XLSX;const report=JSON.parse(document.getElementById("reportData").textContent);const els={downloadBtn:document.querySelector("#downloadBtn"),tabs:document.querySelector("#tabs"),previewTable:document.querySelector("#previewTable"),tableTitle:document.querySelector("#tableTitle"),tableMeta:document.querySelector("#tableMeta"),chartPanel:document.querySelector("#chartPanel"),monthMetric:document.querySelector("#monthMetric"),unitMetric:document.querySelector("#unitMetric"),totalMetric:document.querySelector("#totalMetric"),statusMetric:document.querySelector("#statusMetric")};let activeTab="加總";els.downloadBtn.addEventListener("click",()=>downloadWorkbook(report));renderAll();function renderAll(){const totalRow=report.parentRows.at(-1);els.monthMetric.textContent=report.months.length;els.unitMetric.textContent=Math.max(0,report.totalRows.length-1).toLocaleString("zh-TW");els.totalMetric.textContent=Number(totalRow.at(-1)||0).toLocaleString("zh-TW");els.statusMetric.textContent="已載入";renderTabs();renderActiveTab()}function getTabs(){return[{name:"加總",rows:report.totalRows},{name:"母單位月彙總",rows:report.parentRows},{name:"直條圖",rows:report.chartRows},...report.monthlyReports.map(item=>({name:item.sheetName,rows:[["單位","總數(正常車及警示車)","正常車","警示車","失車","註(吊)銷","典當車","AB車","權利車","偽造車","警示車_未分類"],...item.rows.map(row=>["單位","總數(正常車及警示車)","正常車","警示車","失車","註(吊)銷","典當車","AB車","權利車","偽造車","警示車_未分類"].map(header=>row[header]))]}))]}function renderTabs(){els.tabs.replaceChildren(...getTabs().map(tab=>{const button=document.createElement("button");button.className="tab-button"+(activeTab===tab.name?" active":"");button.type="button";button.textContent=tab.name;button.addEventListener("click",()=>{activeTab=tab.name;renderTabs();renderActiveTab()});return button}))}function renderActiveTab(){const tab=getTabs().find(item=>item.name===activeTab);if(!tab)return;els.tableTitle.textContent=tab.name;els.tableMeta.textContent=Math.max(0,tab.rows.length-1).toLocaleString("zh-TW")+" 筆資料";renderTable(tab.rows);renderChart(tab.name==="直條圖"?tab.rows:null)}function renderTable(rows){if(!rows.length){els.previewTable.replaceChildren();return}const [headers,...bodyRows]=rows;const thead=document.createElement("thead");const headRow=document.createElement("tr");headers.forEach((header,index)=>{const th=document.createElement("th");th.textContent=header;if(index>0)th.classList.add("number");headRow.append(th)});thead.append(headRow);const tbody=document.createElement("tbody");bodyRows.forEach(row=>{const tr=document.createElement("tr");row.forEach((cell,index)=>{const td=document.createElement("td");td.textContent=typeof cell==="number"?cell.toLocaleString("zh-TW"):String(cell||"");if(index>0&&typeof cell==="number")td.classList.add("number");tr.append(td)});tbody.append(tr)});els.previewTable.replaceChildren(thead,tbody)}function renderChart(rows){if(!rows){els.chartPanel.hidden=true;els.chartPanel.replaceChildren();return}const data=rows.slice(1);const max=Math.max(...data.map(row=>Number(row[1])||0),1);els.chartPanel.replaceChildren(...data.map(([label,value])=>{const row=document.createElement("div");row.className="bar-row";const name=document.createElement("span");name.textContent=label;const track=document.createElement("div");track.className="bar-track";const fill=document.createElement("div");fill.className="bar-fill";const numericValue=Number(value)||0;const percent=numericValue>0?numericValue/max*100:0;fill.style.width=percent+"%";fill.style.minWidth=numericValue>0?"2px":"0";track.append(fill);const total=document.createElement("strong");total.textContent=Number(value).toLocaleString("zh-TW");row.append(name,track,total);return row}));els.chartPanel.hidden=false}function downloadWorkbook(report){const workbook=XLSX.utils.book_new();for(const monthlyReport of report.monthlyReports){const worksheet=XLSX.utils.aoa_to_sheet(monthlyReport.rawRows);styleWorksheetColumns(worksheet,monthlyReport.rawRows);XLSX.utils.book_append_sheet(workbook,worksheet,monthlyReport.sheetName)}appendSheet(workbook,"加總",report.totalRows);appendSheet(workbook,"母單位月彙總",report.parentRows);appendSheet(workbook,"直條圖",report.chartRows);XLSX.writeFile(workbook,"總表.xlsx",{compression:true})}function appendSheet(workbook,name,rows){const worksheet=XLSX.utils.aoa_to_sheet(rows);styleWorksheetColumns(worksheet,rows);XLSX.utils.book_append_sheet(workbook,worksheet,name)}function styleWorksheetColumns(worksheet,rows){const colCount=rows.reduce((max,row)=>Math.max(max,row.length),0);worksheet["!cols"]=Array.from({length:colCount},(_,colIndex)=>{const width=rows.reduce((max,row)=>{const text=String(row[colIndex]||"");return Math.max(max,Math.min(text.length+4,32))},colIndex===0?18:10);return{wch:width}})}})();`;

function appendSheet(workbook, name, rows) {
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  styleWorksheetColumns(worksheet, rows);
  XLSX.utils.book_append_sheet(workbook, worksheet, name);
}

function styleWorksheetColumns(worksheet, rows) {
  const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  worksheet["!cols"] = Array.from({ length: colCount }, (_, colIndex) => {
    const width = rows.reduce((max, row) => {
      const text = String(row[colIndex] ?? "");
      return Math.max(max, Math.min(text.length + 4, 32));
    }, colIndex === 0 ? 18 : 10);
    return { wch: width };
  });
}

function showMessage(messages) {
  els.messagePanel.innerHTML = messages.map((message) => `<div>${escapeHtml(message)}</div>`).join("");
  els.messagePanel.hidden = false;
}

function clearMessage() {
  els.messagePanel.hidden = true;
  els.messagePanel.replaceChildren();
}

function resetResults(status) {
  state.report = null;
  state.activeTab = null;
  els.downloadBtn.disabled = true;
  if (els.downloadFinalBtn) els.downloadFinalBtn.disabled = true;
  els.monthMetric.textContent = "0";
  els.unitMetric.textContent = "0";
  els.totalMetric.textContent = "0";
  els.statusMetric.textContent = status;
  els.tabs.replaceChildren();
  els.previewTable.replaceChildren();
  els.chartPanel.replaceChildren();
  els.chartPanel.hidden = true;
  els.tableTitle.textContent = "尚未產生結果";
  els.tableMeta.textContent = "上傳月報後會顯示總表內容。";
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
