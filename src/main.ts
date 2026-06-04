import * as XLSX from "xlsx";
import { analyzeWorkbook, splitList } from "./analysis";
import { renderCharts, resizeCharts } from "./charts";
import { defaultConfig, availableKeywordFields } from "./defaults";
import { exportResult, downloadChartImage, type ExportOptions } from "./export";
import type { AnalysisConfig, AnalysisResult, ClassificationRule, KeywordRule } from "./types";
import "./styles.css";

type SortKey = "downtime" | "date" | "machineType";

const STORAGE_PREFIX = "faw-config-";
const USERS_KEY = "faw-users";
const CURRENT_USER_KEY = "faw-current-user";

let currentUser = "";
let config: AnalysisConfig = structuredClone(defaultConfig);
let workbook: XLSX.WorkBook | null = null;
let result: AnalysisResult | null = null;
let selectedMonth = "合计";
let activeChart = "pareto";
let searchText = "";
let monthFilter = "全部";
let typeFilter = "全部";
let sortKey: SortKey = "downtime";
let sidebarOpen = true;
let activePage = "mainPage";
let exportModalVisible = false;

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("App root missing");

function loadUsers(): string[] {
  try { return JSON.parse(localStorage.getItem(USERS_KEY) || "[]"); } catch { return []; }
}

function saveUsers(users: string[]): void {
  localStorage.setItem(USERS_KEY, JSON.stringify(users));
}

function loadConfig(username: string): AnalysisConfig | null {
  try {
    const raw = localStorage.getItem(STORAGE_PREFIX + username);
    if (!raw) return null;
    return JSON.parse(raw) as AnalysisConfig;
  } catch { return null; }
}

function saveConfig(): void {
  if (!currentUser) return;
  localStorage.setItem(STORAGE_PREFIX + currentUser, JSON.stringify(config));
}

function loginUser(username: string): void {
  currentUser = username;
  localStorage.setItem(CURRENT_USER_KEY, username);
  const saved = loadConfig(username);
  if (saved) {
    config = saved;
  } else {
    config = structuredClone(defaultConfig);
  }
  const users = loadUsers();
  if (!users.includes(username)) {
    users.push(username);
    saveUsers(users);
  }
  updateLoginUI();
  syncConfigToUI();
  renderRuleEditors();
  if (workbook) runAnalysis();
}

function logoutUser(): void {
  currentUser = "";
  localStorage.removeItem(CURRENT_USER_KEY);
  updateLoginUI();
}

function syncConfigToUI(): void {
  const sheetsInput = document.querySelector<HTMLInputElement>("#sheetsInput");
  const deptInput = document.querySelector<HTMLInputElement>("#departmentInput");
  const minDowntimeInput = document.querySelector<HTMLInputElement>("#minDowntimeInput");
  const maxDowntimeInput = document.querySelector<HTMLInputElement>("#maxDowntimeInput");
  const highlightInput = document.querySelector<HTMLTextAreaElement>("#highlightInput");
  if (sheetsInput) sheetsInput.value = config.sheets.join(",");
  if (deptInput) deptInput.value = config.departmentFilter;
  if (minDowntimeInput) minDowntimeInput.value = String(config.minDowntime);
  if (maxDowntimeInput) maxDowntimeInput.value = config.maxDowntime ? String(config.maxDowntime) : "";
  if (highlightInput) highlightInput.value = config.highlightKeywords.join(",");
}

function updateLoginUI(): void {
  const loginArea = document.querySelector<HTMLSpanElement>("#loginArea");
  if (!loginArea) return;
  if (currentUser) {
    loginArea.innerHTML = `
      <span class="user-badge">${escapeHtml(currentUser)}</span>
      <button id="logoutBtn" class="secondary" style="min-height:30px;font-size:12px;padding:0 10px;">切换账号</button>
    `;
    document.querySelector("#logoutBtn")?.addEventListener("click", () => { logoutUser(); showLoginModal(); });
  } else {
    loginArea.innerHTML = `<button id="loginBtn" class="secondary" style="min-height:30px;font-size:12px;padding:0 10px;">登录</button>`;
    document.querySelector("#loginBtn")?.addEventListener("click", showLoginModal);
  }
}

function showLoginModal(): void {
  const users = loadUsers();
  const existing = users.map(u => `<button class="user-pick-btn">${escapeHtml(u)}</button>`).join("");
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.innerHTML = `
    <div class="modal login-modal">
      <h3>选择或创建账号</h3>
      ${existing ? `<div class="user-pick-list">${existing}</div>` : ""}
      <div class="login-new">
        <input id="newUsername" placeholder="输入新账号名" maxlength="20" />
        <button id="loginNewBtn" class="primary">进入</button>
      </div>
      ${currentUser ? `<button id="loginCancelBtn" class="secondary" style="margin-top:8px;">返回</button>` : ""}
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.querySelectorAll(".user-pick-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.body.removeChild(overlay);
      loginUser(btn.textContent || "");
    });
  });
  overlay.querySelector("#loginNewBtn")?.addEventListener("click", () => {
    const input = overlay.querySelector<HTMLInputElement>("#newUsername");
    const name = input?.value.trim();
    if (!name) return;
    document.body.removeChild(overlay);
    loginUser(name);
  });
  overlay.querySelector("#loginCancelBtn")?.addEventListener("click", () => {
    document.body.removeChild(overlay);
    if (currentUser) updateLoginUI();
  });
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) { document.body.removeChild(overlay); if (currentUser) updateLoginUI(); }
  });
}

function showExportModal(): void {
  if (!result) return;
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.innerHTML = `
    <div class="modal export-modal">
      <h3>导出选项</h3>
      <div class="export-checks">
        <label class="check-label"><input type="checkbox" id="expPareto" checked /> 停机柏拉图</label>
        <label class="check-label"><input type="checkbox" id="expMttr" checked /> MTTR / MTBF</label>
        <label class="check-label"><input type="checkbox" id="expTrend" checked /> 故障推移</label>
        <label class="check-label"><input type="checkbox" id="expData" checked /> Excel 数据表</label>
      </div>
      <label style="margin-top:10px;">
        图表月份
        <select id="expMonth">
          ${["合计", ...result.months].map(m => `<option value="${m}" ${m === selectedMonth ? "selected" : ""}>${m}</option>`).join("")}
        </select>
      </label>
      <div class="export-actions">
        <button id="exportDoBtn" class="primary">导出</button>
        <button id="exportCancelBtn" class="secondary">取消</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.querySelector("#exportDoBtn")?.addEventListener("click", () => {
    const expMonth = (overlay.querySelector<HTMLSelectElement>("#expMonth"))?.value || "合计";
    const doPareto = (overlay.querySelector<HTMLInputElement>("#expPareto"))?.checked;
    const doMttr = (overlay.querySelector<HTMLInputElement>("#expMttr"))?.checked;
    const doTrend = (overlay.querySelector<HTMLInputElement>("#expTrend"))?.checked;
    const doData = (overlay.querySelector<HTMLInputElement>("#expData"))?.checked;
    document.body.removeChild(overlay);
    if (!result) return;

    if (doData) exportResult(result, config);

    if (expMonth !== selectedMonth) {
      renderCharts(result, expMonth);
      setTimeout(() => {
        if (doPareto) downloadChartImage("pareto", expMonth);
        if (doMttr) downloadChartImage("mttr", expMonth);
        if (doTrend) downloadChartImage("trend", expMonth);
        renderCharts(result!, selectedMonth);
      }, 500);
    } else {
      if (doPareto) downloadChartImage("pareto", expMonth);
      if (doMttr) downloadChartImage("mttr", expMonth);
      if (doTrend) downloadChartImage("trend", expMonth);
    }
  });
  overlay.querySelector("#exportCancelBtn")?.addEventListener("click", () => {
    document.body.removeChild(overlay);
  });
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) document.body.removeChild(overlay);
  });
}

// ======= init app =======
app.innerHTML = `
  <header class="topbar">
    <div>
      <h1>设备故障数据分析</h1>
      <p>上传 Excel，配置筛选关键词，生成明细、柏拉图、MTTR/MTBF 和故障推移。</p>
    </div>
    <div class="topbar-right">
      <span id="loginArea"></span>
      <button id="exportBtn" class="primary" disabled>导出 Excel</button>
    </div>
  </header>

  <main class="layout">
    <aside class="sidebar">
      <nav class="side-nav" aria-label="功能导航">
        <div class="side-brand">
          <span>设</span>
          <strong>故障分析</strong>
        </div>
        <button id="sidebarToggle" class="sidebar-toggle" aria-expanded="true" aria-label="收起规则">☰</button>

        <button class="side-nav-item active" data-page="mainPage">
            <span class="nav-icon">主</span>
            <span>首页</span>
        </button>
        <button class="side-nav-item" data-page="basicPanel">
            <span class="nav-icon">基</span>
            <span>基础规则</span>
        </button>
        <button class="side-nav-item" data-page="keywordPanel">
            <span class="nav-icon">关</span>
            <span>关键词</span>
        </button>
        <button class="side-nav-item" data-page="classPanel">
            <span class="nav-icon">分</span>
            <span>分类规则</span>
        </button>
      </nav>
    </aside>

    <button id="sidebarOpenToggle" class="sidebar-open-toggle" aria-label="展开规则">☰</button>

    <section class="workspace">
      <section id="mainPage" class="page-view active">
        <section class="panel upload-panel main-upload">
          <div>
            <h2>文件上传</h2>
            <div id="fileMeta" class="muted">尚未上传文件</div>
          </div>
          <label class="dropzone">
            <input id="fileInput" type="file" accept=".xlsx,.xls" />
            <span>选择 Excel 文件</span>
            <small>支持 L1/L2/L3/L4 工作表</small>
          </label>
        </section>
        <div id="warnings" class="warnings"></div>

        <section class="metrics">
          <div><span id="totalRecords">0</span><small>筛选记录</small></div>
          <div><span id="totalDowntime">0</span><small>停机分钟</small></div>
          <div><span id="totalTypes">0</span><small>设备类型</small></div>
          <div><span id="totalMonths">0</span><small>月份</small></div>
        </section>

        <section class="panel result-panel">
          <div class="result-head">
            <h2>图表分析</h2>
            <select id="chartMonth"></select>
          </div>
          <div class="tabs">
            <button data-chart="pareto" class="active">停机柏拉图</button>
            <button data-chart="mttr">MTTR/MTBF</button>
            <button data-chart="trend">故障推移</button>
          </div>
          <div id="paretoChart" class="chart active"></div>
          <div id="mttrChart" class="chart"></div>
          <div id="trendChart" class="chart"></div>
        </section>

        <section class="panel result-panel">
          <div class="result-head">
            <h2>筛选明细</h2>
            <div class="table-tools">
              <input id="tableSearch" placeholder="搜索问题、机器、责任人" />
              <select id="monthFilter"></select>
              <select id="typeFilter"></select>
              <select id="sortSelect">
                <option value="downtime">按停机时长</option>
                <option value="date">按日期</option>
                <option value="machineType">按设备类型</option>
              </select>
            </div>
          </div>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>日期</th><th>线体</th><th>时间</th><th>停机</th><th>设备类型</th><th>问题描述</th><th>责任</th>
                </tr>
              </thead>
              <tbody id="recordRows"></tbody>
            </table>
          </div>
        </section>
      </section>

      <section id="basicPanel" class="page-view panel config-workspace">
        <div class="result-head">
          <h2>基础规则</h2>
        </div>
        <div class="form-grid">
          <label>工作表<input id="sheetsInput" /></label>
          <label>责任部门包含<input id="departmentInput" /></label>
          <label>最小停机时长(min)<input id="minDowntimeInput" type="number" min="0" step="1" /></label>
          <label>最大停机时长(min)<input id="maxDowntimeInput" type="number" min="0" step="1" placeholder="不限制" /></label>
        </div>
        <div style="margin-top:14px;text-align:right;">
          <button id="saveBasicBtn" class="primary">保存并应用</button>
        </div>
      </section>

      <section id="keywordPanel" class="page-view panel config-workspace">
          <div class="result-head">
            <h2>关键词</h2>
          </div>
          <label>
            <span class="label-row">高亮关键词<span class="field-note">在故障内容中高亮以下关键词</span></span>
            <textarea id="highlightInput"></textarea>
          </label>
          <div class="mini-title-row">
            <div class="mini-title">包含规则</div>
            <button id="addIncludeBtn" class="icon-btn" title="新增包含规则">+</button>
          </div>
          <div id="includeRules"></div>
          <div class="mini-title-row">
            <div class="mini-title">排除规则</div>
            <button id="addExcludeBtn" class="icon-btn" title="新增排除规则">+</button>
          </div>
          <div id="excludeRules"></div>
          <div style="margin-top:14px;text-align:right;">
            <button id="saveKeywordBtn" class="primary">保存并应用</button>
          </div>
      </section>

      <section id="classPanel" class="page-view panel config-workspace">
          <div class="result-head">
            <h2>分类规则</h2>
          </div>
          <div class="section-title">
            <span></span>
            <button id="addClassBtn" class="icon-btn" title="新增分类">+</button>
          </div>
          <div class="class-header">
            <span>机器</span>
            <span>关键词</span>
          </div>
          <div id="classRules"></div>
          <div style="margin-top:14px;text-align:right;">
            <button id="saveClassBtn" class="primary">保存并应用</button>
          </div>
      </section>

    </section>
  </main>
`;

bindEvents();
renderRuleEditors();
renderEmptyState();
updateSidebarState();
updatePageView();
updateLoginUI();

// Auto-login from last session
const lastUser = localStorage.getItem(CURRENT_USER_KEY);
if (lastUser && loadConfig(lastUser)) {
  loginUser(lastUser);
}

function bindEvents(): void {
  document.querySelector("#sidebarToggle")?.addEventListener("click", () => {
    sidebarOpen = !sidebarOpen;
    updateSidebarState();
  });
  document.querySelector("#sidebarOpenToggle")?.addEventListener("click", () => {
    sidebarOpen = true;
    updateSidebarState();
  });
  document.querySelectorAll<HTMLButtonElement>(".side-nav-item").forEach((button) => {
    button.addEventListener("click", () => {
      sidebarOpen = true;
      updateSidebarState();
      activePage = button.dataset.page ?? "mainPage";
      updatePageView();
    });
  });

  document.querySelector<HTMLInputElement>("#fileInput")?.addEventListener("change", async (event) => {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    if (!file.name.match(/\.xlsx?$/i)) {
      setWarnings(["请上传 .xlsx 或 .xls 文件"]);
      return;
    }
    const buffer = await file.arrayBuffer();
    workbook = XLSX.read(buffer, { type: "array" });
    document.querySelector("#fileMeta")!.textContent = `${file.name} · ${workbook.SheetNames.length} 个工作表`;
    runAnalysis();
  });

  // Basic panel - no auto-run, just track changes
  bindInputNoRun("#sheetsInput", (value) => { config.sheets = splitList(value); });
  bindInputNoRun("#departmentInput", (value) => { config.departmentFilter = value.trim(); });
  bindInputNoRun("#minDowntimeInput", (value) => { config.minDowntime = Number(value) || 0; });
  bindInputNoRun("#maxDowntimeInput", (value) => { config.maxDowntime = Number(value) || 0; });
  bindInputNoRun("#highlightInput", (value) => { config.highlightKeywords = splitList(value); });

  // Save buttons
  document.querySelector("#saveBasicBtn")?.addEventListener("click", () => { saveConfig(); runAnalysis(); });
  document.querySelector("#saveKeywordBtn")?.addEventListener("click", () => { saveConfig(); runAnalysis(); });
  document.querySelector("#saveClassBtn")?.addEventListener("click", () => { saveConfig(); runAnalysis(); });

  document.querySelector("#addIncludeBtn")?.addEventListener("click", () => {
    config.includeKeywords.push({ fields: ["description"], keywords: [], matchMode: "any" });
    renderRuleEditors();
  });
  document.querySelector("#addExcludeBtn")?.addEventListener("click", () => {
    config.excludeKeywords.push({ fields: ["description"], keywords: [], matchMode: "any" });
    renderRuleEditors();
  });
  document.querySelector("#addClassBtn")?.addEventListener("click", () => {
    config.classificationRules.push({ type: "新分类", keywords: [] });
    renderRuleEditors();
  });
  document.querySelector("#exportBtn")?.addEventListener("click", () => {
    if (result) showExportModal();
  });

  document.querySelector("#chartMonth")?.addEventListener("change", (event) => {
    selectedMonth = (event.target as HTMLSelectElement).value;
    if (result) renderCharts(result, selectedMonth);
  });

  document.querySelectorAll<HTMLButtonElement>(".tabs button").forEach((button) => {
    button.addEventListener("click", () => {
      activeChart = button.dataset.chart ?? "pareto";
      updateChartVisibility();
    });
  });

  bindInputNoRun("#tableSearch", (value) => { searchText = value.trim(); renderTable(); });
  document.querySelector("#monthFilter")?.addEventListener("change", (event) => { monthFilter = (event.target as HTMLSelectElement).value; renderTable(); });
  document.querySelector("#typeFilter")?.addEventListener("change", (event) => { typeFilter = (event.target as HTMLSelectElement).value; renderTable(); });
  document.querySelector("#sortSelect")?.addEventListener("change", (event) => { sortKey = (event.target as HTMLSelectElement).value as SortKey; renderTable(); });
  window.addEventListener("resize", resizeCharts);
}

function bindInputNoRun(selector: string, handler: (value: string) => void): void {
  document.querySelector<HTMLInputElement | HTMLTextAreaElement>(selector)?.addEventListener("input", (event) => {
    handler((event.target as HTMLInputElement).value);
  });
}

function runAnalysis(): void {
  if (!workbook) {
    renderEmptyState();
    return;
  }
  result = analyzeWorkbook(workbook, config);
  renderResult();
}

function renderRuleEditors(): void {
  renderKeywordRules("includeRules", config.includeKeywords, "include");
  renderKeywordRules("excludeRules", config.excludeKeywords, "exclude");
  renderClassRules();
}

function renderKeywordRules(containerId: string, rules: KeywordRule[], kind: "include" | "exclude"): void {
  const container = document.querySelector<HTMLDivElement>(`#${containerId}`);
  if (!container) return;
  container.innerHTML = rules.map((rule, index) => `
    <div class="rule-row">
      <select data-kind="${kind}" data-index="${index}" data-prop="field">
        ${availableKeywordFields.map((field) => `<option value="${field.key}" ${rule.fields.includes(field.key) ? "selected" : ""}>${field.label}</option>`).join("")}
      </select>
      <select data-kind="${kind}" data-index="${index}" data-prop="mode">
        <option value="any" ${rule.matchMode === "any" ? "selected" : ""}>任一命中</option>
        <option value="all" ${rule.matchMode === "all" ? "selected" : ""}>全部命中</option>
      </select>
      <input data-kind="${kind}" data-index="${index}" data-prop="keywords" value="${escapeHtml(rule.keywords.join(","))}" placeholder="关键词，逗号分隔" />
      <button data-kind="${kind}" data-index="${index}" data-prop="delete" class="icon-btn danger" title="删除">×</button>
    </div>
  `).join("");

  container.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLButtonElement>("[data-kind]").forEach((element) => {
    element.addEventListener("input", updateKeywordRule);
    element.addEventListener("click", updateKeywordRule);
  });
}

function updateKeywordRule(event: Event): void {
  const target = event.target as HTMLInputElement | HTMLSelectElement | HTMLButtonElement;
  const kind = target.dataset.kind as "include" | "exclude";
  const index = Number(target.dataset.index);
  const prop = target.dataset.prop;
  const rules = kind === "include" ? config.includeKeywords : config.excludeKeywords;
  if (prop === "delete") {
    rules.splice(index, 1);
    renderRuleEditors();
    return;
  }
  if (prop === "field") rules[index].fields = [(target as HTMLSelectElement).value];
  if (prop === "mode") rules[index].matchMode = (target as HTMLSelectElement).value as "any" | "all";
  if (prop === "keywords") rules[index].keywords = splitList((target as HTMLInputElement).value);
}

function renderClassRules(): void {
  const container = document.querySelector<HTMLDivElement>("#classRules");
  if (!container) return;
  container.innerHTML = config.classificationRules.map((rule, index) => `
    <div class="class-row">
      <input data-class-index="${index}" data-prop="type" value="${escapeHtml(rule.type)}" />
      <textarea data-class-index="${index}" data-prop="keywords">${escapeHtml(rule.keywords.join(","))}</textarea>
      <button data-class-index="${index}" data-prop="delete" class="icon-btn danger" title="删除">×</button>
    </div>
  `).join("");

  container.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLButtonElement>("[data-class-index]").forEach((element) => {
    element.addEventListener("input", updateClassRule);
    element.addEventListener("click", updateClassRule);
  });
}

function updateClassRule(event: Event): void {
  const target = event.target as HTMLInputElement | HTMLTextAreaElement | HTMLButtonElement;
  const index = Number(target.dataset.classIndex);
  const prop = target.dataset.prop as keyof ClassificationRule | "delete";
  if (prop === "delete") {
    config.classificationRules.splice(index, 1);
    renderRuleEditors();
    return;
  }
  if (prop === "type") config.classificationRules[index].type = (target as HTMLInputElement).value.trim();
  if (prop === "keywords") config.classificationRules[index].keywords = splitList((target as HTMLTextAreaElement).value);
}

function renderResult(): void {
  if (!result) return;
  setWarnings(result.warnings);
  document.querySelector("#exportBtn")?.removeAttribute("disabled");
  document.querySelector("#totalRecords")!.textContent = String(result.records.length);
  document.querySelector("#totalDowntime")!.textContent = String(Math.round(result.records.reduce((sum, record) => sum + record.downtime, 0)));
  document.querySelector("#totalTypes")!.textContent = String(result.typeSummary.length);
  document.querySelector("#totalMonths")!.textContent = String(result.months.length);

  fillSelect("#chartMonth", ["合计", ...result.months], selectedMonth);
  fillSelect("#monthFilter", ["全部", ...result.months], monthFilter);
  fillSelect("#typeFilter", ["全部", ...result.typeSummary.map((row) => row.type)], typeFilter);
  renderCharts(result, selectedMonth);
  updateChartVisibility();
  renderTable();
}

function renderTable(): void {
  const tbody = document.querySelector<HTMLTableSectionElement>("#recordRows");
  if (!tbody || !result) return;

  const rows = result.records
    .filter((record) => monthFilter === "全部" || record.date.startsWith(monthFilter))
    .filter((record) => typeFilter === "全部" || record.machineType === typeFilter)
    .filter((record) => !searchText || [record.description, record.machine, record.owner, record.department].some((value) => value.includes(searchText)))
    .sort((a, b) => {
      if (sortKey === "date") return b.date.localeCompare(a.date);
      if (sortKey === "machineType") return a.machineType.localeCompare(b.machineType);
      return b.downtime - a.downtime;
    })
    .slice(0, 300);

  tbody.innerHTML = rows.length ? rows.map((record) => `
    <tr class="${record.highlighted ? "highlight" : ""}">
      <td>${escapeHtml(record.date)}</td>
      <td>${escapeHtml(record.line)}</td>
      <td>${escapeHtml(record.startTime)}-${escapeHtml(record.endTime)}</td>
      <td><span class="${downtimeClass(record.downtime)}">${record.downtime}</span></td>
      <td>${escapeHtml(record.machineType)}</td>
      <td>${escapeHtml(record.description)}</td>
      <td>${escapeHtml(record.department)} / ${escapeHtml(record.owner)}</td>
    </tr>
  `).join("") : `<tr><td colspan="7" class="empty">暂无符合条件的数据</td></tr>`;
}

function renderEmptyState(): void {
  setWarnings([]);
  document.querySelector("#totalRecords")!.textContent = "0";
  document.querySelector("#totalDowntime")!.textContent = "0";
  document.querySelector("#totalTypes")!.textContent = "0";
  document.querySelector("#totalMonths")!.textContent = "0";
  document.querySelector<HTMLTableSectionElement>("#recordRows")!.innerHTML = `<tr><td colspan="7" class="empty">上传文件后显示筛选明细</td></tr>`;
}

function setWarnings(warnings: string[]): void {
  const element = document.querySelector<HTMLDivElement>("#warnings");
  if (!element) return;
  element.innerHTML = warnings.length ? warnings.map((warning) => `<span>${escapeHtml(warning)}</span>`).join("") : "";
}

function fillSelect(selector: string, values: string[], selected: string): void {
  const select = document.querySelector<HTMLSelectElement>(selector);
  if (!select) return;
  const nextSelected = values.includes(selected) ? selected : values[0];
  select.innerHTML = values.map((value) => `<option value="${escapeHtml(value)}" ${value === nextSelected ? "selected" : ""}>${escapeHtml(value)}</option>`).join("");
  if (selector === "#chartMonth") selectedMonth = nextSelected;
  if (selector === "#monthFilter") monthFilter = nextSelected;
  if (selector === "#typeFilter") typeFilter = nextSelected;
}

function updateChartVisibility(): void {
  document.querySelectorAll(".tabs button").forEach((button) => button.classList.toggle("active", (button as HTMLButtonElement).dataset.chart === activeChart));
  document.querySelectorAll(".chart").forEach((chart) => chart.classList.toggle("active", chart.id.startsWith(activeChart)));
  resizeCharts();
}

function updateSidebarState(): void {
  const layout = document.querySelector<HTMLElement>(".layout");
  const button = document.querySelector<HTMLButtonElement>("#sidebarToggle");
  const openButton = document.querySelector<HTMLButtonElement>("#sidebarOpenToggle");
  layout?.classList.toggle("sidebar-collapsed", !sidebarOpen);
  if (button) {
    button.textContent = "☰";
    button.setAttribute("aria-label", sidebarOpen ? "收起规则" : "展开规则");
    button.setAttribute("aria-expanded", String(sidebarOpen));
  }
  if (openButton) openButton.hidden = true;
  window.setTimeout(resizeCharts, 180);
}

function updatePageView(): void {
  document.querySelectorAll<HTMLElement>(".page-view").forEach((panel) => {
    panel.classList.toggle("active", panel.id === activePage);
  });
  document.querySelectorAll<HTMLButtonElement>(".side-nav-item").forEach((button) => {
    button.classList.toggle("active", button.dataset.page === activePage);
  });
  window.setTimeout(resizeCharts, 60);
}

function downtimeClass(value: number): string {
  if (value > 480) return "dt dt-yellow";
  if (value > 120) return "dt dt-red";
  if (value > 60) return "dt dt-pink";
  return "dt";
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;" }[char] ?? char));
}
