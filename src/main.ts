import * as XLSX from "xlsx";
import { analyzeWorkbook, splitList } from "./analysis";
import { renderCharts, resizeCharts, renderParetoInline, renderMttrInline, renderTrendInline, renderDailyTrendInline, getChartInstance } from "./charts";
import { defaultConfig, availableKeywordFields } from "./defaults";
import { exportFullReport, type ExportOptions } from "./export";
import type { AnalysisConfig, AnalysisResult, ClassificationRule, KeywordRule, FaultRecord } from "./types";
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
let selectedIds = new Set<string>();
let currentPage = 1;
let pageSize = 20;
let lineFilter = "全部";
let sortOrder: "asc" | "desc" = "desc";
let _filteredCache: FaultRecord[] | null = null;
let _filteredCacheKey = "";

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("App root missing");

// ===== Toast =====
function showToast(message: string): void {
  const existing = document.querySelector<HTMLDivElement>(".toast");
  if (existing) existing.remove();
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  document.body.appendChild(toast);
  requestAnimationFrame(() => {
    toast.classList.add("toast-show");
    setTimeout(() => {
      toast.classList.remove("toast-show");
      setTimeout(() => toast.remove(), 400);
    }, 2000);
  });
}

// ===== User / Config =====
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
  if (saved) config = saved; else config = structuredClone(defaultConfig);
  const users = loadUsers();
  if (!users.includes(username)) { users.push(username); saveUsers(users); }
  updateLoginUI();
  syncConfigToUI();
  renderRuleEditors();
  if (workbook) runAnalysis();
}

function logoutUser(): void { currentUser = ""; localStorage.removeItem(CURRENT_USER_KEY); updateLoginUI(); }

function deleteUser(username: string): void {
  const users = loadUsers().filter(u => u !== username);
  saveUsers(users);
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
    loginArea.innerHTML = `<span class="user-badge">${escapeHtml(currentUser)}</span><button id="logoutBtn" class="secondary" style="min-height:30px;font-size:12px;padding:0 10px;">切换账号</button>`;
    document.querySelector("#logoutBtn")?.addEventListener("click", () => { logoutUser(); showLoginModal(); });
  } else {
    loginArea.innerHTML = `<button id="loginBtn" class="secondary" style="min-height:30px;font-size:12px;padding:0 10px;">登录</button>`;
    document.querySelector("#loginBtn")?.addEventListener("click", showLoginModal);
  }
}

function showLoginModal(): void {
  const users = loadUsers();
  const existing = users.map(u => `<div class="user-pick-row"><button class="user-pick-btn">${escapeHtml(u)}</button><button class="user-del-btn" data-username="${escapeHtml(u)}" title="删除账号（保留数据）">×</button></div>`).join("");
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.innerHTML = `<div class="modal login-modal"><h3>选择或创建账号</h3>${existing ? `<div class="user-pick-list">${existing}</div>` : ""}<div class="login-new"><input id="newUsername" placeholder="输入新账号名" maxlength="20" /><button id="loginNewBtn" class="primary">进入</button></div>${currentUser ? `<button id="loginCancelBtn" class="secondary" style="margin-top:8px;">返回</button>` : ""}</div>`;
  document.body.appendChild(overlay);
  overlay.querySelectorAll(".user-pick-btn").forEach(btn => {
    btn.addEventListener("click", () => { document.body.removeChild(overlay); loginUser(btn.textContent || ""); });
  });
  overlay.querySelectorAll<HTMLButtonElement>(".user-del-btn").forEach(delBtn => {
    delBtn.addEventListener("click", (e) => { e.stopPropagation(); const name = delBtn.dataset.username || ""; deleteUser(name); document.body.removeChild(overlay); showLoginModal(); });
  });
  overlay.querySelector("#loginNewBtn")?.addEventListener("click", () => {
    const input = overlay.querySelector<HTMLInputElement>("#newUsername");
    const name = input?.value.trim(); if (!name) return;
    document.body.removeChild(overlay); loginUser(name);
  });
  overlay.querySelector("#loginCancelBtn")?.addEventListener("click", () => { document.body.removeChild(overlay); if (currentUser) updateLoginUI(); });
  overlay.addEventListener("click", (e) => { if (e.target === overlay) { document.body.removeChild(overlay); if (currentUser) updateLoginUI(); } });
}

function showExportModal(): void {
  if (!result) return;
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.innerHTML = `<div class="modal export-modal"><h3>导出选项</h3><div class="export-checks"><label class="check-label"><input type="checkbox" id="expPareto" checked /> 停机柏拉图</label><label class="check-label"><input type="checkbox" id="expMttr" checked /> MTTR / MTBF</label><label class="check-label"><input type="checkbox" id="expTrend" checked /> 故障推移</label><label class="check-label"><input type="checkbox" id="expData" checked /> 数据明细</label></div><label style="margin-top:10px;">图表月份<select id="expMonth">${["合计", ...result.months].map(m => `<option value="${m}" ${m === selectedMonth ? "selected" : ""}>${m}</option>`).join("")}</select></label><div class="export-actions"><button id="exportDoBtn" class="primary">导出 Excel</button><button id="exportCancelBtn" class="secondary">取消</button></div></div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("#exportDoBtn")?.addEventListener("click", async () => {
    const expMonth = (overlay.querySelector<HTMLSelectElement>("#expMonth"))?.value || "合计";
    const opts: ExportOptions = {
      pareto: (overlay.querySelector<HTMLInputElement>("#expPareto"))?.checked ?? true,
      mttr: (overlay.querySelector<HTMLInputElement>("#expMttr"))?.checked ?? true,
      trend: (overlay.querySelector<HTMLInputElement>("#expTrend"))?.checked ?? true,
      data: (overlay.querySelector<HTMLInputElement>("#expData"))?.checked ?? true,
      month: expMonth
    };
    document.body.removeChild(overlay);
    if (!result) return;
    try { await exportFullReport(result, config, opts); showToast("✓ 导出成功"); }
    catch (err) { console.error("Export failed:", err); showToast("导出失败，请重试"); }
  });
  overlay.querySelector("#exportCancelBtn")?.addEventListener("click", () => document.body.removeChild(overlay));
  overlay.addEventListener("click", (e) => { if (e.target === overlay) document.body.removeChild(overlay); });
}

// ===== Edit modal =====
function showEditModal(record?: FaultRecord): void {
  const isEdit = !!record;
  const overlay = document.createElement("div");
  overlay.className = "modal-overlay";
  overlay.innerHTML = `<div class="modal edit-modal"><h3>${isEdit ? "修改记录" : "新增记录"}</h3><div class="edit-form"><label>日期<input id="editDate" value="${escapeHtml(record?.date ?? "")}" /></label><label>线体<input id="editLine" value="${escapeHtml(record?.line ?? "")}" /></label><label>起始时间<input id="editStartTime" value="${escapeHtml(record?.startTime ?? "")}" /></label><label>截止时间<input id="editEndTime" value="${escapeHtml(record?.endTime ?? "")}" /></label><label>停机时长(min)<input id="editDowntime" type="number" step="0.1" value="${record?.downtime ?? ""}" /></label><label>设备类型<input id="editMachineType" value="${escapeHtml(record?.machineType ?? "")}" /></label><label>机器<input id="editMachine" value="${escapeHtml(record?.machine ?? "")}" /></label><label>问题描述<textarea id="editDescription">${escapeHtml(record?.description ?? "")}</textarea></label><label>责任部门<input id="editDepartment" value="${escapeHtml(record?.department ?? "")}" /></label><label>责任人<input id="editOwner" value="${escapeHtml(record?.owner ?? "")}" /></label></div><div class="edit-actions"><button id="editSaveBtn" class="primary">保存</button><button id="editCancelBtn" class="secondary">取消</button></div></div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("#editCancelBtn")?.addEventListener("click", () => document.body.removeChild(overlay));
  overlay.addEventListener("click", (e) => { if (e.target === overlay) document.body.removeChild(overlay); });
  overlay.querySelector("#editSaveBtn")?.addEventListener("click", () => {
    const read = (id: string) => (overlay.querySelector<HTMLInputElement | HTMLTextAreaElement>(`#${id}`))?.value ?? "";
    const newRec: FaultRecord = {
      id: record?.id ?? `manual-${Date.now()}`,
      sheet: record?.sheet ?? "",
      date: read("editDate"),
      line: read("editLine"),
      startTime: read("editStartTime"),
      endTime: read("editEndTime"),
      downtime: Number(read("editDowntime")) || 0,
      machine: read("editMachine"),
      machineType: read("editMachineType"),
      description: read("editDescription"),
      department: read("editDepartment"),
      owner: read("editOwner"),
      highlighted: record?.highlighted ?? false
    };
    if (!result) return;
    if (isEdit) {
      const idx = result.records.findIndex(r => r.id === record!.id);
      if (idx >= 0) result.records[idx] = newRec;
    } else {
      result.records.unshift(newRec);
    }
    document.body.removeChild(overlay);
    renderResult();
    showToast(isEdit ? "✓ 修改成功" : "✓ 新增成功");
  });
}

// ======= init app =======
app.innerHTML = `<header class="topbar"><div><h1>设备故障数据分析</h1><p>上传 Excel，配置筛选关键词，生成明细、柏拉图、MTTR/MTBF 和故障推移。</p></div><div class="topbar-right"><span id="loginArea"></span><button id="exportBtn" class="primary" disabled>导出 Excel</button></div></header><main class="layout"><aside class="sidebar"><nav class="side-nav" aria-label="功能导航"><div class="side-brand"><span>设</span><strong>故障分析</strong></div><button id="sidebarToggle" class="sidebar-toggle" aria-expanded="true" aria-label="收起规则">☰</button><button class="side-nav-item active" data-page="mainPage"><span class="nav-icon">主</span><span>首页</span></button><button class="side-nav-item" data-page="basicPanel"><span class="nav-icon">基</span><span>基础规则</span></button><button class="side-nav-item" data-page="keywordPanel"><span class="nav-icon">关</span><span>关键词</span></button><button class="side-nav-item" data-page="classPanel"><span class="nav-icon">分</span><span>分类规则</span></button></nav></aside><button id="sidebarOpenToggle" class="sidebar-open-toggle" aria-label="展开规则">☰</button><section class="workspace"><section id="mainPage" class="page-view active"><section class="panel upload-panel main-upload"><div><h2>文件上传</h2><div id="fileMeta" class="muted">尚未上传文件</div></div><label class="dropzone"><input id="fileInput" type="file" accept=".xlsx,.xls" /><span>选择 Excel 文件</span><small>支持 L1/L2/L3/L4 工作表</small></label></section><div id="warnings" class="warnings"></div><section class="metrics"><div><span id="totalRecords">0</span><small>筛选记录</small></div><div><span id="totalDowntime">0</span><small>停机分钟</small></div><div><span id="totalTypes">0</span><small>设备类型</small></div><div><span id="totalMonths">0</span><small>月份</small></div></section><section class="panel result-panel"><div class="result-head"><h2>图表分析</h2><select id="chartMonth"></select><button id="refreshChartsBtn" class="secondary" style="margin-left:8px;">刷新图表</button></div><div class="tabs"><button data-chart="pareto" class="active">停机柏拉图</button><button data-chart="mttr">MTTR/MTBF</button><button data-chart="trend">故障推移</button></div><div id="paretoChart" class="chart active"></div><div id="mttrChart" class="chart"></div><div id="trendChart" class="chart"></div></section><section class="panel result-panel"><div class="result-head"><h2>筛选明细</h2><div class="table-tools"><input id="tableSearch" placeholder="搜索问题、机器、责任人" /><select id="monthFilter"></select><select id="typeFilter"></select><select id="lineFilter"><option value="全部">全部线体</option></select><select id="sortSelect"><option value="downtime">按停机时长</option><option value="date">按日期</option><option value="machineType">按设备类型</option></select><select id="sortOrderSelect"><option value="desc">降序</option><option value="asc">升序</option></select></div></div><div class="data-toolbar" id="dataToolbar"><label class="check-label"><input type="checkbox" id="selectAllCheck" /> 全选</label><span id="selectionCount" class="muted"></span><button id="addRecordBtn" class="secondary">新增</button><button id="editRecordBtn" class="secondary" disabled>修改</button><button id="deleteRecordsBtn" class="secondary danger-btn" disabled>删除</button></div><div class="table-wrap"><table><thead><tr><th style="width:36px;"></th><th>日期</th><th>线体</th><th>时间</th><th>停机</th><th>设备类型</th><th>问题描述</th><th>责任</th></tr></thead><tbody id="recordRows"></tbody></table></div><div class="pagination" id="pagination"></div></section></section><section id="basicPanel" class="page-view panel config-workspace"><div class="result-head"><h2>基础规则</h2></div><div class="form-grid"><label>工作表<input id="sheetsInput" /></label><label>责任部门包含<input id="departmentInput" /></label><label>最小停机时长(min)<input id="minDowntimeInput" type="number" min="0" step="1" /></label><label>最大停机时长(min)<input id="maxDowntimeInput" type="number" min="0" step="1" placeholder="不限制" /></label></div><div style="margin-top:14px;text-align:right;"><button id="saveBasicBtn" class="primary">保存并应用</button></div></section><section id="keywordPanel" class="page-view panel config-workspace"><div class="result-head"><h2>关键词</h2></div><label><span class="label-row">高亮关键词<span class="field-note">在故障内容中高亮以下关键词</span></span><textarea id="highlightInput"></textarea></label><div class="mini-title-row"><div class="mini-title">包含规则</div><button id="addIncludeBtn" class="icon-btn" title="新增包含规则">+</button></div><div id="includeRules"></div><div class="mini-title-row"><div class="mini-title">排除规则</div><button id="addExcludeBtn" class="icon-btn" title="新增排除规则">+</button></div><div id="excludeRules"></div><div style="margin-top:14px;text-align:right;"><button id="saveKeywordBtn" class="primary">保存并应用</button></div></section><section id="classPanel" class="page-view panel config-workspace"><div class="result-head"><h2>分类规则</h2></div><div class="section-title"><span></span><button id="addClassBtn" class="icon-btn" title="新增分类">+</button></div><div class="class-header"><span>机器</span><span>关键词</span></div><div id="classRules"></div><div style="margin-top:14px;text-align:right;"><button id="saveClassBtn" class="primary">保存并应用</button></div></section></section></main>`;

bindEvents();
renderRuleEditors();
renderEmptyState();
updateSidebarState();
updatePageView();
updateLoginUI();

const lastUser = localStorage.getItem(CURRENT_USER_KEY);
if (lastUser && loadConfig(lastUser)) loginUser(lastUser);

function bindEvents(): void {
  document.querySelector("#sidebarToggle")?.addEventListener("click", () => { sidebarOpen = !sidebarOpen; updateSidebarState(); });
  document.querySelector("#sidebarOpenToggle")?.addEventListener("click", () => { sidebarOpen = true; updateSidebarState(); });
  document.querySelectorAll<HTMLButtonElement>(".side-nav-item").forEach((button) => {
    button.addEventListener("click", () => { sidebarOpen = true; updateSidebarState(); activePage = button.dataset.page ?? "mainPage"; updatePageView(); });
  });

  document.querySelector<HTMLInputElement>("#fileInput")?.addEventListener("change", async (event) => {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    if (!file.name.match(/\.xlsx?$/i)) { setWarnings(["请上传 .xlsx 或 .xls 文件"]); return; }
    const buffer = await file.arrayBuffer();
    workbook = XLSX.read(buffer, { type: "array" });
    document.querySelector("#fileMeta")!.textContent = `${file.name} · ${workbook.SheetNames.length} 个工作表`;
    runAnalysis();
  });

  bindInputNoRun("#sheetsInput", (value) => { config.sheets = splitList(value); });
  bindInputNoRun("#departmentInput", (value) => { config.departmentFilter = value.trim(); });
  bindInputNoRun("#minDowntimeInput", (value) => { config.minDowntime = Number(value) || 0; });
  bindInputNoRun("#maxDowntimeInput", (value) => { config.maxDowntime = Number(value) || 0; });
  bindInputNoRun("#highlightInput", (value) => { config.highlightKeywords = splitList(value); });

  document.querySelector("#saveBasicBtn")?.addEventListener("click", () => { saveConfig(); runAnalysis(); showToast("✓ 保存成功"); });
  document.querySelector("#saveKeywordBtn")?.addEventListener("click", () => { saveConfig(); runAnalysis(); showToast("✓ 保存成功"); });
  document.querySelector("#saveClassBtn")?.addEventListener("click", () => { saveConfig(); runAnalysis(); showToast("✓ 保存成功"); });

  document.querySelector("#addIncludeBtn")?.addEventListener("click", () => { config.includeKeywords.push({ fields: ["description"], keywords: [], matchMode: "any" }); renderRuleEditors(); });
  document.querySelector("#addExcludeBtn")?.addEventListener("click", () => { config.excludeKeywords.push({ fields: ["description"], keywords: [], matchMode: "any" }); renderRuleEditors(); });
  document.querySelector("#addClassBtn")?.addEventListener("click", () => { config.classificationRules.push({ type: "新分类", keywords: [] }); renderRuleEditors(); });
  document.querySelector("#exportBtn")?.addEventListener("click", () => { if (result) showExportModal(); });

  // Data toolbar
  document.querySelector("#selectAllCheck")?.addEventListener("change", (e) => toggleSelectAll((e.target as HTMLInputElement).checked));
  document.querySelector("#addRecordBtn")?.addEventListener("click", () => showEditModal());
  document.querySelector("#editRecordBtn")?.addEventListener("click", () => { if (selectedIds.size === 1 && result) { const rec = result.records.find(r => r.id === [...selectedIds][0]); if (rec) showEditModal(rec); } });
  document.querySelector("#deleteRecordsBtn")?.addEventListener("click", () => deleteSelectedRecords());

  document.querySelector("#chartMonth")?.addEventListener("change", (event) => { selectedMonth = (event.target as HTMLSelectElement).value; if (result) renderActiveChartOnly(); });
  document.querySelector("#refreshChartsBtn")?.addEventListener("click", () => { if (result) { _invalidateFilterCache(); recalcSummaries(); renderActiveChartOnly(); showToast("✓ 图表已刷新"); } });
  document.querySelectorAll<HTMLButtonElement>(".tabs button").forEach((button) => {
    button.addEventListener("click", () => { activeChart = button.dataset.chart ?? "pareto"; updateChartVisibility(); renderActiveChartOnly(); });
  });

  // 搜索防抖 300ms —— 避免每次按键都重渲整张表
  let searchTimer: ReturnType<typeof setTimeout> | null = null;
  const searchInput = document.querySelector<HTMLInputElement>("#tableSearch");
  searchInput?.addEventListener("input", (event) => {
    const value = (event.target as HTMLInputElement).value.trim();
    if (searchTimer) clearTimeout(searchTimer);
    searchTimer = setTimeout(() => { searchText = value; currentPage = 1; _invalidateFilterCache(); renderTable(); }, 300);
  });
  
  document.querySelector("#monthFilter")?.addEventListener("change", (event) => { monthFilter = (event.target as HTMLSelectElement).value; currentPage = 1; selectedIds.clear(); _invalidateFilterCache(); renderTable(); });
  document.querySelector("#typeFilter")?.addEventListener("change", (event) => { typeFilter = (event.target as HTMLSelectElement).value; currentPage = 1; selectedIds.clear(); _invalidateFilterCache(); renderTable(); });
  document.querySelector("#lineFilter")?.addEventListener("change", (event) => { lineFilter = (event.target as HTMLSelectElement).value; currentPage = 1; selectedIds.clear(); _invalidateFilterCache(); renderTable(); });
  document.querySelector("#sortSelect")?.addEventListener("change", (event) => { sortKey = (event.target as HTMLSelectElement).value as SortKey; currentPage = 1; _invalidateFilterCache(); renderTable(); });
  document.querySelector("#sortOrderSelect")?.addEventListener("change", (event) => { sortOrder = (event.target as HTMLSelectElement).value as "asc" | "desc"; currentPage = 1; _invalidateFilterCache(); renderTable(); });
  window.addEventListener("resize", resizeCharts);

  // 行 checkbox 事件委托 —— 只绑定一次
  const tableBody = document.querySelector<HTMLTableSectionElement>("#recordRows");
  tableBody?.addEventListener("change", (event) => {
    const target = event.target as HTMLInputElement;
    if (target.classList.contains("row-check")) {
      const id = target.dataset.id ?? "";
      if (target.checked) selectedIds.add(id); else selectedIds.delete(id);
      renderTable();
    }
  });
}

function bindInputNoRun(selector: string, handler: (value: string) => void): void {
  document.querySelector<HTMLInputElement | HTMLTextAreaElement>(selector)?.addEventListener("input", (event) => { handler((event.target as HTMLInputElement).value); });
}

function runAnalysis(): void {
  _invalidateFilterCache();
  if (!workbook) { renderEmptyState(); return; }
  result = analyzeWorkbook(workbook, config);
  currentPage = 1;
  selectedIds.clear();
  renderResult();
}

// ===== Data ops =====

function _invalidateFilterCache(): void { _filteredCache = null; _filteredCacheKey = ""; }

function toggleSelectAll(checked: boolean): void {
  if (!result) return;
  const filtered = getFilteredRows();
  if (checked) filtered.forEach(r => selectedIds.add(r.id));
  else selectedIds.clear();
  renderTable();
}

function deleteSelectedRecords(): void {
  if (!result || selectedIds.size === 0) return;
  if (!confirm(`确定删除选中的 ${selectedIds.size} 条记录吗？`)) return;
  result.records = result.records.filter(r => !selectedIds.has(r.id));
  selectedIds.clear();
  currentPage = 1;
  _invalidateFilterCache();
  renderResult();
  showToast("✓ 删除成功");
}

function getFilteredRows(): FaultRecord[] {
  if (!result) return [];
  const cacheKey = `${monthFilter}|${typeFilter}|${lineFilter}|${searchText}|${sortKey}|${sortOrder}|${result.records.length}`;
  if (_filteredCache && _filteredCacheKey === cacheKey) return _filteredCache;
  let filtered = result.records
    .filter(r => monthFilter === "全部" || r.date.startsWith(monthFilter))
    .filter(r => typeFilter === "全部" || r.machineType === typeFilter)
    .filter(r => lineFilter === "全部" || r.line === lineFilter)
    .filter(r => !searchText || [r.description, r.machine, r.owner, r.department].some(v => v.includes(searchText)))
    .sort((a, b) => {
      const mul = sortOrder === "asc" ? 1 : -1;
      if (sortKey === "date") return mul * a.date.localeCompare(b.date);
      if (sortKey === "machineType") return mul * a.machineType.localeCompare(b.machineType);
      return mul * (a.downtime - b.downtime);
    });
  _filteredCache = filtered;
  _filteredCacheKey = cacheKey;
  return filtered;
}

function recalcSummaries(): void {
  if (!result || !workbook) return;
  result = analyzeWorkbook(workbook, config);
}

// ===== Rule editors =====
function renderRuleEditors(): void {
  renderKeywordRules("includeRules", config.includeKeywords, "include");
  renderKeywordRules("excludeRules", config.excludeKeywords, "exclude");
  renderClassRules();
}

function renderKeywordRules(containerId: string, rules: KeywordRule[], kind: "include" | "exclude"): void {
  const container = document.querySelector<HTMLDivElement>(`#${containerId}`);
  if (!container) return;
  container.innerHTML = rules.map((rule, index) => `<div class="rule-row"><select data-kind="${kind}" data-index="${index}" data-prop="field">${availableKeywordFields.map(f => `<option value="${f.key}" ${rule.fields.includes(f.key) ? "selected" : ""}>${f.label}</option>`).join("")}</select><select data-kind="${kind}" data-index="${index}" data-prop="mode"><option value="any" ${rule.matchMode === "any" ? "selected" : ""}>任一命中</option><option value="all" ${rule.matchMode === "all" ? "selected" : ""}>全部命中</option></select><input data-kind="${kind}" data-index="${index}" data-prop="keywords" value="${escapeHtml(rule.keywords.join(","))}" placeholder="关键词，逗号分隔" /><button data-kind="${kind}" data-index="${index}" data-prop="delete" class="icon-btn danger" title="删除">×</button></div>`).join("");
  container.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLButtonElement>("[data-kind]").forEach((el) => { el.addEventListener("input", updateKeywordRule); el.addEventListener("click", updateKeywordRule); });
}

function updateKeywordRule(event: Event): void {
  const target = event.target as HTMLInputElement | HTMLSelectElement | HTMLButtonElement;
  const kind = target.dataset.kind as "include" | "exclude";
  const index = Number(target.dataset.index);
  const prop = target.dataset.prop;
  const rules = kind === "include" ? config.includeKeywords : config.excludeKeywords;
  if (prop === "delete") { rules.splice(index, 1); renderRuleEditors(); return; }
  if (prop === "field") rules[index].fields = [(target as HTMLSelectElement).value];
  if (prop === "mode") rules[index].matchMode = (target as HTMLSelectElement).value as "any" | "all";
  if (prop === "keywords") rules[index].keywords = splitList((target as HTMLInputElement).value);
}

function renderClassRules(): void {
  const container = document.querySelector<HTMLDivElement>("#classRules");
  if (!container) return;
  container.innerHTML = config.classificationRules.map((rule, index) => `<div class="class-row"><input data-class-index="${index}" data-prop="type" value="${escapeHtml(rule.type)}" /><textarea data-class-index="${index}" data-prop="keywords">${escapeHtml(rule.keywords.join(","))}</textarea><button data-class-index="${index}" data-prop="delete" class="icon-btn danger" title="删除">×</button></div>`).join("");
  container.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLButtonElement>("[data-class-index]").forEach((el) => { el.addEventListener("input", updateClassRule); el.addEventListener("click", updateClassRule); });
}

function updateClassRule(event: Event): void {
  const target = event.target as HTMLInputElement | HTMLTextAreaElement | HTMLButtonElement;
  const index = Number(target.dataset.classIndex);
  const prop = target.dataset.prop as keyof ClassificationRule | "delete";
  if (prop === "delete") { config.classificationRules.splice(index, 1); renderRuleEditors(); return; }
  if (prop === "type") config.classificationRules[index].type = (target as HTMLInputElement).value.trim();
  if (prop === "keywords") config.classificationRules[index].keywords = splitList((target as HTMLTextAreaElement).value);
}

// ===== Result rendering =====
function renderResult(): void {
  if (!result) return;
  setWarnings(result.warnings);
  document.querySelector("#exportBtn")?.removeAttribute("disabled");
  document.querySelector("#totalRecords")!.textContent = String(result.records.length);
  document.querySelector("#totalDowntime")!.textContent = String(Math.round(result.records.reduce((sum, r) => sum + r.downtime, 0)));
  document.querySelector("#totalTypes")!.textContent = String(result.typeSummary.length);
  document.querySelector("#totalMonths")!.textContent = String(result.months.length);
  fillSelect("#chartMonth", ["合计", ...result.months], selectedMonth);
  fillSelect("#monthFilter", ["全部", ...result.months], monthFilter);
  fillSelect("#typeFilter", ["全部", ...result.typeSummary.map(r => r.type)], typeFilter);
  // 填充线体筛选
  const lines = [...new Set(result.records.map(r => r.line).filter(Boolean))].sort();
  fillSelect("#lineFilter", ["全部", ...lines], lineFilter);
  _invalidateFilterCache();
  renderActiveChartOnly();
  updateChartVisibility();
  renderTable();
}

function renderTable(): void {
  const tbody = document.querySelector<HTMLTableSectionElement>("#recordRows");
  if (!tbody || !result) return;

  const filtered = getFilteredRows();
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  if (currentPage > totalPages) currentPage = totalPages;
  const pageRows = filtered.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  const allChecked = filtered.length > 0 && filtered.every(r => selectedIds.has(r.id));

  // 使用 DocumentFragment 批量插入，减少 DOM 重排
  const frag = document.createDocumentFragment();
  const tbodyNew = document.createElement("tbody");
  tbodyNew.id = "recordRows";
  tbodyNew.innerHTML = pageRows.length ? pageRows.map(r => `<tr class="${r.highlighted ? "highlight" : ""}"><td><input type="checkbox" class="row-check" data-id="${r.id}" ${selectedIds.has(r.id) ? "checked" : ""} /></td><td>${escapeHtml(r.date)}</td><td>${escapeHtml(r.line)}</td><td>${escapeHtml(r.startTime)}-${escapeHtml(r.endTime)}</td><td><span class="${downtimeClass(r.downtime)}">${r.downtime}</span></td><td>${escapeHtml(r.machineType)}</td><td>${escapeHtml(r.description)}</td><td>${escapeHtml(r.department)} / ${escapeHtml(r.owner)}</td></tr>`).join("") : `<tr><td colspan="8" class="empty">暂无符合条件的数据</td></tr>`;
  tbody.replaceWith(tbodyNew);
  // 重新绑定事件委托
  tbodyNew.addEventListener("change", (event) => {
    const target = event.target as HTMLInputElement;
    if (target.classList.contains("row-check")) {
      const id = target.dataset.id ?? "";
      if (target.checked) selectedIds.add(id); else selectedIds.delete(id);
      renderTable();
    }
  });

  // Selection state
  const selectAll = document.querySelector<HTMLInputElement>("#selectAllCheck");
  if (selectAll) selectAll.checked = allChecked;
  const selCount = document.querySelector("#selectionCount");
  if (selCount) selCount.textContent = selectedIds.size > 0 ? `已选 ${selectedIds.size} 条` : "";
  const editBtn = document.querySelector<HTMLButtonElement>("#editRecordBtn");
  const delBtn = document.querySelector<HTMLButtonElement>("#deleteRecordsBtn");
  if (editBtn) editBtn.disabled = selectedIds.size !== 1;
  if (delBtn) delBtn.disabled = selectedIds.size === 0;

  // Pagination
  renderPagination(totalPages);
}

function renderPagination(totalPages: number): void {
  const pg = document.querySelector("#pagination");
  if (!pg) return;

  const pageButtons: string[] = [];
  const maxVisible = 7;
  let start = Math.max(1, currentPage - Math.floor(maxVisible / 2));
  let end = Math.min(totalPages, start + maxVisible - 1);
  if (end - start + 1 < maxVisible) start = Math.max(1, end - maxVisible + 1);

  for (let i = start; i <= end; i++) {
    pageButtons.push(`<button class="pg-btn ${i === currentPage ? "pg-active" : ""}" data-page="${i}">${i}</button>`);
  }

  pg.innerHTML = `
    <div class="pg-left">
      <button class="pg-btn" data-page="${currentPage - 1}" ${currentPage <= 1 ? "disabled" : ""}>上一页</button>
      ${pageButtons.join("")}
      <button class="pg-btn" data-page="${currentPage + 1}" ${currentPage >= totalPages ? "disabled" : ""}>下一页</button>
      <span class="pg-info">共 ${totalPages} 页</span>
    </div>
    <div class="pg-right">
      <span class="pg-info">每页</span>
      <select id="pageSizeSelect">
        ${[20, 50, 100].map(s => `<option value="${s}" ${s === pageSize ? "selected" : ""}>${s}</option>`).join("")}
      </select>
      <span class="pg-info">条</span>
    </div>
  `;

  pg.querySelectorAll<HTMLButtonElement>(".pg-btn[data-page]").forEach(btn => {
    btn.addEventListener("click", () => {
      const p = Number(btn.dataset.page);
      if (p >= 1 && p <= totalPages) { currentPage = p; renderTable(); }
    });
  });
  pg.querySelector("#pageSizeSelect")?.addEventListener("change", (e) => {
    pageSize = Number((e.target as HTMLSelectElement).value);
    currentPage = 1;
    renderTable();
  });
}

function renderEmptyState(): void {
  setWarnings([]);
  document.querySelector("#totalRecords")!.textContent = "0";
  document.querySelector("#totalDowntime")!.textContent = "0";
  document.querySelector("#totalTypes")!.textContent = "0";
  document.querySelector("#totalMonths")!.textContent = "0";
  document.querySelector<HTMLTableSectionElement>("#recordRows")!.innerHTML = `<tr><td colspan="8" class="empty">上传文件后显示筛选明细</td></tr>`;
  const pg = document.querySelector("#pagination");
  if (pg) pg.innerHTML = "";
}

function setWarnings(warnings: string[]): void {
  const el = document.querySelector<HTMLDivElement>("#warnings");
  if (!el) return;
  el.innerHTML = warnings.length ? warnings.map(w => `<span>${escapeHtml(w)}</span>`).join("") : "";
}

function fillSelect(selector: string, values: string[], selected: string): void {
  const select = document.querySelector<HTMLSelectElement>(selector);
  if (!select) return;
  const nextSelected = values.includes(selected) ? selected : values[0];
  select.innerHTML = values.map(v => `<option value="${escapeHtml(v)}" ${v === nextSelected ? "selected" : ""}>${escapeHtml(v)}</option>`).join("");
  if (selector === "#chartMonth") selectedMonth = nextSelected;
  if (selector === "#monthFilter") monthFilter = nextSelected;
  if (selector === "#typeFilter") typeFilter = nextSelected;
  if (selector === "#lineFilter") lineFilter = nextSelected;
}

function renderActiveChartOnly(): void {
  if (!result) return;
  const typeData = selectedMonth === "合计" ? result.typeSummary : result.typeSummaryByMonth[selectedMonth] ?? [];
  if (activeChart === "pareto") {
    const p = getChartInstance("pareto");
    if (p) renderParetoInline(p, typeData, selectedMonth);
  } else if (activeChart === "mttr") {
    const m = getChartInstance("mttr");
    if (m) renderMttrInline(m, typeData, selectedMonth);
  } else {
    const t = getChartInstance("trend");
    if (t) {
      if (selectedMonth === "合计") renderTrendInline(t, result);
      else renderDailyTrendInline(t, result, selectedMonth);
    }
  }
}

function updateChartVisibility(): void {
  document.querySelectorAll(".tabs button").forEach(b => b.classList.toggle("active", (b as HTMLButtonElement).dataset.chart === activeChart));
  document.querySelectorAll(".chart").forEach(c => c.classList.toggle("active", c.id.startsWith(activeChart)));
  resizeCharts();
}

function updateSidebarState(): void {
  const layout = document.querySelector<HTMLElement>(".layout");
  const button = document.querySelector<HTMLButtonElement>("#sidebarToggle");
  const openButton = document.querySelector<HTMLButtonElement>("#sidebarOpenToggle");
  layout?.classList.toggle("sidebar-collapsed", !sidebarOpen);
  if (button) { button.textContent = "☰"; button.setAttribute("aria-label", sidebarOpen ? "收起规则" : "展开规则"); button.setAttribute("aria-expanded", String(sidebarOpen)); }
  if (openButton) openButton.hidden = true;
  window.setTimeout(resizeCharts, 180);
}

function updatePageView(): void {
  document.querySelectorAll<HTMLElement>(".page-view").forEach(p => p.classList.toggle("active", p.id === activePage));
  document.querySelectorAll<HTMLButtonElement>(".side-nav-item").forEach(b => b.classList.toggle("active", b.dataset.page === activePage));
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
