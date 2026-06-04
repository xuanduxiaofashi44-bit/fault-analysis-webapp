import type { AnalysisConfig, AnalysisResult } from "./types";
import { buildDailySummary } from "./analysis";
import * as XLSX from "xlsx";

export type ExportOptions = {
  pareto: boolean;
  mttr: boolean;
  trend: boolean;
  month: string;
  data: boolean;
};

// ===== 快速导出 (PNG图片) =====

export function exportDataOnly(result: AnalysisResult, _config: AnalysisConfig): void {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.records.map((record) => ({
    工作表: record.sheet, 日期: record.date, 线体: record.line,
    起始时间: record.startTime, 截止时间: record.endTime,
    "停机时长(min)": record.downtime, 原机器: record.machine,
    设备类型: record.machineType, 问题描述: record.description,
    责任部门: record.department, 责任人: record.owner
  }))), "筛选明细");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.typeSummary.map((row) => ({
    设备类型: row.type, 故障次数: row.count, "停机总时长(min)": row.downtime,
    占比: percent(row.share), 累计占比: percent(row.cumulativeShare),
    "MTTR(min)": row.mttr, "MTBF(h)": row.mtbf
  }))), "分类汇总");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.monthSummary.map((row) => ({
    月份: row.month, 天数: row.days, 故障次数: row.count,
    "停机总时长(min)": row.downtime, "日均故障时长(min/天)": row.dailyDowntime,
    "故障率(%)": row.faultRate, "MTTR(min)": row.mttr, "MTBF(h)": row.mtbf
  }))), "月度趋势");
  XLSX.writeFile(workbook, `设备故障分析_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function percent(value: number): string {
  return `${Math.round(value * 1000) / 10}%`;
}

// ===== ECharts-based chart image generation =====
async function renderChartImage(chartOption: object): Promise<Uint8Array | null> {
  try {
    const echarts = await import("echarts");
    const container = document.createElement("div");
    container.style.cssText = "position:fixed;left:-9999px;top:-9999px;width:800px;height:400px;";
    document.body.appendChild(container);
    const chart = echarts.init(container);
    chart.setOption({ ...chartOption, animation: false });
    await new Promise<void>((resolve) => {
      let done = false;
      chart.on("finished", () => { if (!done) { done = true; resolve(); } });
      setTimeout(() => { if (!done) { done = true; resolve(); } }, 2000);
    });
    const dataUrl = chart.getDataURL({ type: "png", pixelRatio: 2, backgroundColor: "#ffffff" });
    chart.dispose();
    document.body.removeChild(container);
    const base64 = dataUrl.split(",")[1];
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  } catch {
    return null;
  }
}

function buildParetoOption(rows: { type: string; downtime: number; cumulativeShare: number }[], title: string) {
  const categories = rows.map(r => r.type);
  return {
    title: { text: title, left: "center", textStyle: { fontSize: 14 } },
    tooltip: { trigger: "axis" },
    legend: { data: ["停机时长(min)", "累计占比"], bottom: 0, textStyle: { fontSize: 12 } },
    grid: { left: 60, right: 60, top: 50, bottom: 40 },
    xAxis: { type: "category", data: categories, axisLabel: { rotate: 30, fontSize: 10 } },
    yAxis: [
      { type: "value", name: "分钟", nameTextStyle: { fontSize: 11 } },
      { type: "value", name: "%", min: 0, max: 100, nameTextStyle: { fontSize: 11 } }
    ],
    series: [
      { name: "停机时长(min)", type: "bar", data: rows.map(r => r.downtime), itemStyle: { color: "#5470c6" }, barMaxWidth: 40 },
      { name: "累计占比", type: "line", yAxisIndex: 1, data: rows.map(r => +(r.cumulativeShare * 100).toFixed(1)), lineStyle: { color: "#e74c3c" }, itemStyle: { color: "#e74c3c" }, symbol: "circle" }
    ]
  };
}

function buildMttrOption(rows: { type: string; mttr: number; mtbf: number }[], title: string) {
  const categories = rows.map(r => r.type);
  return {
    title: { text: title, left: "center", textStyle: { fontSize: 14 } },
    tooltip: { trigger: "axis" },
    legend: { data: ["MTTR(min)", "MTBF(h)"], bottom: 0, textStyle: { fontSize: 12 } },
    grid: { left: 60, right: 60, top: 50, bottom: 40 },
    xAxis: { type: "category", data: categories, axisLabel: { rotate: 30, fontSize: 10 } },
    yAxis: { type: "value", name: "时间", nameTextStyle: { fontSize: 11 } },
    series: [
      { name: "MTTR(min)", type: "line", data: rows.map(r => +r.mttr.toFixed(1)), lineStyle: { color: "#5470c6" }, itemStyle: { color: "#5470c6" }, symbol: "circle" },
      { name: "MTBF(h)", type: "line", data: rows.map(r => +r.mtbf.toFixed(1)), lineStyle: { color: "#91cc75" }, itemStyle: { color: "#91cc75" }, symbol: "diamond" }
    ]
  };
}

function buildTrendOption(rows: { month: string; downtime: number; faultRate: number }[], title: string) {
  const categories = rows.map(r => r.month);
  return {
    title: { text: title, left: "center", textStyle: { fontSize: 14 } },
    tooltip: { trigger: "axis" },
    legend: { data: ["故障率(%)", "停机总时长(min)"], bottom: 0, textStyle: { fontSize: 10 } },
    grid: { left: 60, right: 60, top: 50, bottom: 40 },
    xAxis: { type: "category", data: categories, axisLabel: { fontSize: 10 } },
    yAxis: [
      { type: "value", name: "分钟/小时", nameTextStyle: { fontSize: 11 } },
      { type: "value", name: "%", nameTextStyle: { fontSize: 11 } }
    ],
    series: [
      { name: "故障率(%)", type: "bar", data: rows.map(r => +r.faultRate.toFixed(2)), itemStyle: { color: "#fac858" }, barMaxWidth: 30, yAxisIndex: 1 },
      { name: "停机总时长(min)", type: "line", data: rows.map(r => r.downtime), lineStyle: { color: "#5470c6" }, itemStyle: { color: "#5470c6" }, symbol: "circle" }
    ]
  };
}

export async function exportFullReport(result: AnalysisResult, _config: AnalysisConfig, opts: ExportOptions): Promise<void> {
  const ExcelJS = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  wb.creator = "故障分析系统";

  if (opts.data) {
    const ws1 = wb.addWorksheet("筛选明细");
    ws1.columns = [
      { header: "工作表", key: "sheet", width: 10 },
      { header: "日期", key: "date", width: 12 },
      { header: "线体", key: "line", width: 8 },
      { header: "起始时间", key: "startTime", width: 10 },
      { header: "截止时间", key: "endTime", width: 10 },
      { header: "停机时长(min)", key: "downtime", width: 14 },
      { header: "原机器", key: "machine", width: 12 },
      { header: "设备类型", key: "machineType", width: 12 },
      { header: "问题描述", key: "description", width: 40 },
      { header: "责任部门", key: "department", width: 10 },
      { header: "责任人", key: "owner", width: 8 }
    ];
    ws1.addRows(result.records);
    formatHeader(ws1);
  }

  const typeRows = result.typeSummary;
  const ws2 = wb.addWorksheet("分类汇总");
  ws2.columns = [
    { header: "设备类型", key: "type", width: 12 },
    { header: "故障次数", key: "count", width: 10 },
    { header: "停机总时长(min)", key: "downtime", width: 16 },
    { header: "占比", key: "share", width: 8 },
    { header: "累计占比", key: "cumulativeShare", width: 10 },
    { header: "MTTR(min)", key: "mttr", width: 10 },
    { header: "MTBF(h)", key: "mtbf", width: 10 }
  ];
  typeRows.forEach((r, _i) => ws2.addRow({ type: r.type, count: r.count, downtime: r.downtime, share: percent(r.share), cumulativeShare: r.cumulativeShare, mttr: r.mttr, mtbf: r.mtbf }));
  formatHeader(ws2);

  if (opts.pareto) {
    const paretoImg = await renderChartImage(buildParetoOption(typeRows, opts.month + " 停机柏拉图"));
    if (paretoImg) {
      const imgId = wb.addImage({ buffer: paretoImg as any, extension: "png" });
      ws2.addImage(imgId, { tl: { col: 9, row: 0 }, ext: { width: 480, height: 300 } });
    }
  }
  if (opts.mttr) {
    const mttrImg = await renderChartImage(buildMttrOption(typeRows, opts.month + " MTTR/MTBF"));
    if (mttrImg) {
      const imgId = wb.addImage({ buffer: mttrImg as any, extension: "png" });
      ws2.addImage(imgId, { tl: { col: 9, row: 22 }, ext: { width: 480, height: 300 } });
    }
  }

  const monthRows = result.monthSummary;
  const ws3 = wb.addWorksheet("月度趋势");
  ws3.columns = [
    { header: "月份", key: "month", width: 10 },
    { header: "天数", key: "days", width: 6 },
    { header: "故障次数", key: "count", width: 10 },
    { header: "停机总时长(min)", key: "downtime", width: 16 },
    { header: "日均故障时长", key: "dailyDowntime", width: 14 },
    { header: "故障率(%)", key: "faultRate", width: 12 },
    { header: "MTTR(min)", key: "mttr", width: 10 },
    { header: "MTBF(h)", key: "mtbf", width: 10 }
  ];
  monthRows.forEach(r => ws3.addRow(r));
  formatHeader(ws3);

  if (opts.trend) {
    const trendImg = await renderChartImage(buildTrendOption(monthRows, "月度故障推移"));
    if (trendImg) {
      const imgId = wb.addImage({ buffer: trendImg as any, extension: "png" });
      ws3.addImage(imgId, { tl: { col: 10, row: 0 }, ext: { width: 500, height: 320 } });
    }
  }

  if (opts.month !== "合计") {
    const dailyRows = buildDailySummary(result.records, opts.month);
    const ws4 = wb.addWorksheet(opts.month + " 每日趋势");
    ws4.columns = [
      { header: "日期", key: "day", width: 12 },
      { header: "故障次数", key: "count", width: 10 },
      { header: "停机总时长(min)", key: "downtime", width: 16 },
      { header: "故障率(%)", key: "faultRate", width: 12 },
      { header: "MTTR(min)", key: "mttr", width: 10 },
      { header: "MTBF(h)", key: "mtbf", width: 10 }
    ];
    dailyRows.forEach(r => ws4.addRow(r));
    formatHeader(ws4);
    const dailyTrendImg = await renderChartImage(buildTrendOption(
      dailyRows.map(r => ({ month: r.day, downtime: r.downtime, faultRate: r.faultRate })),
      opts.month + " 每日故障推移"
    ));
    if (dailyTrendImg) {
      const imgId = wb.addImage({ buffer: dailyTrendImg as any, extension: "png" });
      ws4.addImage(imgId, { tl: { col: 8, row: 0 }, ext: { width: 500, height: 320 } });
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  downloadBlob(buf, `设备故障分析_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function formatHeader(ws: any): void {
  if (!ws.getRow) return;
  const row = ws.getRow(1);
  row.font = { bold: true };
  row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0F4FA" } };
}

function downloadBlob(buf: ArrayBuffer, filename: string): void {
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// ===== GitHub Actions 原生图表导出 =====

const GH_API = "https://api.github.com";
const GH_REPO = "xuanduxiaofashi44-bit/fault-analysis-webapp";

interface ExportPayload {
  month: string;
  records: any[];
  typeSummary: any[];
  monthSummary: any[];
  dailySummary: any[];
}

async function ghFetch(token: string, path: string, options: RequestInit = {}): Promise<Response> {
  const res = await fetch(`${GH_API}${path}`, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `token ${token}`,
      Accept: "application/vnd.github.v3+json",
    },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`GitHub API ${res.status}: ${text.substring(0, 200)}`);
  }
  return res;
}

export async function exportViaGitHubActions(
  result: AnalysisResult,
  opts: ExportOptions,
  token: string,
  onStatus: (msg: string) => void
): Promise<void> {
  if (!token) throw new Error("请先在设置中填写 GitHub Token");

  // 1. Build payload
  const payload: ExportPayload = {
    month: opts.month,
    records: opts.data ? result.records : [],
    typeSummary: (opts.pareto || opts.mttr) ? result.typeSummary : [],
    monthSummary: opts.trend ? result.monthSummary : [],
    dailySummary: [],
  };
  if (opts.month !== "合计") {
    payload.dailySummary = buildDailySummary(result.records, opts.month);
  }

  // 2. Push data file to repo
  const ts = Date.now();
  const filePath = `data/exports/export_${ts}.json`;
  const jsonStr = JSON.stringify(payload, null, 2);
  const b64 = btoa(unescape(encodeURIComponent(jsonStr)));

  onStatus("正在上传数据...");
  await ghFetch(token, `/repos/${GH_REPO}/contents/${filePath}`, {
    method: "PUT",
    body: JSON.stringify({
      message: `export data ${ts}`,
      content: b64,
      branch: "main",
    }),
  });

  // 3. Trigger workflow
  onStatus("正在触发导出流程...");
  const wfResp = await ghFetch(token, `/repos/${GH_REPO}/actions/workflows`);
  const wfData = await wfResp.json();
  const exportWf = wfData.workflows.find((w: any) => w.name === "Export Excel with Charts" || w.path?.includes("export-excel"));
  if (!exportWf) throw new Error("未找到导出 workflow");

  const dispatchResp = await ghFetch(token, `/repos/${GH_REPO}/actions/workflows/${exportWf.id}/dispatches`, {
    method: "POST",
    body: JSON.stringify({
      ref: "main",
      inputs: { data_file: filePath },
    }),
  });

  // 4. Poll for the new workflow run
  onStatus("正在生成图表(约20-40秒)...");
  await new Promise(r => setTimeout(r, 3000));

  let runId: number | null = null;
  for (let attempt = 0; attempt < 30; attempt++) {
    const runsResp = await ghFetch(token, `/repos/${GH_REPO}/actions/workflows/${exportWf.id}/runs?per_page=5`);
    const runsData = await runsResp.json();
    const run = runsData.workflow_runs?.find((r: any) =>
      r.head_branch === "main" &&
      r.event === "workflow_dispatch" &&
      new Date(r.created_at).getTime() > ts - 60000
    );
    if (run) {
      runId = run.id;
      if (run.status === "completed") {
        if (run.conclusion === "success") break;
        throw new Error(`导出失败: ${run.conclusion}`);
      }
    }
    onStatus(`正在生成图表(${attempt * 3 + 3}秒)...`);
    await new Promise(r => setTimeout(r, 3000));
  }

  if (!runId) throw new Error("无法找到导出的 workflow 运行记录");

  // 5. Get artifact and download
  onStatus("正在下载文件...");
  const artResp = await ghFetch(token, `/repos/${GH_REPO}/actions/runs/${runId}/artifacts`);
  const artData = await artResp.json();
  const artifact = artData.artifacts?.find((a: any) => a.name === "设备故障分析_图表");
  if (!artifact) throw new Error("未找到导出产物");

  // Download via redirect
  const dlResp = await fetch(`${GH_API}/repos/${GH_REPO}/actions/artifacts/${artifact.id}/zip`, {
    headers: { Authorization: `token ${token}` },
    redirect: "follow",
  });
  if (!dlResp.ok) throw new Error("下载失败");

  const blob = await dlResp.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `设备故障分析_图表_${new Date().toISOString().slice(0, 10)}.zip`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
