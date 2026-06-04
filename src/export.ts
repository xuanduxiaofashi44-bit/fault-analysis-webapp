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

export function exportDataOnly(result: AnalysisResult, config: AnalysisConfig): void {
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

export async function exportFullReport(result: AnalysisResult, _config: AnalysisConfig, opts: ExportOptions): Promise<void> {
  const ExcelJS = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  wb.creator = "故障分析系统";

  // ===== Sheet 1: 筛选明细 =====
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

  // ===== Sheet 2: 分类汇总 + 柏拉图/MTTR数据 =====
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
  typeRows.forEach((r, i) => ws2.addRow({ type: r.type, count: r.count, downtime: r.downtime, share: percent(r.share), cumulativeShare: r.cumulativeShare, mttr: r.mttr, mtbf: r.mtbf }));
  formatHeader(ws2);

  const typeCount = typeRows.length;
  const lastRow = typeCount + 1;
  const catRange = `分类汇总!$A$2:$A$${lastRow}`;

  // Pareto chart
  if (opts.pareto) {
    addParetoChart(ws2, wb, typeCount, lastRow, `"${opts.month}" 停机柏拉图`);
  }
  // MTTR chart
  if (opts.mttr) {
    addMttrChart(ws2, wb, typeCount, lastRow, `"${opts.month}" MTTR/MTBF`);
  }

  // ===== Sheet 3: 月度趋势 =====
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

  const monthCount = monthRows.length;
  const monthLastRow = monthCount + 1;

  if (opts.trend) {
    addTrendChart(ws3, wb, monthCount, monthLastRow, "月度故障推移");
  }

  // ===== Sheet 4: 每日趋势 (仅当选择了具体月份) =====
  if (opts.month !== "合计") {
    const dailyRows = buildDailySummary(result.records, opts.month);
    const ws4 = wb.addWorksheet(`${opts.month} 每日趋势`);
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
    const dailyCount = dailyRows.length;
    const dailyLastRow = dailyCount + 1;
    addDailyTrendChart(ws4, wb, dailyCount, dailyLastRow, `${opts.month} 每日故障推移`);
  }

  const buf = await wb.xlsx.writeBuffer();
  downloadBlob(buf, `设备故障分析_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function addParetoChart(ws: any, wb: any, count: number, lastRow: number, title: string): void {
  const chart = wb.addChart ? wb.addChart('columnBar', []) : null;
  if (!chart && wb.createChart) {
    // fallback: try worksheet-level chart
  }
  // Use worksheet.addChart for embedded charts
  if (typeof ws.addChart !== 'function') return;
  try {
    // Bar + Line combo - exceljs uses 'bar' for bar charts
    const paretoChart = ws.addChart('bar', [], {
      title: { text: title },
      legend: { position: 'bottom' }
    });
    if (!paretoChart) return;
    // downtime bar
    paretoChart.addSeries({
      name: '停机时长(min)',
      categories: `分类汇总!$A$2:$A$${lastRow}`,
      values: `分类汇总!$C$2:$C$${lastRow}`,
      type: 'bar'
    });
    // cumulative share line on secondary axis
    paretoChart.addSeries({
      name: '累计占比',
      categories: `分类汇总!$A$2:$A$${lastRow}`,
      values: `分类汇总!$E$2:$E$${lastRow}`,
      type: 'line',
      secondaryAxis: true
    });
  } catch { /* chart creation may fail in some environments */ }
}

function addMttrChart(ws: any, wb: any, count: number, lastRow: number, title: string): void {
  if (typeof ws.addChart !== 'function') return;
  try {
    const chart = ws.addChart('line', [], {
      title: { text: title },
      legend: { position: 'bottom' }
    });
    if (!chart) return;
    chart.addSeries({
      name: 'MTTR(min)',
      categories: `分类汇总!$A$2:$A$${lastRow}`,
      values: `分类汇总!$F$2:$F$${lastRow}`
    });
    chart.addSeries({
      name: 'MTBF(h)',
      categories: `分类汇总!$A$2:$A$${lastRow}`,
      values: `分类汇总!$G$2:$G$${lastRow}`
    });
  } catch {}
}

function addTrendChart(ws: any, wb: any, count: number, lastRow: number, title: string): void {
  if (typeof ws.addChart !== 'function') return;
  try {
    const chart = ws.addChart('bar', [], {
      title: { text: title },
      legend: { position: 'bottom' }
    });
    if (!chart) return;
    chart.addSeries({
      name: '故障率(%)',
      categories: `月度趋势!$A$2:$A$${lastRow}`,
      values: `月度趋势!$F$2:$F$${lastRow}`,
      type: 'bar',
      secondaryAxis: true
    });
    chart.addSeries({
      name: '停机总时长(min)',
      categories: `月度趋势!$A$2:$A$${lastRow}`,
      values: `月度趋势!$D$2:$D$${lastRow}`,
      type: 'line'
    });
    chart.addSeries({
      name: 'MTTR(min)',
      categories: `月度趋势!$A$2:$A$${lastRow}`,
      values: `月度趋势!$G$2:$G$${lastRow}`,
      type: 'line'
    });
    chart.addSeries({
      name: 'MTBF(h)',
      categories: `月度趋势!$A$2:$A$${lastRow}`,
      values: `月度趋势!$H$2:$H$${lastRow}`,
      type: 'line'
    });
  } catch {}
}

function addDailyTrendChart(ws: any, wb: any, count: number, lastRow: number, title: string): void {
  if (typeof ws.addChart !== 'function') return;
  const sheetName = ws.name || '每日趋势';
  try {
    const chart = ws.addChart('bar', [], {
      title: { text: title },
      legend: { position: 'bottom' }
    });
    if (!chart) return;
    chart.addSeries({
      name: '故障率(%)',
      categories: `'${sheetName}'!$A$2:$A$${lastRow}`,
      values: `'${sheetName}'!$D$2:$D$${lastRow}`,
      type: 'bar',
      secondaryAxis: true
    });
    chart.addSeries({
      name: '停机总时长(min)',
      categories: `'${sheetName}'!$A$2:$A$${lastRow}`,
      values: `'${sheetName}'!$C$2:$C$${lastRow}`,
      type: 'line'
    });
    chart.addSeries({
      name: 'MTTR(min)',
      categories: `'${sheetName}'!$A$2:$A$${lastRow}`,
      values: `'${sheetName}'!$E$2:$E$${lastRow}`,
      type: 'line'
    });
    chart.addSeries({
      name: 'MTBF(h)',
      categories: `'${sheetName}'!$A$2:$A$${lastRow}`,
      values: `'${sheetName}'!$F$2:$F$${lastRow}`,
      type: 'line'
    });
  } catch {}
}

function formatHeader(ws: any): void {
  if (!ws.getRow) return;
  const row = ws.getRow(1);
  row.font = { bold: true };
  row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F4FA' } };
}

function downloadBlob(buf: ArrayBuffer, filename: string): void {
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
