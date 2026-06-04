import * as XLSX from "xlsx";
import type { AnalysisConfig, AnalysisResult } from "./types";
import { getChartInstance } from "./charts";

export type ExportOptions = {
  pareto: boolean;
  mttr: boolean;
  trend: boolean;
  month: string;
};

export function exportResult(result: AnalysisResult, config: AnalysisConfig): void {
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.records.map((record) => ({
    工作表: record.sheet,
    日期: record.date,
    线体: record.line,
    起始时间: record.startTime,
    截止时间: record.endTime,
    "停机时长(min)": record.downtime,
    原机器: record.machine,
    设备类型: record.machineType,
    问题描述: record.description,
    责任部门: record.department,
    责任人: record.owner
  }))), "筛选明细");

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.typeSummary.map((row) => ({
    设备类型: row.type,
    故障次数: row.count,
    "停机总时长(min)": row.downtime,
    占比: percent(row.share),
    累计占比: percent(row.cumulativeShare),
    "MTTR(min)": row.mttr,
    "MTBF(h)": row.mtbf
  }))), "分类汇总");

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(result.monthSummary.map((row) => ({
    月份: row.month,
    天数: row.days,
    故障次数: row.count,
    "停机总时长(min)": row.downtime,
    "日均故障时长(min/天)": row.dailyDowntime,
    "故障率(%)": row.faultRate,
    "MTTR(min)": row.mttr,
    "MTBF(h)": row.mtbf
  }))), "月度趋势");

  const configRows = [
    { 项目: "工作表", 值: config.sheets.join(",") },
    { 项目: "责任部门", 值: config.departmentFilter },
    { 项目: "最小停机时长", 值: String(config.minDowntime) },
    { 项目: "最大停机时长", 值: config.maxDowntime ? String(config.maxDowntime) : "不限制" },
    { 项目: "包含关键词", 值: JSON.stringify(config.includeKeywords) },
    { 项目: "排除关键词", 值: JSON.stringify(config.excludeKeywords) },
    { 项目: "高亮关键词", 值: config.highlightKeywords.join(",") },
    { 项目: "分类规则", 值: JSON.stringify(config.classificationRules) },
    { 项目: "字段映射", 值: JSON.stringify(config.columnMap) }
  ];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(configRows), "规则配置");

  XLSX.writeFile(workbook, `设备故障分析_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

export function downloadChartImage(kind: "pareto" | "mttr" | "trend", label: string): void {
  const chart = getChartInstance(kind);
  if (!chart) return;
  const dataUrl = chart.getDataURL({ type: "png", pixelRatio: 2, backgroundColor: "#fff" });
  const link = document.createElement("a");
  link.href = dataUrl;
  link.download = `设备故障分析_${label}_${kind}_${new Date().toISOString().slice(0, 10)}.png`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function percent(value: number): string {
  return `${Math.round(value * 1000) / 10}%`;
}
