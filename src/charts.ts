import * as echarts from "echarts";
import type { AnalysisResult, DailySummary, TypeSummary } from "./types";
import { buildDailySummary } from "./analysis";

let paretoChart: echarts.ECharts | null = null;
let mttrChart: echarts.ECharts | null = null;
let trendChart: echarts.ECharts | null = null;

export function getChartInstance(kind: "pareto" | "mttr" | "trend"): echarts.ECharts | null {
  if (kind === "pareto") return paretoChart;
  if (kind === "mttr") return mttrChart;
  return trendChart;
}

function _safeInit(el: HTMLDivElement, existing: echarts.ECharts | null): echarts.ECharts {
  if (existing && !existing.isDisposed()) return existing;
  if (existing) existing.dispose();
  const instance = echarts.init(el);
  if (el.offsetWidth > 0 && el.offsetHeight > 0) {
    instance.resize();
  }
  return instance;
}

/** 懒初始化 + 防dispose：始终返回可用的 echarts 实例 */
export function ensureChart(kind: "pareto" | "mttr" | "trend"): echarts.ECharts {
  if (kind === "pareto") {
    const el = document.querySelector<HTMLDivElement>("#paretoChart");
    if (el) paretoChart = _safeInit(el, paretoChart);
    return paretoChart!;
  }
  if (kind === "mttr") {
    const el = document.querySelector<HTMLDivElement>("#mttrChart");
    if (el) mttrChart = _safeInit(el, mttrChart);
    return mttrChart!;
  }
  const el = document.querySelector<HTMLDivElement>("#trendChart");
  if (el) trendChart = _safeInit(el, trendChart);
  return trendChart!;
}

export function renderCharts(result: AnalysisResult, selectedMonth: string): void {
  const paretoElement = document.querySelector<HTMLDivElement>("#paretoChart");
  const mttrElement = document.querySelector<HTMLDivElement>("#mttrChart");
  const trendElement = document.querySelector<HTMLDivElement>("#trendChart");
  if (!paretoElement || !mttrElement || !trendElement) return;

  paretoChart = _safeInit(paretoElement, paretoChart);
  mttrChart = _safeInit(mttrElement, mttrChart);
  trendChart = _safeInit(trendElement, trendChart);

  const typeData = selectedMonth === "合计" ? result.typeSummary : result.typeSummaryByMonth[selectedMonth] ?? [];
  renderParetoInline(paretoChart, typeData, selectedMonth);
  renderMttrInline(mttrChart, typeData, selectedMonth);
  if (selectedMonth === "合计") {
    renderTrendInline(trendChart, result);
  } else {
    renderDailyTrendInline(trendChart, result, selectedMonth);
  }
}

export function resizeCharts(): void {
  paretoChart?.resize();
  mttrChart?.resize();
  trendChart?.resize();
}

export function renderParetoInline(chart: echarts.ECharts, rows: TypeSummary[], label: string): void {
  chart.setOption({
    title: { text: `${label} 停机柏拉图`, left: 8, textStyle: { fontSize: 14 } },
    color: ["#2f80ed", "#f2994a"],
    tooltip: { trigger: "axis" },
    legend: { top: 24 },
    grid: { top: 72, left: 54, right: 58, bottom: 42 },
    xAxis: { type: "category", data: rows.map((row) => row.type), axisLabel: { interval: 0 } },
    yAxis: [
      { type: "value", name: "停机时长", axisLabel: { formatter: "{value} min" } },
      { type: "value", name: "累计占比", min: 0, max: 1, axisLabel: { formatter: (value: number) => `${Math.round(value * 100)}%` } }
    ],
    series: [
      {
        name: "停机时长(min)",
        type: "bar",
        data: rows.map((row) => row.downtime),
        label: { show: true, position: "top", fontSize: 10 }
      },
      {
        name: "累计占比",
        type: "line",
        yAxisIndex: 1,
        data: rows.map((row) => Number(row.cumulativeShare.toFixed(3))),
        label: { show: true, formatter: (item: { value: number }) => `${Math.round(item.value * 100)}%`, fontSize: 10 }
      }
    ]
  });
}

export function renderDailyTrendInline(chart: echarts.ECharts, result: AnalysisResult, month: string, deviceCount: number = 1): void {
  const daily = buildDailySummary(result.records, month, deviceCount);
  chart.setOption({
    title: { text: `${month} 每日故障推移`, left: 8, textStyle: { fontSize: 14 } },
    color: ["#eb5757", "#2f80ed"],
    tooltip: { trigger: "axis" },
    legend: { top: 24 },
    grid: { top: 72, left: 54, right: 48, bottom: 42 },
    xAxis: { type: "category", data: daily.map((row) => row.day.slice(8)), name: "日" },
    yAxis: [
      { type: "value", name: "数值" },
      { type: "value", name: "故障率(%)" }
    ],
    series: [
      { name: "故障率(%)", type: "bar", yAxisIndex: 1, data: daily.map((row) => row.faultRate), label: { show: true, position: "top", fontSize: 10 } },
      { name: "停机总时长(min)", type: "line", data: daily.map((row) => row.downtime) }
    ]
  });
}

export function renderMttrInline(chart: echarts.ECharts, rows: TypeSummary[], label: string): void {
  chart.setOption({
    title: { text: `${label} MTTR / MTBF`, left: 8, textStyle: { fontSize: 14 } },
    color: ["#27ae60", "#9b51e0"],
    tooltip: { trigger: "axis" },
    legend: { top: 24 },
    grid: { top: 72, left: 54, right: 48, bottom: 42 },
    xAxis: { type: "category", data: rows.map((row) => row.type), axisLabel: { interval: 0 } },
    yAxis: { type: "value", name: "小时 / 分钟" },
    series: [
      { name: "MTTR(min)", type: "line", data: rows.map((row) => row.mttr), label: { show: true, fontSize: 10 } },
      { name: "MTBF(h)", type: "line", data: rows.map((row) => row.mtbf), label: { show: true, fontSize: 10 } }
    ]
  });
}

export function renderTrendInline(chart: echarts.ECharts, result: AnalysisResult): void {
  chart.setOption({
    title: { text: "月度故障推移", left: 8, textStyle: { fontSize: 14 } },
    color: ["#eb5757", "#2f80ed"],
    tooltip: { trigger: "axis" },
    legend: { top: 24 },
    grid: { top: 72, left: 54, right: 48, bottom: 42 },
    xAxis: { type: "category", data: result.monthSummary.map((row) => row.month) },
    yAxis: [
      { type: "value", name: "数值" },
      { type: "value", name: "故障率(%)" }
    ],
    series: [
      { name: "故障率(%)", type: "bar", yAxisIndex: 1, data: result.monthSummary.map((row) => row.faultRate), label: { show: true, position: "top", fontSize: 10 } },
      { name: "停机总时长(min)", type: "line", data: result.monthSummary.map((row) => row.downtime) }
    ]
  });
}
