import * as XLSX from "xlsx";
import type { AnalysisConfig, AnalysisResult, DailySummary, FaultRecord, KeywordRule, MonthSummary, TypeSummary } from "./types";

const fieldGetters: Record<string, (record: FaultRecord) => string> = {
  description: (record) => record.description,
  machine: (record) => record.machine,
  department: (record) => record.department,
  line: (record) => record.line,
  owner: (record) => record.owner
};

export function splitList(value: string): string[] {
  return value
    .split(/[,，\n]/)
    .map((item) => item.trim())
    .filter(Boolean);
}

export function daysInMonth(month: string): number {
  const [year, monthIndex] = month.split("-").map(Number);
  if (!year || !monthIndex) return 30;
  return new Date(year, monthIndex, 0).getDate();
}

export function normalizeDate(value: unknown): string {
  if (value instanceof Date) {
    return formatDate(value);
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return formatDate(parsed);
  return raw.slice(0, 10);
}

export function normalizeTime(value: unknown): string {
  if (value instanceof Date) {
    return `${String(value.getHours()).padStart(2, "0")}:${String(value.getMinutes()).padStart(2, "0")}`;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const totalMinutes = Math.round(value * 24 * 60);
    const hours = Math.floor(totalMinutes / 60) % 24;
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }
  return String(value ?? "").trim();
}

export function analyzeWorkbook(workbook: XLSX.WorkBook, config: AnalysisConfig): AnalysisResult {
  const warnings: string[] = [];
  const records: FaultRecord[] = [];

  for (const sheetName of config.sheets) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      warnings.push(`未找到工作表：${sheetName}`);
      continue;
    }

    const rows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: "" });
    if (rows.length < 2) {
      warnings.push(`工作表为空：${sheetName}`);
      continue;
    }

    const header = rows[0].map((cell) => String(cell ?? "").replace(/\n/g, "").trim());
    const required = Object.values(config.columnMap);
    const missing = required.filter((column) => !header.includes(column));
    if (missing.length) {
      warnings.push(`${sheetName} 缺少列：${missing.join("、")}`);
    }

    const indexOf = (key: string) => header.indexOf(config.columnMap[key] ?? key);
    const indexes = {
      date: indexOf("date"),
      line: indexOf("line"),
      startTime: indexOf("startTime"),
      endTime: indexOf("endTime"),
      downtime: indexOf("downtime"),
      machine: indexOf("machine"),
      description: indexOf("description"),
      department: indexOf("department"),
      owner: indexOf("owner")
    };

    rows.slice(1).forEach((row, index) => {
      const get = (key: keyof typeof indexes): unknown => {
        const column = indexes[key];
        return column >= 0 ? row[column] : "";
      };
      const downtime = Number.parseFloat(String(get("downtime") || "0")) || 0;
      const record: FaultRecord = {
        id: `${sheetName}-${index}`,
        sheet: sheetName,
        date: normalizeDate(get("date")),
        line: String(get("line") ?? "").trim(),
        startTime: normalizeTime(get("startTime")),
        endTime: normalizeTime(get("endTime")),
        downtime: Math.round(downtime * 10) / 10,
        machine: String(get("machine") ?? "").trim(),
        machineType: "",
        description: String(get("description") ?? "").trim(),
        department: String(get("department") ?? "").trim(),
        owner: String(get("owner") ?? "").trim(),
        highlighted: false
      };
      if (record.date || record.description || record.machine) records.push(record);
    });
  }

  const filtered = records
    .filter((record) => !config.departmentFilter || record.department.includes(config.departmentFilter))
    .filter((record) => record.downtime >= config.minDowntime)
    .filter((record) => !config.maxDowntime || record.downtime <= config.maxDowntime)
    .filter((record) => matchesRuleGroup(record, config.includeKeywords, true))
    .filter((record) => !matchesRuleGroup(record, config.excludeKeywords, false))
    .map((record) => {
      const machineType = classify(record, config);
      const highlighted = config.highlightKeywords.some((keyword) => keyword && record.description.includes(keyword));
      return { ...record, machineType, highlighted };
    });

  const deduped = dedupeRecords(filtered).sort((a, b) => b.downtime - a.downtime);
  const months = [...new Set(deduped.map((record) => record.date.slice(0, 7)).filter(Boolean))].sort();
  const typeSummary = buildTypeSummary(deduped);
  const typeSummaryByMonth = Object.fromEntries(months.map((month) => [month, buildTypeSummary(deduped.filter((record) => record.date.startsWith(month)), month)]));
  const dc = getDeviceCount(config);
  const monthSummary = buildMonthSummary(deduped, months, dc);

  return { records: deduped, warnings, months, typeSummary, typeSummaryByMonth, monthSummary };
}

function formatDate(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function matchesRuleGroup(record: FaultRecord, rules: KeywordRule[], emptyDefault: boolean): boolean {
  const activeRules = rules.filter((rule) => rule.fields.length && rule.keywords.length);
  if (!activeRules.length) return emptyDefault;
  return activeRules.some((rule) => {
    const haystack = rule.fields.map((field) => fieldGetters[field]?.(record) ?? "").join(" ");
    return rule.matchMode === "all"
      ? rule.keywords.every((keyword) => haystack.includes(keyword))
      : rule.keywords.some((keyword) => haystack.includes(keyword));
  });
}

function classify(record: { description: string; machine: string }, config: AnalysisConfig): string {
  const machineRules = config.classificationRules.filter(r => r.field === "machine");
  const descRules = config.classificationRules.filter(r => r.field !== "machine");
  for (const rules of [machineRules, descRules]) {
    for (const rule of rules) {
      const haystack = rule.field === "machine" ? (record.machine ?? "") : (record.description ?? "");
      if (rule.keywords.some((keyword) => keyword && haystack.includes(keyword))) {
        return rule.type;
      }
    }
  }
return "未分类";
}

function dedupeRecords(records: FaultRecord[]): FaultRecord[] {
  const seen = new Set<string>();
  const output: FaultRecord[] = [];
  for (const record of records) {
    const key = `${record.date}|${record.line}|${record.startTime}|${record.endTime}`;
    if (seen.has(key)) continue;
    seen.add(key);
    output.push(record);
  }
  return output;
}

export function buildTypeSummary(records: FaultRecord[], month?: string): TypeSummary[] {
  const totalDowntime = records.reduce((sum, record) => sum + record.downtime, 0);
  const days = month ? daysInMonth(month) : Math.max(1, [...new Set(records.map((record) => record.date))].length || 1);
  const grouped = new Map<string, { count: number; downtime: number }>();
  for (const record of records) {
    const current = grouped.get(record.machineType) ?? { count: 0, downtime: 0 };
    current.count += 1;
    current.downtime += record.downtime;
    grouped.set(record.machineType, current);
  }

  let cumulative = 0;
  return [...grouped.entries()]
    .map(([type, value]) => ({ type, ...value }))
    .sort((a, b) => b.downtime - a.downtime)
    .map((item) => {
      const share = totalDowntime ? item.downtime / totalDowntime : 0;
      cumulative += share;
      return {
        type: item.type,
        count: item.count,
        downtime: round1(item.downtime),
        share,
        cumulativeShare: cumulative,
        mttr: item.count ? round1(item.downtime / item.count) : 0,
        mtbf: item.count ? round1((days * 24 * 60) / item.count / 60) : 0
      };
    });
}

export function buildMonthSummary(records: FaultRecord[], months: string[], deviceCount: number = 1): MonthSummary[] {
  return months.map((month) => {
    const monthRecords = records.filter((record) => record.date.startsWith(month));
    const days = daysInMonth(month);
    const downtime = monthRecords.reduce((sum, record) => sum + record.downtime, 0);
    const count = monthRecords.length;
    return {
      month,
      days,
      count,
      downtime: round1(downtime),
      dailyDowntime: round1(downtime / days),
      faultRate: round1((downtime / (days * 24 * 60 * deviceCount)) * 100),
      mttr: count ? round1(downtime / count) : 0,
      mtbf: count ? round1((days * 24 * 60) / count / 60) : 0
    };
  });
}

export function getDeviceCount(config: AnalysisConfig): number {
  return config.lineCount * config.classificationRules.length;
}

function round1(value: number): number {
  return Math.round(value * 10) / 10;
}



export function buildDailySummary(records: FaultRecord[], month: string, deviceCount: number = 1): DailySummary[] {
  const days = daysInMonth(month);
  const result: DailySummary[] = [];
  for (let d = 1; d <= days; d++) {
    const day = month + "-" + String(d).padStart(2, "0");
    const dayRecords = records.filter((record) => record.date === day);
    const downtime = dayRecords.reduce((sum, record) => sum + record.downtime, 0);
    const count = dayRecords.length;
    result.push({
      day,
      count,
      downtime: round1(downtime),
      faultRate: round1((downtime / (24 * 60 * deviceCount)) * 100),
      mttr: count ? round1(downtime / count) : 0,
      mtbf: count ? round1(24 * 60 / count / 60) : 0
    });
  }
  return result;
}
