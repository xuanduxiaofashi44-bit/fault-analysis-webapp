export type MatchMode = "any" | "all";

export type KeywordRule = {
  fields: string[];
  keywords: string[];
  matchMode: MatchMode;
};

export type ClassificationRule = {
  type: string;
  keywords: string[];
};

export type AnalysisConfig = {
  sheets: string[];
  minDowntime: number;
  maxDowntime: number;
  departmentFilter: string;
  lineCount: number;
  includeKeywords: KeywordRule[];
  excludeKeywords: KeywordRule[];
  highlightKeywords: string[];
  classificationRules: ClassificationRule[];
  columnMap: Record<string, string>;
};

export type FaultRecord = {
  id: string;
  sheet: string;
  date: string;
  line: string;
  startTime: string;
  endTime: string;
  downtime: number;
  machine: string;
  machineType: string;
  description: string;
  department: string;
  owner: string;
  highlighted: boolean;
};

export type TypeSummary = {
  type: string;
  count: number;
  downtime: number;
  share: number;
  cumulativeShare: number;
  mttr: number;
  mtbf: number;
};

export type MonthSummary = {
  month: string;
  days: number;
  count: number;
  downtime: number;
  dailyDowntime: number;
  faultRate: number;
  mttr: number;
  mtbf: number;
};

export type DailySummary = {
  day: string;
  count: number;
  downtime: number;
  faultRate: number;
  mttr: number;
  mtbf: number;
};

export type AnalysisResult = {
  records: FaultRecord[];
  warnings: string[];
  months: string[];
  typeSummary: TypeSummary[];
  typeSummaryByMonth: Record<string, TypeSummary[]>;
  monthSummary: MonthSummary[];
};
