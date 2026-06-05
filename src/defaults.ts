import type { AnalysisConfig } from "./types";

export const defaultConfig: AnalysisConfig = {
  sheets: ["L1", "L2", "L3", "L4"],
  minDowntime: 10,
  maxDowntime: 0,
  departmentFilter: "设备",
  lineCount: 4,
  includeKeywords: [],
  excludeKeywords: [
    {
      fields: ["machine"],
      keywords: ["注液", "包膜", "尺寸", "薄膜", "氦检", "烘干", "激光", "密封", "撕膜"],
      matchMode: "any"
    },
    {
      fields: ["description"],
      keywords: ["分容静置时间未到", "来料去预分选", "分容不出料", "分容未出料", "分容堵料"],
      matchMode: "any"
    }
  ],
  highlightKeywords: ["注液", "包膜", "尺寸", "薄膜", "氦检", "烘干", "激光", "密封", "待料", "工艺", "来料"],
  classificationRules: [
    { type: "WCS", field: "description", keywords: ["WCS", "wcs", "调度系统", "上位机", "通讯异常", "通讯故障", "通信异常", "通信故障", "网络通讯", "网络断开", "网络中断", "调度异常", "WMS", "wms"] },
    { type: "堆垛机", field: "description", keywords: ["堆垛机", "堆垛", "货叉", "载货台", "穿梭车", "RGV", "rgv", "堆垛车"] },
    { type: "物流线", field: "description", keywords: ["物流线", "物流", "提升机", "输送线", "输送", "滚筒", "倍速链", "倍速", "链条", "传送带", "传送", "皮带", "移载", "顶升", "转角", "转弯", "分拣", "拉带"] },
    { type: "机械手", field: "description", keywords: ["机械手", "机器人", "机械臂", "抓取", "夹爪", "夹具", "抓手", "robot", "Robot", "三坐标", "两坐标", "分选", "预分选", "压料"] },
    { type: "OCV", field: "description", keywords: ["OCV", "ocv", "开路电压", "开路"] },
    { type: "DCIR", field: "description", keywords: ["DCIR", "dcir", "直流内阻", "内阻"] },
    { type: "分容", field: "description", keywords: ["分容", "容量", "充放电", "充放", "放电", "充电", "针床", "探针", "压床", "库位", "托盘", "料盒"] },
    { type: "化成", field: "description", keywords: ["化成", "化成柜", "负压", "真空", "温控", "温度", "湿度", "注液", "溢液"] }
  ],
  columnMap: {
    date: "日期",
    line: "线体",
    startTime: "起始时间",
    endTime: "截止时间",
    downtime: "停机时长（min)",
    machine: "机器",
    description: "问题描述",
    department: "责任部门",
    owner: "责任人"
  }
};

export const availableKeywordFields = [
  { key: "description", label: "问题描述" },
  { key: "machine", label: "机器" },
  { key: "department", label: "责任部门" },
  { key: "line", label: "线体" },
  { key: "owner", label: "责任人" }
];

