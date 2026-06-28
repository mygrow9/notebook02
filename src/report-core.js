export const CANONICAL_HEADERS = [
  "單位",
  "總數(正常車及警示車)",
  "正常車",
  "警示車",
  "失車",
  "註(吊)銷",
  "典當車",
  "AB車",
  "權利車",
  "偽造車",
  "警示車_未分類",
];

const NUMERIC_HEADERS = CANONICAL_HEADERS.slice(1);
const KNOWN_PARENT_UNITS = [
  "高市警局保大",
  "高市警局交大",
  "高市警局婦幼隊",
  "高市警局少年隊",
  "高市警局捷運隊",
  "高市刑大",
];

export function normalizeHeader(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

export function parseNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = String(value).replace(/,/g, "").trim();
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function monthSheetName(month) {
  const numeric = parseInt(month, 10);
  if (!Number.isInteger(numeric) || numeric < 1 || numeric > 12) {
    throw new Error(`無法辨識月份：${month}`);
  }
  return String(numeric).padStart(2, "0");
}

export function inferParentUnit(unitName) {
  const unit = String(unitName ?? "").trim();
  const known = KNOWN_PARENT_UNITS.find((name) => unit.startsWith(name));
  if (known) return known;

  const divisionIndex = unit.indexOf("分局");
  if (divisionIndex >= 0) {
    return unit.slice(0, divisionIndex + "分局".length);
  }

  return unit;
}

function makeEmptyTotals() {
  return Object.fromEntries(NUMERIC_HEADERS.map((header) => [header, 0]));
}

function rowToCanonicalArray(unit, totals) {
  return [unit, ...NUMERIC_HEADERS.map((header) => totals[header] ?? 0)];
}

export function buildReport(monthlyReports) {
  const sortedReports = [...monthlyReports].sort((a, b) => a.month - b.month);
  const months = sortedReports.map((report) => report.month);
  const unitTotals = new Map();
  const parentMonthTotals = new Map();

  for (const report of sortedReports) {
    for (const row of report.rows) {
      const unit = String(row["單位"] ?? "").trim();
      if (!unit) continue;

      if (!unitTotals.has(unit)) unitTotals.set(unit, makeEmptyTotals());
      const unitBucket = unitTotals.get(unit);
      for (const header of NUMERIC_HEADERS) {
        unitBucket[header] += parseNumber(row[header]);
      }

      if (unit === "小計") continue;

      const parentUnit = inferParentUnit(unit);
      if (!parentMonthTotals.has(parentUnit)) {
        parentMonthTotals.set(parentUnit, Object.fromEntries(months.map((month) => [month, 0])));
      }
      parentMonthTotals.get(parentUnit)[report.month] += parseNumber(row["總數(正常車及警示車)"]);
    }
  }

  const totalRows = [
    CANONICAL_HEADERS,
    ...Array.from(unitTotals.entries()).map(([unit, totals]) => rowToCanonicalArray(unit, totals)),
  ];

  const parentData = Array.from(parentMonthTotals.entries())
    .map(([parentUnit, monthValues]) => {
      const monthlyValues = months.map((month) => monthValues[month] ?? 0);
      const total = monthlyValues.reduce((sum, value) => sum + value, 0);
      return { parentUnit, monthlyValues, total };
    })
    .sort((a, b) => b.total - a.total || a.parentUnit.localeCompare(b.parentUnit, "zh-Hant"));

  const parentRows = [
    ["母單位", ...months.map((month) => `${month}月`), "總計"],
    ...parentData.map((item) => [item.parentUnit, ...item.monthlyValues, item.total]),
  ];

  const grandMonthlyValues = months.map((_, index) =>
    parentData.reduce((sum, item) => sum + item.monthlyValues[index], 0),
  );
  const grandTotal = grandMonthlyValues.reduce((sum, value) => sum + value, 0);
  parentRows.push(["總計", ...grandMonthlyValues, grandTotal]);

  const chartRows = [
    ["母單位", "總計"],
    ...parentData.map((item) => [item.parentUnit, item.total]),
  ];

  return {
    months,
    totalRows,
    parentRows,
    chartRows,
    monthlyReports: sortedReports,
  };
}
