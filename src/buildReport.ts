import ExcelJS from "exceljs";
import type { BadgeConsumption, BillingData, ReportRow, Tenant } from "./types";

const MOMS_RATE = 0.25;

/** Levenshtein distance, capped early — used to spot likely badge typos. */
function editDistance(a: string, b: string): number {
  if (Math.abs(a.length - b.length) > 2) return 99;
  const dp = Array.from({ length: a.length + 1 }, (_, i) => i);
  for (let j = 1; j <= b.length; j++) {
    let prev = dp[0];
    dp[0] = j;
    for (let i = 1; i <= a.length; i++) {
      const tmp = dp[i];
      dp[i] = Math.min(
        dp[i] + 1,
        dp[i - 1] + 1,
        prev + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
      prev = tmp;
    }
  }
  return dp[a.length];
}

export const SWEDISH_MONTHS = [
  "Januari", "Februari", "Mars", "April", "Maj", "Juni",
  "Juli", "Augusti", "September", "Oktober", "November", "December",
];

function round(n: number, decimals: number): number {
  const f = 10 ** decimals;
  return Math.round(n * f) / f;
}

/** Join consumption to tenants by badge id and compute costs. */
export function buildBillingData(
  tenants: Tenant[],
  consumption: BadgeConsumption[],
): BillingData {
  const byBadge = new Map(consumption.map((c) => [c.badgeId, c]));
  const tenantBadges = new Set(tenants.map((t) => t.badgeId));

  const rows: ReportRow[] = [];
  for (const tenant of tenants) {
    const c = byBadge.get(tenant.badgeId);
    if (!c || c.kwh <= 0) continue;
    const kwh = round(c.kwh, 3);
    const cost = round(kwh * tenant.pricePerKwh, 2);
    const moms = round(cost * MOMS_RATE, 2);
    rows.push({ tenant, kwh, cost, moms, total: round(cost + moms, 2) });
  }
  rows.sort((a, b) => {
    const na = parseFloat(a.tenant.tagNr);
    const nb = parseFloat(b.tenant.tagNr);
    if (Number.isFinite(na) && Number.isFinite(nb) && na !== nb) return na - nb;
    return a.tenant.tagNr.localeCompare(b.tenant.tagNr, "sv");
  });

  const unmatched = consumption
    .filter((c) => c.kwh > 0 && !tenantBadges.has(c.badgeId))
    .map((c) => {
      const near = tenants.find((t) => editDistance(t.badgeId, c.badgeId) <= 1);
      return {
        consumption: c,
        suggestion: near ? { badgeId: near.badgeId, namn: near.namn } : null,
      };
    });

  const charging = new Set(rows.map((r) => r.tenant.badgeId));
  const nonCharging = tenants.filter((t) => !charging.has(t.badgeId));

  return { rows, unmatched, nonCharging };
}

export function reportFileName(year: number, month: number): string {
  return `Debitering Billaddare ${SWEDISH_MONTHS[month - 1]} ${year}.xlsx`;
}

const MONEY_FMT = "#,##0.00";
const KWH_FMT = "#,##0.000";
const THIN = { style: "thin" as const, color: { argb: "FF999999" } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

/**
 * Build the simplified output workbook: one tenant per charging row, with a
 * title, header, and totals row.
 */
export async function buildWorkbook(
  data: BillingData,
  year: number,
  month: number,
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Debitering Billaddare";
  const ws = wb.addWorksheet("Debitering");

  const headers = [
    "Tag Nr", "Namn", "Lägenhet", "Badge Id", "Total kWh",
    "Pris/kWh", "Kostnad (ex moms)", "Moms 25%", "Att betala", "Not.",
  ];
  ws.columns = [
    { width: 8 }, { width: 24 }, { width: 18 }, { width: 12 }, { width: 11 },
    { width: 10 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 22 },
  ];

  // Title row.
  ws.mergeCells(1, 1, 1, headers.length);
  const title = ws.getCell(1, 1);
  title.value = `Debitering Billaddare – Brf Lavetten – ${SWEDISH_MONTHS[month - 1]} ${year}`;
  title.font = { bold: true, size: 14 };
  title.alignment = { horizontal: "left" };
  ws.getRow(1).height = 22;

  // Header row.
  const headerRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
    cell.border = BORDER;
    cell.alignment = { vertical: "middle", wrapText: true };
  });
  headerRow.height = 28;

  // Data rows.
  let r = 4;
  for (const row of data.rows) {
    const cells: (string | number)[] = [
      row.tenant.tagNr,
      row.tenant.namn,
      row.tenant.apartment,
      row.tenant.badgeId,
      row.kwh,
      row.tenant.pricePerKwh,
      row.cost,
      row.moms,
      row.total,
      row.tenant.external ? "Grannförening – faktureras" : "",
    ];
    const wsRow = ws.getRow(r);
    cells.forEach((v, i) => {
      const cell = wsRow.getCell(i + 1);
      cell.value = v;
      cell.border = BORDER;
    });
    wsRow.getCell(5).numFmt = KWH_FMT;
    wsRow.getCell(6).numFmt = MONEY_FMT;
    wsRow.getCell(7).numFmt = MONEY_FMT;
    wsRow.getCell(8).numFmt = MONEY_FMT;
    wsRow.getCell(9).numFmt = MONEY_FMT;
    r++;
  }

  // Totals row.
  const totalRow = ws.getRow(r);
  const sum = (pick: (row: ReportRow) => number) =>
    round(data.rows.reduce((acc, x) => acc + pick(x), 0), 2);
  totalRow.getCell(2).value = "Summa";
  totalRow.getCell(2).font = { bold: true };
  totalRow.getCell(5).value = round(data.rows.reduce((a, x) => a + x.kwh, 0), 3);
  totalRow.getCell(7).value = sum((x) => x.cost);
  totalRow.getCell(8).value = sum((x) => x.moms);
  totalRow.getCell(9).value = sum((x) => x.total);
  totalRow.getCell(5).numFmt = KWH_FMT;
  [7, 8, 9].forEach((c) => (totalRow.getCell(c).numFmt = MONEY_FMT));
  for (let c = 1; c <= headers.length; c++) {
    const cell = totalRow.getCell(c);
    cell.border = { ...BORDER, top: { style: "double", color: { argb: "FF666666" } } };
    if (c >= 5) cell.font = { bold: true };
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
