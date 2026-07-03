import ExcelJS from "exceljs";
import type { BadgeConsumption, ExportResult } from "./types";
import { normalizeBadgeId } from "./parseMall";

function cellText(value: ExcelJS.CellValue): string {
  if (value == null) return "";
  if (typeof value === "object") {
    const v = value as { text?: string; result?: unknown; richText?: { text: string }[] };
    if (v.richText) return v.richText.map((r) => r.text).join("");
    if (typeof v.text === "string") return v.text;
    if (v.result != null) return String(v.result);
    return "";
  }
  return String(value);
}

/** Parse a Hager start date. Values look like "02/06/2026 10:46:20" (DD/MM/YYYY). */
function parseDate(value: ExcelJS.CellValue): { year: number; month: number } | null {
  if (value instanceof Date) {
    return { year: value.getFullYear(), month: value.getMonth() + 1 };
  }
  const m = cellText(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  return { year: parseInt(m[3], 10), month: parseInt(m[2], 10) };
}

/**
 * Read the Hager export. Data is in the "overview" sheet with a header row:
 * badgeid, badgename, energychargedkwh, startdatetransaction, ... We locate
 * columns by header name (case-insensitive) so column order can drift safely.
 */
export async function parseExport(buffer: ArrayBuffer): Promise<ExportResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const ws =
    wb.worksheets.find((s) => s.name.toLowerCase() === "overview") ??
    wb.worksheets.find((s) => s.state !== "hidden") ??
    wb.worksheets[0];
  if (!ws) throw new Error("Exportfilen saknar kalkylblad.");

  // Map header name -> column number from the first row.
  const headerRow = ws.getRow(1);
  const col: Record<string, number> = {};
  headerRow.eachCell((cell, c) => {
    const key = cellText(cell.value).trim().toLowerCase();
    if (key) col[key] = c;
  });

  const badgeCol = col["badgeid"];
  const kwhCol = col["energychargedkwh"];
  const nameCol = col["badgename"];
  const startCol = col["startdatetransaction"];
  if (!badgeCol || !kwhCol) {
    throw new Error(
      "Kunde inte hitta kolumnerna 'badgeid' och 'energychargedkwh' i exporten.",
    );
  }

  const byBadge = new Map<string, BadgeConsumption>();
  const monthCounts = new Map<string, number>();

  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const badgeId = normalizeBadgeId(cellText(row.getCell(badgeCol).value));
    if (!badgeId) continue;

    const kwh = parseFloat(cellText(row.getCell(kwhCol).value).replace(",", "."));
    if (!Number.isFinite(kwh)) continue;

    const existing = byBadge.get(badgeId);
    if (existing) {
      existing.kwh += kwh;
      existing.sessions += 1;
    } else {
      byBadge.set(badgeId, {
        badgeId,
        kwh,
        sessions: 1,
        badgeName: nameCol ? cellText(row.getCell(nameCol).value).trim() : "",
      });
    }

    if (startCol) {
      const d = parseDate(row.getCell(startCol).value);
      if (d) {
        const key = `${d.year}-${d.month}`;
        monthCounts.set(key, (monthCounts.get(key) ?? 0) + 1);
      }
    }
  }

  // Billing period = the most frequent month among transaction start dates.
  let period: { year: number; month: number } | null = null;
  let best = 0;
  for (const [key, count] of monthCounts) {
    if (count > best) {
      best = count;
      const [y, m] = key.split("-").map(Number);
      period = { year: y, month: m };
    }
  }

  return { consumption: [...byBadge.values()], period };
}
