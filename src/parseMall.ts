import ExcelJS from "exceljs";
import type { Tenant } from "./types";

/** Normalize a badge id: trim, strip a leading "badge id:" label, lowercase. */
export function normalizeBadgeId(raw: string): string {
  return raw
    .replace(/badge\s*id\s*:?/i, "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

const BADGE_RE = /badge\s*id\s*:?\s*([0-9a-fA-F]{4,})/i;

function cellText(value: ExcelJS.CellValue): string {
  if (value == null) return "";
  if (typeof value === "object") {
    // Rich text / hyperlink / formula result objects.
    const v = value as { text?: string; result?: unknown; richText?: { text: string }[] };
    if (v.richText) return v.richText.map((r) => r.text).join("");
    if (typeof v.text === "string") return v.text;
    if (v.result != null) return String(v.result);
    return "";
  }
  return String(value);
}

/**
 * The MALL file lays each tenant out as a 5-row bordered box:
 *   row+0 (header): "Tag Nr." | "Namn/ Adress" | ...
 *   row+1: A=Tag Nr | B=Namn | C=phone
 *   row+2: C=email
 *   row+3: C="Badge Id: xxxx" | D="204044001-XXXX" | E="moms 25% :"
 *   row+4: E="Att Betala..."
 * The kWh input cell sits in column F of the box; its formula multiplier
 * (F..*3 or F..*3.8) encodes the price per kWh.
 *
 * We scan every row for a "Badge Id:" cell in column C, then read the
 * surrounding cells relative to it.
 */
export async function parseMall(buffer: ArrayBuffer): Promise<Tenant[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("MALL-filen saknar kalkylblad.");

  const tenants: Tenant[] = [];
  const seen = new Set<string>();

  ws.eachRow((row, rowNumber) => {
    const cCell = cellText(row.getCell("C").value);
    const match = cCell.match(BADGE_RE);
    if (!match) return;

    const badgeId = normalizeBadgeId(match[1]);
    if (!badgeId || seen.has(badgeId)) return;

    // Box first data row = 2 rows above the "Badge Id:" row.
    const firstRow = ws.getRow(rowNumber - 2);
    const tagNr = cellText(firstRow.getCell("A").value).trim();
    const namn = cellText(firstRow.getCell("B").value).trim();
    const apartment = cellText(row.getCell("D").value).trim();

    // Price from the box's F-column formula multiplier; default 3.00.
    const { price, external } = readPrice(ws, rowNumber);

    tenants.push({ tagNr, namn, apartment, badgeId, pricePerKwh: price, external });
    seen.add(badgeId);
  });

  return tenants;
}

/**
 * Look at the box's F cells (rows firstRow..lastRow) for a formula like
 * "F123*3" or "F123*3.8" and extract the multiplier. Also detect the external
 * "Grannförening" block (price 3.8 or an explicit "3,80" price label nearby).
 */
function readPrice(
  ws: ExcelJS.Worksheet,
  badgeRow: number,
): { price: number; external: boolean } {
  const first = badgeRow - 2;
  const last = badgeRow + 1;
  for (let r = first; r <= last; r++) {
    const cell = ws.getRow(r).getCell("F");
    const val = cell.value;
    const formula =
      val && typeof val === "object" && "formula" in val
        ? (val as { formula: string }).formula
        : "";
    const m = formula.match(/\*\s*([0-9]+(?:\.[0-9]+)?)/);
    if (m) {
      const price = parseFloat(m[1]);
      // Only treat cost-multiplier formulas (not the *0.25 moms line) as price.
      if (price >= 1) return { price, external: price >= 3.8 };
    }
  }
  return { price: 3.0, external: false };
}
