import { readFile } from "node:fs/promises";
import { parseMall } from "./src/parseMall.ts";
import { parseExport } from "./src/parseExport.ts";
import { buildBillingData, buildWorkbook, reportFileName } from "./src/buildReport.ts";

const toAB = (buf) => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

const mallBuf = await readFile("0 MALL FÖR DEBITERING GÖR KOPIA..xlsx");
const expBuf = await readFile("example_export.xlsm");

const tenants = await parseMall(toAB(mallBuf));
console.log(`\n=== TENANTS (${tenants.length}) ===`);
for (const t of tenants) {
  console.log(
    `tag ${t.tagNr.padEnd(4)} ${t.namn.padEnd(24)} lgh ${t.apartment.padEnd(16)} badge ${t.badgeId.padEnd(10)} pris ${t.pricePerKwh}${t.external ? " [EXT]" : ""}`,
  );
}

const { consumption, period } = await parseExport(toAB(expBuf));
console.log(`\n=== EXPORT period=${JSON.stringify(period)} badges=${consumption.length} ===`);
for (const c of consumption.sort((a, b) => a.badgeId.localeCompare(b.badgeId))) {
  console.log(`badge ${c.badgeId.padEnd(10)} ${c.kwh.toFixed(3).padStart(9)} kWh  x${c.sessions}  ${c.badgeName}`);
}

const data = buildBillingData(tenants, consumption);
console.log(`\n=== REPORT ROWS (${data.rows.length}) ===`);
for (const r of data.rows) {
  console.log(
    `tag ${r.tenant.tagNr.padEnd(4)} ${r.tenant.namn.padEnd(24)} ${r.kwh.toFixed(3).padStart(9)} kWh  kostnad ${r.cost.toFixed(2).padStart(8)}  moms ${r.moms.toFixed(2).padStart(7)}  att betala ${r.total.toFixed(2).padStart(8)}`,
  );
}
console.log(`\n=== UNMATCHED (${data.unmatched.length}) ===`);
for (const u of data.unmatched)
  console.log(
    `badge ${u.consumption.badgeId} ${u.consumption.kwh.toFixed(3)} kWh (${u.consumption.badgeName})` +
      (u.suggestion ? `  → liknar ${u.suggestion.badgeId} (${u.suggestion.namn})` : ""),
  );

const blob = await buildWorkbook(data, period?.year ?? 2026, period?.month ?? 6);
console.log(`\n=== WORKBOOK ${reportFileName(period?.year ?? 2026, period?.month ?? 6)} = ${blob.size} bytes ===`);
