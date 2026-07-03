import { useMemo, useState } from "react";
import { parseMall } from "./parseMall";
import { parseExport } from "./parseExport";
import {
  SWEDISH_MONTHS,
  buildBillingData,
  buildWorkbook,
  reportFileName,
} from "./buildReport";
import type { BillingData, Tenant } from "./types";

type FileState = { name: string; buffer: ArrayBuffer } | null;

const money = (n: number) =>
  n.toLocaleString("sv-SE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const kwh = (n: number) =>
  n.toLocaleString("sv-SE", { minimumFractionDigits: 3, maximumFractionDigits: 3 });

export function App() {
  const [mallFile, setMallFile] = useState<FileState>(null);
  const [exportFile, setExportFile] = useState<FileState>(null);
  const [tenants, setTenants] = useState<Tenant[] | null>(null);
  const [data, setData] = useState<BillingData | null>(null);
  const [year, setYear] = useState<number>(new Date().getFullYear());
  const [month, setMonth] = useState<number>(new Date().getMonth() + 1);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function readFile(file: File): Promise<FileState> {
    return { name: file.name, buffer: await file.arrayBuffer() };
  }

  async function process(mall: FileState, exp: FileState) {
    if (!mall || !exp) return;
    setBusy(true);
    setError(null);
    try {
      const parsedTenants = await parseMall(mall.buffer);
      if (parsedTenants.length === 0) {
        throw new Error(
          "Hittade inga taggar i MALL-filen. Kontrollera att rätt fil laddades upp.",
        );
      }
      const { consumption, period } = await parseExport(exp.buffer);
      setTenants(parsedTenants);
      setData(buildBillingData(parsedTenants, consumption));
      if (period) {
        setYear(period.year);
        setMonth(period.month);
      }
    } catch (e) {
      setData(null);
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  async function onPick(kind: "mall" | "export", file?: File) {
    if (!file) return;
    const state = await readFile(file);
    const mall = kind === "mall" ? state : mallFile;
    const exp = kind === "export" ? state : exportFile;
    if (kind === "mall") setMallFile(state);
    else setExportFile(state);
    if (mall && exp) await process(mall, exp);
  }

  async function download() {
    if (!data) return;
    const blob = await buildWorkbook(data, year, month);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = reportFileName(year, month);
    a.click();
    URL.revokeObjectURL(url);
  }

  const totals = useMemo(() => {
    if (!data) return null;
    return data.rows.reduce(
      (acc, r) => ({
        kwh: acc.kwh + r.kwh,
        cost: acc.cost + r.cost,
        moms: acc.moms + r.moms,
        total: acc.total + r.total,
      }),
      { kwh: 0, cost: 0, moms: 0, total: 0 },
    );
  }, [data]);

  return (
    <div className="page">
      <header>
        <h1>Debitering Billaddare</h1>
        <p className="sub">Brf Lavetten i Varberg</p>
      </header>

      <section className="uploads">
        <FileInput
          label="1. MALL-fil (taggregister)"
          hint="0 MALL FÖR DEBITERING …"
          fileName={mallFile?.name}
          accept=".xlsx"
          onPick={(f) => onPick("mall", f)}
        />
        <FileInput
          label="2. Hager-export"
          hint="export … .xlsm / .xlsx"
          fileName={exportFile?.name}
          accept=".xlsm,.xlsx"
          onPick={(f) => onPick("export", f)}
        />
      </section>

      {busy && <p className="info">Bearbetar …</p>}
      {error && <p className="error">⚠ {error}</p>}

      {data && (
        <>
          <section className="controls">
            <label>
              Månad
              <select value={month} onChange={(e) => setMonth(Number(e.target.value))}>
                {SWEDISH_MONTHS.map((m, i) => (
                  <option key={m} value={i + 1}>{m}</option>
                ))}
              </select>
            </label>
            <label>
              År
              <input
                type="number"
                value={year}
                onChange={(e) => setYear(Number(e.target.value))}
              />
            </label>
            <button className="primary" onClick={download} disabled={data.rows.length === 0}>
              Ladda ner rapport
            </button>
          </section>

          {tenants && (
            <p className="info">
              {tenants.length} taggar i registret · {data.rows.length} laddade denna period
            </p>
          )}

          {data.unmatched.length > 0 && (
            <div className="warn">
              <strong>⚠ Taggar som laddat men saknas i MALL-filen</strong>
              <p className="warn-sub">
                Dessa har inte tagits med i rapporten. Kontrollera Badge Id i MALL-filen.
              </p>
              <ul>
                {data.unmatched.map((u) => (
                  <li key={u.consumption.badgeId}>
                    <code>{u.consumption.badgeId}</code> — {kwh(u.consumption.kwh)} kWh
                    {u.consumption.badgeName ? ` (${u.consumption.badgeName})` : ""}
                    {u.suggestion && (
                      <em>
                        {" "}→ liknar <code>{u.suggestion.badgeId}</code> ({u.suggestion.namn})
                        — möjlig felstavning i MALL-filen
                      </em>
                    )}
                  </li>
                ))}
              </ul>
            </div>
          )}

          {data.rows.length === 0 ? (
            <p className="info">Ingen laddning matchade registret för denna period.</p>
          ) : (
            <table className="report">
              <thead>
                <tr>
                  <th>Tag</th><th>Namn</th><th>Lägenhet</th><th>Badge Id</th>
                  <th className="num">kWh</th><th className="num">Pris</th>
                  <th className="num">Kostnad</th><th className="num">Moms</th>
                  <th className="num">Att betala</th>
                </tr>
              </thead>
              <tbody>
                {data.rows.map((r) => (
                  <tr key={r.tenant.badgeId}>
                    <td>{r.tenant.tagNr}</td>
                    <td>{r.tenant.namn}{r.tenant.external ? " (grannförening)" : ""}</td>
                    <td>{r.tenant.apartment}</td>
                    <td><code>{r.tenant.badgeId}</code></td>
                    <td className="num">{kwh(r.kwh)}</td>
                    <td className="num">{money(r.tenant.pricePerKwh)}</td>
                    <td className="num">{money(r.cost)}</td>
                    <td className="num">{money(r.moms)}</td>
                    <td className="num">{money(r.total)}</td>
                  </tr>
                ))}
              </tbody>
              {totals && (
                <tfoot>
                  <tr>
                    <td colSpan={4}>Summa</td>
                    <td className="num">{kwh(totals.kwh)}</td>
                    <td></td>
                    <td className="num">{money(totals.cost)}</td>
                    <td className="num">{money(totals.moms)}</td>
                    <td className="num">{money(totals.total)}</td>
                  </tr>
                </tfoot>
              )}
            </table>
          )}
        </>
      )}
    </div>
  );
}

function FileInput(props: {
  label: string;
  hint: string;
  fileName?: string;
  accept: string;
  onPick: (file?: File) => void;
}) {
  return (
    <label className={`drop ${props.fileName ? "has-file" : ""}`}>
      <span className="drop-label">{props.label}</span>
      <span className="drop-hint">{props.fileName ?? props.hint}</span>
      <input
        type="file"
        accept={props.accept}
        onChange={(e) => props.onPick(e.target.files?.[0])}
      />
    </label>
  );
}
