// Builds the comps workbook: one tab per company, a Comps summary and a Sources tab.
// Banking conventions: blue font = hard-coded input from a filing, black = formula.

import ExcelJS from "exceljs";
import { DURATION_ITEMS, INSTANT_ITEMS, type Extracted, type Sourced } from "./financials";
import { filingIndexUrl } from "./sec";
import type { Quote } from "./price";

export type CompanyInput = { ticker: string; data: Extracted; quote: Quote | null };

const BLUE = { argb: "FF0000FF" };
const NUM = "#,##0.0;(#,##0.0);\"–\"";
const EPS = "0.00;(0.00)";
const MULT = "0.0\"x\";(0.0\"x\")";
const HEADER_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F3864" } };
const INPUT_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };

const scale = (s: Sourced, unit: string) => (unit === "USD/shares" ? s.val : s.val / 1e6);
const ref = (sheet: string, cell: string) => `'${sheet.replace(/'/g, "''")}'!${cell}`;
const sheetName = (ticker: string) => ticker.replace(/[\\/?*[\]:]/g, "-").slice(0, 31);

type SourceRow = { ticker: string; item: string; period: string; s: Sourced; cik: number };

// Rows on each company tab that the Comps tab points at.
type Cells = { name: string; price: string; equity: string; ev: string; rev: string; ebitda: string; ebit: string; ni: string; evRev: string; evEbitda: string; evEbit: string; pe: string };

function headerRow(row: ExcelJS.Row) {
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.eachCell((c) => (c.fill = HEADER_FILL));
}

function companySheet(wb: ExcelJS.Workbook, c: CompanyInput, sources: SourceRow[]): Cells {
  const name = sheetName(c.ticker);
  const ws = wb.addWorksheet(name);
  const d = c.data;
  ws.columns = [{ width: 38 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 }];

  ws.getCell("A1").value = `${d.name} (${c.ticker})`;
  ws.getCell("A1").font = { bold: true, size: 14 };
  ws.getCell("A2").value = "$ in millions, except per-share data. Shares in millions. Source: SEC EDGAR XBRL filings.";
  ws.getCell("A3").value = "Blue = value from a filing (see Sources tab). Black = formula. A blank cell means no matching tag was found: check the filing.";
  ws.getCell("A2").font = ws.getCell("A3").font = { italic: true, color: { argb: "FF595959" } };

  // Columns B..D = fiscal years (oldest first), E = prior YTD, F = current YTD, G = LTM.
  const fyCols = ["B", "C", "D"].slice(3 - d.fyEnds.length);
  const head = ws.getRow(5);
  head.values = ["Income statement and cash flow", "", "", "", "YTD prior", "YTD current", "LTM"];
  d.fyEnds.forEach((end, i) => (ws.getCell(`${fyCols[i]}5`).value = `FY ${end.slice(0, 4)}`));
  headerRow(head);
  const dates = ws.getRow(6);
  d.fyEnds.forEach((end, i) => (ws.getCell(`${fyCols[i]}6`).value = `ends ${end}`));
  if (d.ytd) {
    ws.getCell("E6").value = `${d.ytd.priorStart} to ${d.ytd.priorEnd}`;
    ws.getCell("F6").value = `${d.ytd.start} to ${d.ytd.end}`;
  }
  ws.getCell("G6").value = d.ytd ? `through ${d.ytd.end}` : "= latest FY";
  dates.font = { italic: true, size: 9, color: { argb: "FF595959" } };

  const order = ["revenue", "cogs", "grossProfit", "sga", "rnd", "ebit", "da", "ebitda", "interestExpense", "pretaxIncome", "taxes", "netIncome", "epsDiluted", "dilutedShares", "sbc", "cfo", "capex", "fcf"];
  const rowOf: Record<string, number> = {};
  order.forEach((k, i) => (rowOf[k] = 7 + i));
  const lastFy = fyCols.at(-1) ?? "D";
  const allCols = [...fyCols, "E", "F"];

  for (const key of order) {
    const r = rowOf[key];
    const def = DURATION_ITEMS.find((x) => x.key === key);
    const fmt = def?.unit === "USD/shares" ? EPS : NUM;
    ws.getCell(`A${r}`).value = def?.label ?? (key === "ebitda" ? "EBITDA (EBIT + D&A)" : "Free cash flow (CFO − capex)");

    if (!def || (key === "grossProfit" && !d.duration.grossProfit.fy.some(Boolean))) {
      // Derived rows: written as formulas in every column, including LTM.
      const f = (col: string) =>
        key === "ebitda" ? `${col}${rowOf.ebit}+${col}${rowOf.da}` :
        key === "fcf" ? `${col}${rowOf.cfo}-${col}${rowOf.capex}` :
        `${col}${rowOf.revenue}-${col}${rowOf.cogs}`;
      for (const col of [...allCols, "G"]) ws.getCell(`${col}${r}`).value = { formula: f(col) };
      if (key === "ebitda") ws.getRow(r).font = { bold: true };
    } else {
      const item = d.duration[key];
      const put = (col: string, s: Sourced | null, period: string) => {
        if (!s) return;
        const cell = ws.getCell(`${col}${r}`);
        cell.value = scale(s, def.unit);
        cell.font = { color: BLUE };
        sources.push({ ticker: c.ticker, item: def.label, period, s, cik: d.cik });
      };
      item.fy.forEach((s, i) => put(fyCols[i], s, `FY ending ${d.fyEnds[i]}`));
      put("E", item.ytdPrior, "YTD prior");
      put("F", item.ytd, "YTD current");

      const ltm = ws.getCell(`G${r}`);
      const lastFyVal = item.fy.at(-1);
      if (!d.ytd) {
        if (lastFyVal) ltm.value = { formula: `${lastFy}${r}` };
      } else if (key === "dilutedShares") {
        if (item.ytd) ltm.value = { formula: `F${r}` };
      } else if (lastFyVal && item.ytd && item.ytdPrior) {
        ltm.value = { formula: `${lastFy}${r}+F${r}-E${r}` };
      }
    }
    for (const col of [...allCols, "G"]) ws.getCell(`${col}${r}`).numFmt = fmt;
  }

  // Balance sheet, latest reported date.
  let r = rowOf.fcf + 2;
  const bsHead = ws.getRow(r);
  bsHead.values = [`Balance sheet${d.balanceDate ? ` as of ${d.balanceDate}` : ""}`, "$m"];
  headerRow(bsHead);
  const bs: Record<string, string> = {};
  for (const def of INSTANT_ITEMS) {
    r++;
    ws.getCell(`A${r}`).value = def.label;
    const s = d.instant[def.key];
    const cell = ws.getCell(`B${r}`);
    if (s) {
      cell.value = scale(s, def.unit);
      cell.font = { color: BLUE };
      sources.push({ ticker: c.ticker, item: def.label, period: `as of ${s.end}`, s, cik: d.cik });
    }
    cell.numFmt = NUM;
    bs[def.key] = `B${r}`;
  }

  // Market data and the bridge from equity value to enterprise value.
  r += 2;
  const mHead = ws.getRow(r);
  mHead.values = ["Valuation", ""];
  headerRow(mHead);
  const line = (label: string, value: ExcelJS.CellValue, fmt: string, bold = false) => {
    r++;
    ws.getCell(`A${r}`).value = label;
    const cell = ws.getCell(`B${r}`);
    cell.value = value;
    cell.numFmt = fmt;
    if (bold) ws.getRow(r).font = { bold: true };
    return `B${r}`;
  };

  const price = line(c.quote ? `Share price ($, ${c.quote.source})` : "Share price ($) — enter today's price", c.quote?.price ?? null, "0.00");
  ws.getCell(price).font = { color: BLUE };
  ws.getCell(price).fill = INPUT_FILL;
  const sh = d.sharesOutstanding;
  const shares = line(sh ? `Shares outstanding (m, cover page ${sh.end})` : "Shares outstanding (m) — enter from 10-Q cover", sh ? sh.val / 1e6 : null, NUM);
  ws.getCell(shares).font = { color: BLUE };
  if (sh) sources.push({ ticker: c.ticker, item: "Shares outstanding", period: `as of ${sh.end}`, s: sh, cik: d.cik });
  else ws.getCell(shares).fill = INPUT_FILL;

  const equity = line("Equity value", { formula: `${price}*${shares}` }, NUM, true);
  const debt = line("+ Total debt", { formula: `${bs.shortTermDebt}+${bs.currentLtd}+${bs.longTermDebt}` }, NUM);
  const pref = line("+ Preferred stock", { formula: bs.preferred }, NUM);
  const nci = line("+ Noncontrolling interest", { formula: bs.nci }, NUM);
  const cash = line("− Cash and short-term investments", { formula: `${bs.cash}+${bs.stInvestments}` }, NUM);
  const ev = line("Enterprise value", { formula: `${equity}+${debt}+${pref}+${nci}-${cash}` }, NUM, true);

  r++;
  const g = (k: string) => `G${rowOf[k]}`;
  const evRev = line("EV / LTM revenue", { formula: `IFERROR(${ev}/${g("revenue")},"NM")` }, MULT);
  const evEbitda = line("EV / LTM EBITDA", { formula: `IFERROR(IF(${g("ebitda")}<=0,"NM",${ev}/${g("ebitda")}),"NM")` }, MULT);
  const evEbit = line("EV / LTM EBIT", { formula: `IFERROR(IF(${g("ebit")}<=0,"NM",${ev}/${g("ebit")}),"NM")` }, MULT);
  const pe = line("P / E (LTM diluted EPS)", { formula: `IFERROR(IF(${g("epsDiluted")}<=0,"NM",${price}/${g("epsDiluted")}),"NM")` }, MULT);
  ws.views = [{ state: "frozen", ySplit: 5, xSplit: 1 }];

  const at = (cell: string) => ref(name, cell);
  return {
    name: d.name, price: at(price), equity: at(equity), ev: at(ev),
    rev: at(g("revenue")), ebitda: at(g("ebitda")), ebit: at(g("ebit")), ni: at(g("netIncome")),
    evRev: at(evRev), evEbitda: at(evEbitda), evEbit: at(evEbit), pe: at(pe),
  };
}

export async function buildWorkbook(target: CompanyInput, peers: CompanyInput[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "IB Comps Tool";
  const comps = wb.addWorksheet("Comps");
  const sources: SourceRow[] = [];
  const all = [target, ...peers];
  const cells = all.map((c) => companySheet(wb, c, sources));

  comps.getCell("A1").value = `Trading comparables: ${target.data.name}`;
  comps.getCell("A1").font = { bold: true, size: 14 };
  comps.getCell("A2").value = `$ in millions. Built ${new Date().toISOString().slice(0, 10)} from SEC EDGAR. Verify key figures against the filings before use.`;
  comps.getCell("A2").font = { italic: true, color: { argb: "FF595959" } };

  const cols = ["Company", "Ticker", "Share price", "Equity value", "Enterprise value", "LTM revenue", "LTM EBITDA", "LTM EBIT", "LTM net income", "EV / Revenue", "EV / EBITDA", "EV / EBIT", "P / E"];
  comps.columns = cols.map((_, i) => ({ width: i === 0 ? 34 : 14 }));
  const head = comps.getRow(4);
  head.values = cols;
  headerRow(head);

  const fmts = [undefined, undefined, "0.00", NUM, NUM, NUM, NUM, NUM, NUM, MULT, MULT, MULT, MULT];
  const writeRow = (rowNum: number, c: CompanyInput, x: Cells) => {
    const row = comps.getRow(rowNum);
    row.values = [x.name, c.ticker, ...[x.price, x.equity, x.ev, x.rev, x.ebitda, x.ebit, x.ni, x.evRev, x.evEbitda, x.evEbit, x.pe].map((f) => ({ formula: f }))];
    fmts.forEach((f, i) => f && (row.getCell(i + 1).numFmt = f));
  };

  writeRow(5, target, cells[0]);
  comps.getRow(5).font = { bold: true };
  comps.getCell("A6").value = "Peers";
  comps.getCell("A6").font = { italic: true };
  const first = 7;
  peers.forEach((p, i) => writeRow(first + i, p, cells[i + 1]));
  const last = first + peers.length - 1;

  if (peers.length) {
    const stats: [string, (range: string) => string][] = [
      ["Mean", (rg) => `AVERAGE(${rg})`],
      ["Median", (rg) => `MEDIAN(${rg})`],
      ["25th percentile", (rg) => `PERCENTILE(${rg},0.25)`],
      ["75th percentile", (rg) => `PERCENTILE(${rg},0.75)`],
    ];
    stats.forEach(([label, fn], i) => {
      const row = comps.getRow(last + 2 + i);
      row.getCell(1).value = `Peer ${label.toLowerCase()}`;
      row.font = { bold: true };
      ["J", "K", "L", "M"].forEach((col) => {
        const cell = comps.getCell(`${col}${last + 2 + i}`);
        cell.value = { formula: `IFERROR(${fn(`${col}${first}:${col}${last}`)},"NM")` };
        cell.numFmt = MULT;
      });
    });
  }
  comps.views = [{ state: "frozen", ySplit: 4, xSplit: 1 }];

  const src = wb.addWorksheet("Sources");
  src.columns = [
    { header: "Ticker", width: 10 }, { header: "Line item", width: 34 }, { header: "Period", width: 26 },
    { header: "Value as filed", width: 18 }, { header: "XBRL tag", width: 60 }, { header: "Form", width: 8 },
    { header: "Filed", width: 12 }, { header: "Filing", width: 14 },
  ];
  headerRow(src.getRow(1));
  for (const row of sources) {
    const added = src.addRow([row.ticker, row.item, row.period, row.s.val, row.s.tag, row.s.form, row.s.filed, ""]);
    added.getCell(4).numFmt = "#,##0.####";
    added.getCell(8).value = { text: "Open filing", hyperlink: filingIndexUrl(row.cik, row.s.accn) };
    added.getCell(8).font = { color: BLUE, underline: true };
  }
  src.views = [{ state: "frozen", ySplit: 1 }];

  return Buffer.from(await wb.xlsx.writeBuffer());
}
