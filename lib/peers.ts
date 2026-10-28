// Rule-based peer suggestions: same SIC industry code, listed on an exchange,
// closest in revenue. No AI needed; the student edits the list anyway.

import { getCompaniesBySic, getFrame, getSubmissions, getTickers, type TickerEntry } from "./sec";

const REVENUE_TAGS = ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"];

// SIC codes are narrow, so some industries have only a handful of US-GAAP filers
// (Nike's 3021 has Deckers and Crocs). When the target's own code runs short,
// these neighbouring codes fill the list after the same-code peers.
const RELATED_SIC: Record<string, string[]> = {
  "3021": ["3140", "2300"], // footwear ↔ apparel
  "3140": ["3021", "2300"],
  "2300": ["2320", "2330", "3021", "3140"],
  "2320": ["2300", "2330"],
  "2330": ["2300", "2320"],
  "2080": ["2082", "2086"], // beverages
  "2082": ["2080", "2086"],
  "2086": ["2080", "2082"],
  "2834": ["2836"], // pharma ↔ biologics
  "2836": ["2834"],
  "3674": ["3672", "3577"], // semiconductors
  "7370": ["7372", "7374"], // software and services
  "7372": ["7370", "7374"],
  "7374": ["7370", "7372"],
  "5311": ["5331", "5399"], // general merchandise retail
  "5331": ["5311", "5399"],
  "5399": ["5311", "5331"],
  "5812": ["5810"], // restaurants
  "1311": ["2911", "1381"], // oil and gas
  "2911": ["1311"],
  "3711": ["3714"], // autos
  "3714": ["3711"],
  "6021": ["6022"], // banks
  "6022": ["6021"],
};

async function revenueByCik(year: number): Promise<Map<number, number>> {
  const out = new Map<number, number>();
  for (const tag of REVENUE_TAGS) {
    try {
      for (const row of await getFrame(tag, year)) if (!out.has(row.cik)) out.set(row.cik, row.val);
    } catch {
      // A missing frame only weakens the size ranking.
    }
  }
  return out;
}

export type PeerSuggestion = TickerEntry & { revenue: number | null; reason: string };

export async function suggestPeers(cik: number, limit = 6): Promise<{ sicDescription: string; peers: PeerSuggestion[] }> {
  const target = await getSubmissions(cik);
  if (!target.sic) return { sicDescription: "", peers: [] };

  const tickers = await getTickers();
  const primary = new Map<number, TickerEntry>();
  for (const t of tickers) if (!primary.has(t.cik)) primary.set(t.cik, t);

  const year = new Date().getUTCFullYear() - 1;
  let revenue = await revenueByCik(year);
  if (!revenue.has(cik)) revenue = await revenueByCik(year - 1);
  const targetRev = revenue.get(cik);
  const distance = (rev: number) => (targetRev ? Math.abs(Math.log(rev / targetRev)) : -rev);

  const seen = new Set<number>([cik]);
  const peers: PeerSuggestion[] = [];
  for (const sic of [target.sic, ...(RELATED_SIC[target.sic] ?? [])]) {
    if (peers.length >= limit) break;
    // Companies with no US-GAAP revenue are shells or IFRS filers (such as On Holding),
    // whose tabs would come out empty, so they are left for the student to add by hand.
    const ranked = [...new Set(await getCompaniesBySic(sic))]
      .filter((c) => !seen.has(c) && primary.has(c) && revenue.has(c))
      .map((c) => ({ ...primary.get(c)!, revenue: revenue.get(c)! }))
      .sort((a, b) => distance(a.revenue) - distance(b.revenue));
    if (!ranked.length) continue;
    const description = sic === target.sic ? target.sicDescription : (await getSubmissions(ranked[0].cik)).sicDescription;
    const reason = sic === target.sic ? description.toLowerCase() || `SIC ${sic}` : `related industry: ${description.toLowerCase() || `SIC ${sic}`}`;
    for (const p of ranked.slice(0, limit - peers.length)) {
      seen.add(p.cik);
      peers.push({ ...p, reason });
    }
  }

  return { sicDescription: target.sicDescription, peers };
}
