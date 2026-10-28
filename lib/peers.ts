// Rule-based peer suggestions: same SIC industry code, listed on an exchange,
// closest in revenue. No AI needed; the student edits the list anyway.

import { getCompaniesBySic, getFrame, getSubmissions, getTickers, type TickerEntry } from "./sec";

const REVENUE_TAGS = ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"];

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

  const [sameSic, tickers] = await Promise.all([getCompaniesBySic(target.sic), getTickers()]);
  const primary = new Map<number, TickerEntry>();
  for (const t of tickers) if (!primary.has(t.cik)) primary.set(t.cik, t);

  const candidates = [...new Set(sameSic)].filter((c) => c !== cik && primary.has(c)).map((c) => primary.get(c)!);

  const year = new Date().getUTCFullYear() - 1;
  let revenue = await revenueByCik(year);
  if (!revenue.has(cik)) revenue = await revenueByCik(year - 1);
  const targetRev = revenue.get(cik);

  const ranked = candidates
    .map((c) => ({ ...c, revenue: revenue.get(c.cik) ?? null }))
    .sort((a, b) => {
      if (a.revenue === null) return 1;
      if (b.revenue === null) return -1;
      if (!targetRev) return b.revenue - a.revenue;
      return Math.abs(Math.log(a.revenue / targetRev)) - Math.abs(Math.log(b.revenue / targetRev));
    });

  return {
    sicDescription: target.sicDescription,
    peers: ranked.slice(0, limit).map((p) => ({ ...p, reason: target.sicDescription.toLowerCase() || `SIC ${target.sic}` })),
  };
}
