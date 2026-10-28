// Optional share price lookup. With no FMP_API_KEY the workbook leaves the
// price as a highlighted input cell for the student to fill.

export type Quote = { price: number; source: string };

export async function getPrice(ticker: string): Promise<Quote | null> {
  const key = process.env.FMP_API_KEY;
  if (!key) return null;
  try {
    const url = `https://financialmodelingprep.com/stable/quote?symbol=${encodeURIComponent(ticker)}&apikey=${key}`;
    const res = await fetch(url, { next: { revalidate: 900 } });
    if (!res.ok) return null;
    const [row] = (await res.json()) as { price?: number }[];
    return typeof row?.price === "number" ? { price: row.price, source: "Financial Modeling Prep" } : null;
  } catch {
    return null;
  }
}
