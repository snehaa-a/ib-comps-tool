import { NextResponse } from "next/server";
import { getTickers } from "@/lib/sec";

export async function GET(req: Request) {
  const q = (new URL(req.url).searchParams.get("q") ?? "").trim();
  if (!q) return NextResponse.json({ results: [] });
  try {
    const tickers = await getTickers();
    const upper = q.toUpperCase();
    const lower = q.toLowerCase();
    const seen = new Set<number>();
    const results = [
      ...tickers.filter((t) => t.ticker === upper),
      ...tickers.filter((t) => t.name.toLowerCase().startsWith(lower)),
      ...tickers.filter((t) => t.ticker.startsWith(upper)),
      // Match whole words in the name, so "NKE" doesn't pull in "Cullen/Frost Bankers".
      ...tickers.filter((t) => t.name.toLowerCase().split(/[^a-z0-9&]+/).some((w) => w.startsWith(lower))),
    ].filter((t) => !seen.has(t.cik) && seen.add(t.cik)).slice(0, 8);
    return NextResponse.json({ results });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 502 });
  }
}
