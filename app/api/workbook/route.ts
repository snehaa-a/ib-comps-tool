import { extract } from "@/lib/financials";
import { getPrice } from "@/lib/price";
import { getCompanyFacts, getTickers } from "@/lib/sec";
import { buildWorkbook, type CompanyInput } from "@/lib/workbook";

export const maxDuration = 60;
const MAX_PEERS = 10;

async function load(ticker: string, cik: number): Promise<CompanyInput> {
  const [facts, quote] = await Promise.all([getCompanyFacts(cik), getPrice(ticker)]);
  return { ticker, data: extract(facts), quote };
}

export async function POST(req: Request) {
  let body: { target?: string; peers?: string[] };
  try {
    body = await req.json();
  } catch {
    return Response.json({ error: "Invalid request" }, { status: 400 });
  }
  const target = body.target?.toUpperCase();
  const peers = [...new Set((body.peers ?? []).map((p) => p.toUpperCase()))].filter((p) => p !== target).slice(0, MAX_PEERS);
  if (!target) return Response.json({ error: "Missing target ticker" }, { status: 400 });

  try {
    const tickers = await getTickers();
    const cikOf = (t: string) => tickers.find((x) => x.ticker === t)?.cik;
    const missing = [target, ...peers].filter((t) => !cikOf(t));
    if (missing.length) return Response.json({ error: `Not found in SEC ticker list: ${missing.join(", ")}` }, { status: 400 });

    // Sequential on purpose: keeps us well under SEC's 10 requests per second.
    const loaded: CompanyInput[] = [];
    for (const t of [target, ...peers]) loaded.push(await load(t, cikOf(t)!));

    const xlsx = await buildWorkbook(loaded[0], loaded.slice(1));
    return new Response(new Uint8Array(xlsx), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${target}_comps.xlsx"`,
      },
    });
  } catch (e) {
    return Response.json({ error: (e as Error).message }, { status: 502 });
  }
}
