// Thin client for the SEC EDGAR APIs. All calls run on the server.
// SEC asks for a descriptive User-Agent with a contact email and at most
// 10 requests per second: https://www.sec.gov/os/accessing-edgar-data

export type TickerEntry = { cik: number; ticker: string; name: string };

export type FactEntry = {
  start?: string;
  end: string;
  val: number;
  accn: string;
  fy?: number;
  fp?: string;
  form: string;
  filed: string;
  frame?: string;
};

export type CompanyFacts = {
  cik: number;
  entityName: string;
  facts: Record<string, Record<string, { label?: string; units: Record<string, FactEntry[]> }>>;
};

export type Submissions = {
  cik: string;
  name: string;
  tickers: string[];
  sic: string;
  sicDescription: string;
  fiscalYearEnd: string; // "MMDD"
};

const DAY = 24 * 60 * 60;

function userAgent(): string {
  // Tolerate a value pasted with quotes, which SEC would otherwise reject.
  const ua = (process.env.SEC_USER_AGENT ?? "").trim().replace(/^["']|["']$/g, "").trim();
  if (!ua) {
    throw new Error("SEC_USER_AGENT is not set. Add it in your Vercel project settings (see README).");
  }
  if (!/\S+@\S+\.\S+/.test(ua)) {
    throw new Error('SEC_USER_AGENT needs an app name and a contact email, like "IB Comps Builder you@school.edu".');
  }
  return ua;
}

// Spaces requests out so a burst of company pulls stays under SEC's limit.
let lastCall = 0;
async function throttle() {
  const wait = lastCall + 120 - Date.now();
  lastCall = Math.max(Date.now(), lastCall + 120);
  if (wait > 0) await new Promise((r) => setTimeout(r, wait));
}

async function secFetch(url: string, revalidate: number): Promise<Response> {
  await throttle();
  const res = await fetch(url, {
    headers: { "User-Agent": userAgent(), Accept: "application/json, application/atom+xml" },
    next: { revalidate },
  });
  if (res.status === 403) {
    throw new Error(
      "SEC refused the request (403). Check that SEC_USER_AGENT is an app name plus a real contact email with no quotes, then redeploy.",
    );
  }
  if (!res.ok) throw new Error(`SEC request failed (${res.status}): ${url}`);
  return res;
}

const memo = new Map<string, { at: number; value: unknown }>();
async function cached<T>(key: string, ttlMs: number, load: () => Promise<T>): Promise<T> {
  const hit = memo.get(key);
  if (hit && Date.now() - hit.at < ttlMs) return hit.value as T;
  const value = await load();
  memo.set(key, { at: Date.now(), value });
  if (memo.size > 200) memo.delete(memo.keys().next().value as string);
  return value;
}

export const padCik = (cik: number | string) => String(cik).padStart(10, "0");

export function getTickers(): Promise<TickerEntry[]> {
  return cached("tickers", 12 * 3600_000, async () => {
    const res = await secFetch("https://www.sec.gov/files/company_tickers.json", DAY);
    const raw = (await res.json()) as Record<string, { cik_str: number; ticker: string; title: string }>;
    return Object.values(raw).map((r) => ({ cik: r.cik_str, ticker: r.ticker.toUpperCase(), name: r.title }));
  });
}

export function getSubmissions(cik: number): Promise<Submissions> {
  return cached(`sub:${cik}`, 6 * 3600_000, async () => {
    const res = await secFetch(`https://data.sec.gov/submissions/CIK${padCik(cik)}.json`, DAY);
    const j = await res.json();
    return {
      cik: j.cik,
      name: j.name,
      tickers: j.tickers ?? [],
      sic: j.sic ?? "",
      sicDescription: j.sicDescription ?? "",
      fiscalYearEnd: j.fiscalYearEnd ?? "",
    };
  });
}

export function getCompanyFacts(cik: number): Promise<CompanyFacts> {
  return cached(`facts:${cik}`, 6 * 3600_000, async () => {
    // companyfacts files are often over Next's 2 MB data cache limit, so only memoize in memory.
    const res = await secFetch(`https://data.sec.gov/api/xbrl/companyfacts/CIK${padCik(cik)}.json`, 0);
    return (await res.json()) as CompanyFacts;
  });
}

// Companies filed under one SIC code, newest filers first. EDGAR serves this as an Atom feed.
export async function getCompaniesBySic(sic: string, pages = 3): Promise<number[]> {
  const ciks: number[] = [];
  for (let page = 0; page < pages; page++) {
    const url =
      `https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&SIC=${encodeURIComponent(sic)}` +
      `&owner=include&count=100&start=${page * 100}&output=atom`;
    const xml = await (await secFetch(url, 7 * DAY)).text();
    const found = [...xml.matchAll(/<cik>(\d+)<\/cik>/gi)].map((m) => Number(m[1]));
    ciks.push(...found);
    if (found.length < 100) break;
  }
  return ciks;
}

export const filingIndexUrl = (cik: number, accn: string) =>
  `https://www.sec.gov/Archives/edgar/data/${cik}/${accn.replace(/-/g, "")}/`;

export type FrameRow = { cik: number; entityName: string; val: number };

// One line item across every filer for a calendar year, used to size-rank peers.
export function getFrame(tag: string, year: number): Promise<FrameRow[]> {
  return cached(`frame:${tag}:${year}`, 24 * 3600_000, async () => {
    const res = await secFetch(`https://data.sec.gov/api/xbrl/frames/us-gaap/${tag}/USD/CY${year}.json`, 0);
    return ((await res.json()) as { data: FrameRow[] }).data;
  });
}
