// Turns a company's raw XBRL facts into the standard line items a comps sheet needs.
// Every value keeps the tag and filing it came from, so the workbook can cite it.

import type { CompanyFacts, FactEntry } from "./sec";

export type Sourced = {
  val: number;
  tag: string;
  start?: string;
  end: string;
  form: string;
  filed: string;
  accn: string;
};

type ItemDef = { key: string; label: string; unit: "USD" | "USD/shares" | "shares"; tags: string[] };

// Candidate tags in order of preference. Companies switch tags over the years
// (SalesRevenueNet became RevenueFromContractWithCustomer... in 2018), so each
// period falls back through the list on its own.
export const DURATION_ITEMS: ItemDef[] = [
  { key: "revenue", label: "Revenue", unit: "USD", tags: ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet", "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueGoodsNet"] },
  { key: "cogs", label: "Cost of revenue", unit: "USD", tags: ["CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold", "CostOfServices"] },
  { key: "grossProfit", label: "Gross profit", unit: "USD", tags: ["GrossProfit"] },
  { key: "sga", label: "SG&A", unit: "USD", tags: ["SellingGeneralAndAdministrativeExpense"] },
  { key: "rnd", label: "R&D", unit: "USD", tags: ["ResearchAndDevelopmentExpense", "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"] },
  { key: "ebit", label: "Operating income (EBIT)", unit: "USD", tags: ["OperatingIncomeLoss"] },
  { key: "interestExpense", label: "Interest expense", unit: "USD", tags: ["InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt"] },
  { key: "pretaxIncome", label: "Pre-tax income", unit: "USD", tags: ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"] },
  { key: "taxes", label: "Income taxes", unit: "USD", tags: ["IncomeTaxExpenseBenefit"] },
  { key: "netIncome", label: "Net income", unit: "USD", tags: ["NetIncomeLoss", "ProfitLoss"] },
  { key: "epsDiluted", label: "Diluted EPS", unit: "USD/shares", tags: ["EarningsPerShareDiluted", "EarningsPerShareBasicAndDiluted"] },
  { key: "dilutedShares", label: "Diluted shares", unit: "shares", tags: ["WeightedAverageNumberOfDilutedSharesOutstanding"] },
  { key: "da", label: "D&A", unit: "USD", tags: ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization", "DepreciationAmortizationAndAccretionNet", "Depreciation"] },
  { key: "sbc", label: "Stock-based comp", unit: "USD", tags: ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"] },
  { key: "cfo", label: "Cash from operations", unit: "USD", tags: ["NetCashProvidedByUsedInOperatingActivities", "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"] },
  { key: "capex", label: "Capex", unit: "USD", tags: ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"] },
];

export const INSTANT_ITEMS: ItemDef[] = [
  { key: "cash", label: "Cash and equivalents", unit: "USD", tags: ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents", "Cash"] },
  { key: "stInvestments", label: "Short-term investments", unit: "USD", tags: ["ShortTermInvestments", "MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesDebtSecuritiesCurrent"] },
  { key: "shortTermDebt", label: "Short-term borrowings", unit: "USD", tags: ["ShortTermBorrowings", "CommercialPaper"] },
  { key: "currentLtd", label: "Current portion of long-term debt", unit: "USD", tags: ["LongTermDebtCurrent", "LongTermDebtAndCapitalLeaseObligationsCurrent"] },
  { key: "longTermDebt", label: "Long-term debt", unit: "USD", tags: ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations"] },
  { key: "preferred", label: "Preferred stock", unit: "USD", tags: ["PreferredStockValue"] },
  { key: "nci", label: "Noncontrolling interest", unit: "USD", tags: ["MinorityInterest"] },
  { key: "operatingLeases", label: "Operating lease liabilities (memo)", unit: "USD", tags: ["OperatingLeaseLiability"] },
  { key: "equity", label: "Shareholders' equity", unit: "USD", tags: ["StockholdersEquity"] },
];

const ANCHOR_TAGS = ["NetIncomeLoss", "ProfitLoss", "OperatingIncomeLoss", "RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues"];
const ANNUAL_FORMS = new Set(["10-K", "10-K/A", "10-KT"]);
const QUARTER_FORMS = new Set(["10-Q", "10-Q/A"]);

const days = (start: string, end: string) => (Date.parse(end) - Date.parse(start)) / 86_400_000;
const near = (a: string, b: string, tolerance = 7) => Math.abs(days(a, b)) <= tolerance;
const shiftYear = (d: string, years: number) => {
  const t = new Date(d + "T00:00:00Z");
  t.setUTCFullYear(t.getUTCFullYear() + years);
  return t.toISOString().slice(0, 10);
};

export type Extracted = {
  cik: number;
  name: string;
  fyEnds: string[]; // up to 3, oldest first
  ytd: { start: string; end: string; priorStart: string; priorEnd: string } | null;
  duration: Record<string, { fy: (Sourced | null)[]; ytd: Sourced | null; ytdPrior: Sourced | null }>;
  balanceDate: string | null;
  instant: Record<string, Sourced | null>;
  sharesOutstanding: Sourced | null;
};

function entries(facts: CompanyFacts, tag: string, unit: string): FactEntry[] {
  return facts.facts["us-gaap"]?.[tag]?.units?.[unit] ?? [];
}

// Latest-filed entry wins, so restated figures replace originals.
function pick(list: FactEntry[], tag: string, match: (e: FactEntry) => boolean): Sourced | null {
  let best: FactEntry | null = null;
  for (const e of list) if (match(e) && (!best || e.filed > best.filed)) best = e;
  return best && { val: best.val, tag, start: best.start, end: best.end, form: best.form, filed: best.filed, accn: best.accn };
}

function firstAcrossTags(facts: CompanyFacts, def: ItemDef, match: (e: FactEntry) => boolean): Sourced | null {
  for (const tag of def.tags) {
    const hit = pick(entries(facts, tag, def.unit), tag, match);
    if (hit) return hit;
  }
  return null;
}

function anchorEntries(facts: CompanyFacts): FactEntry[] {
  return ANCHOR_TAGS.flatMap((t) => entries(facts, t, "USD")).filter((e) => e.start);
}

export function fiscalYearEnds(facts: CompanyFacts, count = 3): string[] {
  const ends = new Set<string>();
  for (const e of anchorEntries(facts)) {
    const len = days(e.start!, e.end);
    if (ANNUAL_FORMS.has(e.form) && len >= 350 && len <= 380) ends.add(e.end);
  }
  return [...ends].sort().slice(-count);
}

export function ytdPeriod(facts: CompanyFacts, lastFyEnd: string | undefined): Extracted["ytd"] {
  const quarterly = anchorEntries(facts).filter((e) => QUARTER_FORMS.has(e.form));
  if (!quarterly.length) return null;
  const end = quarterly.reduce((m, e) => (e.end > m ? e.end : m), "");
  if (lastFyEnd && end <= lastFyEnd) return null;
  // Year-to-date = the longest period ending on the latest quarter that is under a year.
  let start = "";
  for (const e of quarterly) {
    if (e.end === end && days(e.start!, end) < 350 && (!start || e.start! < start)) start = e.start!;
  }
  if (!start) return null;
  return { start, end, priorStart: shiftYear(start, -1), priorEnd: shiftYear(end, -1) };
}

export function extract(facts: CompanyFacts): Extracted {
  const fyEnds = fiscalYearEnds(facts);
  const ytd = ytdPeriod(facts, fyEnds.at(-1));

  const duration: Extracted["duration"] = {};
  for (const def of DURATION_ITEMS) {
    const fy = fyEnds.map((end) =>
      firstAcrossTags(facts, def, (e) => !!e.start && e.end === end && days(e.start, e.end) >= 350 && days(e.start, e.end) <= 380),
    );
    const span = ytd ? days(ytd.start, ytd.end) : 0;
    const inPeriod = (s: string, en: string) => (e: FactEntry) =>
      !!e.start && near(e.end, en) && near(e.start, s) && Math.abs(days(e.start, e.end) - span) <= 7;
    duration[def.key] = {
      fy,
      ytd: ytd ? firstAcrossTags(facts, def, inPeriod(ytd.start, ytd.end)) : null,
      ytdPrior: ytd ? firstAcrossTags(facts, def, inPeriod(ytd.priorStart, ytd.priorEnd)) : null,
    };
  }

  const balanceDate =
    entries(facts, "Assets", "USD")
      .filter((e) => ANNUAL_FORMS.has(e.form) || QUARTER_FORMS.has(e.form))
      .reduce<string | null>((m, e) => (!m || e.end > m ? e.end : m), null);
  const instant: Extracted["instant"] = {};
  for (const def of INSTANT_ITEMS) {
    instant[def.key] = balanceDate ? firstAcrossTags(facts, def, (e) => !e.start && e.end === balanceDate) : null;
  }

  return {
    cik: facts.cik,
    name: facts.entityName,
    fyEnds,
    ytd,
    duration,
    balanceDate,
    instant,
    sharesOutstanding: sharesOutstanding(facts),
  };
}

// Cover-page share count from the latest filing, summed across share classes.
export function sharesOutstanding(facts: CompanyFacts): Sourced | null {
  const list = facts.facts.dei?.EntityCommonStockSharesOutstanding?.units?.shares ?? [];
  if (!list.length) return null;
  const latest = list.reduce((m, e) => (e.filed > m.filed ? e : m));
  const sameFiling = list.filter((e) => e.accn === latest.accn && e.end === latest.end);
  return {
    val: sameFiling.reduce((s, e) => s + e.val, 0),
    tag: "dei:EntityCommonStockSharesOutstanding",
    end: latest.end,
    form: latest.form,
    filed: latest.filed,
    accn: latest.accn,
  };
}
