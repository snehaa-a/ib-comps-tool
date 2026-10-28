import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { extract } from "@/lib/financials";
import type { CompanyFacts, FactEntry } from "@/lib/sec";
import { buildWorkbook } from "@/lib/workbook";

// A small hand-made facts file shaped like SEC's companyfacts JSON, with a May fiscal year end.
const e = (start: string | undefined, end: string, val: number, form: string, filed: string, accn = `acc-${filed}`): FactEntry =>
  ({ start, end, val, form, filed, accn });

function fixture(): CompanyFacts {
  const usd = (list: FactEntry[]) => ({ units: { USD: list } });
  return {
    cik: 123,
    entityName: "Example Co",
    facts: {
      dei: {
        EntityCommonStockSharesOutstanding: {
          units: {
            shares: [
              e(undefined, "2025-06-20", 1_000_000_000, "10-K", "2025-07-20", "k25"),
              e(undefined, "2025-09-25", 1_200_000_000, "10-Q", "2025-10-05", "q1"),
              e(undefined, "2025-09-25", 300_000_000, "10-Q", "2025-10-05", "q1"),
            ],
          },
        },
      },
      "us-gaap": {
        // Tag switch: older years under SalesRevenueNet, newer under the ASC 606 tag.
        SalesRevenueNet: usd([e("2022-06-01", "2023-05-31", 900e6, "10-K", "2023-07-20")]),
        RevenueFromContractWithCustomerExcludingAssessedTax: usd([
          e("2023-06-01", "2024-05-31", 1000e6, "10-K", "2024-07-20"),
          e("2024-06-01", "2025-05-31", 1100e6, "10-K", "2025-07-20"),
          e("2024-06-01", "2024-08-31", 260e6, "10-Q", "2025-10-05"),
          e("2025-06-01", "2025-08-31", 300e6, "10-Q", "2025-10-05"),
        ]),
        NetIncomeLoss: usd([
          e("2022-06-01", "2023-05-31", 90e6, "10-K", "2023-07-20"),
          e("2023-06-01", "2024-05-31", 100e6, "10-K", "2024-07-20"),
          // Restated in a later 10-K: the later filing should win.
          e("2023-06-01", "2024-05-31", 95e6, "10-K", "2025-07-20"),
          e("2024-06-01", "2025-05-31", 120e6, "10-K", "2025-07-20"),
          e("2024-06-01", "2024-08-31", 30e6, "10-Q", "2025-10-05"),
          e("2025-06-01", "2025-08-31", 35e6, "10-Q", "2025-10-05"),
        ]),
        OperatingIncomeLoss: usd([
          e("2024-06-01", "2025-05-31", 150e6, "10-K", "2025-07-20"),
          e("2024-06-01", "2024-08-31", 40e6, "10-Q", "2025-10-05"),
          e("2025-06-01", "2025-08-31", 45e6, "10-Q", "2025-10-05"),
        ]),
        Assets: usd([e(undefined, "2025-05-31", 5000e6, "10-K", "2025-07-20"), e(undefined, "2025-08-31", 5100e6, "10-Q", "2025-10-05")]),
        CashAndCashEquivalentsAtCarryingValue: usd([e(undefined, "2025-08-31", 400e6, "10-Q", "2025-10-05")]),
        LongTermDebtNoncurrent: usd([e(undefined, "2025-08-31", 800e6, "10-Q", "2025-10-05")]),
      },
    },
  };
}

describe("extract", () => {
  const d = extract(fixture());

  it("finds the last three fiscal years and the open YTD period", () => {
    expect(d.fyEnds).toEqual(["2023-05-31", "2024-05-31", "2025-05-31"]);
    expect(d.ytd).toEqual({ start: "2025-06-01", end: "2025-08-31", priorStart: "2024-06-01", priorEnd: "2024-08-31" });
  });

  it("falls back across tags per period", () => {
    expect(d.duration.revenue.fy.map((s) => s?.tag)).toEqual([
      "SalesRevenueNet",
      "RevenueFromContractWithCustomerExcludingAssessedTax",
      "RevenueFromContractWithCustomerExcludingAssessedTax",
    ]);
    expect(d.duration.revenue.ytd?.val).toBe(300e6);
    expect(d.duration.revenue.ytdPrior?.val).toBe(260e6);
  });

  it("prefers the latest filed (restated) value", () => {
    expect(d.duration.netIncome.fy[1]?.val).toBe(95e6);
  });

  it("reads the balance sheet at the latest date and sums share classes", () => {
    expect(d.balanceDate).toBe("2025-08-31");
    expect(d.instant.cash?.val).toBe(400e6);
    expect(d.instant.shortTermDebt).toBeNull();
    expect(d.sharesOutstanding?.val).toBe(1_500_000_000);
  });
});

describe("buildWorkbook", () => {
  it("writes inputs, LTM formulas and comps links", async () => {
    const data = extract(fixture());
    const buf = await buildWorkbook({ ticker: "EXM", data, quote: { price: 20, source: "test" } }, [
      { ticker: "PEER", data, quote: null },
    ]);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    expect(wb.worksheets.map((w) => w.name)).toEqual(["Comps", "EXM", "PEER", "Sources"]);

    const ws = wb.getWorksheet("EXM")!;
    expect(ws.getCell("D7").value).toBe(1100); // FY revenue in $m
    expect(ws.getCell("G7").value).toMatchObject({ formula: "D7+F7-E7" });

    const comps = wb.getWorksheet("Comps")!;
    expect(comps.getCell("B5").value).toBe("EXM");
    expect((comps.getCell("J5").value as { formula: string }).formula).toMatch(/^'EXM'!B\d+$/);
    expect((comps.getCell("J9").value as { formula: string }).formula).toBe('IFERROR(AVERAGE(J7:J7),"NM")');

    const sources = wb.getWorksheet("Sources")!;
    expect(sources.rowCount).toBeGreaterThan(10);
  });
});
