# IB Comps Builder

A free web app for investment banking students. Type a public company, confirm it, review a suggested peer set, and download an Excel trading comps workbook built from SEC filings. Students spend their time evaluating the numbers instead of collecting them.

## What the workbook contains

- **Comps tab:** the target and each peer on one row with share price, equity value, enterprise value, LTM revenue, EBITDA, EBIT and net income, and EV/Revenue, EV/EBITDA, EV/EBIT and P/E. Peer mean, median, 25th and 75th percentile.
- **One tab per company:** last three fiscal years, prior and current year-to-date, and LTM (`FY + YTD current − YTD prior`) for the income statement and cash flow items. Also the latest balance sheet and the bridge from equity value to enterprise value.
- **Sources tab:** every hard-coded number with its XBRL tag, period, form, filing date and a link to the filing.

Blue font marks values from filings and black marks formulas. Yellow cells are inputs for the student, such as the share price when no price key is set.

## How it works

- Company lookup uses SEC's `company_tickers.json`.
- Peers are companies with the same SIC industry code that have a listed ticker, ranked by how close their revenue is to the target's (SEC `frames` API).
- Financials come from SEC's `companyfacts` XBRL API. Each line item tries a list of standard us-gaap tags in order, period by period (`lib/financials.ts`).
- The `.xlsx` is built with ExcelJS (`lib/workbook.ts`).

It runs with no AI and no paid services.

## Deploy for free on Vercel

1. Push this repo to GitHub.
2. Sign in at vercel.com with GitHub. Click **Add New → Project**, then import this repo.
3. Under **Environment Variables**, add:
   - `SEC_USER_AGENT`: your app name and a contact email, for example `IB Comps Builder you@school.edu`. SEC requires this.
   - `FMP_API_KEY` (optional): a free key from financialmodelingprep.com that fills share prices automatically.
4. Click **Deploy**. You get a free `*.vercel.app` link to share.

## Run locally

```bash
npm install
cp .env.example .env.local   # then fill in SEC_USER_AGENT
npm run dev
npm test
```

## Known limits

- Covers SEC filers only, so no private companies and most non-US companies are out.
- Forward (NTM) estimates aren't available from free sources.
- Some companies use custom tags, which leave a blank cell with a note. Check the filing for those.
- EBITDA is EBIT plus cash-flow D&A, without adjustments for one-off items.
