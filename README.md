# PocketSmith Tax Return Prep (Australia - Employee)

A web app that ingests PocketSmith CSV exports and prepares an accountant-ready Australian employee tax return submission draft.

## What it does

- Uploads PocketSmith transaction CSV exports in-browser.
- Defaults FY selection to the most recently ended Australian FY (Sydney time).
- Classifies transactions with merchant-name-first logic (income, deductions, non-deductible, transfers, review items).
- Generates:
  - On-screen schedules
  - Downloadable submission markdown pack
  - Downloadable classified transaction CSV

## Merchant intelligence (new)

The site can enrich **every merchant in the selected FY** by calling backend APIs that:

- Canonicalize noisy bank statement merchant strings into stable merchant keys (strip IDs, receipt noise, transfer artifacts).
- Group transaction variants under the same merchant entity before enrichment/classification.
- Search ABR/ABN records by merchant/business name.
- Attempt ABN match extraction for each merchant.
- Pull ABN details (entity type, main place of business where available).
- Run web search to infer business type/category.
- Feed enrichment back into transaction classification.
- Export merchant intelligence as CSV.

## Important assumptions

- Transaction merchant text is the primary classification signal.
- PocketSmith categories are only low-confidence fallback when merchant text is inconclusive.
- Merchant enrichment is probabilistic and should still be accountant-reviewed.

## Run

1. Install dependencies:

```bash
npm install
```

2. Start the local web server:

```bash
npm start
```

3. Open:

```text
http://localhost:3000
```

## Files

- `index.html`: UI and workflow
- `styles.css`: Styling
- `app.js`: CSV parsing, FY logic, analysis engine, frontend enrichment workflow
- `server.js`: backend enrichment endpoints (ABR + web search)
