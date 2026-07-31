# Inventory Management System — AI Warehouse Assistant

Built on the same pattern as your Finance Management System: FastAPI backend,
an `.xlsx` file stored in a GitHub repo as the "database" (no SQL), a simple
shared-login frontend, and a chat bot powered by Gemini.

## What's different from the finance project

- **No fixed master template.** In the finance system, line items were
  locked in `AI_Financial_Report_Template.xlsx` forever. Here, products and
  locations are open-ended — the assistant creates a new one the first time
  you mention it, and reuses the exact existing name if you refer to
  something already known (it's told the current product/location list on
  every message, so "Dell Laptop" and "dell laptops" don't become two
  different products).
- **FIFO batches, not a single quantity per product.** Every delivery
  creates a dated batch. Every dispatch/transfer/write-off consumes the
  **oldest** batch first for that exact product + location. This is what
  makes "what's been sitting around for 3+ months" an exact query instead
  of a guess.
- **The bot answers questions too**, not just records transactions. Gemini
  only classifies *what* is being asked — the actual quantities always come
  straight from the ledger in Python, never from the model, so answers
  can't be hallucinated.

## Files

| File | Purpose |
|---|---|
| `main.py` | Entire backend (FastAPI). Deploy this. |
| `requirements.txt` | Python deps. |
| `index.html` | Login page. |
| `dashboard.html` | Summary cards, stock-by-location chart, aging alerts. |
| `assistant.html` | The chat UI — this is the AI Warehouse Assistant. |
| `reports.html` | Full stock table, aging report, transaction history. |
| `app.js` | Shared JS helpers used by all pages. |
| `style.css` | Shared styling. |

You'll also need a `chart.umd.js` file next to the HTML files (the Chart.js
UMD build) — same as your finance project already does.

## Environment variables (set in Vercel Project Settings, or a `.env` locally)

| Variable | Purpose |
|---|---|
| `GITHUB_TOKEN` | GitHub Personal Access Token, Contents: Read & Write |
| `GITHUB_REPO` | `your-username/your-repo` |
| `GITHUB_BRANCH` | usually `main` |
| `EXCEL_PATH` | where the live file lives, e.g. `data/inventory_data.xlsx` |
| `ACCESS_CODE` | shared login password |
| `GEMINI_API_KEY` | your Google Gemini API key |
| `AGING_THRESHOLD_DAYS` | optional, defaults to `90` (~3 months) |

Model is fixed to `gemini-3.1-flash-lite` in `main.py` — not configurable via env var.

No template `.xlsx` file needs to ship with this one — the workbook is
created automatically (with just the `Batches` and `Transaction Log`
sheets) the first time anything is saved.

## Try it via the chat bot

- "received 200 units of Dell Laptop in Mumbai Warehouse"
- "dispatched 50 Dell Laptop from Mumbai Warehouse"
- "moved 30 Dell Laptop from Mumbai Warehouse to Delhi Warehouse"
- "lost 5 Dell Laptop in Delhi Warehouse due to water damage"
- "how much Dell Laptop do we have and where is it kept?"
- "what's been sitting in stock for too long?"

## Things worth deciding before you go live

- **Ambiguous product names**: if two very differently-named batches of the
  same real item get created by mistake, there's no merge tool yet — you'd
  edit the downloaded `.xlsx` directly and re-upload, or ask me to add a
  `/api/merge-products` endpoint.
- **Units**: the model infers the unit ("units", "boxes", "kg") from your
  message; it isn't validated against a fixed list, so keep phrasing
  consistent for the same product if you want the totals to add up cleanly.
- **Undo** only reverses the single most recent *applied* transaction, not
  an arbitrary one further back.
