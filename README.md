# Ntelegence Reports

Automated call intelligence report generation and delivery system.

Generates branded `.docx` executive summaries for all Ntelegence clients — on schedule or on demand — and delivers via Resend email. Completely independent of TeleTraxx/Replit.

## Report Types

| Type | Audience | Scope |
|------|----------|-------|
| `REGIONAL` | Marketing Director, Regional VP | All stores for one client, side-by-side |
| `STORE` | Store Manager, Ops Lead | Single location, full agent scorecards |
| `INDIVIDUAL` | Agent + Direct Manager | Single agent vs store average |
| `CORP` | PE owners, C-suite | Cross-client portfolio rollup |

## Architecture

```
lib/
  docx-builder.js   Shared formatting library (colors, tables, KPI tiles)
  db.js             Live data from Render PostgreSQL
  clients.js        Client registry — add clients here
  report-types.js   Four report tier definitions
  mailer.js         Resend email delivery
  runner.js         Orchestrates build + send

builders/
  regional.js       Multi-store rollup
  store.js          Single location deep dive
  individual.js     Agent scorecard with vs-store comparison
  corp.js           Portfolio level

server.js           Express: POST /generate, POST /generate/download
scheduler.js        node-cron: auto-runs per client/type schedule
cli.js              Local one-off generation
```

## Setup

### 1. Environment variables

```bash
cp .env.example .env
# Fill in DATABASE_URL, RESEND_API_KEY, REPORTS_API_KEY
```

| Variable | Where to get it |
|----------|----------------|
| `DATABASE_URL` | Render dashboard → PostgreSQL → Connect → External Database URL |
| `RESEND_API_KEY` | resend.com → API Keys |
| `REPORTS_API_KEY` | Any random string: `node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"` |

### 2. Install

```bash
npm install
```

### 3. Run locally

```bash
# Generate a report (saves to ./output/)
node cli.js --client rnr-tire-midwest --type REGIONAL --from 2026-05-31 --to 2026-06-07

# Generate + email
node cli.js --client rnr-tire-midwest --type REGIONAL --from 2026-05-31 --to 2026-06-07 --send

# Store report
node cli.js --client rnr-tire-midwest --type STORE --from 2026-05-31 --to 2026-06-07 --store stl

# Agent scorecard
node cli.js --client rnr-tire-midwest --type INDIVIDUAL --from 2026-05-31 --to 2026-06-07 --store stl --agent "Colin Johns"

# HTTP server (for TeleTraxx button)
node server.js

# Scheduler (cron auto-send)
node scheduler.js
```

## Deploy to Render

### Web Service (HTTP server — for TeleTraxx on-demand button)

| Setting | Value |
|---------|-------|
| Build command | `npm install` |
| Start command | `node server.js` |
| Environment | `DATABASE_URL`, `RESEND_API_KEY`, `REPORTS_API_KEY` |

### Background Worker (scheduled auto-send)

Same repo, same env vars, different start command:

| Setting | Value |
|---------|-------|
| Start command | `node scheduler.js` |
| Environment | + `START_MODE=scheduler` |

## TeleTraxx Integration

The "Generate Report" button in TeleTraxx calls:

```
POST https://your-render-url.onrender.com/generate
x-api-key: your_REPORTS_API_KEY

{
  "clientId":   "rnr-tire-midwest",
  "reportType": "REGIONAL",
  "dateFrom":   "2026-05-31",
  "dateTo":     "2026-06-07"
}
```

Returns `{ "success": true }` and emails the report to configured recipients.

For a download link instead of email:

```
POST /generate/download
→ Returns .docx file directly
```

## Adding a New Client

1. Open `lib/clients.js`
2. Copy the `_TEMPLATE` block at the bottom
3. Fill in `tenantId`, `stores`, `recipients`, `schedules`
4. Set `active: true`

That's it. The scheduler and runner pick it up automatically.

## Adding a New Store

In `lib/clients.js`, add to the client's `stores` array:

```js
{ id: 'newstore', label: 'New Store', location: 'New Store', trackingLabel: 'New Store Sales' }
```

---

Prepared by Ntelegence · jscherer@ntelegence.com · ntelegence.com
