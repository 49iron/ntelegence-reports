/**
 * Ntelegence Reports — HTTP Server
 *
 * Endpoints:
 *   GET  /health                  → 200 OK (Render health check)
 *   POST /generate                → generate + email a report on demand
 *   POST /generate/download       → generate + return as file download (no email)
 *
 * Called by TeleTraxx "Generate Report" button.
 * Protected by REPORTS_API_KEY env var.
 */

require('dotenv').config();
const express            = require('express');
const { Packer }         = require('docx');
const { generateReport, generateAndSend } = require('./lib/runner');

const app  = express();
const PORT = process.env.PORT || 4000;

app.use(express.json());

// ── Auth middleware ────────────────────────────────────────────────────────
app.use((req, res, next) => {
  if (req.path === '/health') return next();
  const key = req.headers['x-api-key'] ?? req.query.apiKey;
  if (!process.env.REPORTS_API_KEY || key === process.env.REPORTS_API_KEY) return next();
  res.status(401).json({ error: 'Unauthorized' });
});

// ── Health check ──────────────────────────────────────────────────────────
app.get('/health', (_, res) => res.json({ status: 'ok', ts: new Date().toISOString() }));

// ── Generate + email ──────────────────────────────────────────────────────
/**
 * POST /generate
 * Body: {
 *   clientId:   string,
 *   reportType: 'REGIONAL' | 'STORE' | 'INDIVIDUAL' | 'CORP',
 *   dateFrom:   'YYYY-MM-DD',
 *   dateTo:     'YYYY-MM-DD',
 *   storeId?:   string,
 *   agentName?: string,
 *   recipients?: string[],   // override default recipients
 *   manual?:    object,      // manually-supplied data (for clients without DB tenantId)
 * }
 */
app.post('/generate', async (req, res) => {
  try {
    const { clientId, reportType, dateFrom, dateTo, storeId, agentName, recipients, manual } = req.body;
    if (!clientId || !reportType || !dateFrom || !dateTo) {
      return res.status(400).json({ error: 'clientId, reportType, dateFrom, dateTo are required' });
    }

    await generateAndSend({ clientId, reportType, dateFrom, dateTo, storeId, agentName, recipients, manual });
    res.json({ success: true, message: `${reportType} report for ${clientId} sent.` });
  } catch (err) {
    console.error('/generate error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── Generate + download (no email) ────────────────────────────────────────
app.post('/generate/download', async (req, res) => {
  try {
    const { clientId, reportType, dateFrom, dateTo, storeId, agentName, manual } = req.body;
    if (!clientId || !reportType || !dateFrom || !dateTo) {
      return res.status(400).json({ error: 'clientId, reportType, dateFrom, dateTo are required' });
    }

    const buffer   = await generateReport({ clientId, reportType, dateFrom, dateTo, storeId, agentName, manual });
    const filename = `${clientId}_${reportType}_${dateTo}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('/generate/download error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── Start ─────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`Ntelegence Reports server listening on :${PORT}`);
});

module.exports = app;
