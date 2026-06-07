// ─────────────────────────────────────────────────────────────────────────────
// ADD TO server/routes.ts  (or server/routes/reports.ts if it exists)
// Adds POST /api/reports/generate for the TeleTraxx "Generate Report" button
// ─────────────────────────────────────────────────────────────────────────────

// At the top of routes.ts, add this import:
//   import { generateReportOnDemand } from './reportScheduler';

// Then add this route in the registerRoutes function:

app.post('/api/reports/generate', async (req, res) => {
  try {
    const { clientId, reportType, dateFrom, dateTo, recipients } = req.body;

    if (!clientId || !reportType || !dateFrom || !dateTo) {
      return res.status(400).json({ error: 'clientId, reportType, dateFrom, dateTo required' });
    }

    // Fire and forget — report sends in background, returns immediately
    generateReportOnDemand({ clientId, reportType, dateFrom, dateTo, recipients })
      .then(() => console.log(`[Reports] On-demand complete: ${clientId}/${reportType}`))
      .catch((err: Error) => console.error(`[Reports] On-demand failed: ${err.message}`));

    res.json({ success: true, message: `Generating ${reportType} for ${clientId} — will email when ready.` });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});
