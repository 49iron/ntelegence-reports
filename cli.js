#!/usr/bin/env node
/**
 * Ntelegence Reports — CLI
 * Usage: node cli.js --client rnr-tire-midwest --type REGIONAL --from 2026-05-31 --to 2026-06-07 [--send]
 *
 * --send   emails the report to default recipients
 * omit     saves .docx locally to ./output/
 */

require('dotenv').config();
const path  = require('path');
const fs    = require('fs');
const { generateReport, generateAndSend } = require('./lib/runner');
const { getClient }  = require('./lib/clients');
const { getReportType } = require('./lib/report-types');

const args = process.argv.slice(2);
const get  = (flag) => { const i = args.indexOf(flag); return i !== -1 ? args[i + 1] : null; };

const clientId   = get('--client')  ?? get('-c');
const reportType = get('--type')    ?? get('-t');
const dateFrom   = get('--from')    ?? get('-f');
const dateTo     = get('--to')      ?? get('-d');
const storeId    = get('--store')   ?? null;
const agentName  = get('--agent')   ?? null;
const send       = args.includes('--send');

if (!clientId || !reportType || !dateFrom || !dateTo) {
  console.error(`
Usage:
  node cli.js --client <clientId> --type <REGIONAL|STORE|INDIVIDUAL|CORP> --from YYYY-MM-DD --to YYYY-MM-DD [--store <storeId>] [--agent <name>] [--send]

Examples:
  node cli.js --client rnr-tire-midwest --type REGIONAL  --from 2026-05-31 --to 2026-06-07
  node cli.js --client rnr-tire-midwest --type STORE     --from 2026-05-31 --to 2026-06-07 --store stl
  node cli.js --client rnr-tire-midwest --type INDIVIDUAL --from 2026-05-31 --to 2026-06-07 --store stl --agent "Colin Johns"
  node cli.js --client rnr-tire-midwest --type REGIONAL  --from 2026-05-31 --to 2026-06-07 --send
`);
  process.exit(1);
}

(async () => {
  try {
    const client     = getClient(clientId);
    const reportTypeObj = getReportType(reportType);
    const filename   = `${client.name.replace(/\s+/g,'_')}_${reportTypeObj.label.replace(/\s+/g,'_')}_${dateTo}.docx`;
    const outDir     = path.join(__dirname, 'output');

    if (send) {
      await generateAndSend({ clientId, reportType, dateFrom, dateTo, storeId, agentName });
    } else {
      const buffer = await generateReport({ clientId, reportType, dateFrom, dateTo, storeId, agentName });
      fs.mkdirSync(outDir, { recursive: true });
      const outPath = path.join(outDir, filename);
      fs.writeFileSync(outPath, buffer);
      console.log(`\n📄  Saved: ${outPath}`);
    }
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
})();
