/**
 * Ntelegence Reports — Report Runner
 *
 * generateReport(options) → Buffer (docx)
 * generateAndSend(options) → sends via Resend
 *
 * options: {
 *   clientId:   string              — key from clients.js
 *   reportType: 'REGIONAL' | 'STORE' | 'INDIVIDUAL' | 'CORP'
 *   dateFrom:   string              — 'YYYY-MM-DD'
 *   dateTo:     string              — 'YYYY-MM-DD'
 *   storeId:    string?             — required for STORE / INDIVIDUAL
 *   agentName:  string?             — required for INDIVIDUAL
 *   recipients: string[]?           — override default recipients
 *   send:       boolean?            — if true, email the report
 * }
 */

require('dotenv').config();
const { Packer }          = require('docx');
const { getClient }       = require('./clients');
const { getReportType }   = require('./report-types');
const db                  = require('./db');
const { sendReport }      = require('./mailer');
const buildRegional       = require('../builders/regional');
const buildStore          = require('../builders/store');
const buildIndividual     = require('../builders/individual');
const buildCorp           = require('../builders/corp');

const BUILDERS = {
  REGIONAL:   buildRegional,
  STORE:      buildStore,
  INDIVIDUAL: buildIndividual,
  CORP:       buildCorp,
};

async function generateReport(opts) {
  const client     = getClient(opts.clientId);
  const reportType = getReportType(opts.reportType);
  const builder    = BUILDERS[reportType.id.toUpperCase()];

  if (!builder) throw new Error(`No builder implemented for report type "${opts.reportType}"`);

  console.log(`▶  Generating ${reportType.label} — ${client.name} — ${opts.dateFrom} → ${opts.dateTo}`);

  // Pull live data from DB (skipped for clients with no tenantId — manual mode)
  let data = null;
  if (client.tenantId) {
    data = await db.getReportData(
      client.tenantId,
      opts.dateFrom,
      opts.dateTo,
      client.stores.map(s => s.location)
    );
  }

  // Build document
  const doc    = builder({ client, reportType, data, opts });
  const buffer = await Packer.toBuffer(doc);

  console.log(`✓  Built ${reportType.label} (${Math.round(buffer.length / 1024)} KB)`);
  return buffer;
}

async function generateAndSend(opts) {
  const client     = getClient(opts.clientId);
  const reportType = getReportType(opts.reportType);

  const buffer    = await generateReport(opts);
  const recipients= opts.recipients ?? client.recipients[opts.reportType] ?? ['jscherer@ntelegence.com'];
  const dateLabel = `${fmtDate(opts.dateFrom)} – ${fmtDate(opts.dateTo)}`;
  const filename  = `${client.name.replace(/\s+/g, '_')}_${reportType.label.replace(/\s+/g, '_')}_${opts.dateTo}.docx`;

  await sendReport({
    to:         recipients,
    subject:    `${client.name} — ${reportType.label} | ${dateLabel}`,
    clientName: client.name,
    period:     dateLabel,
    docxBuffer: buffer,
    filename,
  });

  return buffer;
}

function fmtDate(d) {
  // 'YYYY-MM-DD' → 'Jun 7, 2026'
  const dt = new Date(d + 'T12:00:00Z');
  return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' });
}

module.exports = { generateReport, generateAndSend };
