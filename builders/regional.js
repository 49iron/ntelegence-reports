/**
 * Regional Report Builder
 * Audience: Marketing Director, Regional VP
 * Scope: all stores for one client, side-by-side comparison
 */

const {
  Document, Packer,
} = require('docx');

const B = require('../lib/docx-builder');

/**
 * @param {object} p
 * @param {object} p.client     — from clients.js
 * @param {object} p.reportType — from report-types.js
 * @param {object} p.data       — from db.getReportData(), or null (manual mode)
 * @param {object} p.opts       — original runner options
 * @param {object} p.manual     — optional manually-supplied data object (used when tenantId is null)
 */
function build({ client, reportType, data, opts, manual }) {
  // Support manual data injection (for clients not yet in DB)
  const d = manual ?? data ?? {};

  // ── Derive store labels from client config ─────────────────────────────
  const stores    = client.stores;
  const storeLabels = stores.map(s => s.label);

  // ── KPI tiles — pulled from data or manual override ───────────────────
  const kpi = d.kpi ?? [
    { label: 'TOTAL INBOUND CALLS', value: d.totalCalls   ?? '—', sub: 'All inbound calls' },
    { label: 'ANSWER RATE',         value: d.answerRate   ?? '—', sub: d.answeredLine ?? '', color: 'green' },
    { label: 'SALES OPPORTUNITIES', value: d.salesOpps    ?? '—', sub: d.salesOppsSub ?? '' },
    { label: 'APPOINTMENTS SET',    value: d.apptsSet     ?? '—', sub: d.apptsSub ?? '' },
    { label: 'AVG CALL SCORE',      value: d.avgScore     ?? '—', sub: d.avgScoreSub ?? '', color: 'amber' },
    { label: 'CALLS ASSESSED',      value: d.callsAssessed?? '—', sub: d.callsAssessedSub ?? '' },
  ];

  // ── Store breakdown table ──────────────────────────────────────────────
  const storeBreakdownHeaders = ['Metric', ...storeLabels, 'Combined'];
  const storeBreakdownWidths  = [3000, ...stores.map(() => Math.floor(5760 / stores.length)), 1600];
  const storeBreakdownRows    = d.storeBreakdown ?? [
    ['Total inbound calls',    ...stores.map(s => d[s.id]?.total    ?? '—'), d.totalCalls ?? '—'],
    ['Answer rate',            ...stores.map(s => d[s.id]?.answerRate?? '—'), d.answerRate ?? '—'],
    ['Missed calls',           ...stores.map(s => d[s.id]?.missed   ?? '—'), d.missed     ?? '—'],
    ['Sales opportunities',    ...stores.map(s => d[s.id]?.salesOpps?? '—'), d.salesOpps  ?? '—'],
    ['Appointments confirmed', ...stores.map(s => d[s.id]?.appts    ?? '—'), d.apptsSet   ?? '—'],
    ['Appointments tentative', ...stores.map(s => d[s.id]?.tentative?? '—'), d.tentative  ?? '—'],
    ['Avg talk time',          ...stores.map(s => d[s.id]?.talkTime ?? '—'), d.talkTime   ?? '—'],
    ['Calls scored',           ...stores.map(s => d[s.id]?.scored   ?? '—'), d.scored     ?? '—'],
    ['Avg call score',         ...stores.map(s => d[s.id]?.avgScore ?? '—'), d.avgScore   ?? '—'],
  ];

  // ── Call intent ────────────────────────────────────────────────────────
  const intentHeaders = ['Reason for Call', ...storeLabels, 'Combined %'];
  const intentWidths  = [3000, ...stores.map(() => Math.floor(4160 / stores.length)), 2200];
  const intentRows    = d.callIntent ?? [];

  // ── Behavioral scoring ─────────────────────────────────────────────────
  const criteria        = d.criteria ?? [];
  const behavHeaders    = ['Behavioral Criteria', ...storeLabels.map(l => `${l} (n=${d[l.toLowerCase().replace(/ /g,'_')]?.scored ?? 0})`), 'Combined'];
  const behavWidths     = [3200, ...stores.map(() => Math.floor(4560 / stores.length)), 1600];
  const behavRows       = criteria.map(c => [
    c.label,
    ...stores.map(s => d[s.id]?.behavioral?.[c.key] ?? '—'),
    d.behavioral?.[c.key] ?? '—',
  ]);

  // ── Findings & next steps ─────────────────────────────────────────────
  const findings  = d.findings  ?? [];
  const nextSteps = d.nextSteps ?? { immediate: [], thirtyDay: [], ongoing: [] };

  // ── Date label ────────────────────────────────────────────────────────
  const dateLine = opts.dateLine ?? `Data period: ${opts.dateFrom} to ${opts.dateTo}  |  ${storeLabels.join('  ·  ')}`;

  // ── Build document ────────────────────────────────────────────────────
  const body = [
    B.kpiTable(kpi),
    B.spacer(40),
    B.colorKey(),

    B.sectionHeader('STORE-BY-STORE BREAKDOWN'),
    B.spacer(80),
    B.dataTable(storeBreakdownHeaders, storeBreakdownRows, storeBreakdownWidths),
    B.spacer(140),

    B.sectionHeader('CALL INTENT — REASON FOR CALL'),
    B.spacer(80),
    ...(intentRows.length
      ? [B.dataTable(intentHeaders, intentRows, intentWidths)]
      : [B.para([B.txt('No call intent data available for this period.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(60),
    ...(d.intentNote
      ? [B.para([B.txt(d.intentNote, { size: 16, italics: true, color: B.C.gray })], { spacing: { after: 40 } })]
      : []),

    B.pageBreak(),
    B.sectionHeader('CSR BEHAVIORAL SCORING — % COMPLIANCE'),
    B.spacer(60),
    B.para([
      B.txt('Scoring based on the RNR \'Thanks for Calling LIVE!\' 50-point rubric.  ', { size: 16, color: B.C.gray }),
      B.txt('Red',   { size: 16, bold: true, color: B.C.red }),   B.txt(' = 0–20%   ', { size: 16, color: B.C.gray }),
      B.txt('Amber', { size: 16, bold: true, color: B.C.amber }), B.txt(' = 21–49%   ', { size: 16, color: B.C.gray }),
      B.txt('Green', { size: 16, bold: true, color: B.C.green }), B.txt(' = 50%+',     { size: 16, color: B.C.gray }),
    ], { spacing: { after: 80 } }),
    ...(behavRows.length
      ? [B.dataTable(behavHeaders, behavRows, behavWidths, { colorPct: true })]
      : [B.para([B.txt('No behavioral scoring data available for this period.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(140),

    B.pageBreak(),
    B.sectionHeader('KEY FINDINGS & IMMEDIATE OPPORTUNITIES'),
    B.spacer(80),
    ...(findings.length
      ? [B.findingsTable(findings)]
      : [B.para([B.txt('No findings recorded for this period.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(120),

    B.sectionHeader('RECOMMENDED NEXT STEPS'),
    B.spacer(80),
    B.nextStepsTable(nextSteps),
    B.spacer(80),
  ];

  return B.buildDoc({
    title:       client.name,
    subtitle:    `Call Intelligence Executive Summary  —  ${stores.length}-Store ${client.brand ?? 'Report'}`,
    dateLine,
    clientLabel: `${client.name}`,
    month:       opts.month ?? opts.dateTo.slice(0, 7),
  }, body);
}

module.exports = build;
