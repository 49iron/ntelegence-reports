/**
 * Store Report Builder
 * Audience: Store Manager, Ops Lead
 * Scope: single location — full KPIs, call intent, behavioral, all agent scorecards
 */

const B = require('../lib/docx-builder');

function build({ client, reportType, data, opts, manual }) {
  const d       = manual ?? data ?? {};
  const store   = client.stores.find(s => s.id === opts.storeId) ?? client.stores[0];
  const label   = store?.label ?? 'Store';
  const dateLine= opts.dateLine ?? `Data period: ${opts.dateFrom} to ${opts.dateTo}  |  ${label}`;

  const kpi = d.kpi ?? [
    { label: 'TOTAL INBOUND CALLS', value: d.totalCalls    ?? '—', sub: 'All inbound calls' },
    { label: 'ANSWER RATE',         value: d.answerRate    ?? '—', sub: d.answeredLine ?? '', color: 'green' },
    { label: 'SALES OPPORTUNITIES', value: d.salesOpps     ?? '—', sub: d.salesOppsSub ?? '' },
    { label: 'APPOINTMENTS SET',    value: d.apptsSet      ?? '—', sub: d.apptsSub ?? '' },
    { label: 'AVG CALL SCORE',      value: d.avgScore      ?? '—', sub: d.avgScoreSub ?? '', color: 'amber' },
    { label: 'CALLS ASSESSED',      value: d.callsAssessed ?? '—', sub: d.callsAssessedSub ?? '' },
  ];

  const intentHeaders = ['Reason for Call', label, 'Combined %'];
  const intentWidths  = [3600, 2280, 3480];
  const intentRows    = d.callIntent ?? [];

  const criteria   = d.criteria ?? [];
  const behavHeaders = ['Behavioral Criteria', `${label} (n=${d.scored ?? 0})`];
  const behavWidths  = [6360, 3000];
  const behavRows    = criteria.map(c => [c.label, d.behavioral?.[c.key] ?? '—']);

  const agents    = d.agents ?? [];
  const findings  = d.findings  ?? [];
  const nextSteps = d.nextSteps ?? { immediate: [], thirtyDay: [], ongoing: [] };

  const body = [
    B.kpiTable(kpi),
    B.spacer(40),
    B.colorKey(),

    B.sectionHeader('CALL INTENT — REASON FOR CALL'),
    B.spacer(80),
    ...(intentRows.length
      ? [B.dataTable(intentHeaders, intentRows, intentWidths)]
      : [B.para([B.txt('No call intent data for this period.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(140),

    B.pageBreak(),
    B.sectionHeader('CSR BEHAVIORAL SCORING — % COMPLIANCE'),
    B.spacer(60),
    B.para([
      B.txt('Scoring based on the ', { size: 16, color: B.C.gray }),
      B.txt(`RNR 'Thanks for Calling LIVE!'`, { size: 16, bold: true, color: B.C.navy }),
      B.txt(' 50-point rubric.  ', { size: 16, color: B.C.gray }),
      B.txt('Red',   { size: 16, bold: true, color: B.C.red }),   B.txt(' = 0–20%   ', { size: 16, color: B.C.gray }),
      B.txt('Amber', { size: 16, bold: true, color: B.C.amber }), B.txt(' = 21–49%   ', { size: 16, color: B.C.gray }),
      B.txt('Green', { size: 16, bold: true, color: B.C.green }), B.txt(' = 50%+',     { size: 16, color: B.C.gray }),
    ], { spacing: { after: 80 } }),
    ...(behavRows.length
      ? [B.dataTable(behavHeaders, behavRows, behavWidths, { colorPct: true })]
      : [B.para([B.txt('No behavioral scoring data for this period.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(60),
    ...(d.behavNote ? [B.para([B.txt(d.behavNote, { size: 15, italics: true, color: B.C.gray })], { spacing: { after: 40 } })] : []),

    B.pageBreak(),
    B.sectionHeader('CSR BEHAVIORAL SCORING — AGENT SCORECARDS'),
    B.spacer(100),
    ...B.agentCardsSection(agents, criteria.map(c => c.label), label.toUpperCase()),

    B.pageBreak(),
    B.sectionHeader('KEY FINDINGS & IMMEDIATE OPPORTUNITIES'),
    B.spacer(80),
    ...(findings.length ? [B.findingsTable(findings)] : [B.para([B.txt('No findings for this period.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(120),

    B.sectionHeader('RECOMMENDED NEXT STEPS'),
    B.spacer(80),
    B.nextStepsTable(nextSteps),
    B.spacer(80),
  ];

  return B.buildDoc({
    title:       `${client.name}  —  ${label}`,
    subtitle:    'Store Call Intelligence Report',
    dateLine,
    clientLabel: client.name,
    month:       opts.month ?? opts.dateTo.slice(0, 7),
  }, body);
}

module.exports = build;
