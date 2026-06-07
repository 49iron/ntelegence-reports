/**
 * Corporate Report Builder
 * Audience: PE owners, C-suite, Ntelegence internal
 * Scope: cross-client portfolio KPIs, top-line trends, exec findings only
 */

const { Table, TableRow, TableCell, AlignmentType, BorderStyle, WidthType, ShadingType } = require('docx');
const B = require('../lib/docx-builder');

function build({ client, reportType, data, opts, manual }) {
  const d       = manual ?? data ?? {};
  const dateLine= opts.dateLine ?? `Data period: ${opts.dateFrom} to ${opts.dateTo}`;

  // Portfolio table: one row per client/location group
  const portfolioRows  = d.portfolioRows  ?? [];
  const portfolioHeaders = ['Client / Location', 'Total Calls', 'Answer Rate', 'Sales Opps', 'Appts', 'Avg Score', 'Scored'];
  const portfolioWidths  = [2800, 1100, 1100, 1100, 800, 1100, 1360];

  const findings  = d.findings  ?? [];
  const nextSteps = d.nextSteps ?? { immediate: [], thirtyDay: [], ongoing: [] };

  // ── Portfolio KPI summary tiles ────────────────────────────────────────
  const kpi = d.kpi ?? [
    { label: 'TOTAL CALLS',         value: d.totalCalls  ?? '—', sub: 'Across all accounts' },
    { label: 'AVG ANSWER RATE',     value: d.answerRate  ?? '—', sub: 'All locations', color: 'green' },
    { label: 'TOTAL SALES OPPS',    value: d.salesOpps   ?? '—', sub: d.salesOppsSub ?? '' },
    { label: 'TOTAL APPTS',         value: d.appts       ?? '—', sub: 'Confirmed' },
    { label: 'AVG CALL SCORE',      value: d.avgScore    ?? '—', sub: 'Across scored calls', color: 'amber' },
    { label: 'ACTIVE LOCATIONS',    value: d.locations   ?? '—', sub: 'Reporting this period' },
  ];

  const body = [
    B.kpiTable(kpi),
    B.spacer(40),
    B.colorKey(),

    B.sectionHeader('PORTFOLIO BREAKDOWN — ALL ACCOUNTS'),
    B.spacer(80),
    ...(portfolioRows.length
      ? [B.dataTable(portfolioHeaders, portfolioRows, portfolioWidths)]
      : [B.para([B.txt('No portfolio data available for this period.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(140),

    B.pageBreak(),
    B.sectionHeader('KEY FINDINGS & PORTFOLIO INSIGHTS'),
    B.spacer(80),
    ...(findings.length ? [B.findingsTable(findings)] : [B.para([B.txt('No findings recorded.', { size: 17, italics: true, color: B.C.gray })])]),
    B.spacer(120),

    B.sectionHeader('RECOMMENDED NEXT STEPS'),
    B.spacer(80),
    B.nextStepsTable(nextSteps),
    B.spacer(80),
  ];

  return B.buildDoc({
    title:       opts.portfolioName ?? 'Ntelegence',
    subtitle:    'Communication Intelligence — Portfolio Summary',
    dateLine,
    clientLabel: 'Portfolio Report',
    month:       opts.month ?? opts.dateTo.slice(0, 7),
  }, body);
}

module.exports = build;
