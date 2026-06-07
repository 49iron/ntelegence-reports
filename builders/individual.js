/**
 * Individual Agent Scorecard Builder
 * Audience: Agent + Direct Manager
 * Scope: single agent — personal KPIs, full criterion detail, vs store avg, coaching notes
 */

const { Table, TableRow, TableCell, AlignmentType, BorderStyle, WidthType, ShadingType } = require('docx');
const B = require('../lib/docx-builder');

function build({ client, reportType, data, opts, manual }) {
  const d         = manual ?? data ?? {};
  const agentName = opts.agentName ?? d.agentName ?? 'Agent';
  const store     = client.stores.find(s => s.id === opts.storeId);
  const storeLabel= store?.label ?? d.storeLabel ?? 'Store';
  const dateLine  = opts.dateLine ?? `Data period: ${opts.dateFrom} to ${opts.dateTo}  |  ${storeLabel}`;

  const criteria  = d.criteria ?? [];
  const agentScores= d.agentScores ?? {};   // { criterion_key: pct_string }
  const storeAvg  = d.storeAvg ?? {};       // { criterion_key: pct_string }

  // ── Personal KPI tiles ─────────────────────────────────────────────────
  const kpi = d.kpi ?? [
    { label: 'TOTAL CALLS',         value: d.totalCalls ?? '—', sub: 'Calls handled' },
    { label: 'CALLS SCORED',        value: d.scored     ?? '—', sub: 'Scored this period' },
    { label: 'AVG CALL SCORE',      value: d.avgScore   ?? '—', sub: `vs store avg ${d.storeAvgScore ?? '—'}`, color: 'amber' },
    { label: 'SALES OPPORTUNITIES', value: d.salesOpps  ?? '—', sub: 'Identified' },
    { label: 'APPTS CONFIRMED',     value: d.appts      ?? '—', sub: 'Booked' },
    { label: 'AVG TALK TIME',       value: d.talkTime   ?? '—', sub: 'Per answered call' },
  ];

  // ── Criterion detail table: agent % | store avg % | gap ────────────────
  const criteriaRows = criteria.map((c, i) => {
    const agPct   = agentScores[c.key] ?? '—';
    const stPct   = storeAvg[c.key]   ?? '—';
    const agNum   = parseFloat(agPct);
    const stNum   = parseFloat(stPct);
    const gap     = (!isNaN(agNum) && !isNaN(stNum)) ? `${agNum >= stNum ? '+' : ''}${Math.round(agNum - stNum)}%` : '—';
    const gapColor= (!isNaN(agNum) && !isNaN(stNum)) ? (agNum >= stNum ? B.C.green : B.C.red) : B.C.gray;
    const fill    = i % 2 === 0 ? B.C.white : B.C.offwhite;
    const thin    = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
    const borders = { top: thin, bottom: thin, left: thin, right: thin };

    return new TableRow({ children: [
      new TableCell({ borders, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: 3600, type: WidthType.DXA },
        children: [B.para([B.txt(c.label, { size: 17 })], { alignment: AlignmentType.LEFT })] }),
      new TableCell({ borders, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: 1920, type: WidthType.DXA },
        children: [B.para([B.txt(agPct, { size: 17, bold: true, color: B.pctColor ? B.C.black : B.C.black })], { alignment: AlignmentType.CENTER })] }),
      new TableCell({ borders, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: 1920, type: WidthType.DXA },
        children: [B.para([B.txt(stPct, { size: 17, color: B.C.gray })], { alignment: AlignmentType.CENTER })] }),
      new TableCell({ borders, shading: { fill, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, width: { size: 1920, type: WidthType.DXA },
        children: [B.para([B.txt(gap,   { size: 17, bold: true, color: gapColor })], { alignment: AlignmentType.CENTER })] }),
    ]});
  });

  const criterionTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3600, 1920, 1920, 1920],
    rows: [
      new TableRow({ tableHeader: true, children: [
        new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' } }, shading: { fill: B.C.navy, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, width: { size: 3600, type: WidthType.DXA },
          children: [B.para([B.txt('Behavioral Criteria', { bold: true, color: B.C.white, size: 17 })], { alignment: AlignmentType.CENTER })] }),
        new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' } }, shading: { fill: B.C.navy, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, width: { size: 1920, type: WidthType.DXA },
          children: [B.para([B.txt('My Score', { bold: true, color: B.C.white, size: 17 })], { alignment: AlignmentType.CENTER })] }),
        new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' } }, shading: { fill: B.C.navy, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, width: { size: 1920, type: WidthType.DXA },
          children: [B.para([B.txt('Store Avg', { bold: true, color: B.C.white, size: 17 })], { alignment: AlignmentType.CENTER })] }),
        new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }, right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' } }, shading: { fill: B.C.navy, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, width: { size: 1920, type: WidthType.DXA },
          children: [B.para([B.txt('vs. Store', { bold: true, color: B.C.white, size: 17 })], { alignment: AlignmentType.CENTER })] }),
      ]}),
      ...criteriaRows,
    ],
  });

  const coachingNotes = d.coachingNotes ?? [];

  const body = [
    B.kpiTable(kpi),
    B.spacer(40),
    B.colorKey(),

    B.sectionHeader(`BEHAVIORAL SCORECARD — ${agentName.toUpperCase()}`),
    B.spacer(80),
    criterionTable,
    B.spacer(60),
    B.para([B.txt(`Strongest behavior: ${d.strongest ?? '—'}  |  Biggest opportunity: ${d.weakest ?? '—'}`,
      { size: 17, bold: true, color: B.C.navy })], { spacing: { after: 40 } }),

    ...(coachingNotes.length ? [
      B.pageBreak(),
      B.sectionHeader('COACHING PRIORITIES & NEXT STEPS'),
      B.spacer(80),
      B.findingsTable(coachingNotes),
      B.spacer(80),
    ] : []),
  ];

  return B.buildDoc({
    title:       agentName,
    subtitle:    `Agent Scorecard  —  ${storeLabel}`,
    dateLine,
    clientLabel: client.name,
    month:       opts.month ?? opts.dateTo.slice(0, 7),
  }, body);
}

module.exports = build;
