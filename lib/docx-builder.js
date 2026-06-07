/**
 * Ntelegence Report Builder — Shared Library
 * Used by all client exec summary scripts.
 * Brand colors, helpers, and reusable section builders.
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak, Header, Footer, PageNumber, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');
const path = require('path');

// ── Brand ─────────────────────────────────────────────────────────────────
const C = {
  navy:     '1F3864',
  mid:      '2B6589',
  light:    'D9E8F0',
  green:    '1E7145',
  greenBg:  'E2EFDA',
  amber:    '7F6000',
  amberBg:  'FFF2CC',
  red:      '9C0006',
  redBg:    'FFE7E7',
  white:    'FFFFFF',
  offwhite: 'F2F7FA',
  gray:     '595959',
  black:    '000000',
};

// ── Border presets ────────────────────────────────────────────────────────
const noBorder   = { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' };
const noBorders  = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const thin       = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const thinBorders= { top: thin, bottom: thin, left: thin, right: thin };

// ── Core helpers ──────────────────────────────────────────────────────────
function txt(text, opts = {}) {
  return new TextRun({ text: String(text ?? ''), font: 'Arial', ...opts });
}

function para(children, opts = {}) {
  if (typeof children === 'string') children = [txt(children)];
  return new Paragraph({ children, ...opts });
}

function spacer(pts = 80) {
  return para('', { spacing: { before: pts, after: 0 } });
}

function pageBreak() {
  return para([new PageBreak()]);
}

// ── Section header (navy bar with white text) ─────────────────────────────
function sectionHeader(text) {
  return new Paragraph({
    children: [txt(text, { color: C.white, bold: true, size: 20 })],
    shading:  { fill: C.navy, type: ShadingType.CLEAR },
    spacing:  { before: 180, after: 80 },
    indent:   { left: 100, right: 100 },
    keepNext: true,
  });
}

// ── Sub-header (mid-blue rule + text) ────────────────────────────────────
function subHeader(text) {
  return new Paragraph({
    children: [txt(text, { bold: true, size: 20, color: C.mid })],
    spacing:  { after: 80 },
    border:   { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.mid } },
  });
}

// ── Color-code a percentage string ────────────────────────────────────────
function pctColor(cell) {
  const n = parseFloat(String(cell));
  if (isNaN(n) || !String(cell).includes('%')) return C.black;
  if (n >= 50) return C.green;
  if (n >= 21) return C.amber;
  return C.red;
}

// ── KPI tile table ────────────────────────────────────────────────────────
/**
 * tiles: array of { label, value, sub, color? }
 * color options: 'green', 'amber', 'red', 'light' (default)
 */
function kpiTable(tiles) {
  const bgMap = { green: C.greenBg, amber: C.amberBg, red: C.redBg, light: C.light };
  const colW  = Math.floor(9360 / tiles.length);

  const cells = tiles.map(t => new TableCell({
    borders:  thinBorders,
    shading:  { fill: bgMap[t.color] ?? C.light, type: ShadingType.CLEAR },
    margins:  { top: 100, bottom: 100, left: 100, right: 100 },
    width:    { size: colW, type: WidthType.DXA },
    children: [
      para([txt(t.value, { bold: true, size: 36, color: C.navy })], { alignment: AlignmentType.CENTER }),
      para([txt(t.label, { bold: true, size: 17, color: C.navy })],  { alignment: AlignmentType.CENTER }),
      para([txt(t.sub,   { size: 15,             color: C.gray })],  { alignment: AlignmentType.CENTER }),
    ],
  }));

  return new Table({
    width:        { size: 9360, type: WidthType.DXA },
    columnWidths: tiles.map(() => colW),
    rows:         [new TableRow({ children: cells })],
  });
}

// ── Color key line ────────────────────────────────────────────────────────
function colorKey() {
  return para([
    txt('Color key:  ', { size: 16, color: C.gray }),
    txt('Green = strong   ', { size: 16, bold: true, color: C.green }),
    txt('Amber = needs attention   ', { size: 16, bold: true, color: C.amber }),
    txt('Red = critical gap', { size: 16, bold: true, color: C.red }),
  ], { spacing: { after: 140 } });
}

// ── Generic data table ─────────────────────────────────────────────────────
/**
 * headers:   string[]
 * rows:      string[][]
 * colWidths: number[]  (DXA, must sum to table width)
 * opts.colorPct: true = color-code % cells
 */
function dataTable(headers, rows, colWidths, opts = {}) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders:  thinBorders,
      shading:  { fill: C.navy, type: ShadingType.CLEAR },
      margins:  { top: 80, bottom: 80, left: 100, right: 100 },
      width:    { size: colWidths[i], type: WidthType.DXA },
      children: [para([txt(h, { bold: true, color: C.white, size: 17 })], { alignment: AlignmentType.CENTER })],
    })),
  });

  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => {
      const isLabel = ci === 0;
      const fill    = ri % 2 === 0 ? C.white : C.offwhite;
      const color   = (!isLabel && opts.colorPct) ? pctColor(cell) : C.black;
      return new TableCell({
        borders:  thinBorders,
        shading:  { fill, type: ShadingType.CLEAR },
        margins:  { top: 60, bottom: 60, left: 100, right: 100 },
        width:    { size: colWidths[ci], type: WidthType.DXA },
        children: [para(
          [txt(String(cell ?? ''), { bold: isLabel, size: 17, color })],
          { alignment: isLabel ? AlignmentType.LEFT : AlignmentType.CENTER }
        )],
      });
    }),
  }));

  return new Table({
    width:        { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows:         [headerRow, ...dataRows],
  });
}

// ── Agent scorecard (single card, full width) ─────────────────────────────
/**
 * agent: { name, store, total, scored, avg, scores: [pct strings] }
 * criteria: string[] — criterion labels (same order as scores)
 */
function agentCard(agent, criteria) {
  const headerRow = new TableRow({ children: [
    new TableCell({
      columnSpan: 2,
      borders:    noBorders,
      shading:    { fill: C.mid, type: ShadingType.CLEAR },
      margins:    { top: 80, bottom: 80, left: 100, right: 100 },
      width:      { size: 4480, type: WidthType.DXA },
      children:   [para([
        txt(agent.name, { bold: true, size: 19, color: C.white }),
        txt(`   ${agent.store}  ·  ${agent.total} calls  ·  ${agent.scored} scored  ·  Avg `, { size: 16, color: C.light }),
        txt(`${agent.avg} / 50`, { bold: true, size: 16, color: C.white }),
      ])],
    }),
  ]});

  const criteriaRows = criteria.map((label, i) => {
    const pct  = agent.scores[i] ?? '—';
    const fill = i % 2 === 0 ? C.white : C.offwhite;
    const color= pctColor(pct);
    return new TableRow({ children: [
      new TableCell({
        borders:  thinBorders,
        shading:  { fill, type: ShadingType.CLEAR },
        margins:  { top: 40, bottom: 40, left: 80, right: 60 },
        width:    { size: 2880, type: WidthType.DXA },
        children: [para([txt(label, { size: 16 })], { alignment: AlignmentType.LEFT })],
      }),
      new TableCell({
        borders:  thinBorders,
        shading:  { fill, type: ShadingType.CLEAR },
        margins:  { top: 40, bottom: 40, left: 60, right: 80 },
        width:    { size: 1600, type: WidthType.DXA },
        children: [para([txt(pct, { size: 16, color, bold: parseFloat(pct) >= 50 })],
          { alignment: AlignmentType.RIGHT })],
      }),
    ]});
  });

  return new Table({
    width:        { size: 4480, type: WidthType.DXA },
    columnWidths: [2880, 1600],
    rows:         [headerRow, ...criteriaRows],
  });
}

// ── Agent scorecards 2-up layout ──────────────────────────────────────────
function agentCardsSection(agents, criteria, storeLabel) {
  const blocks = [subHeader(storeLabel), spacer(80)];

  if (!agents || agents.length === 0) {
    blocks.push(para([txt('No scored calls from verified agents for this period.',
      { size: 17, color: C.gray, italics: true })], { spacing: { after: 40 } }));
    return blocks;
  }

  for (let i = 0; i < agents.length; i += 2) {
    const left  = agentCard(agents[i], criteria);
    const right = agents[i + 1] ? agentCard(agents[i + 1], criteria) : null;

    blocks.push(new Table({
      width:        { size: 9360, type: WidthType.DXA },
      columnWidths: [4480, 400, 4480],
      rows: [new TableRow({ children: [
        new TableCell({ borders: noBorders, width: { size: 4480, type: WidthType.DXA }, children: [left] }),
        new TableCell({ borders: noBorders, width: { size: 400,  type: WidthType.DXA }, children: [para('')] }),
        new TableCell({ borders: noBorders, width: { size: 4480, type: WidthType.DXA }, children: right ? [right] : [para('')] }),
      ]})],
    }));
    blocks.push(spacer(120));
  }
  return blocks;
}

// ── Key findings table ────────────────────────────────────────────────────
/**
 * findings: [{ icon: '!' | '~' | '+', text }]
 */
function findingsTable(findings) {
  const iconColor = { '!': C.red, '~': C.amber, '+': C.green };
  const rowFills  = [C.redBg, C.white, C.amberBg, C.white, C.greenBg, C.white];

  const rows = findings.map((f, i) => new TableRow({ children: [
    new TableCell({
      borders:       thinBorders,
      shading:       { fill: rowFills[i % rowFills.length], type: ShadingType.CLEAR },
      margins:       { top: 80, bottom: 80, left: 120, right: 80 },
      width:         { size: 500,  type: WidthType.DXA },
      verticalAlign: VerticalAlign.TOP,
      children:      [para([txt(f.icon, { bold: true, size: 22, color: iconColor[f.icon] ?? C.gray })],
        { alignment: AlignmentType.CENTER })],
    }),
    new TableCell({
      borders:  thinBorders,
      shading:  { fill: rowFills[i % rowFills.length], type: ShadingType.CLEAR },
      margins:  { top: 80, bottom: 80, left: 100, right: 120 },
      width:    { size: 8860, type: WidthType.DXA },
      children: [para([txt(f.text, { size: 18 })], { alignment: AlignmentType.LEFT })],
    }),
  ]}));

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [500, 8860], rows });
}

// ── Next steps 3-column table ─────────────────────────────────────────────
/**
 * steps: { immediate: string[], thirtyDay: string[], ongoing: string[] }
 */
function nextStepsTable(steps) {
  const col = (heading, items, fill) => new TableCell({
    borders:  thinBorders,
    shading:  { fill, type: ShadingType.CLEAR },
    margins:  { top: 100, bottom: 100, left: 120, right: 120 },
    width:    { size: 3120, type: WidthType.DXA },
    children: [
      para([txt(heading, { bold: true, size: 18, color: C.navy })], { spacing: { after: 80 } }),
      ...items.map(t => para([txt('• ' + t, { size: 17 })], { spacing: { after: 60 } })),
    ],
  });

  return new Table({
    width:        { size: 9360, type: WidthType.DXA },
    columnWidths: [3120, 3120, 3120],
    rows: [new TableRow({ children: [
      col('Immediate (Days 1–14)',       steps.immediate,  C.redBg),
      col('30-Day Coaching Priorities',  steps.thirtyDay,  C.amberBg),
      col('Ongoing / Platform',          steps.ongoing,    C.greenBg),
    ]})],
  });
}

// ── Document wrapper ──────────────────────────────────────────────────────
/**
 * Build and write a complete report docx.
 * config: { title, subtitle, dateLine, clientLabel, month, outputPath }
 * children: Paragraph[] | Table[] — the full body content
 */
function buildReport(config, children) {
  const doc = new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
    sections: [{
      properties: {
        page: {
          size:   { width: 12240, height: 15840 },
          margin: { top: 720, right: 900, bottom: 720, left: 900 },
        },
      },
      headers: {
        default: new Header({ children: [new Paragraph({
          children: [
            txt('CONFIDENTIAL  |  Ntelegence Communication Intelligence', { size: 14, color: C.gray }),
            txt(`\t\t${config.clientLabel}  |  ${config.month}`,          { size: 14, color: C.gray }),
          ],
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          border:   { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.mid } },
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          children: [
            txt('Prepared by Ntelegence  |  jscherer@ntelegence.com  |  ntelegence.com', { size: 14, color: C.gray }),
            txt('\t\tPage ', { size: 14, color: C.gray }),
            new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.gray }),
          ],
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          border:   { top: { style: BorderStyle.SINGLE, size: 2, color: C.mid } },
        })] }),
      },
      children: [
        // Title block
        para([txt(config.title,    { bold: true, size: 44, color: C.navy })], { alignment: AlignmentType.CENTER, spacing: { before: 60, after: 40 } }),
        para([txt(config.subtitle, { size: 22, color: C.mid })],              { alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
        para([txt(config.dateLine, { size: 18, color: C.gray })],             { alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
        ...children,
      ],
    }],
  });

  return Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(config.outputPath, buf);
    console.log(`✓  ${path.basename(config.outputPath)}`);
    return config.outputPath;
  });
}

module.exports = {
  C, txt, para, spacer, pageBreak,
  sectionHeader, subHeader, colorKey,
  kpiTable, dataTable,
  agentCard, agentCardsSection,
  findingsTable, nextStepsTable,
  buildReport,
};

// ── buildDoc — returns Document object (no file write, no email) ──────────
// Used by builders that return the doc to runner.js for buffer generation.
function buildDoc(config, children) {
  return new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
    sections: [{
      properties: {
        page: {
          size:   { width: 12240, height: 15840 },
          margin: { top: 720, right: 900, bottom: 720, left: 900 },
        },
      },
      headers: {
        default: new Header({ children: [new Paragraph({
          children: [
            txt('CONFIDENTIAL  |  Ntelegence Communication Intelligence', { size: 14, color: C.gray }),
            txt(`\t\t${config.clientLabel}  |  ${config.month}`,          { size: 14, color: C.gray }),
          ],
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          border:   { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.mid } },
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          children: [
            txt('Prepared by Ntelegence  |  jscherer@ntelegence.com  |  ntelegence.com', { size: 14, color: C.gray }),
            txt('\t\tPage ', { size: 14, color: C.gray }),
            new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.gray }),
          ],
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          border:   { top: { style: BorderStyle.SINGLE, size: 2, color: C.mid } },
        })] }),
      },
      children: [
        para([txt(config.title,    { bold: true, size: 44, color: C.navy })], { alignment: AlignmentType.CENTER, spacing: { before: 60, after: 40 } }),
        para([txt(config.subtitle, { size: 22, color: C.mid })],              { alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
        para([txt(config.dateLine, { size: 18, color: C.gray })],             { alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
        ...children,
      ],
    }],
  });
}

module.exports = {
  C, txt, para, spacer, pageBreak, pctColor,
  sectionHeader, subHeader, colorKey,
  kpiTable, dataTable,
  agentCard, agentCardsSection,
  findingsTable, nextStepsTable,
  buildReport, buildDoc,
};
