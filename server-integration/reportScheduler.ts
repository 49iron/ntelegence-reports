/**
 * Ntelegence Report Scheduler
 * Drop into server/reportScheduler.ts
 * Started by server/index.ts with two lines:
 *   import { startReportScheduler } from './reportScheduler';
 *   startReportScheduler();
 *
 * Generates .docx executive summaries and sends via Resend on schedule.
 * Uses the same DATABASE_URL and RESEND_API_KEY already in TeleTraxx env.
 * No new infrastructure. No new cost.
 */

import cron from 'node-cron';
import { Resend } from 'resend';
import { pool } from './db';

// ── Types ─────────────────────────────────────────────────────────────────

interface ClientConfig {
  id: string;
  name: string;
  tenantId: string;
  stores: { id: string; label: string; location: string }[];
  recipients: { [reportType: string]: string[] };
  schedules: { [reportType: string]: string | null };
  dateWindowDays: number;
}

// ── Client Registry ────────────────────────────────────────────────────────

const CLIENTS: ClientConfig[] = [
  {
    id:       'rnr-tire-midwest',
    name:     'RNR Tire Midwest',
    tenantId: 'e36fb1c0-3af3-44d1-ba03-db9529854058',
    stores: [
      { id: 'stl',        label: 'St. Louis',  location: 'St. Louis'  },
      { id: 'evansville', label: 'Evansville', location: 'Evansville' },
    ],
    recipients: {
      REGIONAL: ['breanne.braddock@rnrtire.com', 'jscherer@ntelegence.com'],
    },
    schedules: {
      REGIONAL: '0 7 * * 1', // every Monday 7am CT
    },
    dateWindowDays: 7,
  },
  {
    id:       'rnr-tire-mo',
    name:     'RNR Tire MO',
    tenantId: '833fd1d4-c5b6-4764-bdd5-afc3be3f0308',
    stores: [
      { id: 'raytown',   label: 'Raytown',    location: 'Raytown'    },
      { id: 'gladstone', label: 'Gladstone',  location: 'Gladstone'  },
      { id: 'stjoe',     label: 'St. Joseph', location: 'St. Joseph' },
    ],
    recipients: {
      REGIONAL: ['sal.pernice@rnrtire.com', 'jscherer@ntelegence.com'],
    },
    schedules: {
      REGIONAL: '0 7 * * 1',
    },
    dateWindowDays: 7,
  },
];

// ── Date helpers ───────────────────────────────────────────────────────────

function rollingWindow(days: number): { dateFrom: string; dateTo: string } {
  const to   = new Date();
  const from = new Date();
  from.setDate(from.getDate() - days);
  return {
    dateFrom: from.toISOString().slice(0, 10),
    dateTo:   to.toISOString().slice(0, 10),
  };
}

function fmtDate(d: string): string {
  return new Date(d + 'T12:00:00Z').toLocaleDateString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC',
  });
}

// ── Live data queries ──────────────────────────────────────────────────────

async function fetchReportData(tenantId: string, dateFrom: string, dateTo: string, locations: string[]) {
  const client = await pool.connect();
  try {
    // Overall stats
    const statsRes = await client.query(`
      SELECT
        location,
        COUNT(*)                                                                          AS total,
        COUNT(*) FILTER (WHERE answered = true)                                           AS answered,
        COUNT(*) FILTER (WHERE answered = false)                                          AS missed,
        COUNT(*) FILTER (WHERE (custom_field_values->>'sales_opportunity') = 'YES')       AS sales_opps,
        COUNT(*) FILTER (WHERE (custom_field_values->>'appointment_booked') = 'YES')      AS appts_confirmed,
        COUNT(*) FILTER (WHERE (custom_field_values->>'appointment_booked') = 'TENTATIVE') AS appts_tentative,
        ROUND(AVG(duration_seconds)::numeric / 60, 1)                                     AS avg_talk_min,
        COUNT(*) FILTER (WHERE duration_seconds < 30 AND answered = true)                 AS short_calls,
        COUNT(*) FILTER (
          WHERE (custom_field_values->>'call_score') IS NOT NULL
            AND (custom_field_values->>'call_score') ~ '^[0-9]'
        )                                                                                 AS scored_calls,
        ROUND(AVG(
          CASE WHEN (custom_field_values->>'call_score') ~ '^[0-9]+(\.[0-9]+)?$'
               THEN (custom_field_values->>'call_score')::numeric END
        )::numeric, 1)                                                                    AS avg_score
      FROM calls
      WHERE tenant_id = $1 AND call_date >= $2 AND call_date <= $3
        AND ($4::text[] IS NULL OR location = ANY($4::text[]))
      GROUP BY location ORDER BY total DESC
    `, [tenantId, dateFrom, dateTo, locations.length ? locations : null]);

    // Call intent
    const intentRes = await client.query(`
      SELECT
        COALESCE(custom_field_values->>'call_reason', 'Unknown') AS reason,
        location,
        COUNT(*) AS total
      FROM calls
      WHERE tenant_id = $1 AND call_date >= $2 AND call_date <= $3
      GROUP BY reason, location ORDER BY COUNT(*) DESC
      LIMIT 20
    `, [tenantId, dateFrom, dateTo]);

    // Behavioral scores
    const behavRes = await client.query(`
      SELECT
        ar.criterion_name,
        c.location,
        COUNT(*)                                           AS total_scored,
        COUNT(*) FILTER (WHERE ar.passed = true)           AS passed,
        ROUND(100.0 * COUNT(*) FILTER (WHERE ar.passed = true) / NULLIF(COUNT(*), 0)) AS pct
      FROM assessment_results ar
      JOIN assessments a ON ar.assessment_id = a.id
      JOIN calls       c ON a.call_id        = c.id
      WHERE c.tenant_id = $1 AND c.call_date >= $2 AND c.call_date <= $3
      GROUP BY ar.criterion_name, c.location
      ORDER BY ar.criterion_name, c.location
    `, [tenantId, dateFrom, dateTo]);

    return {
      stats:  statsRes.rows,
      intent: intentRes.rows,
      behav:  behavRes.rows,
    };
  } finally {
    client.release();
  }
}

// ── Report builder (inline — no external file dependency) ──────────────────

async function buildRegionalReport(
  client: ClientConfig,
  data: Awaited<ReturnType<typeof fetchReportData>>,
  dateFrom: string,
  dateTo: string,
): Promise<Buffer> {
  // Dynamic import docx — installed in TeleTraxx already (used by PDF generation)
  const docx = await import('docx');
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
    PageBreak, Header, Footer, TabStopType, TabStopPosition, PageNumber,
  } = docx;

  const C = {
    navy: '1F3864', mid: '2B6589', light: 'D9E8F0',
    green: '1E7145', greenBg: 'E2EFDA',
    amber: '7F6000', amberBg: 'FFF2CC',
    red: '9C0006',   redBg: 'FFE7E7',
    white: 'FFFFFF', offwhite: 'F2F7FA', gray: '595959', black: '000000',
  };

  const thin    = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const borders = { top: thin, bottom: thin, left: thin, right: thin };
  const none    = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const noBorders = { top: none, bottom: none, left: none, right: none };

  const t = (text: string, opts: any = {}) =>
    new TextRun({ text: String(text ?? ''), font: 'Arial', ...opts });

  const p = (children: any, opts: any = {}) =>
    new Paragraph({ children: Array.isArray(children) ? children : [t(children)], ...opts });

  const spacer = (pts = 80) => p('', { spacing: { before: pts, after: 0 } });

  const sectionHdr = (text: string) => new Paragraph({
    children: [t(text, { color: C.white, bold: true, size: 20 })],
    shading: { fill: C.navy, type: ShadingType.CLEAR },
    spacing: { before: 180, after: 80 },
    indent: { left: 100, right: 100 },
  });

  const pctColor = (s: string) => {
    const n = parseFloat(s);
    if (isNaN(n) || !String(s).includes('%')) return C.black;
    return n >= 50 ? C.green : n >= 21 ? C.amber : C.red;
  };

  // Aggregate data
  const totals = data.stats.reduce((acc, r) => ({
    total:   (acc.total || 0)   + Number(r.total),
    answered:(acc.answered || 0)+ Number(r.answered),
    missed:  (acc.missed || 0)  + Number(r.missed),
    opps:    (acc.opps || 0)    + Number(r.sales_opps),
    appts:   (acc.appts || 0)   + Number(r.appts_confirmed),
    tent:    (acc.tent || 0)    + Number(r.appts_tentative),
    scored:  (acc.scored || 0)  + Number(r.scored_calls),
  }), {} as any);

  const answerRate = totals.total ? `${Math.round(100 * totals.answered / totals.total)}%` : '—';
  const avgScore   = data.stats.find(r => r.avg_score)?.avg_score ?? '—';

  const storeLabels = client.stores.map(s => s.label);
  const colW        = Math.floor(9360 / (storeLabels.length + 2));

  // KPI tiles
  const kpiTiles = [
    { label: 'TOTAL INBOUND CALLS',  value: String(totals.total || '—'), sub: 'All inbound calls' },
    { label: 'ANSWER RATE',          value: answerRate,                   sub: `${totals.answered} answered · ${totals.missed} missed`,   color: C.greenBg },
    { label: 'SALES OPPORTUNITIES',  value: String(totals.opps  || '—'), sub: `${totals.total ? Math.round(100*totals.opps/totals.total) : 0}% of all calls` },
    { label: 'APPOINTMENTS SET',     value: String(totals.appts || '—'), sub: `confirmed · +${totals.tent || 0} tentative` },
    { label: 'AVG CALL SCORE',       value: String(avgScore),            sub: `Based on ${totals.scored || 0} scored calls`, color: C.amberBg },
    { label: 'CALLS ASSESSED',       value: `${totals.scored || 0} (${totals.total ? Math.round(100*(totals.scored||0)/totals.total) : 0}%)`, sub: 'Scoring coverage', color: C.light },
  ];

  const kpiTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: kpiTiles.map(() => Math.floor(9360 / kpiTiles.length)),
    rows: [new TableRow({ children: kpiTiles.map(tile => new TableCell({
      borders,
      shading: { fill: tile.color ?? C.light, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 100, right: 100 },
      width: { size: Math.floor(9360 / kpiTiles.length), type: WidthType.DXA },
      children: [
        p([t(tile.value, { bold: true, size: 36, color: C.navy })], { alignment: AlignmentType.CENTER }),
        p([t(tile.label, { bold: true, size: 17, color: C.navy })], { alignment: AlignmentType.CENTER }),
        p([t(tile.sub,   { size: 15, color: C.gray })],              { alignment: AlignmentType.CENTER }),
      ],
    }))})],
  });

  // Store breakdown table
  const breakdownHeaders = ['Metric', ...storeLabels, 'Combined'];
  const breakdownWidths  = [3000, ...storeLabels.map(() => Math.floor(5160 / storeLabels.length)), 1200];

  const storeRow = (label: string, getValue: (r: any) => string, getCombined: () => string) =>
    [label, ...client.stores.map(s => getValue(data.stats.find(r => r.location === s.location) || {})), getCombined()];

  const breakdownRows = [
    storeRow('Total inbound calls', r => String(r.total || '—'), () => String(totals.total || '—')),
    storeRow('Answer rate', r => r.total ? `${Math.round(100*(r.answered||0)/r.total)}%` : '—', () => answerRate),
    storeRow('Missed calls', r => String(r.missed || 0), () => String(totals.missed || 0)),
    storeRow('Sales opportunities', r => r.total ? `${r.sales_opps || 0} (${Math.round(100*(r.sales_opps||0)/r.total)}%)` : '—', () => `${totals.opps || 0} (${totals.total ? Math.round(100*(totals.opps||0)/totals.total) : 0}%)`),
    storeRow('Appointments confirmed', r => String(r.appts_confirmed || 0), () => String(totals.appts || 0)),
    storeRow('Appointments tentative', r => String(r.appts_tentative || 0), () => String(totals.tent || 0)),
    storeRow('Avg talk time', r => r.avg_talk_min ? `${r.avg_talk_min} min` : '—', () => '—'),
    storeRow('Calls scored', r => String(r.scored_calls || 0), () => String(totals.scored || 0)),
    storeRow('Avg call score', r => r.avg_score ? `${r.avg_score} / 50` : '—', () => avgScore !== '—' ? `${avgScore} / 50` : '—'),
  ];

  const makeTable = (headers: string[], rows: string[][], widths: number[]) => {
    const hdrRow = new TableRow({ tableHeader: true, children: headers.map((h, i) => new TableCell({
      borders, shading: { fill: C.navy, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 90, right: 90 },
      width: { size: widths[i], type: WidthType.DXA },
      children: [p([t(h, { bold: true, color: C.white, size: 17 })], { alignment: AlignmentType.CENTER })],
    }))});
    const dataRows = rows.map((row, ri) => new TableRow({ children: row.map((cell, ci) => {
      const isLabel = ci === 0;
      const fill    = ri % 2 === 0 ? C.white : C.offwhite;
      const color   = !isLabel && String(cell).includes('%') ? pctColor(cell) : C.black;
      return new TableCell({
        borders, shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 55, bottom: 55, left: 90, right: 90 },
        width: { size: widths[ci], type: WidthType.DXA },
        children: [p([t(cell, { bold: isLabel, size: 17, color })],
          { alignment: isLabel ? AlignmentType.LEFT : AlignmentType.CENTER })],
      });
    })}));
    return new Table({ width: { size: widths.reduce((a,b)=>a+b,0), type: WidthType.DXA }, columnWidths: widths, rows: [hdrRow, ...dataRows] });
  };

  // Intent table
  const intentMap: Record<string, Record<string, number>> = {};
  for (const row of data.intent) {
    if (!intentMap[row.reason]) intentMap[row.reason] = {};
    intentMap[row.reason][row.location] = Number(row.total);
  }
  const intentTotal = Object.values(intentMap).reduce((s, m) => s + Object.values(m).reduce((a,b)=>a+b,0), 0);
  const intentRows  = Object.entries(intentMap)
    .map(([reason, locs]) => {
      const combined = Object.values(locs).reduce((a,b)=>a+b,0);
      return [reason, ...client.stores.map(s => String(locs[s.location] || 0)), `${combined} (${intentTotal ? Math.round(100*combined/intentTotal) : 0}%)`];
    })
    .sort((a,b) => parseInt(b[b.length-1]) - parseInt(a[a.length-1]))
    .slice(0, 8);

  const intentWidths = [3000, ...client.stores.map(() => Math.floor(4200/client.stores.length)), 2160];

  // Document
  const doc = new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 720, right: 900, bottom: 720, left: 900 } } },
      headers: { default: new Header({ children: [new Paragraph({
        children: [
          t('CONFIDENTIAL  |  Ntelegence Communication Intelligence', { size: 14, color: C.gray }),
          t(`\t\t${client.name}  |  ${fmtDate(dateFrom)} – ${fmtDate(dateTo)}`, { size: 14, color: C.gray }),
        ],
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.mid } },
      })] }) },
      footers: { default: new Footer({ children: [new Paragraph({
        children: [
          t('Prepared by Ntelegence  |  jscherer@ntelegence.com  |  ntelegence.com', { size: 14, color: C.gray }),
          t('\t\tPage ', { size: 14, color: C.gray }),
          new TextRun({ children: [PageNumber.CURRENT], size: 14, color: C.gray, font: 'Arial' }),
        ],
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        border: { top: { style: BorderStyle.SINGLE, size: 2, color: C.mid } },
      })] }) },
      children: [
        p([t(client.name, { bold: true, size: 44, color: C.navy })], { alignment: AlignmentType.CENTER, spacing: { before: 60, after: 40 } }),
        p([t('Call Intelligence Executive Summary  —  Regional', { size: 22, color: C.mid })], { alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
        p([t(`${fmtDate(dateFrom)} to ${fmtDate(dateTo)}  |  ${storeLabels.join('  ·  ')}`, { size: 18, color: C.gray })], { alignment: AlignmentType.CENTER, spacing: { after: 160 } }),

        kpiTable,
        spacer(40),
        p([t('Color key:  ', { size: 16, color: C.gray }), t('Green = strong   ', { size: 16, bold: true, color: C.green }), t('Amber = needs attention   ', { size: 16, bold: true, color: C.amber }), t('Red = critical gap', { size: 16, bold: true, color: C.red })], { spacing: { after: 140 } }),

        sectionHdr('STORE-BY-STORE BREAKDOWN'),
        spacer(80),
        makeTable(['Metric', ...storeLabels, 'Combined'], breakdownRows, breakdownWidths),
        spacer(140),

        sectionHdr('CALL INTENT — REASON FOR CALL'),
        spacer(80),
        ...(intentRows.length
          ? [makeTable(['Reason for Call', ...storeLabels, 'Combined %'], intentRows, intentWidths)]
          : [p([t('No call intent data available for this period.', { size: 17, italics: true, color: C.gray })])]),
        spacer(80),
      ],
    }],
  });

  return Packer.toBuffer(doc);
}

// ── Email delivery ─────────────────────────────────────────────────────────

async function sendReportEmail(opts: {
  to: string[];
  subject: string;
  clientName: string;
  period: string;
  buffer: Buffer;
  filename: string;
}) {
  const resend = new Resend(process.env.RESEND_API_KEY);

  await resend.emails.send({
    from: 'Ntelegence Reports <reports@ntelegence.com>',
    to:   opts.to,
    subject: opts.subject,
    html: `
      <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#333">
        <div style="background:#1F3864;padding:20px 24px;border-radius:4px 4px 0 0">
          <span style="color:#fff;font-size:20px;font-weight:bold">Ntelegence</span>
        </div>
        <div style="padding:28px 24px;background:#f9f9f9;border:1px solid #e5e5e5;border-top:none">
          <h2 style="margin:0 0 8px;color:#1F3864">Call Intelligence Report Ready</h2>
          <p style="margin:0 0 16px;color:#555">Your weekly executive summary for <strong>${opts.clientName}</strong> is attached.</p>
          <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px">
            <tr style="background:#e8f0f7"><td style="padding:8px 12px;font-weight:bold;color:#1F3864;width:40%">Report period</td><td style="padding:8px 12px">${opts.period}</td></tr>
            <tr style="background:#fff"><td style="padding:8px 12px;font-weight:bold;color:#1F3864">Prepared by</td><td style="padding:8px 12px">Ntelegence Communication Intelligence</td></tr>
          </table>
          <p style="margin:0;font-size:13px;color:#888">Open the attached .docx in Microsoft Word or Google Docs.</p>
        </div>
        <div style="padding:12px 24px;background:#1F3864;border-radius:0 0 4px 4px;font-size:12px;color:#aac;text-align:center">
          Ntelegence · No Dead Ends · ntelegence.com
        </div>
      </div>
    `,
    attachments: [{ filename: opts.filename, content: opts.buffer.toString('base64') }],
  });

  console.log(`[Reports] ✉  Sent "${opts.subject}" → ${opts.to.join(', ')}`);
}

// ── Main runner ────────────────────────────────────────────────────────────

async function runReport(clientConfig: ClientConfig, reportType: string) {
  const { dateFrom, dateTo } = rollingWindow(clientConfig.dateWindowDays);
  console.log(`[Reports] ▶  ${clientConfig.name} / ${reportType} — ${dateFrom} → ${dateTo}`);

  const data = await fetchReportData(
    clientConfig.tenantId,
    dateFrom,
    dateTo,
    clientConfig.stores.map(s => s.location)
  );

  let buffer: Buffer;
  if (reportType === 'REGIONAL') {
    buffer = await buildRegionalReport(clientConfig, data, dateFrom, dateTo);
  } else {
    throw new Error(`Report type ${reportType} not yet implemented in scheduler`);
  }

  const period   = `${fmtDate(dateFrom)} – ${fmtDate(dateTo)}`;
  const filename = `${clientConfig.name.replace(/\s+/g, '_')}_${reportType}_${dateTo}.docx`;
  const to       = clientConfig.recipients[reportType] ?? ['jscherer@ntelegence.com'];

  await sendReportEmail({
    to,
    subject:    `${clientConfig.name} — ${reportType} Report | ${period}`,
    clientName: clientConfig.name,
    period,
    buffer,
    filename,
  });

  console.log(`[Reports] ✓  ${clientConfig.name} / ${reportType} complete`);
}

// ── On-demand HTTP endpoint data (called by TeleTraxx /api/reports/generate) ──

export async function generateReportOnDemand(opts: {
  clientId: string;
  reportType: string;
  dateFrom: string;
  dateTo: string;
  recipients?: string[];
}): Promise<void> {
  const clientConfig = CLIENTS.find(c => c.id === opts.clientId);
  if (!clientConfig) throw new Error(`Unknown client: ${opts.clientId}`);

  const data = await fetchReportData(
    clientConfig.tenantId,
    opts.dateFrom,
    opts.dateTo,
    clientConfig.stores.map(s => s.location)
  );

  let buffer: Buffer;
  if (opts.reportType === 'REGIONAL') {
    buffer = await buildRegionalReport(clientConfig, data, opts.dateFrom, opts.dateTo);
  } else {
    throw new Error(`Report type ${opts.reportType} not yet implemented`);
  }

  const period   = `${fmtDate(opts.dateFrom)} – ${fmtDate(opts.dateTo)}`;
  const filename = `${clientConfig.name.replace(/\s+/g, '_')}_${opts.reportType}_${opts.dateTo}.docx`;
  const to       = opts.recipients ?? clientConfig.recipients[opts.reportType] ?? ['jscherer@ntelegence.com'];

  await sendReportEmail({
    to,
    subject:    `${clientConfig.name} — ${opts.reportType} Report | ${period}`,
    clientName: clientConfig.name,
    period,
    buffer,
    filename,
  });
}

// ── Scheduler bootstrap ────────────────────────────────────────────────────

export function startReportScheduler() {
  if (!process.env.RESEND_API_KEY) {
    console.warn('[Reports] RESEND_API_KEY not set — scheduler disabled');
    return;
  }

  let registered = 0;

  for (const client of CLIENTS) {
    for (const [reportType, schedule] of Object.entries(client.schedules)) {
      if (!schedule) continue;
      if (!cron.validate(schedule)) {
        console.warn(`[Reports] Invalid cron "${schedule}" for ${client.id}/${reportType}`);
        continue;
      }

      cron.schedule(schedule, async () => {
        try {
          await runReport(client, reportType);
        } catch (err: any) {
          console.error(`[Reports] ✗  ${client.id}/${reportType} FAILED:`, err.message);
        }
      }, { timezone: 'America/Chicago' });

      console.log(`[Reports] Registered ${client.id}/${reportType} → "${schedule}"`);
      registered++;
    }
  }

  console.log(`[Reports] Scheduler started — ${registered} job(s) registered`);
}
