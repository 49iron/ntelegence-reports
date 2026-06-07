/**
 * Ntelegence Reports — Report Type Definitions
 *
 * Four report tiers, each with different scope, sections, and recipients.
 *
 * CORP        → Multi-client or ownership-level rollup. Highest abstraction.
 *               Audience: PE owners, C-suite, Ntelegence internal.
 *               Sections: portfolio KPIs, cross-location trends, exec findings only.
 *
 * REGIONAL    → Multi-store rollup for a single client/brand.
 *               Audience: VP/Director of Marketing/Ops, regional managers.
 *               Sections: combined KPIs, store comparison, behavioral summary, findings, next steps.
 *
 * STORE       → Single location deep-dive.
 *               Audience: store manager, ops lead.
 *               Sections: store KPIs, call intent, behavioral compliance, agent scorecards, coaching priorities.
 *
 * INDIVIDUAL  → Single agent scorecard report.
 *               Audience: agent + direct manager.
 *               Sections: personal KPIs, criterion-by-criterion detail, trend vs store avg, coaching notes.
 */

const REPORT_TYPES = {

  CORP: {
    id:          'corp',
    label:       'Corporate Summary',
    description: 'Portfolio-level rollup across all clients or a client group',
    sections: [
      'kpi_tiles',
      'portfolio_breakdown',   // table: client × KPI
      'cross_location_trends', // chart data for volume, score trends
      'key_findings',
      'next_steps',
    ],
    agentDetail:  false,
    storeDetail:  false,
    behavioral:   false,       // summary % only, no per-criterion table
  },

  REGIONAL: {
    id:          'regional',
    label:       'Regional Summary',
    description: 'Multi-store rollup for a single brand/client',
    sections: [
      'kpi_tiles',
      'store_breakdown',       // table: store × KPI
      'call_intent',
      'behavioral_summary',    // per-criterion, all stores side-by-side
      'key_findings',
      'next_steps',
    ],
    agentDetail:  false,
    storeDetail:  true,
    behavioral:   true,
  },

  STORE: {
    id:          'store',
    label:       'Store Report',
    description: 'Single location full detail',
    sections: [
      'kpi_tiles',
      'call_intent',
      'behavioral_compliance', // full criterion table for this store
      'agent_scorecards',      // all agents at this store
      'key_findings',
      'next_steps',
    ],
    agentDetail:  true,
    storeDetail:  false,
    behavioral:   true,
  },

  INDIVIDUAL: {
    id:          'individual',
    label:       'Agent Scorecard',
    description: 'Single agent performance report',
    sections: [
      'agent_kpis',            // calls, scored, avg score, opps, appts
      'criterion_detail',      // full breakdown with trend vs store avg
      'coaching_notes',        // AI-generated coaching priorities
    ],
    agentDetail:  true,
    storeDetail:  false,
    behavioral:   true,
  },

};

/**
 * Validate that a report type string is recognized.
 */
function getReportType(typeId) {
  const t = REPORT_TYPES[typeId?.toUpperCase()];
  if (!t) throw new Error(`Unknown report type "${typeId}". Valid: ${Object.keys(REPORT_TYPES).join(', ')}`);
  return t;
}

module.exports = { REPORT_TYPES, getReportType };
