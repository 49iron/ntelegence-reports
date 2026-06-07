/**
 * Ntelegence Reports — Database Layer
 * Connects to the same Render PostgreSQL instance as TeleTraxx.
 * All queries return plain JS objects ready for report templates.
 */

const { Pool } = require('pg');

let _pool = null;

function pool() {
  if (!_pool) {
    _pool = new Pool({
      connectionString: process.env.DATABASE_URL,
      ssl: { rejectUnauthorized: false }, // required for Render
    });
  }
  return _pool;
}

async function query(sql, params = []) {
  const client = await pool().connect();
  try {
    const res = await client.query(sql, params);
    return res.rows;
  } finally {
    client.release();
  }
}

// ── Core queries ──────────────────────────────────────────────────────────

/**
 * Returns basic call volume stats for a tenant + date range.
 * Splits by location if locations array is provided.
 */
async function getCallStats(tenantId, dateFrom, dateTo, locations = []) {
  const rows = await query(`
    SELECT
      location,
      COUNT(*)                                                        AS total,
      COUNT(*) FILTER (WHERE answered = true)                         AS answered,
      COUNT(*) FILTER (WHERE answered = false)                        AS missed,
      COUNT(*) FILTER (WHERE (custom_field_values->>'sales_opportunity') = 'YES') AS sales_opps,
      COUNT(*) FILTER (WHERE (custom_field_values->>'appointment_booked') = 'YES') AS appts_confirmed,
      COUNT(*) FILTER (WHERE (custom_field_values->>'appointment_booked') = 'TENTATIVE') AS appts_tentative,
      ROUND(AVG(duration_seconds)::numeric / 60, 1)                  AS avg_talk_min,
      COUNT(*) FILTER (WHERE duration_seconds < 30 AND answered = true) AS short_calls,
      COUNT(*) FILTER (WHERE (custom_field_values->>'call_score') IS NOT NULL
                         AND (custom_field_values->>'call_score') != '') AS scored_calls,
      ROUND(AVG(
        CASE WHEN (custom_field_values->>'call_score') ~ '^[0-9]+(\.[0-9]+)?$'
             THEN (custom_field_values->>'call_score')::numeric END
      )::numeric, 1)                                                  AS avg_score
    FROM calls
    WHERE tenant_id   = $1
      AND call_date  >= $2
      AND call_date  <= $3
      ${locations.length ? `AND location = ANY($4)` : ''}
    GROUP BY location
    ORDER BY total DESC
  `, locations.length ? [tenantId, dateFrom, dateTo, locations] : [tenantId, dateFrom, dateTo]);

  return rows;
}

/**
 * Call intent / reason for call breakdown.
 */
async function getCallIntent(tenantId, dateFrom, dateTo) {
  return query(`
    SELECT
      COALESCE(custom_field_values->>'call_reason', 'Unknown') AS reason,
      location,
      COUNT(*) AS total
    FROM calls
    WHERE tenant_id  = $1
      AND call_date >= $2
      AND call_date <= $3
    GROUP BY reason, location
    ORDER BY COUNT(*) DESC
  `, [tenantId, dateFrom, dateTo]);
}

/**
 * Behavioral scoring compliance per criterion per location.
 * Returns rows: { criterion, location, pct }
 */
async function getBehavioralScores(tenantId, dateFrom, dateTo) {
  return query(`
    SELECT
      ar.criterion_name,
      c.location,
      COUNT(*)                                       AS total_scored,
      COUNT(*) FILTER (WHERE ar.passed = true)       AS passed,
      ROUND(
        100.0 * COUNT(*) FILTER (WHERE ar.passed = true) / NULLIF(COUNT(*), 0)
      )                                              AS pct
    FROM assessment_results ar
    JOIN assessments        a  ON ar.assessment_id = a.id
    JOIN calls              c  ON a.call_id        = c.id
    WHERE c.tenant_id  = $1
      AND c.call_date >= $2
      AND c.call_date <= $3
    GROUP BY ar.criterion_name, c.location
    ORDER BY ar.criterion_name, c.location
  `, [tenantId, dateFrom, dateTo]);
}

/**
 * Per-agent stats + behavioral scores.
 * Returns agents sorted by avg_score DESC.
 */
async function getAgentStats(tenantId, dateFrom, dateTo, location = null) {
  const agents = await query(`
    SELECT
      c.agent_name,
      c.location,
      COUNT(*)                                                          AS total_calls,
      COUNT(*) FILTER (WHERE a.id IS NOT NULL)                          AS scored_calls,
      ROUND(AVG(
        CASE WHEN (c.custom_field_values->>'call_score') ~ '^[0-9]+(\.[0-9]+)?$'
             THEN (c.custom_field_values->>'call_score')::numeric END
      )::numeric, 1)                                                    AS avg_score
    FROM calls c
    LEFT JOIN assessments a ON a.call_id = c.id
    WHERE c.tenant_id  = $1
      AND c.call_date >= $2
      AND c.call_date <= $3
      ${location ? 'AND c.location = $4' : ''}
    GROUP BY c.agent_name, c.location
    HAVING COUNT(*) FILTER (WHERE a.id IS NOT NULL) > 0
    ORDER BY avg_score DESC NULLS LAST
  `, location ? [tenantId, dateFrom, dateTo, location] : [tenantId, dateFrom, dateTo]);

  // Fetch criterion-level scores for each agent
  for (const agent of agents) {
    const scores = await query(`
      SELECT
        ar.criterion_name,
        ROUND(100.0 * COUNT(*) FILTER (WHERE ar.passed) / NULLIF(COUNT(*), 0)) AS pct
      FROM assessment_results ar
      JOIN assessments a ON ar.assessment_id = a.id
      JOIN calls       c ON a.call_id        = c.id
      WHERE c.tenant_id  = $1
        AND c.call_date >= $2
        AND c.call_date <= $3
        AND c.agent_name = $4
        ${location ? 'AND c.location = $5' : ''}
      GROUP BY ar.criterion_name
      ORDER BY ar.criterion_name
    `, location ? [tenantId, dateFrom, dateTo, agent.agent_name, location]
               : [tenantId, dateFrom, dateTo, agent.agent_name]);

    agent.criterion_scores = scores; // [{ criterion_name, pct }]
  }

  return agents;
}

/**
 * Convenience: pull everything needed for a full report in one call.
 */
async function getReportData(tenantId, dateFrom, dateTo, locations = []) {
  const [callStats, callIntent, behavioral, agents] = await Promise.all([
    getCallStats(tenantId, dateFrom, dateTo, locations),
    getCallIntent(tenantId, dateFrom, dateTo),
    getBehavioralScores(tenantId, dateFrom, dateTo),
    getAgentStats(tenantId, dateFrom, dateTo),
  ]);
  return { callStats, callIntent, behavioral, agents };
}

async function closePool() {
  if (_pool) await _pool.end();
}

module.exports = { query, getCallStats, getCallIntent, getBehavioralScores, getAgentStats, getReportData, closePool };
