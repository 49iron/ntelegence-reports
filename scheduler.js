/**
 * Ntelegence Reports — Scheduler
 * Registers cron jobs for every active client × report type.
 * Run: node scheduler.js
 * On Render: set as the start command for the background worker.
 */

require('dotenv').config();
const cron               = require('node-cron');
const { getActiveClients } = require('./lib/clients');
const { generateAndSend }  = require('./lib/runner');
const { closePool }        = require('./lib/db');

function rollingDateRange(days) {
  const to   = new Date();
  const from = new Date();
  from.setDate(from.getDate() - days);
  return {
    dateFrom: from.toISOString().slice(0, 10),
    dateTo:   to.toISOString().slice(0, 10),
  };
}

function registerJobs() {
  const clients = getActiveClients();
  let jobCount  = 0;

  for (const client of clients) {
    for (const [reportType, schedule] of Object.entries(client.schedules)) {
      if (!schedule) continue;

      const tag = `[${client.id} / ${reportType}]`;

      if (!cron.validate(schedule)) {
        console.warn(`⚠  Invalid cron "${schedule}" for ${tag} — skipping`);
        continue;
      }

      cron.schedule(schedule, async () => {
        console.log(`⏰  ${tag} triggered at ${new Date().toISOString()}`);
        try {
          const { dateFrom, dateTo } = rollingDateRange(client.dateWindow?.days ?? 7);
          await generateAndSend({ clientId: client.id, reportType, dateFrom, dateTo });
          console.log(`✓  ${tag} complete`);
        } catch (err) {
          console.error(`✗  ${tag} FAILED:`, err.message);
          // TODO: alert via Resend or Slack webhook
        }
      }, { timezone: 'America/Chicago' });

      console.log(`  ✓ Registered ${tag} → "${schedule}"`);
      jobCount++;
    }
  }

  console.log(`\nScheduler ready — ${jobCount} job(s) registered.`);
}

registerJobs();

// Graceful shutdown
process.on('SIGTERM', async () => {
  console.log('SIGTERM — shutting down scheduler');
  await closePool();
  process.exit(0);
});
