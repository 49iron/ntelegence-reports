/**
 * Ntelegence Reports — Client Registry
 *
 * Single source of truth for every client, their DB tenant IDs,
 * store locations, report schedules, and recipients per report type.
 *
 * TO ADD A CLIENT: copy the _TEMPLATE block at the bottom, fill in, and add to CLIENTS.
 * TO ADD A STORE:  add to the client's stores[] array.
 * TO CHANGE SCHEDULE: edit the cron string (standard 5-field cron).
 */

const CLIENTS = {

  // ── RNR Tire Midwest ───────────────────────────────────────────────────
  'rnr-tire-midwest': {
    id:           'rnr-tire-midwest',
    name:         'RNR Tire Midwest',
    tenantId:     'e36fb1c0-3af3-44d1-ba03-db9529854058',
    brand:        'RNR Tire Express',
    active:       true,

    stores: [
      { id: 'stl',       label: 'St. Louis',  location: 'St. Louis',  trackingLabel: 'St. Louis Sales' },
      { id: 'evansville',label: 'Evansville', location: 'Evansville', trackingLabel: 'Evansville Sales' },
    ],

    // Behavioral rubric used for scoring (must match assessment_rubrics in DB)
    rubric: 'Thanks for Calling LIVE!',

    // Recipients per report type
    recipients: {
      REGIONAL:   ['breanne.braddock@rnrtire.com', 'jscherer@ntelegence.com'],
      STORE:      ['breanne.braddock@rnrtire.com', 'jscherer@ntelegence.com'],
      INDIVIDUAL: ['jscherer@ntelegence.com'],  // add store manager when confirmed
      CORP:       ['jscherer@ntelegence.com'],
    },

    // Cron schedules (when to auto-send each report type)
    // Format: minute hour day-of-month month day-of-week
    schedules: {
      REGIONAL:   '0 7 * * 1',   // every Monday 7am
      STORE:      '0 7 * * 1',   // every Monday 7am
      INDIVIDUAL: null,           // on-demand only for now
      CORP:       null,
    },

    // Date window for scheduled runs (rolling)
    dateWindow: { days: 7 },    // last 7 days

    // Unverified agents — excluded from scorecards until Breanne confirms
    unverifiedAgents: ['Beckwith', 'Herman', 'Dante', 'Dennis', 'Shamiya'],
  },

  // ── RNR Tire MO (B&B Rentals) ─────────────────────────────────────────
  'rnr-tire-mo': {
    id:       'rnr-tire-mo',
    name:     'RNR Tire MO',
    tenantId: '833fd1d4-c5b6-4764-bdd5-afc3be3f0308',
    brand:    'RNR Tire Express',
    active:   true,

    stores: [
      { id: 'raytown',   label: 'Raytown',    location: 'Raytown',    trackingLabel: 'Raytown' },
      { id: 'gladstone', label: 'Gladstone',  location: 'Gladstone',  trackingLabel: 'Gladstone' },
      { id: 'stjoe',     label: 'St. Joseph', location: 'St. Joseph', trackingLabel: 'St. Joseph',
        // St. Joseph has persistent answer rate issues — suppress sensitive columns in client-facing reports
        suppressColumns: ['answer_rate', 'missed_calls'] },
    ],

    rubric: null, // assessment_enabled = FALSE for RNR MO

    recipients: {
      REGIONAL:   ['sal.pernice@rnrtire.com', 'jscherer@ntelegence.com'],
      STORE:      ['jscherer@ntelegence.com'],
      INDIVIDUAL: ['jscherer@ntelegence.com'],
      CORP:       ['jscherer@ntelegence.com'],
    },

    schedules: {
      REGIONAL:   '0 7 * * 1',   // every Monday 7am
      STORE:      null,
      INDIVIDUAL: null,
      CORP:       null,
    },

    dateWindow: { days: 7 },

    unverifiedAgents: [],
  },

  // ── Big Brand Tire ─────────────────────────────────────────────────────
  'big-brand-tire': {
    id:       'big-brand-tire',
    name:     'Big Brand Tire',
    tenantId: null,             // BBT lives in CTM white-label, not TeleTraxx DB — manual data entry for now
    brand:    'Big Brand Tire',
    active:   false,            // set true when TeleTraxx tenant is created

    stores: [],                 // populated when active

    rubric: null,

    recipients: {
      REGIONAL:   ['matt.guilford@bigbrandtire.com', 'jscherer@ntelegence.com'],
      STORE:      ['jscherer@ntelegence.com'],
      INDIVIDUAL: ['jscherer@ntelegence.com'],
      CORP:       ['jscherer@ntelegence.com'],
    },

    schedules: {
      REGIONAL:   '0 7 * * 1',
      STORE:      null,
      INDIVIDUAL: null,
      CORP:       null,
    },

    dateWindow: { days: 7 },
    unverifiedAgents: [],
  },

  // ── City Toyota (demo) ─────────────────────────────────────────────────
  'city-toyota': {
    id:       'city-toyota',
    name:     'City Toyota',
    tenantId: null,             // demo/simulator only
    brand:    'Toyota',
    active:   false,

    stores: [
      { id: 'main', label: 'Main Store', location: 'Main Store', trackingLabel: 'Main Store' },
    ],

    rubric: null,

    recipients: {
      REGIONAL:   ['jscherer@ntelegence.com'],
      STORE:      ['jscherer@ntelegence.com'],
      INDIVIDUAL: ['jscherer@ntelegence.com'],
      CORP:       ['jscherer@ntelegence.com'],
    },

    schedules: {
      REGIONAL:   null,  // demo — no auto-send
      STORE:      null,
      INDIVIDUAL: null,
      CORP:       null,
    },

    dateWindow: { days: 7 },
    unverifiedAgents: [],
  },

};

// ── Helpers ───────────────────────────────────────────────────────────────

function getClient(clientId) {
  const c = CLIENTS[clientId];
  if (!c) throw new Error(`Unknown client "${clientId}". Valid: ${Object.keys(CLIENTS).join(', ')}`);
  return c;
}

function getActiveClients() {
  return Object.values(CLIENTS).filter(c => c.active);
}

function getStore(clientId, storeId) {
  const client = getClient(clientId);
  const store  = client.stores.find(s => s.id === storeId);
  if (!store) throw new Error(`Unknown store "${storeId}" for client "${clientId}"`);
  return store;
}

module.exports = { CLIENTS, getClient, getActiveClients, getStore };
