/**
 * ============================================================================
 *  DOMAIN & MAILBOX AUDIT  —  GoDaddy  ×  Apollo  ×  Warmup Inbox
 * ============================================================================
 *  What it does
 *   1. Pulls every domain from GoDaddy (API).
 *   2. Pulls every linked sending mailbox from Apollo (API).
 *   3. Pulls every inbox from Warmup Inbox (API), incl. warmup status & score.
 *      (Anything pasted in the "Warmup Import" tab is merged in as a backup.)
 *   4. Checks each domain's MX record (is email even set up on it?).
 *   5. Matches mailboxes → domains and gives every domain a verdict:
 *        In use / Review / Terminate candidate / Excluded / Inactive.
 *   6. Keeps a running mailbox registry: new mailboxes are added
 *      automatically, removed ones are marked "Removed" (never deleted),
 *      and every change is written to the "Change Log" tab.
 *
 *  Setup (one time)
 *   - Extensions → Apps Script in your Google Sheet, paste this file.
 *   - Project Settings → Script Properties, add:
 *        GODADDY_KEY      your GoDaddy production API key
 *        GODADDY_SECRET   your GoDaddy production API secret
 *          (or GODADDY_PAT  if you use a GoDaddy Personal Access Token)
 *        APOLLO_API_KEY   Apollo MASTER API key (needed for email_accounts)
 *        WARMUP_API_KEY   Warmup Inbox API key (Warmup Inbox → API access)
 *        NOTIFY_EMAIL     (optional) who gets an email when something changes
 *   - Reload the sheet → menu "Domain Audit" → "Set up sheets"
 *   - Add your keep-forever domains (your primary domain etc.) in "Exclusions"
 *   - "Run audit now", then "Install hourly trigger".
 * ============================================================================
 */

const CFG = {
  SHEETS: {
    SUMMARY: 'Summary',
    DOMAINS: 'Domain Audit',
    MAILBOXES: 'Mailboxes',
    WARMUP: 'Warmup Import',
    EXCLUDE: 'Exclusions',
    LOG: 'Change Log',
  },
  GODADDY_BASE: 'https://api.godaddy.com',
  APOLLO_URL: 'https://api.apollo.io/api/v1/email_accounts',
  WARMUP_URL: 'https://api.warmupinbox.com/v1/inboxes',
  CHECK_MX: true,
  TRIGGER_HOURS: 1, // how often the audit refreshes
  DEFAULT_EXCLUSIONS: [['example.com', 'Primary company domain – replace with yours']],
};

const VERDICT = {
  EXCLUDED: 'Excluded (keep)',
  INACTIVE: 'Not active at GoDaddy',
  IN_USE: 'In use',
  REVIEW_INACTIVE: 'Review – only inactive Apollo mailboxes',
  REVIEW_WARMUP: 'Review – warming only, not in Apollo',
  REVIEW_MX: 'Review – email set up, no tracked mailbox',
  TERMINATE: 'Terminate candidate',
};

const VERDICT_COLOR = {
  [VERDICT.EXCLUDED]: '#e8eaed',
  [VERDICT.INACTIVE]: '#e8eaed',
  [VERDICT.IN_USE]: '#d9ead3',
  [VERDICT.REVIEW_INACTIVE]: '#fff2cc',
  [VERDICT.REVIEW_WARMUP]: '#fff2cc',
  [VERDICT.REVIEW_MX]: '#fce5cd',
  [VERDICT.TERMINATE]: '#f4cccc',
};

const HEADERS = {
  DOMAINS: ['Domain', 'Verdict', 'GoDaddy status', 'Expires', 'Auto-renew',
    'MX / email provider', 'Apollo mailboxes (active)', 'Apollo mailboxes (inactive)',
    'Warmup Inbox mailboxes', 'Total unique mailboxes', 'Mailbox list',
    'First seen', 'Last checked'],
  MAILBOXES: ['Email', 'Mailbox domain', 'GoDaddy domain', 'Domain in GoDaddy?',
    'In Apollo?', 'Apollo active?', 'Apollo account type', 'In Warmup Inbox?',
    'Warmup status', 'Warmup score', 'Status', 'First seen', 'Last seen'],
  LOG: ['Timestamp', 'Event', 'Item', 'Detail'],
  EXCLUDE: ['Domain', 'Reason'],
  WARMUP: ['Optional backup: paste extra Warmup Inbox emails here (any layout – every email on this tab is read)'],
};

// ---------------------------------------------------------------------------
// Menu & triggers
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Domain Audit')
    .addItem('Run audit now', 'runAudit')
    .addSeparator()
    .addItem('Set up sheets', 'setupSheets')
    .addItem('Install hourly trigger', 'installTrigger')
    .addItem('Remove triggers', 'removeTriggers')
    .addToUi();
}

function installTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('runAudit').timeBased().everyHours(CFG.TRIGGER_HOURS).create();
  toast_('Audit will refresh every ' + CFG.TRIGGER_HOURS + ' hour(s).');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runAudit')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const ensure = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (headers && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  ensure(CFG.SHEETS.SUMMARY);
  ensure(CFG.SHEETS.DOMAINS, HEADERS.DOMAINS);
  ensure(CFG.SHEETS.MAILBOXES, HEADERS.MAILBOXES);
  ensure(CFG.SHEETS.WARMUP, HEADERS.WARMUP);
  const ex = ensure(CFG.SHEETS.EXCLUDE, HEADERS.EXCLUDE);
  if (ex.getLastRow() === 1) {
    ex.getRange(2, 1, CFG.DEFAULT_EXCLUSIONS.length, 2).setValues(CFG.DEFAULT_EXCLUSIONS);
  }
  ensure(CFG.SHEETS.LOG, HEADERS.LOG);
  toast_('Sheets ready. Add keep-forever domains in "Exclusions", then Run audit now.');
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
function runAudit() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // another run is in progress
  try {
    setupSheets();
    const ss = SpreadsheetApp.getActive();
    const now = new Date();

    // 1. Pull sources
    const gdDomains = fetchGoDaddyDomains_();
    const apollo = fetchApolloMailboxes_();
    const warmup = fetchWarmupMailboxes_().concat(
      readWarmupMailboxes_(ss).map(email => ({ email, status: 'manual', score: '' })));
    const exclusions = readExclusions_(ss);
    const mx = CFG.CHECK_MX ? lookupMx_(gdDomains.map(d => d.domain)) : {};

    // 2. Previous state (for first-seen + change detection)
    const prevDomains = readKeyed_(ss.getSheetByName(CFG.SHEETS.DOMAINS));
    const prevMailboxes = readKeyed_(ss.getSheetByName(CFG.SHEETS.MAILBOXES));

    // 3. Build
    const result = buildAudit_({ gdDomains, apollo, warmup, exclusions, mx, prevDomains, prevMailboxes, now });

    // 4. Write
    writeTable_(ss.getSheetByName(CFG.SHEETS.DOMAINS), HEADERS.DOMAINS, result.domainRows, 1);
    writeTable_(ss.getSheetByName(CFG.SHEETS.MAILBOXES), HEADERS.MAILBOXES, result.mailboxRows);
    writeSummary_(ss.getSheetByName(CFG.SHEETS.SUMMARY), result, now,
      { gd: gdDomains.length, apollo: apollo.length, warmup: warmup.length });
    if (result.changes.length) {
      const log = ss.getSheetByName(CFG.SHEETS.LOG);
      log.getRange(log.getLastRow() + 1, 1, result.changes.length, 4).setValues(result.changes);
      notify_(result.changes);
    }
  } catch (e) {
    console.error(e);
    logError_(e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Core matching logic (pure – no Google services, easy to test)
// ---------------------------------------------------------------------------
function buildAudit_(input) {
  const { gdDomains, apollo, warmup, exclusions, mx, prevDomains, prevMailboxes, now } = input;
  const ts = now;
  const changes = [];
  // On the very first run everything is "new" — don't flood the log/email.
  const firstMailboxRun = Object.keys(prevMailboxes).length === 0;
  const firstDomainRun = Object.keys(prevDomains).length === 0;

  const gdSet = gdDomains.map(d => d.domain);
  // Map a mailbox host (incl. subdomains like mail.foo.com) to its GoDaddy domain.
  const toGdDomain = host => {
    let best = '';
    gdSet.forEach(d => {
      if ((host === d || host.endsWith('.' + d)) && d.length > best.length) best = d;
    });
    return best;
  };

  // ---- Merge mailboxes from both sources ----
  const mb = {}; // email -> record
  const touch = email => {
    if (!mb[email]) {
      const host = email.split('@')[1] || '';
      mb[email] = { email, host, gd: toGdDomain(host), inApollo: false, apolloActive: false,
        apolloType: '', inWarmup: false, warmupStatus: '', warmupScore: '' };
    }
    return mb[email];
  };
  apollo.forEach(a => {
    const r = touch(a.email);
    r.inApollo = true;
    r.apolloActive = r.apolloActive || a.active;
    r.apolloType = a.type || r.apolloType;
  });
  warmup.forEach(w => {
    const r = touch(w.email);
    r.inWarmup = true;
    // API data wins over a manual paste
    if (!r.warmupStatus || r.warmupStatus === 'manual') r.warmupStatus = w.status || '';
    if (w.score !== '' && w.score !== undefined && w.score !== null) r.warmupScore = w.score;
  });

  // ---- Mailbox rows (current + previously seen, now removed) ----
  const mailboxRows = [];
  Object.keys(mb).sort().forEach(email => {
    const r = mb[email];
    const prev = prevMailboxes[email];
    const firstSeen = prev ? prev['First seen'] : ts;
    if (!prev && !firstMailboxRun) {
      changes.push([ts, 'New mailbox', email,
        [r.inApollo ? 'Apollo' : '', r.inWarmup ? 'Warmup Inbox' : ''].filter(Boolean).join(' + ')]);
    } else if (prev && prev['Status'] === 'Removed') {
      changes.push([ts, 'Mailbox re-appeared', email, '']);
    }
    mailboxRows.push([email, r.host, r.gd || '—', r.gd ? 'Yes' : 'No (other registrar?)',
      yn_(r.inApollo), r.inApollo ? yn_(r.apolloActive) : '—', r.apolloType || '—',
      yn_(r.inWarmup), r.warmupStatus || '—', r.warmupScore === '' ? '—' : r.warmupScore,
      'Active', firstSeen, ts]);
  });
  Object.keys(prevMailboxes).forEach(email => {
    if (mb[email]) return;
    const p = prevMailboxes[email];
    if (p['Status'] !== 'Removed') changes.push([ts, 'Mailbox removed', email, 'No longer in Apollo or Warmup Inbox']);
    mailboxRows.push(HEADERS.MAILBOXES.map(h => {
      if (h === 'Status') return 'Removed';
      if (h === 'In Apollo?' || h === 'In Warmup Inbox?') return 'No';
      return p[h] === undefined ? '' : p[h];
    }));
  });

  // ---- Domain rows ----
  const byDomain = {};
  Object.values(mb).forEach(r => {
    if (!r.gd) return;
    (byDomain[r.gd] = byDomain[r.gd] || []).push(r);
  });

  const domainRows = [];
  const counts = {};
  gdDomains.slice().sort((a, b) => a.domain.localeCompare(b.domain)).forEach(d => {
    const list = byDomain[d.domain] || [];
    const apActive = list.filter(r => r.inApollo && r.apolloActive).length;
    const apInactive = list.filter(r => r.inApollo && !r.apolloActive).length;
    const wu = list.filter(r => r.inWarmup).length;
    const mxVal = mx[d.domain] || (CFG.CHECK_MX ? 'Lookup failed' : '—');
    const status = String(d.status || '').toUpperCase();

    let verdict;
    if (exclusions[d.domain] !== undefined) verdict = VERDICT.EXCLUDED;
    else if (status && status !== 'ACTIVE') verdict = VERDICT.INACTIVE;
    else if (apActive > 0) verdict = VERDICT.IN_USE;
    else if (apInactive > 0) verdict = VERDICT.REVIEW_INACTIVE;
    else if (wu > 0) verdict = VERDICT.REVIEW_WARMUP;
    else if (/google|microsoft/i.test(mxVal)) verdict = VERDICT.REVIEW_MX;
    else verdict = VERDICT.TERMINATE;
    counts[verdict] = (counts[verdict] || 0) + 1;

    const prev = prevDomains[d.domain];
    if (!prev && !firstDomainRun) changes.push([ts, 'New domain', d.domain, 'Found in GoDaddy']);
    else if (prev && prev['Verdict'] !== verdict) changes.push([ts, 'Verdict changed', d.domain, prev['Verdict'] + ' → ' + verdict]);

    domainRows.push([d.domain, verdict, d.status || '', d.expires ? new Date(d.expires) : '',
      d.renewAuto === undefined ? '' : yn_(d.renewAuto), mxVal, apActive, apInactive, wu,
      list.length, list.map(r => r.email).sort().join(', '),
      prev ? prev['First seen'] : ts, ts]);
  });
  Object.keys(prevDomains).forEach(dom => {
    if (!gdSet.includes(dom)) changes.push([ts, 'Domain removed', dom, 'No longer in GoDaddy']);
  });

  const orphanMailboxes = Object.values(mb).filter(r => !r.gd).map(r => r.email);
  return { domainRows, mailboxRows, changes, counts, orphanMailboxes };
}

// ---------------------------------------------------------------------------
// Sources
// ---------------------------------------------------------------------------
function fetchGoDaddyDomains_() {
  const p = props_();
  const pat = p.getProperty('GODADDY_PAT');
  const key = p.getProperty('GODADDY_KEY');
  const secret = p.getProperty('GODADDY_SECRET');
  if (!pat && !(key && secret)) {
    throw new Error('Missing GoDaddy credentials: set GODADDY_KEY + GODADDY_SECRET (or GODADDY_PAT) in Script Properties.');
  }
  const auth = pat ? 'Bearer ' + pat : 'sso-key ' + key + ':' + secret;
  const out = [];
  let marker = '';
  for (let page = 0; page < 50; page++) {
    const url = CFG.GODADDY_BASE + '/v1/domains?limit=1000' + (marker ? '&marker=' + encodeURIComponent(marker) : '');
    const res = httpJson_(url, { Authorization: auth }, 'GoDaddy');
    if (!Array.isArray(res) || !res.length) break;
    res.forEach(d => out.push({
      domain: String(d.domain).toLowerCase().trim(),
      status: d.status,
      expires: d.expires,
      renewAuto: d.renewAuto,
    }));
    if (res.length < 1000) break;
    marker = res[res.length - 1].domain;
  }
  return out;
}

function fetchApolloMailboxes_() {
  const key = props_().getProperty('APOLLO_API_KEY');
  if (!key) throw new Error('Missing APOLLO_API_KEY (must be a master key) in Script Properties.');
  const json = httpJson_(CFG.APOLLO_URL, { 'X-Api-Key': key, 'Cache-Control': 'no-cache' }, 'Apollo');
  const list = json.email_accounts || json.data || [];
  return list.map(a => ({
    email: String(a.email || '').toLowerCase().trim(),
    active: a.active !== false,
    type: a.type || a.provider || '',
  })).filter(a => a.email.includes('@'));
}

/**
 * Warmup Inbox API: GET /v1/inboxes  (header: x-api-key)
 * Each inbox has: inbox_id, email, status (running/paused/banned/error/suspended),
 * type (google/office/smtp_imap), score (landing score).
 */
function fetchWarmupMailboxes_() {
  const key = props_().getProperty('WARMUP_API_KEY');
  if (!key) { console.warn('WARMUP_API_KEY not set – skipping Warmup Inbox API.'); return []; }
  const json = httpJson_(CFG.WARMUP_URL, { 'x-api-key': key }, 'Warmup Inbox');
  const list = Array.isArray(json) ? json : (json.inboxes || json.data || json.items || json.results || []);
  return list.map(i => ({
    email: String(i.email || '').toLowerCase().trim(),
    status: i.status || '',
    score: i.score === undefined || i.score === null ? '' : i.score,
  })).filter(i => i.email.includes('@'));
}

/**
 * Backup source: reads every email address found anywhere on the
 * "Warmup Import" tab (useful for inboxes in a second Warmup Inbox account).
 */
function readWarmupMailboxes_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.WARMUP);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
  const re = /[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/gi;
  const set = {};
  values.forEach(row => row.forEach(cell => {
    (String(cell).match(re) || []).forEach(e => { set[e.toLowerCase()] = true; });
  }));
  return Object.keys(set);
}

function readExclusions_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.EXCLUDE);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([d, reason]) => {
    if (d) out[String(d).toLowerCase().trim()] = reason || '';
  });
  return out;
}

/** MX lookup via Google Public DNS (no key needed), batched. */
function lookupMx_(domains) {
  const out = {};
  for (let i = 0; i < domains.length; i += 50) {
    const chunk = domains.slice(i, i + 50);
    const reqs = chunk.map(d => ({ url: 'https://dns.google/resolve?type=MX&name=' + encodeURIComponent(d), muteHttpExceptions: true }));
    let resps;
    try { resps = UrlFetchApp.fetchAll(reqs); } catch (e) { continue; }
    resps.forEach((r, idx) => {
      const d = chunk[idx];
      try {
        const j = JSON.parse(r.getContentText());
        const hosts = (j.Answer || []).filter(a => a.type === 15)
          .map(a => a.data.split(' ').pop().replace(/\.$/, '').toLowerCase());
        if (!hosts.length) out[d] = 'No MX (no email)';
        else if (hosts.some(h => h.includes('google'))) out[d] = 'Google Workspace';
        else if (hosts.some(h => h.includes('outlook') || h.includes('microsoft'))) out[d] = 'Microsoft 365';
        else out[d] = 'Other: ' + hosts[0];
      } catch (e) { out[d] = 'Lookup failed'; }
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------
/** Reads a sheet into { firstColumnValue: {header: value} }. */
function readKeyed_(sh) {
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  data.slice(1).forEach(row => {
    const k = String(row[0] || '').toLowerCase().trim();
    if (!k) return;
    const o = {};
    headers.forEach((h, i) => { o[h] = row[i]; });
    out[k] = o;
  });
  return out;
}

function writeTable_(sh, headers, rows, verdictColIdx) {
  sh.clearContents();
  sh.clearFormats();
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  if (verdictColIdx !== undefined) {
    const bg = rows.map(r => headers.map(() => VERDICT_COLOR[r[verdictColIdx]] || null));
    sh.getRange(2, 1, rows.length, headers.length).setBackgrounds(bg);
  }
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, rows.length + 1, headers.length).createFilter();
  sh.autoResizeColumns(1, Math.min(headers.length, 10));
}

function writeSummary_(sh, result, now, src) {
  sh.clear();
  const rows = [
    ['Domain & Mailbox Audit', ''],
    ['Last refreshed', now],
    ['', ''],
    ['Domains in GoDaddy', src.gd],
    ['Mailboxes in Apollo', src.apollo],
    ['Mailboxes in Warmup Inbox', src.warmup],
    ['', ''],
    ['Verdict', 'Domains'],
  ];
  Object.values(VERDICT).forEach(v => rows.push([v, result.counts[v] || 0]));
  rows.push(['', '']);
  rows.push(['Mailboxes on domains NOT in GoDaddy', result.orphanMailboxes.length]);
  rows.push(['  (list)', result.orphanMailboxes.join(', ')]);
  rows.push(['Changes this run', result.changes.length]);
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  sh.getRange('A8:B8').setFontWeight('bold');
  Object.values(VERDICT).forEach((v, i) => sh.getRange(9 + i, 1, 1, 2).setBackground(VERDICT_COLOR[v]));
  sh.autoResizeColumns(1, 2);
}

function notify_(changes) {
  const to = props_().getProperty('NOTIFY_EMAIL');
  if (!to) return;
  const important = changes.filter(c => c[1] !== 'Verdict changed' || /Terminate/.test(c[3]));
  if (!important.length) return;
  const body = important.map(c => '• ' + c[1] + ': ' + c[2] + (c[3] ? ' (' + c[3] + ')' : '')).join('\n') +
    '\n\nSheet: ' + SpreadsheetApp.getActive().getUrl();
  MailApp.sendEmail(to, 'Domain & Mailbox Audit – ' + important.length + ' change(s)', body);
}

function logError_(e) {
  try {
    const log = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.LOG);
    if (log) log.appendRow([new Date(), 'ERROR', '', String(e && e.message || e)]);
  } catch (ignore) {}
}

// ---------------------------------------------------------------------------
// Generic helpers
// ---------------------------------------------------------------------------
function httpJson_(url, headers, label) {
  for (let attempt = 0; attempt < 4; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: Object.assign({ Accept: 'application/json' }, headers),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText());
    if (code === 429 || code >= 500) { Utilities.sleep(2000 * (attempt + 1)); continue; }
    let hint = '';
    if (label === 'GoDaddy' && (code === 401 || code === 403)) hint = ' — check key/secret are PRODUCTION keys and the account has API access.';
    if (label === 'Apollo' && code === 403) hint = ' — this endpoint needs an Apollo MASTER API key.';
    throw new Error(label + ' API ' + code + ': ' + res.getContentText().slice(0, 300) + hint);
  }
  throw new Error(label + ' API: rate-limited after retries.');
}

function props_() { return PropertiesService.getScriptProperties(); }
function yn_(b) { return b ? 'Yes' : 'No'; }
function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg, 'Domain Audit', 6); } catch (e) {} }
