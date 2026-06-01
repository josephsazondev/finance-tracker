// ============================================================
// FINANCE TRACKER — APPS SCRIPT (Lean Version)
// All aggregation is done by Google Sheets formulas.
// This script only reads/writes rows and serves JSON.
//
// Sheet names (must match exactly):
//   Config, AccountSetup, MonthlyInput, Dashboard
// ============================================================

const S_CONFIG   = 'Config';
const S_ACCOUNTS = 'AccountSetup';
const S_INPUT    = 'MonthlyInput';
const S_DASH     = 'Dashboard';

// ============================================================
// ROUTER
// ============================================================

function doGet(e)  { return route(e); }
function doPost(e) { return route(e); }

function route(e) {
  try {
    const p      = e.parameter || {};
    const path   = p.path   || '';
    const action = p.action || '';

    if (!validateApiKey(p.api_key || ''))
      return respond({ error: 'Invalid API key' }, e);

    // Parse body — comes in as ?body=JSON (JSONP) or postData
    let body = {};
    if (e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch(_) {}
    } else if (p.body) {
      try { body = JSON.parse(p.body); } catch(_) {}
    }

    const act = body.action || action;

    if (path === 'dashboard')                        return getDashboard(e);
    if (path === 'entries'  && !act)                 return getEntries(p, e);
    if (path === 'entries'  && act === 'create')     return createEntry(body, e);
    if (path === 'entries'  && act === 'batch_create') return batchCreateEntries(body, e);
    if (path === 'entries'  && act === 'update')     return updateEntry(body, e);
    if (path === 'entries'  && act === 'delete')     return deleteEntry(body, e);
    if (path === 'entries'  && act === 'copy_previous') return copyPrevious(p, e);
    if (path === 'accounts' && !act)                 return getAccounts(e);
    if (path === 'accounts' && act === 'create')     return createAccount(body, e);
    if (path === 'accounts' && act === 'update')     return updateAccount(body, e);
    if (path === 'accounts' && act === 'delete')     return deleteAccount(body, e);
    if (path === 'accounts' && act === 'set_currency') return setLocalCurrency(body, e);

    return respond({ error: 'Unknown route: ' + path + '/' + act }, e);
  } catch(err) {
    return respond({ error: err.message, stack: err.stack }, e);
  }
}

// ============================================================
// AUTH
// ============================================================

function validateApiKey(key) {
  const rows = getSheet(S_CONFIG).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'api_key') return String(rows[i][1]) === String(key);
  }
  return false;
}

// ============================================================
// DASHBOARD  — reads pre-calculated Dashboard sheet
// ============================================================

function getDashboard(e) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const dash   = ss.getSheetByName(S_DASH);
  const config = ss.getSheetByName(S_CONFIG);

  // ── Config values ──────────────────────────────────────
  const cfgMap = sheetToMap(config);
  const localCurrency = cfgMap['local_currency'] || 'PHP';

  // ── Dashboard rows ─────────────────────────────────────
  // Expected headers (row 1):
  //   month | net_worth | total_assets | total_debt |
  //   cash  | emergency_fund | investment | other_assets |
  //   credit_card | personal_loan | other_debt
  const dashData = dash.getDataRange().getValues();
  if (dashData.length < 2) return respond({ error: 'Dashboard sheet has no data rows' }, e);

  const headers = dashData[0].map(h => String(h).trim().toLowerCase().replace(/ /g, '_'));

  const idx = (name) => headers.indexOf(name);
  const I = {
    month:          idx('month'),
    net_worth:      idx('net_worth'),
    total_assets:   idx('total_assets'),
    total_debt:     idx('total_debt'),
    cash:           idx('cash'),
    emergency_fund: idx('emergency_fund'),
    investment:     idx('investment'),
    other_assets:   idx('other_assets'),
  };

  // Validate required columns exist
  const missing = Object.entries(I).filter(([,v]) => v === -1).map(([k]) => k);
  if (missing.length) return respond({ error: 'Dashboard sheet missing columns: ' + missing.join(', ') }, e);

  // Build history array (skip header row)
  const history = [];
  for (let i = 1; i < dashData.length; i++) {
    const row = dashData[i];
    const month = toYYYYMM(row[I.month]);
    if (!month) continue; // skip blank rows

    history.push({
      month,
      netWorth:    num(row[I.net_worth]),
      totalAssets: num(row[I.total_assets]),
      totalDebt:   num(row[I.total_debt]),
      categoryTotals: {
        'Cash':           num(row[I.cash]),
        'Emergency Fund': num(row[I.emergency_fund]),
        'Investment':     num(row[I.investment]),
        'Other Assets':   num(row[I.other_assets]),
      }
    });
  }

  // Sort by month string ascending
  history.sort((a, b) => a.month.localeCompare(b.month));

  if (!history.length) return respond({ error: 'No rows found in Dashboard sheet' }, e);

  const current = history[history.length - 1];
  const prev    = history.length > 1 ? history[history.length - 2] : null;
  const change  = prev ? current.netWorth - prev.netWorth : 0;
  const changePct = prev && prev.netWorth ? (change / prev.netWorth) * 100 : 0;

  // ── Asset allocation (current month, assets only) ──────
  const assetAllocation = Object.entries(current.categoryTotals)
    .filter(([, v]) => v > 0)
    .map(([category, amount]) => ({ category, amount }));

  // ── Asset goals ────────────────────────────────────────
  const GOAL_CATS = ['Cash', 'Emergency Fund', 'Investment', 'Other Assets'];
  const assetGoals = GOAL_CATS.map(cat => {
    const key    = 'goal_' + cat.replace(/ /g, '_');
    const target = num(cfgMap[key] || cfgMap['goal_' + cat] || 0);
    const curr   = current.categoryTotals[cat] || 0;
    return { category: cat, current: curr, target, progress: target ? Math.min(100, curr / target * 100) : 0 };
  });

  // ── Investment milestones ──────────────────────────────
  const investment = current.categoryTotals['Investment'] || 0;
  const milestones = [];
  for (let n = 1; n <= 10; n++) {
    const t = num(cfgMap['milestone_' + n] || 0);
    if (!t) break;
    milestones.push({
      name:     'Milestone #' + n,
      target:   t,
      current:  investment,
      progress: Math.min(100, investment / t * 100)
    });
  }

  // ── FIRE metrics ───────────────────────────────────────
  const coastTarget = num(cfgMap['coast_fire_target'] || 0);
  const fireMetrics = {
    coastFireTarget:    coastTarget,
    currentInvestment:  investment,
    remaining:          Math.max(0, coastTarget - investment),
    progress:           coastTarget ? Math.min(100, investment / coastTarget * 100) : 0,
    retirementYearly:   investment * 0.04,
    retirementMonthly:  investment * 0.04 / 12,
  };

  return respond({
    localCurrency,
    currentMonth:       current.month,
    netWorth:           current.netWorth,
    totalAssets:        current.totalAssets,
    totalDebt:          current.totalDebt,
    netWorthChange:     change,
    netWorthChangePct:  changePct,
    netWorthHistory:    history,
    assetAllocation,
    assetGoals,
    investmentMilestones: milestones,
    fireMetrics,
  }, e);
}

// ============================================================
// ENTRIES  — month stored/matched as YYYY-MM string
// ============================================================

function getEntries(p, e) {
  const month    = String(p.month || '').trim();
  const sheet    = getSheet(S_INPUT);
  const accSheet = getSheet(S_ACCOUNTS);
  const accMap   = accountLookup(accSheet);
  const rows     = sheet.getDataRange().getValues();
  const entries  = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const rowMonth = toYYYYMM(r[1]);
    if (month && rowMonth !== month) continue;
    const account = String(r[2]);
    const info    = accMap[account] || {};
    entries.push({
      id:          r[0],
      month:       rowMonth,
      account,
      amount:      num(r[3]),
      type:        info.type     || '',
      category:    info.category || '',
      currency:    info.currency || '',
      amountLocal: num(r[7]),
    });
  }

  return respond({ entries }, e);
}

function createEntry(body, e) {
  const sheet  = getSheet(S_INPUT);
  const id     = nextId(sheet);
  const month  = toYYYYMM(body.month) || String(body.month);
  sheet.appendRow([id, month, body.account, num(body.amount)]);
  copyFormulasToLastRow(sheet);
  return respond({ success: true, id }, e);
}

function batchCreateEntries(body, e) {
  const entries = body.entries;
  if (!entries || !entries.length) return respond({ error: 'No entries provided' }, e);

  const sheet = getSheet(S_INPUT);
  let id = nextId(sheet);
  const created = [];

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const month = toYYYYMM(entry.month) || String(entry.month);
    sheet.appendRow([id, month, entry.account, num(entry.amount)]);
    copyFormulasToLastRow(sheet);
    created.push(id);
    id++;
  }

  return respond({ success: true, ids: created, count: created.length }, e);
}

function updateEntry(body, e) {
  const sheet = getSheet(S_INPUT);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.id)) {
      const month = toYYYYMM(body.month) || String(body.month);
      sheet.getRange(i + 1, 2).setValue(month);
      sheet.getRange(i + 1, 3).setValue(body.account);
      sheet.getRange(i + 1, 4).setValue(num(body.amount));
      return respond({ success: true }, e);
    }
  }
  return respond({ error: 'Entry not found' }, e);
}

function deleteEntry(body, e) {
  const sheet = getSheet(S_INPUT);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.id)) {
      sheet.deleteRow(i + 1);
      return respond({ success: true }, e);
    }
  }
  return respond({ error: 'Entry not found' }, e);
}

function copyPrevious(p, e) {
  const month = String(p.month || '').trim();
  if (!month) return respond({ error: 'month required' }, e);

  // Derive previous month purely from YYYY-MM string
  const [y, m] = month.split('-').map(Number);
  const prevMonth = m === 1
    ? (y - 1) + '-12'
    : y + '-' + String(m - 1).padStart(2, '0');

  const sheet    = getSheet(S_INPUT);
  const accSheet = getSheet(S_ACCOUNTS);
  const accMap   = accountLookup(accSheet);
  const rows     = sheet.getDataRange().getValues();

  // All accounts
  const allAccounts = Object.entries(accMap).map(([name, info]) => ({ account: name, ...info }));

  // Previous month amounts keyed by account name
  const prevAmounts = {};
  for (let i = 1; i < rows.length; i++) {
    if (toYYYYMM(rows[i][1]) === prevMonth) {
      prevAmounts[String(rows[i][2])] = num(rows[i][3]);
    }
  }

  const entries = allAccounts.map(acc => ({
    account:     acc.account,
    type:        acc.type,
    category:    acc.category,
    currency:    acc.currency,
    amount:      prevAmounts[acc.account] !== undefined ? prevAmounts[acc.account] : 0,
    hasPrevious: prevAmounts[acc.account] !== undefined,
  }));

  return respond({ entries, previousMonth: prevMonth, targetMonth: month }, e);
}

// ============================================================
// ACCOUNTS — full CRUD
// ============================================================

function getAccounts(e) {
  const cfgMap = sheetToMap(getSheet(S_CONFIG));
  const rows   = getSheet(S_ACCOUNTS).getDataRange().getValues();
  const accounts = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][1]) continue;
    accounts.push({ id: rows[i][0], account: rows[i][1], type: rows[i][2], category: rows[i][3], currency: rows[i][4], creditLimit: rows[i][5] || '' });
  }
  return respond({ localCurrency: cfgMap['local_currency'] || 'PHP', accounts }, e);
}

function createAccount(body, e) {
  const sheet = getSheet(S_ACCOUNTS);
  const id    = nextId(sheet);
  sheet.appendRow([id, body.account, body.type, body.category, body.currency || 'PHP', parseFloat(body.creditLimit) || '']);
  return respond({ success: true, id }, e);
}

function updateAccount(body, e) {
  const sheet = getSheet(S_ACCOUNTS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.id)) {
      sheet.getRange(i + 1, 2).setValue(body.account);
      sheet.getRange(i + 1, 3).setValue(body.type);
      sheet.getRange(i + 1, 4).setValue(body.category);
      sheet.getRange(i + 1, 5).setValue(body.currency || 'PHP');
      sheet.getRange(i + 1, 6).setValue(parseFloat(body.creditLimit) || '');
      return respond({ success: true }, e);
    }
  }
  return respond({ error: 'Account not found' }, e);
}

function deleteAccount(body, e) {
  const sheet = getSheet(S_ACCOUNTS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.id)) {
      sheet.deleteRow(i + 1);
      return respond({ success: true }, e);
    }
  }
  return respond({ error: 'Account not found' }, e);
}

function setLocalCurrency(body, e) {
  const sheet = getSheet(S_CONFIG);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'local_currency') {
      sheet.getRange(i + 1, 2).setValue(body.currency);
      return respond({ success: true, localCurrency: body.currency }, e);
    }
  }
  sheet.appendRow(['local_currency', body.currency]);
  return respond({ success: true, localCurrency: body.currency }, e);
}

// ============================================================
// HELPERS
// ============================================================

function getSheet(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: "' + name + '"');
  return s;
}

// Convert anything to a YYYY-MM string.
// Accepts: "2026-05", "2026-05-01", a Date object, a serial number.
// Rejects anything that can't resolve to a valid month — returns ''.
function toYYYYMM(value) {
  if (!value && value !== 0) return '';
  const s = String(value).trim();

  // Already YYYY-MM or YYYY-MM-DD
  const iso = s.match(/^(\d{4})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2];

  // Google Sheets date serial (number)
  if (/^\d+$/.test(s)) {
    // Sheets epoch: Dec 30 1899
    const d = new Date(1899, 11, 30);
    d.setDate(d.getDate() + parseInt(s));
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    return y + '-' + m;
  }

  // Try native Date parse as last resort
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  }

  return '';
}

function num(v) {
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

// Build a { key: value } map from a two-column config sheet.
// Also handles "goal_Emergency Fund" key style with spaces.
function sheetToMap(sheet) {
  const rows = sheet.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) map[String(rows[i][0]).trim()] = rows[i][1];
  }
  return map;
}

// Build account lookup: { accountName: { type, category, currency } }
function accountLookup(sheet) {
  const rows = sheet.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1]) map[String(rows[i][1])] = { type: rows[i][2], category: rows[i][3], currency: rows[i][4] };
  }
  return map;
}

// Copy formulas from the second-to-last row into the last row for columns E-H.
// This lets Google Sheets autopopulate type, category, currency, and
// amount_in_local_currency via the existing sheet formulas.
function copyFormulasToLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // need at least row 2 (header) + 1 data row to copy from
  const sourceRow = lastRow - 1;
  // Columns E(5) through H(8)
  sheet.getRange(sourceRow, 5, 1, 4).copyTo(
    sheet.getRange(lastRow, 5, 1, 4),
    SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
    false
  );
}

function nextId(sheet) {
  const rows = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < rows.length; i++) {
    const n = parseInt(rows[i][0]);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function respond(data, e) {
  const cb = (e?.parameter || {}).callback;
  if (cb) {
    const out = ContentService.createTextOutput(cb + '(' + JSON.stringify(data) + ')');
    out.setMimeType(ContentService.MimeType.JAVASCRIPT);
    return out;
  }
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}