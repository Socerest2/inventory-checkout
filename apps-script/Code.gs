/**
 * Inventory check-out backend and admin portal.
 * Paste into a Google Apps Script project (with Admin.html) and deploy twice:
 *   public API:   Execute as Me, access Anyone
 *   admin portal: Execute as User accessing, access Anyone with Google account
 * See README.md.
 */

const SPREADSHEET_ID = '18dOTIriWlQVI4ZQFUH_BV6XI2x2KFmoa9ijRz4nNkV4';
const LOG_SHEET = 'Checkout Log';
const LOG_HEADERS = ['ID', 'Checked Out', 'Office', 'Item', 'Box / Zone', 'Qty', 'Name', 'NetID', 'Event', 'Status', 'Checked In'];

// Used until an admin saves settings in the portal.
const DEFAULT_CONFIG = {
  title: 'Inventory Check Out',
  requireEvent: false,
  offices: [
    { name: 'OFSA Office', sheet: 'OFSA Basement Inventory', firstRow: 3 },
    { name: 'DM Office', sheet: 'DM Office', firstRow: 3 },
  ],
};

function doGet(e) {
  if (e.parameter.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('Inventory Admin')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return handle_(() => getData_(e.parameter.office));
}

function doPost(e) {
  return handle_(() => {
    const body = JSON.parse(e.postData.contents);
    return withLock_(() => {
      if (body.action === 'checkout') return checkout_(body);
      if (body.action === 'checkin') return checkin_(body);
      throw new Error('Unknown action: ' + body.action);
    });
  });
}

function handle_(fn) {
  let out;
  try {
    out = { ok: true, data: fn() };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ---------- Settings ----------

function getConfig_() {
  const raw = PropertiesService.getScriptProperties().getProperty('config');
  return raw ? JSON.parse(raw) : DEFAULT_CONFIG;
}

function findOffice_(name) {
  const office = getConfig_().offices.find(o => o.name === name);
  if (!office) throw new Error('Unknown office: ' + name);
  return office;
}

function saveConfig_(c) {
  const sheetNames = ss_().getSheets().map(s => s.getName());
  const offices = (c.offices || []).map(o => ({
    name: String(o.name || '').trim(),
    sheet: String(o.sheet || ''),
    firstRow: Math.max(1, Math.floor(Number(o.firstRow)) || 1),
  }));
  if (!offices.length) throw new Error('Add at least one office.');
  offices.forEach(o => {
    if (!o.name) throw new Error('Every office needs a name.');
    if (!sheetNames.includes(o.sheet)) throw new Error('Sheet tab not found: ' + o.sheet);
  });
  if (new Set(offices.map(o => o.name)).size !== offices.length) throw new Error('Office names must be different.');
  const cfg = {
    title: String(c.title || '').trim() || DEFAULT_CONFIG.title,
    requireEvent: !!c.requireEvent,
    offices,
    admins: getConfig_().admins || [],
  };
  writeConfig_(cfg);
  return cfg;
}

function writeConfig_(cfg) {
  PropertiesService.getScriptProperties().setProperty('config', JSON.stringify(cfg));
}

// ---------- Inventory ----------

function sheetFor_(office) {
  const sh = ss_().getSheetByName(office.sheet);
  if (!sh) throw new Error('Sheet tab not found: ' + office.sheet);
  return sh;
}

// Columns A:C of every row from firstRow down, blank rows dropped.
function readRows_(office) {
  const sh = sheetFor_(office);
  const first = Number(office.firstRow) || 1;
  const last = sh.getLastRow();
  if (last < first) return [];
  return sh.getRange(first, 1, last - first + 1, 3).getDisplayValues()
    .map((values, i) => ({ row: first + i, values }))
    .filter(r => r.values.some(v => String(v).trim()));
}

// Turns one messy inventory row into one or more {name, box, qty} items.
function parseRow_(box, desc, qty) {
  box = String(box || '').trim();
  desc = String(desc || '').trim();
  qty = String(qty || '').trim();
  if (!desc && !qty) return [];

  // "Yellow Wristband/Orange Wristband" with "80/105" -> two items.
  if (desc) {
    const names = desc.split('/').map(s => s.trim()).filter(String);
    const qtys = qty.split('/').map(s => s.trim());
    if (names.length > 1 && names.length === qtys.length) {
      return names.map((n, i) => ({ name: n, box, qty: qtys[i] }));
    }
    return [{ name: desc, box, qty }];
  }

  // No description: "10 unused FTK ribbons, 3 rolls of raffle tickets" -> one item per part.
  return qty.split(/[,;]/).map(s => s.trim()).filter(String).map(part => {
    const m = part.match(/^(~?\d[\d+]*)\s+(.*)$/);
    return m ? { name: m[2], box, qty: m[1] } : { name: part, box, qty: '' };
  });
}

// Guards admin edits against rows that moved since the portal loaded them.
function checkRow_(sh, office, row, original) {
  if (row < (Number(office.firstRow) || 1)) throw new Error('Row is above the inventory.');
  const current = sh.getRange(row, 1, 1, 3).getDisplayValues()[0];
  if (JSON.stringify(current) !== JSON.stringify(original)) {
    throw new Error('That row changed in the sheet since you loaded it. Reload and try again.');
  }
}

function saveRow_(p) {
  const office = findOffice_(p.office);
  const sh = sheetFor_(office);
  checkRow_(sh, office, p.row, p.original);
  sh.getRange(p.row, 1, 1, 3).setValues([p.values]);
}

function addRow_(p) {
  const office = findOffice_(p.office);
  const rows = readRows_(office);
  const row = rows.length ? rows[rows.length - 1].row + 1 : Number(office.firstRow) || 1;
  sheetFor_(office).getRange(row, 1, 1, 3).setValues([p.values]);
}

function deleteRow_(p) {
  const office = findOffice_(p.office);
  const sh = sheetFor_(office);
  checkRow_(sh, office, p.row, p.original);
  sh.deleteRow(p.row);
}

// ---------- Checkout log ----------

function logSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(LOG_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function logRows_() {
  const sh = logSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, LOG_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2,
    id: String(r[0]),
    out: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    office: r[2], item: r[3], box: r[4], qty: r[5],
    name: r[6], netid: r[7], event: r[8], status: r[9],
    in: r[10] instanceof Date ? r[10].toISOString() : String(r[10] || ''),
  }));
}

function findLog_(id) {
  const entry = logRows_().find(l => l.id === String(id));
  if (!entry) throw new Error('Checkout not found');
  return entry;
}

function getData_(officeName) {
  const cfg = getConfig_();
  const office = cfg.offices.find(o => o.name === officeName) || cfg.offices[0];
  if (!office) throw new Error('No offices are set up yet.');

  const items = [];
  readRows_(office).forEach(r => {
    parseRow_(r.values[0], r.values[1], r.values[2]).forEach(it => items.push(Object.assign(it, { row: r.row })));
  });

  // Log, people and events cover every office: carts can span offices,
  // and autofill should work on any tab.
  const log = logRows_();
  const people = {};
  log.forEach(l => { if (l.netid) people[String(l.netid).toLowerCase()] = l.name; });
  const events = Array.from(new Set(log.map(l => l.event).filter(String)));

  return {
    config: { title: cfg.title, requireEvent: !!cfg.requireEvent, offices: cfg.offices.map(o => o.name) },
    office: office.name,
    items,
    out: log.filter(l => l.status === 'OUT').reverse(),
    recent: log.slice(-25).reverse(),
    people,
    events,
  };
}

// Takes a cart (`items`: [{office, item, box, qty}]) or a single item; one log row per item.
function checkout_(b) {
  const required = ['name', 'netid'].concat(getConfig_().requireEvent ? ['event'] : []);
  required.forEach(k => {
    if (!String(b[k] || '').trim()) throw new Error('Missing ' + k);
  });
  const items = b.items || [{ office: b.office, item: b.item, box: b.box, qty: b.qty }];
  if (!items.length) throw new Error('Nothing to check out');
  items.forEach(it => {
    if (!String(it.item || '').trim()) throw new Error('Missing item');
    findOffice_(it.office);
  });
  const now = new Date();
  const rows = items.map(it => [
    Utilities.getUuid().slice(0, 8), now, it.office, it.item, it.box || '', Number(it.qty) || 1,
    b.name.trim(), b.netid.trim().toLowerCase(), (b.event || '').trim(), 'OUT', '',
  ]);
  const sh = logSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows);
  return { ids: rows.map(r => r[0]) };
}

// Takes `ids` (several) or `id` (one).
function checkin_(b) {
  const log = logRows_();
  const entries = (b.ids || [b.id]).map(id => {
    const entry = log.find(l => l.id === String(id));
    if (!entry) throw new Error('Checkout not found');
    return entry;
  }).filter(e => e.status === 'OUT');
  if (!entries.length) throw new Error('Already checked in');
  const sh = logSheet_();
  const now = new Date();
  entries.forEach(e => sh.getRange(e.row, 10, 1, 2).setValues([['IN', now]]));
  return { ids: entries.map(e => e.id) };
}

function updateLog_(p) {
  const entry = findLog_(p.id);
  const status = p.status === 'IN' ? 'IN' : 'OUT';
  const checkedIn = status === 'OUT' ? '' : (entry.in ? new Date(entry.in) : new Date());
  logSheet_().getRange(entry.row, 3, 1, 9).setValues([[
    p.office, p.item, p.box, Number(p.qty) || 1, p.name,
    String(p.netid || '').trim().toLowerCase(), p.event, status, checkedIn,
  ]]);
}

// ---------- Admin portal ----------

function sheetEditors_() {
  const ss = ss_();
  return ss.getEditors().concat(ss.getOwner() || []).map(u => u.getEmail().toLowerCase());
}

// Admins are the emails in the admin list (Settings). Until that list is
// set, every editor of the spreadsheet is an admin.
function requireAdmin_() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  if (!email) throw new Error('Sign in to Google to use the admin portal.');
  const admins = getConfig_().admins || [];
  if (admins.length) {
    if (!admins.includes(email)) throw new Error(email + ' is not an admin. Ask an admin to add you in Settings.');
    return email;
  }
  let editors;
  try {
    editors = sheetEditors_();
  } catch (err) {
    throw new Error(email + ' does not have edit access to the inventory sheet.');
  }
  if (!editors.includes(email)) throw new Error(email + ' does not have edit access to the inventory sheet.');
  return email;
}

// Adds an admin and, if needed, shares the spreadsheet with them as an
// editor (the portal runs as the signed-in user, so admins need edit access).
function addAdmin_(p, me) {
  const email = String(p.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a full email address.');
  const cfg = getConfig_();
  // The first admin added also locks in whoever is adding, so they aren't shut out.
  const admins = new Set((cfg.admins || []).length ? cfg.admins : [me]);
  admins.add(email);
  let shared = false;
  if (!sheetEditors_().includes(email)) {
    ss_().addEditor(email);
    shared = true;
  }
  cfg.admins = [...admins];
  writeConfig_(cfg);
  return { admins: cfg.admins, shared };
}

function removeAdmin_(p, me) {
  const email = String(p.email || '').trim().toLowerCase();
  if (email === me) throw new Error("You can't remove yourself. Ask another admin.");
  const cfg = getConfig_();
  cfg.admins = (cfg.admins || []).filter(a => a !== email);
  writeConfig_(cfg);
  return { admins: cfg.admins };
}

// Single entry point for Admin.html (google.script.run).
function adminApi(action, p) {
  const email = requireAdmin_();
  return withLock_(() => {
    switch (action) {
      case 'init': {
        const ss = ss_();
        return {
          email,
          config: getConfig_(),
          sheets: ss.getSheets().map(s => s.getName()).filter(n => n !== LOG_SHEET),
          sheetUrl: ss.getUrl(),
        };
      }
      case 'log': return logRows_().reverse();
      case 'updateLog': return updateLog_(p);
      case 'checkin': return checkin_(p);
      case 'deleteLog': logSheet_().deleteRow(findLog_(p.id).row); return null;
      case 'rows': return readRows_(findOffice_(p.office));
      case 'saveRow': return saveRow_(p);
      case 'addRow': return addRow_(p);
      case 'deleteRow': return deleteRow_(p);
      case 'saveConfig': return saveConfig_(p.config);
      case 'addAdmin': return addAdmin_(p, email);
      case 'removeAdmin': return removeAdmin_(p, email);
      default: throw new Error('Unknown action: ' + action);
    }
  });
}
