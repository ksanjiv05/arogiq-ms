/**
 * @OnlyCurrentDoc
 *
 * Sheet Forms - Google Apps Script backend
 * ----------------------------------------
 * Open the Google Sheet, choose Extensions > Apps Script, replace the editor
 * contents with this file, then deploy it as a web app:
 *   Execute as: Me    Who has access: Anyone
 * Full steps are in README.md.
 *
 * Optional script properties (Apps Script > Project Settings > Script properties).
 * Changes take effect immediately; no new deployment is needed.
 *   ACCESS_KEY  Key everyone must enter to use the app. Strongly recommended,
 *               because the web app URL is public once it is in config.js.
 *   ADMIN_KEY   Key that also allows saving and deleting column views.
 *               If it is not set, everyone who can use the app can manage views.
 *
 * What the app respects from the sheet:
 *   - Hidden tabs are not shown.
 *   - Formula cells, and columns filled by array formulas, are read-only.
 *   - A tab whose cell A1 holds a QUERY/IMPORTRANGE/FILTER-style formula is read-only.
 *   - Protected ranges (not "Show a warning") are read-only.
 *   - Data validation set to "Reject input" is enforced.
 */

const VERSION = 5;
const MAX_VIEWS = 50;
const MAX_VIEW_BYTES = 8500; // Apps Script allows 9 KB per property value
const LOCK_WAIT_MS = 20000;
const VIEW_PREFIX = 'view:';

const RANGE_REF = "(?:'[^']*'!|[\\w.]+!)?\\$?[A-Z]{1,3}\\$?\\d*\\s*:\\s*\\$?[A-Z]{1,3}\\$?\\d*";

// Formulas that can fill cells beyond their own (array results).
const ARRAY_RE = new RegExp(
  '\\b(ARRAYFORMULA|QUERY|FILTER|SORTN?|UNIQUE|IMPORTRANGE|IMPORTDATA|IMPORTHTML|IMPORTXML|IMPORTFEED|' +
  'FLATTEN|MAP|BYROW|BYCOL|MAKEARRAY|SCAN|SEQUENCE|TRANSPOSE|SPLIT|VSTACK|HSTACK|TOCOL|TOROW|' +
  'CHOOSECOLS|CHOOSEROWS|WRAPROWS|WRAPCOLS|ARRAY_CONSTRAIN|MMULT|FREQUENCY|GOOGLEFINANCE)\\s*\\(' +
  '|^=\\s*\\{' +
  '|^=\\s*' + RANGE_REF + '\\s*$',
  'i'
);

// In cell A1, these fill the whole tab, header row included: a report tab.
const REPORT_RE = new RegExp(
  '^=\\s*(QUERY|IMPORTRANGE|FILTER|SORTN?|UNIQUE|VSTACK|IMPORTDATA|IMPORTHTML|IMPORTXML|IMPORTFEED)\\s*\\(' +
  '|^=\\s*' + RANGE_REF + '\\s*$',
  'i'
);

// ---------------------------------------------------------------- entry points

function doGet() {
  return json_({ ok: true, data: { service: 'Sheet Forms', version: VERSION, status: 'running' } });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = { ok: true, data: route_(req && typeof req === 'object' ? req : {}) };
  } catch (err) {
    out = err && err.code
      ? { ok: false, code: err.code, error: err.message }
      : { ok: false, code: 'error', error: 'Server error: ' + (err && err.message ? err.message : String(err)) };
  }
  return json_(out);
}

function route_(req) {
  const ctx = access_(req.key);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) fail_('setup', "This script is not attached to a Google Sheet. Create it from the sheet's Extensions > Apps Script menu.");

  switch (req.action) {
    case 'info': return info_(ss, ctx);
    case 'read': return read_(ss, req, ctx);
    case 'update': return withLock_(() => update_(ss, req));
    case 'append': return withLock_(() => append_(ss, req));
    case 'saveView': requireAdmin_(ctx); return withLock_(() => saveView_(ss, req));
    case 'deleteView': requireAdmin_(ctx); return withLock_(() => deleteView_(ss, req));
    default: return fail_('bad', 'Unknown request.');
  }
}

// ---------------------------------------------------------------- access

function access_(key) {
  const props = PropertiesService.getScriptProperties();
  const accessKey = String(props.getProperty('ACCESS_KEY') || '').trim();
  const adminKey = String(props.getProperty('ADMIN_KEY') || '').trim();
  const given = String(key || '').trim();

  if (adminKey && given === adminKey) return { admin: true };
  const accessOk = !accessKey || given === accessKey;
  // Slow down guessing: any key that matches neither configured key waits.
  if (given && (accessKey || adminKey) && given !== accessKey) Utilities.sleep(1000);
  if (!accessOk) fail_('auth', given ? 'That access key is not correct.' : 'Enter the access key to open this sheet.');
  return { admin: !adminKey };
}

function requireAdmin_(ctx) {
  if (!ctx.admin) fail_('forbidden', 'Only an admin can save or delete views.');
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) fail_('busy', 'The sheet is busy. Please try again.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- actions

function info_(ss, ctx) {
  return {
    title: ss.getName(),
    tabs: visibleSheets_(ss).map((s) => ({ id: s.getSheetId(), name: s.getName() })),
    canManageViews: ctx.admin
  };
}

function read_(ss, req, ctx) {
  const sheet = sheet_(ss, req.tabId);
  const snap = snapshot_(sheet);
  const rows = [];
  for (let r = 1; r < snap.lastDataRow; r++) {
    const cells = snap.display[r];
    if (cells.every((_, c) => isBlank_(snap, r, c))) continue;
    const f = [];
    for (let c = 0; c < snap.lastCol; c++) if (snap.formulas[r][c]) f.push(c);
    rows.push({ n: r + 1, cells: cells, f: f });
  }
  return {
    headers: snap.headers,
    columns: columnMeta_(sheet, snap),
    rows: rows,
    views: getViews_(sheet),
    canManageViews: ctx.admin
  };
}

function update_(ss, req) {
  const sheet = sheet_(ss, req.tabId);
  const snap = snapshot_(sheet);
  checkHeaders_(snap, req.headers);

  const row = Number(req.row);
  if (!Number.isInteger(row) || row < 2 || row > snap.lastRow) return { missing: true };

  const current = snap.display[row - 1];
  const formulas = snap.formulas[row - 1];
  if (!req.force && !rowMatches_(current, formulas, snap.computed, req.expected)) {
    return { conflict: true, current: current };
  }

  const writes = entries_(req.changes, snap);
  if (!writes.length) fail_('empty', 'There are no changes to save.');

  const prot = protections_(sheet);
  const rules = snap.rules[row - 1];
  writes.forEach((w) => {
    if (snap.computed[w.c] || formulas[w.c]) fail_('formula', '"' + w.name + '" is calculated by a formula and can\'t be edited here.');
    if (prot.isProtected(row, w.c + 1)) fail_('protected', '"' + w.name + '" is protected in the sheet and can\'t be edited here.');
    w.value = cellValue_(w.text, rules[w.c], w.name);
  });

  const skipped = writeCells_(sheet, row, writes, snap.anchors);

  const range = sheet.getRange(row, 1, 1, snap.lastCol);
  const f = [];
  range.getFormulas()[0].forEach((x, c) => { if (x) f.push(c); });
  return { row: row, cells: range.getDisplayValues()[0], f: f, skipped: skipped };
}

function append_(ss, req) {
  const sheet = sheet_(ss, req.tabId);
  const snap = snapshot_(sheet);
  if (!snap.lastCol) fail_('bad', 'This tab has no header row.');
  checkHeaders_(snap, req.headers);

  const prev = snap.lastDataRow; // 1 means only the header row has content
  const target = prev + 1;
  const maxRows = sheet.getMaxRows();
  // A row below the data holding other formulas (a totals row, say) is never
  // written into: a fresh row is inserted above it instead.
  const place = targetInfo_(sheet, snap);
  const inserted = place.foreign;

  let writes = entries_(req.values, snap)
    .filter((w) => !snap.computed[w.c] && !snap.auto[w.c] && w.text.trim() !== '');
  if (!writes.length) fail_('empty', 'Fill in at least one field.');

  // Validate everything before changing the sheet. The new row gets the
  // dropdowns/checkboxes of the row above, or keeps its own if it is the first.
  const prot = protections_(sheet);
  const isProtected = inserted ? prot.isProtectedInserted : prot.isProtected;
  const rules = prev >= 2
    ? snap.rules[prev - 1]
    : (target <= maxRows ? sheet.getRange(target, 1, 1, snap.lastCol).getDataValidations()[0] : []);
  writes.forEach((w) => {
    if (isProtected(target, w.c + 1)) fail_('protected', '"' + w.name + '" is protected in the sheet and can\'t be filled in here.');
    w.value = cellValue_(w.text, rules[w.c], w.name);
  });

  // Cells of the target row that already hold a (per-row) formula are left alone.
  const skipped = [];
  const hadFormula = (c) => !inserted && target <= snap.lastRow && Boolean(snap.formulas[target - 1][c]);
  writes = writes.filter((w) => {
    if (!hadFormula(w.c)) return true;
    skipped.push(w.c);
    return false;
  });

  // Formulas below the insert point move down one row.
  const anchors = snap.anchors.map((a) => (inserted && a.r >= target - 1 ? { r: a.r + 1, c: a.c, before: a.before } : a));
  const copied = [];
  try {
    if (inserted) sheet.insertRowBefore(target);
    else if (target > maxRows) sheet.insertRowsAfter(maxRows, target - maxRows);

    // Same formatting, dropdowns/checkboxes and per-row formulas as the row above.
    if (prev >= 2) {
      const source = sheet.getRange(prev, 1, 1, snap.lastCol);
      const dest = sheet.getRange(target, 1, 1, snap.lastCol);
      source.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      source.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      snap.auto.forEach((isAuto, c) => {
        if (!isAuto) return;
        sheet.getRange(prev, c + 1).copyTo(sheet.getRange(target, c + 1), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
        if (!hadFormula(c)) copied.push(c);
        if (ARRAY_RE.test(snap.formulas[prev - 1][c])) anchors.push({ r: target - 1, c: c, before: '' });
      });
    }
    return { row: target, skipped: skipped.concat(writeCells_(sheet, target, writes, anchors)) };
  } catch (e) {
    // Leave the sheet as it was: drop the inserted row, or the copied formulas.
    try {
      if (inserted) sheet.deleteRow(target);
      else copied.forEach((c) => sheet.getRange(target, c + 1).clearContent());
      SpreadsheetApp.flush();
    } catch (e2) {
      // best effort
    }
    throw e;
  }
}

function saveView_(ss, req) {
  const sheet = sheet_(ss, req.tabId);
  const input = req.view && typeof req.view === 'object' ? req.view : {};
  const id = validViewId_(input.id) ? String(input.id) : '';
  const views = getViews_(sheet);
  const existing = id ? views.find((v) => v.id === id) : null;
  if (!existing && views.length >= MAX_VIEWS) fail_('bad', 'There are too many views. Delete one first.');

  const view = {
    id: existing ? existing.id : (id || Utilities.getUuid().replace(/-/g, '').slice(0, 10)),
    name: String(input.name || '').trim().slice(0, 60) || 'Untitled view',
    view: cleanList_(input.view),
    edit: cleanList_(input.edit),
    order: existing ? existing.order : views.reduce((m, v) => Math.max(m, Number(v.order) || 0), 0) + 1
  };
  const text = JSON.stringify(view);
  if (utf8Length_(text) > MAX_VIEW_BYTES) fail_('bad', 'This view lists too many columns to save.');
  try {
    viewStore_().setProperty(viewKey_(sheet, view.id), text);
  } catch (e) {
    fail_('bad', 'View storage for this spreadsheet is full. Delete some views first.');
  }
  return { views: getViews_(sheet), id: view.id };
}

function deleteView_(ss, req) {
  const sheet = sheet_(ss, req.tabId);
  if (validViewId_(req.id)) viewStore_().deleteProperty(viewKey_(sheet, String(req.id)));
  return { views: getViews_(sheet) };
}

// ---------------------------------------------------------------- reading the sheet

function visibleSheets_(ss) {
  return ss.getSheets().filter((s) => !s.isSheetHidden() && s.getType() === SpreadsheetApp.SheetType.GRID);
}

function sheet_(ss, tabId) {
  const id = Number(tabId);
  const sheet = tabId === null || tabId === '' || tabId === undefined
    ? null
    : visibleSheets_(ss).find((s) => s.getSheetId() === id);
  if (!sheet) fail_('tab', 'That tab no longer exists or is hidden.');
  return sheet;
}

// Everything the actions need to know about a tab.
function snapshot_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return {
      lastRow: 0, lastCol: 0, display: [], formulas: [], rules: [], headers: [],
      computed: [], auto: [], anchors: [], fillers: [], report: false, lastDataRow: 1
    };
  }
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const snap = {
    lastRow: lastRow,
    lastCol: lastCol,
    display: range.getDisplayValues(),
    formulas: range.getFormulas(),
    rules: range.getDataValidations()
  };
  snap.headers = headerNames_(snap.display[0]);

  // Array formulas ("anchors"). One that is alone in its column (no formula
  // just above or below it) fills that column with its results, so the whole
  // column is read-only. Repeated per-row formulas are not.
  snap.anchors = [];
  for (let r = 0; r < lastRow; r++) {
    for (let c = 0; c < lastCol; c++) {
      const f = snap.formulas[r][c];
      if (f && ARRAY_RE.test(f)) snap.anchors.push({ r: r, c: c, before: snap.display[r][c] });
    }
  }
  // An array formula fills its column when it is in the header row, or when
  // it is alone in its column (no formula just above or below) and heads the
  // column's data (nothing typed above it). A totals formula under typed
  // values, or repeated per-row formulas, do not count.
  const alone = (a) =>
    (a.r === 0 || !snap.formulas[a.r - 1][a.c]) &&
    (a.r + 1 >= lastRow || !snap.formulas[a.r + 1][a.c]);
  const typedAbove = (a) => {
    for (let r = 1; r < a.r; r++) if (!snap.formulas[r][a.c] && !isBlank_(snap, r, a.c)) return true;
    return false;
  };
  snap.fillers = snap.anchors.filter((a) => a.r === 0 || (alone(a) && !typedAbove(a)));
  snap.report = Boolean(snap.formulas[0][0]) && REPORT_RE.test(snap.formulas[0][0]);
  snap.computed = snap.headers.map((_, c) => snap.report || snap.fillers.some((a) => a.c === c));

  // Last row (1-based) holding something a person entered. On a report tab,
  // the last row with anything in it.
  snap.lastDataRow = 1;
  for (let r = lastRow - 1; r >= 1; r--) {
    const filled = snap.display[r].some((_, c) => snap.report
      ? !isBlank_(snap, r, c)
      : !snap.computed[c] && !snap.formulas[r][c] && !isBlank_(snap, r, c));
    if (filled) {
      snap.lastDataRow = r + 1;
      break;
    }
  }

  // Per-row formula columns (the formula is repeated on each row): new rows get it copied.
  const last = snap.lastDataRow - 1;
  snap.auto = snap.headers.map((_, c) => {
    if (snap.computed[c] || last < 1 || !snap.formulas[last][c]) return false;
    if (last === 1) return !ARRAY_RE.test(snap.formulas[last][c]);
    return Boolean(snap.formulas[last - 1][c]);
  });
  return snap;
}

// Empty, or an unticked checkbox nobody has touched.
function isBlank_(snap, r, c) {
  const v = snap.display[r][c];
  if (v === '') return true;
  const rule = snap.rules[r] && snap.rules[r][c];
  if (!rule || snap.formulas[r][c]) return false;
  const cb = checkboxValues_(rule);
  return Boolean(cb) && v.toUpperCase() === cb.unchecked.toUpperCase();
}

// The row a new record would go in (just below the data). Formulas already
// there are fine when they are per-row formulas filled down in advance: the
// same as the row above or below, or referring only to their own row.
// Anything else (a totals row, say) makes the row "foreign".
function targetInfo_(sheet, snap) {
  const target = snap.lastDataRow + 1;
  const none = { foreign: false, prefilled: [] };
  if (!snap.lastCol || target > snap.lastRow || snap.report) return none;
  const ownAnchor = (c) => snap.fillers.some((a) => a.r === target - 1 && a.c === c);
  const cols = [];
  snap.formulas[target - 1].forEach((f, c) => { if (f && !ownAnchor(c)) cols.push(c); });
  if (!cols.length) return none;

  const r1c1 = (row) => sheet.getRange(row, 1, 1, snap.lastCol).getFormulasR1C1()[0];
  const here = r1c1(target);
  const above = snap.lastDataRow >= 2 ? r1c1(snap.lastDataRow) : [];
  const below = target + 1 <= snap.lastRow ? r1c1(target + 1) : [];
  const ownRowOnly = (f) => !/R\[-?[1-9]\d*\]|R\d+/.test(f);
  const ok = (c) => (above[c] && here[c] === above[c]) || (below[c] && here[c] === below[c]) || ownRowOnly(here[c]);
  const foreign = cols.some((c) => !ok(c));
  return { foreign: foreign, prefilled: foreign ? [] : cols };
}

function headerNames_(row) {
  const seen = new Set();
  return row.map((v, c) => {
    const base = String(v).trim() || 'Column ' + colLetter_(c);
    let name = base;
    let n = 2;
    while (seen.has(name)) name = base + ' (' + n++ + ')';
    seen.add(name);
    return name;
  });
}

function checkHeaders_(snap, headers) {
  const same = Array.isArray(headers) &&
    headers.length === snap.headers.length &&
    headers.every((h, i) => String(h) === snap.headers[i]);
  if (!same) fail_('columns', "The sheet's columns were changed. Loading the latest version, please try again.");
}

// Only cells people type into are compared; formula results change on their own.
function rowMatches_(current, formulas, computed, expected) {
  if (!Array.isArray(expected)) return false;
  for (let c = 0; c < current.length; c++) {
    if (formulas[c] || computed[c]) continue;
    const want = expected[c] === null || expected[c] === undefined ? '' : String(expected[c]);
    if (String(current[c]) !== want) return false;
  }
  return true;
}

// Input hints for the phone form: dropdowns, checkboxes, read-only columns.
function columnMeta_(sheet, snap) {
  if (!snap.lastCol) return [];
  const prot = protections_(sheet);
  const template = Math.max(2, snap.lastDataRow);
  let rules = null;
  if (template <= snap.lastRow) rules = snap.rules[template - 1];
  else if (template <= sheet.getMaxRows()) rules = sheet.getRange(template, 1, 1, snap.lastCol).getDataValidations()[0];

  const place = targetInfo_(sheet, snap);
  const next = snap.lastDataRow + 1;
  return snap.headers.map((_, c) => {
    const meta = {
      type: 'text',
      computed: snap.computed[c],
      auto: snap.auto[c] || place.prefilled.indexOf(c) >= 0,
      locked: prot.isProtected(2, c + 1) &&
        (place.foreign ? prot.isProtectedInserted(next, c + 1) : prot.isProtected(next, c + 1))
    };
    const rule = rules && rules[c];
    if (rule) Object.assign(meta, ruleMeta_(rule));
    return meta;
  });
}

function checkboxValues_(rule) {
  if (rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) return null;
  const args = rule.getCriteriaValues() || [];
  return args.length
    ? { checked: String(args[0]), unchecked: args.length > 1 ? String(args[1]) : '', custom: true }
    : { checked: 'TRUE', unchecked: 'FALSE', custom: false };
}

// All options of a dropdown rule, or null when the rule is not a dropdown.
function ruleOptions_(rule) {
  const T = SpreadsheetApp.DataValidationCriteria;
  const type = rule.getCriteriaType();
  if (type !== T.VALUE_IN_LIST && type !== T.VALUE_IN_RANGE) return null;
  const args = rule.getCriteriaValues() || [];
  let options = [];
  try {
    options = type === T.VALUE_IN_LIST
      ? (args[0] || [])
      : args[0].getDisplayValues().reduce((all, r) => all.concat(r), []);
  } catch (e) {
    options = [];
  }
  return unique_(options.map(String).filter((v) => v !== ''));
}

function ruleMeta_(rule) {
  const cb = checkboxValues_(rule);
  if (cb) return { type: 'checkbox', checked: cb.checked, unchecked: cb.unchecked };
  const options = ruleOptions_(rule);
  if (options && options.length) {
    return {
      type: 'choice',
      options: options.slice(0, 500),
      truncated: options.length > 500,
      strict: !rule.getAllowInvalid()
    };
  }
  return {};
}

function protections_(sheet) {
  const P = SpreadsheetApp.ProtectionType;
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  // A range reaching the last row/column of the grid also covers rows/columns added later.
  const box = (rg) => {
    const r2 = rg.getLastRow();
    const c2 = rg.getLastColumn();
    return { r1: rg.getRow(), r2: r2 >= maxRows ? Infinity : r2, c1: rg.getColumn(), c2: c2 >= maxCols ? Infinity : c2 };
  };
  const inside = (b, row, col) => row >= b.r1 && row <= b.r2 && col >= b.c1 && col <= b.c2;
  const ranges = sheet.getProtections(P.RANGE).filter((p) => !p.isWarningOnly()).map((p) => box(p.getRange()));
  const whole = sheet.getProtections(P.SHEET).filter((p) => !p.isWarningOnly());
  const open = whole.length ? whole[0].getUnprotectedRanges().map(box) : [];
  const isProtected = (row, col) =>
    ranges.some((b) => inside(b, row, col)) || (whole.length > 0 && !open.some((b) => inside(b, row, col)));
  // A row inserted above `row` is covered only by ranges that span the gap.
  const spans = (b, row, col) => b.r1 < row && b.r2 >= row && col >= b.c1 && col <= b.c2;
  const isProtectedInserted = (row, col) =>
    ranges.some((b) => spans(b, row, col)) || (whole.length > 0 && !open.some((b) => spans(b, row, col)));
  return { isProtected: isProtected, isProtectedInserted: isProtectedInserted };
}

// ---------------------------------------------------------------- writing

// Request maps of {columnIndex: value}; keys must be plain column numbers.
function entries_(obj, snap) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return [];
  return Object.keys(obj).map((k) => {
    if (!/^(0|[1-9]\d*)$/.test(k) || Number(k) >= snap.lastCol) fail_('bad', 'Unknown column.');
    const v = obj[k];
    const c = Number(k);
    return { c: c, name: snap.headers[c], text: v === null || v === undefined ? '' : String(v) };
  });
}

// What to store for text typed in the app: { value, literal }.
// literal = store exactly this text (never a formula, number or date).
function cellValue_(text, rule, name) {
  if (rule) {
    const cb = checkboxValues_(rule);
    if (cb && !cb.custom) {
      const up = text.trim().toUpperCase();
      if (up === 'TRUE') return { value: true };
      if (up === 'FALSE' || up === '') return { value: false };
    }
    const options = cb ? null : ruleOptions_(rule);
    if (options && options.length && !rule.getAllowInvalid() && text !== '' && options.indexOf(text) < 0) {
      fail_('invalid', '"' + name + '" only accepts one of its dropdown options.');
    }
  }
  if (text === '') return { value: '' };

  const t = text.trim();
  const literal =
    /^[=+@']/.test(t) ||                              // would become a formula (or lose the apostrophe)
    (t.charAt(0) === '-' && !/^-[\d.,]*\d[\d.,]*%?$/.test(t)) ||
    /^0\d+$/.test(t) ||                               // phone numbers / IDs keep their leading zero
    /^\d{16,}$/.test(t);                              // long ID numbers keep every digit
  return { value: text, literal: literal };
}

// Writes the values into one row. A value that would make an array formula
// show #REF! (because it lands where the formula puts its results, or changes
// what the formula returns) is left out and returned as skipped, and the cell
// keeps what it had. If the sheet rejects a value, every cell is put back.
function writeCells_(sheet, row, writes, anchors) {
  if (!writes.length) return [];
  const cell = (w) => sheet.getRange(row, w.c + 1);
  const put = (w, value) => {
    if (value.literal) cell(w).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(String(value.value)).build());
    else cell(w).setValue(value.value);
  };
  const lastWritten = writes.reduce((m, w) => Math.max(m, w.c), 0);
  const before = sheet.getRange(row, 1, 1, lastWritten + 1).getValues()[0];
  // Puts a cell's old value back exactly (text stays text). If the old value
  // breaks the cell's own validation rule (rule added later), the rule is
  // lifted for the moment it takes to put it back.
  const restore = (w) => {
    const v = before[w.c];
    const value = typeof v === 'string' ? (v === '' ? { value: '' } : { value: v, literal: true }) : { value: v };
    try {
      put(w, value);
    } catch (e) {
      const rule = cell(w).getDataValidation();
      cell(w).clearDataValidations();
      put(w, value);
      if (rule) cell(w).setDataValidation(rule);
    }
  };
  // Clears the cells and reports which ones a formula fills on its own
  // (array output: empty underneath, so they must stay empty). Only cells
  // below/right of an array formula can be its output.
  const clearAndFindOutput = (list) => {
    list.forEach((w) => cell(w).setValue(''));
    SpreadsheetApp.flush();
    const shown = sheet.getRange(row, 1, 1, lastWritten + 1).getDisplayValues()[0];
    return new Set(list.filter((w) => shown[w.c] !== '').map((w) => w.c));
  };
  const mayBeOutput = (w) => anchors.some((a) => a.r <= row - 1 && a.c <= w.c && !(a.r === row - 1 && a.c === w.c));

  const done = [];
  try {
    writes.forEach((w) => { put(w, w.value); done.push(w); });
    SpreadsheetApp.flush();
  } catch (e) {
    const notRestored = [];
    let output = new Set();
    try { output = clearAndFindOutput(done.filter(mayBeOutput)); } catch (e2) { /* keep going */ }
    done.forEach((w) => {
      if (output.has(w.c)) return;
      try { restore(w); } catch (e3) { notRestored.push(w.name); }
    });
    try { SpreadsheetApp.flush(); } catch (e4) { /* keep going */ }
    const failed = writes[Math.min(done.length, writes.length - 1)];
    const reason = '"' + failed.name + '": ' + (e && e.message ? e.message : String(e));
    fail_('invalid', notRestored.length
      ? reason + ' These fields may still hold the new value: ' + notRestored.join(', ') + '.'
      : 'Nothing was saved. ' + reason);
  }

  if (!anchors.length || !newlyBroken_(sheet, anchors)) return [];

  // Is it our values? With all of them cleared, the formulas should recover.
  const output = clearAndFindOutput(writes);
  if (newlyBroken_(sheet, anchors)) {
    // Not caused by these values; keep them.
    writes.forEach((w) => put(w, w.value));
    SpreadsheetApp.flush();
    return [];
  }

  // Put the values back one at a time; leave out the ones that break a formula.
  const skipped = [];
  writes.forEach((w) => {
    put(w, w.value);
    SpreadsheetApp.flush();
    if (!newlyBroken_(sheet, anchors)) return;
    try {
      if (output.has(w.c)) cell(w).setValue('');
      else restore(w);
      SpreadsheetApp.flush();
      skipped.push(w.c);
    } catch (e) {
      put(w, w.value); // the old value can't go back; keep the new one
      SpreadsheetApp.flush();
    }
  });
  return skipped;
}

function newlyBroken_(sheet, anchors) {
  let r1 = Infinity, r2 = -1, c1 = Infinity, c2 = -1;
  anchors.forEach((a) => {
    r1 = Math.min(r1, a.r); r2 = Math.max(r2, a.r);
    c1 = Math.min(c1, a.c); c2 = Math.max(c2, a.c);
  });
  const now = sheet.getRange(r1 + 1, c1 + 1, r2 - r1 + 1, c2 - c1 + 1).getDisplayValues();
  return anchors.some((a) => now[a.r - r1][a.c - c1] === '#REF!' && a.before !== '#REF!');
}

// ---------------------------------------------------------------- views storage
// One document property per view, so each view gets the full 9 KB value limit.

function viewStore_() {
  return PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
}

function viewKey_(sheet, id) {
  return VIEW_PREFIX + sheet.getSheetId() + ':' + id;
}

function validViewId_(id) {
  return typeof id === 'string' && /^[A-Za-z0-9_-]{1,40}$/.test(id);
}

function getViews_(sheet) {
  const prefix = VIEW_PREFIX + sheet.getSheetId() + ':';
  const all = viewStore_().getProperties();
  const views = [];
  Object.keys(all).forEach((k) => {
    if (k.indexOf(prefix) !== 0) return;
    try {
      const v = JSON.parse(all[k]);
      if (v && validViewId_(v.id)) views.push(v);
    } catch (e) {
      // ignore unreadable entries
    }
  });
  views.sort((a, b) => (Number(a.order) || 0) - (Number(b.order) || 0) || String(a.name).localeCompare(String(b.name)));
  return views;
}

function cleanList_(list) {
  return Array.isArray(list) ? unique_(list.map((x) => String(x).slice(0, 200))).slice(0, 300) : [];
}

// ---------------------------------------------------------------- utils

function unique_(list) {
  return Array.from(new Set(list));
}

function utf8Length_(s) {
  let n = 0;
  for (const ch of s) {
    const cp = ch.codePointAt(0);
    n += cp < 0x80 ? 1 : cp < 0x800 ? 2 : cp < 0x10000 ? 3 : 4;
  }
  return n;
}

function colLetter_(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function fail_(code, message) {
  const err = new Error(message);
  err.code = code;
  throw err;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
