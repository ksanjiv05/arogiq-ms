(() => {
  'use strict';

  const CONFIG = window.SHEET_FORMS_CONFIG || {};
  const API_URL = String(CONFIG.apiUrl || '').trim();
  const PAGE = 100;                 // rows rendered per "Show more" step
  const STALE_MS = 2 * 60 * 1000;   // reload data when coming back to the app after this long

  const $ = (id) => document.getElementById(id);

  // ---------- device storage (only conveniences: access key, last tab/view, layout) ----------
  function hash(s) {
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
    return (h >>> 0).toString(36);
  }
  const NS = `sf.${hash(API_URL)}.`;
  const store = {
    get(k) {
      try { const v = localStorage.getItem(NS + k); return v === null ? null : JSON.parse(v); } catch { return null; }
    },
    set(k, v) {
      try { localStorage.setItem(NS + k, JSON.stringify(v)); } catch { /* storage unavailable */ }
    },
    del(k) {
      try { localStorage.removeItem(NS + k); } catch { /* storage unavailable */ }
    }
  };

  // ---------- helpers ----------
  function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function option(value, text) {
    const o = el('option', null, text);
    o.value = String(value);
    return o;
  }

  function linkButton(text, onClick) {
    const b = el('button', 'linkbtn', text);
    b.type = 'button';
    b.addEventListener('click', onClick);
    return b;
  }

  const pad = (cells, width) => Array.from({ length: width }, (_, c) => String((cells && cells[c]) ?? ''));
  const byNumber = (a, b) => a - b;

  // Toasts and the progress bar must sit inside an open modal dialog, or the
  // dialog (which is in the browser's top layer) covers them.
  function host(node) {
    const target = document.querySelector('dialog[open]') || document.body;
    if (node.parentNode !== target) target.append(node);
    node.classList.toggle('in-dialog', target !== document.body);
  }

  function hostOverlays() {
    host($('toast'));
    host($('progress'));
  }

  let toastTimer = null;
  function toast(message, isError = false) {
    const t = $('toast');
    host(t);
    t.textContent = message;
    t.classList.toggle('error', isError);
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, isError ? 7000 : 2500);
  }

  let busyCount = 0;
  function busy(on) {
    busyCount = Math.max(0, busyCount + (on ? 1 : -1));
    const p = $('progress');
    host(p);
    p.hidden = busyCount === 0;
  }

  // ---------- state ----------
  const state = {
    key: store.get('key') || '',
    title: '',
    tabs: [],
    tabId: null,
    loadedTab: null,   // tab whose rows are in state.rows
    loadFailed: false,
    canManage: false,
    headers: [],
    columns: [],       // per column: { type, options, strict, checked, unchecked, computed, auto, locked }
    rows: [],          // { n: sheet row number, cells: string[], f: Set of formula column indexes }
    stats: [],
    views: [],         // saved views for this tab, shared by everyone
    viewId: null,
    current: { view: [], edit: [] },  // column indexes in use right now
    custom: false,     // columns changed on this device and not saved
    query: '',
    shown: PAGE,
    layout: store.get('layout') || (window.matchMedia('(min-width: 720px)').matches ? 'table' : 'cards'),
    loadedAt: 0,
    writeSeq: 0,       // bumped whenever a save result arrives
    viewSaving: false,
    form: null,
    draft: null,
    lastError: ''
  };

  // ---------- API ----------
  class ApiError extends Error {
    constructor(message, code) { super(message); this.code = code; }
  }

  async function call(action, params = {}) {
    let res;
    try {
      // text/plain keeps this a "simple" request, which Apps Script web apps accept cross-origin.
      res = await fetch(API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ ...params, action, key: state.key })
      });
    } catch {
      throw new ApiError('Could not reach the sheet. Check your internet connection. If it keeps failing, the Apps Script must be deployed with "Who has access: Anyone".', 'network');
    }
    let data = null;
    try { data = await res.json(); } catch { /* not JSON */ }
    if (!data || typeof data.ok !== 'boolean') {
      throw new ApiError(`The sheet service sent an unexpected response (${res.status}). Check the Apps Script URL in js/config.js.`, 'bad-response');
    }
    if (!data.ok) throw new ApiError(data.error || 'Request failed.', data.code || 'error');
    return data.data;
  }

  async function run(task) {
    busy(true);
    try { await task(); return true; }
    catch (err) { handleError(err); return false; }
    finally { busy(false); }
  }

  function handleError(err) {
    const code = err && err.code;
    state.lastError = (err && err.message) || 'Something went wrong.';
    if (code === 'auth') {
      closeDialogs();
      showKey(state.lastError);
      return;
    }
    toast(state.lastError, true);
    if (code === 'columns' || code === 'tab') {
      closeDialogs();
      setTimeout(() => (code === 'tab' ? run(loadInfo) : reloadTab()), 0);
    }
  }

  // ---------- loading ----------
  async function loadInfo() {
    const info = await call('info');
    state.title = info.title || 'Sheet Forms';
    state.tabs = Array.isArray(info.tabs) ? info.tabs : [];
    state.canManage = Boolean(info.canManageViews);
    $('docTitle').textContent = state.title;
    document.title = state.title;

    if (!state.tabs.length) {
      showMessage(state.title, ['There are no visible tabs in this sheet.']);
      return;
    }
    const ids = state.tabs.map((t) => t.id);
    const remembered = store.get('tab');
    if (!ids.includes(state.tabId)) {
      state.tabId = ids.includes(remembered) ? remembered : ids[0];
      state.viewId = null;
      state.custom = false;
    }
    const select = $('tabSelect');
    select.replaceChildren(...state.tabs.map((t) => option(t.id, t.name)));
    select.value = String(state.tabId);
    select.hidden = state.tabs.length < 2;
    $('pickers').hidden = select.hidden && $('viewSelect').hidden;

    showMain();
    renderFoot();
    await loadTab();
  }

  async function loadTab() {
    const tabId = state.tabId;
    const seq = state.writeSeq;
    state.loadFailed = false;
    if (state.loadedTab !== tabId) renderList(); // shows "Loading…"
    const data = await call('read', { tabId });
    if (tabId !== state.tabId) return; // the user switched tabs meanwhile
    if (seq !== state.writeSeq) return loadTab(); // a save landed meanwhile; read again

    const headers = Array.isArray(data.headers) ? data.headers.map(String) : [];
    const width = headers.length;
    const sameColumns = state.loadedTab === tabId && headers.join('\u0000') === state.headers.join('\u0000');

    state.headers = headers;
    state.columns = Array.from({ length: width }, (_, c) => (data.columns && data.columns[c]) || { type: 'text' });
    state.rows = (data.rows || []).map((r) => ({ n: r.n, cells: pad(r.cells, width), f: new Set(r.f || []) }));
    state.views = Array.isArray(data.views) ? data.views : [];
    if (typeof data.canManageViews === 'boolean') state.canManage = data.canManageViews;
    state.loadedTab = tabId;
    state.loadedAt = Date.now();

    computeStats();
    selectView(sameColumns && state.custom);
    renderList();
    renderFoot();
  }

  async function reloadTab() {
    const tabId = state.tabId;
    const ok = await run(async () => {
      try {
        await loadTab();
      } catch (err) {
        if (tabId === state.tabId) throw err; // ignore failures for a tab the user already left
      }
    });
    if (!ok && tabId === state.tabId && state.loadedTab !== state.tabId) {
      state.loadFailed = true;
      renderList();
    }
  }

  // Per-column hints for a comfortable phone keyboard and suggestions.
  function computeStats() {
    state.stats = state.headers.map((_, c) => {
      let filled = 0, numeric = 0, negative = false, long = false;
      const distinct = new Set();
      for (const row of state.rows) {
        const v = row.cells[c];
        if (!v) continue;
        filled++;
        if (/^-?[\d,]*\.?\d+$/.test(v)) numeric++;
        if (v.startsWith('-')) negative = true;
        if (v.length > 60 || v.includes('\n')) long = true;
        if (distinct.size <= 15) distinct.add(v);
      }
      const isChoice = distinct.size >= 1 && distinct.size <= 15 && filled >= distinct.size * 2;
      return {
        // iPhone's decimal keypad has no minus key, so only use it for columns without negatives.
        numericPad: filled > 0 && numeric / filled >= 0.8 && !negative,
        long,
        options: isChoice ? [...distinct].sort() : []
      };
    });
  }

  // ---------- views ----------
  function editable(c) {
    const m = state.columns[c];
    return !(m && (m.computed || m.locked));
  }

  function defaultView() {
    return {
      id: '',
      name: 'Default',
      view: state.headers.slice(0, 4),
      edit: state.headers.filter((_, c) => editable(c))
    };
  }

  const allViews = () => (state.views.length ? state.views : [defaultView()]);
  const activeView = () => allViews().find((v) => v.id === state.viewId) || allViews()[0];

  function toIndexes(names) {
    const out = new Set();
    for (const name of names || []) {
      const i = state.headers.indexOf(String(name));
      if (i >= 0) out.add(i);
    }
    return [...out].sort(byNumber);
  }

  function applyView(view) {
    state.viewId = view.id;
    state.current = { view: toIndexes(view.view), edit: toIndexes(view.edit).filter(editable) };
    state.custom = false;
  }

  function selectView(keepCustom) {
    const ids = allViews().map((v) => v.id);
    if (!ids.includes(state.viewId)) {
      const remembered = store.get(`view.${state.tabId}`);
      state.viewId = ids.includes(remembered) ? remembered : ids[0];
      keepCustom = false;
    }
    if (!keepCustom) applyView(activeView());
    renderViewSelect();
  }

  function renderViewSelect() {
    const select = $('viewSelect');
    const views = allViews();
    select.replaceChildren(...views.map((v) =>
      option(v.id, v.id === state.viewId && state.custom ? `${v.name} (changed)` : v.name)));
    select.value = String(state.viewId);
    select.hidden = views.length < 2;
    $('pickers').hidden = $('tabSelect').hidden && select.hidden;
  }

  // ---------- list ----------
  function matchingRows() {
    const q = state.query.trim().toLowerCase();
    if (!q) return state.rows;
    return state.rows.filter((row) => row.cells.some((v) => v.toLowerCase().includes(q)));
  }

  function setListButtons() {
    const ready = state.loadedTab === state.tabId;
    $('addBtn').disabled = !ready;
    $('columnsBtn').disabled = !ready;
  }

  function renderList() {
    const list = $('list');
    $('layoutBtn').textContent = state.layout === 'cards' ? 'Table' : 'Cards';
    setListButtons();

    if (state.loadedTab !== state.tabId) {
      if (state.loadFailed) {
        const retry = el('button', 'btn primary', 'Try again');
        retry.type = 'button';
        retry.addEventListener('click', reloadTab);
        const box = el('div', 'empty');
        box.append(el('p', null, `Couldn't load this tab. ${state.lastError}`), retry);
        list.replaceChildren(box);
      } else {
        list.replaceChildren(el('p', 'empty', 'Loading…'));
      }
      return;
    }
    if (!state.headers.length) {
      list.replaceChildren(el('p', 'empty', 'This tab is empty. Put column names in row 1 of the sheet, then tap Refresh.'));
      return;
    }

    const notes = [];
    if (state.custom) {
      const note = el('p', 'count');
      note.append(document.createTextNode('Columns changed on this device. '), linkButton('Reset', () => {
        applyView(activeView());
        renderViewSelect();
        renderList();
      }));
      notes.push(note);
    }

    if (!state.current.view.length) {
      list.replaceChildren(...notes, el('p', 'empty', 'No columns are shown. Tap "Columns" to choose what to show.'));
      return;
    }
    const rows = matchingRows();
    if (!rows.length) {
      list.replaceChildren(...notes, el('p', 'empty', state.query ? 'No rows match your search.' : 'No rows yet. Tap "+ Add row" to create one.'));
      return;
    }

    const visible = rows.slice(0, state.shown);
    const nodes = [...notes, el('p', 'count', `${rows.length} row${rows.length === 1 ? '' : 's'} · tap a row to edit`)];
    nodes.push(state.layout === 'cards' ? renderCards(visible) : renderTable(visible));
    if (rows.length > visible.length) {
      const more = el('button', 'btn block', `Show more (${rows.length - visible.length} more)`);
      more.type = 'button';
      more.addEventListener('click', () => { state.shown += PAGE; renderList(); });
      nodes.push(more);
    }
    list.replaceChildren(...nodes);
  }

  function renderCards(rows) {
    const frag = document.createDocumentFragment();
    const [first, ...rest] = state.current.view;
    for (const row of rows) {
      const card = el('button', 'card');
      card.type = 'button';
      const title = el('div', 'card-title');
      title.append(el('span', 'card-name', row.cells[first] || '(blank)'), el('span', 'rownum', `Row ${row.n}`));
      card.append(title);
      if (rest.length) {
        const dl = el('dl');
        for (const c of rest) dl.append(el('dt', null, state.headers[c]), el('dd', null, row.cells[c]));
        card.append(dl);
      }
      card.addEventListener('click', () => openEdit(row));
      frag.append(card);
    }
    return frag;
  }

  function renderTable(rows) {
    const wrap = el('div', 'table-wrap');
    const table = el('table');
    const head = el('tr');
    head.append(el('th', null, '#'));
    for (const c of state.current.view) head.append(el('th', null, state.headers[c]));
    const thead = el('thead');
    thead.append(head);
    const tbody = el('tbody');
    for (const row of rows) {
      const tr = el('tr');
      tr.tabIndex = 0;
      tr.append(el('td', 'muted', String(row.n)));
      for (const c of state.current.view) tr.append(el('td', null, row.cells[c]));
      tr.addEventListener('click', () => openEdit(row));
      tr.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter') return;
        e.preventDefault(); // otherwise the same key press also submits the form that opens
        openEdit(row);
      });
      tbody.append(tr);
    }
    table.append(thead, tbody);
    wrap.append(table);
    return wrap;
  }

  function renderFoot() {
    const parts = [];
    if (state.canManage) parts.push(el('span', null, 'You can manage views.'));
    parts.push(state.key
      ? linkButton('Forget access key on this device', () => {
        state.key = '';
        store.del('key');
        toast('Access key removed from this device.');
        start();
      })
      : linkButton('Enter an access key', () => showKey('Enter your access key.')));
    $('foot').replaceChildren(...parts);
  }

  // ---------- column chooser ----------
  function openColumns() {
    if (state.loadedTab !== state.tabId) return;
    if (!state.headers.length) {
      toast('This tab has no columns yet.', true);
      return;
    }
    const v = activeView();
    const admin = state.canManage;
    state.draft = { view: new Set(state.current.view), edit: new Set(state.current.edit) };
    $('viewNameField').hidden = !admin;
    $('viewName').value = v.name;
    $('viewAdmin').hidden = !admin;
    $('deleteView').hidden = !v.id;
    $('columnsApply').textContent = admin ? 'Save' : 'Apply';
    $('columnsNote').textContent = admin
      ? (v.id ? 'Saving updates this view for everyone who uses the app.' : 'Saving makes this the default view for everyone who uses the app.')
      : 'Your changes apply on this device until the page is reloaded. Only an admin can save views for everyone.';
    renderColumnsList();
    setColumnsBusy(state.viewSaving);
    $('columnsDlg').showModal();
    hostOverlays();
  }

  function renderColumnsList() {
    const draft = state.draft;
    $('columnsList').replaceChildren(...state.headers.map((name, c) => {
      const row = el('div', 'col-row');
      const label = el('span', 'name', name);
      const meta = state.columns[c] || {};
      if (meta.computed) label.append(el('span', 'tag', 'formula'));
      else if (meta.locked) label.append(el('span', 'tag', 'protected'));
      row.append(label);
      for (const kind of ['view', 'edit']) {
        const wrap = el('label');
        const check = el('input');
        check.type = 'checkbox';
        check.checked = draft[kind].has(c);
        if (kind === 'edit' && !editable(c)) check.disabled = true;
        check.addEventListener('change', () => {
          if (check.checked) draft[kind].add(c); else draft[kind].delete(c);
        });
        wrap.append(check, document.createTextNode(kind === 'view' ? 'Show' : 'Edit'));
        row.append(wrap);
      }
      return row;
    }));
  }

  function setColumnsBusy(on) {
    for (const id of ['columnsApply', 'saveAsNew', 'deleteView']) $(id).disabled = on;
  }

  function draftColumns() {
    return {
      view: [...state.draft.view].sort(byNumber),
      edit: [...state.draft.edit].filter(editable).sort(byNumber)
    };
  }

  async function saveView(def) {
    setColumnsBusy(true);
    state.viewSaving = true;
    const tabId = state.tabId;
    const draft = state.draft;
    await run(async () => {
      const res = await call('saveView', {
        tabId,
        view: {
          id: def.id || '',
          name: def.name,
          view: def.view.map((c) => state.headers[c]),
          edit: def.edit.map((c) => state.headers[c])
        }
      });
      if (tabId !== state.tabId) return;
      state.views = res.views || [];
      state.viewId = res.id;
      store.set(`view.${tabId}`, res.id);
      if (state.draft === draft) {
        applyView(activeView());
        $('columnsDlg').close();
      }
      renderViewSelect();
      renderList();
      toast('View saved for everyone.');
    });
    state.viewSaving = false;
    setColumnsBusy(false);
  }

  // ---------- forms ----------
  function buildField(c, value, lockReason) {
    const meta = state.columns[c] || {};
    const stat = state.stats[c] || {};
    const id = `field-${c}`;
    const wrap = el('div', 'field');
    const label = el('label', 'field-label', state.headers[c]);
    label.htmlFor = id;
    let input;
    let get;

    if (meta.type === 'checkbox' && !lockReason) {
      const on = meta.checked || 'TRUE';
      const off = meta.unchecked ?? 'FALSE';
      wrap.classList.add('check');
      input = el('input', 'toggle');
      input.type = 'checkbox';
      input.checked = value.toUpperCase() === on.toUpperCase();
      get = () => (input.checked ? on : off);
      wrap.append(label, input);
    } else if (meta.type === 'choice' && meta.strict && !meta.truncated && !lockReason) {
      input = el('select');
      input.append(option('', '- Select -'));
      for (const o of meta.options) input.append(option(o, o));
      if (value && !meta.options.includes(value)) input.append(option(value, value));
      input.value = value;
      get = () => input.value;
      wrap.append(label, input);
    } else {
      const multiline = stat.long || value.length > 60 || value.includes('\n');
      input = el(multiline ? 'textarea' : 'input');
      if (multiline) {
        input.rows = 3;
      } else {
        input.type = 'text';
        input.autocomplete = 'off';
        if (stat.numericPad) input.inputMode = 'decimal';
      }
      input.value = value;
      get = () => input.value;
      wrap.append(label, input);

      const options = meta.type === 'choice' ? meta.options : stat.options;
      if (!multiline && !lockReason && options && options.length) {
        const list = el('datalist');
        list.id = `options-${c}`;
        for (const o of options) list.append(option(o, o));
        input.setAttribute('list', list.id);
        wrap.append(list);
      }
    }
    input.id = id;

    if (lockReason) {
      input.disabled = true;
      wrap.append(el('span', 'hint', lockReason));
    }
    return { c, input, get, orig: get(), locked: Boolean(lockReason), wrap };
  }

  function openRowDialog(title, saveText, fields) {
    $('rowTitle').textContent = title;
    $('rowSave').textContent = saveText;
    $('rowSave').dataset.label = saveText;
    $('rowSave').disabled = false;
    // On phones the keyboard's Enter key moves to the next field instead of submitting.
    const inputs = (fields || []).filter((f) => !f.locked && f.input.tagName === 'INPUT' && f.input.type === 'text');
    inputs.forEach((f, i) => { f.input.enterKeyHint = i === inputs.length - 1 ? 'done' : 'next'; });
    $('rowDlg').showModal();
    hostOverlays();
    $('rowBody').scrollTop = 0;
  }

  function openEdit(row) {
    if (state.loadedTab !== state.tabId) return;
    const body = $('rowBody');
    if (!state.current.edit.length) {
      const pick = el('button', 'btn primary block', 'Choose columns to edit');
      pick.type = 'button';
      pick.addEventListener('click', () => { $('rowDlg').close(); openColumns(); });
      body.replaceChildren(el('p', 'empty', 'No columns are selected for editing. Pick the columns you want in this form.'), pick);
      state.form = null;
      $('rowSave').hidden = true;
      openRowDialog(`Row ${row.n}`, 'Save', []);
      return;
    }

    const context = state.current.view
      .filter((c) => !state.current.edit.includes(c) && row.cells[c])
      .map((c) => `${state.headers[c]}: ${row.cells[c]}`);

    const fields = state.current.edit.map((c) =>
      buildField(c, row.cells[c], row.f.has(c) ? 'Calculated by a formula. Change it in Google Sheets.' : ''));

    const nodes = fields.map((f) => f.wrap);
    if (context.length) nodes.unshift(el('div', 'context', context.join('\n')));
    body.replaceChildren(...nodes);
    state.form = {
      mode: 'edit',
      tabId: state.loadedTab,
      headers: state.headers.slice(),
      row: { n: row.n, cells: row.cells.slice(), f: new Set(row.f) },
      computed: state.columns.map((m) => Boolean(m && m.computed)),
      fields
    };
    $('rowSave').hidden = false;
    openRowDialog(`Edit row ${row.n}`, 'Save', fields);
  }

  function openAdd() {
    if (state.loadedTab !== state.tabId) return;
    if (!state.headers.length) {
      toast('This tab has no header row. Add column names in row 1 first.', true);
      return;
    }
    const auto = [];
    const locked = [];
    const fields = [];
    state.headers.forEach((name, c) => {
      const meta = state.columns[c] || {};
      if (meta.computed || meta.auto) auto.push(name);
      else if (meta.locked) locked.push(name);
      else fields.push(buildField(c, '', ''));
    });
    if (!fields.length) {
      toast('No column in this tab can be filled in from the app.', true);
      return;
    }
    const nodes = fields.map((f) => f.wrap);
    if (auto.length) nodes.push(el('p', 'muted small', `Filled in automatically: ${auto.join(', ')}`));
    if (locked.length) nodes.push(el('p', 'muted small', `Protected in the sheet: ${locked.join(', ')}`));
    $('rowBody').replaceChildren(...nodes);
    state.form = { mode: 'add', tabId: state.loadedTab, headers: state.headers.slice(), fields };
    $('rowSave').hidden = false;
    openRowDialog('Add new row', 'Add row', fields);
  }

  function changedValues(form) {
    const out = {};
    for (const f of form.fields) {
      if (f.locked) continue;
      const v = f.get();
      if (v !== f.orig) out[f.c] = v;
    }
    return out;
  }

  const isDirty = (form) => Boolean(form) && Object.keys(changedValues(form)).length > 0;

  // Closes the row dialog only if it still shows this form.
  function finishForm(form) {
    if (state.form === form && $('rowDlg').open) $('rowDlg').close();
  }

  function warnSkipped(form, skipped, doneText) {
    if (!skipped || !skipped.length) {
      toast(doneText);
      return;
    }
    const names = skipped.map((c) => form.headers[c]).join(', ');
    toast(`${doneText} Not saved: ${names}. Saving ${skipped.length === 1 ? 'it' : 'them'} would break a formula in the sheet.`, true);
  }

  function conflictText(form, now, changes) {
    const h = form.headers;
    const quote = (v) => (v === '' ? '(blank)' : `"${v}"`);
    const theirs = h.map((_, c) => c)
      .filter((c) => !form.row.f.has(c) && !form.computed[c] && now[c] !== form.row.cells[c])
      .slice(0, 8)
      .map((c) => `• ${h[c]}: ${quote(form.row.cells[c])} → ${quote(now[c])}${c in changes ? '  (you changed this too)' : ''}`);
    const mine = Object.keys(changes).map((c) => `• ${h[c]}: ${quote(changes[c])}`);
    return `Someone else changed row ${form.row.n} after you opened it.\n\n` +
      (theirs.length ? `Their changes:\n${theirs.join('\n')}\n\n` : '') +
      `Your changes:\n${mine.join('\n')}\n\n` +
      'OK: save your changes anyway (only the fields you changed).\n' +
      'Cancel: discard your changes and load the latest data.';
  }

  async function saveEdit(form, force = false) {
    const changes = changedValues(form);
    if (!Object.keys(changes).length) {
      finishForm(form);
      toast('No changes to save.');
      return;
    }
    const res = await call('update', {
      tabId: form.tabId,
      row: form.row.n,
      headers: form.headers,
      expected: form.row.cells,
      changes,
      force
    });
    if (!res.conflict && !res.missing) state.writeSeq++;

    if (res.missing) {
      finishForm(form);
      toast(`Row ${form.row.n} no longer exists in the sheet. Showing the latest data.`, true);
      await loadTab();
      return;
    }
    if (res.conflict) {
      const now = pad(res.current, form.headers.length);
      if (window.confirm(conflictText(form, now, changes))) {
        await saveEdit(form, true);
      } else {
        finishForm(form);
        await loadTab();
      }
      return;
    }

    if (state.loadedTab === form.tabId) {
      const row = state.rows.find((r) => r.n === res.row);
      if (row) {
        row.cells = pad(res.cells, state.headers.length);
        row.f = new Set(res.f || []);
      }
      renderList();
    }
    finishForm(form);
    warnSkipped(form, res.skipped, `Row ${form.row.n} saved.`);
  }

  async function saveAdd(form) {
    const values = changedValues(form);
    if (!Object.values(values).some((v) => String(v).trim() !== '')) {
      toast('Fill in at least one field.', true);
      return;
    }
    const res = await call('append', { tabId: form.tabId, headers: form.headers, values });
    state.writeSeq++;
    finishForm(form);
    warnSkipped(form, res.skipped, `Row ${res.row} added.`);
    if (state.tabId === form.tabId && !(await run(loadTab))) {
      toast(`Row ${res.row} was added, but the list could not be refreshed. Tap Refresh.`, true);
    }
  }

  function confirmDiscard() {
    const form = state.form;
    if (form && form.saving) {
      toast('Still saving…');
      return false;
    }
    if (!form || !isDirty(form)) return true;
    return window.confirm('Discard what you typed?');
  }

  // ---------- screens ----------
  function hideScreens() {
    for (const id of ['screenMsg', 'screenKey', 'main']) $(id).hidden = true;
  }

  function showMessage(title, paragraphs, action) {
    hideScreens();
    $('msgTitle').textContent = title;
    const nodes = paragraphs.map((p) => el('p', 'muted', p));
    if (action) {
      const btn = el('button', 'btn primary block', action.label);
      btn.type = 'button';
      btn.addEventListener('click', action.onClick);
      nodes.push(btn);
    }
    $('msgBody').replaceChildren(...nodes);
    $('screenMsg').hidden = false;
  }

  function showKey(message) {
    hideScreens();
    $('keyMsg').textContent = message || 'Enter the access key you were given.';
    $('keyInput').value = '';
    $('screenKey').hidden = false;
    $('keyInput').focus();
  }

  function showMain() {
    hideScreens();
    $('main').hidden = false;
  }

  function closeDialogs() {
    for (const dlg of document.querySelectorAll('dialog[open]')) dlg.close();
  }

  async function start() {
    showMessage('Sheet Forms', ['Loading…']);
    const ok = await run(loadInfo);
    if (ok) return;
    if (!$('screenMsg').hidden) {
      showMessage('Could not open the sheet', [state.lastError], { label: 'Try again', onClick: start });
    } else if (!$('main').hidden && state.loadedTab !== state.tabId) {
      state.loadFailed = true;
      renderList();
    }
  }

  // ---------- events ----------
  $('keyForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const key = $('keyInput').value.trim();
    if (!key) return;
    state.key = key;
    store.set('key', key);
    start();
  });

  $('refreshBtn').addEventListener('click', reloadTab);
  $('columnsBtn').addEventListener('click', openColumns);
  $('addBtn').addEventListener('click', openAdd);

  $('layoutBtn').addEventListener('click', () => {
    state.layout = state.layout === 'cards' ? 'table' : 'cards';
    store.set('layout', state.layout);
    renderList();
  });

  $('tabSelect').addEventListener('change', (e) => {
    state.tabId = Number(e.target.value);
    store.set('tab', state.tabId);
    state.viewId = null;
    state.custom = false;
    state.query = '';
    state.shown = PAGE;
    state.rows = [];
    state.loadedTab = null;
    $('search').value = '';
    reloadTab();
  });

  $('viewSelect').addEventListener('change', (e) => {
    const view = allViews().find((v) => v.id === e.target.value);
    if (!view) return;
    applyView(view);
    store.set(`view.${state.tabId}`, view.id);
    state.shown = PAGE;
    renderViewSelect();
    renderList();
  });

  $('search').addEventListener('input', (e) => {
    state.query = e.target.value;
    state.shown = PAGE;
    renderList();
  });

  $('columnsDlg').addEventListener('click', (e) => {
    const bulk = e.target.dataset && e.target.dataset.bulk;
    if (!bulk || !state.draft) return;
    const [kind, mode] = bulk.split('-');
    state.draft[kind] = new Set(mode === 'all'
      ? state.headers.map((_, c) => c).filter((c) => kind === 'view' || editable(c))
      : []);
    renderColumnsList();
  });

  $('columnsApply').addEventListener('click', () => {
    if (!state.draft) return;
    const cols = draftColumns();
    if (!state.canManage) {
      state.current = cols;
      state.custom = true;
      state.shown = PAGE;
      $('columnsDlg').close();
      renderViewSelect();
      renderList();
      return;
    }
    const v = activeView();
    saveView({ id: v.id, name: $('viewName').value.trim() || v.name || 'Untitled view', ...cols });
  });

  $('saveAsNew').addEventListener('click', () => {
    if (!state.draft) return;
    const typed = $('viewName').value.trim();
    const suggestion = typed && typed !== activeView().name ? typed : `${activeView().name} copy`;
    const name = window.prompt('Name for the new view:', suggestion);
    if (name === null || !name.trim()) return;
    saveView({ id: '', name: name.trim(), ...draftColumns() });
  });

  $('deleteView').addEventListener('click', async () => {
    const v = activeView();
    if (!v.id || !window.confirm(`Delete the view "${v.name}" for everyone?`)) return;
    const tabId = state.tabId;
    const draft = state.draft;
    setColumnsBusy(true);
    state.viewSaving = true;
    await run(async () => {
      const res = await call('deleteView', { tabId, id: v.id });
      if (tabId !== state.tabId) return;
      state.views = res.views || [];
      state.viewId = null;
      store.del(`view.${tabId}`);
      if (state.draft === draft) $('columnsDlg').close();
      selectView(false);
      renderList();
      toast('View deleted.');
    });
    state.viewSaving = false;
    setColumnsBusy(false);
  });

  // Cancel buttons; the row form asks before throwing away typed values.
  document.addEventListener('click', (e) => {
    const closer = e.target.closest && e.target.closest('[data-close]');
    if (!closer) return;
    const dlg = closer.closest('dialog');
    if (!dlg) return;
    if (dlg.id === 'rowDlg' && !confirmDiscard()) return;
    dlg.close();
  });

  // Esc key, or the Android back gesture.
  $('rowDlg').addEventListener('cancel', (e) => {
    if (!confirmDiscard()) e.preventDefault();
  });

  for (const dlg of document.querySelectorAll('dialog')) {
    dlg.addEventListener('close', () => {
      if (dlg.id === 'columnsDlg') state.draft = null;
      hostOverlays();
    });
  }

  // Enter in a one-line field moves to the next field rather than submitting.
  $('rowForm').addEventListener('keydown', (e) => {
    const t = e.target;
    if (e.key !== 'Enter' || e.isComposing || e.keyCode === 229 || t.tagName !== 'INPUT') return;
    e.preventDefault();
    const fields = [...$('rowBody').querySelectorAll('input, select, textarea')]
      .filter((n) => !n.disabled && n.type !== 'hidden');
    const next = fields[fields.indexOf(t) + 1];
    if (next) next.focus(); else t.blur();
  });

  $('rowForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const form = state.form;
    const save = $('rowSave');
    if (!form || form.saving) return;
    form.saving = true;
    save.disabled = true;
    save.textContent = 'Saving…';
    await run(() => (form.mode === 'edit' ? saveEdit(form) : saveAdd(form)));
    form.saving = false;
    if (state.form === form) {
      save.disabled = false;
      save.textContent = save.dataset.label || 'Save';
    }
  });

  // Pick up other people's changes when coming back to the app.
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState !== 'visible' || $('main').hidden) return;
    if (document.querySelector('dialog[open]') || busyCount > 0) return;
    if (Date.now() - state.loadedAt > STALE_MS) reloadTab();
  });

  // ---------- start ----------
  // A share link can carry the key: https://.../#key=YOUR-KEY
  // (the part after # never reaches a server). Returns true if a key was found.
  function readKeyFromLink() {
    const raw = location.hash.replace(/^#/, '');
    const parts = raw.split('&');
    const found = parts.find((p) => p.startsWith('key='));
    if (found === undefined) return false;
    let key = found.slice(4);
    try { key = decodeURIComponent(key); } catch { /* use as is */ }
    key = key.trim();
    state.key = key;
    if (key) store.set('key', key); else store.del('key');
    const rest = parts.filter((p) => p !== found).join('&');
    history.replaceState(null, '', location.pathname + location.search + (rest ? `#${rest}` : ''));
    return true;
  }

  window.addEventListener('hashchange', () => {
    if (readKeyFromLink() && API_URL) start();
  });

  readKeyFromLink();
  if (!API_URL) {
    showMessage('Setup needed', [
      'This app is not connected to a Google Sheet yet.',
      'Deploy apps-script/Code.gs from your Google Sheet as a web app, then put its URL in js/config.js. README.md has the steps.'
    ]);
  } else {
    start();
  }
})();
