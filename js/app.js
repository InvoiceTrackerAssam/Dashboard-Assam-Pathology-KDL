/* Assam Pathology Project – Daily Dashboard
 * Runs fully in the browser. Uploaded Excel files are aggregated per day and
 * saved in this browser's IndexedDB. Re-uploading a date replaces that date.
 */
(function () {
  'use strict';

  // ---------------------------------------------------------------- storage
  const DB_NAME = 'assam-pathology-dashboard';
  const KEY = 'store';

  function emptyStore() {
    // days[YYYY-MM-DD] = { p, t, r, fac:{name:{p,t,r,d}}, dist:{name:{p,t,r}}, tst:{name:{t,r}}, file, at }
    // p = patients, t = tests, r = revenue (₹), d = district of facility
    return { version: 1, demo: false, days: {}, uploads: [], mappings: {} };
  }

  function idbOpen() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore('kv');
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function loadStore() {
    try {
      const db = await idbOpen();
      const val = await new Promise((resolve, reject) => {
        const r = db.transaction('kv').objectStore('kv').get(KEY);
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
      return val ? Object.assign(emptyStore(), val) : emptyStore();
    } catch (e) {
      console.warn('IndexedDB unavailable, trying localStorage', e);
      try { const s = localStorage.getItem(DB_NAME); if (s) return Object.assign(emptyStore(), JSON.parse(s)); } catch (_) { /* ignore */ }
      return emptyStore();
    }
  }
  async function saveStore() {
    try {
      const db = await idbOpen();
      await new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readwrite');
        tx.objectStore('kv').put(store, KEY);
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {
      try { localStorage.setItem(DB_NAME, JSON.stringify(store)); } catch (e2) {
        toast('Could not save data in this browser. Please take a backup.', true);
      }
    }
  }

  // ---------------------------------------------------------------- helpers
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  const fmtInt = (n) => Math.round(n || 0).toLocaleString('en-IN');
  const fmtDec = (n, d = 2) => (n || 0).toLocaleString('en-IN', { minimumFractionDigits: d, maximumFractionDigits: d });
  const fmtINR = (n, d = 0) => '₹' + fmtDec(n, d);
  function fmtINRShort(n) {
    const a = Math.abs(n || 0);
    if (a >= 1e7) return '₹' + fmtDec(n / 1e7, 2) + ' Cr';
    if (a >= 1e5) return '₹' + fmtDec(n / 1e5, 2) + ' L';
    return fmtINR(n);
  }
  const div = (a, b) => (b ? a / b : 0);

  const pad = (n) => String(n).padStart(2, '0');
  const ymd = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;
  function fmtDate(s, withYear = true) {
    if (!s) return '';
    const [y, m, d] = s.split('-').map(Number);
    return `${pad(d)}-${MONTHS[m - 1]}` + (withYear ? `-${y}` : '');
  }
  function addDays(s, n) {
    const [y, m, d] = s.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d + n));
    return ymd(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
  }
  function todayYmd(offset = 0) {
    const d = new Date();
    d.setDate(d.getDate() + offset);
    return ymd(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }
  function validYmd(y, m, d) {
    if (y < 100) y += 2000;
    if (y < 2000 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCMonth() !== m - 1) return null;
    return ymd(y, m, d);
  }

  // Indian convention: numeric dates are DD/MM/YYYY.
  function parseDate(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) {
        const p = XLSX.SSF.parse_date_code(v);
        return p ? validYmd(p.y, p.m, p.d) : null;
      }
      return null;
    }
    if (v instanceof Date) return validYmd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    const s = String(v).trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return validYmd(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
    if (m) return validYmd(+m[3], +m[2], +m[1]);
    m = s.match(/^(\d{1,2})(?:st|nd|rd|th)?[-\s/.,]*([A-Za-z]{3,9})[-\s/.,]*(\d{2,4})/);
    if (m) {
      const mi = MONTHS.findIndex((x) => m[2].toLowerCase().startsWith(x.toLowerCase()));
      if (mi >= 0) return validYmd(+m[3], mi + 1, +m[1]);
    }
    m = s.match(/^([A-Za-z]{3,9})[-\s/.]*(\d{1,2})(?:st|nd|rd|th)?[-\s/.,]+(\d{4})/);
    if (m) {
      const mi = MONTHS.findIndex((x) => m[1].toLowerCase().startsWith(x.toLowerCase()));
      if (mi >= 0) return validYmd(+m[3], mi + 1, +m[2]);
    }
    if (/^\d{5}(\.\d+)?$/.test(s)) return parseDate(Number(s));
    return null;
  }

  function parseNum(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const s = String(v).replace(/rs\.?|inr|₹|,|\s/gi, '');
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  const cleanText = (v) => (v == null ? '' : String(v).replace(/\s+/g, ' ').trim());

  let toastTimer;
  function toast(msg, isError = false) {
    const t = $('toast');
    t.textContent = msg;
    t.className = 'toast' + (isError ? ' error' : '');
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, isError ? 7000 : 4500);
  }

  function downloadBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  // ---------------------------------------------------------------- column detection
  // Order matters: earlier fields claim a header first.
  const FIELDS = [
    { key: 'date', label: 'Date', modes: ['line', 'summary'], kw: [/date/, /^day$/] },
    { key: 'patients', label: 'Patient count', modes: ['summary'], kw: [/no\.?\s*of\s*patients?/, /patients?\s*count/, /^(total\s*)?patients$/, /footfall/, /^patient\s*nos?$/] },
    { key: 'qty', label: 'Test count / Qty', modes: ['line', 'summary'], kw: [/^qty$/, /quantity/, /no\.?\s*of\s*(tests?|investigations?)/, /tests?\s*count/, /^total\s*tests?$/, /^tests$/, /tests?\s*(done|performed)/] },
    { key: 'patientId', label: 'Patient ID / Reg. No.', modes: ['line'], kw: [/patient.*\b(id|no|code|number)\b/, /uhid/, /\bmrn\b/, /reg(istration|n)?\.?\s*(no|id|number)/, /barcode/, /lab\s*(no|id)/, /visit\s*(no|id)/, /^sid$/, /sample\s*(id|no)/, /accession/, /bill\s*no/, /^patient$/] },
    { key: 'testName', label: 'Test name', modes: ['line'], kw: [/test\s*name/, /investigation/, /^test$/, /service/, /parameter/, /test\s*desc/, /test\s*code/, /profile/] },
    { key: 'amount', label: 'Amount / Revenue (₹)', modes: ['line', 'summary'], kw: [/net\s*amount/, /amount/, /revenue/, /\brate\b/, /\b(rs|inr)\b/, /price/, /value/, /billing/, /\bamt\b/, /charges?/, /tariff/] },
    { key: 'district', label: 'District', modes: ['line', 'summary'], kw: [/district/] },
    { key: 'facility', label: 'Facility / Hospital / Centre', modes: ['line', 'summary'], kw: [/hospital/, /facility/, /cent(re|er)/, /institution/, /^lab$/, /lab\s*name/, /site/, /\b(phc|chc|fru|sdch|mch|bphc|dh)\b/, /location/, /unit/] },
  ];
  const REQUIRED = { line: ['patientId'], summary: ['patients', 'qty'] };

  function normHeader(h) { return cleanText(h).toLowerCase().replace(/[_]+/g, ' '); }

  function findHeaderRow(rows) {
    let best = 0, bestScore = -1;
    for (let i = 0; i < Math.min(rows.length, 25); i++) {
      const r = rows[i] || [];
      let score = 0, filled = 0;
      for (const c of r) {
        if (c == null || c === '') continue;
        filled++;
        const h = normHeader(c);
        if (typeof c === 'string' && FIELDS.some((f) => f.kw.some((k) => k.test(h)))) score += 2;
      }
      score += filled * 0.01;
      if (score > bestScore) { best = i; bestScore = score; }
    }
    return best;
  }

  function autoMap(headers) {
    const map = {};
    const used = new Set();
    const norm = headers.map(normHeader);
    for (const f of FIELDS) {
      for (const k of f.kw) {
        const idx = norm.findIndex((h, i) => h && !used.has(i) && k.test(h));
        if (idx >= 0) { map[f.key] = idx; used.add(idx); break; }
      }
    }
    return map;
  }
  const headerSig = (headers) => headers.map(normHeader).filter(Boolean).join('|');

  function dateFromFileName(name) {
    const base = name.replace(/\.[^.]+$/, '');
    const m = base.match(/(\d{1,2}[-_. ]\d{1,2}[-_. ]\d{2,4}|\d{4}[-_. ]\d{1,2}[-_. ]\d{1,2}|\d{1,2}[-_. ]?[A-Za-z]{3,9}[-_. ]?\d{2,4})/);
    return m ? parseDate(m[1].replace(/_/g, '-')) : null;
  }

  // ---------------------------------------------------------------- aggregation
  const TOTAL_RE = /^\s*(grand\s*)?(sub\s*)?total\b/i;

  function aggregate(dataRows, map, mode, fallbackDate) {
    const days = {};
    const res = { days, rows: 0, skipped: 0, noDate: 0, noPatient: 0 };
    const sets = {}; // date -> { all:Set, fac:{name:Set}, dist:{name:Set} }

    for (const row of dataRows) {
      if (!row || row.every((c) => c == null || c === '')) continue;
      if (row.some((c) => typeof c === 'string' && TOTAL_RE.test(c))) { res.skipped++; continue; }

      const date = map.date != null ? parseDate(row[map.date]) : fallbackDate;
      if (!date) { res.noDate++; continue; }

      const fac = (map.facility != null && cleanText(row[map.facility])) || 'Unspecified';
      const dist = (map.district != null && cleanText(row[map.district])) || 'Unspecified';
      const amt = map.amount != null ? parseNum(row[map.amount]) : 0;

      let p, t;
      if (mode === 'line') {
        const pid = cleanText(row[map.patientId]);
        const test = map.testName != null ? cleanText(row[map.testName]) : '';
        if (!pid) {
          if (test || amt) res.noPatient++;
          continue;
        }
        t = map.qty != null ? (parseNum(row[map.qty]) || 1) : 1;
        if (!sets[date]) sets[date] = { all: new Set(), fac: {}, dist: {} };
        const S = sets[date];
        // patient counted once per day overall, once per facility, once per district
        p = 0;
        if (!S.all.has(pid)) { S.all.add(pid); p = 1; }
        const fs = S.fac[fac] || (S.fac[fac] = new Set());
        const fp = fs.has(pid) ? 0 : (fs.add(pid), 1);
        const ds = S.dist[dist] || (S.dist[dist] = new Set());
        const dp = ds.has(pid) ? 0 : (ds.add(pid), 1);
        const D = days[date] || (days[date] = newDay());
        D.p += p; D.t += t; D.r += amt;
        const F = D.fac[fac] || (D.fac[fac] = { p: 0, t: 0, r: 0, d: dist });
        F.p += fp; F.t += t; F.r += amt;
        const DS = D.dist[dist] || (D.dist[dist] = { p: 0, t: 0, r: 0 });
        DS.p += dp; DS.t += t; DS.r += amt;
        if (test) {
          const T = D.tst[test] || (D.tst[test] = { t: 0, r: 0 });
          T.t += t; T.r += amt;
        }
      } else {
        p = parseNum(row[map.patients]);
        t = parseNum(row[map.qty]);
        if (!p && !t && !amt) continue;
        const D = days[date] || (days[date] = newDay());
        D.p += p; D.t += t; D.r += amt;
        const F = D.fac[fac] || (D.fac[fac] = { p: 0, t: 0, r: 0, d: dist });
        F.p += p; F.t += t; F.r += amt;
        const DS = D.dist[dist] || (D.dist[dist] = { p: 0, t: 0, r: 0 });
        DS.p += p; DS.t += t; DS.r += amt;
      }
      res.rows++;
    }
    return res;
  }
  function newDay() { return { p: 0, t: 0, r: 0, fac: {}, dist: {}, tst: {} }; }

  // ---------------------------------------------------------------- import flow
  let pending = null; // { file, wb, sheet, rows, headerIdx, headers, map, mode }

  async function handleFile(file) {
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', raw: false, cellDates: false });
      pending = { file, wb };
      const sheetSel = $('impSheet');
      sheetSel.innerHTML = wb.SheetNames.map((n) => `<option>${esc(n)}</option>`).join('');
      // choose the sheet whose header row matches the most fields
      let bestSheet = wb.SheetNames[0], bestCount = -1;
      for (const n of wb.SheetNames) {
        const rows = sheetRows(wb, n);
        const hi = findHeaderRow(rows);
        const c = Object.keys(autoMap(rows[hi] || [])).length * 100000 + rows.length;
        if (c > bestCount) { bestCount = c; bestSheet = n; }
      }
      sheetSel.value = bestSheet;
      $('impSheetWrap').hidden = wb.SheetNames.length < 2;
      $('impFile').textContent = `${file.name} · ${(file.size / 1024).toFixed(0)} KB`;
      loadSheet(bestSheet, true);
      $('importDlg').returnValue = '';
      $('importDlg').showModal();
    } catch (e) {
      console.error(e);
      toast('Could not read this file. Please check it is a valid Excel/CSV file.', true);
    }
  }

  function sheetRows(wb, name) {
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
  }

  function loadSheet(name, useSaved) {
    const rows = sheetRows(pending.wb, name);
    const headerIdx = findHeaderRow(rows);
    const headers = (rows[headerIdx] || []).map((h, i) => cleanText(h) || `Column ${i + 1}`);
    let map = autoMap(rows[headerIdx] || []);
    let mode = map.patientId != null ? 'line' : (map.patients != null ? 'summary' : 'line');
    const saved = useSaved && store.mappings[headerSig(headers)];
    if (saved) {
      mode = saved.mode;
      map = {};
      for (const [k, h] of Object.entries(saved.map)) {
        const i = headers.indexOf(h);
        if (i >= 0) map[k] = i;
      }
    }
    Object.assign(pending, { sheet: name, rows, headerIdx, headers, map, mode });
    $('impMode').value = mode;
    const guess = dateFromFileName(pending.file.name);
    $('impDate').value = guess || todayYmd(-1);
    renderMapping();
  }

  function renderMapping() {
    const { headers, map, mode } = pending;
    const opts = (sel) => `<option value="">— not in file —</option>` +
      headers.map((h, i) => `<option value="${i}" ${sel === i ? 'selected' : ''}>${esc(h)}</option>`).join('');
    $('impMap').innerHTML = FIELDS.filter((f) => f.modes.includes(mode)).map((f) => `
      <label><span class="${REQUIRED[mode].includes(f.key) ? 'req' : ''}">${esc(f.label)}</span>
        <select data-field="${f.key}">${opts(map[f.key])}</select>
      </label>`).join('');
    $('impDateWrap').hidden = map.date != null;

    const prev = pending.rows.slice(pending.headerIdx, pending.headerIdx + 6);
    const cols = Math.min(headers.length, 12);
    $('impPreview').innerHTML = prev.map((r, ri) => `<tr>${Array.from({ length: cols }, (_, i) => {
      const v = r[i] == null ? '' : r[i];
      return ri === 0 ? `<th>${esc(v)}</th>` : `<td>${esc(v)}</td>`;
    }).join('')}</tr>`).join('');
    updateImportSummary();
  }

  function updateImportSummary() {
    const { map, mode } = pending;
    const box = $('impSummary');
    const missing = REQUIRED[mode].filter((k) => map[k] == null);
    if (missing.length) {
      const names = missing.map((k) => FIELDS.find((f) => f.key === k).label).join(', ');
      box.innerHTML = `<span class="err">Please select the column for: ${esc(names)}</span>` +
        (mode === 'line' ? `<div class="muted small" style="margin-top:6px">A patient ID / registration no. column is needed to count unique patients. If your file is a daily summary with patient & test counts, change “File format” above.</div>` : '');
      $('impGo').disabled = true;
      pending.result = null;
      return;
    }
    const fallback = map.date == null ? $('impDate').value : null;
    if (map.date == null && !fallback) {
      box.innerHTML = '<span class="err">Please select the report date for this file.</span>';
      $('impGo').disabled = true;
      return;
    }
    const res = aggregate(pending.rows.slice(pending.headerIdx + 1), map, mode, fallback);
    pending.result = res;
    const dates = Object.keys(res.days).sort();
    const tot = dates.reduce((a, d) => { const x = res.days[d]; a.p += x.p; a.t += x.t; a.r += x.r; return a; }, { p: 0, t: 0, r: 0 });
    const existing = dates.filter((d) => store.days[d]);
    let html = '';
    if (!dates.length) {
      html = '<span class="err">No valid data rows found. Please check the selected columns.</span>';
    } else {
      const range = dates.length === 1 ? fmtDate(dates[0]) : `${fmtDate(dates[0])} to ${fmtDate(dates[dates.length - 1])} (${dates.length} days)`;
      html = `<div>Data date: <b>${range}</b> · ${fmtInt(res.rows)} rows read</div>
        <div class="stats">
          <div><span class="muted small">Patients</span><b>${fmtInt(tot.p)}</b></div>
          <div><span class="muted small">Tests</span><b>${fmtInt(tot.t)}</b></div>
          <div><span class="muted small">Revenue</span><b>${fmtINR(tot.r)}</b></div>
          <div><span class="muted small">Revenue / patient</span><b>${fmtINR(div(tot.r, tot.p), 2)}</b></div>
          <div><span class="muted small">Tests / patient</span><b>${fmtDec(div(tot.t, tot.p))}</b></div>
        </div>`;
      const warns = [];
      if (existing.length) warns.push(`Data already exists for ${existing.length === 1 ? fmtDate(existing[0]) : existing.length + ' of these dates'} — it will be <b>replaced</b> with this file.`);
      if (res.noDate) warns.push(`${fmtInt(res.noDate)} rows skipped because the date could not be read.`);
      if (res.noPatient) warns.push(`${fmtInt(res.noPatient)} rows skipped because Patient ID is blank.`);
      if (res.skipped) warns.push(`${fmtInt(res.skipped)} total/sub-total rows ignored.`);
      if (map.amount == null) warns.push('No amount column selected — revenue will be shown as ₹0.');
      html += warns.map((w) => `<div class="warn">${w}</div>`).join('');
    }
    box.innerHTML = html;
    $('impGo').disabled = !dates.length;
  }

  async function commitImport() {
    const { result, map, mode, headers, file } = pending;
    if (!result) return;
    if (store.demo) { store = emptyStore(); }
    const dates = Object.keys(result.days).sort();
    const now = new Date().toISOString();
    let tot = { p: 0, t: 0, r: 0 };
    for (const d of dates) {
      const day = result.days[d];
      day.file = file.name; day.at = now;
      store.days[d] = day;
      tot.p += day.p; tot.t += day.t; tot.r += day.r;
    }
    store.uploads.unshift({ file: file.name, at: now, dates, mode, rows: result.rows, ...tot });
    store.uploads = store.uploads.slice(0, 200);
    const namedMap = {};
    for (const [k, i] of Object.entries(map)) if (i != null) namedMap[k] = headers[i];
    store.mappings[headerSig(headers)] = { mode, map: namedMap };
    await saveStore();
    // jump the dashboard to the newest uploaded date
    filters.reportDate = dates[dates.length - 1];
    pending = null;
    render();
    toast(`Imported ${dates.length === 1 ? fmtDate(dates[0]) : dates.length + ' days'}: ${fmtInt(tot.p)} patients, ${fmtInt(tot.t)} tests, ${fmtINR(tot.r)}`);
  }

  // ---------------------------------------------------------------- dashboard state
  let store = emptyStore();
  const filters = { reportDate: null, period: 'mtd', from: null, district: '', facility: '' };
  const charts = {};

  function allDates() { return Object.keys(store.days).sort(); }

  function metricsFor(day) {
    if (!day) return null;
    if (filters.facility) return day.fac[filters.facility] || null;
    if (filters.district) return day.dist[filters.district] || null;
    return day;
  }

  function periodRange(dates) {
    const to = filters.reportDate;
    let from;
    switch (filters.period) {
      case 'mtd': from = to.slice(0, 8) + '01'; break;
      case '7': from = addDays(to, -6); break;
      case '30': from = addDays(to, -29); break;
      case '90': from = addDays(to, -89); break;
      case 'all': from = dates[0]; break;
      case 'custom': from = filters.from && filters.from <= to ? filters.from : to.slice(0, 8) + '01'; break;
    }
    return { from, to };
  }

  // ---------------------------------------------------------------- render
  function render() {
    const dates = allDates();
    const has = dates.length > 0;
    $('emptyState').hidden = has;
    $('dashboard').hidden = !has;
    $('demoBanner').hidden = !store.demo;
    if (!has) return;

    if (!filters.reportDate || !store.days[filters.reportDate] && filters.reportDate > dates[dates.length - 1]) {
      filters.reportDate = dates[dates.length - 1];
    }
    if (filters.reportDate < dates[0]) filters.reportDate = dates[0];
    $('fReportDate').value = filters.reportDate;
    $('fReportDate').min = dates[0];
    $('fReportDate').max = dates[dates.length - 1];

    renderFilterOptions(dates);
    const { from, to } = periodRange(dates);
    $('fFromWrap').hidden = filters.period !== 'custom';
    $('fFrom').value = from;
    $('fFrom').max = to;

    const lastUp = store.uploads[0];
    $('lastUpdated').textContent = lastUp ? `Last upload: ${lastUp.file} · ${new Date(lastUp.at).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}` : '';

    renderDay(dates);
    const series = buildSeries(from, to);
    renderPeriod(series, from, to);
    renderCharts(series);
    renderBreakdowns(from, to);
    renderDaily(dates);
    renderUploads();
  }

  function renderFilterOptions(dates) {
    const districts = new Set(), facilities = new Map();
    for (const d of dates) {
      const day = store.days[d];
      Object.keys(day.dist).forEach((x) => districts.add(x));
      for (const [f, v] of Object.entries(day.fac)) facilities.set(f, v.d);
    }
    const dl = [...districts].sort();
    const onlyUnspecified = (arr) => arr.length <= 1 && (arr[0] || 'Unspecified') === 'Unspecified';
    $('fDistrictWrap').hidden = onlyUnspecified(dl);
    if (!dl.includes(filters.district)) filters.district = '';
    $('fDistrict').innerHTML = '<option value="">All districts</option>' +
      dl.map((x) => `<option ${x === filters.district ? 'selected' : ''}>${esc(x)}</option>`).join('');

    const fl = [...facilities.entries()].filter(([, d]) => !filters.district || d === filters.district).map(([f]) => f).sort();
    $('fFacilityWrap').hidden = onlyUnspecified([...facilities.keys()]);
    if (!fl.includes(filters.facility)) filters.facility = '';
    $('fFacility').innerHTML = '<option value="">All facilities</option>' +
      fl.map((x) => `<option ${x === filters.facility ? 'selected' : ''}>${esc(x)}</option>`).join('');
  }

  function scopeLabel() {
    return filters.facility || filters.district || 'All Assam';
  }

  function deltaHtml(cur, prev, isMoney) {
    if (prev == null || !isFinite(prev)) return '';
    if (!prev) return cur ? '<span class="delta flat">new</span>' : '';
    const pct = ((cur - prev) / Math.abs(prev)) * 100;
    if (Math.abs(pct) < 0.05) return '<span class="delta flat">● 0.0%</span>';
    const up = pct > 0;
    return `<span class="delta ${up ? 'up' : 'down'}" title="Previous: ${isMoney ? fmtINR(prev, 2) : fmtDec(prev)}">${up ? '▲' : '▼'} ${fmtDec(Math.abs(pct), 1)}%</span>`;
  }

  function kpiTiles(m, prev, extra = {}) {
    const P = prev || null;
    const tiles = [
      { label: 'Patients', value: fmtInt(m.p), d: P && deltaHtml(m.p, P.p), meta: extra.p },
      { label: 'Tests', value: fmtInt(m.t), d: P && deltaHtml(m.t, P.t), meta: extra.t },
      { label: 'Revenue', value: fmtINRShort(m.r), title: fmtINR(m.r, 2), d: P && deltaHtml(m.r, P.r, true), meta: extra.r },
      { label: 'Avg revenue / patient', value: fmtINR(div(m.r, m.p), 2), d: P && deltaHtml(div(m.r, m.p), div(P.r, P.p), true) },
      { label: 'Avg revenue / test', value: fmtINR(div(m.r, m.t), 2), d: P && deltaHtml(div(m.r, m.t), div(P.r, P.t), true) },
      { label: 'Avg tests / patient', value: fmtDec(div(m.t, m.p)), d: P && deltaHtml(div(m.t, m.p), div(P.t, P.p)) },
    ];
    return tiles.map((t) => `
      <div class="kpi">
        <div class="label">${t.label}</div>
        <div class="value" ${t.title ? `title="${t.title}"` : ''}>${t.value}</div>
        <div class="meta">${t.d || ''}${t.d && t.meta ? ' · ' : ''}${t.meta || ''}</div>
      </div>`).join('');
  }

  function renderDay(dates) {
    const d = filters.reportDate;
    const m = metricsFor(store.days[d]) || { p: 0, t: 0, r: 0 };
    const idx = dates.indexOf(d);
    const prevDate = idx > 0 ? dates[idx - 1] : dates.filter((x) => x < d).pop();
    const prev = prevDate ? (metricsFor(store.days[prevDate]) || { p: 0, t: 0, r: 0 }) : null;
    $('dayTitle').textContent = `Day — ${fmtDate(d)} · ${scopeLabel()}`;
    $('dayCompare').textContent = store.days[d]
      ? (prevDate ? `Change vs ${fmtDate(prevDate)}` : 'First day of data')
      : 'No data uploaded for this date';
    $('dayKpis').innerHTML = kpiTiles(m, prev);
  }

  function buildSeries(from, to) {
    const out = [];
    for (let d = from; d <= to; d = addDays(d, 1)) {
      const day = store.days[d];
      const m = day ? metricsFor(day) || { p: 0, t: 0, r: 0 } : null;
      out.push({ date: d, m });
    }
    return out;
  }

  function renderPeriod(series, from, to) {
    const withData = series.filter((s) => s.m);
    const tot = withData.reduce((a, s) => { a.p += s.m.p; a.t += s.m.t; a.r += s.m.r; return a; }, { p: 0, t: 0, r: 0 });
    const n = withData.length;
    $('periodTitle').textContent = `Period — ${fmtDate(from)} to ${fmtDate(to)}`;
    $('periodInfo').textContent = `${n} day${n === 1 ? '' : 's'} with data · ${scopeLabel()}`;
    $('periodKpis').innerHTML = kpiTiles(tot, null, {
      p: n ? `Daily avg ${fmtInt(tot.p / n)}` : '',
      t: n ? `Daily avg ${fmtInt(tot.t / n)}` : '',
      r: n ? `Daily avg ${fmtINRShort(tot.r / n)}` : '',
    });
  }

  // ---------------------------------------------------------------- charts
  function css(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }

  function renderCharts(series) {
    const labels = series.map((s) => fmtDate(s.date, false));
    const v = (fn) => series.map((s) => (s.m ? fn(s.m) : null));
    const c1 = css('--series-1'), c2 = css('--series-2');
    const legend = (el, items) => { if (el) el.innerHTML = items.map(([n, c]) => `<span><i style="background:${c}"></i>${n}</span>`).join(''); };

    legend($('lgCounts'), [['Patients', c1], ['Tests', c2]]);
    legend($('lgAvgRev'), [['Revenue / patient', c1], ['Revenue / test', c2]]);

    drawChart('chCounts', 'line', labels, [
      { label: 'Patients', data: v((m) => m.p), color: c1 },
      { label: 'Tests', data: v((m) => m.t), color: c2 },
    ], (x) => fmtInt(x));
    drawChart('chRevenue', 'bar', labels, [
      { label: 'Revenue', data: v((m) => m.r), color: c1 },
    ], (x) => fmtINR(x), (x) => fmtINRShort(x));
    drawChart('chAvgRev', 'line', labels, [
      { label: 'Revenue / patient', data: v((m) => div(m.r, m.p)), color: c1 },
      { label: 'Revenue / test', data: v((m) => div(m.r, m.t)), color: c2 },
    ], (x) => fmtINR(x, 2), (x) => '₹' + fmtInt(x));
    drawChart('chTpp', 'line', labels, [
      { label: 'Tests / patient', data: v((m) => div(m.t, m.p)), color: c1 },
    ], (x) => fmtDec(x), (x) => fmtDec(x, 1));
  }

  function drawChart(id, type, labels, sets, tipFmt, axisFmt) {
    const text2 = css('--text-2'), text3 = css('--text-3'), grid = css('--grid'), surface = css('--surface');
    const many = labels.length > 45;
    const datasets = sets.map((s) => type === 'bar' ? {
      label: s.label, data: s.data, backgroundColor: s.color, hoverBackgroundColor: s.color,
      borderRadius: { topLeft: 4, topRight: 4 }, borderSkipped: 'bottom', maxBarThickness: 28, categoryPercentage: 0.8, barPercentage: 0.9,
    } : {
      label: s.label, data: s.data, borderColor: s.color, backgroundColor: s.color, borderWidth: 2, tension: 0.25,
      pointRadius: many ? 0 : 3, pointHoverRadius: 5, pointBackgroundColor: s.color, pointBorderColor: surface, pointBorderWidth: 2,
      spanGaps: false,
    });
    const options = {
      responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: surface, titleColor: css('--text'), bodyColor: text2, borderColor: css('--border'), borderWidth: 1,
          padding: 10, boxPadding: 4, usePointStyle: true,
          callbacks: { label: (ctx) => ` ${ctx.dataset.label}: ${ctx.parsed.y == null ? '—' : tipFmt(ctx.parsed.y)}` },
        },
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: text3, maxRotation: 0, autoSkip: true, autoSkipPadding: 12 }, border: { color: grid } },
        y: { beginAtZero: true, grid: { color: grid }, border: { display: false }, ticks: { color: text3, maxTicksLimit: 6, callback: (x) => (axisFmt || tipFmt)(x) } },
      },
    };
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart($(id), { type, data: { labels, datasets }, options });
  }

  // ---------------------------------------------------------------- tables
  const sortState = { facility: { key: 'r', dir: -1 } };

  function renderBreakdowns(from, to) {
    const dist = {}, fac = {}, tst = {};
    let anyTests = false;
    for (let d = from; d <= to; d = addDays(d, 1)) {
      const day = store.days[d];
      if (!day) continue;
      for (const [k, v] of Object.entries(day.dist)) {
        const o = dist[k] || (dist[k] = { p: 0, t: 0, r: 0 });
        o.p += v.p; o.t += v.t; o.r += v.r;
      }
      for (const [k, v] of Object.entries(day.fac)) {
        if (filters.district && v.d !== filters.district) continue;
        const o = fac[k] || (fac[k] = { p: 0, t: 0, r: 0, d: v.d });
        o.p += v.p; o.t += v.t; o.r += v.r;
      }
      for (const [k, v] of Object.entries(day.tst || {})) {
        anyTests = true;
        const o = tst[k] || (tst[k] = { t: 0, r: 0 });
        o.t += v.t; o.r += v.r;
      }
    }

    const distKeys = Object.keys(dist);
    const showDist = !(distKeys.length <= 1 && (distKeys[0] || 'Unspecified') === 'Unspecified');
    $('districtCard').hidden = !showDist;
    if (showDist) $('tblDistrict').innerHTML = metricTable('District', dist, 'r', -1, filters.district);

    const facKeys = Object.keys(fac);
    const showFac = !(facKeys.length <= 1 && (facKeys[0] || 'Unspecified') === 'Unspecified');
    $('facilityCard').hidden = !showFac;
    if (showFac) {
      const q = $('facSearch').value.trim().toLowerCase();
      const filtered = Object.fromEntries(Object.entries(fac).filter(([k, v]) => !q || k.toLowerCase().includes(q) || (v.d || '').toLowerCase().includes(q)));
      const s = sortState.facility;
      $('tblFacility').innerHTML = metricTable('Facility', filtered, s.key, s.dir, filters.facility, true);
    }

    $('testCard').hidden = !anyTests;
    if (anyTests) {
      const totT = Object.values(tst).reduce((a, v) => a + v.t, 0);
      const list = Object.entries(tst).sort((a, b) => b[1].t - a[1].t);
      const max = list.length ? list[0][1].t : 1;
      $('testNote').textContent = filters.district || filters.facility ? 'All facilities' : `${list.length} tests`;
      $('tblTests').innerHTML = `<thead><tr><th>Test</th><th>Count</th><th>Share</th><th>Revenue</th><th>Avg ₹</th></tr></thead><tbody>` +
        list.slice(0, 50).map(([k, v]) => `<tr>
          <td class="wrap">${esc(k)}</td>
          <td><div class="bar-cell"><span class="bar" style="width:${Math.max(2, (v.t / max) * 60)}px"></span>${fmtInt(v.t)}</div></td>
          <td>${fmtDec(div(v.t, totT) * 100, 1)}%</td>
          <td>${fmtINR(v.r)}</td>
          <td>${fmtINR(div(v.r, v.t))}</td></tr>`).join('') + '</tbody>';
    }
    // facility table always takes a full row; district/test share a row when both exist
    $('districtCard').parentElement.hidden = !(showDist || showFac || anyTests);
    $('districtCard').style.gridColumn = anyTests ? '' : '1 / -1';
    $('testCard').style.gridColumn = showDist ? '' : '1 / -1';
  }

  function metricTable(title, obj, sortKey, dir, highlight, sortable) {
    const rows = Object.entries(obj).map(([k, v]) => ({ k, ...v, rpp: div(v.r, v.p), rpt: div(v.r, v.t), tpp: div(v.t, v.p) }));
    rows.sort((a, b) => sortKey === 'k' ? dir * a.k.localeCompare(b.k) : dir * (a[sortKey] - b[sortKey]));
    const tot = rows.reduce((a, r) => { a.p += r.p; a.t += r.t; a.r += r.r; return a; }, { p: 0, t: 0, r: 0 });
    const cols = [['k', title], ['p', 'Patients'], ['t', 'Tests'], ['r', 'Revenue'], ['rpp', '₹/Patient'], ['rpt', '₹/Test'], ['tpp', 'Tests/Pt']];
    const head = cols.map(([k, l]) => {
      const arrow = sortable && sortKey === k ? (dir < 0 ? ' ▼' : ' ▲') : '';
      return `<th ${sortable ? `class="sortable" data-sort="${k}"` : ''}>${l}${arrow}</th>`;
    }).join('');
    if (!rows.length) return `<thead><tr>${head}</tr></thead><tbody><tr class="empty-row"><td colspan="7">No matching rows</td></tr></tbody>`;
    return `<thead><tr>${head}</tr></thead><tbody>` + rows.map((r) => `<tr ${r.k === highlight ? 'style="font-weight:600"' : ''}>
      <td class="wrap">${esc(r.k)}</td><td>${fmtInt(r.p)}</td><td>${fmtInt(r.t)}</td><td>${fmtINR(r.r)}</td>
      <td>${fmtINR(r.rpp)}</td><td>${fmtINR(r.rpt)}</td><td>${fmtDec(r.tpp)}</td></tr>`).join('') +
      `</tbody><tfoot><tr><td>Total (${rows.length})</td><td>${fmtInt(tot.p)}</td><td>${fmtInt(tot.t)}</td><td>${fmtINR(tot.r)}</td>
      <td>${fmtINR(div(tot.r, tot.p))}</td><td>${fmtINR(div(tot.r, tot.t))}</td><td>${fmtDec(div(tot.t, tot.p))}</td></tr></tfoot>`;
  }

  function dailyRows(dates) {
    return dates.slice().reverse().map((d) => {
      const m = metricsFor(store.days[d]) || { p: 0, t: 0, r: 0 };
      return { d, p: m.p, t: m.t, r: m.r, rpp: div(m.r, m.p), rpt: div(m.r, m.t), tpp: div(m.t, m.p), file: store.days[d].file || '' };
    });
  }

  function renderDaily(dates) {
    const rows = dailyRows(dates);
    $('tblDaily').innerHTML = `<thead><tr><th>Date</th><th>Patients</th><th>Tests</th><th>Revenue</th><th>Avg ₹ / Patient</th><th>Avg ₹ / Test</th><th>Avg Tests / Patient</th><th>Source file</th><th></th></tr></thead><tbody>` +
      rows.map((r) => `<tr ${r.d === filters.reportDate ? 'style="font-weight:600"' : ''}>
        <td>${fmtDate(r.d)}</td><td>${fmtInt(r.p)}</td><td>${fmtInt(r.t)}</td><td>${fmtINR(r.r)}</td>
        <td>${fmtINR(r.rpp, 2)}</td><td>${fmtINR(r.rpt, 2)}</td><td>${fmtDec(r.tpp)}</td>
        <td class="muted small">${esc(r.file)}</td>
        <td><button class="row-del" type="button" data-del="${r.d}" title="Delete data for ${fmtDate(r.d)}" aria-label="Delete data for ${fmtDate(r.d)}">✕</button></td></tr>`).join('') + '</tbody>';
  }

  function renderUploads() {
    const ups = store.uploads;
    $('tblUploads').innerHTML = `<thead><tr><th>Uploaded on</th><th>File</th><th>Data date(s)</th><th>Format</th><th>Patients</th><th>Tests</th><th>Revenue</th></tr></thead><tbody>` +
      (ups.length ? ups.slice(0, 60).map((u) => `<tr>
        <td>${new Date(u.at).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}</td>
        <td class="wrap">${esc(u.file)}</td>
        <td>${u.dates.length === 1 ? fmtDate(u.dates[0]) : `${fmtDate(u.dates[0])} – ${fmtDate(u.dates[u.dates.length - 1])}`}</td>
        <td>${u.mode === 'summary' ? 'Summary' : 'Line list'}</td>
        <td>${fmtInt(u.p)}</td><td>${fmtInt(u.t)}</td><td>${fmtINR(u.r)}</td></tr>`).join('')
        : '<tr class="empty-row"><td colspan="7">No uploads yet</td></tr>') + '</tbody>';
  }

  // ---------------------------------------------------------------- menu actions
  function exportDaily() {
    const dates = allDates();
    if (!dates.length) return toast('No data to export yet.', true);
    const rows = dailyRows(dates).reverse().map((r) => ({
      'Date': fmtDate(r.d), 'Patients': r.p, 'Tests': r.t, 'Revenue (₹)': Math.round(r.r * 100) / 100,
      'Avg Revenue / Patient (₹)': Math.round(r.rpp * 100) / 100, 'Avg Revenue / Test (₹)': Math.round(r.rpt * 100) / 100,
      'Avg Tests / Patient': Math.round(r.tpp * 100) / 100,
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 13 }, { wch: 10 }, { wch: 10 }, { wch: 14 }, { wch: 24 }, { wch: 22 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Daily Summary');
    const sfx = filters.facility || filters.district ? '_' + (filters.facility || filters.district).replace(/[^\w]+/g, '_') : '';
    XLSX.writeFile(wb, `Assam_Pathology_Daily_Summary${sfx}_${todayYmd()}.xlsx`);
  }

  function downloadTemplate() {
    const wb = XLSX.utils.book_new();
    const line = [
      ['Date', 'Patient ID', 'Patient Name', 'District', 'Facility', 'Test Name', 'Amount'],
      ['24-10-2026', 'KDL-0001', 'Sample Patient 1', 'Kamrup Metro', 'Example Hospital A', 'CBC', 150],
      ['24-10-2026', 'KDL-0001', 'Sample Patient 1', 'Kamrup Metro', 'Example Hospital A', 'Blood Sugar (F)', 50],
      ['24-10-2026', 'KDL-0002', 'Sample Patient 2', 'Nagaon', 'Example Hospital B', 'LFT', 400],
    ];
    const summary = [
      ['Date', 'District', 'Facility', 'No of Patients', 'No of Tests', 'Revenue'],
      ['24-10-2026', 'Kamrup Metro', 'Example Hospital A', 120, 410, 61500],
      ['24-10-2026', 'Nagaon', 'Example Hospital B', 85, 260, 38900],
    ];
    const w1 = XLSX.utils.aoa_to_sheet(line);
    w1['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 16 }, { wch: 24 }, { wch: 22 }, { wch: 10 }];
    const w2 = XLSX.utils.aoa_to_sheet(summary);
    w2['!cols'] = [{ wch: 12 }, { wch: 16 }, { wch: 24 }, { wch: 15 }, { wch: 12 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, w1, 'Line list (one row per test)');
    XLSX.utils.book_append_sheet(wb, w2, 'Daily summary');
    XLSX.writeFile(wb, 'Assam_Pathology_Upload_Template.xlsx');
  }

  function backup() {
    const blob = new Blob([JSON.stringify(store)], { type: 'application/json' });
    downloadBlob(blob, `Assam_Pathology_Dashboard_Backup_${todayYmd()}.json`);
  }

  async function restore(file) {
    try {
      const data = JSON.parse(await file.text());
      if (!data || typeof data.days !== 'object') throw new Error('bad file');
      if (allDates().length && !confirm('Restoring will replace all data currently in this dashboard. Continue?')) return;
      store = Object.assign(emptyStore(), data);
      await saveStore();
      filters.reportDate = null;
      render();
      toast(`Backup restored: ${allDates().length} days of data.`);
    } catch (e) {
      toast('This is not a valid dashboard backup file.', true);
    }
  }

  async function clearAll() {
    if (!confirm(store.demo ? 'Clear the sample data?' : 'This will permanently delete ALL uploaded data from this browser. Take a backup first if needed. Continue?')) return;
    const mappings = store.demo ? {} : store.mappings;
    store = emptyStore();
    store.mappings = mappings;
    await saveStore();
    filters.reportDate = null;
    render();
    toast('All data cleared.');
  }

  function toggleTheme() {
    const root = document.documentElement;
    const isDark = root.dataset.theme ? root.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
    root.dataset.theme = isDark ? 'light' : 'dark';
    try { localStorage.setItem('apd-theme', root.dataset.theme); } catch (_) { /* ignore */ }
    render();
  }

  async function loadDemo() {
    // 45 days of illustrative data across a few districts
    const districts = { 'Kamrup Metro': ['GMCH Guwahati', 'Dispur FRU', 'Sonapur CHC'], 'Nagaon': ['Nagaon BPCH', 'Kaliabor CHC'], 'Dibrugarh': ['AMCH Dibrugarh', 'Chabua CHC'], 'Cachar': ['SMCH Silchar', 'Lakhipur CHC'], 'Jorhat': ['JMCH Jorhat'] };
    const tests = [['CBC', 150, 30], ['Blood Sugar (F)', 50, 22], ['LFT', 400, 10], ['KFT', 400, 9], ['Lipid Profile', 350, 7], ['Urine R/E', 80, 12], ['TSH', 250, 5], ['HbA1c', 300, 4], ['Malaria Ag', 120, 6], ['Dengue NS1', 500, 3]];
    const wsum = tests.reduce((a, t) => a + t[2], 0);
    store = emptyStore();
    store.demo = true;
    const end = todayYmd(-1);
    let seed = 7;
    const rnd = () => { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; };
    for (let i = 44; i >= 0; i--) {
      const d = addDays(end, -i);
      const day = newDay();
      const dow = new Date(d + 'T00:00:00').getDay();
      for (const [dist, facs] of Object.entries(districts)) {
        const DS = day.dist[dist] = { p: 0, t: 0, r: 0 };
        for (const f of facs) {
          const base = f.includes('MCH') ? 320 : f.includes('FRU') || f.includes('BPCH') ? 140 : 60;
          const p = Math.round(base * (dow === 0 ? 0.45 : 0.85 + rnd() * 0.3) * (1 + (44 - i) * 0.003));
          let t = 0, r = 0;
          for (const [name, price, w] of tests) {
            const n = Math.round(p * (w / wsum) * (2.6 + rnd() * 0.5));
            t += n; r += n * price;
            const T = day.tst[name] || (day.tst[name] = { t: 0, r: 0 });
            T.t += n; T.r += n * price;
          }
          day.fac[f] = { p, t, r, d: dist };
          DS.p += p; DS.t += t; DS.r += r;
          day.p += p; day.t += t; day.r += r;
        }
      }
      day.file = 'Sample data';
      day.at = new Date().toISOString();
      store.days[d] = day;
    }
    await saveStore();
    filters.reportDate = null;
    render();
  }

  // ---------------------------------------------------------------- events
  function bind() {
    const fileInput = $('fileInput');
    const openPicker = () => { fileInput.value = ''; fileInput.click(); };
    $('btnUpload').addEventListener('click', openPicker);
    $('btnUpload2').addEventListener('click', openPicker);
    fileInput.addEventListener('change', () => handleFile(fileInput.files[0]));
    $('btnDemo').addEventListener('click', loadDemo);

    // drag & drop anywhere on the page
    let dragDepth = 0;
    document.addEventListener('dragenter', (e) => { e.preventDefault(); dragDepth++; $('dropZone').classList.add('over'); });
    document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; $('dropZone').classList.remove('over'); } });
    document.addEventListener('dragover', (e) => e.preventDefault());
    document.addEventListener('drop', (e) => {
      e.preventDefault(); dragDepth = 0; $('dropZone').classList.remove('over');
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleFile(f);
    });

    // menu
    const menuBtn = $('btnMenu'), menu = $('menuList');
    menuBtn.addEventListener('click', (e) => { e.stopPropagation(); menu.hidden = !menu.hidden; menuBtn.setAttribute('aria-expanded', String(!menu.hidden)); });
    document.addEventListener('click', () => { menu.hidden = true; menuBtn.setAttribute('aria-expanded', 'false'); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') menu.hidden = true; });
    document.addEventListener('click', (e) => {
      const a = e.target.closest('[data-action]');
      if (!a) return;
      const act = a.dataset.action;
      if (act === 'template') downloadTemplate();
      else if (act === 'export') exportDaily();
      else if (act === 'backup') backup();
      else if (act === 'restore') { $('restoreInput').value = ''; $('restoreInput').click(); }
      else if (act === 'clear') clearAll();
      else if (act === 'theme') toggleTheme();
    });
    $('restoreInput').addEventListener('change', (e) => restore(e.target.files[0]));

    // import dialog
    $('impSheet').addEventListener('change', (e) => loadSheet(e.target.value, true));
    $('impMode').addEventListener('change', (e) => { pending.mode = e.target.value; renderMapping(); });
    $('impMap').addEventListener('change', (e) => {
      const f = e.target.dataset.field;
      if (!f) return;
      const val = e.target.value === '' ? null : Number(e.target.value);
      // a column can only be used for one field
      for (const k of Object.keys(pending.map)) if (val != null && pending.map[k] === val) delete pending.map[k];
      if (val == null) delete pending.map[f]; else pending.map[f] = val;
      renderMapping();
    });
    $('impDate').addEventListener('change', updateImportSummary);
    $('importDlg').addEventListener('close', () => {
      if ($('importDlg').returnValue === 'default' && pending && pending.result) commitImport();
      else pending = null;
    });

    // filters
    $('fReportDate').addEventListener('change', (e) => { if (e.target.value) { filters.reportDate = e.target.value; render(); } });
    $('fPeriod').addEventListener('change', (e) => { filters.period = e.target.value; render(); });
    $('fFrom').addEventListener('change', (e) => { filters.from = e.target.value; render(); });
    $('fDistrict').addEventListener('change', (e) => { filters.district = e.target.value; filters.facility = ''; render(); });
    $('fFacility').addEventListener('change', (e) => { filters.facility = e.target.value; render(); });
    $('facSearch').addEventListener('input', () => {
      const { from, to } = periodRange(allDates());
      renderBreakdowns(from, to);
    });
    $('tblFacility').addEventListener('click', (e) => {
      const th = e.target.closest('th[data-sort]');
      if (!th) return;
      const s = sortState.facility, k = th.dataset.sort;
      s.dir = s.key === k ? -s.dir : (k === 'k' ? 1 : -1);
      s.key = k;
      const { from, to } = periodRange(allDates());
      renderBreakdowns(from, to);
    });
    $('tblDaily').addEventListener('click', async (e) => {
      const b = e.target.closest('[data-del]');
      if (!b) return;
      const d = b.dataset.del;
      if (!confirm(`Delete all data for ${fmtDate(d)}? You can upload the file again later.`)) return;
      delete store.days[d];
      await saveStore();
      if (filters.reportDate === d) filters.reportDate = null;
      render();
      toast(`Data for ${fmtDate(d)} deleted.`);
    });

    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => render());
  }

  async function init() {
    try { const t = localStorage.getItem('apd-theme'); if (t) document.documentElement.dataset.theme = t; } catch (_) { /* ignore */ }
    if (window.Chart) {
      Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
      Chart.defaults.font.size = 12;
    }
    bind();
    store = await loadStore();
    render();
  }

  init();
})();
