/* ============================================================
   DASHBOARD PENJUALAN 37 OUTLET
   - Membaca sheet INPUT PENJUALAN + MASTER OUTLET dari file .xlsx
   - Mode data: file GitHub (otomatis), Google Sheets (real-time),
     unggah manual, dan data contoh
   ============================================================ */
'use strict';

/* ---------------- KONFIGURASI ---------------- */
const CFG = {
  cats: ['Dine in', 'Take away', 'GoFood', 'GrabFood', 'ShopeeFood', 'ESB Order'],
  catKeys: ['dineIn', 'takeAway', 'goFood', 'grabFood', 'shopeeFood', 'esbOrder'],
  catColors: ['#22d3ee', '#8b5cf6', '#34d399', '#fbbf24', '#f97316', '#f472b6'],
  months: ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'],
  monthsShort: ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'],
  refreshMs: 60000,
  defaultDataUrl: 'data/penjualan.xlsx'
};
const LS_KEY = 'dash37.v1';

/* ---------------- STATE ---------------- */
const S = {
  mode: 'github',            // github | gsheets | upload | demo
  dataUrl: CFG.defaultDataUrl,
  gsUrl: '',
  records: [],
  outlets: {},
  outletOrder: [],
  year: null,
  monMonth: null,
  search: '',
  lastUpdate: null,
  demo: false,
  loading: false,
  view: null,
  charts: {}
};

/* ---------------- UTIL ---------------- */
const $ = s => document.querySelector(s);
const nfIDR = new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 });
const nfNum = new Intl.NumberFormat('id-ID', { maximumFractionDigits: 0 });

function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function normName(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function fmtRp(v) { return (v == null || isNaN(v)) ? '—' : nfIDR.format(Math.round(v)); }
function trim1(x) { return (Math.round(x * 10) / 10).toLocaleString('id-ID', { maximumFractionDigits: 1 }); }
function fmtShort(v) {
  if (v == null || isNaN(v)) return '—';
  const a = Math.abs(v);
  if (a >= 1e12) return 'Rp ' + trim1(v / 1e12) + ' T';
  if (a >= 1e9) return 'Rp ' + trim1(v / 1e9) + ' M';
  if (a >= 1e6) return 'Rp ' + trim1(v / 1e6) + ' jt';
  if (a >= 1e3) return 'Rp ' + trim1(v / 1e3) + ' rb';
  return 'Rp ' + nfNum.format(Math.round(v));
}
function fmtShortN(v) {
  if (v == null || isNaN(v)) return '—';
  const a = Math.abs(v);
  if (a >= 1e12) return trim1(v / 1e12) + ' T';
  if (a >= 1e9) return trim1(v / 1e9) + ' M';
  if (a >= 1e6) return trim1(v / 1e6) + ' jt';
  if (a >= 1e3) return trim1(v / 1e3) + ' rb';
  return nfNum.format(Math.round(v));
}
function fmtPct(p, digits) {
  if (p == null || isNaN(p) || !isFinite(p)) return '—';
  return (p * 100).toLocaleString('id-ID', { maximumFractionDigits: digits == null ? 1 : digits }) + '%';
}
function timeStr(d) { return d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }); }
function cacheBust(u) { return u + (u.includes('?') ? '&' : '?') + 't=' + Date.now(); }
function debounce(fn, ms) { let t; return function () { clearTimeout(t); const a = arguments; t = setTimeout(() => fn.apply(null, a), ms); }; }

/* ---------------- PARSER NILAI ---------------- */
function monthFromName(s) {
  const t = String(s).toLowerCase().slice(0, 3);
  const list = ['jan', 'feb', 'mar', 'apr', 'mei', 'jun', 'jul', 'agu', 'sep', 'okt', 'nov', 'des'];
  const i = list.indexOf(t);
  if (i >= 0) return i + 1;
  const en = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const j = en.indexOf(t);
  return j >= 0 ? j + 1 : null;
}

function parsePeriod(v) {
  try {
    if (v instanceof Date && !isNaN(v)) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    // Nomor serial Excel (1900-system, 25569 = 1970-01-01) — deterministik lintas zona waktu
    if (typeof v === 'number' && isFinite(v) && v > 20000 && v < 90000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
    }
    if (typeof v === 'string') {
      const s = v.trim(); if (!s) return null;
      let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
      if (m) return { y: +m[1], m: +m[2], d: +m[3] };
      m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
      if (m) {
        let d = +m[1], mo = +m[2], y = +m[3]; if (y < 100) y += 2000;
        if (mo === 1 && d !== 1) { const t = d; d = mo; mo = t; } // format mm/dd tertukar
        return { y, m: mo, d };
      }
      m = s.match(/^(\d{1,2})[\s\-]([A-Za-z]{3,9})[\s\-.](\d{2,4})$/);
      if (m) { const mo = monthFromName(m[2]); let y = +m[3]; if (y < 100) y += 2000; if (mo) return { y, m: mo, d: +m[1] }; }
      m = s.match(/^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})$/);
      if (m) { const mo = monthFromName(m[1]); if (mo) return { y: +m[3], m: mo, d: +m[2] }; }
      const d = new Date(s);
      if (!isNaN(d)) return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
    }
  } catch (e) { /* abaikan */ }
  return null;
}

function parseNum(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).replace(/rp/gi, '').replace(/[^0-9.,\-]/g, '').trim();
  if (!s || s === '-') return 0;
  const neg = s.startsWith('-'); s = s.replace(/-/g, '');
  const hasDot = s.includes('.'), hasCom = s.includes(',');
  if (hasDot && hasCom) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(/,/g, '.');
    else s = s.replace(/,/g, '');
  } else if (hasCom) {
    const n = (s.match(/,/g) || []).length;
    s = (n === 1 && /,\d{1,2}$/.test(s)) ? s.replace(',', '.') : s.replace(/,/g, '');
  } else if (hasDot) {
    const n = (s.match(/\./g) || []).length;
    s = (n === 1 && /\.\d{1,2}$/.test(s)) ? s : s.replace(/\./g, '');
  }
  const out = parseFloat(s);
  return isFinite(out) ? (neg ? -out : out) : 0;
}

/* ---------------- PARSER TABEL ---------------- */
function findHeaderRow(grid, kind) {
  let loose = -1;
  for (let i = 0; i < Math.min(grid.length, 15); i++) {
    const cells = (grid[i] || []).map(c => normName(c).toLowerCase());
    const j = cells.join('|');
    const filled = cells.filter(c => c).length;
    // Aturan ketat: baris header sejati (banyak sel terisi, bukan sel kalimat instruksi)
    if (filled >= 3) {
      if (kind === 'input' && cells.some(c => c.indexOf('periode') === 0 || c.indexOf('tanggal') === 0) &&
          cells.some(c => c.indexOf('outlet') >= 0 && c.indexOf('total') < 0)) return i;
      if (kind === 'master' && cells.some(c => c.indexOf('nama outlet') === 0) &&
          cells.some(c => c.indexOf('jenis') === 0 || c.indexOf('wilayah') === 0)) return i;
    }
    // Cadangan longgar
    if (loose < 0) {
      if (kind === 'input' && /periode|tanggal/.test(j) && /outlet/.test(j)) loose = i;
      else if (kind === 'master' && /nama outlet/.test(j) && /(jenis|wilayah)/.test(j)) loose = i;
    }
  }
  return loose;
}

function mapInputCols(cells) {
  const idx = {};
  cells.forEach((c, i) => {
    const t = normName(c).toLowerCase();
    if (idx.period == null && /periode|tanggal/.test(t)) idx.period = i;
    if (idx.outlet == null && /outlet/.test(t) && !/total/.test(t)) idx.outlet = i;
    if (idx.dineIn == null && /dine/.test(t)) idx.dineIn = i;
    if (idx.takeAway == null && /take/.test(t)) idx.takeAway = i;
    if (idx.goFood == null && /go\s*food/.test(t)) idx.goFood = i;
    if (idx.grabFood == null && /grab/.test(t)) idx.grabFood = i;
    if (idx.shopeeFood == null && /shopee/.test(t)) idx.shopeeFood = i;
    if (idx.esbOrder == null && /esb/.test(t)) idx.esbOrder = i;
    if (idx.total == null && /total/.test(t)) idx.total = i;
    if (idx.note == null && /catatan/.test(t)) idx.note = i;
  });
  return idx;
}

function parseGridRecords(grid) {
  const hr = findHeaderRow(grid, 'input');
  if (hr < 0) throw new Error('Baris header "Periode Bulan / Nama Outlet" tidak ditemukan di sheet INPUT PENJUALAN.');
  const idx = mapInputCols((grid[hr] || []).map(c => c == null ? '' : c));
  const recs = [];
  let skipped = 0;
  for (let r = hr + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const outlet = idx.outlet != null ? normName(row[idx.outlet]) : '';
    if (!outlet) continue;
    const p = idx.period != null ? parsePeriod(row[idx.period]) : null;
    if (!p) { skipped++; continue; }
    const cats = CFG.catKeys.map(k => idx[k] != null ? Math.max(0, parseNum(row[idx[k]])) : 0);
    let total = cats.reduce((a, b) => a + b, 0);
    if (total === 0 && idx.total != null) {
      const tv = Math.max(0, parseNum(row[idx.total]));
      if (tv > 0) total = tv;
    }
    recs.push({ outlet, y: p.y, m: p.m, d: p.d, cats, total, note: idx.note != null ? normName(row[idx.note]) : '' });
  }
  return { records: recs, skipped };
}

function parseGridOutlets(grid) {
  const hr = findHeaderRow(grid, 'master');
  if (hr < 0) return null;
  const cells = (grid[hr] || []).map(c => normName(c).toLowerCase());
  let iName = -1, iW = -1, iJ = -1;
  cells.forEach((t, i) => {
    if (iName < 0 && /nama outlet/.test(t)) iName = i;
    if (iW < 0 && /wilayah/.test(t)) iW = i;
    if (iJ < 0 && /jenis/.test(t)) iJ = i;
  });
  if (iName < 0) return null;
  const clean = x => { const s = normName(x); return (!s || s === '*' || s === '-') ? '' : s; };
  const map = {}, order = [];
  for (let r = hr + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const name = normName(row[iName]);
    if (!name || map[name]) continue;
    map[name] = { wilayah: iW >= 0 ? clean(row[iW]) : '', jenis: iJ >= 0 ? clean(row[iJ]) : '' };
    order.push(name);
  }
  return { map, order };
}

function finishParse(gridIn, master) {
  const { records, skipped } = parseGridRecords(gridIn);
  const outlets = Object.assign({}, master && master.map || {});
  const order = (master && master.order || []).slice();
  for (const r of records) {
    if (!outlets[r.outlet]) { outlets[r.outlet] = { wilayah: '', jenis: '' }; order.push(r.outlet); }
  }
  return { records, outlets, order, skipped };
}

function parseWorkbook(buf) {
  if (typeof XLSX === 'undefined') throw new Error('Library pembaca Excel gagal dimuat. Periksa koneksi internet lalu muat ulang halaman.');
  const wb = XLSX.read(buf, { type: 'array' }); // serial tanggal tetap angka: konversi manual di parsePeriod
  const nameIn = wb.SheetNames.find(n => /input/i.test(n));
  if (!nameIn) throw new Error('Sheet "INPUT PENJUALAN" tidak ditemukan dalam file.');
  const gridIn = XLSX.utils.sheet_to_json(wb.Sheets[nameIn], { header: 1, raw: true, defval: null });
  let master = null;
  const nameMa = wb.SheetNames.find(n => /master/i.test(n));
  if (nameMa) {
    const g = XLSX.utils.sheet_to_json(wb.Sheets[nameMa], { header: 1, raw: true, defval: null });
    master = parseGridOutlets(g);
  }
  return finishParse(gridIn, master);
}

function csvToGrid(text) {
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur); cur = ''; rows.push(row); row = [];
    } else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

/* ---------------- SUMBER DATA ---------------- */
function gsId(url) { const m = String(url || '').match(/\/d\/([a-zA-Z0-9-_]{15,})/); return m ? m[1] : null; }

async function fetchGSheets(url) {
  const id = gsId(url);
  if (!id) throw new Error('Link Google Sheets tidak valid. Salin link lengkap dari browser.');
  const base = 'https://docs.google.com/spreadsheets/d/' + id + '/gviz/tq?tqx=out:csv&sheet=';
  const [inRes, maRes] = await Promise.allSettled([
    fetch(base + encodeURIComponent('INPUT PENJUALAN'), { cache: 'no-store' }),
    fetch(base + encodeURIComponent('MASTER OUTLET'), { cache: 'no-store' })
  ]);
  if (inRes.status !== 'fulfilled' || !inRes.value.ok) {
    throw new Error('Tidak bisa membaca sheet "INPUT PENJUALAN". Pastikan file dibagikan "Anyone with the link" dan nama sheet tidak diubah.');
  }
  const gridIn = csvToGrid(await inRes.value.text());
  let master = null;
  if (maRes.status === 'fulfilled' && maRes.value.ok) {
    master = parseGridOutlets(csvToGrid(await maRes.value.text()));
  }
  return finishParse(gridIn, master);
}

function genDemo() {
  const rnd = (a => () => { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; })(20261109);
  const y = new Date().getFullYear(), nowM = new Date().getMonth() + 1;
  const WIL = ['Jakarta Pusat', 'Jakarta Selatan', 'Jakarta Barat', 'Tangerang', 'Bekasi'];
  const names = [], map = {};
  for (let i = 1; i <= 37; i++) {
    const n = 'Outlet ' + String(i).padStart(2, '0');
    names.push(n);
    map[n] = { wilayah: WIL[i % 5], jenis: i <= 3 ? 'Pusat' : 'Mitra' };
  }
  const recs = [];
  names.forEach((n, i) => {
    const scale = (i % 4 === 0 ? 2.1 : i % 3 === 0 ? 1.45 : 1) * (0.85 + rnd() * 0.4);
    const base = 95e6 * scale;
    for (let m = 1; m <= nowM; m++) {
      const grow = 1 + (m - 1) * 0.013 + (rnd() - 0.45) * 0.16;
      const total = Math.round(base * grow / 1e5) * 1e5;
      const w = [0.42, 0.18, 0.16, 0.12, 0.07, 0.05].map(x => x * (0.72 + rnd() * 0.55));
      const ws = w.reduce((a, b) => a + b, 0);
      const cats = w.map(x => Math.round(total * x / ws / 1e4) * 1e4);
      recs.push({ outlet: n, y, m, d: 1, cats, total: cats.reduce((a, b) => a + b, 0), note: '' });
    }
  });
  return { records: recs, outlets: map, order: names, skipped: 0 };
}

/* ---------------- MUAT DATA ---------------- */
async function loadAll(silent) {
  if (S.loading) return;
  S.loading = true; setBadge('loading');
  try {
    let parsed, label = '';
    if (S.mode === 'demo') { parsed = genDemo(); label = 'data contoh'; }
    else if (S.mode === 'gsheets') { parsed = await fetchGSheets(S.gsUrl); label = 'Google Sheets'; }
    else {
      const res = await fetch(cacheBust(S.dataUrl), { cache: 'no-store' });
      if (!res.ok) throw new Error('"' + S.dataUrl + '" tidak ditemukan (HTTP ' + res.status + ').');
      const ct = res.headers.get('content-type') || '';
      const isCsv = /\.csv($|\?)/i.test(S.dataUrl) || ct.includes('text/csv');
      if (isCsv) {
        const grid = csvToGrid(await res.text());
        parsed = finishParse(grid, null);
      } else {
        parsed = parseWorkbook(await res.arrayBuffer());
      }
      label = S.dataUrl;
    }
    applyData(parsed);
    S.lastUpdate = new Date(); S.demo = (S.mode === 'demo');
    updateStatus();
    if (!silent) toast('ok', 'Data dimuat: ' + S.records.length + ' baris input dari ' + label + '.');
  } catch (e) {
    console.error(e);
    setBadge('error');
    if (!silent) toast('err', 'Gagal memuat data: ' + e.message);
    if (!S.records.length) showEmpty(true);
  } finally {
    S.loading = false;
  }
}

function applyData(parsed, opts) {
  opts = opts || {};
  S.records = parsed.records || [];
  S.outlets = parsed.outlets || {};
  S.outletOrder = parsed.order || [];
  const years = [...new Set(S.records.map(r => r.y))].sort((a, b) => b - a);
  if (!years.length) years.push(new Date().getFullYear());
  if (S.year == null || !years.includes(S.year)) S.year = years[0];
  fillYearSelect(years);
  S.monMonth = null;
  renderAll();
  showEmpty(S.records.length === 0);
}

function handleFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) { toast('err', 'Format harus .xlsx, .xls, atau .csv'); return; }
  const rd = new FileReader();
  rd.onload = () => {
    try {
      let parsed;
      if (/\.csv$/i.test(file.name)) {
        const grid = csvToGrid(String(rd.result));
        parsed = finishParse(grid, null);
      } else {
        parsed = parseWorkbook(rd.result);
      }
      S.mode = 'upload'; S.demo = false;
      applyData(parsed);
      S.lastUpdate = new Date();
      updateStatus();
      toast('ok', 'File "' + file.name + '" dimuat: ' + S.records.length + ' baris (hanya tampilan lokal, tidak tersimpan).');
    } catch (e) {
      console.error(e);
      toast('err', 'Gagal membaca file: ' + e.message);
    }
  };
  rd.onerror = () => toast('err', 'File tidak bisa dibaca.');
  if (/\.csv$/i.test(file.name)) rd.readAsText(file); else rd.readAsArrayBuffer(file);
}

/* ---------------- AGREGASI ---------------- */
function computeView() {
  const y = S.year;
  const recs = S.records.filter(r => r.y === y);
  const months = Array.from({ length: 12 }, () => ({ total: 0, cats: [0, 0, 0, 0, 0, 0], rows: 0, outlets: new Map() }));
  const om = {};
  const catTotals = [0, 0, 0, 0, 0, 0];
  let yearTotal = 0;
  for (const r of recs) {
    const M = months[r.m - 1];
    M.total += r.total; M.rows++;
    r.cats.forEach((v, i) => { M.cats[i] += v; catTotals[i] += v; });
    const c = M.outlets.get(r.outlet) || 0; M.outlets.set(r.outlet, c + 1);
    yearTotal += r.total;
    const o = om[r.outlet] || (om[r.outlet] = { total: 0, cats: [0, 0, 0, 0, 0, 0], lastM: 0 });
    o.total += r.total;
    r.cats.forEach((v, i) => o.cats[i] += v);
    if (r.m > o.lastM) o.lastM = r.m;
  }
  const dataMonths = months.map((M, i) => M.rows > 0 ? i + 1 : 0).filter(Boolean);
  const lastM = dataMonths.length ? Math.max.apply(null, dataMonths) : 0;
  const today = new Date();
  const isCurYear = y === today.getFullYear();
  const refM = isCurYear ? today.getMonth() + 1 : (lastM || 1);
  const curTotal = months[refM - 1].total;
  const prevM = refM > 1 ? refM - 1 : 0;
  const prevTotal = prevM ? months[prevM - 1].total : 0;
  // Growth hanya dihitung bila kedua bulan memiliki data
  const growth = (curTotal > 0 && prevTotal > 0) ? (curTotal / prevTotal - 1) : null;

  let pu = 0, mi = 0, nPu = 0, nMi = 0, nUn = 0;
  for (const name of Object.keys(om)) {
    const j = ((S.outlets[name] && S.outlets[name].jenis) || '').toLowerCase();
    if (j.indexOf('pusat') === 0) { pu += om[name].total; nPu++; }
    else if (j.indexOf('mitra') === 0) { mi += om[name].total; nMi++; }
    else nUn++;
  }

  const rank = Object.keys(om).map(name => ({
    name, total: om[name].total, cats: om[name].cats, lastM: om[name].lastM,
    meta: S.outlets[name] || { wilayah: '', jenis: '' }
  })).sort((a, b) => b.total - a.total);
  for (const n of S.outletOrder) {
    if (!om[n]) rank.push({ name: n, total: 0, cats: [0, 0, 0, 0, 0, 0], lastM: 0, meta: S.outlets[n] || { wilayah: '', jenis: '' } });
  }

  const view = { y, recs, months, catTotals, yearTotal, om, rank, dataMonths, lastM, refM, prevM, curTotal, prevTotal, growth, pu, mi, nPu, nMi, nUn, isCurYear };
  S.view = view;
  return view;
}

/* ---------------- RENDER ---------------- */
function animNum(el, to, fmt) {
  const from = parseFloat(el.dataset.v || '0');
  el.dataset.v = to;
  const reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
  if (reduce || from === to) { el.innerHTML = fmt(to); return; }
  const t0 = performance.now(), dur = 750;
  (function step(t) {
    const p = Math.min((t - t0) / dur, 1);
    const e = 1 - Math.pow(1 - p, 3);
    el.innerHTML = fmt(from + (to - from) * e);
    if (p < 1) requestAnimationFrame(step);
  })(t0);
}

function growthBadge(g) {
  if (g == null) return '<span class="badge badge-flat">—</span>';
  if (Math.abs(g) < 0.0005) return '<span class="badge badge-flat">0%</span>';
  return g > 0
    ? '<span class="badge badge-up">&#8593; ' + fmtPct(g) + '</span>'
    : '<span class="badge badge-down">&#8595; ' + fmtPct(Math.abs(g)) + '</span>';
}

function bigIdx(cats) {
  let bi = -1, bv = 0;
  for (let i = 0; i < cats.length; i++) if (cats[i] > bv) { bv = cats[i]; bi = i; }
  return bi;
}

function fillYearSelect(years) {
  const sel = $('#yearSelect');
  sel.innerHTML = years.map(y => '<option value="' + y + '">' + y + '</option>').join('');
  sel.value = String(S.year);
}

function fillMonthSelect() {
  const sel = $('#monMonth');
  sel.innerHTML = CFG.months.map((m, i) => '<option value="' + (i + 1) + '">' + m + '</option>').join('');
}

function renderAll() {
  const v = computeView();
  renderSummary(v);
  renderRank();
  renderMonitoring();
  renderBulanan(v);
  const notice = $('#yearNotice');
  if (S.records.length > 0 && v.recs.length === 0) {
    notice.hidden = false;
    notice.innerHTML = 'Tidak ada data input untuk tahun <b>' + v.y + '</b>. Data tersedia untuk tahun: <b>' +
      [...new Set(S.records.map(r => r.y))].sort().reverse().join(', ') + '</b>.';
  } else notice.hidden = true;
}

function renderSummary(v) {
  $('#kpiYearLabel').textContent = v.y;
  $('#kpiYearLabel2').textContent = v.y;
  $('#chMonthlyYear').textContent = '· ' + v.y;
  $('#rankYear').textContent = '· ' + v.y;
  $('#chTrendYear').textContent = '· ' + v.y;
  $('#trendYear').textContent = '· ' + v.y;

  animNum($('#kpiTotal'), v.yearTotal, fmtRp);
  const activeOutlets = Object.keys(v.om).length;
  $('#kpiTotalSub').textContent = v.dataMonths.length
    ? v.dataMonths.length + ' bulan berdata · ' + activeOutlets + ' outlet aktif'
    : 'Belum ada data pada tahun ini';

  const today = new Date();
  const isRunningMonth = v.isCurYear && v.refM === today.getMonth() + 1;
  $('#kpiMonthLabel').textContent = isRunningMonth ? 'Omset Bulan Berjalan' : 'Omset ' + CFG.months[v.refM - 1];
  animNum($('#kpiMonth'), v.curTotal, fmtRp);
  $('#kpiMonthSub').textContent = v.months[v.refM - 1].rows
    ? CFG.months[v.refM - 1] + ' ' + v.y + (isRunningMonth ? ' · periode sampai tanggal ' + today.getDate() : '')
    : 'Belum ada laporan untuk ' + CFG.months[v.refM - 1];

  $('#kpiGrowthLabel').textContent = 'Growth vs Bulan Lalu';
  const gEl = $('#kpiGrowth');
  if (v.growth == null) { gEl.dataset.v = '0'; gEl.textContent = '—'; }
  else animNum(gEl, v.growth, g => g >= 0
    ? '<span class="trend-up">&#8593; ' + fmtPct(g) + '</span>'
    : '<span class="trend-down">&#8595; ' + fmtPct(Math.abs(g)) + '</span>');
  $('#kpiGrowthSub').textContent = v.prevM
    ? (v.curTotal === 0 && v.prevTotal > 0
        ? 'Belum ada laporan untuk ' + CFG.months[v.refM - 1]
        : CFG.monthsShort[v.refM - 1] + ' vs ' + CFG.monthsShort[v.prevM - 1] + (v.prevTotal === 0 ? ' · bulan lalu belum berdata' : ''))
    : 'Januari belum punya pembanding';

  animNum($('#kpiReports'), v.recs.length, x => nfNum.format(Math.round(x)));
  const totalOut = S.outletOrder.length || Math.max(37, activeOutlets);
  $('#kpiReportsSub').textContent = CFG.months[v.refM - 1] + ': ' + v.months[v.refM - 1].outlets.size + '/' + totalOut + ' outlet melapor';

  if (S.charts.monthly) {
    S.charts.monthly.data.datasets[0].data = v.months.map(M => M.total);
    S.charts.monthly.update();
  }

  if (S.charts.cats) {
    S.charts.cats.data.datasets[0].data = v.catTotals.slice();
    S.charts.cats.options.plugins.centerText.value = v.yearTotal > 0 ? fmtShortN(v.yearTotal) : '—';
    S.charts.cats.update();
  }
  const sum = v.yearTotal || 1;
  $('#catLegend').innerHTML = CFG.cats.map((c, i) => {
    const val = v.catTotals[i];
    return '<span class="lg-item"><span class="lg-dot" style="background:' + CFG.catColors[i] + '"></span>' +
      esc(c) + ' <b>' + (val > 0 ? fmtPct(val / sum) : '0%') + '</b></span>';
  }).join('');

  // Pusat vs Mitra
  const pmTotal = v.pu + v.mi;
  $('#pmPusat').textContent = fmtRp(v.pu);
  $('#pmMitra').textContent = fmtRp(v.mi);
  $('#pmPusatN').textContent = v.nPu;
  $('#pmMitraN').textContent = v.nMi;
  $('#pmPusatPct').textContent = pmTotal > 0 ? fmtPct(v.pu / pmTotal, 0) : '0%';
  $('#pmMitraPct').textContent = pmTotal > 0 ? fmtPct(v.mi / pmTotal, 0) : '0%';
  $('#pmFillPusat').style.width = pmTotal > 0 ? (v.pu / pmTotal * 100) + '%' : '50%';
  $('#pmFillMitra').style.width = pmTotal > 0 ? (v.mi / pmTotal * 100) + '%' : '50%';
  let note = 'Klasifikasi dari kolom "Jenis Outlet" di MASTER OUTLET';
  if (v.nUn > 0) note += ' · ' + v.nUn + ' outlet belum diklasifikasi';
  $('#pmNote').textContent = note;
}

function renderRank() {
  const v = S.view; if (!v) return;
  const q = (S.search || '').toLowerCase();
  const max = v.rank.length ? v.rank[0].total : 0;
  const rows = v.rank.filter(r => !q ||
    r.name.toLowerCase().includes(q) ||
    (r.meta.wilayah || '').toLowerCase().includes(q) ||
    (r.meta.jenis || '').toLowerCase().includes(q));
  $('#rankCount').textContent = rows.length + ' outlet';
  const tbody = $('#rankBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="cell-muted" style="text-align:center;padding:26px">Tidak ada outlet yang cocok dengan pencarian.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((r, i) => {
    const rk = i + 1;
    const cls = r.total > 0 && rk <= 3 ? ' class="rank-badge rank-' + rk + '"' : ' class="rank-badge"';
    const pct = v.yearTotal > 0 ? r.total / v.yearTotal : 0;
    const bar = max > 0 ? r.total / max * 100 : 0;
    const bi = bigIdx(r.cats);
    const chip = bi < 0
      ? '<span class="cell-muted">—</span>'
      : '<span class="cat-chip" style="color:' + CFG.catColors[bi] + ';border-color:' + CFG.catColors[bi] + '55;background:' + CFG.catColors[bi] + '14">' + esc(CFG.cats[bi]) + '</span>';
    return '<tr>' +
      '<td class="t-center"><span' + cls + '>' + (r.total > 0 ? rk : '–') + '</span></td>' +
      '<td><div class="outlet-name">' + esc(r.name) + '</div></td>' +
      '<td class="cell-muted">' + esc(r.meta.wilayah || '—') + '</td>' +
      '<td class="cell-muted">' + esc(r.meta.jenis || '—') + '</td>' +
      '<td class="t-right num">' + (r.total > 0 ? fmtRp(r.total) : '<span class="cell-muted">belum ada data</span>') + '</td>' +
      '<td class="t-wide"><div class="contrib-wrap"><div class="contrib-bar"><div class="contrib-fill" style="width:' + bar + '%"></div></div><span class="contrib-pct">' + (r.total > 0 ? fmtPct(pct) : '0%') + '</span></div></td>' +
      '<td>' + chip + '</td></tr>';
  }).join('');
  if (S.charts.top) {
    const top = rows.filter(r => r.total > 0).slice(0, 8);
    S.charts.top.data.labels = top.map(r => r.name);
    S.charts.top.data.datasets[0].data = top.map(r => r.total);
    S.charts.top.update();
  }
}

function renderMonitoring() {
  const v = S.view; if (!v) return;
  const m = S.monMonth || v.refM;
  $('#monMonth').value = String(m);
  const M = v.months[m - 1];
  const totalOut = S.outletOrder.length || 37;
  const uniq = M.outlets.size;
  const dups = [...M.outlets.keys()].filter(n => M.outlets.get(n) > 1);
  const missing = S.outletOrder.filter(n => !M.outlets.has(n));
  $('#monChip').textContent = uniq + '/' + totalOut;
  $('#monBar').style.width = (totalOut ? uniq / totalOut * 100 : 0) + '%';
  let info = M.rows + ' baris laporan untuk ' + CFG.months[m - 1] + ' ' + v.y + '.';
  if (dups.length) info += ' ' + dups.length + ' outlet tercatat lebih dari satu kali pada bulan ini.';
  $('#monInfo').textContent = info;
  let html = dups.map(n => '<span class="dup-chip">DUPLIKAT: ' + esc(n) + '</span>').join('');
  html += missing.slice(0, 10).map(n => '<span class="miss-chip">' + esc(n) + '</span>').join('');
  if (missing.length > 10) html += '<span class="miss-chip">+' + (missing.length - 10) + ' outlet lain</span>';
  if (!missing.length && !dups.length && totalOut > 0) html = '<span class="badge badge-up">Semua outlet sudah melapor</span>';
  $('#monMissing').innerHTML = html;
}

function renderBulanan(v) {
  const isRunning = v.isCurYear && v.refM === new Date().getMonth() + 1;
  $('#momCurLabel').textContent = isRunning ? 'Omset Bulan Berjalan' : 'Omset ' + CFG.months[v.refM - 1];
  animNum($('#momCur'), v.curTotal, fmtRp);
  $('#momCurSub').textContent = CFG.months[v.refM - 1] + ' ' + v.y + (v.months[v.refM - 1].rows ? '' : ' · belum ada laporan');
  $('#momPrevLabel').textContent = v.prevM ? 'Omset ' + CFG.months[v.prevM - 1] : 'Bulan Sebelumnya';
  animNum($('#momPrev'), v.prevTotal, fmtRp);
  $('#momPrevSub').textContent = v.prevM ? CFG.months[v.prevM - 1] + ' ' + v.y : '—';
  const gEl = $('#momGrowth');
  if (v.growth == null) { gEl.dataset.v = '0'; gEl.textContent = '—'; }
  else animNum(gEl, v.growth, g => g >= 0
    ? '<span class="trend-up">&#8593; ' + fmtPct(g) + '</span>'
    : '<span class="trend-down">&#8595; ' + fmtPct(Math.abs(g)) + '</span>');
  $('#momGrowthSub').textContent = (v.prevM && v.growth != null) ? 'Selisih: ' + fmtRp(v.curTotal - v.prevTotal) : 'Selisih: —';

  if (S.charts.trend) {
    const n = Math.max(v.lastM, 1);
    S.charts.trend.data.labels = CFG.monthsShort.slice(0, n);
    S.charts.trend.data.datasets.forEach((ds, i) => { ds.data = v.months.map(M => M.cats[i]).slice(0, n); });
    S.charts.trend.update();
  }

  const tbody = $('#trendBody');
  tbody.innerHTML = CFG.months.map((name, i) => {
    const M = v.months[i];
    const g = i > 0 && v.months[i - 1].total > 0 ? (M.total / v.months[i - 1].total - 1) : null;
    return '<tr>' +
      '<td class="outlet-name">' + name + '</td>' +
      '<td class="t-right num">' + (M.total > 0 ? fmtRp(M.total) : '<span class="cell-muted">—</span>') + '</td>' +
      '<td class="t-center">' + growthBadge(g) + '</td>' +
      M.cats.map(cval => '<td class="t-right num cell-muted">' + (cval > 0 ? fmtShortN(cval) : '—') + '</td>').join('') +
      '</tr>';
  }).join('');
}

/* ---------------- CHARTS ---------------- */
const centerTextPlugin = {
  id: 'centerText',
  afterDraw(c) {
    const o = c.config.options.plugins && c.config.options.plugins.centerText;
    if (!o) return;
    const a = c.chartArea; if (!a) return;
    const x = (a.left + a.right) / 2, y = (a.top + a.bottom) / 2;
    const ctx = c.ctx;
    ctx.save();
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.font = '600 10px Inter, sans-serif';
    ctx.fillStyle = 'rgba(139,150,173,.95)';
    ctx.fillText(o.title || 'TOTAL', x, y - 12);
    ctx.font = '800 17px Sora, Inter, sans-serif';
    ctx.fillStyle = '#e8edf7';
    ctx.fillText(o.value || '—', x, y + 9);
    ctx.restore();
  }
};

function gradBarV(ctx) {
  const a = ctx.chart.chartArea; if (!a) return 'rgba(34,211,238,.75)';
  const g = ctx.chart.ctx.createLinearGradient(0, a.bottom, 0, a.top);
  g.addColorStop(0, 'rgba(34,211,238,.08)');
  g.addColorStop(1, 'rgba(34,211,238,.88)');
  return g;
}
function gradBarH(ctx) {
  const a = ctx.chart.chartArea; if (!a) return 'rgba(139,92,246,.75)';
  const g = ctx.chart.ctx.createLinearGradient(a.left, 0, a.right, 0);
  g.addColorStop(0, 'rgba(139,92,246,.85)');
  g.addColorStop(1, 'rgba(34,211,238,.75)');
  return g;
}

function setupCharts() {
  if (typeof Chart === 'undefined') {
    toast('err', 'Library grafik (Chart.js) gagal dimuat. Periksa koneksi internet lalu muat ulang.');
    return;
  }
  Chart.defaults.color = 'rgba(139,150,173,.9)';
  Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
  Chart.defaults.font.size = 11;
  const tip = {
    backgroundColor: 'rgba(8,12,26,.95)',
    borderColor: 'rgba(255,255,255,.14)', borderWidth: 1,
    titleFont: { family: "'Sora', sans-serif", weight: '700', size: 12 },
    padding: 12, cornerRadius: 10, boxPadding: 6
  };

  S.charts.monthly = new Chart($('#chMonthly'), {
    type: 'bar',
    data: {
      labels: CFG.monthsShort.slice(),
      datasets: [{
        data: Array(12).fill(0),
        backgroundColor: gradBarV,
        hoverBackgroundColor: 'rgba(129,230,251,.95)',
        borderRadius: 9, borderSkipped: false, maxBarThickness: 46
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: Object.assign({}, tip, { callbacks: { label: c => ' ' + fmtRp(c.parsed.y) } })
      },
      scales: {
        x: { grid: { display: false }, border: { display: false } },
        y: {
          beginAtZero: true, border: { display: false },
          grid: { color: 'rgba(148,163,184,.07)' },
          ticks: { callback: v => fmtShortN(v), maxTicksLimit: 6 }
        }
      }
    }
  });

  S.charts.cats = new Chart($('#chCats'), {
    type: 'doughnut',
    data: {
      labels: CFG.cats.slice(),
      datasets: [{
        data: Array(6).fill(0),
        backgroundColor: CFG.catColors.slice(),
        borderColor: 'rgba(5,8,16,.85)', borderWidth: 3, hoverOffset: 10
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '68%',
      plugins: {
        legend: { display: false },
        centerText: { title: 'TOTAL ' + (S.year || ''), value: '—' },
        tooltip: Object.assign({}, tip, {
          callbacks: {
            label: c => {
              const arr = c.dataset.data || [];
              const sum = arr.reduce((a, b) => a + b, 0);
              const pct = sum > 0 ? ' (' + fmtPct(c.parsed / sum) + ')' : '';
              return ' ' + c.label + ': ' + fmtRp(c.parsed) + pct;
            }
          }
        })
      }
    },
    plugins: [centerTextPlugin]
  });

  S.charts.top = new Chart($('#chTop'), {
    type: 'bar',
    data: { labels: [], datasets: [{ data: [], backgroundColor: gradBarH, hoverBackgroundColor: 'rgba(196,181,253,.95)', borderRadius: 7, borderSkipped: false, maxBarThickness: 24 }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: Object.assign({}, tip, { callbacks: { label: c => ' ' + fmtRp(c.parsed.x) } })
      },
      scales: {
        x: { beginAtZero: true, border: { display: false }, grid: { color: 'rgba(148,163,184,.07)' }, ticks: { callback: v => fmtShortN(v), maxTicksLimit: 6 } },
        y: { grid: { display: false }, border: { display: false }, ticks: { font: { size: 10 } } }
      }
    }
  });

  S.charts.trend = new Chart($('#chTrend'), {
    type: 'line',
    data: {
      labels: CFG.monthsShort.slice(),
      datasets: CFG.cats.map((c, i) => ({
        label: c, data: Array(12).fill(0),
        borderColor: CFG.catColors[i], backgroundColor: 'transparent',
        borderWidth: 2, tension: .38, pointRadius: 2.5, pointHoverRadius: 6,
        pointBackgroundColor: CFG.catColors[i], pointBorderColor: 'rgba(5,8,16,.9)', pointBorderWidth: 1
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'line', boxWidth: 26, boxHeight: 2, padding: 14, font: { size: 10.5 } } },
        tooltip: Object.assign({}, tip, { callbacks: { label: c => ' ' + c.dataset.label + ': ' + fmtRp(c.parsed.y) } })
      },
      scales: {
        x: { grid: { display: false }, border: { display: false } },
        y: {
          beginAtZero: true, border: { display: false },
          grid: { color: 'rgba(148,163,184,.07)' },
          ticks: { callback: v => fmtShortN(v), maxTicksLimit: 6 }
        }
      }
    }
  });
}

/* ---------------- STATUS & UI KECIL ---------------- */
function setBadge(state) {
  const b = $('#liveBadge');
  b.classList.remove('is-live', 'is-loading', 'is-error', 'is-demo', 'is-upload');
  b.classList.add('is-' + state);
  const map = { loading: 'MEMUAT', error: 'OFFLINE', demo: 'MODE CONTOH', upload: 'FILE LOKAL', live: 'LIVE' };
  $('#liveText').textContent = map[state] || 'LIVE';
}

function updateStatus() {
  if (S.mode === 'demo') setBadge('demo');
  else if (S.mode === 'upload') setBadge('upload');
  else if (S.lastUpdate) { setBadge('live'); $('#liveText').textContent = 'LIVE · ' + timeStr(S.lastUpdate); }
  else setBadge('error');

  $('#srcLabel').textContent = 'Sumber data: ' + (
    S.mode === 'github' ? S.dataUrl :
    S.mode === 'gsheets' ? 'Google Sheets' :
    S.mode === 'demo' ? 'Data contoh' : 'Unggahan lokal');
  $('#footLast').textContent = 'Terakhir dimuat: ' + (S.lastUpdate ? timeStr(S.lastUpdate) : '—');
  $('#statRows').textContent = S.records.length;
  $('#statLast').textContent = S.lastUpdate ? timeStr(S.lastUpdate) : '—';
  $('#demoBanner').hidden = !S.demo;
}

function showEmpty(show) {
  $('#emptyState').hidden = !show;
  $('#mainContent').querySelectorAll('.view').forEach(v => {
    const on = !show && v.id === 'view-' + currentView;
    v.classList.toggle('active', on);
    v.hidden = !on;
  });
  $('#yearSelect').disabled = show;
}

let currentView = 'ringkasan';
function switchView(name) {
  currentView = name;
  document.querySelectorAll('.tab').forEach(t => {
    const on = t.dataset.view === name;
    t.classList.toggle('active', on);
    t.setAttribute('aria-selected', on ? 'true' : 'false');
  });
  document.querySelectorAll('.view').forEach(v => {
    const on = v.id === 'view-' + name;
    v.classList.toggle('active', on);
    v.hidden = !on;
  });
}

function toast(type, msg) {
  const box = $('#toasts');
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.textContent = msg;
  box.appendChild(el);
  setTimeout(() => { el.classList.add('hide'); setTimeout(() => el.remove(), 400); }, 5200);
}

/* ---------------- PENGATURAN ---------------- */
function saveSettings() {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify({
      mode: (S.mode === 'upload' || S.mode === 'demo') ? 'github' : S.mode,
      dataUrl: S.dataUrl, gsUrl: S.gsUrl
    }));
  } catch (e) { /* localStorage bisa gagal di beberapa browser */ }
}
function loadSaved() {
  try {
    const o = JSON.parse(localStorage.getItem(LS_KEY) || 'null');
    if (o && o.mode) return o;
  } catch (e) { /* abaikan */ }
  return { mode: 'github', dataUrl: CFG.defaultDataUrl, gsUrl: '' };
}

function openModal() {
  syncSettingsUI();
  $('#modal').hidden = false;
}
function closeModal() { $('#modal').hidden = true; }

function syncSettingsUI() {
  const mode = ['github', 'gsheets', 'upload'].includes(S.mode) ? S.mode : 'github';
  document.querySelectorAll('input[name="dataMode"]').forEach(r => { r.checked = r.value === mode; });
  document.querySelectorAll('.mode-item').forEach(l => {
    const r = l.querySelector('input');
    l.classList.toggle('sel', r.checked);
    l.style.opacity = r.checked ? '1' : '.62';
  });
  $('#inputDataUrl').value = S.dataUrl;
  $('#inputGsUrl').value = S.gsUrl;
}

function applySettingsFromModal() {
  const checked = document.querySelector('input[name="dataMode"]:checked');
  const mode = checked ? checked.value : 'github';
  const url = $('#inputDataUrl').value.trim() || CFG.defaultDataUrl;
  const gs = $('#inputGsUrl').value.trim();
  if (mode === 'gsheets' && !gsId(gs)) {
    toast('err', 'Link Google Sheets belum benar. Contoh: https://docs.google.com/spreadsheets/d/xxxx/edit');
    return;
  }
  S.mode = mode; S.dataUrl = url; S.gsUrl = gs;
  saveSettings();
  closeModal();
  loadAll();
}

function enterDemo() {
  S.mode = 'demo'; S.demo = true;
  const parsed = genDemo();
  applyData(parsed);
  S.lastUpdate = new Date();
  updateStatus();
  toast('info', 'Data contoh dimuat — angka hanya ilustrasi tampilan.');
}
function exitDemo() {
  const st = loadSaved();
  S.mode = st.mode; S.dataUrl = st.dataUrl || CFG.defaultDataUrl; S.gsUrl = st.gsUrl || '';
  syncSettingsUI();
  loadAll();
}

/* ---------------- EVENT ---------------- */
function bindEvents() {
  document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => switchView(t.dataset.view)));

  $('#yearSelect').addEventListener('change', e => { S.year = +e.target.value; renderAll(); });
  $('#monMonth').addEventListener('change', e => { S.monMonth = +e.target.value; renderMonitoring(); });
  $('#outletSearch').addEventListener('input', debounce(e => { S.search = e.target.value.trim(); renderRank(); }, 160));

  $('#btnSettings').addEventListener('click', openModal);
  $('#btnEmptySettings').addEventListener('click', openModal);
  $('#btnCloseModal').addEventListener('click', closeModal);
  $('#modal').addEventListener('click', e => { if (e.target === $('#modal')) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

  document.querySelectorAll('input[name="dataMode"]').forEach(r =>
    r.addEventListener('change', () => { document.querySelectorAll('.mode-item').forEach(l => {
      const rr = l.querySelector('input');
      l.classList.toggle('sel', rr.checked);
      l.style.opacity = rr.checked ? '1' : '.62';
    }); }));

  $('#btnSaveSettings').addEventListener('click', applySettingsFromModal);
  $('#btnLoadNow').addEventListener('click', () => loadAll());
  $('#btnPickFile').addEventListener('click', () => $('#fileInput').click());
  $('#fileInput').addEventListener('change', e => { handleFile(e.target.files[0]); e.target.value = ''; });

  $('#btnDemo').addEventListener('click', enterDemo);
  $('#btnExitDemo').addEventListener('click', exitDemo);

  // Drag & drop file ke seluruh halaman
  let dragDepth = 0;
  window.addEventListener('dragenter', e => { e.preventDefault(); dragDepth++; $('#dropHint').hidden = false; });
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('dragleave', e => { e.preventDefault(); dragDepth--; if (dragDepth <= 0) { dragDepth = 0; $('#dropHint').hidden = true; } });
  window.addEventListener('drop', e => {
    e.preventDefault(); dragDepth = 0; $('#dropHint').hidden = true;
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
}

/* ---------------- INIT ---------------- */
function init() {
  bindEvents();
  fillMonthSelect();
  setupCharts();

  const st = loadSaved();
  S.mode = st.mode; S.dataUrl = st.dataUrl || CFG.defaultDataUrl; S.gsUrl = st.gsUrl || '';
  const qs = new URLSearchParams(location.search);
  if (qs.get('src')) { S.mode = 'github'; S.dataUrl = qs.get('src'); }

  syncSettingsUI();
  showEmpty(true);
  setBadge('loading');
  loadAll();

  setInterval(() => {
    if ((S.mode === 'github' || S.mode === 'gsheets') && !document.hidden) loadAll(true);
  }, CFG.refreshMs);

  document.addEventListener('visibilitychange', () => {
    if (!document.hidden && S.lastUpdate && (S.mode === 'github' || S.mode === 'gsheets') &&
        Date.now() - S.lastUpdate > CFG.refreshMs) loadAll(true);
  });
}

document.addEventListener('DOMContentLoaded', init);
