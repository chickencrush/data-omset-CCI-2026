/* ==========================================================
   Sales Command Center — 37 Outlet
   ========================================================== */

const CATS = [
  { key: 'dineIn',     label: 'Dine in',    color: '#00d9ff' },
  { key: 'takeAway',   label: 'Take away',  color: '#a855f7' },
  { key: 'goFood',     label: 'GoFood',     color: '#00ffa3' },
  { key: 'grabFood',   label: 'GrabFood',   color: '#ff8c42' },
  { key: 'shopeeFood', label: 'ShopeeFood', color: '#ff4d6d' },
  { key: 'esbOrder',   label: 'ESB Order',  color: '#ffd23f' }
];
const MONTHS = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];

const state = {
  outlets: [],
  penjualan: [],
  year: new Date().getFullYear(),
  workbook: null,
  charts: {}
};

/* ============ UTILITIES ============ */
const N = v => Number(v) || 0;
const total = p => CATS.reduce((s, c) => s + N(p[c.key]), 0);
const fmtRp = n => 'Rp ' + Math.round(N(n)).toLocaleString('id-ID');
const fmtShort = n => {
  n = N(n);
  if (n >= 1e9) return 'Rp ' + (n / 1e9).toFixed(1) + ' M';
  if (n >= 1e6) return 'Rp ' + (n / 1e6).toFixed(1) + ' jt';
  if (n >= 1e3) return 'Rp ' + (n / 1e3).toFixed(0) + ' rb';
  return 'Rp ' + n;
};
const fmtPct = n => N(n).toFixed(1) + '%';
const pad = n => String(n).padStart(2, '0');
const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') {
    try { const d = XLSX.SSF.parse_date_code(v); return new Date(d.y, d.m - 1, d.d); }
    catch { return null; }
  }
  if (typeof v === 'string') {
    const d = new Date(v);
    return isNaN(d) ? null : d;
  }
  return null;
}

function toast(msg, type = 'info') {
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.textContent = msg;
  document.body.appendChild(el);
  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(() => {
    el.classList.remove('show');
    setTimeout(() => el.remove(), 350);
  }, 3200);
}

/* ============ AGGREGATIONS ============ */
const yearRows = () => state.penjualan.filter(p => p.tanggal && p.tanggal.getFullYear() === state.year);
const sumRows = rows => rows.reduce((s, p) => s + total(p), 0);

function sumCat(rows) {
  const o = {};
  CATS.forEach(c => o[c.key] = 0);
  rows.forEach(p => CATS.forEach(c => o[c.key] += N(p[c.key])));
  return o;
}

function sumByOutlet(rows) {
  const o = {};
  rows.forEach(p => { o[p.outlet] = (o[p.outlet] || 0) + total(p); });
  return o;
}

/* ============ PARSING ============ */
function parseWorkbook(wb) {
  const outlets = [];
  const penjualan = [];

  // MASTER OUTLET
  const wsM = wb.Sheets['MASTER OUTLET'];
  if (wsM) {
    const rows = XLSX.utils.sheet_to_json(wsM, { header: 1, defval: '' });
    for (let i = 3; i < rows.length; i++) {
      const r = rows[i];
      if (!r[1]) continue;
      outlets.push({
        no: r[0],
        nama: String(r[1]).trim(),
        wilayah: r[2] || '',
        jenis: String(r[3] || '').trim()
      });
      if (outlets.length >= 37) break;
    }
  }

  // INPUT PENJUALAN
  const wsP = wb.Sheets['INPUT PENJUALAN'];
  if (wsP) {
    const rows = XLSX.utils.sheet_to_json(wsP, { header: 1, defval: '' });
    for (let i = 4; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0] || !r[1]) continue;
      const tgl = toDate(r[0]);
      if (!tgl) continue;
      penjualan.push({
        tanggal: tgl,
        outlet: String(r[1]).trim(),
        dineIn: N(r[2]),
        takeAway: N(r[3]),
        goFood: N(r[4]),
        grabFood: N(r[5]),
        shopeeFood: N(r[6]),
        esbOrder: N(r[7]),
        catatan: r[9] || ''
      });
    }
  }

  return { outlets, penjualan };
}

/* ============ RENDER ============ */
function renderAll() {
  populateYearSelect();
  renderKPIs();
  renderCharts();
  renderPusatMitra();
  renderRanking();
  renderGrowth();
  renderMonitoring();
  renderDataTable();
}

function populateYearSelect() {
  const sel = document.getElementById('yearSelect');
  const years = [...new Set(state.penjualan.map(p => p.tanggal.getFullYear()))].sort((a, b) => b - a);
  if (!years.length) years.push(new Date().getFullYear());
  if (!years.includes(state.year)) state.year = years[0];
  sel.innerHTML = years.map(y => `<option value="${y}">${y}</option>`).join('');
  sel.value = state.year;
}

function renderKPIs() {
  const yr = yearRows();
  const totalAll = sumRows(yr);

  const now = new Date();
  const cm = now.getMonth(), cy = now.getFullYear();
  const curM = state.penjualan.filter(p => p.tanggal.getFullYear() === cy && p.tanggal.getMonth() === cm);
  const lastM = state.penjualan.filter(p => p.tanggal.getFullYear() === cy && p.tanggal.getMonth() === cm - 1);
  const curTotal = sumRows(curM);
  const lastTotal = sumRows(lastM);
  const growth = lastTotal > 0 ? (curTotal / lastTotal - 1) * 100 : 0;

  document.getElementById('kpiTotal').textContent = fmtRp(totalAll);
  document.getElementById('kpiTotalSub').textContent = `${yr.length} baris · tahun ${state.year}`;
  document.getElementById('kpiTrx').textContent = yr.length.toLocaleString('id-ID');
  document.getElementById('kpiMonth').textContent = fmtRp(curTotal);
  document.getElementById('kpiMonthSub').textContent = `${MONTHS[cm]} ${cy}`;

  const growthEl = document.getElementById('kpiGrowth');
  growthEl.textContent = (growth >= 0 ? '+' : '') + fmtPct(growth);
  growthEl.style.color = growth >= 0 ? 'var(--good)' : 'var(--bad)';
  document.getElementById('kpiGrowthSub').textContent = `vs ${MONTHS[(cm - 1 + 12) % 12]}`;
}

function renderCharts() {
  // Kategori
  const cat = sumCat(yearRows());
  if (state.charts.kategori) state.charts.kategori.destroy();
  state.charts.kategori = new Chart(document.getElementById('chartKategori'), {
    type: 'doughnut',
    data: {
      labels: CATS.map(c => c.label),
      datasets: [{
        data: CATS.map(c => cat[c.key]),
        backgroundColor: CATS.map(c => c.color),
        borderWidth: 0,
        hoverOffset: 10
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '68%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: { color: '#8ba3c7', padding: 14, font: { size: 11, family: 'Inter' }, boxWidth: 12, boxHeight: 12, usePointStyle: true, pointStyle: 'circle' }
        },
        tooltip: {
          backgroundColor: 'rgba(10,14,31,0.95)',
          borderColor: 'rgba(255,255,255,0.1)',
          borderWidth: 1,
          padding: 12,
          titleColor: '#e8f0ff',
          bodyColor: '#8ba3c7',
          callbacks: { label: ctx => ctx.label + ': ' + fmtRp(ctx.raw) }
        }
      }
    }
  });

  // Bulanan
  const monthly = new Array(12).fill(0);
  yearRows().forEach(p => monthly[p.tanggal.getMonth()] += total(p));
  const cm = new Date().getMonth();

  if (state.charts.bulan) state.charts.bulan.destroy();
  state.charts.bulan = new Chart(document.getElementById('chartBulan'), {
    type: 'bar',
    data: {
      labels: MONTHS,
      datasets: [{
        label: 'Omset',
        data: monthly,
        backgroundColor: monthly.map((v, i) =>
          (i === cm && state.year === new Date().getFullYear())
            ? 'rgba(168,85,247,0.9)'
            : 'rgba(0,217,255,0.55)'
        ),
        borderRadius: 8,
        borderSkipped: false
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(10,14,31,0.95)',
          borderColor: 'rgba(255,255,255,0.1)',
          borderWidth: 1,
          padding: 12,
          titleColor: '#e8f0ff',
          bodyColor: '#8ba3c7',
          callbacks: { label: ctx => fmtRp(ctx.raw) }
        }
      },
      scales: {
        x: { ticks: { color: '#8ba3c7', font: { size: 11 } }, grid: { display: false } },
        y: {
          ticks: { color: '#8ba3c7', font: { size: 11 }, callback: v => fmtShort(v) },
          grid: { color: 'rgba(255,255,255,0.05)' }
        }
      }
    }
  });
}

function renderPusatMitra() {
  const yr = yearRows();
  const jenisMap = {};
  state.outlets.forEach(o => jenisMap[o.nama] = o.jenis);

  let pusat = 0, mitra = 0, other = 0;
  yr.forEach(p => {
    const j = jenisMap[p.outlet];
    const t = total(p);
    if (j === 'Pusat') pusat += t;
    else if (j === 'Mitra') mitra += t;
    else other += t;
  });
  const tot = pusat + mitra + other || 1;

  document.getElementById('pusatMitra').innerHTML = `
    <div class="pm-bar">
      <div class="pm-seg pusat" style="width:${(pusat / tot * 100).toFixed(1)}%"></div>
      <div class="pm-seg mitra" style="width:${(mitra / tot * 100).toFixed(1)}%"></div>
      <div class="pm-seg other" style="width:${(other / tot * 100).toFixed(1)}%"></div>
    </div>
    <div class="pm-legend">
      <div><span class="dot" style="background:#00d9ff"></span> Pusat <strong>${fmtRp(pusat)}</strong> <em>${(pusat / tot * 100).toFixed(1)}%</em></div>
      <div><span class="dot" style="background:#a855f7"></span> Mitra <strong>${fmtRp(mitra)}</strong> <em>${(mitra / tot * 100).toFixed(1)}%</em></div>
      <div><span class="dot" style="background:#4a5568"></span> Belum klasifikasi <strong>${fmtRp(other)}</strong> <em>${(other / tot * 100).toFixed(1)}%</em></div>
    </div>
  `;
}

function renderRanking() {
  const totals = sumByOutlet(yearRows());
  const list = Object.entries(totals).map(([nama, v]) => ({ nama, v })).sort((a, b) => b.v - a.v);
  const grandTotal = list.reduce((s, x) => s + x.v, 0) || 1;

  const q = (document.getElementById('searchOutlet')?.value || '').toLowerCase().trim();
  const filtered = q ? list.filter(x => x.nama.toLowerCase().includes(q)) : list;

  document.getElementById('rankingList').innerHTML = filtered.length ? filtered.map(x => {
    const realRank = list.indexOf(x) + 1;
    const pct = x.v / grandTotal * 100;
    const badge = realRank <= 5 ? 'top' : (realRank >= list.length - 4 ? 'bottom' : '');
    return `
      <div class="rank-row">
        <div class="rank-num ${badge}">${realRank}</div>
        <div class="rank-body">
          <div class="rank-head"><strong>${x.nama}</strong><span>${fmtRp(x.v)}</span></div>
          <div class="rank-bar"><div style="width:${pct.toFixed(1)}%"></div></div>
        </div>
        <div class="rank-pct">${pct.toFixed(1)}%</div>
      </div>`;
  }).join('') : '<div class="empty">Belum ada data</div>';
}

function renderGrowth() {
  const now = new Date();
  const cm = now.getMonth(), cy = now.getFullYear();
  const cur = {}, last = {};

  state.penjualan.forEach(p => {
    if (p.tanggal.getFullYear() !== cy) return;
    const m = p.tanggal.getMonth();
    if (m === cm) cur[p.outlet] = (cur[p.outlet] || 0) + total(p);
    else if (m === cm - 1) last[p.outlet] = (last[p.outlet] || 0) + total(p);
  });

  const names = [...new Set([...Object.keys(cur), ...Object.keys(last)])];
  const list = names.map(nama => {
    const c = cur[nama] || 0, l = last[nama] || 0;
    const g = l > 0 ? (c / l - 1) * 100 : (c > 0 ? 999 : 0);
    return { nama, c, l, g };
  }).sort((a, b) => b.g - a.g);

  document.getElementById('growthList').innerHTML = `
    <div class="section-title">${MONTHS[cm]} ${cy} vs ${MONTHS[(cm - 1 + 12) % 12]}</div>
    ${list.length ? list.slice(0, 15).map(x => `
      <div class="growth-row">
        <div class="g-name">${x.nama}</div>
        <div class="g-val">${fmtRp(x.c)}</div>
        <div class="g-pct ${x.g >= 0 ? 'up' : 'down'}">${x.g >= 999 ? 'NEW' : (x.g >= 0 ? '+' : '') + x.g.toFixed(1) + '%'}</div>
      </div>`).join('') : '<div class="empty">Belum ada data bulan ini</div>'}
  `;
}

function renderMonitoring() {
  const now = new Date();
  const cm = now.getMonth(), cy = now.getFullYear();
  const targetDays = now.getDate();

  const counts = {};
  state.penjualan.forEach(p => {
    if (p.tanggal.getFullYear() === cy && p.tanggal.getMonth() === cm) {
      counts[p.outlet] = (counts[p.outlet] || 0) + 1;
    }
  });

  const outlets = state.outlets.length
    ? state.outlets.map(o => o.nama)
    : [...new Set(state.penjualan.map(p => p.outlet))];

  const done = outlets.filter(n => (counts[n] || 0) >= targetDays).length;
  const partial = outlets.filter(n => (counts[n] || 0) > 0 && (counts[n] || 0) < targetDays).length;
  const none = outlets.length - done - partial;

  document.getElementById('monitoring').innerHTML = `
    <div class="mon-stats">
      <div><span>${done}</span><small>Lengkap</small></div>
      <div><span>${partial}</span><small>Sebagian</small></div>
      <div><span>${none}</span><small>Belum input</small></div>
      <div><span>${outlets.length}</span><small>Total outlet</small></div>
    </div>
    <div class="section-title">Detail — target ${targetDays} hari</div>
    <div class="mon-list">
      ${outlets.map(nama => {
        const c = counts[nama] || 0;
        const pct = Math.min(100, c / targetDays * 100);
        const cls = c >= targetDays ? 'ok' : (c > 0 ? 'partial' : 'none');
        return `
          <div class="mon-row">
            <div class="mon-name">${nama}</div>
            <div class="mon-bar"><div class="${cls}" style="width:${pct}%"></div></div>
            <div class="mon-count">${c}/${targetDays}</div>
          </div>`;
      }).join('')}
    </div>
  `;
}

function renderDataTable() {
  const rows = [...state.penjualan]
    .map((p, i) => ({ ...p, _i: i }))
    .sort((a, b) => b.tanggal - a.tanggal)
    .slice(0, 200);

  document.getElementById('dataTable').innerHTML = `
    <div class="data-meta">Menampilkan ${rows.length} baris terbaru dari total ${state.penjualan.length}</div>
    <div class="data-scroll">
      <table class="data-table">
        <thead><tr>
          <th>Tanggal</th><th>Outlet</th><th>Dine in</th><th>Take away</th>
          <th>GoFood</th><th>GrabFood</th><th>ShopeeFood</th><th>ESB</th><th>Total</th><th>×</th>
        </tr></thead>
        <tbody>
          ${rows.length ? rows.map(p => `
            <tr>
              <td>${p.tanggal.toLocaleDateString('id-ID')}</td>
              <td>${p.outlet}</td>
              <td>${N(p.dineIn).toLocaleString('id-ID')}</td>
              <td>${N(p.takeAway).toLocaleString('id-ID')}</td>
              <td>${N(p.goFood).toLocaleString('id-ID')}</td>
              <td>${N(p.grabFood).toLocaleString('id-ID')}</td>
              <td>${N(p.shopeeFood).toLocaleString('id-ID')}</td>
              <td>${N(p.esbOrder).toLocaleString('id-ID')}</td>
              <td><strong>${fmtRp(total(p))}</strong></td>
              <td><button class="del-btn" data-idx="${p._i}">✕</button></td>
            </tr>`).join('') : '<tr><td colspan="10" class="empty">Belum ada data</td></tr>'}
        </tbody>
      </table>
    </div>`;

  document.querySelectorAll('.del-btn').forEach(b => {
    b.onclick = () => {
      const idx = +b.dataset.idx;
      if (confirm('Hapus baris ini?')) {
        state.penjualan.splice(idx, 1);
        saveToStorage();
        renderAll();
        toast('Baris dihapus', 'info');
      }
    };
  });
}

/* ============ STORAGE ============ */
function saveToStorage() {
  try {
    localStorage.setItem('dash_penjualan', JSON.stringify(
      state.penjualan.map(p => ({ ...p, tanggal: p.tanggal.toISOString() }))
    ));
    localStorage.setItem('dash_year', String(state.year));
    localStorage.setItem('dash_outlets', JSON.stringify(state.outlets));
  } catch (e) { console.warn(e); }
}

function loadFromStorage() {
  try {
    const rawOutlets = localStorage.getItem('dash_outlets');
    if (rawOutlets) state.outlets = JSON.parse(rawOutlets);
    const raw = localStorage.getItem('dash_penjualan');
    if (raw) state.penjualan = JSON.parse(raw).map(p => ({ ...p, tanggal: new Date(p.tanggal) }));
    const y = localStorage.getItem('dash_year');
    if (y) state.year = +y;
  } catch (e) { console.warn(e); }
}

/* ============ BINDERS ============ */
function bindTabs() {
  document.querySelectorAll('.tab').forEach(t => {
    t.onclick = () => {
      document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('active'));
      t.classList.add('active');
      document.getElementById('tab-' + t.dataset.tab).classList.add('active');
    };
  });
}

function bindFileUpload() {
  const input = document.getElementById('fileInput');
  input.onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: true });
        const parsed = parseWorkbook(wb);
        state.workbook = wb;
        state.outlets = parsed.outlets;
        state.penjualan = parsed.penjualan;

        const dash = wb.Sheets['DASHBOARD'];
        if (dash && dash['B3'] && dash['B3'].v) {
          const y = N(dash['B3'].v);
          if (y > 2000) state.year = y;
        }

        saveToStorage();
        renderAll();
        toast(`✓ ${parsed.penjualan.length} baris · ${parsed.outlets.length} outlet berhasil dimuat`, 'success');
      } catch (err) {
        console.error(err);
        toast('Gagal baca file: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    input.value = '';
  };
}

function bindAddForm() {
  const modal = document.getElementById('modal');
  const form = document.getElementById('formInput');
  const outletSel = document.getElementById('fOutlet');

  document.getElementById('btnAdd').onclick = () => {
    const list = state.outlets;
    outletSel.innerHTML = list.length
      ? list.map(o => `<option value="${o.nama}">${o.nama}</option>`).join('')
      : '<option value="">(Upload Excel dulu)</option>';
    form.reset();
    document.getElementById('fTanggal').value = todayStr();
    if (list.length) outletSel.selectedIndex = 0;
    modal.classList.add('show');
  };

  const close = () => modal.classList.remove('show');
  document.getElementById('fCancel').onclick = close;
  document.getElementById('fCancel2').onclick = close;
  modal.onclick = e => { if (e.target === modal) close(); };

  form.onsubmit = e => {
    e.preventDefault();
    const tglStr = document.getElementById('fTanggal').value;
    const outlet = outletSel.value;
    if (!tglStr || !outlet) { toast('Isi tanggal dan outlet', 'error'); return; }
    const [y, m, d] = tglStr.split('-').map(Number);

    state.penjualan.push({
      tanggal: new Date(y, m - 1, d),
      outlet,
      dineIn: N(document.getElementById('fDineIn').value),
      takeAway: N(document.getElementById('fTakeAway').value),
      goFood: N(document.getElementById('fGoFood').value),
      grabFood: N(document.getElementById('fGrabFood').value),
      shopeeFood: N(document.getElementById('fShopeeFood').value),
      esbOrder: N(document.getElementById('fEsbOrder').value),
      catatan: document.getElementById('fCatatan').value
    });

    saveToStorage();
    renderAll();
    close();
    toast('✓ Data berhasil ditambahkan', 'success');
  };
}

function bindExport() {
  document.getElementById('btnExport').onclick = () => {
    try {
      const wb = XLSX.utils.book_new();

      // MASTER OUTLET
      const masterAoa = [
        ['MASTER DATA OUTLET'],
        [],
        ['No', 'Nama Outlet', 'Wilayah', 'Jenis Outlet'],
        ...state.outlets.map((o, i) => [i + 1, o.nama, o.wilayah, o.jenis])
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(masterAoa), 'MASTER OUTLET');

      // INPUT PENJUALAN
      const penjAoa = [
        ['PANEL INPUT PENJUALAN HARIAN — INPUT HANYA DI SHEET INI'],
        ['Satu baris = satu outlet per tanggal'],
        [],
        ['Tanggal', 'Nama Outlet', 'Dine in', 'Take away', 'GoFood', 'GrabFood', 'ShopeeFood', 'ESB Order', 'TOTAL OMSET', 'Catatan'],
        ...state.penjualan.map(p => [
          p.tanggal, p.outlet, p.dineIn, p.takeAway, p.goFood, p.grabFood,
          p.shopeeFood, p.esbOrder, null, p.catatan || ''
        ])
      ];
      const wsP = XLSX.utils.aoa_to_sheet(penjAoa);

      // Total formula
      state.penjualan.forEach((p, i) => {
        const r = 5 + i;
        wsP['I' + r] = { t: 'n', v: total(p), f: `IF(OR(A${r}="",B${r}=""),"",SUM(C${r}:H${r}))` };
      });

      // Date format
      for (let i = 0; i < state.penjualan.length; i++) {
        const cell = wsP['A' + (5 + i)];
        if (cell && cell.t === 'd') cell.z = 'yyyy-mm-dd';
      }

      XLSX.utils.book_append_sheet(wb, wsP, 'INPUT PENJUALAN');

      // DASHBOARD (ringkas)
      const yr = yearRows();
      const cat = sumCat(yr);
      const dashAoa = [
        ['DASHBOARD PENJUALAN'], [],
        ['Tahun Analisis', state.year], [],
        ['Total Omset Tahun', sumRows(yr)],
        ['Total Transaksi', yr.length], [],
        ['Kategori', 'Total Omset', 'Kontribusi'],
        ...CATS.map(c => [c.label, cat[c.key], sumRows(yr) ? cat[c.key] / sumRows(yr) : 0])
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dashAoa), 'DASHBOARD');

      // Preserve sheet lain dari file asli
      if (state.workbook) {
        for (const name of state.workbook.SheetNames) {
          if (['MASTER OUTLET', 'INPUT PENJUALAN', 'DASHBOARD'].includes(name)) continue;
          if (!wb.Sheets[name]) {
            XLSX.utils.book_append_sheet(wb, state.workbook.Sheets[name], name);
          }
        }
      }

      const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
      const blob = new Blob([buf], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Data_Penjualan_${state.year}_${todayStr()}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      toast('✓ Excel berhasil di-export', 'success');
    } catch (err) {
      console.error(err);
      toast('Gagal export: ' + err.message, 'error');
    }
  };
}

function bindSearch() {
  const el = document.getElementById('searchOutlet');
  if (el) el.oninput = renderRanking;
}

function bindYearSelect() {
  document.getElementById('yearSelect').onchange = e => {
    state.year = +e.target.value;
    saveToStorage();
    renderAll();
  };
}

/* ============ INIT ============ */
document.addEventListener('DOMContentLoaded', () => {
  bindTabs();
  bindFileUpload();
  bindAddForm();
  bindExport();
  bindSearch();
  bindYearSelect();
  loadFromStorage();
  renderAll();
});