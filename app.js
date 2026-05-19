/* ═══════════════════════════════════════════════════
   DATA STORE
═══════════════════════════════════════════════════ */
let barang = [
  {id:1,kode:'ATK-001',nama:'Kertas A4 80gsm',kategori:'Kertas',stok:45,min:20,satuan:'Rim',harga:45000},
  {id:2,kode:'ATK-002',nama:'Pulpen Ballpoint Biru',kategori:'Alat Tulis',stok:8,min:20,satuan:'Lusin',harga:18000},
  {id:3,kode:'ATK-003',nama:'Pensil 2B',kategori:'Alat Tulis',stok:0,min:5,satuan:'Lusin',harga:15000},
  {id:4,kode:'ATK-004',nama:'Stapler Besar',kategori:'Lainnya',stok:3,min:2,satuan:'Pcs',harga:35000},
  {id:5,kode:'ATK-005',nama:'Map Plastik',kategori:'File & Arsip',stok:30,min:15,satuan:'Pcs',harga:5000},
  {id:6,kode:'ATK-006',nama:'Tinta Printer Hitam',kategori:'Elektronik',stok:4,min:5,satuan:'Botol',harga:85000},
  {id:7,kode:'ATK-007',nama:'Penggaris 30cm',kategori:'Alat Tulis',stok:12,min:5,satuan:'Pcs',harga:8000},
  {id:8,kode:'ATK-008',nama:'Lem Stick',kategori:'Lainnya',stok:6,min:10,satuan:'Pcs',harga:12000},
  {id:9,kode:'ATK-009',nama:'Amplop Coklat A4',kategori:'File & Arsip',stok:100,min:50,satuan:'Pcs',harga:1500},
  {id:10,kode:'ATK-010',nama:'Spidol Whiteboard',kategori:'Alat Tulis',stok:2,min:8,satuan:'Pcs',harga:22000},
];

let permintaan = [
  {id:1,noReq:'REQ-001',tgl:'2025-05-10',pemohon:'Budi Santoso',divisi:'Marketing',barangId:2,qty:2,keperluan:'Kebutuhan tim marketing',status:'Approved'},
  {id:2,noReq:'REQ-002',tgl:'2025-05-12',pemohon:'Siti Rahayu',divisi:'HRD',barangId:6,qty:1,keperluan:'Printer HRD rusak',status:'Pending'},
  {id:3,noReq:'REQ-003',tgl:'2025-05-14',pemohon:'Ahmad Fauzi',divisi:'Logistik',barangId:8,qty:3,keperluan:'Packing dokumen pengiriman',status:'Rejected'},
  {id:4,noReq:'REQ-004',tgl:'2025-05-15',pemohon:'Dewi Lestari',divisi:'Keuangan',barangId:1,qty:5,keperluan:'Bulanan divisi keuangan',status:'Approved'},
  {id:5,noReq:'REQ-005',tgl:'2025-05-16',pemohon:'Rudi Hartono',divisi:'IT',barangId:10,qty:4,keperluan:'Rapat presentasi mingguan',status:'Pending'},
];

let riwayat = [
  {id:1,noTrx:'TRX-001',tgl:'2025-05-08',jenis:'Masuk',barangId:1,qty:20,pelaku:'Admin',ket:'Pembelian rutin bulanan'},
  {id:2,noTrx:'TRX-002',tgl:'2025-05-10',jenis:'Keluar',barangId:2,qty:2,pelaku:'Budi Santoso',ket:'Distribusi Marketing'},
  {id:3,noTrx:'TRX-003',tgl:'2025-05-12',jenis:'Masuk',barangId:6,qty:3,pelaku:'Admin',ket:'Restock tinta printer'},
  {id:4,noTrx:'TRX-004',tgl:'2025-05-14',jenis:'Keluar',barangId:5,qty:10,pelaku:'Ahmad Fauzi',ket:'Arsip dokumen Q1 2025'},
  {id:5,noTrx:'TRX-005',tgl:'2025-05-15',jenis:'Keluar',barangId:1,qty:5,pelaku:'Dewi Lestari',ket:'Keuangan bulanan'},
  {id:6,noTrx:'TRX-006',tgl:'2025-05-16',jenis:'Masuk',barangId:3,qty:10,pelaku:'Admin',ket:'Restock pensil'},
  {id:7,noTrx:'TRX-007',tgl:'2025-05-16',jenis:'Keluar',barangId:10,qty:3,pelaku:'Rudi Hartono',ket:'Rapat mingguan IT'},
];

let nextBarangId = 11, nextReqId = 6, nextTrxId = 8;
let nextReqNum = 6, nextTrxNum = 8;

/* ═══════════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════════ */
const getBarang = id => barang.find(b => b.id === id);
const fmt = n => 'Rp ' + n.toLocaleString('id-ID');
const today = () => new Date().toISOString().slice(0, 10);

function statusStok(b) {
  if (b.stok === 0) return { label:'Habis', cls:'badge-out' };
  if (b.stok < b.min) return { label:'Rendah', cls:'badge-low' };
  return { label:'Aman', cls:'badge-ok' };
}

function badge(label, cls) {
  return `<span class="badge ${cls}"><span class="badge-dot"></span>${label}</span>`;
}

function stockBar(stok, min) {
  const pct = min === 0 ? 100 : Math.min(100, Math.round((stok / (min * 2)) * 100));
  const color = stok === 0 ? '#EF4444' : stok < min ? '#F59E0B' : '#10B981';
  return `<div class="stock-wrap">
    <span style="font-family:var(--ff-display);font-weight:600;font-size:15px;color:var(--text-1);min-width:28px">${stok}</span>
    <div class="stock-bar-bg"><div class="stock-bar-fill" style="width:${pct}%;background:${color}"></div></div>
  </div>`;
}

/* ═══════════════════════════════════════════════════
   EXCEL DOWNLOAD
═══════════════════════════════════════════════════ */
function downloadExcel(type) {
  const wb = XLSX.utils.book_new();
  const tgl = new Date().toLocaleDateString('id-ID', {day:'2-digit', month:'long', year:'numeric'});

  if (type === 'stok' || type === 'ringkasan') {
    const stokData = barang.map(b => {
      const masukList  = riwayat.filter(r => r.barangId === b.id && r.jenis === 'Masuk');
      const keluarList = riwayat.filter(r => r.barangId === b.id && r.jenis === 'Keluar');
      const totalMasuk  = masukList.reduce((a, r) => a + r.qty, 0);
      const totalKeluar = keluarList.reduce((a, r) => a + r.qty, 0);
      const status = statusStok(b).label;
      return {
        'Kode Barang':          b.kode,
        'Nama Barang':          b.nama,
        'Kategori':             b.kategori,
        'Satuan':               b.satuan,
        'Stok Tersisa':         b.stok,
        'Min. Stok':            b.min,
        'Total Barang Masuk':   totalMasuk,
        'Total Barang Keluar':  totalKeluar,
        'Harga Satuan (Rp)':    b.harga,
        'Nilai Stok (Rp)':      b.stok * b.harga,
        'Status':               status,
      };
    });

    const ws1 = XLSX.utils.json_to_sheet(stokData);
    ws1['!cols'] = [
      {wch:12},{wch:28},{wch:14},{wch:10},{wch:14},{wch:12},
      {wch:20},{wch:22},{wch:18},{wch:18},{wch:10}
    ];
    XLSX.utils.book_append_sheet(wb, ws1, 'Data Stok Barang');
  }

  if (type === 'transaksi' || type === 'ringkasan') {
    const trxData = [...riwayat].sort((a,b) => b.id - a.id).map(r => {
      const b = getBarang(r.barangId);
      return {
        'No. Transaksi': r.noTrx,
        'Tanggal':       r.tgl,
        'Jenis':         r.jenis,
        'Nama Barang':   b ? b.nama : '-',
        'Kode Barang':   b ? b.kode : '-',
        'Kategori':      b ? b.kategori : '-',
        'Jumlah':        r.qty,
        'Satuan':        b ? b.satuan : '-',
        'Pelaku':        r.pelaku,
        'Keterangan':    r.ket,
      };
    });

    const ws2 = XLSX.utils.json_to_sheet(trxData);
    ws2['!cols'] = [
      {wch:14},{wch:12},{wch:10},{wch:28},{wch:12},
      {wch:14},{wch:10},{wch:10},{wch:18},{wch:30}
    ];
    XLSX.utils.book_append_sheet(wb, ws2, 'Riwayat Transaksi');
  }

  if (type === 'ringkasan') {
    const totalNilai  = barang.reduce((a, b) => a + b.stok * b.harga, 0);
    const totalMasuk  = riwayat.filter(r => r.jenis === 'Masuk').reduce((a, r) => a + r.qty, 0);
    const totalKeluar = riwayat.filter(r => r.jenis === 'Keluar').reduce((a, r) => a + r.qty, 0);

    const summaryRows = [
      ['RINGKASAN LAPORAN ATK', ''],
      ['Tanggal Cetak:', tgl],
      ['', ''],
      ['INFORMASI STOK', ''],
      ['Total Jenis Barang',   barang.length],
      ['Stok Status Aman',     barang.filter(b => b.stok >= b.min).length],
      ['Stok Status Rendah',   barang.filter(b => b.stok > 0 && b.stok < b.min).length],
      ['Stok Status Habis',    barang.filter(b => b.stok === 0).length],
      ['Estimasi Nilai Stok (Rp)', totalNilai],
      ['', ''],
      ['INFORMASI TRANSAKSI', ''],
      ['Total Unit Barang Masuk',  totalMasuk],
      ['Total Unit Barang Keluar', totalKeluar],
      ['Total Transaksi',          riwayat.length],
      ['', ''],
      ['INFORMASI PERMINTAAN', ''],
      ['Total Permintaan',         permintaan.length],
      ['Permintaan Disetujui',     permintaan.filter(p => p.status === 'Approved').length],
      ['Permintaan Ditolak',       permintaan.filter(p => p.status === 'Rejected').length],
      ['Permintaan Pending',       permintaan.filter(p => p.status === 'Pending').length],
      ['', ''],
      ['DISTRIBUSI STOK PER KATEGORI', ''],
      ...['Kertas','Alat Tulis','File & Arsip','Elektronik','Lainnya'].map(cat => {
        const total = barang.filter(b => b.kategori === cat).reduce((a, b) => a + b.stok, 0);
        return [cat, total];
      }),
    ];

    const ws3 = XLSX.utils.aoa_to_sheet(summaryRows);
    ws3['!cols'] = [{wch:30}, {wch:20}];
    XLSX.utils.book_append_sheet(wb, ws3, 'Ringkasan Eksekutif');
  }

  const fileNames = {
    stok:      `ATK_Data_Stok_${today()}.xlsx`,
    transaksi: `ATK_Riwayat_Transaksi_${today()}.xlsx`,
    ringkasan: `ATK_Laporan_Lengkap_${today()}.xlsx`,
  };

  XLSX.writeFile(wb, fileNames[type]);

  const labels = {
    stok:      'Data Stok Barang',
    transaksi: 'Riwayat Transaksi',
    ringkasan: 'Laporan Lengkap'
  };
  toast(`📊 ${labels[type]} berhasil diunduh!`, 'success');
}

/* ═══════════════════════════════════════════════════
   NAVIGATION
═══════════════════════════════════════════════════ */
const pageTitles = {
  dashboard:'Dashboard', stok:'Data Barang',
  permintaan:'Permintaan', riwayat:'Riwayat Transaksi', laporan:'Laporan'
};

function navigate(page, el) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + page).classList.add('active');
  if (el) el.classList.add('active');
  document.getElementById('topbar-title').textContent = pageTitles[page];
  if (page === 'dashboard')  renderDashboard();
  if (page === 'stok')       renderStok();
  if (page === 'permintaan') renderPermintaan();
  if (page === 'riwayat')    renderRiwayat();
  if (page === 'laporan')    setTimeout(renderLaporan, 50);
}

/* ═══════════════════════════════════════════════════
   RENDER: DASHBOARD
═══════════════════════════════════════════════════ */
function renderDashboard() {
  const aman    = barang.filter(b => b.stok >= b.min).length;
  const rendah  = barang.filter(b => b.stok > 0 && b.stok < b.min).length;
  const habis   = barang.filter(b => b.stok === 0).length;
  const pending = permintaan.filter(p => p.status === 'Pending').length;

  document.getElementById('kpi-grid').innerHTML = `
    <div class="kpi-card kpi-blue">
      <div class="kpi-icon">📦</div>
      <div class="kpi-label">Total Jenis Barang</div>
      <div class="kpi-value">${barang.length}</div>
      <div class="kpi-sub">jenis ATK terdaftar</div>
    </div>
    <div class="kpi-card kpi-green">
      <div class="kpi-icon">✅</div>
      <div class="kpi-label">Stok Aman</div>
      <div class="kpi-value">${aman}</div>
      <div class="kpi-sub">barang dalam kondisi baik</div>
    </div>
    <div class="kpi-card kpi-gold">
      <div class="kpi-icon">⚠️</div>
      <div class="kpi-label">Stok Rendah</div>
      <div class="kpi-value">${rendah}</div>
      <div class="kpi-sub">perlu segera restock</div>
    </div>
    <div class="kpi-card kpi-red">
      <div class="kpi-icon">🚨</div>
      <div class="kpi-label">Stok Habis</div>
      <div class="kpi-value">${habis}</div>
      <div class="kpi-sub">tidak tersedia</div>
    </div>
  `;

  document.getElementById('badge-pending').textContent = pending;

  const kritis = barang.filter(b => statusStok(b).label !== 'Aman');
  document.getElementById('alert-dashboard').innerHTML = kritis.length
    ? `<div class="alert alert-warning">
        <span class="alert-icon">⚠️</span>
        <div class="alert-text"><strong>${kritis.length} barang</strong> memerlukan perhatian: ${kritis.map(b=>b.nama).join(', ')}.</div>
      </div>`
    : '';

  document.getElementById('tbl-alert').innerHTML = kritis.length
    ? kritis.map(b => {
        const s = statusStok(b);
        return `<tr class="${s.label==='Habis'?'row-danger':'row-warn'}">
          <td class="td-name">${b.nama}</td>
          <td class="td-num">${b.stok}</td>
          <td style="color:var(--text-3)">${b.min}</td>
          <td>${badge(s.label, s.cls)}</td>
        </tr>`;
      }).join('')
    : `<tr><td colspan="4"><div class="empty-state"><div class="empty-icon">✅</div><p>Semua stok aman</p></div></td></tr>`;

  const sorted = [...riwayat].sort((a,b)=>b.id-a.id).slice(0,5);
  document.getElementById('tbl-recent').innerHTML = sorted.map(r => {
    const b = getBarang(r.barangId);
    return `<tr>
      <td class="td-name" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${b ? b.nama : '—'}</td>
      <td>${badge(r.jenis, r.jenis==='Masuk'?'badge-in':'badge-out2')}</td>
      <td class="td-num">${r.qty}</td>
      <td style="color:var(--text-3);font-size:12px">${r.tgl}</td>
    </tr>`;
  }).join('');
}

/* ═══════════════════════════════════════════════════
   RENDER: STOK
═══════════════════════════════════════════════════ */
function renderStok() {
  const q   = (document.getElementById('search-stok')?.value || '').toLowerCase();
  const kat = document.getElementById('filter-kategori')?.value || '';
  const sts = document.getElementById('filter-status')?.value || '';

  const list = barang.filter(b => {
    const matchQ   = !q || b.nama.toLowerCase().includes(q) || b.kode.toLowerCase().includes(q);
    const matchKat = !kat || b.kategori === kat;
    const matchSts = !sts || statusStok(b).label === sts;
    return matchQ && matchKat && matchSts;
  });

  document.getElementById('tbl-stok').innerHTML = list.length === 0
    ? `<tr><td colspan="9"><div class="empty-state"><div class="empty-icon">📦</div><p>Tidak ada barang ditemukan</p></div></td></tr>`
    : list.map(b => {
        const s = statusStok(b);
        const pct = b.min === 0 ? 100 : Math.min(100, Math.round((b.stok / (b.min * 2)) * 100));
        const barColor = b.stok === 0 ? '#EF4444' : b.stok < b.min ? '#F59E0B' : '#10B981';
        return `<tr id="row-${b.id}">
          <td class="td-code">${b.kode}</td>
          <td class="td-name">${b.nama}</td>
          <td style="color:var(--text-3)">${b.kategori}</td>
          <td class="stok-cell" id="stok-cell-${b.id}" onclick="openInlineEditor(${b.id}, event)">
            <div class="stok-display">
              <span class="stok-num" id="stok-num-${b.id}">${b.stok}</span>
              <div class="stok-bar-bg" style="min-width:50px">
                <div class="stok-bar-fill" id="stok-bar-${b.id}" style="width:${pct}%;background:${barColor}"></div>
              </div>
              <span class="stok-edit-hint">✏️ klik edit</span>
            </div>
          </td>
          <td style="color:var(--text-3)">${b.min}</td>
          <td style="color:var(--text-3)">${b.satuan}</td>
          <td style="color:var(--cyan);font-size:13px">${fmt(b.harga)}</td>
          <td id="stok-badge-${b.id}">${badge(s.label, s.cls)}</td>
          <td>
            <div style="display:flex;gap:6px">
              <button class="btn btn-sm" onclick="openEditBarang(${b.id})" title="Edit Data">✏️</button>
              <button class="btn btn-sm btn-danger" onclick="hapusBarang(${b.id})" title="Hapus">🗑️</button>
            </div>
          </td>
        </tr>`;
      }).join('');
}

/* ─── INLINE EDITOR ─── */
let activeEditorId = null;
let currentEditorType = 'Masuk';

function openInlineEditor(id, event) {
  event.stopPropagation();
  closeInlineEditor();

  const b = getBarang(id);
  activeEditorId = id;

  const cell = document.getElementById(`stok-cell-${id}`);
  const editor = document.createElement('div');
  editor.className = 'inline-editor';
  editor.id = 'inline-editor';
  editor.onclick = e => e.stopPropagation();
  editor.innerHTML = `
    <div class="inline-editor-title">✏️ Edit Stok — ${b.nama}</div>
    <div class="inline-editor-row">
      <div class="inline-type-toggle">
        <button class="inline-type-btn active-in" id="btn-type-masuk" onclick="setEditorType('Masuk')">📥 Masuk</button>
        <button class="inline-type-btn" id="btn-type-keluar" onclick="setEditorType('Keluar')">📤 Keluar</button>
      </div>
      <input class="inline-qty-input" id="inline-qty" type="number" min="1" value="1"
        onkeydown="handleInlineKey(event, ${id})"
        onfocus="this.select()">
      <button class="inline-confirm" onclick="confirmInlineEdit(${id})">Simpan</button>
      <button class="inline-cancel" onclick="closeInlineEditor()">✕</button>
    </div>
    <div class="inline-editor-note">
      <span>Stok saat ini:</span>
      <strong style="color:var(--text-1);font-family:var(--ff-display)">${b.stok}</strong>
      <span>${b.satuan} &nbsp;·&nbsp; Enter untuk simpan, Esc untuk batal</span>
    </div>
  `;

  cell.style.position = 'relative';
  cell.appendChild(editor);
  setTimeout(() => document.getElementById('inline-qty')?.focus(), 50);
  setTimeout(() => document.addEventListener('click', outsideClickHandler), 10);
}

function outsideClickHandler(e) {
  if (!document.getElementById('inline-editor')?.contains(e.target)) {
    closeInlineEditor();
  }
}

function closeInlineEditor() {
  const editor = document.getElementById('inline-editor');
  if (editor) editor.remove();
  document.removeEventListener('click', outsideClickHandler);
  activeEditorId = null;
}

function setEditorType(type) {
  currentEditorType = type;
  const btnMasuk  = document.getElementById('btn-type-masuk');
  const btnKeluar = document.getElementById('btn-type-keluar');
  if (!btnMasuk || !btnKeluar) return;
  btnMasuk.className  = 'inline-type-btn' + (type === 'Masuk'  ? ' active-in'  : '');
  btnKeluar.className = 'inline-type-btn' + (type === 'Keluar' ? ' active-out' : '');
}

function handleInlineKey(event, id) {
  if (event.key === 'Enter')  confirmInlineEdit(id);
  if (event.key === 'Escape') closeInlineEditor();
}

function confirmInlineEdit(id) {
  const b   = getBarang(id);
  const qty = parseInt(document.getElementById('inline-qty')?.value) || 0;
  const type = currentEditorType;

  if (qty <= 0) { toast('Jumlah harus lebih dari 0', 'error'); return; }
  if (type === 'Keluar' && qty > b.stok) {
    toast(`Stok tidak cukup! Tersisa ${b.stok} ${b.satuan}`, 'error'); return;
  }

  b.stok += type === 'Masuk' ? qty : -qty;

  riwayat.push({
    id: nextTrxId++,
    noTrx: `TRX-${String(nextTrxNum++).padStart(3,'0')}`,
    tgl: today(), jenis: type, barangId: id, qty, pelaku: 'Admin', ket: 'Update stok langsung'
  });

  closeInlineEditor();

  const pct = b.min === 0 ? 100 : Math.min(100, Math.round((b.stok / (b.min * 2)) * 100));
  const barColor = b.stok === 0 ? '#EF4444' : b.stok < b.min ? '#F59E0B' : '#10B981';
  const numEl   = document.getElementById(`stok-num-${id}`);
  const barEl   = document.getElementById(`stok-bar-${id}`);
  const badgeEl = document.getElementById(`stok-badge-${id}`);
  const row     = document.getElementById(`row-${id}`);

  if (numEl)   numEl.textContent = b.stok;
  if (barEl)   { barEl.style.width = pct + '%'; barEl.style.background = barColor; }
  if (badgeEl) { const s = statusStok(b); badgeEl.innerHTML = badge(s.label, s.cls); }

  if (row) {
    row.classList.remove('flash-in', 'flash-out');
    void row.offsetWidth;
    row.classList.add(type === 'Masuk' ? 'flash-in' : 'flash-out');
    setTimeout(() => row.classList.remove('flash-in', 'flash-out'), 1000);
  }

  renderDashboard();
  toast(`${type === 'Masuk' ? '📥' : '📤'} Stok ${b.nama} ${type === 'Masuk' ? 'ditambah' : 'dikurangi'} ${qty} ${b.satuan} → sisa ${b.stok}`, type === 'Masuk' ? 'success' : 'warning');
}

/* ═══════════════════════════════════════════════════
   RENDER: PERMINTAAN
═══════════════════════════════════════════════════ */
function renderPermintaan() {
  const q   = (document.getElementById('search-req')?.value || '').toLowerCase();
  const sts = document.getElementById('filter-status-req')?.value || '';

  const list = permintaan.filter(p => {
    const b = getBarang(p.barangId);
    const matchQ   = !q || p.pemohon.toLowerCase().includes(q) || (b && b.nama.toLowerCase().includes(q));
    const matchSts = !sts || p.status === sts;
    return matchQ && matchSts;
  }).sort((a,b) => b.id - a.id);

  const badgeCls = { Pending:'badge-pending', Approved:'badge-approved', Rejected:'badge-rejected' };

  document.getElementById('tbl-permintaan').innerHTML = list.length === 0
    ? `<tr><td colspan="9"><div class="empty-state"><div class="empty-icon">📋</div><p>Tidak ada permintaan</p></div></td></tr>`
    : list.map(p => {
        const b = getBarang(p.barangId);
        const actions = p.status === 'Pending'
          ? `<button class="btn btn-sm btn-success" onclick="setujuPermintaan(${p.id})">✅ Setuju</button>
             <button class="btn btn-sm btn-danger" onclick="tolakPermintaan(${p.id})">❌ Tolak</button>`
          : `<span style="color:var(--text-3);font-size:12px">—</span>`;
        return `<tr>
          <td class="td-code">${p.noReq}</td>
          <td style="color:var(--text-3);font-size:12px;white-space:nowrap">${p.tgl}</td>
          <td class="td-name">${p.pemohon}</td>
          <td style="color:var(--text-3)">${p.divisi}</td>
          <td style="color:var(--text-1)">${b ? b.nama : '—'}</td>
          <td class="td-num">${p.qty} <span style="font-size:11px;color:var(--text-3);font-family:var(--ff-body)">${b ? b.satuan : ''}</span></td>
          <td style="color:var(--text-3);font-size:12px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${p.keperluan}">${p.keperluan}</td>
          <td>${badge(p.status, badgeCls[p.status])}</td>
          <td><div style="display:flex;gap:6px;flex-wrap:wrap">${actions}</div></td>
        </tr>`;
      }).join('');

  document.getElementById('badge-pending').textContent = permintaan.filter(p => p.status === 'Pending').length;
}

/* ═══════════════════════════════════════════════════
   RENDER: RIWAYAT
═══════════════════════════════════════════════════ */
function renderRiwayat() {
  const q     = (document.getElementById('search-trx')?.value || '').toLowerCase();
  const jenis = document.getElementById('filter-jenis-trx')?.value || '';

  const list = riwayat.filter(r => {
    const b = getBarang(r.barangId);
    const matchQ = !q || r.pelaku.toLowerCase().includes(q) || (b && b.nama.toLowerCase().includes(q));
    const matchJ = !jenis || r.jenis === jenis;
    return matchQ && matchJ;
  }).sort((a,b) => b.id - a.id);

  document.getElementById('tbl-riwayat').innerHTML = list.length === 0
    ? `<tr><td colspan="8"><div class="empty-state"><div class="empty-icon">🔄</div><p>Belum ada transaksi</p></div></td></tr>`
    : list.map(r => {
        const b = getBarang(r.barangId);
        return `<tr>
          <td class="td-code">${r.noTrx}</td>
          <td style="color:var(--text-3);font-size:12px;white-space:nowrap">${r.tgl}</td>
          <td>${badge(r.jenis, r.jenis==='Masuk'?'badge-in':'badge-out2')}</td>
          <td class="td-name">${b ? b.nama : '—'}</td>
          <td class="td-num">${r.qty}</td>
          <td style="color:var(--text-3)">${b ? b.satuan : ''}</td>
          <td style="color:var(--text-2)">${r.pelaku}</td>
          <td style="color:var(--text-3);font-size:12px">${r.ket}</td>
        </tr>`;
      }).join('');
}

/* ═══════════════════════════════════════════════════
   RENDER: LAPORAN
═══════════════════════════════════════════════════ */
function renderLaporan() {
  const totalKeluar = riwayat.filter(r=>r.jenis==='Keluar').reduce((a,r)=>a+r.qty,0);
  const totalMasuk  = riwayat.filter(r=>r.jenis==='Masuk').reduce((a,r)=>a+r.qty,0);
  const nilaiStok   = barang.reduce((a,b)=>a+b.stok*b.harga,0);
  const approved    = permintaan.filter(p=>p.status==='Approved').length;

  document.getElementById('kpi-laporan').innerHTML = `
    <div class="kpi-card kpi-blue">
      <div class="kpi-icon">📤</div>
      <div class="kpi-label">Total Unit Keluar</div>
      <div class="kpi-value">${totalKeluar}</div>
      <div class="kpi-sub">unit terdistribusi</div>
    </div>
    <div class="kpi-card kpi-green">
      <div class="kpi-icon">📥</div>
      <div class="kpi-label">Total Unit Masuk</div>
      <div class="kpi-value">${totalMasuk}</div>
      <div class="kpi-sub">unit diterima</div>
    </div>
    <div class="kpi-card kpi-gold">
      <div class="kpi-icon">✅</div>
      <div class="kpi-label">Req. Disetujui</div>
      <div class="kpi-value">${approved}</div>
      <div class="kpi-sub">dari ${permintaan.length} permintaan</div>
    </div>
    <div class="kpi-card kpi-blue">
      <div class="kpi-icon">💰</div>
      <div class="kpi-label">Nilai Stok</div>
      <div class="kpi-value" style="font-size:20px">${fmt(nilaiStok)}</div>
      <div class="kpi-sub">estimasi nilai persediaan</div>
    </div>
  `;

  const cats      = ['Kertas','Alat Tulis','File & Arsip','Elektronik','Lainnya'];
  const catColors = ['#3B82F6','#10B981','#F59E0B','#8B5CF6','#EF4444'];
  const catStok   = cats.map(c => barang.filter(b=>b.kategori===c).reduce((a,b)=>a+b.stok,0));
  const maxStok   = Math.max(...catStok, 1);

  document.getElementById('chart-kategori').innerHTML = cats.map((c,i) => `
    <div class="bar-row">
      <div class="bar-label">${c}</div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${Math.round(catStok[i]/maxStok*100)}%;background:${catColors[i]}">
          ${catStok[i]}
        </div>
      </div>
    </div>
  `).join('');

  const aman   = barang.filter(b=>b.stok>=b.min).length;
  const rendah = barang.filter(b=>b.stok>0&&b.stok<b.min).length;
  const habis  = barang.filter(b=>b.stok===0).length;
  drawDonut([aman,rendah,habis],['#10B981','#F59E0B','#EF4444'],['Aman','Rendah','Habis']);

  const reqCount = {};
  permintaan.forEach(p => { reqCount[p.barangId] = (reqCount[p.barangId]||0) + p.qty; });
  const top      = Object.entries(reqCount).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const rankCls  = ['r1','r2','r3','r4','r5'];

  document.getElementById('top-barang').innerHTML = top.length === 0
    ? '<p style="color:var(--text-3);font-size:13px">Belum ada data permintaan.</p>'
    : top.map(([id,qty],i) => {
        const b = getBarang(Number(id));
        return `<div class="top-item">
          <div class="top-rank ${rankCls[i]}">#${i+1}</div>
          <div class="top-info">
            <div class="top-name">${b ? b.nama : '—'}</div>
            <div class="top-cat">${b ? b.kategori : ''}</div>
          </div>
          <div class="top-qty">${qty} <span style="font-size:12px;color:var(--text-3);font-family:var(--ff-body)">${b?b.satuan:''}</span></div>
        </div>`;
      }).join('');
}

function drawDonut(data, colors, labels) {
  const canvas = document.getElementById('donut-canvas');
  const ctx = canvas.getContext('2d');
  const total = data.reduce((a,v)=>a+v,0);
  if (total === 0) return;
  const cx = 70, cy = 70, r = 55, inner = 32;
  ctx.clearRect(0,0,140,140);
  let angle = -Math.PI/2;
  data.forEach((v,i) => {
    const slice = (v/total)*Math.PI*2;
    ctx.beginPath();
    ctx.moveTo(cx,cy);
    ctx.arc(cx,cy,r,angle,angle+slice);
    ctx.closePath();
    ctx.fillStyle = colors[i];
    ctx.fill();
    angle += slice;
  });
  ctx.beginPath();
  ctx.arc(cx,cy,inner,0,Math.PI*2);
  ctx.fillStyle = '#1F2937';
  ctx.fill();
  ctx.fillStyle = '#F9FAFB';
  ctx.font = 'bold 18px Syne,sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(total, cx, cy);

  document.getElementById('donut-legend').innerHTML = data.map((v,i)=>`
    <div class="legend-item">
      <div class="legend-dot" style="background:${colors[i]}"></div>
      <span>${labels[i]}: <strong style="color:var(--text-1)">${v}</strong></span>
    </div>
  `).join('');
}

/* ═══════════════════════════════════════════════════
   MODAL FORMS
═══════════════════════════════════════════════════ */
function openModal(type, data) {
  const overlay = document.getElementById('modal-overlay');
  const title   = document.getElementById('modal-title');
  const body    = document.getElementById('modal-body');
  const footer  = document.getElementById('modal-footer');

  if (type === 'tambah-barang') {
    title.textContent = '➕ Tambah Barang Baru';
    body.innerHTML = `
      <div class="form-grid">
        <div class="form-group">
          <label class="form-label">Kode Barang</label>
          <input class="form-input" id="f-kode" placeholder="ATK-0XX" value="ATK-0${String(nextBarangId).padStart(2,'0')}">
        </div>
        <div class="form-group">
          <label class="form-label">Kategori</label>
          <select class="form-select" id="f-kat">
            <option>Kertas</option><option>Alat Tulis</option>
            <option>File & Arsip</option><option>Elektronik</option><option>Lainnya</option>
          </select>
        </div>
        <div class="form-group span-2">
          <label class="form-label">Nama Barang</label>
          <input class="form-input" id="f-nama" placeholder="Contoh: Buku Tulis A5">
        </div>
        <div class="form-group">
          <label class="form-label">Stok Awal</label>
          <input class="form-input" id="f-stok" type="number" min="0" value="0">
        </div>
        <div class="form-group">
          <label class="form-label">Min. Stok</label>
          <input class="form-input" id="f-min" type="number" min="1" value="5">
        </div>
        <div class="form-group">
          <label class="form-label">Satuan</label>
          <input class="form-input" id="f-satuan" placeholder="Pcs / Rim / Lusin">
        </div>
        <div class="form-group">
          <label class="form-label">Harga Satuan (Rp)</label>
          <input class="form-input" id="f-harga" type="number" min="0" value="0">
        </div>
      </div>`;
    footer.innerHTML = `
      <button class="btn" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="simpanBarang()">💾 Simpan</button>`;

  } else if (type === 'edit-barang') {
    const b = data;
    title.textContent = '✏️ Edit Barang';
    body.innerHTML = `
      <div class="form-grid">
        <div class="form-group">
          <label class="form-label">Kode Barang</label>
          <input class="form-input" id="f-kode" value="${b.kode}">
        </div>
        <div class="form-group">
          <label class="form-label">Kategori</label>
          <select class="form-select" id="f-kat">
            ${['Kertas','Alat Tulis','File & Arsip','Elektronik','Lainnya'].map(k=>`<option${k===b.kategori?' selected':''}>${k}</option>`).join('')}
          </select>
        </div>
        <div class="form-group span-2">
          <label class="form-label">Nama Barang</label>
          <input class="form-input" id="f-nama" value="${b.nama}">
        </div>
        <div class="form-group">
          <label class="form-label">Min. Stok</label>
          <input class="form-input" id="f-min" type="number" value="${b.min}">
        </div>
        <div class="form-group">
          <label class="form-label">Satuan</label>
          <input class="form-input" id="f-satuan" value="${b.satuan}">
        </div>
        <div class="form-group span-2">
          <label class="form-label">Harga Satuan (Rp)</label>
          <input class="form-input" id="f-harga" type="number" value="${b.harga}">
        </div>
      </div>`;
    footer.innerHTML = `
      <button class="btn" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="updateBarang(${b.id})">💾 Simpan Perubahan</button>`;

  } else if (type === 'update-stok') {
    const b = data;
    title.textContent = `📥 Update Stok: ${b.nama}`;
    body.innerHTML = `
      <div style="background:rgba(0,0,0,0.2);border-radius:var(--radius);padding:14px 16px;margin-bottom:18px;font-size:13px;color:var(--text-2)">
        Stok saat ini: <strong style="color:var(--text-1);font-size:20px;font-family:var(--ff-display)">${b.stok}</strong> ${b.satuan}
        &nbsp;·&nbsp; Min. Stok: <strong>${b.min}</strong>
      </div>
      <div class="form-grid">
        <div class="form-group">
          <label class="form-label">Jenis Transaksi</label>
          <select class="form-select" id="f-jenis">
            <option value="Masuk">📥 Masuk (Tambah Stok)</option>
            <option value="Keluar">📤 Keluar (Kurangi Stok)</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Jumlah</label>
          <input class="form-input" id="f-qty" type="number" min="1" value="1">
        </div>
        <div class="form-group span-2">
          <label class="form-label">Keterangan</label>
          <input class="form-input" id="f-ket" placeholder="Contoh: Pembelian rutin bulanan">
        </div>
      </div>`;
    footer.innerHTML = `
      <button class="btn" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="updateStok(${b.id})">💾 Simpan Transaksi</button>`;

  } else if (type === 'buat-permintaan') {
    title.textContent = '📤 Buat Permintaan ATK';
    body.innerHTML = `
      <div class="form-grid">
        <div class="form-group">
          <label class="form-label">Nama Pemohon</label>
          <input class="form-input" id="f-pemohon" placeholder="Nama lengkap Anda">
        </div>
        <div class="form-group">
          <label class="form-label">Divisi / Departemen</label>
          <input class="form-input" id="f-divisi" placeholder="Contoh: Marketing">
        </div>
        <div class="form-group span-2">
          <label class="form-label">Barang yang Diminta</label>
          <select class="form-select" id="f-barang-req">
            ${barang.map(b=>`<option value="${b.id}">${b.nama} (Stok: ${b.stok} ${b.satuan})</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Jumlah</label>
          <input class="form-input" id="f-qty-req" type="number" min="1" value="1">
        </div>
        <div class="form-group">
          <label class="form-label">Tanggal</label>
          <input class="form-input" id="f-tgl-req" type="date" value="${today()}">
        </div>
        <div class="form-group span-2">
          <label class="form-label">Keperluan / Keterangan</label>
          <textarea class="form-textarea" id="f-ket-req" placeholder="Jelaskan keperluan pengambilan barang..."></textarea>
        </div>
      </div>`;
    footer.innerHTML = `
      <button class="btn" onclick="closeModal()">Batal</button>
      <button class="btn btn-gold" onclick="simpanPermintaan()">📤 Kirim Permintaan</button>`;

  } else if (type === 'catat-transaksi') {
    title.textContent = '✏️ Catat Transaksi Baru';
    body.innerHTML = `
      <div class="form-grid">
        <div class="form-group span-2">
          <label class="form-label">Barang</label>
          <select class="form-select" id="f-barang-trx">
            ${barang.map(b=>`<option value="${b.id}">${b.nama} (${b.kode}) - Stok: ${b.stok} ${b.satuan}</option>`).join('')}
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Jenis Transaksi</label>
          <select class="form-select" id="f-jenis-trx">
            <option value="Masuk">📥 Masuk</option>
            <option value="Keluar">📤 Keluar</option>
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Jumlah</label>
          <input class="form-input" id="f-qty-trx" type="number" min="1" value="1">
        </div>
        <div class="form-group">
          <label class="form-label">Pelaku</label>
          <input class="form-input" id="f-pelaku-trx" placeholder="Nama Admin/Karyawan">
        </div>
        <div class="form-group">
          <label class="form-label">Tanggal</label>
          <input class="form-input" id="f-tgl-trx" type="date" value="${today()}">
        </div>
        <div class="form-group span-2">
          <label class="form-label">Keterangan</label>
          <input class="form-input" id="f-ket-trx" placeholder="Keterangan transaksi...">
        </div>
      </div>`;
    footer.innerHTML = `
      <button class="btn" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" onclick="simpanTransaksi()">💾 Simpan</button>`;
  }

  overlay.classList.add('open');
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('open');
}

function handleOverlayClick(e) {
  if (e.target === document.getElementById('modal-overlay')) closeModal();
}

/* ═══════════════════════════════════════════════════
   CRUD OPERATIONS
═══════════════════════════════════════════════════ */
function simpanBarang() {
  const nama = document.getElementById('f-nama').value.trim();
  if (!nama) { toast('Nama barang wajib diisi', 'error'); return; }
  barang.push({
    id:       nextBarangId++,
    kode:     document.getElementById('f-kode').value.trim() || `ATK-${String(nextBarangId).padStart(3,'0')}`,
    nama,
    kategori: document.getElementById('f-kat').value,
    stok:     parseInt(document.getElementById('f-stok').value) || 0,
    min:      parseInt(document.getElementById('f-min').value) || 5,
    satuan:   document.getElementById('f-satuan').value || 'Pcs',
    harga:    parseInt(document.getElementById('f-harga').value) || 0,
  });
  closeModal(); renderStok(); renderDashboard();
  toast('Barang berhasil ditambahkan! ✅', 'success');
}

function openEditBarang(id) {
  openModal('edit-barang', getBarang(id));
}

function updateBarang(id) {
  const b    = getBarang(id);
  b.kode     = document.getElementById('f-kode').value.trim();
  b.nama     = document.getElementById('f-nama').value.trim();
  b.kategori = document.getElementById('f-kat').value;
  b.min      = parseInt(document.getElementById('f-min').value) || 5;
  b.satuan   = document.getElementById('f-satuan').value;
  b.harga    = parseInt(document.getElementById('f-harga').value) || 0;
  closeModal(); renderStok();
  toast('Data barang diperbarui! ✅', 'success');
}

function hapusBarang(id) {
  if (!confirm('Hapus barang ini dari daftar?')) return;
  barang = barang.filter(b => b.id !== id);
  renderStok(); renderDashboard();
  toast('Barang dihapus.', 'warning');
}

function openUpdateStok(id) {
  openModal('update-stok', getBarang(id));
}

function updateStok(id) {
  const b     = getBarang(id);
  const jenis = document.getElementById('f-jenis').value;
  const qty   = parseInt(document.getElementById('f-qty').value) || 0;
  const ket   = document.getElementById('f-ket').value || '-';
  if (jenis === 'Keluar' && qty > b.stok) { toast('Stok tidak mencukupi!', 'error'); return; }
  b.stok += jenis === 'Masuk' ? qty : -qty;
  riwayat.push({
    id: nextTrxId++,
    noTrx: `TRX-${String(nextTrxNum++).padStart(3,'0')}`,
    tgl: today(), jenis, barangId: id, qty, pelaku: 'Admin', ket
  });
  closeModal(); renderStok(); renderDashboard();
  toast(`Stok ${jenis === 'Masuk' ? 'ditambahkan' : 'dikurangi'} ${qty} unit ✅`, 'success');
}

function simpanPermintaan() {
  const pemohon = document.getElementById('f-pemohon').value.trim();
  if (!pemohon) { toast('Nama pemohon wajib diisi', 'error'); return; }
  permintaan.push({
    id:       nextReqId++,
    noReq:    `REQ-${String(nextReqNum++).padStart(3,'0')}`,
    tgl:      document.getElementById('f-tgl-req').value,
    pemohon,
    divisi:   document.getElementById('f-divisi').value,
    barangId: parseInt(document.getElementById('f-barang-req').value),
    qty:      parseInt(document.getElementById('f-qty-req').value) || 1,
    keperluan:document.getElementById('f-ket-req').value || '-',
    status:   'Pending'
  });
  closeModal();
  navigate('permintaan', document.querySelector('[onclick*="permintaan"]'));
  toast('Permintaan berhasil dikirim! 📤', 'success');
}

function setujuPermintaan(id) {
  const p = permintaan.find(x => x.id === id);
  const b = getBarang(p.barangId);
  if (p.qty > b.stok) { toast('Stok tidak mencukupi untuk menyetujui permintaan ini!', 'error'); return; }
  p.status = 'Approved';
  b.stok  -= p.qty;
  riwayat.push({
    id: nextTrxId++,
    noTrx: `TRX-${String(nextTrxNum++).padStart(3,'0')}`,
    tgl: today(), jenis: 'Keluar', barangId: p.barangId, qty: p.qty, pelaku: p.pemohon, ket: p.keperluan
  });
  renderPermintaan(); renderDashboard();
  toast(`Permintaan ${p.noReq} disetujui ✅`, 'success');
}

function tolakPermintaan(id) {
  permintaan.find(x => x.id === id).status = 'Rejected';
  renderPermintaan(); renderDashboard();
  toast('Permintaan ditolak.', 'warning');
}

function simpanTransaksi() {
  const barangId = parseInt(document.getElementById('f-barang-trx').value);
  const b        = getBarang(barangId);
  const jenis    = document.getElementById('f-jenis-trx').value;
  const qty      = parseInt(document.getElementById('f-qty-trx').value) || 0;
  const pelaku   = document.getElementById('f-pelaku-trx').value.trim() || 'Admin';
  const ket      = document.getElementById('f-ket-trx').value || '-';
  if (jenis === 'Keluar' && qty > b.stok) { toast('Stok tidak mencukupi!', 'error'); return; }
  b.stok += jenis === 'Masuk' ? qty : -qty;
  riwayat.push({
    id: nextTrxId++,
    noTrx: `TRX-${String(nextTrxNum++).padStart(3,'0')}`,
    tgl: document.getElementById('f-tgl-trx').value,
    jenis, barangId, qty, pelaku, ket
  });
  closeModal(); renderRiwayat(); renderDashboard();
  toast('Transaksi berhasil dicatat! ✅', 'success');
}

/* ═══════════════════════════════════════════════════
   TOAST
═══════════════════════════════════════════════════ */
function toast(msg, type = 'success') {
  const container = document.getElementById('toast-container');
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  const icons = { success:'✅', error:'❌', warning:'⚠️' };
  el.innerHTML = `<span>${icons[type]||'ℹ️'}</span> ${msg}`;
  container.appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    el.style.transition = 'opacity 0.3s';
    setTimeout(() => el.remove(), 300);
  }, 3000);
}

/* ═══════════════════════════════════════════════════
   CLOCK
═══════════════════════════════════════════════════ */
function updateClock() {
  const now = new Date();
  document.getElementById('clock').textContent =
    now.toLocaleTimeString('id-ID', {hour:'2-digit', minute:'2-digit', second:'2-digit'});
}
setInterval(updateClock, 1000);
updateClock();

/* ═══════════════════════════════════════════════════
   SIDEBAR TOGGLE (mobile)
═══════════════════════════════════════════════════ */
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

function checkMobile() {
  const isMobile = window.innerWidth <= 900;
  document.getElementById('menu-btn').style.display = isMobile ? 'flex' : 'none';
}
window.addEventListener('resize', checkMobile);
checkMobile();

/* ═══════════════════════════════════════════════════
   INIT
═══════════════════════════════════════════════════ */
renderDashboard();
document.getElementById('badge-pending').textContent = permintaan.filter(p => p.status === 'Pending').length;
