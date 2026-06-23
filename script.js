// ════ STATE & CONSTANTS ═════════
const DEFAULT_PESERTA = [
  { nama: 'ISKANDAR, S.Sos.', jabatan: 'Ketua' },
  { nama: 'DARKASYI ABDUL HAMID, S.Pd.', jabatan: 'Anggota' },
  { nama: 'ABDULLAH, S.Sos.', jabatan: 'Anggota' },
  { nama: 'MASRUR, MA.', jabatan: 'Anggota' },
  { nama: 'HASMUNIR, SH.', jabatan: 'Anggota' },
  { nama: 'ISWANDI, S.Sos.', jabatan: 'Sekretaris' },
  { nama: 'DAHLAN, A.Md.', jabatan: 'Kasubbag Keuangan, Umum, dan Logistik' },
  { nama: 'MASYKUR, S.Pd.I.', jabatan: 'Kasubbag Perencanaan, Data dan Informasi' },
  { nama: 'MAHMUNIR, S.Kom.', jabatan: 'Kasubbag Teknis Penyelenggaraan Pemilu, dan Hukum' },
  { nama: 'MAIMUN MAHMILUL, S.IP.', jabatan: 'Kasubbag Parmas dan SDM' },
  { nama: 'ISNAINI, SE.', jabatan: 'Analis Pengelola Keuangan APBN Ahli Muda' },
  { nama: 'NURHAYATI, A.Md.', jabatan: 'Bendahara Pengeluaran' },
  { nama: 'FAZIL BASRI, S.Kom.', jabatan: 'Notulen' },
];


let currentRole = null; // 'admin' | 'user'
function isAdmin() { return currentRole === 'admin'; }

let pesertaList = JSON.parse(localStorage.getItem('sirapat_peserta') || 'null') || DEFAULT_PESERTA.map(p => ({ ...p }));
const DEFAULT_YTH = [
  'Seluruh Anggota',
  'Sekretaris',
  'Pejabat Struktural dan Pejabat Fungsional',
  'Bendahara Pengeluaran',
];
let ythList = JSON.parse(localStorage.getItem('sirapat_yth') || 'null') || [...DEFAULT_YTH];
let arsipList = JSON.parse(localStorage.getItem('sirapat_arsip') || '[]');
let settings = JSON.parse(localStorage.getItem('sirapat_settings') || 'null') || {
  instansi: 'KIP Kabupaten Pidie Jaya', kota: 'Meureudu',
  ketua: 'Iskandar', sekretaris: 'Iswandi',
  nomorFmt: '[NO]/PK.01-Und/1118/1/[TAHUN]', nomorLast: 0,
  gasUrl: '', urlUnd: '', urlAbs: '', urlRis: '', tplMode: 'auto'
};

// Paksa ganti [BULAN] → 1 jika masih ada
if (settings.nomorFmt?.includes('[BULAN]')) {
  settings.nomorFmt = settings.nomorFmt.replace('[BULAN]', '1');
  localStorage.setItem('sirapat_settings', JSON.stringify(settings));
}
let nomorBALast = parseInt(localStorage.getItem('sirapat_nomorBA') || '0');
const AUTO_PATHS = {
  und: './templates/UND_template.docx',
  abs: './templates/ABSEN_template.docx',
  ris: './templates/RISALAH_template.docx',
  ba: './templates/BA_template.docx'
};

const SLOT_DEFS = [
  {key:'undangan', label:'📨 Undangan', match:/undangan/i},
  {key:'ba',       label:'📋 Berita Acara', match:/berita|^ba_|_ba/i},
  {key:'absen',    label:'✅ Daftar Hadir', match:/absen|hadir/i},
  {key:'risalah',  label:'📝 Risalah', match:/risalah/i},
  {key:'foto',     label:'📸 Foto Dokumentasi', match:/foto|dokumentasi|whatsapp/i, multi:true},
];

const BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const HARI_ID = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const SH_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

let tplMode = settings.tplMode || 'auto';
let calYearInline, calMonthInline;
const today = new Date();
let lastGenId = null;
let lastGenBlobs = null;
let lastGenPrefix = null;
let currentModalId = null;
let uploadFiles = {};

// ════ HELPERS ═════════════════════════════════════════════════
const GAS_URL = 'https://script.google.com/macros/s/AKfycbyVyokwyLQLgL7QqEK-cTb5idjWA-_fVydbzr3t0VxFCj30r7Mw48FP3bdhRtY9VoCe/exec';
function getGasUrl() { return GAS_URL; }

function tahunTerbilang(tahun) {
  const satuan = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan',
    'Sepuluh', 'Sebelas', 'Dua Belas', 'Tiga Belas', 'Empat Belas', 'Lima Belas', 'Enam Belas',
    'Tujuh Belas', 'Delapan Belas', 'Sembilan Belas'];
  const puluhan = ['', '', 'Dua Puluh', 'Tiga Puluh', 'Empat Puluh', 'Lima Puluh',
    'Enam Puluh', 'Tujuh Puluh', 'Delapan Puluh', 'Sembilan Puluh'];

  const ribuan = Math.floor(tahun / 1000);
  const ratusan = Math.floor((tahun % 1000) / 100);
  const sisa = tahun % 100;

  let hasil = '';
  if (ribuan) hasil += (ribuan === 1 ? 'Seribu' : satuan[ribuan] + ' Ribu') + ' ';
  if (ratusan) hasil += (ratusan === 1 ? 'Seratus' : satuan[ratusan] + ' Ratus') + ' ';
  if (sisa > 0 && sisa < 20) hasil += satuan[sisa];
  else if (sisa >= 20) hasil += puluhan[Math.floor(sisa / 10)] + (sisa % 10 ? ' ' + satuan[sisa % 10] : '');

  return hasil.trim();
}

function buildNomorBA(no, tgl) {
  const tahun = tgl instanceof Date ? tgl.getFullYear() : today.getFullYear();
  return `${no}/PK.01-BA/1118/2/${tahun}`;
}

// ★ FIX TIMEZONE: parse tanggal string "YYYY-MM-DD" sebagai local time
function parseTanggal(str) {
  if (!str) return new Date(NaN);
  if (str.includes('T')) return new Date(str);
  return new Date(str.replace(/-/g, '/') + ' 00:00:00');
}

function getMenitDariJam(jamStr) {
  if (!jamStr || !jamStr.includes(':')) return 0;
  const parts = jamStr.split(':');
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}
function buildNomor(no, tgl) {
  return settings.nomorFmt
    .replace('[NO]', String(no))
    .replace('[BULAN]', tgl instanceof Date ? tgl.getMonth() + 1 : no)
    .replace('[TAHUN]', tgl instanceof Date ? tgl.getFullYear() : today.getFullYear());
}
function tglGeneret() { return `${today.getDate()} ${BULAN_ID[today.getMonth()]} ${today.getFullYear()}`; }
function tglFull(d) { return `${d.getDate()} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`; }
function saveLocal() { localStorage.setItem('sirapat_arsip', JSON.stringify(arsipList)); }

function dlBlob(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

function isImage(n) { return /\.(jpe?g|png|webp|gif|bmp|svg)$/i.test(n || ''); }
function isPdf(n) { return /\.pdf$/i.test(n || ''); }
function fmtSize(b) {
  if (!b) return '';
  if (b < 1024) return b + 'B';
  if (b < 1048576) return (b / 1024).toFixed(1) + 'KB';
  return (b / 1048576).toFixed(1) + 'MB';
}

function shortFileName(name) {
  const map = {
    'undangan': 'Undangan',
    'beritaacara': 'Berita Acara',
    'absenhadir': 'Daftar Hadir',
    'risalah': 'Risalah',
  };
  const base = name.replace(/\.[^.]+$/, '').toLowerCase().replace(/[^a-z]/g, '');
  for (const [k, v] of Object.entries(map)) {
    if (base.includes(k)) return v;
  }
  // fallback: potong 22 karakter
  return name.length > 22 ? name.substring(0, 20) + '…' : name;
}

function statusLbl(s) {
  return { pending: 'Menunggu', uploading: 'Uploading...', done: 'Tersimpan', draft: 'Draft', err: 'Gagal' }[s] || s;
}
function getFileIcon(n) {
  return { pdf: '📕', doc: '📝', docx: '📝', zip: '📦', jpg: '🖼️', jpeg: '🖼️', png: '🖼️', webp: '🖼️', gif: '🖼️', bmp: '🖼️' }
  [(n || '').split('.').pop().toLowerCase()] || '📄';
}
function extractDriveId(url) {
  const m = (url || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}
function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(',')[1]);
    r.onerror = () => rej(new Error('Baca gagal'));
    r.readAsDataURL(file);
  });
}

// ════ GAS API ══════════════════════════════════════════════════
async function gasCall(action, payload = null) {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URL belum diisi');
  let res;
  if (payload) {
    res = await fetch(url, {
      method: 'POST',
      redirect: 'follow',
      body: JSON.stringify({ action, ...payload })
    });
  } else {
    res = await fetch(`${url}?action=${action}`, {
      redirect: 'follow'
    });
  }
  const d = await res.json();
  if (d.error) throw new Error(d.error);
  return d;
}
const gasGet = (action) => gasCall(action);
const gasPost = (payload) => gasCall(payload.action, payload);


// ════ SANITASI ARSIP ══════════════════════════════════════════
function sanitasiField(val, type) {
  const s = String(val || '');
  if (!s) return s;

  if (type === 'tanggal') {
    if (s.includes('T')) {
      const d = new Date(s);
      if (!isNaN(d)) {
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      }
    }
    return s;
  }

  if (type === 'jam') {
    let v = s;
    if (v.includes('T')) {
      try {
        const d = new Date(v);
        if (!isNaN(d)) {
          return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
        }
      } catch { }
    }
    if (v.length > 5 && v.includes(':')) v = v.substring(0, 5);
    if (v.includes(' ')) v = v.split(' ')[0];
    if (/^\d:\d{2}$/.test(v)) v = '0' + v;
    return v;
  }

  if (type === 'tglGeneret') {
    if (s.includes('T')) {
      try {
        const d = new Date(s);
        if (!isNaN(d)) {
          return `${d.getDate()} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`;
        }
      } catch { }
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      try { const p = s.split('-'); return `${parseInt(p[2])} ${BULAN_ID[parseInt(p[1]) - 1]} ${p[0]}`; } catch { }
    }
    return s;
  }

  return s;
}

function sanitasiArsip(list) {
  return list.map(r => ({
    ...r,
    tanggal: sanitasiField(r.tanggal, 'tanggal'),
    jam: sanitasiField(r.jam, 'jam'),
    tglGeneret: sanitasiField(r.tglGeneret, 'tglGeneret'),
    isManual: r.isManual === true,
  }));
}


// ════ SCAN FOLDER DRIVE — sinkron total dgn isi folder Drive ══
async function scanArsipDrive(id, btnEl) {
  if (!getGasUrl()) return;
  const r = arsipList.find(x => x.id === id);
  if (!r) return;
  const folderName = getFolderName(r);

  const teksAsli = btnEl ? btnEl.innerHTML : '';
  if (btnEl) {
    btnEl.disabled = true;
    btnEl.innerHTML = '<span class="cloud-spin" style="display:inline-block;vertical-align:-2px;margin-right:4px"></span>Memindai...';
  }

  try {
    const res = await gasCall('scanFolderArsip', { id, folderName });
    if (!res.success) return;

    r.uploadedFiles = res.uploadedFiles;
    saveLocal();

    const sisaLokal = (uploadFiles[id] || []).filter(f =>
      f._isDraft || f.status === 'pending' || f.status === 'uploading' || f.status === 'err'
    );
    uploadFiles[id] = [
      ...sisaLokal,
      ...res.uploadedFiles.map(f => ({ ...f, file: null, type: f.type || '', _showPreview: false }))
    ];

    renderFileList(id);
    renderArsip();
    refreshStats();
    if (btnEl) showToast(`✓ Scan selesai — ${res.uploadedFiles.length} file ditemukan`, 'success');
  } catch (e) {
    console.warn('scanArsipDrive gagal:', e);
    if (btnEl) showToast('Gagal memindai Drive', 'error');
  } finally {
    if (btnEl) { btnEl.disabled = false; btnEl.innerHTML = teksAsli; }
  }
}
// ════ STATS ═══════════════════════════════════════════════════
function animCount(el, target) {
  let cur = 0;
  const step = Math.max(1, Math.floor(target / 20));
  const t = setInterval(() => { cur = Math.min(cur + step, target); el.textContent = cur; if (cur >= target) clearInterval(t); }, 40);
}

function refreshStats() {
  const yr = today.getFullYear();
  const total = arsipList.length;
  const tiArsip = arsipList.filter(r => parseTanggal(r.tanggal).getFullYear() === yr && !r.isManual);
  const ti = tiArsip.length;
  const avg = total ? Math.round(arsipList.filter(r => !r.isManual).reduce((a, r) => a + (r.peserta || []).length, 0) / total) : 0;

  const ht = document.getElementById('hs-total'); if (ht) animCount(ht, total);
  const hy = document.getElementById('hs-tahun'); if (hy) animCount(hy, ti);
  const hn = document.getElementById('hs-arsip'); if (hn) animCount(hn, total);

  const dt = document.getElementById('dash-total'); if (dt) animCount(dt, total);
  const dy = document.getElementById('dash-tahun'); if (dy) animCount(dy, ti);
  const da = document.getElementById('dash-avg'); if (da) animCount(da, avg);
  const totalDok = arsipList.reduce((acc, r) => {
    const allFiles = [...(uploadFiles[r.id] || []), ...(r.uploadedFiles || [])];
    // Dedup by URL (lebih akurat dari nama)
    const uniqueUrls = new Map();
    allFiles.forEach(f => {
      if (f?.status === 'done' && f?.url && f?.name) {
        const ext = f.name.toLowerCase();
        if (ext.endsWith('.pdf') || isImage(f.name)) {
          uniqueUrls.set(f.url, f);
        }
      }
    });
    return acc + uniqueUrls.size;
  }, 0);
  const dd = document.getElementById('dash-dok'); if (dd) animCount(dd, totalDok);
  const dl = document.getElementById('dash-tahun-lbl'); if (dl) dl.textContent = 'Rapat ' + yr;
  const cl = document.getElementById('chart-tahun-lbl'); if (cl) cl.textContent = yr;

  const months = Array(12).fill(0);
  tiArsip.forEach(r => months[parseTanggal(r.tanggal).getMonth()]++);
  const max = Math.max(...months, 1);
  const bc = document.getElementById('bar-chart-home');
  if (bc) bc.innerHTML = months.map((n, i) =>
    `<div class="bar-group">${n > 0 ? `<div class="bar-val">${n}</div>` : ''}` +
    `<div class="bar" style="height:${Math.round(n / max * 80)}px"><div class="bar-inner" style="height:100%"></div></div>` +
    `<div class="bar-label">${SH_ID[i]}</div></div>`
  ).join('');

  renderUpNext();
  renderRisalahQuick();
  renderHealthMeter();
}
const updateHeroStats = refreshStats;
const renderDashHome = refreshStats;

// ════ NAV HAMBURGER ═══════════════════════════════════════════
function toggleDrawer() {
  const d = document.getElementById('nav-drawer');
  const t = document.getElementById('nav-toggle');
  const open = d.classList.toggle('open');
  t.textContent = open ? '✕' : '☰';
}
function closeDrawer() {
  document.getElementById('nav-drawer').classList.remove('open');
  document.getElementById('nav-toggle').textContent = '☰';
}
document.addEventListener('click', e => {
  const d = document.getElementById('nav-drawer');
  const t = document.getElementById('nav-toggle');
  if (d?.classList.contains('open') && !d.contains(e.target) && !t.contains(e.target)) closeDrawer();
});

// ════ HERO CANVAS ═════════════════════════════════════════════
(function () {
  const canvas = document.getElementById('hero-canvas'); if (!canvas) return;
  const ctx = canvas.getContext('2d'); let W, H, particles = [];
  function resize() { const h = canvas.parentElement; W = canvas.width = h.offsetWidth; H = canvas.height = h.offsetHeight; }
  function mk() { return { x: Math.random() * W, y: H + 10, r: Math.random() * 2.5 + .5, speed: Math.random() * .6 + .3, opacity: Math.random() * .6 + .2, drift: (Math.random() - .5) * .4, life: 0, maxLife: Math.random() * 160 + 80 }; }
  function draw() {
    ctx.clearRect(0, 0, W, H);
    if (particles.length < 55 && Math.random() < .35) particles.push(mk());
    particles = particles.filter(p => {
      p.y -= p.speed; p.x += p.drift; p.life++;
      const t = p.life / p.maxLife, a = t < .2 ? t / .2 : t > .8 ? (1 - t) / .2 : 1;
      ctx.beginPath(); ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(201,147,42,${p.opacity * a})`; ctx.fill();
      return p.life < p.maxLife && p.y > -10;
    });
    requestAnimationFrame(draw);
  }
  resize(); window.addEventListener('resize', resize); draw();
})();

// ════ SYNC INDICATORS ═════════════════════════════════════════
let syncTimer;
function showSync(msg, state = 'syncing') {
  const el = document.getElementById('sync-indicator');
  document.getElementById('sync-dot').className = 'sync-dot ' + state;
  document.getElementById('sync-text').textContent = msg;
  el.classList.add('show');
  clearTimeout(syncTimer);
  if (state !== 'syncing') syncTimer = setTimeout(() => el.classList.remove('show'), 3500);
}
function setHeroSync(state, msg) {
  const dot = document.getElementById('hero-sync-dot'); if (dot) dot.className = 'hero-sync-dot ' + state;
  const tx = document.getElementById('hero-sync-text'); if (tx) tx.textContent = msg;
}
function setCloudBanner(state, msg) {
  const b = document.getElementById('cloud-status-banner'); if (!b) return;
  const sp = document.getElementById('cloud-spin');
  const tx = document.getElementById('cloud-status-text');
  b.style.display = 'flex'; b.className = 'cloud-banner ' + state;
  if (sp) sp.style.display = (state === 'loading') ? 'block' : 'none';
  if (tx) tx.textContent = msg;
  if (state === 'ok' || state === 'err') setTimeout(() => { b.style.display = 'none'; }, 5000);
}

// ════ CLOUD — ARSIP ═══════════════════════════════════════════
// ★ loadArsipFromCloud TIDAK memanggil render apapun — diserahkan ke mulaiAutoSync
async function loadArsipFromCloud() {
  if (!getGasUrl()) {
    setHeroSync('err', 'GAS URL belum diisi');
    setCloudBanner('warn', 'URL GAS belum diisi — data hanya dari browser lokal.');
    return;
  }
  setHeroSync('syncing', 'Menyinkron data cloud...'); showSync('Memuat dari cloud...', 'syncing');
  try {
    const data = await gasCall('getArsip');
    const cloudArsip = sanitasiArsip(data.arsip || []);
    const cloudIds = new Set(cloudArsip.map(r => String(r.id)));
    const localOnly = arsipList.filter(r => !cloudIds.has(String(r.id)));
    arsipList = [...cloudArsip, ...localOnly];
    arsipList.sort((a, b) => String(b.id).localeCompare(String(a.id)));
    saveLocal();
    arsipList.forEach(r => {
      if (r.uploadedFiles?.length && !uploadFiles[r.id]?.length)
        uploadFiles[r.id] = r.uploadedFiles.map(f => ({ ...f, file: null, type: f.type || '', _showPreview: false }));
    });
    for (const item of localOnly) gasCall('simpanArsip', item).catch(() => { });
    setHeroSync('ok', `✓ ${arsipList.length} arsip tersinkron`);
    showSync(`${arsipList.length} arsip tersinkron`, 'ok');
    setCloudBanner('ok', `✓ ${arsipList.length} arsip dimuat dari cloud`);
    // ★ TIDAK memanggil renderArsip / renderCalInline / refreshStats di sini
  } catch (e) {
    setHeroSync('err', 'Gagal sync — menggunakan data lokal');
    setCloudBanner('err', '❌ Gagal memuat dari cloud: ' + e.message);
    showSync('Gagal sync cloud', 'err');
  }
}

// Tombol sync manual dari UI — perlu render ulang setelah selesai
async function refreshArsipCloud() {
  showSync('Menyamakan data dengan Drive...', 'syncing');
  try {
    if (getGasUrl()) await gasCall('bersihkanSemuaArsip');
  } catch (e) {
    console.warn('bersihkanSemuaArsip gagal:', e);
  }
  await loadArsipFromCloud();
  renderCalInline();
  renderArsip();
  refreshStats();
}

async function syncArsipToCloud(item) {
  if (!getGasUrl()) return;
  showSync('Menyimpan ke cloud...', 'syncing');
  try {
    await gasCall('simpanArsip', {
      ...item,
      uploadedFiles: (uploadFiles[item.id] || []).map(f => ({ name: f.name, size: f.size, status: f.status, url: f.url || null }))
    });
    showSync('Tersimpan ke cloud', 'ok');
  } catch { showSync('Gagal sync cloud', 'err'); }
}
async function hapusArsipCloud(id, folderName) {
  if (!getGasUrl()) return;
  try { await gasCall('hapusArsip', { id, folderName }); } catch { }
}

// ════ NOMOR SURAT ═════════════════════════════════════════════
// ★ fetchNomor TIDAK memanggil refreshStats — diserahkan ke mulaiAutoSync
async function fetchNomorBA() {
  const dot = document.getElementById('nomor-ba-dot');
  const inp = document.getElementById('inp-nomor-ba-manual');
  if (!getGasUrl()) {
    if (dot) dot.className = 'nomor-dot err';
    if (inp && !inp.dataset.userEdited) inp.placeholder = '— isi URL Apps Script di Pengaturan';
    return;
  }
  if (dot) dot.className = 'nomor-dot loading';
  if (inp && !inp.dataset.userEdited) { inp.value = ''; inp.placeholder = 'Membaca dari Sheets...'; }
  try {
    const data = await gasCall('getLastNomorBA');
    nomorBALast = data.lastNomorBA;
    localStorage.setItem('sirapat_nomorBA', String(nomorBALast));
    if (dot) dot.className = 'nomor-dot ok';
    if (inp && !inp.dataset.userEdited) {
      const tgl = document.getElementById('inp-tanggal')?.value;
      const d = tgl ? parseTanggal(tgl) : new Date();
      inp.value = buildNomorBA(nomorBALast + 1, d);
      inp.placeholder = 'Nomor BA...';
    }
  } catch (e) {
    if (dot) dot.className = 'nomor-dot err';
    if (inp && !inp.dataset.userEdited) {
      inp.value = buildNomorBA(nomorBALast + 1, new Date());
      inp.placeholder = '❌ Gagal — pakai nomor lokal';
    }
  }
}

async function fetchNomor() {
  const dot = document.getElementById('nomor-dot');
  const inp = document.getElementById('inp-nomor-manual');
  if (!getGasUrl()) { dot.className = 'nomor-dot err'; if (inp) inp.placeholder = '— isi URL Apps Script di Pengaturan'; return; }
  dot.className = 'nomor-dot loading';
  if (inp && !inp.dataset.userEdited) inp.value = '';  // kosongkan dulu saat loading
  if (inp) inp.placeholder = 'Membaca dari Sheets...';
  try {
    const data = await gasCall('getLastNomor');
    settings.nomorLast = data.lastNomor;
    localStorage.setItem('sirapat_settings', JSON.stringify(settings));
    dot.className = 'nomor-dot ok';
    if (inp) inp.placeholder = 'Nomor surat...';
    updateNomorPreview(); // ← isi value jika belum diedit user
  } catch {
    dot.className = 'nomor-dot err';
    if (inp && !inp.dataset.userEdited) inp.value = buildNomor(settings.nomorLast + 1, new Date());
    if (inp) inp.placeholder = '❌ Gagal — pakai nomor lokal';
  }
}

function updateNomorPreview() {
  const tgl = document.getElementById('inp-tanggal').value;
  const d = tgl ? parseTanggal(tgl) : new Date();

  // Nomor Surat
  const inp = document.getElementById('inp-nomor-manual');
  if (inp && !inp.dataset.userEdited) inp.value = buildNomor(settings.nomorLast + 1, d);

  // ★ Nomor BA
  const inpBA = document.getElementById('inp-nomor-ba-manual');
  if (inpBA && !inpBA.dataset.userEdited) inpBA.value = buildNomorBA(nomorBALast + 1, d);
}

// ════ TEMPLATE ════════════════════════════════════════════════
function setTplMode(mode, btn) {
  tplMode = mode; settings.tplMode = mode;
  document.querySelectorAll('.tpl-mode-tab').forEach(b => b.classList.remove('active')); btn.classList.add('active');
  document.querySelectorAll('.tpl-mode-panel').forEach(p => p.classList.remove('active')); document.getElementById('tpl-panel-' + mode).classList.add('active');
}
function getTemplateUrl(key) {
  return tplMode === 'auto' ? AUTO_PATHS[key] : (document.getElementById('url-' + key) || { value: '' }).value.trim();
}

// ════ TEST URLS ═══════════════════════════════════════════════
async function testUrl(url) {
  const r = await fetch(url); if (!r.ok) throw new Error('HTTP ' + r.status);
  const buf = await r.arrayBuffer(); if (buf.byteLength < 200) throw new Error('Bukan docx');
  return buf.byteLength;
}
async function testAutoUrl(key) {
  const b = document.getElementById('st-auto-' + key);
  b.className = 'tpl-badge loading'; b.textContent = 'Mengecek...';
  try { const s = await testUrl(AUTO_PATHS[key]); b.className = 'tpl-badge ok'; b.textContent = `✓ OK (${(s / 1024).toFixed(1)}KB)`; }
  catch (e) { b.className = 'tpl-badge err'; b.textContent = '✗ ' + e.message; }
}
async function testSemuaAuto() {
  await Promise.all(['und', 'abs', 'ris', 'ba'].map(k => testAutoUrl(k)));
}
async function testManualUrl(key) {
  const b = document.getElementById('badge-' + key);
  const url = (document.getElementById('url-' + key) || { value: '' }).value.trim();
  if (!url) { showToast('Masukkan URL', 'error'); return; }
  b.className = 'tpl-badge loading'; b.textContent = 'Mengecek...';
  try { const s = await testUrl(url); b.className = 'tpl-badge ok'; b.textContent = `✓ OK (${(s / 1024).toFixed(1)}KB)`; showToast('✓ OK', 'success'); }
  catch (e) { b.className = 'tpl-badge err'; b.textContent = '✗ ' + e.message; showToast('Gagal: ' + e.message, 'error'); }
}
async function testGasUrl() {
  const url = (document.getElementById('set-gas-url') || { value: '' }).value.trim();
  if (!url) { showToast('Masukkan URL GAS', 'error'); return; }
  const st = document.getElementById('gas-status');
  st.textContent = '⏳ Menguji...'; st.style.color = '#8a6010';
  settings.gasUrl = url; localStorage.setItem('sirapat_settings', JSON.stringify(settings));
  try {
    const d = await gasCall('getLastNomor');
    st.textContent = `✅ Terhubung! Nomor terakhir: ${d.lastNomor}`; st.style.color = '#2e7d32';
    showToast('GAS terhubung!', 'success'); fetchNomor();
  } catch (e) { st.textContent = '❌ Gagal: ' + e.message; st.style.color = '#c62828'; showToast('GAS gagal: ' + e.message, 'error'); }
}

// ════ CEK STATUS CACHE TEMPLATE ═══════════════════════════════
async function cekStatusCacheTemplate() {
  await initTplIDB();
  const keys = ['und', 'abs', 'ris'];
  const labels = { und: 'Undangan', abs: 'Absen', ris: 'Risalah' };
  const hasil = {};
  for (const key of keys) {
    const data = await tplCacheGet(key);
    hasil[key] = !!data;
  }
  const semuaAda = Object.values(hasil).every(Boolean);
  const msg = semuaAda
    ? '✅ Semua template ter-cache — generate bisa offline'
    : `⚠ Cache: ${keys.map(k => `${labels[k]}: ${hasil[k] ? '✓' : '✗'}`).join(' · ')}`;
  showToast(msg, semuaAda ? 'success' : 'info');
}

// ════ PESERTA — GENERATE PAGE ═════════════════════════════════
function renderPesertaGen() {
  document.getElementById('peserta-gen-grid').innerHTML = pesertaList.map((p, i) =>
    `<div class="peserta-item checked" id="pgen-${i}" onclick="togglePeserta(${i})">
      <div class="peserta-check"></div>
      <div class="peserta-info"><div class="peserta-nama">${p.nama}</div><div class="peserta-jabatan">${p.jabatan}</div></div>
    </div>`
  ).join('');
}
function renderYthGen() {
  const el = document.getElementById('yth-gen-grid');
  if (!el) return;
  el.innerHTML = ythList.map((y, i) =>
    `<div class="peserta-item checked" id="ygen-${i}" onclick="toggleYth(${i})">
      <div class="peserta-check"></div>
      <div class="peserta-info">
        <div class="peserta-nama">${y}</div>
        <div class="peserta-jabatan">Penerima Undangan</div>
      </div>
    </div>`
  ).join('');
}
function toggleYth(i) { document.getElementById('ygen-' + i)?.classList.toggle('checked'); }
function getCheckedYth() {
  const result = [];
  let urut = 1;
  ythList.forEach((y, i) => {
    if (document.getElementById('ygen-' + i)?.classList.contains('checked')) {
      result.push({ no: String(urut++), namaYth: y });
    }
  });
  return result;
}
function togglePeserta(i) { document.getElementById('pgen-' + i).classList.toggle('checked'); }
function getCheckedPeserta() { return pesertaList.filter((_, i) => document.getElementById('pgen-' + i)?.classList.contains('checked')); }
function pilihAgenda(el, text) {
  document.querySelectorAll('.agenda-chip').forEach(c => c.classList.remove('active')); el.classList.add('active');
  const ta = document.getElementById('inp-agenda');
  ta.value = text; ta.disabled = !!text; if (!text) { ta.disabled = false; ta.focus(); }
}

// ════ KALENDER ════════════════════════════════════════════════
function initCalInline() { calYearInline = today.getFullYear(); calMonthInline = today.getMonth(); }
function changeMonthInline(d) {
  calMonthInline += d;
  if (calMonthInline < 0) { calMonthInline = 11; calYearInline--; }
  if (calMonthInline > 11) { calMonthInline = 0; calYearInline++; }
  renderCalInline();
}
function renderCalInline() {
  if (calYearInline === undefined) initCalInline();
  document.getElementById('cal-title-inline').textContent = `${BULAN_ID[calMonthInline]} ${calYearInline}`;

  const rapatMap = {};
  arsipList.forEach(r => { (rapatMap[r.tanggal] ??= []).push({ jam: r.jam, tempat: r.tempat, agenda: r.agenda }); });

  const selTgl = document.getElementById('inp-tanggal').value;
  const selJam = document.getElementById('inp-jam').value;
  const selTempat = document.getElementById('inp-tempat').value.trim();
  const firstDay = new Date(calYearInline, calMonthInline, 1).getDay();
  const days = new Date(calYearInline, calMonthInline + 1, 0).getDate();
  const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

  let html = HARI_ID.map((_, i) => `<div class="cal-day-name">${['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'][i]}</div>`).join('');
  for (let i = 0; i < firstDay; i++) html += `<div class="cal-day other-month"></div>`;

  for (let d = 1; d <= days; d++) {
    const ds = `${calYearInline}-${String(calMonthInline + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const events = rapatMap[ds] || [];
    const hasEv = events.length > 0;
    const isBooked = hasEv && ds === selTgl && events.some(e =>
      e.tempat === selTempat && Math.abs(getMenitDariJam(selJam) - getMenitDariJam(e.jam)) < 60
    );

    let cls = 'cal-day';
    if (ds === todayStr) cls += ' today'; else if (ds === selTgl) cls += ' selected';
    cls += isBooked ? ' booked' : hasEv ? ' has-event' : '';

    const tip = hasEv ? events.map(e => `${e.jam} – ${(e.agenda || '').substring(0, 25)}`).join(' | ') : '';

    let badgeHtml = '';
    if (hasEv) {
      const isKonflik = ds === selTgl && events.some(e =>
        e.tempat === selTempat && Math.abs(getMenitDariJam(selJam) - getMenitDariJam(e.jam)) < 60
      );
      badgeHtml = `<div class="cal-badge ${isKonflik ? 'konflik' : ''}">${events.length} Rapat</div>`;
    }

    html += `<div class="${cls}" ${tip ? `data-tip="${tip.replace(/"/g, '&quot;')}"` : ''} onclick="calClickInline('${ds}')">
      <span style="margin-bottom:2px">${d}</span>
      ${badgeHtml}
    </div>`;
  }
  document.getElementById('cal-grid-inline').innerHTML = html;
}

function calClickInline(ds) {
  const events = arsipList.filter(r => r.tanggal === ds);
  if (events.length > 0) { showModalMultiple(ds, events); return; }
  document.getElementById('inp-tanggal').value = ds;
  updateNomorPreview(); checkBooking();
  showPage('generate', document.querySelectorAll('.nav-btn')[1]);
}

function showModalMultiple(ds, events) {
  const d = parseTanggal(ds);
  document.getElementById('modal-title').textContent = `Jadwal: ${d.getDate()} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`;
  const shareBtn = document.getElementById('modal-share-btn');
  if (shareBtn) shareBtn.style.display = 'none';

  let html = '<div class="meeting-list-container">';
  events.forEach(r => {
    html += `
      <div class="meeting-card">
        <div class="mc-header">
          <div class="mc-time">🕒 ${r.jam} WIB</div>
          <div class="mc-participants">👥 ${(r.peserta || []).length} Peserta</div>
        </div>
        <div class="mc-agenda">${r.agenda}</div>
        <div class="mc-location">📍 ${r.tempat}</div>
        <button class="mc-action" onclick="showArsipDetail(${r.id})">
          Kelola Dokumen Rapat <span>➔</span>
        </button>
      </div>`;
  });
  html += '</div>';
  document.getElementById('modal-body').innerHTML = html;
  document.getElementById('modal-overlay').classList.add('open');
}

function checkBooking() {
  const tgl = document.getElementById('inp-tanggal').value;
  const jam = document.getElementById('inp-jam').value;
  const tempat = document.getElementById('inp-tempat').value.trim();
  const warn = document.getElementById('booking-warn');
  const btnGen = document.getElementById('btn-gen');

  if (!jam || jam.length < 5) {
    warn.style.display = 'none'; warn.innerHTML = '';
    if (btnGen) { btnGen.disabled = false; btnGen.style.opacity = ''; btnGen.title = ''; }
    return;
  }

  const menitBaru = getMenitDariJam(jam);
  const konflik = arsipList.find(r => {
    if (r.tanggal !== tgl) return false;
    if (tempat && r.tempat !== tempat) return false;
    const menitLama = getMenitDariJam(r.jam);
    return Math.abs(menitBaru - menitLama) < 60;
  });

  if (konflik) {
    warn.style.display = 'block';
    warn.innerHTML =
      `⚠ <strong>Konflik jadwal!</strong> Jarak antar rapat minimal 1 Jam.<br>` +
      `<span style="font-size:11px;opacity:.85">Sudah ada rapat pukul <strong>${konflik.jam} WIB</strong> (<em>${konflik.agenda.substring(0, 60)}${konflik.agenda.length > 60 ? '...' : ''}</em>)</span><br>` +
      `<span style="font-size:11px;opacity:.7">Silakan ubah jam (berikan jeda minimal 1 jam) atau ganti ruangan.</span>`;
    if (btnGen) { btnGen.disabled = true; btnGen.style.opacity = '0.45'; btnGen.title = 'Ada konflik jadwal'; }
  } else {
    warn.style.display = 'none'; warn.innerHTML = '';
    if (btnGen) { btnGen.disabled = false; btnGen.style.opacity = ''; btnGen.title = ''; }
  }
}

// ════ TEMPLATE CACHE (IndexedDB) ══════════════════════════════
const TPL_IDB_NAME = 'documeet_templates_v1';
let tplIdb = null;

async function initTplIDB() {
  if (tplIdb) return tplIdb;
  return new Promise((res, rej) => {
    const req = indexedDB.open(TPL_IDB_NAME, 1);
    req.onupgradeneeded = e => {
      e.target.result.createObjectStore('templates', { keyPath: 'key' });
    };
    req.onsuccess = e => { tplIdb = e.target.result; res(tplIdb); };
    req.onerror = () => rej(req.error);
  });
}

async function tplCacheSave(key, arrayBuffer) {
  if (!tplIdb) return;
  return new Promise((res, rej) => {
    const tx = tplIdb.transaction('templates', 'readwrite');
    tx.objectStore('templates').put({ key, data: arrayBuffer, ts: Date.now() });
    tx.oncomplete = res;
    tx.onerror = rej;
  });
}

async function tplCacheGet(key) {
  if (!tplIdb) return null;
  return new Promise((res, rej) => {
    const tx = tplIdb.transaction('templates', 'readonly');
    const req = tx.objectStore('templates').get(key);
    req.onsuccess = () => res(req.result ? req.result.data : null);
    req.onerror = rej;
  });
}

async function preloadTemplates() {
  if (!navigator.onLine) return;
  try {
    const keys = ['und', 'abs', 'ris', 'ba'];
    const urls = { und: getTemplateUrl('und'), abs: getTemplateUrl('abs'), ris: getTemplateUrl('ris') };
    let cached = 0;
    for (const key of keys) {
      const url = urls[key];
      if (!url) continue;
      try {
        const r = await fetch(url);
        if (!r.ok) continue;
        const buf = await r.arrayBuffer();
        await tplCacheSave(key, buf);
        cached++;
      } catch { }
    }
    if (cached > 0) console.log(`[DocuMeet] ${cached} template ter-cache offline.`);
  } catch { }
}

// ════ FETCH & INJECT — dengan fallback cache ══════════════════
async function fetchAndInject(url, data, cacheKey) {
  let arrayBuffer = null;

  if (navigator.onLine) {
    try {
      const r = await fetch(url);
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      arrayBuffer = await r.arrayBuffer();
      if (cacheKey) await tplCacheSave(cacheKey, arrayBuffer);
    } catch (e) {
      if (cacheKey) arrayBuffer = await tplCacheGet(cacheKey);
      if (!arrayBuffer) throw new Error(`Gagal fetch "${url}": ${e.message}`);
    }
  } else {
    if (cacheKey) arrayBuffer = await tplCacheGet(cacheKey);
    if (!arrayBuffer) throw new Error(
      `Offline & template "${cacheKey}" belum ter-cache. Buka aplikasi dulu saat online.`
    );
  }

  const zip = new PizZip(arrayBuffer);
  const doc = new window.docxtemplater(zip, {
    paragraphLoop: true, linebreaks: true,
    delimiters: { start: '[[', end: ']]' },
    nullGetter: () => ''
  });
  doc.render(data);
  return doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  });
}

// ════ GENERATE DOKUMEN ════════════════════════════════════════
function setPS(id, state) { const el = document.getElementById(id); if (el) el.className = 'prog-step' + (state ? ' ' + state : ''); }

async function generateDokumen() {
  if (!isAdmin()) { showToast('⛔ Hanya Admin yang bisa generate dokumen.', 'error'); return; }
  const tanggalVal = document.getElementById('inp-tanggal').value;
  const jamVal = document.getElementById('inp-jam').value;
  const tempat = document.getElementById('inp-tempat').value.trim();
  const agenda = document.getElementById('inp-agenda').value.trim();

  if (!tanggalVal) { showToast('Pilih tanggal rapat!', 'error'); return; }
  if (!agenda) { showToast('Isi agenda rapat!', 'error'); return; }

  const menitBaru = getMenitDariJam(jamVal);
  const konflik = arsipList.find(r => {
    if (r.tanggal !== tanggalVal || r.tempat !== tempat) return false;
    return Math.abs(menitBaru - getMenitDariJam(r.jam)) < 60;
  });
  if (konflik) {
    showToast(`❌ Konflik jadwal! "${konflik.agenda.substring(0, 45)}..." sudah terjadwal di waktu & tempat yang sama.`, 'error');
    const w = document.getElementById('booking-warn');
    if (w) { w.style.display = 'block'; w.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
    return;
  }

  const urlUnd = getTemplateUrl('und'), urlAbs = getTemplateUrl('abs'), urlRis = getTemplateUrl('ris'), urlBa = getTemplateUrl('ba');
  if (!urlUnd || !urlAbs || !urlRis) { showToast('URL template belum lengkap!', 'error'); return; }
  const pesertaHadir = getCheckedPeserta();
  if (!pesertaHadir.length) { showToast('Pilih minimal 1 peserta!', 'error'); return; }

  document.getElementById('btn-awan').classList.remove('visible');
  lastGenId = lastGenBlobs = lastGenPrefix = null;

  const tgl = parseTanggal(tanggalVal);
  const hariStr = HARI_ID[tgl.getDay()];
  const tglStr = tglFull(tgl);
  const tglGen = tglGeneret();
  const jamFmt = jamVal + ' WIB s/d Selesai';

  let nextNo = settings.nomorLast + 1;
  let nextNoBA = nomorBALast + 1;
  if (getGasUrl()) {
    try { const dBA = await gasCall('getLastNomorBA'); nextNoBA = dBA.nextNomorBA; } catch { }
  }
  if (getGasUrl()) {
    try { const d = await gasCall('getLastNomor'); nextNo = d.nextNomor; } catch { }
  }

  const inpManual = document.getElementById('inp-nomor-manual');
  const nomorManual = inpManual ? inpManual.value.trim() : '';
  const inpBAManual = document.getElementById('inp-nomor-ba-manual');
  const nomorBAManual = inpBAManual ? inpBAManual.value.trim() : '';
  const data = {
    nomorSurat: nomorManual || buildNomor(nextNo, tgl), hari: hariStr, tanggal: tglStr,
    tanggalHari: `${hariStr}, ${tglStr}`, jam: jamFmt, jamPolos: jamVal, tempat, agenda,
    ketua: settings.ketua, sekretaris: settings.sekretaris, kota: settings.kota,
    kotaTanggal: `${settings.kota}, ${tglStr}`, tahun: String(tgl.getFullYear()),
    bulan: BULAN_ID[tgl.getMonth()], instansi: settings.instansi,
    jumlahPeserta: String(pesertaHadir.length), tgl_generet: tglGen,
    yth: getCheckedYth(),
    nomorBA: nomorBAManual || buildNomorBA(nextNoBA, tgl),
    tglAngka: String(tgl.getDate()),
    tahunTerbilang: tahunTerbilang(tgl.getFullYear()),
    peserta: pesertaHadir.map((p, i) => ({
      no: String(i + 1),
      nama: p.nama,
      namaRingkas: p.nama.split(',')[0].trim()
        .split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' '),
      jabatan: p.jabatan,
      ttd: '',
      ttdNo: (i % 2 === 0) ? String(i + 1) : '          ' + String(i + 1),
    }))
  };

  const btn = document.getElementById('btn-gen');
  const sp = document.getElementById('spinner');
  const tx = document.getElementById('btn-gen-text');
  btn.disabled = true; sp.style.display = 'block';
  document.getElementById('progress-bar').style.display = 'flex';
  ['ps-fetch', 'ps-inject', 'ps-zip', 'ps-done'].forEach(id => setPS(id, ''));

  try {
    setPS('ps-fetch', 'active'); tx.textContent = 'Mengambil template...';
    let blobs;
    try {
      blobs = await Promise.all([
        fetchAndInject(urlUnd, data, 'und'),
        fetchAndInject(urlAbs, data, 'abs'),
        fetchAndInject(urlRis, data, 'ris'),
        fetchAndInject(urlBa, data, 'ba')
      ]);
    }
    catch (e) { setPS('ps-fetch', 'err'); throw e; }
    setPS('ps-fetch', 'done'); setPS('ps-inject', 'done');

    setPS('ps-zip', 'active'); tx.textContent = 'Mengunduh 4 dokumen...';
    const prefix = `Rapat_${tanggalVal.replace(/-/g, '')}`;
    dlBlob(blobs[0], `${prefix}_Undangan.docx`);
    await new Promise(r => setTimeout(r, 300));
    dlBlob(blobs[1], `${prefix}_AbsenHadir.docx`);
    await new Promise(r => setTimeout(r, 300));
    dlBlob(blobs[2], `${prefix}_Risalah.docx`);
    await new Promise(r => setTimeout(r, 300));
    dlBlob(blobs[3], `${prefix}_BeritaAcara.docx`);
    setPS('ps-zip', 'done'); setPS('ps-done', 'done');

    lastGenBlobs = blobs; lastGenPrefix = prefix;

    const arsipId = Date.now();
    const mkDraft = (blob, suffix) => ({
      file: blob, name: `Draft_${prefix}_${suffix}`, size: blob.size, status: 'pending',
      url: null, type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      _blobUrl: null, _showPreview: false, _blob: blob, _isDraft: false
    });
    uploadFiles[arsipId] = [
      mkDraft(blobs[0], 'Undangan.docx'),
      mkDraft(blobs[1], 'AbsenHadir.docx'),
      mkDraft(blobs[2], 'Risalah.docx'),
      mkDraft(blobs[3], 'BeritaAcara.docx'),
    ];

    // newItem.uploadedFiles harus tetap [] — draft hanya di uploadFiles[arsipId] (memory)
    const newItem = {
      id: arsipId, tanggal: tanggalVal, hari: hariStr, jam: jamVal, tempat, agenda,
      nomorSurat: data.nomorSurat, tglGeneret: tglGen,
      peserta: pesertaHadir.map(p => p.nama), uploadedFiles: []
    };  // ← uploadedFiles SELALU []
    arsipList.unshift(newItem); saveLocal();

    settings.nomorLast = nextNo;
    localStorage.setItem('sirapat_settings', JSON.stringify(settings));

    if (getGasUrl()) {
      const folderNameForNomor = `${String(tgl.getDate()).padStart(2, '0')} ${BULAN_ID[tgl.getMonth()]} ${tgl.getFullYear()}`;
      gasCall('simpanNomor', {
        nomorUrut: nextNo,
        nomorSurat: data.nomorSurat,
        tanggal: tanggalVal,        // kirim format YYYY-MM-DD agar bisa di-parse Date()
        agenda: agenda,
        tujuan: '',
        tglGeneret: tglGen,
        pesertaCount: pesertaHadir.length,
        folderName: folderNameForNomor  // untuk cari folder di Drive
      }).catch(() => { });
      // Simpan nomor BA ke sheet BA
      gasCall('simpanNomorBA', {
        noUrut: nextNoBA,
        nomorBA: data.nomorBA,
        tanggal: tanggalVal,
        tentang: `BA RAPAT RUTIN - ${agenda.substring(0, 60)}`,
        linkBA: ''   // bisa diupdate nanti setelah file terupload
      }).catch(() => { });
      // Update state lokal
      nomorBALast = nextNoBA;
      localStorage.setItem('sirapat_nomorBA', String(nomorBALast));

      syncArsipToCloud(newItem);
      const folderName = `${String(tgl.getDate()).padStart(2, '0')} ${BULAN_ID[tgl.getMonth()]} ${tgl.getFullYear()}`;
      setTimeout(() => uploadSemuaFile(arsipId, folderName), 500);
    }

    lastGenId = arsipId;
    document.getElementById('btn-awan').classList.add('visible');
    updateNomorPreview(); renderCalInline(); refreshStats();
    showToast('✓ 4 dokumen berhasil diunduh (Undangan, Daftar hadir, Risalah, BA)!', 'success');
  } catch (err) { console.error(err); showToast('❌ ' + err.message, 'error'); }
  finally { btn.disabled = false; sp.style.display = 'none'; tx.textContent = 'Generate 4 Dokumen'; }
}

function simpanKeAwan() {
  if (!lastGenId) { showToast('Tidak ada rapat yang baru di-generate', 'error'); return; }
  const btnBeranda = document.querySelector('.nav-menu .nav-btn') || document.querySelectorAll('.nav-btn')[0];
  showPage('beranda', btnBeranda);
  setTimeout(() => {
    try {
      const el = document.getElementById('arsip-item-' + lastGenId);
      if (el) { el.classList.add('highlight'); el.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
      showArsipDetail(lastGenId);
      const modalOverlay = document.getElementById('modal-overlay');
      if (modalOverlay) modalOverlay.classList.add('open');
    } catch (err) {
      console.error("Gagal membuka popup detail:", err);
      showToast('Terjadi kesalahan saat membuka detail rapat', 'error');
    }
  }, 300);
}

// ════ ARSIP LIST ══════════════════════════════════════════════
function renderArsip() {
  const q = (document.getElementById('search-inp')?.value || '').toLowerCase();
  const bln = document.getElementById('filter-bulan')?.value || '';
  const thn = document.getElementById('filter-tahun')?.value || '';
  const adaFilter = q || bln || thn;
  const list = arsipList
    .filter(r => {
      if (r.isManual && !adaFilter) return false;
      const d = parseTanggal(r.tanggal);
      if (bln && BULAN_ID[d.getMonth()] !== bln) return false;
      if (thn && String(d.getFullYear()) !== thn) return false;
      if (q && !JSON.stringify(r).toLowerCase().includes(q)) return false;
      return true;
    })
    .sort((a, b) => {
      // Urutkan berdasarkan tanggal+jam rapat (terbaru di atas)
      const dA = parseTanggal(a.tanggal);
      if (a.jam) dA.setHours(...a.jam.split(':').map(Number));
      const dB = parseTanggal(b.tanggal);
      if (b.jam) dB.setHours(...b.jam.split(':').map(Number));
      return dB.getTime() - dA.getTime();
    });

  const el = document.getElementById('arsip-list');
  if (!list.length) {
    el.innerHTML = `<div class="empty-state"><div class="icon">📭</div><h3>Belum ada arsip</h3><p>${getGasUrl() ? 'Data cloud kosong.' : 'Arsip muncul setelah generate pertama.'}</p></div>`;
    return;
  }
  el.innerHTML = list.map(r => {
    const d = parseTanggal(r.tanggal);
    const files = uploadFiles[r.id] || [];
    const allFiles = [...files, ...(r.uploadedFiles || [])];
    const totalCloud = new Set(allFiles.filter(f => f?.status === 'done').map(f => f.name)).size;
    const hasDraft = files.some(f => f._isDraft);

    // ── Cek kelengkapan 5 dokumen (sama persis dgn logika health meter) ──
    const doneFiles = allFiles.filter(f => f?.status === 'done' && f?.name);
    const hasUnd  = doneFiles.some(f => /undangan/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasAbs  = doneFiles.some(f => /hadir/i.test(f.name)    && f.name.toLowerCase().endsWith('.pdf'));
    const hasRis  = doneFiles.some(f => /risalah/i.test(f.name)  && f.name.toLowerCase().endsWith('.pdf'));
    const hasBa   = doneFiles.some(f => /berita/i.test(f.name)   && f.name.toLowerCase().endsWith('.pdf'));
    const hasFoto = doneFiles.some(f => isImage(f.name));
    const isLengkap = hasUnd && hasAbs && hasRis && hasBa && hasFoto;

    return `<div class="arsip-item" id="arsip-item-${r.id}" onclick="showArsipDetail(${r.id})">
      <div class="arsip-date-box"><div class="day">${d.getDate()}</div><div class="month">${SH_ID[d.getMonth()]}</div></div>
      <div class="arsip-info">
        <div class="arsip-title">
          ${r.agenda.substring(0, 60)}${r.agenda.length > 60 ? '...' : ''}
          ${isLengkap ? `<span style="display:inline-flex;align-items:center;gap:3px;font-size:9px;font-weight:700;background:rgba(46,125,50,.12);color:#2e7d32;border:1px solid rgba(46,125,50,.25);border-radius:9px;padding:1px 7px;margin-left:6px;vertical-align:middle">✓ Lengkap</span>` : ''}
        </div>
        <div class="arsip-meta">
          <span>📅 ${r.hari}, ${d.getFullYear()}</span>
          <span>🕐 ${r.jam} WIB</span>
          <span>👥 ${(r.peserta || []).length}</span>
          ${totalCloud ? `<span style="color:var(--blue)">☁ ${totalCloud}</span>` : ''}
          ${hasDraft ? `<span style="color:var(--gold)">📝 draft</span>` : ''}
          ${r.nomorSurat ? `<span>${r.nomorSurat}</span>` : ''}
        </div>
      </div>
      <div class="arsip-actions" onclick="event.stopPropagation()">
        ${isAdmin() ? `<button class="btn-sm" onclick="hapusArsip(${r.id})">Hapus</button>` : ''}
      </div>
    </div>`;
  }).join('');
  if (!isAdmin()) applyRoleUI();
}

function hapusArsip(id) {
  if (!isAdmin()) { showToast('⛔ Hanya Admin yang bisa menghapus arsip.', 'error'); return; }
  const r = arsipList.find(x => x.id === id);
  if (!confirm('Hapus arsip ini? Folder & semua file di Drive untuk rapat ini juga akan dihapus (dipindah ke Trash Drive).')) return;
  const folderName = r ? getFolderName(r) : null;
  arsipList = arsipList.filter(x => x.id !== id);
  saveLocal(); delete uploadFiles[id]; hapusArsipCloud(id, folderName);
  renderArsip(); renderCalInline(); refreshStats();
  showToast('Arsip + folder Drive dihapus', 'info');
}

// ════ MODAL DETAIL ════════════════════════════════════════════
function printDetail() {
  if (!currentModalId) return;
  const r = arsipList.find(x => x.id === currentModalId); if (!r) return;
  const d = parseTanggal(r.tanggal);
  const w = window.open('', '_blank', 'width=800,height=600');
  w.document.write(`<html><head><title>Print Detail Rapat</title>
    <style>@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}</style>
    </head><body><div style="font-family:Arial,sans-serif;max-width:800px;margin:0 auto">
    <h2 style="text-align:center;border-bottom:2px solid #7a1020;color:#7a1020">Detail Arsip Rapat</h2>
    <table style="width:100%;border-collapse:collapse;margin-top:20px;font-size:14px">
      <tr><td style="padding:8px 0;font-weight:600;width:150px;color:#5a3040">Agenda</td><td>${r.agenda}</td></tr>
      <tr><td style="padding:8px 0;font-weight:600;color:#5a3040">Tanggal</td><td>${r.hari}, ${tglFull(d)}</td></tr>
      <tr><td style="padding:8px 0;font-weight:600;color:#5a3040">Pukul</td><td>${r.jam} WIB</td></tr>
      <tr><td style="padding:8px 0;font-weight:600;color:#5a3040">Tempat</td><td>${r.tempat}</td></tr>
      <tr><td style="padding:8px 0;font-weight:600;color:#5a3040">Nomor Surat</td><td>${r.nomorSurat || '-'}</td></tr>
      <tr><td style="padding:8px 0;font-weight:600;color:#5a3040">Tgl Generate</td><td>${r.tglGeneret || '-'}</td></tr>
    </table>
    <h3 style="margin-top:30px;color:#7a1020">Daftar Peserta (${(r.peserta || []).length} Orang)</h3>
    <ol style="padding-left:20px;font-size:14px;line-height:1.6">${(r.peserta || []).map(n => `<li>${n}</li>`).join('')}</ol>
    </div></body></html>`);
  w.document.close(); w.focus();
  setTimeout(() => { w.print(); w.close(); }, 250);
}

function shareFiles() {
  if (!currentModalId) return;
  const r = arsipList.find(x => x.id === currentModalId); if (!r) return;
  const allDone = [...(uploadFiles[currentModalId] || []), ...(r.uploadedFiles || [])]
    .filter(f => f?.status === 'done' && f.url)
    .reduce((acc, f) => { if (!acc.find(x => x.name === f.name)) acc.push(f); return acc; }, []);
  if (!allDone.length) { showToast('Belum ada dokumen tersimpan di Drive.', 'error'); return; }
  const d = parseTanggal(r.tanggal);
  const text = `🗂 Dokumen Rapat: ${r.agenda}\n📅 ${r.hari}, ${tglFull(d)}\n\n` +
    allDone.map(f => `📄 ${f.name}\n${f.url}`).join('\n\n');
  if (navigator.share) navigator.share({ title: 'Dokumen Rapat', text }).catch(() => { });
  else navigator.clipboard.writeText(text).then(() => showToast('✓ Link disalin ke clipboard!', 'success'))
    .catch(() => prompt('Salin teks berikut:', text));
}

function showArsipDetail(id) {
  currentModalId = id;
  const r = arsipList.find(x => x.id === id); if (!r) return;
  const d = parseTanggal(r.tanggal);
  const folderName = `${String(d.getDate()).padStart(2, '0')} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`;
  if (r.uploadedFiles?.length && !uploadFiles[id]?.length)
    uploadFiles[id] = r.uploadedFiles.map(f => ({ ...f, file: null, type: f.type || '', _showPreview: false }));
  const hasDriveFiles = [...(uploadFiles[id] || []), ...(r.uploadedFiles || [])].some(f => f?.status === 'done' && f.url);
  const shareBtn = document.getElementById('modal-share-btn');
  if (shareBtn) shareBtn.style.display = hasDriveFiles ? '' : 'none';
  document.getElementById('modal-title').textContent = 'Detail Rapat';
  document.getElementById('modal-body').innerHTML = `
    <div style="display:flex;justify-content:flex-end;margin-bottom:8px">
   ${isAdmin() ? `<button class="btn-sm" id="btn-edit-detail" onclick="toggleEditDetail(${r.id})">✏ Edit</button>` : ''}
  </div>
  <div id="detail-view-${r.id}">
    <div class="detail-row"><div class="detail-label">Tanggal</div><div class="detail-val">${r.hari}, ${tglFull(d)}</div></div>
    <div class="detail-row"><div class="detail-label">Pukul</div><div class="detail-val">${r.jam} WIB</div></div>
    <div class="detail-row"><div class="detail-label">Tempat</div><div class="detail-val">${r.tempat}</div></div>
    <div class="detail-row"><div class="detail-label">Agenda</div><div class="detail-val">${r.agenda}</div></div>
    <div class="detail-row"><div class="detail-label">Nomor Surat</div><div class="detail-val">${r.nomorSurat || '-'}</div></div>
    <div class="detail-row"><div class="detail-label">Di-generate</div><div class="detail-val">${r.tglGeneret || '-'}</div></div>
    <div class="detail-row" style="border-bottom:none"><div class="detail-label">Peserta (${(r.peserta || []).length})</div>
      <div class="detail-val">${(r.peserta || []).map((n, i) => `${i + 1}. ${n}`).join('<br>')}</div>
    </div>
  </div>
  <div id="detail-edit-${r.id}" style="display:none">
    <div class="field" style="margin-bottom:8px"><label>Tanggal</label>
      <input type="date" id="edit-tgl-${r.id}" value="${r.tanggal}"></div>
    <div class="field" style="margin-bottom:8px"><label>Pukul</label>
      <input type="time" id="edit-jam-${r.id}" value="${r.jam}"></div>
    <div class="field" style="margin-bottom:8px"><label>Tempat</label>
      <input type="text" id="edit-tempat-${r.id}" value="${r.tempat}"></div>
    <div class="field" style="margin-bottom:8px"><label>Agenda</label>
      <textarea id="edit-agenda-${r.id}" rows="3">${r.agenda}</textarea></div>
    <div class="field" style="margin-bottom:8px"><label>Nomor Surat</label>
      <input type="text" id="edit-nomor-${r.id}" value="${r.nomorSurat || ''}"></div>
    <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px">
      <button class="btn-sm" onclick="toggleEditDetail(${r.id})">Batal</button>
      <button class="btn-primary" onclick="simpanEditDetail(${r.id})">💾 Simpan</button>
    </div>
  </div>
    ${renderDraftSection(id)}
    <div class="upload-section">
      <div class="upload-section-title">☁ Upload Dokumen ke Drive <span class="folder-tag">📁 ${folderName}</span>
      <button class="btn-sm" style="margin-left:auto" onclick="scanArsipDrive(${id}, this)">🔄 Scan Drive</button>
</div>
      ${!getGasUrl() ? '<div class="no-gas-warning">⚠ URL Apps Script belum diisi di Pengaturan.</div>' : ''}
      <div class="upload-slots" style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px">
  ${SLOT_DEFS.map(s => `
    <div class="upload-zone${s.key==='foto'?' upload-zone-foto':''}" style="padding:1rem .75rem${s.key==='foto'?';grid-column:1/-1':''}"
         id="dropzone-${id}-${s.key}"
         ondrop="handleDropSlot(event,${id},'${s.key}')"
         ondragover="event.preventDefault();this.classList.add('dragover')"
         ondragleave="this.classList.remove('dragover')"
         onclick="document.getElementById('fi-${id}-${s.key}').click()">
      <div class="upload-zone-icon" style="font-size:1.25rem">${s.label.split(' ')[0]}</div>
      <div class="upload-zone-text"><strong>${s.label.split(' ').slice(1).join(' ')}</strong><br>
        <span style="font-size:10px">${s.multi ? 'klik / drop beberapa foto sekaligus' : 'klik / drop file'}</span></div>
      <input type="file" id="fi-${id}-${s.key}" accept="${s.key==='foto'?'.jpg,.jpeg,.png,.webp,.gif':'.docx,.doc,.pdf,.zip,.jpg,.jpeg,.png,.webp,.gif'}"
             ${s.multi ? 'multiple' : ''}
             onchange="handleFileInputSlot(event,${id},'${s.key}')">
    </div>`).join('')}
</div>
      <div class="uploaded-files" id="file-list-${id}"></div>
      <div class="upload-actions" id="upload-actions-${id}" style="display:none">
       ${isAdmin() ? `<button class="btn-upload-all" id="upload-btn-${id}" onclick="uploadSemuaFile(${id},'${folderName}')">☁ Upload ke Drive</button>` : '<span style="font-size:11px;color:var(--text-muted)">👁 Mode lihat saja</span>'}
      </div>
    </div>`;
  renderFileList(id);
  document.getElementById('modal-overlay').classList.add('open');
  scanArsipDrive(id);
}

function toggleEditDetail(id) {
  const view = document.getElementById(`detail-view-${id}`);
  const edit = document.getElementById(`detail-edit-${id}`);
  const btn = document.getElementById('btn-edit-detail');
  const isEditing = edit.style.display !== 'none';
  view.style.display = isEditing ? '' : 'none';
  edit.style.display = isEditing ? 'none' : '';
  btn.textContent = isEditing ? '✏ Edit' : '✕ Batal';
}

function simpanEditDetail(id) {
  if (!isAdmin()) { showToast('⛔ Akses ditolak.', 'error'); return; }
  const r = arsipList.find(x => x.id === id); if (!r) return;
  const tgl = document.getElementById(`edit-tgl-${id}`).value;
  const jam = document.getElementById(`edit-jam-${id}`).value;
  const tempat = document.getElementById(`edit-tempat-${id}`).value.trim();
  const agenda = document.getElementById(`edit-agenda-${id}`).value.trim();
  const nomor = document.getElementById(`edit-nomor-${id}`).value.trim();

  if (!tgl || !agenda) { showToast('Tanggal dan agenda wajib diisi', 'error'); return; }

  const d = parseTanggal(tgl);
  r.tanggal = tgl;
  r.hari = HARI_ID[d.getDay()];
  r.jam = jam;
  r.tempat = tempat;
  r.agenda = agenda;
  r.nomorSurat = nomor;

  saveLocal();
  syncArsipToCloud(r);
  renderArsip();
  renderCalInline();
  showToast('✓ Detail rapat diperbarui', 'success');
  // Tutup modal lalu buka ulang agar tampilan refresh
  closeModal();
  setTimeout(() => showArsipDetail(id), 200);
}

function renderDraftSection(id) {
  const drafts = (uploadFiles[id] || []).filter(f => f._isDraft && f._blob);
  if (!drafts.length) return '';
  return `<div class="draft-section">
    <div class="draft-section-title">📝 Draft Dokumen Tergenerate</div>
    <div class="draft-files">${drafts.map(f =>
    `<div class="draft-file-item">
        <div class="draft-file-icon">📝</div>
        <div class="draft-file-name">${f.name}</div>
        <span style="font-size:10px;color:var(--text-muted)">${fmtSize(f.size)}</span>
        <button class="draft-file-dl" onclick="downloadDraft(${id},${(uploadFiles[id] || []).indexOf(f)})">⬇ Unduh</button>
      </div>`).join('')}
    </div>
    <div style="font-size:10px;color:var(--text-muted);margin-top:6px">💡 File draft hanya tersedia selama sesi ini.</div>
  </div>`;
}

function downloadDraft(arsipId, fileIdx) {
  const f = (uploadFiles[arsipId] || [])[fileIdx];
  if (!f?._blob) return;
  dlBlob(f._blob, f.name);
}

function renderFileList(id) {
  const files = (uploadFiles[id] || []).filter(f => !f._isDraft);
  const el = document.getElementById(`file-list-${id}`); if (!el) return;
  const actEl = document.getElementById(`upload-actions-${id}`);
  if (actEl) actEl.style.display = files.length ? 'flex' : 'none';
  if (!files.length) { el.innerHTML = ''; return; }

  el.innerHTML = files.map(f => {
    const realIdx = (uploadFiles[id] || []).indexOf(f);
    const isDone = f.status === 'done' && f.url;
    const imgFile = isImage(f.name);
    const pdfFile = isPdf(f.name);
    let btns = '';
    if (isDone) {
      btns += `<a class="file-link" href="${f.url}" target="_blank">Buka ↗</a>`;
      btns += ` <button class="file-preview-btn" onclick="shareSingleFile('${f.url}','${f.name}')" title="Share">🔗</button>`;
    } else {
      btns += `<span class="file-status ${f.status}">${statusLbl(f.status)}</span>`;
    }
    if ((imgFile && (f._blobUrl || isDone)) || (pdfFile && isDone)) {
      btns += ` <button class="file-preview-btn${f._showPreview ? ' active' : ''}" onclick="togglePreview(${id},${realIdx})" title="${f._showPreview ? 'Tutup' : 'Lihat'} preview">👁</button>`;
      btns += ` <button class="file-preview-btn" onclick="printSingleFile(${id},${realIdx})" title="Print">🖨️</button>`;
    }
    let preview = '';
    if (f._showPreview) {
      if (imgFile) {
        const src = f._blobUrl || (isDone ? (extractDriveId(f.url) ? `https://drive.google.com/thumbnail?id=${extractDriveId(f.url)}&sz=w800` : f.url) : '');
        if (src) preview = `<div class="file-preview-area"><img src="${src}" alt="${f.name}" style="max-width:100%;max-height:360px;border-radius:8px;object-fit:contain;display:block;margin:0 auto" onerror="this.style.display='none'"></div>`;
      } else if (pdfFile && isDone) {
        const driveId = extractDriveId(f.url);
        const src = driveId ? `https://drive.google.com/file/d/${driveId}/preview` : f.url;
        preview = `<div class="file-preview-area"><iframe src="${src}" style="width:100%;height:440px;border:none;border-radius:8px" allowfullscreen loading="lazy"></iframe></div>`;
      }
    }
    return `<div class="uploaded-file-item" id="fitem-${id}-${realIdx}" style="flex-direction:column;align-items:stretch">
      <div style="display:flex;align-items:center;gap:9px">
        <div class="file-icon">${getFileIcon(f.name)}</div>
       <div class="file-name" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px" title="${f.name}">${shortFileName(f.name)}</div>
<div class="file-size">${fmtSize(f.size)}</div>
<div style="display:flex;align-items:center;gap:3px;flex-shrink:0;margin-left:auto">
          ${btns}
          ${f.status !== 'uploading' ? `<button onclick="hapusFile(${id},${realIdx})" title="Hapus" style="background:none;border:none;cursor:pointer;font-size:12px;padding:2px 5px;border-radius:3px;color:var(--text-muted)">✕</button>` : ''}
        </div>
      </div>${preview}
    </div>`;
  }).join('');
}


function togglePreview(arsipId, fileIdx) {
  const files = uploadFiles[arsipId]; if (!files?.[fileIdx]) return;
  files[fileIdx]._showPreview = !files[fileIdx]._showPreview;
  renderFileList(arsipId);
}
function revokeBlobUrl(arsipId, fileIdx) {
  const f = (uploadFiles[arsipId] || [])[fileIdx];
  if (f?._blobUrl) URL.revokeObjectURL(f._blobUrl);
}
function closeModal(e) {
  if (!e || e.target.id === 'modal-overlay') {
    document.getElementById('modal-overlay').classList.remove('open');
    currentModalId = null;
  }
}

// ════ UPLOAD ══════════════════════════════════════════════════
function handleFileInput(ev, id) { addFiles(id, Array.from(ev.target.files)); ev.target.value = ''; }
function handleDrop(ev, id) {
  ev.preventDefault();
  document.getElementById(`dropzone-${id}`)?.classList.remove('dragover');
  addFiles(id, Array.from(ev.dataTransfer.files));
}
function addFiles(id, files) {
  uploadFiles[id] ??= [];
  files.forEach(f => {
    const maxMB = f.type.startsWith('image/') ? 20 : 10;
    if (f.size > maxMB * 1024 * 1024) { showToast(`${f.name} terlalu besar (maks ${maxMB}MB)`, 'error'); return; }
    const blobUrl = f.type.startsWith('image/') ? URL.createObjectURL(f) : null;
    const entry = { file: f, name: f.name, size: f.size, status: 'pending', url: null, type: f.type || '', _blobUrl: blobUrl, _showPreview: false };
    uploadFiles[id].push(entry);
    if (!navigator.onLine) {
      const r = arsipList.find(x => x.id === id);
      const folder = r ? getFolderName(r) : String(id);
      idbSave(id, folder, entry).then(updateOfflinePendingCount);
    }
  });
  if (!navigator.onLine) showToast('📶 Offline — file disimpan, auto-upload saat online', 'info');
  renderFileList(id);
}
function renameForSlot(file, slotKey, originalName) {
  if (slotKey === 'foto') return originalName; // foto: nama asli dipertahankan
  const SLOT_PREFIX = {undangan:'Undangan', ba:'BeritaAcara', absen:'AbsenHadir', risalah:'Risalah'};
  const ext = (originalName.split('.').pop() || '').toLowerCase();
  return `${SLOT_PREFIX[slotKey]}.${ext}`;
}

function addFileToSlot(id, file, slotKey) {
  uploadFiles[id] ??= [];
  const maxMB = file.type.startsWith('image/') ? 20 : 10;
  if (file.size > maxMB * 1024 * 1024) { showToast(`${file.name} terlalu besar (maks ${maxMB}MB)`, 'error'); return; }
  const renamedName = renameForSlot(file, slotKey, file.name);
  const blobUrl = file.type.startsWith('image/') ? URL.createObjectURL(file) : null;

  // Slot foto: boleh banyak file. Slot lain: hapus file lama yang belum terupload
  const slot = SLOT_DEFS.find(s => s.key === slotKey);
  if (!slot?.multi) {
    uploadFiles[id] = uploadFiles[id].filter(f => !(f._slot === slotKey && f.status !== 'done'));
  }

  const entry = {
    file, name: renamedName, size: file.size, status: 'pending',
    url: null, type: file.type || '', _blobUrl: blobUrl, _showPreview: false, _slot: slotKey
  };
  uploadFiles[id].push(entry);

  if (!navigator.onLine) {
    const r = arsipList.find(x => x.id === id);
    const folder = r ? getFolderName(r) : String(id);
    idbSave(id, folder, entry).then(updateOfflinePendingCount);
  }
  renderFileList(id);
}

function handleFileInputSlot(ev, id, slotKey) {
  const slot = SLOT_DEFS.find(s => s.key === slotKey);
  if (slot?.multi) {
    Array.from(ev.target.files).forEach(f => addFileToSlot(id, f, slotKey));
  } else {
    const f = ev.target.files[0];
    if (f) addFileToSlot(id, f, slotKey);
  }
  ev.target.value = '';
}

function handleDropSlot(ev, id, slotKey) {
  ev.preventDefault();
  document.getElementById(`dropzone-${id}-${slotKey}`)?.classList.remove('dragover');
  const slot = SLOT_DEFS.find(s => s.key === slotKey);
  if (slot?.multi) {
    Array.from(ev.dataTransfer.files).forEach(f => addFileToSlot(id, f, slotKey));
  } else {
    const f = ev.dataTransfer.files[0];
    if (f) addFileToSlot(id, f, slotKey);
  }
}

function hapusFile(id, i) {
  if (!isAdmin()) { showToast('⛔ Akses ditolak.', 'error'); return; }
  if (!confirm('Hapus file ini? Kalau file sudah tersimpan di Drive, file aslinya juga akan dihapus.')) return;
  if (!uploadFiles[id]) return;

  const f = uploadFiles[id][i];
  const adaDiDrive = f && f.status === 'done' && f.url;

  revokeBlobUrl(id, i); uploadFiles[id].splice(i, 1); renderFileList(id);

  if (adaDiDrive && getGasUrl()) {
    gasCall('hapusFileDrive', { url: f.url }).catch(() => { });
  }

  const r = arsipList.find(x => x.id === id);
  if (r) {
    r.uploadedFiles = uploadFiles[id].filter(f => f.status === 'done').map(f => ({ name: f.name, size: f.size, status: f.status, url: f.url || null }));
    saveLocal();
    if (getGasUrl()) gasCall('updateArsipFiles', { id, uploadedFiles: r.uploadedFiles }).catch(() => { });
  }
  renderArsip(); showToast('File dihapus', 'info');
}

async function uploadSemuaFile(id, folderName) {
  if (!isAdmin()) { showToast('⛔ Hanya Admin yang bisa upload file.', 'error'); return; }
  if (!getGasUrl()) { showToast('URL Apps Script belum diisi!', 'error'); return; }
  const allFiles = uploadFiles[id] || [];
  allFiles.forEach(f => { if (f._isDraft && f._blob && f.status === 'draft') { f.file = f._blob; f.status = 'pending'; } });
  const pending = allFiles.filter(f => f.status === 'pending' || f.status === 'err');
  if (!pending.length) { showToast('Tidak ada file yang perlu diupload.', 'info'); return; }

  const btn = document.getElementById(`upload-btn-${id}`);
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Mengupload...'; }

  let ok = 0, fail = 0;

  for (let i = 0; i < allFiles.length; i++) {
    const f = allFiles[i];
    if (f.status !== 'pending' && f.status !== 'err') continue;
    if (!f.file) { allFiles[i].status = 'err'; fail++; renderFileList(id); continue; }
    allFiles[i].status = 'uploading';
    renderFileList(id);
    try {
      const res = await gasCall('uploadFile', {
        fileName: f.name,
        fileBase64: await toBase64(f.file),
        mimeType: f.file.type || 'application/octet-stream',
        folderName
      });
      if (!res.success) throw new Error(res.error || 'Unknown');
      allFiles[i].status = 'done';
      allFiles[i].url = res.fileUrl;
      allFiles[i].type = allFiles[i].type || f.file?.type || '';
      allFiles[i]._isDraft = false;
      ok++;
    } catch { allFiles[i].status = 'err'; fail++; }
    renderFileList(id);
  }

  if (btn) { btn.disabled = false; btn.textContent = '☁ Upload ke Drive'; }
  showToast(`Upload: ${ok} berhasil${fail ? `, ${fail} gagal` : ''}`, ok ? 'success' : 'error');

  const r = arsipList.find(x => x.id === id);
  if (r) {
    r.uploadedFiles = allFiles
      .filter(f => f.status === 'done' && f.url)
      .map(f => ({ name: f.name, size: f.size, status: f.status, url: f.url }));
    saveLocal();
    if (getGasUrl()) gasCall('updateArsipFiles', { id, uploadedFiles: r.uploadedFiles }).catch(() => { });
  }

  // Update link folder di sheet Surat Keluar setelah upload berhasil
  if (ok > 0 && r?.nomorSurat) {
    const nomorUrut = parseInt(r.nomorSurat);
    if (!isNaN(nomorUrut) && nomorUrut > 0) {
      gasCall('updateLinkFolder', { nomorUrut, folderName }).catch(() => { });
    }
  }

  const shareBtn = document.getElementById('modal-share-btn');
  if (shareBtn) shareBtn.style.display = allFiles.some(f => f.status === 'done' && f.url) ? '' : 'none';
  renderArsip();
  refreshStats();
}

// ════ SHARE & PRINT FILE ══════════════════════════════════════
function shareSingleFile(url, name) {
  navigator.clipboard.writeText(url)
    .then(() => showToast('🔗 Link disalin!', 'success'))
    .catch(() => prompt('Salin link file:', url));
}

function printSingleFile(arsipId, fileIdx) {
  const f = (uploadFiles[arsipId] || [])[fileIdx]; if (!f) return;
  if (isImage(f.name)) {
    const src = f._blobUrl || (f.url ? (extractDriveId(f.url) ? `https://drive.google.com/thumbnail?id=${extractDriveId(f.url)}&sz=w2000` : f.url) : '');
    if (!src) { showToast('File belum siap di-print', 'error'); return; }
    const w = window.open('', '_blank');
    w.document.write(`<html><head><title>Print - ${f.name}</title>
      <style>@media print{@page{margin:0}body{margin:0;display:flex;justify-content:center;align-items:center;height:100vh}img{max-width:100%;max-height:100vh;object-fit:contain}}</style>
      </head><body onload="setTimeout(function(){window.print();window.close()},500)">
      <img src="${src}" alt="${f.name}" style="max-width:100%;max-height:100vh;object-fit:contain">
      </body></html>`);
    w.document.close();
  } else if (isPdf(f.name) && f.url) {
    window.open(f.url, '_blank');
  } else {
    showToast('File belum siap di-print', 'error');
  }
}

// ════ PESERTA — MANAGE PAGE ═══════════════════════════════════
function renderPesertaManage() {
  document.getElementById('peserta-manage-list').innerHTML = pesertaList.map((p, i) =>
    `<div class="peserta-row" draggable="true"
         ondragstart="pDragStart(event,${i})" ondragover="pDragOver(event)"
         ondragenter="pDragEnter(event,${i})" ondragleave="pDragLeave(event)"
         ondrop="pDrop(event,${i})" ondragend="pDragEnd(event)">
      <div class="drag-handle" title="Geser untuk ubah urutan">⠿</div>
      <div class="peserta-num">${i + 1}</div>
      <input type="text" value="${p.nama}" placeholder="Nama + gelar" id="pm-nama-${i}">
      <input type="text" value="${p.jabatan}" placeholder="Jabatan" id="pm-jab-${i}" style="max-width:260px">
      <button class="btn-icon" onclick="hapusPesertaRow(${i})">✕</button>
    </div>`
  ).join('');
}

let pDragIdx = null;
function syncPesertaDOM() {
  pesertaList.forEach((_, i) => {
    const n = document.getElementById('pm-nama-' + i), j = document.getElementById('pm-jab-' + i);
    if (n) pesertaList[i].nama = n.value;
    if (j) pesertaList[i].jabatan = j.value;
  });
}
function pDragStart(e, i) { syncPesertaDOM(); pDragIdx = i; e.dataTransfer.effectAllowed = 'move'; setTimeout(() => e.target.classList.add('dragging'), 0); }
function pDragOver(e) { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; }
function pDragEnter(e, i) { e.preventDefault(); if (i !== pDragIdx) e.currentTarget.classList.add('drop-target'); }
function pDragLeave(e) { e.currentTarget.classList.remove('drop-target'); }
function pDrop(e, i) {
  e.stopPropagation(); e.currentTarget.classList.remove('drop-target');
  if (pDragIdx === null || pDragIdx === i) return;
  pesertaList.splice(i, 0, pesertaList.splice(pDragIdx, 1)[0]);
  renderPesertaManage();
  localStorage.setItem('sirapat_peserta', JSON.stringify(pesertaList));
  renderPesertaGen();
  if (getGasUrl()) gasCall('simpanPeserta', { peserta: pesertaList }).catch(() => { });
}
function pDragEnd(e) {
  e.target.classList.remove('dragging');
  document.querySelectorAll('.peserta-row').forEach(el => el.classList.remove('drop-target'));
  pDragIdx = null;
}
function tambahPeserta() { pesertaList.push({ nama: '', jabatan: '' }); renderPesertaManage(); }
function hapusPesertaRow(i) {
  if (!confirm('Hapus peserta ini?')) return;
  pesertaList.splice(i, 1); renderPesertaManage();
  showToast('Peserta dihapus', 'info');
}
function simpanPeserta() {
  if (!isAdmin()) { showToast('⛔ Akses ditolak.', 'error'); return; }
  pesertaList = pesertaList.map((_, i) => ({
    nama: document.getElementById('pm-nama-' + i)?.value || '',
    jabatan: document.getElementById('pm-jab-' + i)?.value || ''
  })).filter(p => p.nama.trim());
  localStorage.setItem('sirapat_peserta', JSON.stringify(pesertaList));
  renderPesertaGen(); showToast('Daftar peserta disimpan!', 'success');
  if (getGasUrl()) gasCall('simpanPeserta', { peserta: pesertaList }).catch(() => { });
}
function resetPeserta() {
  if (!confirm('Reset ke default?')) return;
  pesertaList = DEFAULT_PESERTA.map(p => ({ ...p }));
  localStorage.setItem('sirapat_peserta', JSON.stringify(pesertaList));
  renderPesertaManage(); renderPesertaGen(); showToast('Direset ke default.', 'info');
}

// ════ YTH LIST ════════════════════════════════════════════════
function renderYthManage() {
  const cont = document.getElementById('yth-manage-list');
  if (!cont) return;
  cont.innerHTML = ythList.map((y, i) =>
    `<div class="peserta-row" style="gap:8px">
      <div class="peserta-num">${i + 1}</div>
      <input type="text" value="${y}" placeholder="Contoh: Seluruh Anggota" id="yth-item-${i}" style="flex:1;font-family:'Inter',sans-serif;font-size:12px;padding:6px 9px;border:1.5px solid #e8ddd5;border-radius:5px;background:#fdfaf7;outline:none">
      <button class="btn-icon" onclick="hapusYthRow(${i})">✕</button>
    </div>`
  ).join('');
}
function tambahYth() {
  ythList.push(''); renderYthManage();
}
function hapusYthRow(i) {
  ythList.splice(i, 1); renderYthManage();
}
function simpanYth() {
  if (!isAdmin()) { showToast('⛔ Akses ditolak.', 'error'); return; }
  ythList = ythList.map((_, i) => document.getElementById('yth-item-' + i)?.value.trim() || '').filter(Boolean);
  localStorage.setItem('sirapat_yth', JSON.stringify(ythList));
  showToast('Daftar Yth. disimpan!', 'success');
  renderYthGen();
  if (getGasUrl()) gasCall('simpanYth', {
    yth: ythList.map((y, i) => ({ no: String(i + 1), namaYth: y }))
  }).catch(() => { });
}
function resetYth() {
  if (!confirm('Reset ke default?')) return;
  ythList = [...DEFAULT_YTH];
  localStorage.setItem('sirapat_yth', JSON.stringify(ythList));
  renderYthManage();
  renderYthGen();
  showToast('Direset ke default.', 'info');
}

// ════ ARSIP MANUAL ════════════════════════════════════════════
async function simpanArsipManual() {
  if (!isAdmin()) { showToast('⛔ Hanya Admin yang bisa menambah arsip manual.', 'error'); return; }
  const tgl = document.getElementById('manual-tgl')?.value;
  if (!tgl) { showToast('Tanggal wajib diisi', 'error'); return; }

  const nomorSurat = document.getElementById('manual-nomor')?.value.trim() || '';
  const d = parseTanggal(tgl);
  const arsipId = Date.now();
  const folderName = getFolderName({ tanggal: tgl });

  const newItem = {
    id: arsipId,
    tanggal: tgl,
    hari: HARI_ID[d.getDay()],
    jam: '10:00',
    tempat: settings.instansi || 'KIP Kabupaten Pidie Jaya',
    agenda: 'Rapat Pleno Rutin',
    nomorSurat,
    tglGeneret: tglGeneret(),
    peserta: pesertaList.map(p => p.nama),
    uploadedFiles: [],
    isManual: true
  };

  arsipList.unshift(newItem);
  saveLocal();

  // Upload file jika ada
  const fileInput = document.getElementById('manual-file-input');
  const file = fileInput?.files[0];
  uploadFiles[arsipId] = [];

  if (file) {
    const entry = {
      file, name: file.name, size: file.size,
      status: 'pending', url: null, type: file.type || '',
      _blobUrl: isImage(file.name) ? URL.createObjectURL(file) : null,
      _showPreview: false
    };
    uploadFiles[arsipId].push(entry);

    if (getGasUrl()) {
      entry.status = 'uploading';
      try {
        const res = await gasCall('uploadFile', {
          fileName: file.name,
          fileBase64: await toBase64(file),
          mimeType: file.type || 'application/octet-stream',
          folderName
        });
        if (!res.success) throw new Error(res.error);
        entry.status = 'done';
        entry.url = res.fileUrl;
        newItem.uploadedFiles = [{ name: entry.name, size: entry.size, status: 'done', url: entry.url }];
      } catch (e) {
        entry.status = 'err';
        showToast('⚠ Gagal upload file: ' + e.message, 'error');
      }
    }
    saveLocal();
  }

  if (getGasUrl()) {
    syncArsipToCloud(newItem);
  }

  // Reset form
  document.getElementById('manual-tgl').value = '';
  document.getElementById('manual-nomor').value = '';
  if (fileInput) fileInput.value = '';
  document.getElementById('manual-dropzone').querySelector('.manual-file-info').textContent = 'Klik atau drop 1 file di sini';
  document.getElementById('manual-file-name').textContent = '';

  refreshStats();
  renderCalInline();
  showToast('✓ Arsip manual tersimpan', 'success');
}

function handleManualFileInput(ev) {
  const file = ev.target.files[0]; if (!file) return;
  document.getElementById('manual-dropzone').querySelector('.manual-file-info').textContent = '✓ File dipilih';
  document.getElementById('manual-file-name').textContent = `📄 ${file.name} (${fmtSize(file.size)})`;
}

function handleManualDrop(ev) {
  ev.preventDefault();
  const dz = document.getElementById('manual-dropzone');
  dz.style.borderColor = '#d0c0c5';
  const file = ev.dataTransfer.files[0]; if (!file) return;
  const dt = new DataTransfer(); dt.items.add(file);
  document.getElementById('manual-file-input').files = dt.files;
  dz.querySelector('.manual-file-info').textContent = '✓ File dipilih';
  document.getElementById('manual-file-name').textContent = `📄 ${file.name} (${fmtSize(file.size)})`;
}

// ════ PENGATURAN ══════════════════════════════════════════════
function loadPengaturan() {
  ['instansi', 'kota', 'ketua', 'sekretaris'].forEach(k => { const el = document.getElementById('set-' + k); if (el) el.value = settings[k] || ''; });
  const fields = { nf: 'set-nomor-fmt', nl: 'set-nomor-last', gu: 'set-gas-url', uu: 'url-und', ua: 'url-abs', ur: 'url-ris' };
  const vals = { nf: settings.nomorFmt, nl: settings.nomorLast || 0, gu: settings.gasUrl, uu: settings.urlUnd, ua: settings.urlAbs, ur: settings.urlRis };
  Object.entries(fields).forEach(([k, id]) => { const el = document.getElementById(id); if (el) el.value = vals[k] || ''; });
  if (settings.tplMode === 'manual') document.querySelectorAll('.tpl-mode-tab')[1]?.click();
}
function simpanPengaturan() {
  if (!isAdmin()) { showToast('⛔ Akses ditolak.', 'error'); return; }
  ['instansi', 'kota', 'ketua', 'sekretaris'].forEach(k => { const el = document.getElementById('set-' + k); if (el) settings[k] = el.value; });
  const nf = document.getElementById('set-nomor-fmt'); if (nf) settings.nomorFmt = nf.value;
  const nl = document.getElementById('set-nomor-last'); if (nl) settings.nomorLast = parseInt(nl.value) || 0;
  const gu = document.getElementById('set-gas-url'); if (gu) settings.gasUrl = gu.value.trim();
  const uu = document.getElementById('url-und'); if (uu) settings.urlUnd = uu.value;
  const ua = document.getElementById('url-abs'); if (ua) settings.urlAbs = ua.value;
  const ur = document.getElementById('url-ris'); if (ur) settings.urlRis = ur.value;
  settings.tplMode = tplMode;
  localStorage.setItem('sirapat_settings', JSON.stringify(settings));
  showToast('Pengaturan disimpan!', 'success'); fetchNomor();
}
function toggleFaq(el) { el.classList.toggle('open'); el.nextElementSibling.classList.toggle('open'); }

// ════ NAV & TOAST ═════════════════════════════════════════════
function showPage(id, btn) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn, .nav-drawer .nav-btn').forEach(b => b.classList.remove('active'));
  const targetPage = document.getElementById('page-' + id);
  if (targetPage) targetPage.classList.add('active');
  if (btn) btn.classList.add('active');
  if (id === 'peserta') { renderPesertaManage(); renderYthManage(); }
  if (id === 'beranda') {
    renderCalInline();
    renderArsip();
    renderUpNext();
    renderRisalahQuick();
    renderHealthMeter();
  }
  if (id === 'pengaturan') loadPengaturan();
}
let toastT;
function showToast(msg, type = 'info') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = `toast ${type} show`;
  clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('show'), 4500);
}

// ════ LOGIN & AUTO SYNC ═══════════════════════════════════════
window.addEventListener('DOMContentLoaded', () => {
  const auth = sessionStorage.getItem('documeet_auth');
  const role = sessionStorage.getItem('documeet_role');
  if (auth === 'true' && role) {
    currentRole = role;
    const screen = document.getElementById('login-screen');
    if (screen) screen.style.display = 'none';
    applyRoleUI();
    mulaiAutoSync();
  }
});

async function loginAdmin() {
  const inp = document.getElementById('admin-pin').value;
  const err = document.getElementById('login-error');
  const btn = document.getElementById('btn-login');
  if (!inp) return;
  btn.textContent = 'Memeriksa PIN...'; btn.disabled = true;
  try {
    const data = await gasCall('getPin');
    if (inp === String(data.pin)) {
      currentRole = 'admin';
      sessionStorage.setItem('documeet_auth', 'true');
      sessionStorage.setItem('documeet_role', 'admin');
    } else if (inp === String(data.pinUser)) {
      currentRole = 'user';
      sessionStorage.setItem('documeet_auth', 'true');
      sessionStorage.setItem('documeet_role', 'user');
    } else {
      tampilError('❌ PIN salah! Silakan coba lagi.');
      return;
    }
    document.getElementById('login-screen').style.display = 'none';
    applyRoleUI();
    showToast(isAdmin() ? '✓ Masuk sebagai Admin' : '✓ Masuk sebagai Pengguna', 'success');
    mulaiAutoSync();
  } catch { tampilError('❌ Gagal terhubung ke Google Sheets.'); }
  function tampilError(p) {
    err.textContent = p; err.style.display = 'block';
    setTimeout(() => { err.style.display = 'none'; }, 3500);
    btn.textContent = 'Masuk'; btn.disabled = false;
  }
}

function applyRoleUI() {
  const admin = isAdmin();

  // Nav: sembunyikan menu khusus admin
  ['generate', 'peserta', 'pengaturan'].forEach(id => {
    document.querySelectorAll(`.nav-btn[onclick*="'${id}'"]`).forEach(el => {
      el.style.display = admin ? '' : 'none';
    });
  });

  // Tombol Keluar
  document.querySelectorAll(`.nav-btn[onclick*="logoutAdmin"]`).forEach(el => {
    el.style.display = admin ? '' : 'flex';
  });

  // Badge role di hero
  const badge = document.getElementById('hero-sync-badge-role');
  if (badge) badge.textContent = admin ? '👑 Admin' : '👁 Pengguna';

  // Sembunyikan tombol hapus arsip untuk user
  if (!admin) {
    document.querySelectorAll('.arsip-actions').forEach(el => el.style.display = 'none');
  }
}

//logout
function logoutAdmin() {
  if (!confirm('Keluar dari sesi?')) return;
  sessionStorage.removeItem('documeet_auth');
  location.reload();
}

// ★ SATU-SATUNYA tempat yang memanggil semua render setelah data siap
async function mulaiAutoSync() {
  // ★ OPSI B: bersihkan dulu seluruh arsip dari file yang sudah dihapus/trashed
  // di Drive SEBELUM membaca arsip — supaya UI yang ditampilkan sudah sesuai
  // kondisi Drive yang sebenarnya, tidak menunggu trigger terjadwal.
  // Tidak blocking fatal kalau gagal (mis. offline) — lanjut pakai data lama.
  setHeroSync('syncing', 'Menyamakan data dengan Drive...');
  try {
    if (getGasUrl()) await gasCall('bersihkanSemuaArsip');
  } catch (e) {
    console.warn('bersihkanSemuaArsip gagal (lanjut tanpa sync):', e);
  }

  Promise.all([
    fetchNomor(),
    fetchNomorBA(),
    loadArsipFromCloud(),
    loadPesertaFromCloud(),
    loadYthFromCloud()
  ]).then(() => {
    // Render semua hanya 1x di sini setelah semua data cloud selesai
    renderCalInline();
    renderArsip();
    refreshStats();       // sudah include renderUpNext + renderRisalahQuick + renderHealthMeter
    updateNomorPreview();
    renderPesertaGen();
    renderYthGen();
    renderYthManage();
    if (document.getElementById('page-peserta')?.classList.contains('active')) renderPesertaManage();
    checkBooking();
  });
}

async function loadPesertaFromCloud() {
  if (!getGasUrl()) return;
  try {
    const data = await gasCall('getPeserta');
    if (data.peserta?.length) {
      pesertaList = data.peserta;
      localStorage.setItem('sirapat_peserta', JSON.stringify(pesertaList));
    }
  } catch (e) { console.error('Gagal memuat peserta dari cloud:', e); }
}

async function loadYthFromCloud() {
  if (!getGasUrl()) return;
  try {
    const data = await gasCall('getYth');
    if (data.yth?.length) {
      // Simpan sebagai array of string (namaYth saja)
      ythList = data.yth.map(y => y.namaYth);
      localStorage.setItem('sirapat_yth', JSON.stringify(ythList));
      renderYthManage();
      renderYthGen();
    }
  } catch (e) { console.error('Gagal memuat Yth dari cloud:', e); }
}

// ════ AGENDA TERDEKAT (UP NEXT) ══════════════════════════════
function renderUpNext() {
  const el = document.getElementById('upnext-list'); if (!el) return;
  const todayMs = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  const nowMs = new Date().getTime();

  const upcoming = arsipList
    .filter(r => !r.isManual)
    .filter(r => {
      const d = parseTanggal(r.tanggal);
      if (r.jam) { const p = r.jam.split(':'); d.setHours(parseInt(p[0], 10), parseInt(p[1], 10), 0, 0); }
      return d.getTime() >= nowMs;
    })
    .sort((a, b) => {
      const dA = parseTanggal(a.tanggal); if (a.jam) dA.setHours(...a.jam.split(':').map(Number));
      const dB = parseTanggal(b.tanggal); if (b.jam) dB.setHours(...b.jam.split(':').map(Number));
      return dA.getTime() - dB.getTime();
    })
    .slice(0, 3);

  if (!upcoming.length) {
    el.innerHTML = '<div class="upnext-empty">📭 Tidak ada rapat mendatang.</div>';
    return;
  }

  el.innerHTML = upcoming.map((r, i) => {
    const d = parseTanggal(r.tanggal);
    const diffDay = Math.round((d.getTime() - todayMs) / 86400000);
    const labelHari = diffDay === 0 ? 'Hari ini' : diffDay === 1 ? 'Besok' : diffDay + ' hari';
    const isSoon = diffDay <= 3;
    const isFirst = i === 0;
    return `<div class="upnext-item${isFirst ? ' next' : ''}" onclick="showArsipDetail(${r.id})">
      <div class="upnext-datebox${isFirst ? '' : ' future'}">
        <span class="ud">${String(d.getDate()).padStart(2, '0')}</span>
        <span class="um">${SH_ID[d.getMonth()].toUpperCase()}</span>
      </div>
      <div class="upnext-info">
        <div class="upnext-agenda">${r.agenda.substring(0, 55)}${r.agenda.length > 55 ? '...' : ''}</div>
        <div class="upnext-meta">
          <span>🕒 ${r.jam} WIB</span>
          <span>📍 ${(r.tempat || '').substring(0, 28)}</span>
        </div>
      </div>
      <span class="upnext-badge ${isSoon ? 'soon' : 'far'}">${labelHari}</span>
    </div>`;
  }).join('');
}

// ════ RISALAH TERAKHIR QUICK ACCESS ══════════════════════════
function renderRisalahQuick() {
  const sub = document.getElementById('risalah-quick-sub');
  const st = document.getElementById('risalah-quick-status');
  if (!sub || !st) return;
  if (!arsipList.length) {
    sub.textContent = 'Belum ada arsip rapat';
    st.className = 'risalah-quick-status none'; st.textContent = '—'; return;
  }

  const nowMs = new Date().getTime();
  const sudahLewat = arsipList
    .filter(r => {
      const d = parseTanggal(r.tanggal);
      if (r.jam) { const p = r.jam.split(':'); d.setHours(parseInt(p[0], 10), parseInt(p[1], 10), 0, 0); }
      return d.getTime() <= nowMs;
    })
    .sort((a, b) => {
      const dA = parseTanggal(a.tanggal); if (a.jam) dA.setHours(...a.jam.split(':').map(Number));
      const dB = parseTanggal(b.tanggal); if (b.jam) dB.setHours(...b.jam.split(':').map(Number));
      return dB.getTime() - dA.getTime();
    });

  if (!sudahLewat.length) {
    sub.textContent = 'Belum ada rapat yang sudah berlangsung';
    st.className = 'risalah-quick-status none'; st.textContent = '—'; return;
  }

  const latest = sudahLewat[0];
  const d = parseTanggal(latest.tanggal);
  sub.textContent = `${(latest.agenda || '').substring(0, 45)}${(latest.agenda || '').length > 45 ? '...' : ''} — ${d.getDate()} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`;

  const allFiles = [...(uploadFiles[latest.id] || []), ...(latest.uploadedFiles || [])];
  const risalahFile = allFiles.find(f => f?.name && /risalah/i.test(f.name) && f.status === 'done');
  if (risalahFile) {
    st.className = 'risalah-quick-status ok'; st.textContent = '☁ Drive';
  } else {
    const hasDraft = (uploadFiles[latest.id] || []).some(f => /risalah/i.test(f.name || ''));
    st.className = 'risalah-quick-status ' + (hasDraft ? 'pending' : 'none');
    st.textContent = hasDraft ? '📝 Draft' : '—';
  }
}

function bukaRisalahTerakhir() {
  if (!arsipList.length) { showToast('Belum ada arsip rapat.', 'error'); return; }
  const nowMs = new Date().getTime();
  const sudahLewat = arsipList
    .filter(r => {
      const d = parseTanggal(r.tanggal);
      if (r.jam) { const p = r.jam.split(':'); d.setHours(parseInt(p[0], 10), parseInt(p[1], 10), 0, 0); }
      return d.getTime() <= nowMs;
    })
    .sort((a, b) => {
      const dA = parseTanggal(a.tanggal); if (a.jam) dA.setHours(...a.jam.split(':').map(Number));
      const dB = parseTanggal(b.tanggal); if (b.jam) dB.setHours(...b.jam.split(':').map(Number));
      return dB.getTime() - dA.getTime();
    });
  if (!sudahLewat.length) { showToast('Belum ada rapat yang sudah berlangsung.', 'error'); return; }
  showArsipDetail(sudahLewat[0].id);
}

// ════ HEALTH METER ════════════════════════════════════════════
function renderHealthMeter() {
  const rowsEl = document.getElementById('health-rows');
  const scoreEl = document.getElementById('health-score');
  const footEl = document.getElementById('health-footer-text');
  if (!rowsEl) return;

  const yr = today.getFullYear();
  const list = arsipList.filter(r => parseTanggal(r.tanggal).getFullYear() === yr && !r.isManual);
  const total = list.length;

  if (!total) {
    rowsEl.innerHTML = '<div style="font-size:11px;color:var(--text-muted);padding:6px 0">Belum ada arsip tahun ini.</div>';
    if (scoreEl) { scoreEl.textContent = '—'; scoreEl.className = 'health-score'; }
    if (footEl) footEl.textContent = 'Tidak ada data';
    return;
  }

  const countDonePdf = (keyword) => list.filter(r => {
    const allFiles = [...(uploadFiles[r.id] || []), ...(r.uploadedFiles || [])];
    return allFiles.some(f =>
      f?.name &&
      new RegExp(keyword, 'i').test(f.name) &&
      f.name.toLowerCase().endsWith('.pdf') &&
      f.status === 'done'
    );
  }).length;

  const undOk = countDonePdf('undangan');
  const absOk = countDonePdf('hadir');
  const risOk = countDonePdf('risalah');
  const baOk = countDonePdf('berita');
  const fotoOk = list.filter(r => {
    const allFiles = [...(uploadFiles[r.id] || []), ...(r.uploadedFiles || [])];
    return allFiles.some(f => f?.name && isImage(f.name) && f.status === 'done');
  }).length;

  const pct = u => Math.round(u / total * 100);
  const pUnd = pct(undOk), pAbs = pct(absOk), pRis = pct(risOk), pFoto = pct(fotoOk), pBa = pct(baOk);
  const overall = Math.round((pUnd + pAbs + pRis + pFoto + pBa) / 5);
  const cls = v => v >= 90 ? 'ok' : v >= 60 ? 'warn' : 'err';

  rowsEl.innerHTML = [
    { icon: '📋', label: 'Berita Acara (PDF)', ok: baOk, pct: pBa },
    { icon: '📨', label: 'Undangan (PDF)', ok: undOk, pct: pUnd },
    { icon: '✅', label: 'Daftar Hadir (PDF)', ok: absOk, pct: pAbs },
    { icon: '📝', label: 'Risalah (PDF)', ok: risOk, pct: pRis },
    { icon: '📸', label: 'Dokumentasi (Foto)', ok: fotoOk, pct: pFoto },
  ].map(row => `
    <div class="health-row">
      <div class="health-row-top">
        <span class="health-row-label">${row.icon} ${row.label}</span>
        <span class="health-row-val ${cls(row.pct)}">${row.ok}/${total}</span>
      </div>
      <div class="health-bar-bg"><div class="health-bar-fill ${cls(row.pct)}" style="width:${row.pct}%"></div></div>
    </div>`).join('');

  if (scoreEl) {
    scoreEl.textContent = overall + '%';
    scoreEl.className = 'health-score ' + (overall >= 90 ? 'high' : overall >= 60 ? 'mid' : 'low');
  }

  const belum = list.filter(r => {
    const allFiles = [...(uploadFiles[r.id] || []), ...(r.uploadedFiles || [])];
    const doneFiles = allFiles.filter(f => f?.status === 'done' && f?.name);
    const hasUnd = doneFiles.some(f => /undangan/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasAbs = doneFiles.some(f => /absen/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasRis = doneFiles.some(f => /risalah/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasBa = doneFiles.some(f => /ba/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasFoto = doneFiles.some(f => isImage(f.name));
    return !(hasUnd && hasAbs && hasRis && hasBa && hasFoto);
  }).length;

  if (footEl) footEl.textContent = belum > 0 ? `${belum} arsip belum lengkap dokumen Drive` : '✓ Semua arsip tahun ini lengkap';
}

function scrollToArsipBelum() {
  const yr = today.getFullYear();
  const lbl = document.getElementById('health-year-lbl');
  if (lbl) lbl.textContent = yr;
  const belum = arsipList.find(r => {
    if (parseTanggal(r.tanggal).getFullYear() !== yr) return false;
    const allFiles = [...(uploadFiles[r.id] || []), ...(r.uploadedFiles || [])];
    const doneFiles = allFiles.filter(f => f?.status === 'done' && f?.name);
    const hasUnd = doneFiles.some(f => /undangan/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasAbs = doneFiles.some(f => /absen/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasRis = doneFiles.some(f => /risalah/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasBa = doneFiles.some(f => /ba/i.test(f.name) && f.name.toLowerCase().endsWith('.pdf'));
    const hasFoto = doneFiles.some(f => isImage(f.name));
    return !(hasUnd && hasAbs && hasRis && hasBa && hasFoto);
  });
  if (!belum) { showToast('Semua arsip tahun ini sudah lengkap!', 'success'); return; }
  const el = document.getElementById('arsip-item-' + belum.id);
  if (el) {
    el.classList.add('highlight');
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    setTimeout(() => el.classList.remove('highlight'), 2500);
  }
  showArsipDetail(belum.id);
}

// ════ OFFLINE MODE ════════════════════════════════════════════
const IDB_NAME = 'documeet_offline_v1';
let idb = null;

async function initIDB() {
  return new Promise((res, rej) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = e => {
      e.target.result.createObjectStore('pending', { keyPath: 'uid', autoIncrement: true });
    };
    req.onsuccess = e => { idb = e.target.result; res(idb); };
    req.onerror = () => rej(req.error);
  });
}

async function idbSave(arsipId, folderName, fileEntry) {
  if (!idb || !fileEntry.file) return;
  const b64 = await toBase64(fileEntry.file);
  return new Promise((res, rej) => {
    const tx = idb.transaction('pending', 'readwrite');
    tx.objectStore('pending').add({ arsipId, folderName, name: fileEntry.name, size: fileEntry.size, type: fileEntry.type || '', b64 });
    tx.oncomplete = res; tx.onerror = rej;
  });
}

async function idbGetAll() {
  if (!idb) return [];
  return new Promise((res, rej) => {
    const tx = idb.transaction('pending', 'readonly');
    const req = tx.objectStore('pending').getAll();
    req.onsuccess = () => res(req.result);
    req.onerror = rej;
  });
}

async function idbClear() {
  if (!idb) return;
  return new Promise((res, rej) => {
    const tx = idb.transaction('pending', 'readwrite');
    tx.objectStore('pending').clear();
    tx.oncomplete = res; tx.onerror = rej;
  });
}

function getFolderName(r) {
  const d = parseTanggal(r.tanggal);
  return `${String(d.getDate()).padStart(2, '0')} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`;
}

async function updateOfflinePendingCount() {
  const items = await idbGetAll();
  const el = document.getElementById('offline-pending-count');
  if (el) el.textContent = items.length > 0 ? `• ${items.length} pending` : '';
}

function setOfflineUI(isOffline) {
  const badge = document.getElementById('offline-badge');
  if (badge) badge.style.display = isOffline ? 'flex' : 'none';
}

window.addEventListener('online', async () => {
  preloadTemplates();
  setOfflineUI(false);
  const items = await idbGetAll();
  if (!items.length) { showToast('🌐 Kembali online', 'success'); return; }
  showToast(`🌐 Online — mengupload ${items.length} file pending...`, 'info');
  let ok = 0;
  for (const item of items) {
    try {
      const res = await gasCall('uploadFile', { fileName: item.name, fileBase64: item.b64, mimeType: item.type || 'application/octet-stream', folderName: item.folderName });
      if (!res.success) throw new Error(res.error);
      const files = uploadFiles[item.arsipId] || [];
      const f = files.find(f => f.name === item.name);
      if (f) { f.status = 'done'; f.url = res.fileUrl; }
      const r = arsipList.find(x => x.id === item.arsipId);
      if (r) {
        r.uploadedFiles = files.filter(f => f.status === 'done').map(f => ({ name: f.name, size: f.size, status: 'done', url: f.url }));
        saveLocal();
        gasCall('updateArsipFiles', { id: item.arsipId, uploadedFiles: r.uploadedFiles }).catch(() => { });
      }
      renderFileList(item.arsipId);
      ok++;
    } catch { }
  }
  await idbClear();
  renderArsip();
  showToast(`✓ ${ok}/${items.length} file berhasil diupload`, 'success');
});

window.addEventListener('offline', () => {
  setOfflineUI(true);
  updateOfflinePendingCount();
  showToast('📶 Offline — file akan auto-upload saat online', 'info');
});

// ════ INIT ════════════════════════════════════════════════════
// Init IDB untuk file offline
initIDB().then(() => {
  setOfflineUI(!navigator.onLine);
  updateOfflinePendingCount();
});

// Init IDB untuk cache template, lalu preload jika online
initTplIDB().then(() => {
  if (navigator.onLine) preloadTemplates();
});

// Render awal dari data lokal (sebelum cloud selesai)
document.getElementById('inp-tanggal').value = today.toISOString().split('T')[0];
initCalInline();
renderCalInline();
renderPesertaGen();
refreshStats();
if (settings.tplMode === 'manual') { tplMode = 'manual'; document.querySelectorAll('.tpl-mode-tab')[1]?.click(); }
