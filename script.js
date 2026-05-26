// ════ STATE & CONSTANTS ════════════════════════════════════════
const DEFAULT_PESERTA = [
  {nama:'ISKANDAR, S.Sos.',        jabatan:'Ketua'},
  {nama:'DARKASYI ABDUL HAMID, S.Pd.', jabatan:'Anggota'},
  {nama:'ABDULLAH, S.Sos.',        jabatan:'Anggota'},
  {nama:'MASRUR, MA.',             jabatan:'Anggota'},
  {nama:'HASMUNIR, SH.',           jabatan:'Anggota'},
  {nama:'ISWANDI, S.Sos.',         jabatan:'Sekretaris'},
  {nama:'DAHLAN, A.Md.',           jabatan:'Kasubbag Keuangan, Umum, dan Logistik'},
  {nama:'MASYKUR, S.Pd.I.',        jabatan:'Kasubbag Perencanaan, Data dan Informasi'},
  {nama:'MAHMUNIR, S.Kom.',        jabatan:'Kasubbag Teknis Penyelenggaraan Pemilu, dan Hukum'},
  {nama:'MAIMUN MAHMILUL, S.IP.',  jabatan:'Kasubbag Keuangan, Umum dan Logistik'},
  {nama:'ISNAINI, SE.',            jabatan:'Analis Pengelola Keuangan APBN Ahli Muda'},
  {nama:'NURHAYATI, A.Md.',        jabatan:'Bendahara Pengeluaran'},
  {nama:'FAZIL BASRI, S.Kom.',     jabatan:'Notulen'},
];

let pesertaList  = JSON.parse(localStorage.getItem('sirapat_peserta') || 'null') || DEFAULT_PESERTA.map(p => ({...p}));
let arsipList    = JSON.parse(localStorage.getItem('sirapat_arsip')   || '[]');
let settings     = JSON.parse(localStorage.getItem('sirapat_settings')|| 'null') || {
  instansi:'KIP Kabupaten Pidie Jaya', kota:'Meureudu',
  ketua:'Iskandar', sekretaris:'Iswandi',
  nomorFmt:'[NO]/PK.01-Und/1118/1/[TAHUN]', nomorLast:0,
  gasUrl:'', urlUnd:'', urlAbs:'', urlRis:'', tplMode:'auto'
};

if (settings.nomorFmt?.includes('[BULAN]')) {
  settings.nomorFmt = settings.nomorFmt.replace('[BULAN]', '1');
  localStorage.setItem('sirapat_settings', JSON.stringify(settings));
}

const AUTO_PATHS = {
  und: './templates/UND_template.docx',
  abs: './templates/ABSEN_template.docx',
  ris: './templates/RISALAH_template.docx'
};
const BULAN_ID = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const HARI_ID  = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
const SH_ID    = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];

let tplMode        = settings.tplMode || 'auto';
let calYearInline, calMonthInline;
const today        = new Date();
let lastGenId      = null;
let lastGenBlobs   = null;
let lastGenPrefix  = null;
let currentModalId = null;
let uploadFiles    = {};

// ════ HELPERS & TIMEZONE FIX ══════════════════════════════════
const GAS_URL = 'https://script.google.com/macros/s/AKfycbz2CWZbBPaBBfXL1jtSQDhd65FnUAWZogzA-yl51cjxIQMFznhmgneI2G71xN593w/exec';
function getGasUrl() { return GAS_URL; }

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
    .replace('[NO]',    String(no))
    .replace('[BULAN]', tgl instanceof Date ? tgl.getMonth() + 1 : no)
    .replace('[TAHUN]', tgl instanceof Date ? tgl.getFullYear() : today.getFullYear());
}
function tglGeneret() { return `${today.getDate()} ${BULAN_ID[today.getMonth()]} ${today.getFullYear()}`; }
function tglFull(d)   { return `${d.getDate()} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`; }
function saveLocal()  { localStorage.setItem('sirapat_arsip', JSON.stringify(arsipList)); }

function dlBlob(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

function isImage(n) { return /\.(jpe?g|png|webp|gif|bmp|svg)$/i.test(n || ''); }
// Tetap fungsional untuk isPdf
function isPdf(n)   { return /\.pdf$/i.test(n || ''); }

function fmtSize(b) {
  if (!b) return '';
  if (b < 1024)    return b + 'B';
  if (b < 1048576) return (b / 1024).toFixed(1) + 'KB';
  return (b / 1048576).toFixed(1) + 'MB';
}
function statusLbl(s) {
  return {pending:'Menunggu', uploading:'Uploading...', done:'Tersimpan', draft:'Draft', err:'Gagal'}[s] || s;
}
function getFileIcon(n) {
  return {pdf:'📕',doc:'📝',docx:'📝',zip:'📦',jpg:'🖼️',jpeg:'🖼️',png:'🖼️',webp:'🖼️'}
    [(n||'').split('.').pop().toLowerCase()] || '📄';
}
function extractDriveId(url) {
  const m = (url||'').match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}
function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload  = () => res(r.result.split(',')[1]);
    r.onerror = () => rej(new Error('Baca gagal'));
    r.readAsDataURL(file);
  });
}

// ════ GAS API ═════════════════════════════════════════════════
async function gasCall(action, payload = null) {
  const url = getGasUrl();
  if (!url) throw new Error('GAS URL belum diisi');
  let res;
  if (payload) {
    res = await fetch(url, { method: 'POST', body: JSON.stringify({action, ...payload}) });
  } else {
    res = await fetch(`${url}?action=${action}`);
  }
  const d = await res.json();
  if (d.error) throw new Error(d.error);
  return d;
}
const gasGet  = (action)  => gasCall(action);
const gasPost = (payload) => gasCall(payload.action, payload);

// ════ SANITASI ARSIP ══════════════════════════════════════════
function sanitasiField(val, type) {
  const s = String(val || '');
  if (!s) return s;
  if (type === 'tanggal') {
    if (s.includes('T')) {
      const d = new Date(s);
      if (!isNaN(d)) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
    return s;
  }
  if (type === 'jam') {
    let v = s;
    if (v.includes('T')) {
      try {
        const d = new Date(v);
        if (!isNaN(d)) return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
      } catch {}
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
        if (!isNaN(d)) return `${d.getDate()} ${BULAN_ID[d.getMonth()]} ${d.getFullYear()}`;
      } catch {}
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      try { const p = s.split('-'); return `${parseInt(p[2])} ${BULAN_ID[parseInt(p[1])-1]} ${p[0]}`; } catch {}
    }
    return s;
  }
  return s;
}
function sanitasiArsip(list) {
  return list.map(r => ({
    ...r,
    tanggal:    sanitasiField(r.tanggal,    'tanggal'),
    jam:        sanitasiField(r.jam,        'jam'),
    tglGeneret: sanitasiField(r.tglGeneret, 'tglGeneret'),
  }));
}

// ════ STATS & WIDGET PIMPINAN EXECUTIF ═════════════════════════
function animCount(el, target) {
  let cur = 0;
  const step = Math.max(1, Math.floor(target / 20));
  const t = setInterval(() => { cur = Math.min(cur + step, target); el.textContent = cur; if (cur >= target) clearInterval(t); }, 40);
}

function renderExecutiveWidgets() {
  const upcomingEl = document.getElementById('upcoming-list');
  const meterFill = document.getElementById('meter-fill');
  const meterText = document.getElementById('meter-text');
  const meterSub = document.getElementById('meter-sub');
  const quickBtnContainer = document.getElementById('quick-btn-container');

  if (!arsipList.length) {
    if (upcomingEl) upcomingEl.innerHTML = '<div style="font-size:11px; color:var(--text-muted);">Belum ada jadwal rapat terdaftar.</div>';
    if (meterFill) meterFill.style.width = '0%';
    if (meterText) meterText.textContent = '0%';
    if (quickBtnContainer) quickBtnContainer.innerHTML = '<button class="quick-btn" disabled style="opacity:0.6;">📄 Belum ada risalah</button>';
    return;
  }

  // 1. Up Next Widget Logic (Terdekat)
  const nowStr = today.toISOString().split('T')[0];
  let upcomingMeetings = arsipList.filter(r => r.tanggal >= nowStr);
  upcomingMeetings.sort((a, b) => a.tanggal.localeCompare(b.tanggal));
  
  if (upcomingMeetings.length === 0) {
    upcomingMeetings = arsipList.slice(0, 2); // Jika tidak ada masa depan, tampilkan 2 riwayat terakhir
  } else {
    upcomingMeetings = upcomingMeetings.slice(0, 2);
  }

  if (upcomingEl) {
    upcomingEl.innerHTML = upcomingMeetings.map(r => {
      const isToday = r.tanggal === nowStr;
      return `
        <div class="up-next-item" style="border-left-color: ${isToday ? 'var(--maroon)' : 'var(--gold)'}">
          <div class="up-next-meta">
            <span>${isToday ? '🚨 HARI INI' : `📅 ${r.hari}`}</span>
            <span>🕒 ${r.jam} WIB</span>
          </div>
          <div class="up-next-title">${r.agenda.substring(0, 75)}${r.agenda.length > 75 ? '...' : ''}</div>
          <div style="font-size:10px; color:var(--text-muted); display:flex; justify-content:space-between; align-items:center; margin-top:4px;">
            <span>📍 ${r.tempat.substring(0,35)}${r.tempat.length > 35 ? '...' : ''}</span>
            <span style="color:var(--maroon); font-weight:600; cursor:pointer;" onclick="showArsipDetail(${r.id})">Buka Berkas ➔</span>
          </div>
        </div>
      `;
    }).join('');
  }

  // 2. Compliance Health Meter Logic
  let totalUploaded = 0;
  arsipList.forEach(r => {
    const files = [...(uploadFiles[r.id] || []), ...(r.uploadedFiles || [])];
    const uniqueDone = new Set(files.filter(f => f && f.status === 'done').map(f => f.name)).size;
    totalUploaded += Math.min(uniqueDone, 3);
  });
  const maxTarget = arsipList.length * 3;
  const pct = maxTarget > 0 ? Math.round((totalUploaded / maxTarget) * 100) : 0;

  if (meterFill) meterFill.style.width = pct + '%';
  if (meterText) meterText.textContent = pct + '%';
  if (meterSub) meterSub.textContent = `${totalUploaded} dari ${maxTarget} dokumen utama terupload ke Google Drive`;

  // 3. Tombol Akses Cepat Risalah Terakhir
  const latestMeeting = arsipList[0];
  if (quickBtnContainer && latestMeeting) {
    const d = parseTanggal(latestMeeting.tanggal);
    quickBtnContainer.innerHTML = `
      <button class="quick-btn" onclick="showArsipDetail(${latestMeeting.id})">
        📄 Baca Risalah Rapat Terakhir (${d.getDate()} ${SH_ID[d.getMonth()]}) ➔
      </button>
    `;
  }
}

function refreshStats() {
  const yr     = today.getFullYear();
  const total  = arsipList.length;
  const tiArsip= arsipList.filter(r => parseTanggal(r.tanggal).getFullYear() === yr);
  const ti     = tiArsip.length;
  const avg    = total ? Math.round(arsipList.reduce((a, r) => a + (r.peserta||[]).length, 0) / total) : 0;

  const ht = document.getElementById('hs-total');  if (ht) animCount(ht, total);
  const hy = document.getElementById('hs-tahun');  if (hy) animCount(hy, ti);
  const hn = document.getElementById('hs-nomor');  if (hn) hn.textContent = '#' + (settings.nomorLast + 1);

  const dt = document.getElementById('dash-total'); if (dt) animCount(dt, total);
  const dy = document.getElementById('dash-tahun'); if (dy) animCount(dy, ti);
  const da = document.getElementById('dash-avg');   if (da) animCount(da, avg);
  const dd = document.getElementById('dash-dok');   if (dd) animCount(dd, total * 3);
  const dl = document.getElementById('dash-tahun-lbl'); if (dl) dl.textContent = 'Rapat ' + yr;
  const cl = document.getElementById('chart-tahun-lbl'); if (cl) cl.textContent = yr;

  const months = Array(12).fill(0);
  tiArsip.forEach(r => months[parseTanggal(r.tanggal).getMonth()]++);
  const max = Math.max(...months, 1);
  const bc = document.getElementById('bar-chart-home');
  if (bc) bc.innerHTML = months.map((n, i) =>
    `<div class="bar-group">${n > 0 ? `<div class="bar-val">${n}</div>` : ''}<code>` +
    `<div class="bar" style="height:${Math.round(n/max*80)}px"><div class="bar-inner" style="height:100%"></div></div>` +
    `</code><div class="bar-label">${SH_ID[i]}</div></div>`
  ).join('');

  renderExecutiveWidgets();
}
const updateHeroStats = refreshStats;
const renderDashHome  = refreshStats;

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

// ════ HERO CANVAS BACKGROUND ══════════════════════════════════
(function () {
  const canvas = document.getElementById('hero-canvas'); if (!canvas) return;
  const ctx = canvas.getContext('2d'); let W, H, particles = [];
  function resize() { const h = canvas.parentElement; W = canvas.width = h.offsetWidth; H = canvas.height = h.offsetHeight; }
  function mk() { return { x:Math.random()*W, y:H+10, r:Math.random()*2.5+.5, speed:Math.random()*.6+.3, opacity:Math.random()*.6+.2, drift:(Math.random()-.5)*.4, life:0, maxLife:Math.random()*160+80 }; }
  function draw() {
    ctx.clearRect(0, 0, W, H);
    if (particles.length < 55 && Math.random() < .35) particles.push(mk());
    particles = particles.filter(p => {
      p.y -= p.speed; p.x += p.drift; p.life++;
      const t = p.life / p.maxLife, a = t < .2 ? t/.2 : t > .8 ? (1-t)/.2 : 1;
      ctx.beginPath(); ctx.arc(p.x, p.y, p.r, 0, Math.PI*2);
      ctx.fillStyle = `rgba(201,147,42,\${p.opacity*a})`; ctx.fill();
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
  const tx  = document.getElementById('hero-sync-text'); if (tx) tx.textContent = msg;
}
function setCloudBanner(state, msg) {
  const b  = document.getElementById('cloud-status-banner'); if (!b) return;
  const sp = document.getElementById('cloud-spin');
  const tx = document.getElementById('cloud-status-text');
  b.style.display = 'flex'; b.className = 'cloud-banner ' + state;
  if (sp) sp.style.display = (state === 'loading') ? 'block' : 'none';
  if (tx) tx.textContent = msg;
  if (state === 'ok' || state === 'err') setTimeout(() => { b.style.display = 'none'; }, 5000);
}

// ════ CLOUD STORAGE CONNECTIVITY ══════════════════════════════
async function loadArsipFromCloud() {
  if (!getGasUrl()) { setHeroSync('err','GAS URL belum diisi'); setCloudBanner('warn','URL GAS belum diisi — data hanya dari browser lokal.'); return; }
  setHeroSync('syncing','Menyinkron data cloud...'); showSync('Memuat dari cloud...','syncing');
  try {
    const data       = await gasCall('getArsip');
    const cloudArsip = sanitasiArsip(data.arsip || []);
    const cloudIds   = new Set(cloudArsip.map(r => String(r.id)));
    const localOnly  = arsipList.filter(r => !cloudIds.has(String(r.id)));
    arsipList = [...cloudArsip, ...localOnly];
    arsipList.sort((a, b) => String(b.id).localeCompare(String(a.id)));
    saveLocal();
    arsipList.forEach(r => {
      if (r.uploadedFiles?.length && !uploadFiles[r.id]?.length)
        uploadFiles[r.id] = r.uploadedFiles.map(f => ({...f, file:null, type:f.type||'', _showPreview:false}));
    });
    for (const item of localOnly) gasCall('simpanArsip', item).catch(() => {});
    setHeroSync('ok', `✓ \${arsipList.length} arsip tersinkron`);
    showSync(`\${arsipList.length} arsip tersinkron`, 'ok');
    setCloudBanner('ok', `✓ \${arsipList.length} arsip dimuat dari cloud`);
    renderArsip(); renderCalInline(); refreshStats();
  } catch (e) {
    setHeroSync('err','Gagal sync — menggunakan data lokal');
    setCloudBanner('err','❌ Gagal memuat dari cloud: ' + e.message);
    showSync('Gagal sync cloud','err');
  }
}
async function refreshArsipCloud() { await loadArsipFromCloud(); }

async function syncArsipToCloud(item) {
  if (!getGasUrl()) return;
  showSync('Menyimpan ke cloud...','syncing');
  try {
    await gasCall('simpanArsip', {
      ...item,
      uploadedFiles: (uploadFiles[item.id]||[]).map(f => ({name:f.name, size:f.size, status:f.status, url:f.url||null}))
    });
    showSync('Tersimpan ke cloud','ok');
  } catch { showSync('Gagal sync cloud','err'); }
}
async function hapusArsipCloud(id) {
  if (!getGasUrl()) return;
  try { await gasCall('hapusArsip', {id}); } catch {}
}

// ════ NOMOR SURAT & BOOKING PROTECTION ═════════════════════════
async function fetchNomor() {
  const dot  = document.getElementById('nomor-dot');
  const prev = document.getElementById('nomor-preview');
  if (!getGasUrl()) { dot.className = 'nomor-dot err'; prev.textContent = '— isi URL Apps Script di Pengaturan'; return; }
  dot.className = 'nomor-dot loading'; prev.textContent = 'Membaca dari Sheets...';
  try {
    const data = await gasCall('getLastNomor');
    settings.nomorLast = data.lastNomor;
    localStorage.setItem('sirapat_settings', JSON.stringify(settings));
    dot.className = 'nomor-dot ok';
    updateNomorPreview(); refreshStats();
  } catch { dot.className = 'nomor-dot err'; prev.textContent = '❌ Gagal — pakai nomor lokal: ' + (settings.nomorLast + 1); }
}
function updateNomorPreview() {
  const tgl  = document.getElementById('inp-tanggal').value;
  const d    = tgl ? parseTanggal(tgl) : new Date();
  const hint = tgl ? ` (urut ke-\${settings.nomorLast+1})` : ' (Silakan tentukan tanggal rapat)';
  document.getElementById('nomor-preview').textContent = buildNomor(settings.nomorLast + 1, d) + hint;
}

function checkBooking() {
  const tgl    = document.getElementById('inp-tanggal').value;
  const jam    = document.getElementById('inp-jam').value;
  const tempat = document.getElementById('inp-tempat').value.trim();
  const warn   = document.getElementById('booking-warn');
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
      `<span style="font-size:11px;opacity:.85">Sudah ada rapat pukul <strong>\${konflik.jam} WIB</strong> (<em>\${konflik.agenda.substring(0,60)}...</em>)</span><br>` +
      `<span style="font-size:11px;opacity:.7">Silakan ubah jam atau ganti ruangan.</span>`;
    if (btnGen) { btnGen.disabled = true; btnGen.style.opacity = '0.45'; btnGen.title = 'Ada konflik jadwal'; }
  } else {
    warn.style.display = 'none'; warn.innerHTML = '';
    if (btnGen) { btnGen.disabled = false; btnGen.style.opacity = ''; btnGen.title = ''; }
  }
}

// ════ KALENDER LOGIK MULTIPLE MEETING ACCENTS ══════════════════
function initCalInline()   { calYearInline = today.getFullYear(); calMonthInline = today.getMonth(); }
function changeMonthInline(d) {
  calMonthInline += d;
  if (calMonthInline < 0)  { calMonthInline = 11; calYearInline--; }
  if (calMonthInline > 11) { calMonthInline = 0;  calYearInline++; }
  renderCalInline();
}
function renderCalInline() {
  if (calYearInline === undefined) initCalInline();
  document.getElementById('cal-title-inline').textContent = `${BULAN_ID[calMonthInline]} ${calYearInline}`;

  const rapatMap = {};
  arsipList.forEach(r => { (rapatMap[r.tanggal] ??= []).push({jam:r.jam, tempat:r.tempat, agenda:r.agenda}); });

  const selTgl    = document.getElementById('inp-tanggal').value;
  const selJam    = document.getElementById('inp-jam').value;
  const selTempat = document.getElementById('inp-tempat').value.trim();
  const firstDay  = new Date(calYearInline, calMonthInline, 1).getDay();
  const days      = new Date(calYearInline, calMonthInline + 1, 0).getDate();
  const todayStr  = today.toISOString().split('T')[0];

  let html = HARI_ID.map((_, i) => `<div class="cal-day-name">${['Min','Sen','Sel','Rab','Kam','Jum','Sab'][i]}</div>`).join('');
  for (let i = 0; i < firstDay; i++) html += `<div class="cal-day other-month"></div>`;

  for (let d = 1; d <= days; d++) {
    const ds      = `${calYearInline}-${String(calMonthInline+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const events  = rapatMap[ds] || [];
    const hasEv   = events.length > 0;
    const isBooked = hasEv && ds === selTgl && events.some(e =>
      e.tempat === selTempat && Math.abs(getMenitDariJam(selJam) - getMenitDariJam(e.jam)) < 60
    );

    let cls = 'cal-day';
    if (ds === todayStr) cls += ' today'; else if (ds === selTgl) cls += ' selected';
    cls += isBooked ? ' booked' : hasEv ? ' has-event' : '';

    const tip = hasEv ? events.map(e => `${e.jam} – ${e.agenda.substring(0,25)}`).join(' | ') : '';

    let badgeHtml = '';
    if (hasEv) {
      const isKonflik = ds === selTgl && events.some(e =>
        e.tempat === selTempat && Math.abs(getMenitDariJam(selJam) - getMenitDariJam(e.jam)) < 60
      );
      badgeHtml = `<div class="cal-badge \${isKonflik ? 'konflik' : ''}">\${events.length} Rapat</div>`;
    }

    html += `<div class="\${cls}" \${tip ? `data-tip="\${tip.replace(/"/g,'&quot;')}"` : ''} onclick="calClickInline('\${ds}')">
      <span style="margin-bottom:2px">\${d}</span>
      \${badgeHtml}
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
  document.getElementById('modal-title').textContent = `Jadwal: \${d.getDate()} \${BULAN_ID[d.getMonth()]} \${d.getFullYear()}`;
  
  let html = '<div class="meeting-list-container">';
  events.forEach(r => {
    html += `
      <div class="meeting-card">
        <div class="mc-header">
          <div class="mc-time">🕒 \${r.jam} WIB</div>
          <div class="mc-participants">👥 \${(r.peserta||[]).length} Peserta</div>
        </div>
        <div class="mc-agenda">\${r.agenda}</div>
        <div class="mc-location">📍 \${r.tempat}</div>
        <button class="mc-action" onclick="showArsipDetail(\${r.id})">
          Lihat Berkas & Detail Rapat <span>➔</span>
        </button>
      </div>`;
  });
  html += '</div>';
  document.getElementById('modal-body').innerHTML = html;
  document.getElementById('modal-overlay').classList.add('open');
}

// ════ GENERATE DOKUMEN ════════════════════════════════════════
function setTplMode(mode, btn) {
  tplMode = mode; settings.tplMode = mode;
  document.querySelectorAll('.tpl-mode-tab').forEach(b => b.classList.remove('active')); btn.classList.add('active');
}
function getTemplateUrl(key) { return AUTO_PATHS[key]; }

async function generateDokumen() {
  const tanggalVal = document.getElementById('inp-tanggal').value;
  const jamVal     = document.getElementById('inp-jam').value;
  const tempat     = document.getElementById('inp-tempat').value.trim();
  const agenda     = document.getElementById('inp-agenda').value.trim();

  if (!tanggalVal) { showToast('Pilih tanggal rapat!','error'); return; }
  if (!agenda)     { showToast('Isi agenda rapat!','error'); return; }

  const menitBaru = getMenitDariJam(jamVal);
  const konflik = arsipList.find(r => {
    if (r.tanggal !== tanggalVal || r.tempat !== tempat) return false;
    return Math.abs(menitBaru - getMenitDariJam(r.jam)) < 60;
  });
  if (konflik) {
    showToast(`❌ Konflik jadwal! Jeda minimal 1 jam.`,'error');
    return;
  }

  const urlUnd = getTemplateUrl('und'), urlAbs = getTemplateUrl('abs'), urlRis = getTemplateUrl('ris');
  const pesertaHadir = getCheckedPeserta();
  if (!pesertaHadir.length) { showToast('Pilih minimal 1 peserta!','error'); return; }

  const tgl      = parseTanggal(tanggalVal);
  const hariStr = HARI_ID[tgl.getDay()];
  const tglStr  = tglFull(tgl);
  const tglGen  = tglGeneret();
  const jamFmt  = jamVal + ' WIB s/d Selesai';

  let nextNo = settings.nomorLast + 1;
  const data = {
    nomorSurat: buildNomor(nextNo, tgl), hari: hariStr, tanggal: tglStr,
    tanggalHari: `\${hariStr}, \${tglStr}`, jam: jamFmt, jamPolos: jamVal, tempat, agenda,
    ketua: settings.ketua, sekretaris: settings.sekretaris, kota: settings.kota,
    kotaTanggal: `\${settings.kota}, \${tglStr}`, tahun: String(tgl.getFullYear()),
    bulan: BULAN_ID[tgl.getMonth()], instansi: settings.instansi,
    jumlahPeserta: String(pesertaHadir.length), tgl_generet: tglGen,
    peserta: pesertaHadir.map((p, i) => ({no:String(i+1), nama:p.nama, jabatan:p.jabatan, ttd:''}))
  };

  const btn = document.getElementById('btn-gen');
  const sp  = document.getElementById('spinner');
  btn.disabled = true; sp.style.display = 'block';

  try {
    const blobs = await Promise.all([fetchAndInject(urlUnd,data), fetchAndInject(urlAbs,data), fetchAndInject(urlRis,data)]);
    const prefix = `Rapat_\${tanggalVal.replace(/-/g,'')}`;
    
    dlBlob(blobs[0], `\${prefix}_Undangan.docx`);
    dlBlob(blobs[1], `\${prefix}_AbsenHadir.docx`);
    dlBlob(blobs[2], `\${prefix}_Risalah.docx`);

    const arsipId = Date.now();
    const mkDraft = (blob, suffix) => ({
      file:blob, name:`Draft_\${prefix}_\${suffix}`, size:blob.size, status:'pending',
      url:null, type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document', _isDraft:false, _blob:blob
    });
    uploadFiles[arsipId] = [mkDraft(blobs[0],'Undangan.docx'), mkDraft(blobs[1],'AbsenHadir.docx'), mkDraft(blobs[2],'Risalah.docx')];

    const newItem = {id:arsipId, tanggal:tanggalVal, hari:hariStr, jam:jamVal, tempat, agenda, nomorSurat: data.nomorSurat, tglGeneret:tglGen, peserta: pesertaHadir.map(p => p.nama), uploadedFiles:[]};
    arsipList.unshift(newItem); saveLocal();

    settings.nomorLast = nextNo;
    localStorage.setItem('sirapat_settings', JSON.stringify(settings));

    if (getGasUrl()) {
      gasCall('simpanNomor', {nomorUrut:nextNo, nomorSurat:data.nomorSurat, tanggal:tglStr, agenda, tujuan:'', tglGeneret:tglGen, pesertaCount:pesertaHadir.length}).catch(()=>{});
      syncArsipToCloud(newItem);
    }

    lastGenId = arsipId;
    document.getElementById('btn-awan').classList.add('visible');
    updateNomorPreview(); renderCalInline(); refreshStats();
    showToast('✓ 3 Dokumen Terdownload!','success');
  } catch (err) { console.error(err); showToast('❌ Gagal','error'); }
  finally { btn.disabled = false; sp.style.display = 'none'; }
}

async function fetchAndInject(url, data) {
  const r = await fetch(url); if (!r.ok) throw new Error('HTTP ' + r.status);
  const zip = new PizZip(await r.arrayBuffer());
  const doc = new window.docxtemplater(zip, {paragraphLoop:true, linebreaks:true, delimiters:{start:'[[',end:']]'}, nullGetter:()=>''});
  doc.render(data);
  return doc.getZip().generate({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression:'DEFLATE'});
}

function simpanKeAwan() {
  if (!lastGenId) return;
  const btnBeranda = document.querySelectorAll('.nav-btn')[0];
  showPage('beranda', btnBeranda);
  setTimeout(() => {
    showArsipDetail(lastGenId);
    document.getElementById('modal-overlay').classList.add('open');
  }, 300);
}

// ════ ARSIP DATA RENDERER ═════════════════════════════════════
function renderArsip() {
  const q   = (document.getElementById('search-inp')?.value || '').toLowerCase();
  const bln = document.getElementById('filter-bulan')?.value || '';
  const thn = document.getElementById('filter-tahun')?.value || '';
  const list = arsipList.filter(r => {
    const d = parseTanggal(r.tanggal);
    if (bln && BULAN_ID[d.getMonth()] !== bln) return false;
    if (thn && String(d.getFullYear()) !== thn) return false;
    if (q && !JSON.stringify(r).toLowerCase().includes(q)) return false;
    return true;
  });
  const el = document.getElementById('arsip-list');
  if (!list.length) {
    el.innerHTML = `<div class="empty-state"><h3>Belum ada arsip</h3></div>`;
    return;
  }
  el.innerHTML = list.map(r => {
    const d = parseTanggal(r.tanggal);
    return `<div class="arsip-item" id="arsip-item-\${r.id}" onclick="showArsipDetail(\${r.id})">
      <div class="arsip-date-box"><div class="day">\${d.getDate()}</div><div class="month">\${SH_ID[d.getMonth()]}</div></div>
      <div class="arsip-info">
        <div class="arsip-title">\${r.agenda.substring(0,60)}...</div>
        <div class="arsip-meta">
          <span>📅 \${r.hari}, \${d.getFullYear()}</span>
          <span>🕐 \${r.jam} WIB</span>
          <span>👥 \${(r.peserta||[]).length}</span>
        </div>
      </div>
      <div class="arsip-actions" onclick="event.stopPropagation()"><button class="btn-sm" onclick="hapusArsip(\${r.id})">Hapus</button></div>
    </div>`;
  }).join('');
}

function hapusArsip(id) {
  if (!confirm('Hapus arsip ini?')) return;
  arsipList = arsipList.filter(r => r.id !== id);
  saveLocal(); delete uploadFiles[id]; hapusArsipCloud(id);
  renderArsip(); renderCalInline(); refreshStats();
}

// ════ MODAL FILE INTERACTIVE MANAGEMENT ═══════════════════════
function showArsipDetail(id) {
  currentModalId = id;
  const r = arsipList.find(x => x.id === id); if (!r) return;
  const d = parseTanggal(r.tanggal);
  const folderName = `\${String(d.getDate()).padStart(2,'0')} \${BULAN_ID[d.getMonth()]} \${d.getFullYear()}`;
  
  document.getElementById('modal-title').textContent = 'Detail & Berkas Rapat';
  document.getElementById('modal-body').innerHTML = `
    <div class="detail-row"><div class="detail-label">Tanggal</div><div class="detail-val">\Ref:\${r.hari}, \${tglFull(d)}</div></div>
    <div class="detail-row"><div class="detail-label">Pukul</div><div class="detail-val">\${r.jam} WIB</div></div>
    <div class="detail-row"><div class="detail-label">Tempat</div><div class="detail-val">\${r.tempat}</div></div>
    <div class="detail-row"><div class="detail-label">Agenda</div><div class="detail-val">\${r.agenda}</div></div>
    \${renderDraftSection(id)}
    <div class="upload-section">
      <div class="upload-section-title">☁ Cloud Server Google Drive <span class="folder-tag">📁 \${folderName}</span></div>
      <div class="upload-zone" onclick="document.getElementById('fi-\${id}').click()">
        <div class="upload-zone-text"><strong>Klik untuk Upload Berkas Final</strong></div>
        <input type="file" id="fi-\${id}" multiple onchange="handleFileInput(event,\${id})">
      </div>
      <div class="uploaded-files" id="file-list-\${id}"></div>
      <div class="upload-actions" id="upload-actions-\${id}" style="display:none; margin-top:8px;">
        <button class="btn-upload-all" id="upload-btn-\${id}" onclick="uploadSemuaFile(\${id},'\${folderName}')">☁ Upload ke Drive</button>
      </div>
    </div>`;
  renderFileList(id);
  document.getElementById('modal-overlay').classList.add('open');
}

function renderDraftSection(id) {
  const drafts = (uploadFiles[id]||[]).filter(f => f.file && f.status === 'pending');
  if (!drafts.length) return '';
  return `<div class="draft-section">
    <div class="draft-section-title">📝 Berkas Tergenerate (Draft)</div>
    <div class="draft-files">\${drafts.map((f, i) => `
      <div class="draft-file-item">
        <div class="draft-file-name">\${f.name}</div>
        <button class="draft-file-dl" onclick="downloadDraft(\${id},\${i})">⬇ Unduh</button>
      </div>`).join('')}
    </div>
  </div>`;
}
function downloadDraft(id, idx) { dlBlob(uploadFiles[id][idx].file, uploadFiles[id][idx].name); }

function renderFileList(id) {
  const files = (uploadFiles[id]||[]);
  const el = document.getElementById(`file-list-\${id}`); if (!el) return;
  const actEl = document.getElementById(`upload-actions-\${id}`);
  if (actEl) actEl.style.display = files.length ? 'flex' : 'none';
  if (!files.length) { el.innerHTML = ''; return; }

  el.innerHTML = files.map((f, realIdx) => {
    const isDone = f.status === 'done' && f.url;
    return `<div class="uploaded-file-item" style="padding: 6px 0; border-bottom: 1px solid #f0e8e0;">
      <div style="display:flex; justify-content:space-between; align-items:center;">
        <span style="font-size:12px;">\${getFileIcon(f.name)} \${f.name}</span>
        <div>
          \${isDone ? `<a href="\${f.url}" target="_blank" class="file-link">Buka Drive ↗</a>` : `<span class="file-status \${f.status}">\${statusLbl(f.status)}</span>`}
          <button onclick="hapusFile(\${id},\${realIdx})" style="margin-left:8px; border:none; background:none; cursor:pointer; color:red;">✕</button>
        </div>
      </div>
    </div>`;
  }).join('');
}

function handleFileInput(ev, id) {
  uploadFiles[id] ??= [];
  Array.from(ev.target.files).forEach(f => {
    uploadFiles[id].push({file:f, name:f.name, size:f.size, status:'pending', url:null, type:f.type||''});
  });
  renderFileList(id);
}
function hapusFile(id, idx) { uploadFiles[id].splice(idx,1); renderFileList(id); }

async function uploadSemuaFile(id, folderName) {
  const allFiles = uploadFiles[id] || [];
  const pending = allFiles.filter(f => f.status === 'pending');
  if (!pending.length) return;
  
  for (let i=0; i<allFiles.length; i++) {
    if (allFiles[i].status !== 'pending') continue;
    allFiles[i].status = 'uploading'; renderFileList(id);
    try {
      const res = await gasCall('uploadFile', {fileName:allFiles[i].name, fileBase64:await toBase64(allFiles[i].file), mimeType:allFiles[i].file.type, folderName});
      allFiles[i].status = 'done'; allFiles[i].url = res.fileUrl;
    } catch { allFiles[i].status = 'err'; }
    renderFileList(id);
  }
  const r = arsipList.find(x => x.id === id);
  if (r) { r.uploadedFiles = allFiles.map(f => ({name:f.name, size:f.size, status:f.status, url:f.url})); saveLocal(); }
  refreshStats();
}

function closeModal(e) { if (!e || e.target.id === 'modal-overlay' || e.target.className === 'modal-close') { document.getElementById('modal-overlay').classList.remove('open'); currentModalId = null; } }
function togglePeserta(i)   { document.getElementById('pgen-'+i).classList.toggle('checked'); }
function getCheckedPeserta(){ return pesertaList.filter((_, i) => document.getElementById('pgen-'+i)?.classList.contains('checked')); }
function pilihAgenda(el, text) { document.querySelectorAll('.agenda-chip').forEach(c => c.classList.remove('active')); el.classList.add('active'); const ta = document.getElementById('inp-agenda'); ta.value = text; }

// ════ ROUTING ROUTINES ════════════════════════════════════════
function showPage(id, btn) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn, .nav-drawer .nav-btn').forEach(b => b.classList.remove('active'));
  
  const targetPage = document.getElementById('page-' + id);
  if (targetPage) targetPage.classList.add('active');
  if (btn) btn.classList.add('active');
  
  if (id === 'peserta')    renderPesertaManage();
  if (id === 'beranda') {  renderCalInline(); renderArsip(); refreshStats(); }
  if (id === 'pengaturan') loadPengaturan();
}

function showToast(msg, type='info') { const t = document.getElementById('toast'); t.textContent = msg; t.className = `toast \${type} show`; setTimeout(() => t.classList.remove('show'), 4000); }
function toggleFaq(el) { el.classList.toggle('open'); el.nextElementSibling.classList.toggle('open'); }

// ════ MANAGEMENT PESERTA Ahli ═════════════════════════════════
function renderPesertaManage() {
  document.getElementById('peserta-manage-list').innerHTML = pesertaList.map((p, i) => `
    <div class="peserta-row">
      <div class="peserta-num">\${i+1}</div>
      <input type="text" value="\${p.nama}" id="pm-nama-\${i}">
      <input type="text" value="\${p.jabatan}" id="pm-jab-\${i}">
    </div>`
  ).join('');
}
function simpanPeserta() {
  pesertaList = pesertaList.map((_,i) => ({ nama: document.getElementById('pm-nama-'+i)?.value||'', jabatan: document.getElementById('pm-jab-'+i)?.value||'' })).filter(p => p.nama.trim());
  localStorage.setItem('sirapat_peserta', JSON.stringify(pesertaList));
  renderPesertaGen(); showToast('Peserta disimpan!','success');
}
function resetPeserta() { pesertaList = DEFAULT_PESERTA.map(p => ({...p})); renderPesertaManage(); }
function loadPengaturan() {}
function simpanPengaturan() { showToast('Pengaturan lokal disimpan!','success'); }

// ════ LOGIN CONTROL ═══════════════════════════════════════════
async function loginAdmin() {
  const inp = document.getElementById('admin-pin').value;
  if (!inp) return;
  try {
    const data = await gasCall('getPin');
    if (inp === String(data.pin)) {
      sessionStorage.setItem('documeet_auth','true');
      document.getElementById('login-screen').style.display='none';
      showToast('✓ Masuk ke DocuMeet','success');
      mulaiAutoSync();
    } else { alert('PIN Salah'); }
  } catch {
    // Fail-safe jika GAS offline
    sessionStorage.setItem('documeet_auth','true');
    document.getElementById('login-screen').style.display='none';
    mulaiAutoSync();
  }
}

function mulaiAutoSync() {
  loadArsipFromCloud().then(() => {
    renderCalInline(); refreshStats(); renderPesertaGen(); checkBooking();
  });
}

window.addEventListener('DOMContentLoaded', () => {
  if (sessionStorage.getItem('documeet_auth') === 'true') {
    document.getElementById('login-screen').style.display='none';
    mulaiAutoSync();
  }
});

// ════ INITIALIZATION ══════════════════════════════════════════
document.getElementById('inp-tanggal').value = today.toISOString().split('T')[0];
initCalInline(); renderCalInline(); renderPesertaGen(); refreshStats();
loadArsipFromCloud();
