/*  ====================================================================
    DASHBOARD SHE — PT PUTRA PERKASA ABADI, JOBSITE MARUWAI COAL
    Berkas server (Google Apps Script) — terikat pada Google Spreadsheet
    yang berisi sheet: KTA, TTA, INSPEKSI, OBSERVASI (+ TRAINING, KONTAK,
    LOG yang dibuat otomatis).

    Tugas berkas ini:
      • doGet()        → menyajikan halaman dashboard (Index.html)
      • bacaSemua()    → mengirim seluruh baris ke halaman
      • simpanBaris()  → menyimpan perubahan satu baris
      • tambahBaris()  → menambah baris baru
      • hapusBaris()   → menghapus baris
      • simpanBanyak() → menyimpan hasil impor sekaligus
      • simpanTraining() / hapusTraining() / simpanKontak()
    Setiap perubahan dicatat ke sheet LOG lengkap dengan email pengubah.
    ==================================================================== */

/* ---------- pengaturan ---------- */
const NAMA_SHEET = {
  kta:'KTA', tta:'TTA', ins:'INSPEKSI', obs:'OBSERVASI',
  training:'TRAINING', kontak:'KONTAK', log:'LOG'
};
/* Nama sheet kadang ditulis berbeda. Daftar padanan di bawah dipakai agar
   skrip tetap menemukan sheet yang benar tanpa perlu ubah kode. */
const PADANAN_SHEET = {
  'KTA'        : ['KTA','HAZARD KTA','KONDISI TIDAK AMAN','DATA KTA'],
  'TTA'        : ['TTA','HAZARD TTA','TINDAKAN TIDAK AMAN','DATA TTA'],
  'INSPEKSI'   : ['INSPEKSI','INSPECTION','INSPEKSI TERENCANA','TINS','DATA INSPEKSI'],
  'OBSERVASI'  : ['OBSERVASI','OBSERVATION','PTO','PLANNED TASK OBSERVATION','TOBS','DATA OBSERVASI']
};

/** Cari sheet berdasarkan nama, abaikan beda spasi/huruf besar-kecil,
 *  lalu coba daftar padanannya. */
function sheetCari_(nama) {
  var bk = bk_(), daftar = bk.getSheets();
  var rapi = function (s) { return String(s).replace(/\s+/g, ' ').trim().toUpperCase(); };
  var target = rapi(nama);
  for (var i = 0; i < daftar.length; i++) if (rapi(daftar[i].getName()) === target) return daftar[i];
  var alias = PADANAN_SHEET[target] || [];
  for (var a = 0; a < alias.length; a++) {
    for (var j = 0; j < daftar.length; j++) if (rapi(daftar[j].getName()) === rapi(alias[a])) return daftar[j];
  }
  return null;
}
const JUDUL_TRAINING = ['ID','NAMA','TANGGAL','TRAINER','PESERTA','DEPT','JENIS','STATUS','HADIR'];
const JUDUL_KONTAK   = ['NRP','NAMA','NO HP','DIPERBARUI'];
const JUDUL_LOG      = ['WAKTU','PENGGUNA','SHEET','NO','KOLOM','DARI','KE'];
const JUDUL_SERT     = ['WAKTU','NAMA','NRP','NRP ALAMTRI','JENIS TRAINING','TANGGAL','NILAI','HASIL','NOMOR','EMAIL','STATUS KIRIM','DIKIRIM OLEH'];

/* nama field yang dipakai halaman  →  kemungkinan judul kolom di sheet */
const FIELD_KE_JUDUL = {
  no:['NO','NO.'],
  tgl:['TGL. PELAPOR','TGL. INSPEKSI','TGL. OBSERVASI','TANGGAL'],
  deptPelapor:['DEPT. PELAPOR'], nrpPelapor:['NRP PELAPOR'], pelapor:['PELAPOR'],
  level:['LEVEL'], jabPelapor:['JABATAN PELAPOR'], posPelapor:['POSISI PELAPOR'],
  kategori:['KATEGORI'], lokasi:['LOKASI'], detail:['DETAIL LOKASI'], deskripsi:['DESKRIPSI'],
  deptPic:['DEPT. PIC'], nrpPic:['NRP PIC'], pic:['PIC'], jabPic:['JABATAN PIC'], posPic:['POSISI PIC'],
  status:['STATUS PELAPORAN','STATUS'], perbaikan:['PERBAIKAN'], iup:['IUP'],
  nrp:['NRP'], nama:['NAMA'], dept:['DEPT'], objek:['INSPEKSI'], catatan:['CATATAN'],
  week:['WEEK'], month:['MONTH']
};

/* ---------- menyajikan halaman ---------- */
function doGet() {
  siapkanSheetBantu_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dashboard SHE — PPA Jobsite Maruwai Coal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------- sumber data ----------
   Seluruh data dibaca dari satu Google Spreadsheet: "PERSONEL DATA MAC".
   Sheet yang dipakai: KTA, TTA, INSPEKSI, OBSERVASI, DATABASE KARYAWAN,
   LAGGING LEADING INDIKATOR, PERUSAHAAN JASA (+ TRAINING, KONTAK, LOG
   yang dibuat otomatis oleh skrip ini).
   Kosongkan ID_DATA bila skrip ingin memakai spreadsheet tempat ia terpasang. */
const ID_DATA = '1J22DsMDmoDu8WvgYKfzW0h2C7cso0n07MYhroCxwXqI';

function bk_() {
  return ID_DATA ? SpreadsheetApp.openById(ID_DATA) : SpreadsheetApp.getActiveSpreadsheet();
}

/** Uji cepat dari editor: daftar sheet, jumlah baris, dan judul kolomnya. */
function cekSheet() {
  var hasil = [];
  bk_().getSheets().forEach(function (sh) {
    var judul = sh.getLastColumn() ? sh.getRange(1, 1, Math.min(3, sh.getLastRow()), sh.getLastColumn()).getValues() : [];
    hasil.push({
      nama: sh.getName(), baris: sh.getLastRow(), kolom: sh.getLastColumn(),
      judulBaris1: (judul[0] || []).slice(0, 12).join(' | '),
      judulBaris2: (judul[1] || []).slice(0, 12).join(' | ')
    });
  });
  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

function sel_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]')
    return Utilities.formatDate(v, bk_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  if (typeof v === 'number') return (Math.round(v) === v) ? String(Math.round(v)) : String(v);
  return String(v).replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function indeksJudul_(judul, field) {
  const atas = judul.map(function (h) { return String(h).toUpperCase().trim(); });
  const kandidat = FIELD_KE_JUDUL[field] || [];
  for (var i = 0; i < kandidat.length; i++) {
    var p = atas.indexOf(kandidat[i]);
    if (p >= 0) return p;
  }
  return -1;
}

/* WEEKNUM(tanggal;2) — minggu dimulai Senin, sama seperti di berkas Excel */
function weeknum2_(ds) {
  if (!ds) return '';
  var b = String(ds).slice(0, 10).split('-');
  if (b.length < 3) return '';
  var t = new Date(+b[0], +b[1] - 1, +b[2]);
  var jan1 = new Date(+b[0], 0, 1);
  var geser = (jan1.getDay() + 6) % 7;
  return Math.floor((Math.round((t - jan1) / 86400000) + geser) / 7) + 1;
}
function bulan_(ds) { var b = String(ds).slice(0, 10).split('-'); return b.length > 1 ? +b[1] : ''; }

function siapkanSheetBantu_() {
  var bk = bk_();
  [[NAMA_SHEET.training, JUDUL_TRAINING], [NAMA_SHEET.kontak, JUDUL_KONTAK],
   [NAMA_SHEET.log, JUDUL_LOG], ['LOG SERTIFIKAT', JUDUL_SERT]]
    .forEach(function (p) {
      var sh = bk.getSheetByName(p[0]);
      if (!sh) {
        sh = bk.insertSheet(p[0]);
        sh.getRange(1, 1, 1, p[1].length).setValues([p[1]]).setFontWeight('bold');
        sh.setFrozenRows(1);
      }
    });
}

function catatLog_(namaSheet, no, jejak) {
  if (!jejak || !jejak.length) return;
  var sh = bk_().getSheetByName(NAMA_SHEET.log);
  if (!sh) return;
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  var waktu = new Date();
  var baris = jejak.map(function (j) { return [waktu, email, namaSheet, no, j.k, j.dari, j.ke]; });
  sh.getRange(sh.getLastRow() + 1, 1, baris.length, 7).setValues(baris);
}

/* ---------- membaca seluruh data ---------- */
function bacaSheetMentah_(nama) {
  var sh = sheetCari_(nama);
  if (!sh) return { kolom: [], baris: [] };
  var nilai = sh.getDataRange().getValues();
  if (!nilai.length) return { kolom: [], baris: [] };
  var kolom = nilai[0].map(function (v) { return String(v).trim(); });
  var baris = [];
  for (var i = 1; i < nilai.length; i++) {
    if (String(nilai[i][0]).trim() === '') continue;      // baris kosong / rumus kosong diabaikan
    baris.push(nilai[i].map(sel_));
  }
  return { kolom: kolom, baris: baris };
}

function bacaTraining_() {
  var sh = bk_().getSheetByName(NAMA_SHEET.training);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < n.length; i++) {
    if (String(n[i][1]).trim() === '') continue;
    out.push({
      _id: sel_(n[i][0]), nama: sel_(n[i][1]), tgl: sel_(n[i][2]), trainer: sel_(n[i][3]),
      peserta: Number(n[i][4]) || 0, dept: sel_(n[i][5]), jenis: sel_(n[i][6]),
      status: sel_(n[i][7]), hadir: Number(n[i][8]) || 0
    });
  }
  return out;
}
function bacaKontak_() {
  var sh = bk_().getSheetByName(NAMA_SHEET.kontak);
  var peta = {};
  if (!sh || sh.getLastRow() < 2) return peta;
  var n = sh.getDataRange().getValues();
  for (var i = 1; i < n.length; i++) {
    var nrp = sel_(n[i][0]), hp = sel_(n[i][2]);
    if (nrp && hp) peta[nrp] = hp;
  }
  return peta;
}
function bacaLog_(batas) {
  var sh = bk_().getSheetByName(NAMA_SHEET.log);
  var peta = {};
  if (!sh || sh.getLastRow() < 2) return peta;
  var total = sh.getLastRow() - 1;
  var ambil = Math.min(total, batas || 1500);
  var n = sh.getRange(sh.getLastRow() - ambil + 1, 1, ambil, 7).getValues();
  n.forEach(function (r) {
    var no = sel_(r[3]); if (!no) return;
    (peta[no] = peta[no] || []).push({
      tgl: (Object.prototype.toString.call(r[0]) === '[object Date]') ? r[0].toISOString() : String(r[0]),
      ubah: [{ k: sel_(r[4]), dari: sel_(r[5]), ke: sel_(r[6]) }],
      oleh: sel_(r[1])
    });
  });
  return peta;
}

/** Dipanggil halaman saat dimuat: seluruh baris + training + kontak + riwayat. */
function bacaSemua() {
  siapkanSheetBantu_();
  var hilang = ['KTA','TTA','INSPEKSI','OBSERVASI'].filter(function (n) { return !sheetCari_(n); });
  var k = bacaSheetMentah_(NAMA_SHEET.kta), t = bacaSheetMentah_(NAMA_SHEET.tta),
      i = bacaSheetMentah_(NAMA_SHEET.ins), o = bacaSheetMentah_(NAMA_SHEET.obs);
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  return {
    kolKTA: k.kolom, kta: k.baris,
    kolTTA: t.kolom, tta: t.baris,
    kolINS: i.kolom, inspeksi: i.baris,
    kolOBS: o.kolom, observasi: o.baris,
    training: bacaTraining_(), kontak: bacaKontak_(), riwayat: bacaLog_(1500),
    user: email, waktu: new Date().toISOString(),
    sheetHilang: hilang
  };
}

/* =====================================================================
   DATA SDM & INDIKATOR — dibaca langsung dari spreadsheet yang sama
   (DATABASE KARYAWAN · LAGGING LEADING INDIKATOR · PERUSAHAAN JASA)
   ===================================================================== */
function bacaSDM() {
  return { karyawan: bacaKaryawan_(), lagging: bacaLagging_(), mitra: bacaMitra_() };
}

/** cari baris judul pada sheet yang punya beberapa baris kepala */
function petaJudul_(nilai, batasBaris) {
  for (var r = 0; r < Math.min(batasBaris || 6, nilai.length); r++) {
    var baris = nilai[r].map(function (v) { return String(v).toUpperCase().trim(); });
    if (baris.indexOf('NRP') >= 0 && baris.indexOf('NAMA') >= 0) {
      var m = {};
      baris.forEach(function (h, i) { if (h && m[h] === undefined) m[h] = i; });
      return { baris: r, peta: m };
    }
  }
  return null;
}

function bacaKaryawan_() {
  var sh = bk_().getSheetByName('DATABASE KARYAWAN');
  if (!sh) return [];
  var n = sh.getDataRange().getValues();
  var j = petaJudul_(n, 6);
  if (!j) return [];
  var p = j.peta, keluar = [];
  var amb = function (r, nama) { return p[nama] === undefined ? '' : sel_(r[p[nama]]); };
  for (var i = j.baris + 1; i < n.length; i++) {
    var nrp = amb(n[i], 'NRP');
    if (!nrp || nrp.toUpperCase() === 'NRP' || amb(n[i], 'NAMA') === '') continue;
    var hp = amb(n[i], 'NO. HP').replace(/[^0-9+]/g, '');
    if (hp && hp.charAt(0) !== '0' && hp.indexOf('62') !== 0) hp = '0' + hp;
    keluar.push([nrp, amb(n[i], 'NAMA'), amb(n[i], 'DEPT').toUpperCase(), amb(n[i], 'JABATAN'),
      amb(n[i], 'POSISI'), amb(n[i], 'PERUSAHAAN').toUpperCase(), hp, amb(n[i], 'POH').toUpperCase(),
      amb(n[i], 'TGL. LAHIR').substring(0, 10), amb(n[i], 'PENDIDIKAN').toUpperCase(),
      amb(n[i], 'EMAIL').toLowerCase(), amb(n[i], 'GENDER').toUpperCase()]);
  }
  return keluar;
}

function angkaInd_(v) {
  if (v === '' || v === null || v === undefined || v === '-') return null;
  if (typeof v === 'number') return Math.round(v * 10000) / 10000;
  var s = String(v).replace('%', '').trim();
  if (/^-?[\d.]+,\d+$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  var f = parseFloat(s);
  return isNaN(f) ? String(v).trim() : f;
}

function bacaLagging_() {
  var sh = bk_().getSheetByName('LAGGING LEADING INDIKATOR');
  var keluar = {};
  if (!sh) return keluar;
  var n = sh.getDataRange().getValues();
  var barisJudul = -1, kolomBlok = [];
  for (var r = 0; r < Math.min(10, n.length); r++) {
    var isi = n[r].map(function (v) { return String(v).trim(); });
    if (isi.indexOf('Year to Date') >= 0) {
      barisJudul = r;
      isi.forEach(function (v, i) { if (v === 'No' && String(isi[i + 1]).trim() === 'Item') kolomBlok.push(i); });
      break;
    }
  }
  if (barisJudul < 0) return keluar;
  kolomBlok.forEach(function (c0) {
    var tahun = '';
    for (var r = 0; r < barisJudul; r++) {
      var m = String(n[r][c0] || '').match(/(20\d{2})/);
      if (m) tahun = m[1];
    }
    if (!tahun) return;
    var daftar = [], grup = 'LAGGING INDICATORS';
    for (var r2 = barisJudul + 1; r2 < n.length; r2++) {
      var no = String(n[r2][c0] || '').trim(), item = String(n[r2][c0 + 1] || '').trim();
      if (no && !item && no === no.toUpperCase() && no.length > 6) { grup = no; continue; }
      if (!item || no === 'No') continue;
      var bulan = [];
      for (var m2 = 0; m2 < 12; m2++) bulan.push(angkaInd_(n[r2][c0 + 3 + m2]));
      daftar.push({
        no: no, item: item.replace(/\s*\*[\s\S]*$/, '').replace(/\s+/g, ' ').trim(), grup: grup,
        thr: String(n[r2][c0 + 2] || '').replace(/\s+/g, ' ').trim(),
        bulan: bulan, ytd: angkaInd_(n[r2][c0 + 15]), ptd: angkaInd_(n[r2][c0 + 16])
      });
    }
    keluar[tahun] = daftar;
  });
  return keluar;
}

function bacaMitra_() {
  var sh = bk_().getSheetByName('PERUSAHAAN JASA');
  if (!sh || sh.getLastRow() < 2) return [];
  var n = sh.getDataRange().getValues();

  /* cari baris judul: baris yang memuat "Nama Perusahaan Jasa" */
  var barisJudul = -1;
  for (var r = 0; r < Math.min(8, n.length); r++) {
    for (var c = 0; c < n[r].length; c++) {
      if (String(n[r][c]).toUpperCase().indexOf('NAMA PERUSAHAAN') >= 0) { barisJudul = r; break; }
    }
    if (barisJudul >= 0) break;
  }
  if (barisJudul < 0) return [];

  /* gabungkan 3 baris kepala agar setiap kolom punya label */
  var label = [];
  for (var c2 = 0; c2 < n[barisJudul].length; c2++) {
    var teks = [];
    for (var r2 = barisJudul; r2 < Math.min(barisJudul + 3, n.length); r2++) {
      var v = String(n[r2][c2] || '').replace(/\s+/g, ' ').trim();
      if (v) teks.push(v.toUpperCase());
    }
    label[c2] = teks.join(' ');
  }
  var cari = function () {
    for (var i = 0; i < arguments.length; i++) {
      for (var c3 = 0; c3 < label.length; c3++) {
        if (label[c3].indexOf(arguments[i]) >= 0) return c3;
      }
    }
    return -1;
  };
  var kol = {
    nama: cari('NAMA PERUSAHAAN'), jenis: cari('JENIS USAHA'), izin: cari('NO. IZIN', 'NO IZIN'),
    bidang: cari('BIDANG USAHA'), kegiatan: cari('KEGIATAN'),
    dari: cari('DARI TGL'), sampai: cari('SAMPAI TGL'), bulan: cari('TOTAL BULAN'),
    lokal: cari('LOKAL'), nasional: cari('NASIONAL'),
    pjo: cari('PENANGGUNG JAWAB OPERASIONAL NAMA', 'PENANGGUNG JAWAB'),
    hp: cari('NO HP', 'NO. HP'), email: cari('EMAIL'), csms: cari('STATUS CSMS')
  };
  var amb = function (baris, k) { return kol[k] >= 0 ? sel_(baris[kol[k]]) : ''; };

  var keluar = [];
  for (var i2 = barisJudul + 1; i2 < n.length; i2++) {
    var nama = amb(n[i2], 'nama');
    if (!nama || nama.toUpperCase().indexOf('NAMA PERUSAHAAN') >= 0) continue;
    keluar.push({
      nama: nama, jenis: amb(n[i2], 'jenis'), izin: amb(n[i2], 'izin'),
      bidang: amb(n[i2], 'bidang'), kegiatan: amb(n[i2], 'kegiatan'),
      kontrakDari: amb(n[i2], 'dari').substring(0, 10), kontrakSampai: amb(n[i2], 'sampai').substring(0, 10),
      bulan: angkaInd_(amb(n[i2], 'bulan')), tkLokal: angkaInd_(amb(n[i2], 'lokal')),
      tkNasional: angkaInd_(amb(n[i2], 'nasional')), pjo: amb(n[i2], 'pjo'),
      hp: amb(n[i2], 'hp'), email: amb(n[i2], 'email'), csms: amb(n[i2], 'csms').toUpperCase()
    });
  }
  return keluar;
}


/* ---------- menulis ---------- */
function barisKe_(sh, no) {
  var kolomNo = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues();
  for (var i = 1; i < kolomNo.length; i++) {
    if (String(kolomNo[i][0]).trim().toUpperCase() === String(no).trim().toUpperCase()) return i + 1;
  }
  return -1;
}

/** Menyimpan perubahan satu baris. obj berisi nama field (bukan judul kolom). */
function simpanBaris(namaSheet, no, obj) {
  var kunci = LockService.getDocumentLock();
  kunci.waitLock(20000);
  try {
    var sh = sheetCari_(namaSheet);
    if (!sh) throw new Error('Sheet ' + namaSheet + ' tidak ada.');
    var judul = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var nomorBaris = barisKe_(sh, no);
    if (nomorBaris < 0) return tambahBaris(namaSheet, obj);      // belum ada → tambahkan
    var isi = sh.getRange(nomorBaris, 1, 1, judul.length).getValues()[0];
    var jejak = [];
    Object.keys(obj).forEach(function (f) {
      if (f.charAt(0) === '_') return;
      var idx = indeksJudul_(judul, f);
      if (idx < 0) return;
      var lama = sel_(isi[idx]), baru = sel_(obj[f]);
      if (lama !== baru) { jejak.push({ k: f, dari: lama, ke: baru }); isi[idx] = obj[f]; }
    });
    /* kolom turunan ikut diperbarui agar sheet tetap selaras */
    var iTgl = indeksJudul_(judul, 'tgl'), iWeek = indeksJudul_(judul, 'week'),
        iBulan = indeksJudul_(judul, 'month'), iIup = indeksJudul_(judul, 'iup'),
        iLok = indeksJudul_(judul, 'lokasi');
    if (iTgl >= 0 && iWeek >= 0) isi[iWeek] = weeknum2_(sel_(isi[iTgl]));
    if (iTgl >= 0 && iBulan >= 0) isi[iBulan] = bulan_(sel_(isi[iTgl]));
    if (namaSheet === NAMA_SHEET.ins && iIup >= 0 && iLok >= 0) isi[iIup] = sel_(isi[iLok]).substring(0, 2).toUpperCase();
    sh.getRange(nomorBaris, 1, 1, judul.length).setValues([isi]);
    catatLog_(namaSheet, no, jejak);
    return { ok: true, baris: nomorBaris, ubah: jejak.length };
  } finally { kunci.releaseLock(); }
}

/** Menambah satu baris baru di akhir sheet. */
function tambahBaris(namaSheet, obj) {
  var kunci = LockService.getDocumentLock();
  kunci.waitLock(20000);
  try {
    var sh = sheetCari_(namaSheet);
    if (!sh) throw new Error('Sheet ' + namaSheet + ' tidak ada.');
    var judul = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var baris = judul.map(function (h, idx) {
      var f = null;
      Object.keys(FIELD_KE_JUDUL).forEach(function (kunciField) {
        if (f) return;
        if (FIELD_KE_JUDUL[kunciField].indexOf(String(h).toUpperCase().trim()) >= 0) f = kunciField;
      });
      if (!f) return '';
      if (f === 'week') return weeknum2_(obj.tgl);
      if (f === 'month') return bulan_(obj.tgl);
      if (f === 'iup' && namaSheet === NAMA_SHEET.ins) return String(obj.lokasi || '').substring(0, 2).toUpperCase();
      return obj[f] === undefined ? '' : obj[f];
    });
    sh.appendRow(baris);
    catatLog_(namaSheet, obj.no, [{ k: 'baris', dari: '', ke: 'ditambahkan lewat dashboard' }]);
    return { ok: true, baris: sh.getLastRow() };
  } finally { kunci.releaseLock(); }
}

/** Menghapus satu baris berdasarkan nomor (kolom pertama). */
function hapusBaris(namaSheet, no) {
  var kunci = LockService.getDocumentLock();
  kunci.waitLock(20000);
  try {
    var sh = sheetCari_(namaSheet);
    var nomorBaris = barisKe_(sh, no);
    if (nomorBaris < 0) return { ok: false, pesan: 'Baris tidak ditemukan.' };
    sh.deleteRow(nomorBaris);
    catatLog_(namaSheet, no, [{ k: 'baris', dari: 'ada', ke: 'dihapus lewat dashboard' }]);
    return { ok: true };
  } finally { kunci.releaseLock(); }
}

/** Menyimpan banyak baris sekaligus (dipakai fitur Impor). */
function simpanBanyak(namaSheet, daftar) {
  var kunci = LockService.getDocumentLock();
  kunci.waitLock(60000);
  try {
    var sh = sheetCari_(namaSheet);
    if (!sh) throw new Error('Sheet ' + namaSheet + ' tidak ada.');
    var judul = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var semua = sh.getDataRange().getValues();
    var petaNo = {};
    for (var i = 1; i < semua.length; i++) petaNo[String(semua[i][0]).trim().toUpperCase()] = i;

    var fieldPerKolom = judul.map(function (h) {
      var f = null;
      Object.keys(FIELD_KE_JUDUL).forEach(function (kf) {
        if (f) return;
        if (FIELD_KE_JUDUL[kf].indexOf(String(h).toUpperCase().trim()) >= 0) f = kf;
      });
      return f;
    });
    var isiBaris = function (obj, asal) {
      return fieldPerKolom.map(function (f, idx) {
        if (!f) return asal ? asal[idx] : '';
        if (f === 'week') return weeknum2_(obj.tgl);
        if (f === 'month') return bulan_(obj.tgl);
        if (f === 'iup' && namaSheet === NAMA_SHEET.ins) return String(obj.lokasi || '').substring(0, 2).toUpperCase();
        if (obj[f] === undefined) return asal ? asal[idx] : '';
        return obj[f];
      });
    };
    var baru = [], diperbarui = 0;
    daftar.forEach(function (obj) {
      var k = String(obj.no || '').trim().toUpperCase();
      if (k && petaNo[k] !== undefined) {
        var idx = petaNo[k];
        semua[idx] = isiBaris(obj, semua[idx]);
        diperbarui++;
      } else {
        baru.push(isiBaris(obj, null));
      }
    });
    if (diperbarui) sh.getRange(1, 1, semua.length, judul.length).setValues(semua.map(function (r) {
      return r.slice(0, judul.length).concat(new Array(Math.max(0, judul.length - r.length)).fill(''));
    }));
    if (baru.length) sh.getRange(sh.getLastRow() + 1, 1, baru.length, judul.length).setValues(baru);
    catatLog_(namaSheet, 'IMPOR', [{ k: 'impor', dari: '', ke: baru.length + ' baris baru, ' + diperbarui + ' diperbarui' }]);
    return { ok: true, baru: baru.length, diperbarui: diperbarui };
  } finally { kunci.releaseLock(); }
}

/* ---------- training & kontak ---------- */
function simpanTraining(obj) {
  var kunci = LockService.getDocumentLock();
  kunci.waitLock(20000);
  try {
    var sh = bk_().getSheetByName(NAMA_SHEET.training);
    var n = sh.getDataRange().getValues();
    var baris = [obj._id, obj.nama, obj.tgl, obj.trainer, obj.peserta, obj.dept, obj.jenis, obj.status, obj.hadir];
    for (var i = 1; i < n.length; i++) {
      if (sel_(n[i][0]) === sel_(obj._id)) { sh.getRange(i + 1, 1, 1, baris.length).setValues([baris]); return { ok: true }; }
    }
    sh.appendRow(baris);
    return { ok: true };
  } finally { kunci.releaseLock(); }
}
function hapusTraining(id) {
  var sh = bk_().getSheetByName(NAMA_SHEET.training);
  var n = sh.getDataRange().getValues();
  for (var i = 1; i < n.length; i++) if (sel_(n[i][0]) === sel_(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  return { ok: false };
}
function simpanKontak(peta) {
  var kunci = LockService.getDocumentLock();
  kunci.waitLock(30000);
  try {
    var sh = bk_().getSheetByName(NAMA_SHEET.kontak);
    var n = sh.getDataRange().getValues();
    var petaBaris = {};
    for (var i = 1; i < n.length; i++) petaBaris[sel_(n[i][0])] = i + 1;
    var waktu = new Date(), baru = [];
    Object.keys(peta).forEach(function (nrp) {
      var isi = [nrp, '', peta[nrp], waktu];
      if (petaBaris[nrp]) {
        isi[1] = sel_(n[petaBaris[nrp] - 1][1]);
        sh.getRange(petaBaris[nrp], 1, 1, 4).setValues([isi]);
      } else baru.push(isi);
    });
    if (baru.length) sh.getRange(sh.getLastRow() + 1, 1, baru.length, 4).setValues(baru);
    return { ok: true };
  } finally { kunci.releaseLock(); }
}

/* ---------- menu bantu di dalam spreadsheet ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard SHE')
    .addItem('Siapkan sheet bantu (TRAINING, KONTAK, LOG)', 'siapkanSheetBantu_')
    .addItem('Buka dashboard', 'bukaDashboard_')
    .addToUi();
}
function bukaDashboard_() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('Alamat dashboard:\n\n' + (url || 'Belum di-deploy sebagai Web app.'));
}

/* =====================================================================
   PUSAT DOKUMEN — membaca folder Google Drive berisi SOP & Instruksi Kerja
   Setiap departemen punya folder sendiri (ID di bawah). Dokumen baru yang
   diunggah ke folder mana pun langsung muncul di website tanpa ubah kode.
   Untuk menambah departemen: tambahkan satu baris pada FOLDER_DEPT.
   ===================================================================== */
const FOLDER_INDUK = '1apGBjjxJqLa1fHh6_P1VTj-xHzvQmxJ-';   // folder SOP PPA MAC
const FOLDER_DEPT = {
  'COE'  : '1mbYEYDzDp_aJrtlXFdhkosfEQAkQm0-i',
  'ENG'  : '1eNHzrsBvmB_OuKLFddRGJLAdmygANZ0u',
  'FLO'  : '1D0lnFrxv-gymvyp8SPWCG813OuUFvdBo',   // folder FALOG
  'HCG'  : '1N5DTrZILsblDFHcfmGLtAKAiHBv-GhfA',   // folder HCGA
  'PLT'  : '1uqHMDUtBybqqIDOGhf8hodbthyF_1xcL',   // folder PLANT
  'PRO'  : '1dM6NAZ2QZm9E1vuy3FK73OnyaprmaUs-',   // folder PRO / PRODUKSI
  'SHE'  : '1e1ts4_sh5dkePOpvodlgCktyihm0Zl0L'
};

function bacaDokumen() {
  var keluar = [], galat = [];
  Object.keys(FOLDER_DEPT).forEach(function (dept) {
    try {
      jelajahFolder_(DriveApp.getFolderById(FOLDER_DEPT[dept]), [], keluar, 0, dept);
    } catch (e) {
      galat.push(dept + ': ' + e.message);
    }
  });
  keluar.sort(function (a, b) { return (a.dept + a.judul).localeCompare(b.dept + b.judul, 'id'); });
  if (galat.length) keluar.push({ judul: '', dept: '', jenis: '', folder: '', link: '', galat: galat.join(' · ') });
  return keluar;
}

function jelajahFolder_(folder, jalur, keluar, dalam, dept) {
  if (dalam > 6 || keluar.length > 2000) return;
  var tz = bk_().getSpreadsheetTimeZone();
  var berkas = folder.getFiles();
  while (berkas.hasNext()) {
    var f = berkas.next();
    var nama = f.getName();
    keluar.push({
      judul: nama.replace(/\.(pdf|docx?|xlsx?|pptx?)$/i, ''),
      nomor: (nama.match(/^[A-Z0-9][A-Z0-9\-\/\.]{3,}/) || [''])[0].replace(/[-\s]+$/, ''),
      jenis: tentukanJenis_(jalur, nama),
      dept: dept || tentukanDept_(jalur),
      folder: jalur.length ? jalur.join(' / ') : (dept || ''),
      link: f.getUrl(),
      unduh: 'https://drive.google.com/uc?export=download&id=' + f.getId(),
      mime: f.getMimeType(),
      ukuran: f.getSize(),
      diperbarui: Utilities.formatDate(f.getLastUpdated(), tz, 'yyyy-MM-dd')
    });
  }
  var sub = folder.getFolders();
  while (sub.hasNext()) {
    var s = sub.next();
    jelajahFolder_(s, jalur.concat([s.getName()]), keluar, dalam + 1, dept);
  }
}

/** Departemen cadangan bila folder tidak terdaftar di FOLDER_DEPT */
function tentukanDept_(jalur) {
  var kode = ['COE','SHE','PRO','PLT','ENG','FLO','HCG','MNG','PJO'];
  for (var i = 0; i < jalur.length; i++) {
    var s = String(jalur[i]).toUpperCase().trim();
    if (kode.indexOf(s) >= 0) return s;
  }
  return jalur.length > 1 ? jalur[1] : (jalur[0] || '-');
}

/** Jenis = SOP atau Instruksi Kerja, dibaca dari nama folder lalu nama berkas */
function tentukanJenis_(jalur, namaBerkas) {
  var teks = (jalur.join(' ') + ' ' + namaBerkas).toUpperCase();
  if (/INSTRUKSI KERJA|(^|[^A-Z])IK([^A-Z]|$)/.test(teks)) return 'Instruksi Kerja';
  if (/STANDAR OPERASIONAL|(^|[^A-Z])SOP([^A-Z]|$)/.test(teks)) return 'SOP';
  return 'Lainnya';
}

/** Uji cepat dari editor Apps Script: lihat berapa dokumen yang terbaca per departemen */
function cekFolderDokumen() {
  var d = bacaDokumen(), hitung = {};
  d.forEach(function (x) { if (x.judul) hitung[x.dept] = (hitung[x.dept] || 0) + 1; });
  Logger.log('Total dokumen terbaca: ' + d.filter(function (x) { return !!x.judul; }).length);
  Logger.log(hitung);
  var g = d.filter(function (x) { return x.galat; });
  if (g.length) Logger.log('Folder bermasalah: ' + g[0].galat);
  return hitung;
}


/* =====================================================================
   SERTIFIKAT TRAINING — kirim ke email peserta + catat riwayatnya
   Gambar sertifikat dibuat di halaman, lalu dikirim ke sini dalam bentuk
   base64 dan dilampirkan pada email.
   ===================================================================== */
function kirimSertifikat(data, base64png) {
  siapkanSheetBantu_();
  var pengirim = '';
  try { pengirim = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  var status = 'TERKIRIM', pesan = '';
  try {
    if (!data || !data.email) throw new Error('Alamat email peserta kosong.');
    var namaBerkas = 'Sertifikat - ' + (data.nama || 'Peserta') + '.png';
    var lampiran = Utilities.newBlob(Utilities.base64Decode(base64png), 'image/png', namaBerkas);
    var isi =
      '<div style="font-family:Arial,sans-serif;font-size:14px;color:#14312B;line-height:1.6">' +
      '<p>Yth. <b>' + (data.nama || '') + '</b>,</p>' +
      '<p>Selamat, Anda dinyatakan lulus mengikuti pelatihan <b>' + (data.jenis || '') + '</b> ' +
      'yang dilaksanakan oleh PT Putra Perkasa Abadi Jobsite Alamtri Mineral Indonesia &mdash; Maruwai Coal.</p>' +
      '<table style="font-size:14px;border-collapse:collapse">' +
      '<tr><td style="padding:2px 12px 2px 0;color:#5C7365">Nomor sertifikat</td><td><b>' + (data.nomor || '-') + '</b></td></tr>' +
      '<tr><td style="padding:2px 12px 2px 0;color:#5C7365">NRP Alamtri</td><td>' + (data.nrpAlamtri || '-') + '</td></tr>' +
      '<tr><td style="padding:2px 12px 2px 0;color:#5C7365">Tanggal pelaksanaan</td><td>' + (data.tanggal || '-') + '</td></tr>' +
      (data.nilai ? '<tr><td style="padding:2px 12px 2px 0;color:#5C7365">Nilai post test</td><td><b>' + data.nilai + '</b> &mdash; ' + (data.status || '') + '</td></tr>' : '') +
      '</table>' +
      '<p>Sertifikat terlampir pada email ini. Simpan sebagai bukti kompetensi Anda.</p>' +
      '<p style="color:#5C7365;font-size:12.5px">Email ini dikirim otomatis oleh SHE ONE &mdash; Database SHE PPA Maruwai Coal.</p></div>';
    MailApp.sendEmail({
      to: data.email,
      subject: 'Sertifikat Training — ' + (data.jenis || '') + ' — ' + (data.nama || ''),
      htmlBody: isi,
      attachments: [lampiran],
      name: 'SHE PPA Maruwai Coal'
    });
  } catch (e) {
    status = 'GAGAL'; pesan = e.message;
  }
  var sh = bk_().getSheetByName('LOG SERTIFIKAT');
  sh.appendRow([new Date(), data.nama || '', data.nrp || '', data.nrpAlamtri || '', data.jenis || '',
                data.tanggal || '', data.nilai || '', data.status || '', data.nomor || '', data.email || '',
                status === 'TERKIRIM' ? 'TERKIRIM' : ('GAGAL: ' + pesan), pengirim]);
  if (status !== 'TERKIRIM') throw new Error(pesan);
  return { ok: true, sisaKuota: MailApp.getRemainingDailyQuota() };
}

/** Catat penerbitan sertifikat tanpa mengirim email (mis. hanya diunduh). */
function catatSertifikat(data) {
  siapkanSheetBantu_();
  var pengirim = '';
  try { pengirim = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  bk_().getSheetByName('LOG SERTIFIKAT').appendRow([
    new Date(), data.nama || '', data.nrp || '', data.nrpAlamtri || '', data.jenis || '',
    data.tanggal || '', data.nilai || '', data.hasil || data.status || '', data.nomor || '',
    data.email || '', data.status || 'DIUNDUH', pengirim]);
  return { ok: true };
}

function riwayatSertifikat() {
  siapkanSheetBantu_();
  var sh = bk_().getSheetByName('LOG SERTIFIKAT');
  if (!sh || sh.getLastRow() < 2) return [];
  var n = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues(), keluar = [];
  n.forEach(function (r) {
    if (!r[1]) return;
    keluar.push({
      waktu: (Object.prototype.toString.call(r[0]) === '[object Date]')
        ? Utilities.formatDate(r[0], bk_().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm') : sel_(r[0]),
      nama: sel_(r[1]), nrp: sel_(r[2]), nrpAlamtri: sel_(r[3]), jenis: sel_(r[4]),
      tanggal: sel_(r[5]).substring(0, 10), nilai: sel_(r[6]), hasil: sel_(r[7]),
      nomor: sel_(r[8]), email: sel_(r[9]), status: sel_(r[10]), oleh: sel_(r[11])
    });
  });
  return keluar.slice(-300);
}

/** Uji kirim dari editor: ganti alamat di bawah lalu Run. */
function ujiKirimSertifikat() {
  Logger.log('Sisa kuota email hari ini: ' + MailApp.getRemainingDailyQuota());
}


/* =====================================================================
   DIAGNOSA — pastikan sheet KTA, TTA, INSPEKSI, dan OBSERVASI terbaca
   Jalankan dari editor Apps Script, lalu buka Execution log.
   ===================================================================== */
function cekDataSAP() {
  var hasil = [];
  ['KTA', 'TTA', 'INSPEKSI', 'OBSERVASI'].forEach(function (nama) {
    var sh = sheetCari_(nama);
    if (!sh) { hasil.push({ diminta: nama, status: 'SHEET TIDAK DITEMUKAN' }); return; }
    var judul = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function (v) { return String(v).trim(); }).filter(String);
    var terpetakan = judul.filter(function (h) {
      var k = h.toUpperCase().replace(/\s+/g, ' ');
      return !!FIELD_KE_JUDUL[Object.keys(FIELD_KE_JUDUL).filter(function (f) {
        return FIELD_KE_JUDUL[f].indexOf(k) >= 0;
      })[0]];
    });
    var isi = 0;
    if (sh.getLastRow() > 1) {
      isi = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
        .filter(function (r) { return String(r[0]).trim() !== ''; }).length;
    }
    hasil.push({
      diminta: nama, ditemukan: sh.getName(), barisBerisi: isi, jumlahKolom: judul.length,
      kolomDikenali: terpetakan.length, judulKolom: judul.join(' | '),
      kolomTidakDikenali: judul.filter(function (h) { return terpetakan.indexOf(h) < 0; }).join(' | ')
    });
  });
  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}
