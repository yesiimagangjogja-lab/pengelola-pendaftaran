/**
 * Dashboard Admin Pendaftaran Magang Multi-Brand
 *
 * Satu brand = satu Google Form. Jawaban form masuk ke Google Sheets,
 * dan dashboard ini membaca sheet itu apa adanya (judul kolom tidak diubah).
 * Kolom admin (Status, Follow-up Terakhir, Link Sertifikat) ditambahkan di akhir sheet.
 */

// Dipakai hanya oleh buatDataContoh(), sebelum data asli disambungkan.
const CONTOH_BRAND = ['MagangJogja', 'FreelanceJogja', 'SkillUp Jogja'];

const ADMIN_COLS = ['Status', 'Follow-up Terakhir', 'Link Sertifikat', 'Jadwal Wawancara', 'Riwayat Status', 'Dokumen Diperiksa', 'Email Terakhir'];
const LOG_COLS = ['Waktu', 'Brand', 'Email', 'Nama', 'Catatan'];
// 9 Nov: admin minta tahap "Ongoing" diganti "Diterima" — yang lolos seleksi berstatus Diterima,
// tidak ada lagi nama Ongoing di dashboard. Yang masih tertulis "Ongoing"/"Berjalan" di sheet lama
// dipetakan ke Diterima lewat STATUS_LAMA, tanpa mengubah isi spreadsheetnya.
const STATUS_LIST = ['Seleksi', 'Diterima', 'Lulus', 'Ditolak', 'Dikeluarkan'];
// Status yang masih proses: orangnya belum bisa dapat sertifikat.
const STATUS_PROSES = ['Seleksi', 'Diterima'];
// Status yang wawancaranya masih akan datang — hanya ini yang dihitung di kartu "Sudah dijadwalkan".
const STATUS_WAWANCARA = ['Seleksi'];
// Nama lama yang mungkin masih ada di sheet, dipetakan ke nama baru.
const STATUS_LAMA = { selesai: 'Lulus', 'tidak jadi': 'Ditolak', do: 'Dikeluarkan', 'drop out': 'Dikeluarkan', berjalan: 'Diterima', ongoing: 'Diterima' };
const JUMLAH_MINGGU_TREND = 8;
const VERSI_APLIKASI = '3.104';
const PROP_BRANDS = 'KONFIG_BRAND';
const PROP_LOG = 'LOG_SPREADSHEET_ID';
const NAMA_LOG = 'Log Follow-up Pendaftaran Magang';
const LOG_SHEET = 'FollowUp';
const FOLDER_SERTIFIKAT = 'Sertifikat Pendaftaran Magang';
const FOLDER_EKSPOR = 'Ekspor Dashboard Magang';

// Kolom admin ditulis singkat di dashboard, panjang di spreadsheet supaya tidak menimpa jawaban form.
const NAMA_KOLOM_ADMIN = {
  status: 'Status',
  followUp: 'Follow-up Terakhir',
  sertifikat: 'Link Sertifikat',
  jadwal: 'Jadwal Wawancara',
  riwayat: 'Riwayat Status',
  dokumen: 'Dokumen Diperiksa',
};

// Pencocok judul kolom form. regex = uji sebagian, string = sama persis.
const ALIAS = {
  timestamp: 'Timestamp',
  nama: /nama lengkap/i,
  tahunLahir: /tahun lahir/i,
  nim: /\bnim\b/i,
  email: /^email$/i,
  hp: /no\. hp aktif yang bisa dihubungi/i,
  kampus: /asal sekolah|asal kampus|asal perguruan tinggi|nama (kampus|sekolah|perguruan tinggi)|universitas|sekolah tinggi/i,
  prodi: /program studi/i,
  fakultas: /fakultas/i,
  kota: /nama kota/i,
  alasan: /mengapa anda ingin/i,
  jenisMagang: /jenis magang/i,
  sistem: /siap dg sistem/i,
  statusDiri: /status anda saat ini/i,
  inggris: /membaca buku berbahasa inggris/i,
  pembimbing: /dosen atau guru pembimbing/i,
  minat: /yang anda minati/i,
  laptop: /alat kerja sendiri/i,
  alat: /alat apa yang anda miliki/i,
  mulai: /berencana mulai/i,
  sumber: /tahu info magang/i,
  cv: /silakan upload\s*:\s*cv/i,
  portofolio: /silakan upload\s*:\s*portofolio/i,
  kegiatan: /kegiatan anda saat ini/i,
  kost: /kost\/kontrakakan/i,
  keluarga: /sudah berkeluarga/i,
  ortu: /wali\/ ortu/i,
  ig: /sosial media/i,
};

// Ditampilkan sebagai ringkasan di tabel dashboard.
const RINGKAS = ['nama', 'email', 'hp', 'nim', 'kampus', 'prodi', 'minat', 'mulai'];

// Label pendek untuk panel detail.
const LABEL_RINGKAS = {
  hp: 'No. HP / WA',
  nim: 'NIM',
  kampus: 'Asal kampus',
  prodi: 'Program studi',
  minat: 'Minat magang',
  mulai: 'Rencana mulai magang',
};


/**
 * Dijalankan kotak error: jawab "kenapa halaman belum bisa dibuka" dengan bukti, bukan tebakan.
 * Sengaja tidak butuh login dan tidak boleh melempar error sama sekali — termasuk kalau
 * penyimpanan internal aplikasi pun sedang ditolak — supaya layarnya selalu ada isi.
 */
function kenapaBelumBisa() {
  const coba = function (label, jalan) {
    let hasil = '';
    try {
      hasil = String(jalan());
    } catch (e) {
      hasil = 'DITOLAK: ' + pesanIzin_(e);
    }
    return { label: label, isi: hasil };
  };
  const baris = [
    coba('akun yang menjalankan', function () {
      return Session.getActiveUser().getEmail() + ' (efektif: ' + Session.getEffectiveUser().getEmail() + ')';
    }),
    coba('penyimpanan konfigurasi aplikasi', function () {
      return PropertiesService.getScriptProperties().getKeys().length + ' pengaturan tersimpan';
    }),
    coba('alamat file log', function () {
      return PropertiesService.getScriptProperties().getProperty(PROP_LOG) || '(belum pernah diisi)';
    }),
    coba('brand yang tersimpan', function () {
      return getBrands_().map(function (b) { return b.name; }).join(', ') || '(kosong)';
    }),
    coba('spreadsheet data kantor', function () {
      return SpreadsheetApp.openById(SS_DATA_LAMA).getName();
    }),
  ];

  const idLog = (function () {
    try {
      return teks_(PropertiesService.getScriptProperties().getProperty(PROP_LOG));
    } catch (e) { return ''; }
  })();
  const file = [];
  if (idLog) {
    file.push({ label: 'file log catatan & konfigurasi', id: idLog, nama: '', bisa: false, alasan: '' });
    try {
      const ss = SpreadsheetApp.openById(idLog);
      file[0].bisa = true;
      file[0].nama = ss.getName();
    } catch (e) {
      file[0].alasan = pesanIzin_(e);
    }
  }

  return {
    ok: true,
    akun: (baris[0].isi.indexOf('DITOLAK') === 0 ? '(tidak terbaca)' : baris[0].isi),
    brand: 0,
    versi: VERSI_APLIKASI,
    baris: baris,
    file: file,
  };
}

/**
 * Dijalankan manual dari editor (dropdown fungsi -> cekIzin -> Jalankan). Mengumpulkan hasil
 * per layanan ke SATU baris, lalu dibuang sebagai error supaya kelihatan merah dan bisa disalin.
 * Sengaja tidak pakai Logger/Console: keduanya juga butuh izin, jadi pesannya bisa hilang.
 * Tidak menulis apa pun ke file mana pun.
 */
function cekIzin() {
  const coba = function (label, jalan) {
    let hasil = '';
    try {
      hasil = String(jalan());
    } catch (e) {
      hasil = 'GAGAL(' + (e.message || e) + ')';
    }
    return label + '=' + hasil;
  };
  const laporan = [
    'versi ' + VERSI_APLIKASI,
    coba('akun', function () {
      return 'aktif:' + Session.getActiveUser().getEmail() + ' efektif:' + Session.getEffectiveUser().getEmail();
    }),
    coba('properties', function () {
      return PropertiesService.getScriptProperties().getKeys().length + ' kunci';
    }),
    coba('log_id', function () {
      return PropertiesService.getScriptProperties().getProperty(PROP_LOG) || '(kosong)';
    }),
    coba('brand', function () {
      return getBrands_().map(function (b) { return b.name; }).join('+') || '(kosong)';
    }),
    coba('sheet_kantor', function () {
      return SpreadsheetApp.openById(SS_DATA_LAMA).getName();
    }),
    coba('drive_cari', function () {
      return DriveApp.getFilesByName(NAMA_LOG).hasNext() ? 'ketemu' : 'tidak ketemu';
    }),
    coba('trigger', function () {
      return ScriptApp.getProjectTriggers().length + ' trigger';
    }),
    // Satu baris per brand: file + tab yang benar-benar dibaca dashboard, jadi brand yang
    // menunjuk ke file orang lain langsung kelihatan, bukan bikin halaman mati diam-diam.
  ].concat(getBrands_().slice(0, 10).map(function (b) {
    return coba('brand[' + b.name + ']', function () {
      const ss = SpreadsheetApp.openById(b.spreadsheetId);
      const ada = ss.getSheetByName(b.sheetName);
      return ss.getName() + ' / ' + b.sheetName + (ada ? ' (tab ada, ' + Math.max(ada.getLastRow() - 1, 0) + ' baris)' : ' (TAB TIDAK ADA)');
    });
  }));
  throw new Error('HASIL CEK IZIN >> ' + laporan.join(' | '));
}

/**
 * Menanya hak akses ke Drive TANPA menulis apa pun: siapa pemilik file ini, dan boleh edit tidak
 * oleh akun yang menjalankan aplikasi. Ini yang membedakan "file-nya belum dibagikan" dari
 * "aplikasinya yang jalan sebagai akun lain".
 */
function cekAksesTulis(urlAtauId) {
  butuhSesi_();
  const id = ambilIdSpreadsheet_(urlAtauId);
  if (!/^[A-Za-z0-9_-]{10,}$/.test(id)) throw new Error('Link atau ID spreadsheet-nya belum dikenali.');
  const hasil = { id: id, nama: '', pemilik: '', akses: '', editor: null, catatan: [] };
  try {
    const file = DriveApp.getFileById(id);
    hasil.nama = file.getName();
    try { hasil.pemilik = file.getOwner().getEmail(); } catch (e) { hasil.pemilik = '(tidak kelihatan)'; }
    try { hasil.akses = String(file.getAccess()).replace('Access.', ''); } catch (e) { hasil.catatan.push('akses: ' + (e.message || e)); }
    try { hasil.editor = file.isEditor(); } catch (e) { hasil.catatan.push('isEditor: ' + (e.message || e)); }
  } catch (e) {
    throw new Error('File itu tidak bisa dibuka sama sekali: ' + pesanIzin_(e));
  }
  try {
    const ss = SpreadsheetApp.openById(id);
    hasil.tab = ss.getSheets().map(function (s) { return s.getName(); }).slice(0, 12);
  } catch (e) {
    hasil.catatan.push('Sheets: ' + pesanIzin_(e));
  }
  hasil.yangJalan = akunAplikasi_() || '(anonim)';
  return { ok: true, hasil: hasil };
}

function doGet(e) {
  const tmpl = HtmlService.createTemplateFromFile('Index');

  tmpl.brandAwal = JSON.stringify(
    String((e && e.parameter && e.parameter.brand) || '')
  );

  return tmpl.evaluate()
    .setTitle('Dashboard Admin Pendaftaran Magang')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Dibuat otomatis oleh tombol "Buat data contoh" di dashboard,
 * atau dijalankan manual dari editor. Membuat spreadsheet berisi 3 brand contoh
 * dengan judul kolom persis seperti form pendaftaran, lalu mendaftarkannya ke dashboard.
 */
function buatDataContoh() {
  butuhSesi_();
  const ss = SpreadsheetApp.create('Data Pendaftaran Magang (Contoh)');

  const header = CONTOH_HEADER.concat(ADMIN_COLS);
  const config = CONTOH_BRAND.map(function (namaBrand, i) {
    // Sheet pertama memakai bawaan spreadsheet: Google tidak mengizinkan semua sheet terhapus.
    const sheet = i === 0 ? ss.getSheets()[0].setName('Form ' + (i + 1)) : ss.insertSheet('Form ' + (i + 1));
    sheet.appendRow(header);
    formatHeader_(sheet);
    CONTOH_ORANG.filter(function (o) { return o.brand === i; })
      .sort(function (a, b) { return new Date(a.timestamp) - new Date(b.timestamp); })
      .forEach(function (o) {
        sheet.appendRow(barisContoh_(header, o));
      });
    return { name: namaBrand, spreadsheetId: ss.getId(), sheetName: sheet.getName() };
  });

  const log = getLogSheet_();
  CONTOH_FOLLOWUP.forEach(function (f) {
    const orang = CONTOH_ORANG.filter(function (o) { return o.email === f.email; })[0];
    log.appendRow([new Date(f.ts), f.brand, f.email, orang ? orang.nama : '', f.catatan]);
  });

  setBrands_(config);
  return { url: ss.getUrl(), brands: config.map(function (c) { return c.name; }) };
}

/**
 * Sambungkan satu brand ke spreadsheet hasil Google Form asli.
 * Contoh: sambungkanBrand('MagangJogja', '1AbCdef...', 'Form Responses 1')
 */
function sambungkanBrand(namaBrand, spreadsheetId, sheetName) {
  butuhSesi_();
  const ss = SpreadsheetApp.openById(String(spreadsheetId).trim());
  const sheet = (sheetName ? ss.getSheetByName(String(sheetName)) : null) || cariSheetForm_(ss) || ss.getSheets()[0];
  if (!sheet) throw new Error('Sheet "' + sheetName + '" tidak ada di spreadsheet itu.');

  // Brand yang sudah ada diperbaiki di tempat, bukan dipindah ke belakang:
  // menekan Simpan pada konfigurasi tidak boleh mengubah urutan bilah kiri.
  const list = getBrands_();
  const punya = list.filter(function (b) { return b.name === namaBrand; })[0];
  if (punya) {
    punya.spreadsheetId = ss.getId();
    punya.sheetName = sheet.getName();
  } else {
    list.push({ name: namaBrand, spreadsheetId: ss.getId(), sheetName: sheet.getName() });
  }
  setBrands_(list);
  return { ok: true, brand: namaBrand, spreadsheet: ss.getName(), sheet: sheet.getName() };
}

function ambilIdSpreadsheet_(idAtauUrl) {
  const s = String(idAtauUrl || '').trim();
  const cocok = s.match(/\/d\/([^\/?#]+)/);
  return cocok ? cocok[1] : s;
}

/**
 * Semua brand dipakai dari SATU spreadsheet yang sudah ada (1 brand = 1 tab di file yang sama).
 * Tidak ada file baru yang dibuat; tab lama dipakai lagi kalau isinya sudah ada.
 */
function pakaiSatuSpreadsheet(idAtauUrl) {
  butuhSesi_();
  const id = ambilIdSpreadsheet_(idAtauUrl);
  if (!id) throw new Error('ID atau link spreadsheet belum diisi.');
  const ss = SpreadsheetApp.openById(id);
  const sheets = ss.getSheets();
  const lama = getBrands_();
  const daftarBaru = [];
  const dipakai = {};

  CONTOH_BRAND.forEach(function (brand, i) {
    const bLama = lama.filter(function (b) { return b.name === brand; })[0] || {};
    // getSheetByName(undefined) ditolak Google, jadi nama lama cuma dipakai kalau benar-benar ada.
    let sheet = ss.getSheetByName(brand) || (bLama.sheetName ? ss.getSheetByName(bLama.sheetName) : null);
    if (!sheet) sheet = ss.getSheetByName('Form ' + (i + 1));
    if (!sheet) sheet = sheets.filter(function (s) { return !dipakai[s.getName()]; })[0];
    if (!sheet) sheet = ss.insertSheet(brand);
    dipakai[sheet.getName()] = true;
    batalCache_(brand);
    daftarBaru.push({
      name: brand,
      spreadsheetId: ss.getId(),
      sheetName: sheet.getName(),
      formId: bLama.formId || '',
      formUrl: bLama.formUrl || '',
    });
  });

  setBrands_(daftarBaru);
  // Sekalian pastikan tiap tab punya header yang benar.
  daftarBaru.forEach(function (b) { getBrandSheet_(b.name, true); });
  batalCache_();
  return { spreadsheet: ss.getName(), url: ss.getUrl(), brands: daftarBaru };
}

/**
 * Memindahkan semua file proyek (spreadsheet brand, log, form, folder sertifikat & ekspor)
 * ke satu folder Drive, supaya tidak berserakan di root.
 */
function pindahFileKeFolder(namaFolder) {
  butuhSesi_();
  const nama = String(namaFolder || '').trim() || 'pengelola pendaftaran magang';
  const ketemu = DriveApp.getFoldersByName(nama);
  const folder = ketemu.hasNext() ? ketemu.next() : DriveApp.createFolder(nama);

  const ids = {};
  getBrands_().forEach(function (b) {
    if (b.spreadsheetId) ids[b.spreadsheetId] = true;
    if (b.formId) ids[b.formId] = true;
  });
  const idLog = PropertiesService.getScriptProperties().getProperty(PROP_LOG);
  if (idLog) ids[idLog] = true;

  const dipindah = [];
  const gagal = [];
  Object.keys(ids).forEach(function (id) {
    try {
      const f = DriveApp.getFileById(id);
      f.moveTo(folder);
      dipindah.push(f.getName());
    } catch (e) {
      gagal.push(id + ': ' + (e.message || e));
    }
  });
  [FOLDER_SERTIFIKAT, FOLDER_EKSPOR].forEach(function (n) {
    const it = DriveApp.getFoldersByName(n);
    if (!it.hasNext()) return;
    try {
      it.next().moveTo(folder);
      dipindah.push(n + '/');
    } catch (e) {
      gagal.push(n + '/: ' + (e.message || e));
    }
  });
  return { folder: folder.getName(), url: folder.getUrl(), dipindah: dipindah, gagal: gagal };
}

// Daftar pertanyaan form, disusun sama seperti form pendaftaran magang yang sudah ada.
// Nama form dikasih tanda (CONTOH) supaya tidak tertukar dengan form resmi kantor.
function namaFormBrand_(brand, singkat) {
  return 'Form Pendaftaran Magang - ' + brand + (singkat ? ' - Singkat' : ' - Lengkap') + ' (CONTOH)';
}

const PERTANYAAN_FORM = [
  { t: 'text', judul: 'Nama Lengkap', wajib: true },
  { t: 'text', judul: 'Tahun Lahir', wajib: true },
  { t: 'text', judul: 'NIM (Nomor Induk Mahasiswa)', wajib: true },
  { t: 'email', judul: 'Email', wajib: true },
  { t: 'choice', judul: 'Jenis Kelamin', opsi: ['Laki-laki', 'Perempuan'], wajib: true },
  { t: 'text', judul: 'No. HP Aktif yang bisa dihubungi dan chat WA', wajib: true },
  { t: 'text', judul: 'Asal Sekolah/Kampus', wajib: true },
  { t: 'text', judul: 'Program Studi', wajib: true },
  { t: 'text', judul: 'Fakultas', wajib: true },
  { t: 'text', judul: 'Nama Kota/Daerah tempat tinggal Anda saat ini', wajib: true },
  { t: 'paragraf', judul: 'Mengapa Anda ingin Magang/PKL disini?', wajib: true },
  {
    t: 'choice', judul: 'Jenis Magang apa yang Anda pilih ?', wajib: true,
    opsi: ['Magang Mandiri (magang atas inisatif diri sendiri, bukan kewajiban dari kampus)', 'Magang Kampus / Reguler (magang karena kewajiban dari kampus)'],
  },
  { t: 'choice', judul: 'Saya Siap dg Sistem Magang ini :', opsi: ['WFO (Work From Office)'], wajib: true },
  {
    t: 'choice', judul: 'Status Anda saat ini :', wajib: true,
    opsi: ['Masih sekolah/masih kuliah', 'Sudah lulus sekolah/kuliah dan belum bekerja', 'Sudah lulus dan sedang bekerja'],
  },
  {
    t: 'choice', judul: 'Apakah Anda bisa membaca buku berbahasa Inggris kah?',
    opsi: ['Saya Bisa', 'Saya Kurang Bisa (bisa sedikit-sedikit)', 'Saya Tidak bisa'],
  },
  { t: 'text', judul: 'No. HP Aktif / WA Dosen atau Guru Pembimbing PKL/Magang beserta Nama dan Jabatan', wajib: true },
  {
    t: 'choice', judul: 'Program Magang/PKL apa yang Anda minati?', wajib: true,
    opsi: ['Project Manager', 'Administrasi', 'HR', 'UI/UX', 'Programmer (Front End/ Backend)', 'Photographer', 'Videographer',
      'Disainer Grafis', 'Social Media Specialist', 'Content Writer', 'Content Planner', 'Marketing dan Sales',
      'Marcomm / Public Relation', 'Digital Marketing', 'Tiktok Creator', 'Las', 'CS', 'Yang lain'],
  },
  { t: 'text', judul: 'Software design yang Anda kuasai (isi strip jika bukan minat ini)' },
  { t: 'text', judul: 'Software editing video yang Anda kuasai (isi strip jika bukan minat ini)' },
  { t: 'text', judul: 'Bahasa Pemrograman yang sudah Anda kuasai (isi strip jika bukan minat ini)' },
  {
    t: 'choice', judul: 'Materi Digital Marketing yang ingin dipraktikkan',
    opsi: ['Digital Marketing Organic', 'Digital Marketing Ads (Fb Ads / Ig Ads)', 'Yang lain'],
  },
  { t: 'choice', judul: 'Apakah Anda ada alat kerja sendiri (LAPTOP) yang bisa dipakai selama Magang/PKL?', opsi: ['YA ADA', 'TIDAK ADA'], wajib: true },
  {
    t: 'choice', judul: 'Jika YA, alat apa yang Anda miliki?',
    opsi: ['Laptop yang sudah terinstal Corel dan Photoshop', 'Laptop yang sudah terinstal Adobe Premiere Pro/Final Cut Pro/Adobe After Effect',
      'Kamera DSLR', 'Laptop/ Netbook', 'Yang lain'],
  },
  { t: 'text', judul: 'Kapan Anda berencana mulai Magang/PKL? beserta durasi magang', wajib: true },
  {
    t: 'choice', judul: 'Anda tahu info magang ini darimana?', wajib: true,
    opsi: ['Website', 'Instagram', 'Twitter', 'Glints', 'Youtube', 'Yang lain'],
  },
  { t: 'file', judul: 'Silakan upload : CV, Scan KTP/KTM, Portofolio disini. Format pdf.', hanyaPdf: true },
  { t: 'file', judul: 'Silakan upload : Portofolio (berupa gambar jpg atau link youtube bagi videografer)' },
  { t: 'paragraf', judul: 'Kegiatan Anda saat ini selain magang/pkl ada kah? Jika ada, mohon sebutkan apa saja :', wajib: true },
  { t: 'choice', judul: 'Apakah Anda membutuhkan info kost/kontrakakan dekat kantor?', opsi: ['YA', 'TIDAK'], wajib: true },
  { t: 'choice', judul: 'Apakah Anda sudah berkeluarga?', opsi: ['YA', 'TIDAK'], wajib: true },
  { t: 'text', judul: 'No. HP Aktif (WA) Wali/ Ortu peserta magang', wajib: true },
  { t: 'text', judul: 'Sosial Media (Instagram) yang dimiliki', wajib: true },
];

/**
 * Membuat Google Form sungguhan untuk tiap brand, lengkap dengan sheet jawabannya,
 * lalu langsung mendaftarkannya ke dashboard. Data contoh tidak dipakai lagi setelah ini.
 */
// Form uji coba cuma 6 pertanyaan; judulnya diambil persis dari daftar lengkap biar kolomnya dikenali dashboard.
const JUDUL_SINGKAT = [
  'Nama Lengkap',
  'Email',
  'No. HP Aktif yang bisa dihubungi dan chat WA',
  'Asal Sekolah/Kampus',
  'Program Magang/PKL apa yang Anda minati?',
  'Kapan Anda berencana mulai Magang/PKL? beserta durasi magang',
];

function pertanyaanDipakai_(singkat) {
  if (!singkat) return PERTANYAAN_FORM;
  return JUDUL_SINGKAT.map(function (j) {
    return PERTANYAAN_FORM.filter(function (q) { return q.judul === j; })[0];
  }).filter(function (q) { return !!q; });
}

function judulPertanyaan_(q) {
  // Kolom unggah berkas sengaja jadi teks: Google menolaknya kalau form tidak wajib login.
  return q.t === 'file' ? q.judul + ' (tulis link berkasnya)' : q.judul;
}

function cariSheetForm_(ss) {
  return ss.getSheets().filter(function (s) {
    const n = s.getName().toLowerCase();
    return n.indexOf('form responses') === 0 || n.indexOf('respons formulir') === 0 || n.indexOf('form response') === 0;
  })[0] || null;
}

/**
 * Sheet tempat form menyimpan jawabannya.
 * Catatan penting: Apps Script TIDAK punya form.getResponseSheet() — itu cuma ada di preview lokal.
 * Yang asli: form.getDestinationType() + form.getDestinationId(), lalu spreadsheet-nya dibuka.
 */
function sheetTanggapanForm_(form) {
  let tipe = '';
  let id = '';
  try { tipe = String(form.getDestinationType()); } catch (e) { return null; }
  if (tipe !== String(FormApp.DestinationType.SPREADSHEET)) return null;
  try { id = String(form.getDestinationId() || ''); } catch (e) { return null; }
  if (!id) return null;
  let ss = null;
  try { ss = SpreadsheetApp.openById(id); } catch (e) { return null; }
  const judul = safeGet_(form, 'getTitle');
  const sheet = ss.getSheetByName(judul) || ss.getSheetByName(judul + ' (Jawaban Formulir)') || cariSheetForm_(ss) || ss.getSheets()[0];
  return { ss: ss, sheet: sheet };
}

function safeGet_(obj, nama) {
  try { return String(obj[nama] ? obj[nama]() : '') || ''; } catch (e) { return ''; }
}

function punyaHeader_(nilai) {
  return barisJudul_(nilai) !== -1;
}

// Judul yang menandakan "ini baris judul beneran": cukup banyak muncul di satu baris.
const POLA_JUDUL_TAB = [
  /(^|[^a-z])timestamp([^a-z]|$)/, /stempel waktu/, /^email$/, /nama lengkap/, /no\. hp/,
  /asal sekolah|asal kampus|universitas|sekolah tinggi/, /program studi/,
];

function skorJudul_(row) {
  const sel = (row || []).map(teks_).map(function (s) { return s.toLowerCase(); });
  return POLA_JUDUL_TAB.reduce(function (n, pola) {
    return n + (sel.some(function (s) { return pola.test(s); }) ? 1 : 0);
  }, 0);
}

/**
 * Baris judul TIDAK selalu baris pertama. Tab induk kantor menaruh blok keterangan warna
 * 8 baris di atas judulnya, jadi dicari sampai baris ke-20 dan dipilih yang paling banyak
 * judulnya (minimal 2). -1 = tidak dikenali, sama seperti dulu.
 */
function barisJudul_(nilai) {
  const daftar = nilai || [];
  const batas = Math.min(daftar.length, 20);
  let terbaik = -1;
  let skor = 1;
  for (let i = 0; i < batas; i++) {
    const s = skorJudul_(daftar[i]);
    if (s > skor) { skor = s; terbaik = i; }
  }
  return terbaik;
}

/** Memecah isi tab: baris judul, baris datanya, dan nomor baris sheet tempat data mulai. */
function pisahTab_(nilai) {
  const idx = barisJudul_(nilai);
  if (idx === -1) return { header: [], rows: [], mulai: 2, judul: -1 };
  return { header: nilai[idx] || [], rows: nilai.slice(idx + 1), mulai: idx + 2, judul: idx };
}

/**
 * Baris judul sebuah tab, dicari sampai baris ke-20 tanpa membaca seluruh tab.
 * Ini yang dipakai fungsi penulisan (status, jadwal, follow-up, sertifikat) supaya kolomnya
 * tidak salah di tab yang bagian atasnya berisi blok keterangan.
 */
function bacaHeaderInfo_(sheet) {
  const baris = Math.max(Math.min(sheet.getMaxRows(), 20), 1);
  const kolom = Math.max(sheet.getLastColumn(), 1);
  const nilai = sheet.getRange(1, 1, baris, kolom).getValues();
  const idx = barisJudul_(nilai);
  return { header: idx === -1 ? [] : (nilai[idx] || []), judul: idx };
}

function bacaHeader_(sheet) {
  return bacaHeaderInfo_(sheet).header;
}

function tulisHeader_(sheet, namaBrand) {
  // Menimpa baris judul di tab yang sudah ada isinya akan menggeser semua kolom. Jangan pernah lakukan itu.
  if (sheet.getLastRow() > 1) {
    throw new Error('Tab "' + sheet.getName() + '" sudah punya ' + (sheet.getLastRow() - 1) + ' baris data tapi baris judulnya belum dikenali. Betulkan judulnya manual di spreadsheet, jangan dari dashboard.');
  }
  const brand = getBrands_().filter(function (b) { return b.name === namaBrand; })[0] || {};
  const judul = (brand.kolomForm && brand.kolomForm.length) ? brand.kolomForm : PERTANYAAN_FORM.map(judulPertanyaan_);
  const header = ['Timestamp'].concat(judul);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  ensureAdminCols_(sheet);
  batalCache_(namaBrand);
}

// Isi sheet dibaca sekali saja per permintaan, supaya muat halaman tidak bolak-balik ke server.
var CACHE_NILAI_ = {};

function nilaiSheet_(kunci, sheet) {
  if (!CACHE_NILAI_[kunci]) CACHE_NILAI_[kunci] = sheet.getDataRange().getValues();
  return CACHE_NILAI_[kunci].slice();
}

function batalCache_(kunci) {
  if (kunci === undefined) CACHE_NILAI_ = {};
  else delete CACHE_NILAI_[kunci];
}

function simpanNamaSheet_(namaBrand, namaSheet) {
  const daftar = getBrands_();
  let berubah = false;
  daftar.forEach(function (b) {
    if (b.name === namaBrand && b.sheetName !== namaSheet) { b.sheetName = namaSheet; berubah = true; }
  });
  if (berubah) setBrands_(daftar);
}

/**
 * Jawaban yang sudah masuk di tab lama ikut dipindahkan ke tab form yang baru,
 * supaya dashboard tidak mendadak kosong setelah form lengkap dibuat.
 * Tab lama TIDAK diubah sedikit pun: tetap jadi arsip di spreadsheet yang sama.
 * Baris yang emailnya sudah ada di tab tujuan dilewati, jadi menekan tombol dua kali
 * tidak menghasilkan daftar pendaftar dobel.
 */
function salinJawabanLama_(sumber, tujuan) {
  if (!sumber || !tujuan || sumber.getName() === tujuan.getName()) return 0;
  const nilai = sumber.getDataRange().getValues();
  const headerLama = ((nilai[0] || [])).map(teks_);
  const headerBaru = ((tujuan.getDataRange().getValues()[0] || [])).map(teks_);
  if (headerLama.length < 2 || headerBaru.length < 2 || nilai.length < 2) return 0;

  const colBaru = kolom_(headerBaru);
  const colLama = kolom_(headerLama);
  const keLama = headerBaru.map(function (h, i) {
    const persis = headerLama.indexOf(h);
    if (persis !== -1) return persis;
    const kunci = Object.keys(ALIAS).filter(function (k) { return cariHeader_(headerBaru, k) === i; })[0];
    return kunci ? cariHeader_(headerLama, kunci) : -1;
  });

  const sudah = {};
  (tujuan.getDataRange().getValues().slice(1)).forEach(function (r) {
    const e = teks_(isi_(r, colBaru.email)).toLowerCase();
    if (e) sudah[e] = true;
  });

  const barisBaru = nilai.slice(1).filter(function (r) {
    if (!r.some(function (v) { return teks_(v) !== ''; })) return false;
    const e = teks_(isi_(r, colLama.email)).toLowerCase();
    return !e || !sudah[e];
  }).map(function (r) {
    return headerBaru.map(function (h, i) { return keLama[i] === -1 ? '' : r[keLama[i]]; });
  });
  if (!barisBaru.length) return 0;

  const mulai = Math.max(tujuan.getLastRow(), 1) + 1;
  tujuan.getRange(mulai, 1, barisBaru.length, headerBaru.length).setValues(barisBaru);
  return barisBaru.length;
}

function buatFormBrand(namaBrand, singkat) {
  butuhSesi_();
  const hasil = [];
  const gagal = [];
  const terdaftar = getBrands_();
  const formHidup = function (b) {
    if (!b || (!b.formId && !b.formUrl)) return false;
    const idForm = b.formId || ambilIdForm_(b.formUrl);
    try {
      return !DriveApp.getFileById(idForm).isTrashed();
    } catch (e) {
      try { FormApp.openById(idForm); return true; } catch (e2) { return false; }
    }
  };
  const hidup = {};
  let adaTautanMati = false;
  terdaftar.forEach(function (b) {
    hidup[b.name] = formHidup(b);
    if (!hidup[b.name] && (b.formId || b.formUrl)) {
      // Form-nya sudah dibuang tapi konfigurasi masih menyebut punya form: tautannya dikosongkan.
      b.formId = '';
      b.formUrl = '';
      adaTautanMati = true;
    }
  });
  if (adaTautanMati) setBrands_(terdaftar);
  const sudahAdaForm = function (nama) { return !!hidup[nama]; };
  const target = namaBrand
    ? [String(namaBrand)]
    : CONTOH_BRAND.filter(function (n) { return !sudahAdaForm(n); });
  const dilewati = CONTOH_BRAND.filter(function (n) { return target.indexOf(n) === -1 && sudahAdaForm(n); });

  target.forEach(function (brand) {
    try {
      // Form baru otomatis punya spreadsheet jawabannya sendiri, jadi tidak perlu bikin file kedua.
      const form = FormApp.create(namaFormBrand_(brand, singkat));
      const pertanyaan = pertanyaanDipakai_(singkat);
      form.setDescription('Form pendaftaran magang ' + brand + ' (form contoh, bukan form resmi kantor). Jawabannya masuk ke tab milik ' + brand + ' di spreadsheet admin.');

      // Tiap brand punya daftar bidang sendiri, jadi form-nya tidak identik dan pilihannya
      // sama dengan yang nanti muncul di grafik brand itu.
      const idxBrand = CONTOH_BRAND.indexOf(brand);
      const minatBrand = idxBrand === -1 || !PROFIL_CONTOH[idxBrand] ? [] : PROFIL_CONTOH[idxBrand].minat.filter(function (m, i, a) {
        return a.indexOf(m) === i;
      });

      const judulKolom = [];
      pertanyaan.forEach(function (q) {
        const judul = judulPertanyaan_(q);
        try {
          let item;
          if (q.t === 'paragraf') item = form.addParagraphTextItem();
          else if (q.t === 'choice') item = form.addMultipleChoiceItem();
          else item = form.addTextItem();
          if (q.t === 'choice') {
            const opsi = (/minati/i.test(q.judul) && minatBrand.length) ? minatBrand : q.opsi;
            item.setChoiceValues(opsi);
          }
          item.setTitle(judul);
          if (q.wajib) item.setRequired(true);
          judulKolom.push(judul);
        } catch (e) {
          gagal.push(brand + ' — kolom "' + judul + '": ' + (e.message || e));
        }
      });

      // Urutan aman: lihat dulu tujuan form yang asli (lewat getDestinationId), baru coba pindahkan
      // ke spreadsheet brand. Jangan mengulang setDestination ke file baru kalau Google sudah menolak.
      const lama = getBrands_().filter(function (b) { return b.name === brand; })[0] || {};
      let ss = null;
      let sheet = null;
      const bawaan = sheetTanggapanForm_(form);
      if (bawaan) { ss = bawaan.ss; sheet = bawaan.sheet; }

      if (lama.spreadsheetId && (!bawaan || bawaan.ss.getId() !== lama.spreadsheetId)) {
        try {
          const ssLama = SpreadsheetApp.openById(lama.spreadsheetId);
          form.setDestination(FormApp.DestinationType.SPREADSHEET, ssLama.getId());
          const setelah = sheetTanggapanForm_(form);
          ss = ssLama;
          sheet = (setelah && setelah.sheet) || ssLama.getSheetByName(lama.sheetName) || cariSheetForm_(ssLama) || ssLama.getSheets()[0];
        } catch (e) {
          gagal.push(brand + ': form sudah jadi, tapi jawabannya tidak bisa diarahkan ke spreadsheet brand (' + (e.message || e) + '). Pakai sheet bawaan form.');
        }
      }

      if (!ss) {
        ss = SpreadsheetApp.create('Data Pendaftaran ' + brand);
        try { form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId()); } catch (e) { gagal.push(brand + ': ' + (e.message || e)); }
        const setelah = sheetTanggapanForm_(form);
        sheet = (setelah && setelah.sheet) || ss.getSheets()[0];
      }
      if (!sheet) throw new Error('Google tidak memberi sheet jawaban untuk form ini. Buka form-nya sekali lewat browser, lalu sambungkan lewat "Tambah brand dari form asli".');
      if (!lama.spreadsheetId) {
        try { ss.setName('Data Pendaftaran ' + brand); } catch (e) { /* nama file bukan masalah besar */ }
      }

      if (!punyaHeader_(sheet.getDataRange().getValues())) {
        const header = ['Timestamp'].concat(judulKolom);
        sheet.getRange(1, 1, 1, header.length).setValues([header]);
      }
      ensureAdminCols_(sheet);
      let dipindah = 0;
      if (lama.sheetName) {
        let tabLama = null;
        try {
          const ssLama = lama.spreadsheetId ? SpreadsheetApp.openById(lama.spreadsheetId) : ss;
          tabLama = ssLama.getSheetByName(lama.sheetName);
        } catch (e) { tabLama = null; }
        dipindah = salinJawabanLama_(tabLama, sheet);
      }
      batalCache_(brand);

      const daftar = getBrands_().filter(function (b) { return b.name !== brand; });
      daftar.push({
        name: brand,
        spreadsheetId: ss.getId(),
        sheetName: sheet.getName(),
        formId: form.getId(),
        formUrl: form.getPublishedUrl(),
        kolomForm: judulKolom,
      });
      setBrands_(daftar);

      hasil.push({
        brand: brand,
        isian: form.getItems().length,
        dipindah: dipindah,
        formUrl: form.getPublishedUrl(),
        editUrl: form.getEditUrl(),
        spreadsheetUrl: ss.getUrl(),
        sheetName: sheet.getName(),
      });
    } catch (e) {
      gagal.push(brand + ': ' + (e.message || e));
    }
  });

  return { brands: hasil, gagal: gagal, dilewati: dilewati };
}

/**
 * Daftar file milik proyek ini di Drive, lengkap dengan tanda file yang masih dipakai dashboard.
 * Dipakai untuk membereskan sisa-sisa percobaan yang gagal.
 * Catatan: kueri DriveApp dipecah per mimeType (kueri gabungan dengan tanda kurung pernah ditolak
 * Google dengan "Argumen tidak valid: q"), nama difilter di JavaScript saja.
 */
function daftarFileProyek() {
  butuhSesi_();
  const dipakai = {};
  const namaFormDipakai = {};
  getBrands_().forEach(function (b) {
    if (b.spreadsheetId) dipakai[b.spreadsheetId] = true;
    if (b.formId) dipakai[b.formId] = true;
    if (b.formUrl) {
      namaFormDipakai[namaFormBrand_(b.name)] = true;
      namaFormDipakai['Form Pendaftaran Magang - ' + b.name] = true;
    }
  });
  const idLog = PropertiesService.getScriptProperties().getProperty(PROP_LOG);
  if (idLog) dipakai[idLog] = true;

  const KATA = ['Data Pendaftaran', 'Form Pendaftaran Magang', 'Log Follow-up Pendaftaran', 'Sertifikat'];
  const cocokNama = function (nama) {
    return KATA.some(function (k) { return nama.indexOf(k) !== -1; });
  };
  const hasil = [];
  const kueri = [
    'mimeType = "' + MimeType.GOOGLE_FORMS + '"',
    'mimeType = "' + MimeType.GOOGLE_SPREADSHEETS + '"',
  ];
  kueri.forEach(function (q) {
    const it = DriveApp.searchFiles(q);
    while (it.hasNext()) {
      const f = it.next();
      const nama = f.getName();
      if (!cocokNama(nama)) continue;
      const form = f.getMimeType() === MimeType.GOOGLE_FORMS;
      hasil.push({
        id: f.getId(),
        nama: nama,
        jenis: form ? 'form' : 'sheet',
        url: f.getUrl(),
        dibuat: tgl_(f.getDateCreated(), true),
        dipakai: !!dipakai[f.getId()] || (form && !!namaFormDipakai[nama]),
      });
    }
  });
  return hasil;
}

/**
 * Membuat form untuk brand yang BELUM punya form saja, isinya pertanyaan lengkap
 * (32 butir, sama seperti form resmi kantor). Singkat=true dipakai kalau cuma mau
 * menguji jalur jawaban form ke tab tanpa mengisi 32 butir.
 * Brand yang sudah punya form dilewati, jadi tombol ini tidak bisa menumbuhkan form kedua
 * untuk brand yang sama (itu sumber tab liar kemarin). Setiap form yang jadi langsung
 * dicatat sebagai konfigurasi di tab Integrasi.
 */
function buatFormContoh(singkat) {
  butuhSesi_();
  const hasil = buatFormBrand(null, singkat === true);
  const dicatat = [];
  (hasil.brands || []).forEach(function (b) {
    const merek = teks_(b.brand);
    const brand = getBrands_().filter(function (x) { return x.name === merek; })[0];
    if (!brand) { dicatat.push({ brand: merek, error: 'brand tidak ketemu di daftar' }); return; }
    try {
      const sudah = integrasiList_().filter(function (r) { return r['Form ID'] === teks_(brand.formId); })[0];
      const namaKonfig = 'Form ' + merek;
      if (sudah) { dicatat.push({ brand: merek, konfigurasi: sudah['Nama Konfigurasi'], sudahAda: true }); return; }
      tambahIntegrasi({
        nama: namaKonfig, brand: merek, program: '', urlForm: brand.formUrl || '',
        urlSpreadsheet: brand.spreadsheetId, namaSheet: brand.sheetName,
        deskripsi: 'Form contoh dibuat dari aplikasi', status: 'Aktif',
      });
      dicatat.push({ brand: merek, konfigurasi: namaKonfig, sheet: brand.sheetName, isian: b.isian, dipindah: b.dipindah || 0 });
    } catch (e) {
      dicatat.push({ brand: merek, error: String(e.message || e) });
    }
  });
  return { ok: true, dibuat: dicatat, gagal: hasil.gagal || [], dilewati: hasil.dilewati || [] };
}

/**
 * Tab jawaban yang tertinggal: bukan tab yang dipakai brand mana pun, dan namanya mirip
 * tab buatan Google Form ("Form Responses N", "Form ... (CONTOH)").
 */
function daftarTabSisa() {
  butuhSesi_();
  const files = {};
  const dipakai = {};
  getBrands_().forEach(function (b) {
    if (!b.spreadsheetId) return;
    files[b.spreadsheetId] = true;
    dipakai[b.spreadsheetId + '|' + b.sheetName] = true;
  });
  const out = [];
  Object.keys(files).forEach(function (id) {
    let ss = null;
    try { ss = SpreadsheetApp.openById(id); } catch (e) { return; }
    ss.getSheets().forEach(function (s) {
      const nama = s.getName();
      if (dipakai[id + '|' + nama]) return;
      const n = nama.toLowerCase();
      const miripTabForm = n.indexOf('form responses') === 0 || n.indexOf('respons formulir') === 0
        || n.indexOf('form ') === 0 || n.indexOf('form pendaftaran') === 0;
      if (!miripTabForm) return;
      out.push({ spreadsheetId: id, spreadsheet: ss.getName(), nama: nama, baris: Math.max(s.getLastRow() - 1, 0) });
    });
  });
  return out;
}

/**
 * Google melarang menghapus tab yang masih ditauti form. Ini melepas tautan form yang
 * menulis ke spreadsheet itu TAPI tidak dipakai brand mana pun (form sisa), jadi tab-nya
 * bisa dihapus. Form-nya sendiri tidak dihapus — cuma tidak lagi menulis ke situ.
 */
function putuskanTautanForm_(spreadsheetId) {
  const dipakai = {};
  getBrands_().forEach(function (b) { if (b.formId) dipakai[b.formId] = true; });
  const dilepas = [];
  let it = null;
  try {
    it = DriveApp.searchFiles('mimeType = "' + MimeType.GOOGLE_FORMS + '"');
  } catch (e) {
    return { dilepas: dilepas, alasan: 'Daftar form di Drive belum boleh dibaca (' + (e.message || e) + ').' };
  }
  let dicek = 0;
  while (it.hasNext() && dicek < 100) {
    const berkas = it.next();
    dicek++;
    if (dipakai[berkas.getId()]) continue;
    let form = null;
    try { form = FormApp.openById(berkas.getId()); } catch (e) { continue; }
    let tujuan = '';
    try { tujuan = String(form.getDestinationId() || ''); } catch (e) { continue; }
    if (tujuan !== String(spreadsheetId)) continue;
    try {
      form.removeDestination();
      dilepas.push(berkas.getName());
    } catch (e) { /* form masih tidak mau dilepas, tab-nya nanti tetap ditolak */ }
  }
  return { dilepas: dilepas, alasan: '' };
}

/** Hanya tab kosong yang dihapus. Tab berisi data apa pun ditolak, meski tidak dipakai brand. */
function buangTabSisa(daftar) {
  butuhSesi_();
  const dipakai = {};
  getBrands_().forEach(function (b) { dipakai[b.spreadsheetId + '|' + b.sheetName] = true; });
  const dihapus = [];
  const ditolak = [];
  const dilepas = [];
  const sudahDilepas = {};

  function hapusSatu(ss, nama) {
    const sheet = ss.getSheetByName(nama);
    if (!sheet) return 'tabnya sudah tidak ada';
    if (ss.getSheets().length <= 1) return 'ini tab satu-satunya di file itu';
    const baris = Math.max(sheet.getLastRow() - 1, 0);
    if (baris > 0) return 'masih ada ' + baris + ' baris data';
    try {
      ss.deleteSheet(sheet);
      return '';
    } catch (e) {
      return String(e.message || e);
    }
  }

  (daftar || []).forEach(function (item) {
    const nama = teks_(item && item.nama);
    if (!nama) return;
    if (dipakai[item.spreadsheetId + '|' + nama]) { ditolak.push(nama + ': masih dipakai brand'); return; }
    let ss = null;
    try { ss = SpreadsheetApp.openById(item.spreadsheetId); } catch (e) { ditolak.push(nama + ': spreadsheet tidak bisa dibuka'); return; }
    let gagal = hapusSatu(ss, nama);
    if (gagal && /formulir|form/i.test(gagal)) {
      // Tab masih ditauti form: lepas dulu sekali, lalu coba hapus lagi.
      if (!sudahDilepas[item.spreadsheetId]) {
        const lepas = putuskanTautanForm_(item.spreadsheetId);
        sudahDilepas[item.spreadsheetId] = true;
        dilepas.push.apply(dilepas, lepas.dilepas);
        if (lepas.alasan) { ditolak.push(nama + ': ' + lepas.alasan); return; }
      }
      gagal = hapusSatu(ss, nama);
    }
    if (gagal) ditolak.push(nama + ': ' + gagal + ' — buang dulu form yang menautkannya lewat tombol "Buang form ke Trash"');
    else dihapus.push(nama);
  });
  batalCache_();
  return { dihapus: dihapus, ditolak: ditolak, dilepas: dilepas };
}

/** Laporan sisa: tab tertinggal + form CONTOH yang tidak dipakai dashboard. */
function daftarSisa() {
  butuhSesi_();
  const tab = daftarTabSisa();
  let form = [];
  let catatan = '';
  try {
    form = daftarFileProyek().filter(function (f) { return f.jenis === 'form' && !f.dipakai; });
  } catch (e) {
    catatan = 'Daftar form di Drive belum bisa dibaca (' + (e.message || e) + '). Izinkan akses Drive sekali: script.google.com → Run → Review permissions → Allow. Tab tetap bisa dibersihkan.';
  }
  return { ok: true, tab: tab, form: form, catatan: catatan };
}

/**
 * Memindahkan file ke Trash Drive (bukan hapus permanen, masih bisa dipulihkan 30 hari).
 * Hanya file proyek ini dan hanya yang tidak dipakai dashboard.
 */
function bersihkanFileProyek(daftarId) {
  butuhSesi_();
  const peta = {};
  daftarFileProyek().forEach(function (f) { peta[f.id] = f; });
  const diTrash = [];
  const ditolak = [];
  (daftarId || []).forEach(function (id) {
    const f = peta[id];
    if (!f) { ditolak.push(String(id) + ': bukan file proyek'); return; }
    if (f.dipakai) { ditolak.push(f.nama + ': masih dipakai dashboard'); return; }
    try {
      DriveApp.getFileById(id).setTrashed(true);
      diTrash.push(f.nama);
    } catch (e) {
      ditolak.push(f.nama + ': ' + (e.message || e));
    }
  });
  return { diTrash: diTrash, ditolak: ditolak };
}

// Nama akun yang membuka, dipakai buat kartu di bawah sidebar.
function emailAktif_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/**
 * Akun yang PAKAI aplikasi untuk membuka spreadsheet/form — bukan akun yang sedang
 * membuka halaman. Ini yang harus ditulis di kolom "Bagikan" kalau filenya punya orang lain.
 */
function akunAplikasi_() {
  try {
    const efektif = Session.getEffectiveUser && Session.getEffectiveUser().getEmail();
    return teks_(efektif) || emailAktif_();
  } catch (e) {
    return emailAktif_();
  }
}

/**
 * Pesan izin dari Google cuma bilang "tidak punya akses" tanpa bilang harus ngapain.
 * Dijadikan satu langkah yang bisa langsung dikerjakan.
 */
function pesanIzin_(e) {
  const asli = String((e && e.message) || e || '');
  if (!/tidak memiliki izin|don't have permission|do not have permission|no permission|access denied|not accessible/i.test(asli)) return asli;
  const akun = akunAplikasi_();
  if (akun) {
    return 'File itu belum boleh dibuka oleh aplikasi. Aplikasi jalan sebagai akun ' + akun +
      ' — buka spreadsheet-nya → Bagikan → tambahkan ' + akun + ' sebagai Editor, lalu simpan lagi.' +
      ' (Kalau linknya kamu ambil dari akun Google lain, pakai akun ' + akun + ' dulu.)';
  }
  return 'File itu belum boleh dibuka aplikasi karena deployment-nya jalan sebagai "siapa saja, termasuk anonim", ' +
    'jadi aplikasi tidak punya akun Google sama sekali. Dua cara: (a) buka spreadsheet-nya → Bagikan → "Siapa saja yang memiliki link" minimal Pembaca, ' +
    'atau (b) buat deployment baru dengan "Jalankan sebagai: Saya" (akses tetap Siapa saja).';
}

function getBootstrap() {
  const brands = getBrands_();
  const pertama = brands[0];
  return {
    brands: brands.map(function (b) { return b.name; }),
    emailAktif: emailAktif_(),
    akunAplikasi: sesiSah_() ? akunAplikasi_() : '',
    akunSebenarnya: (function () {
      try { return teks_(Session.getEffectiveUser().getEmail()); } catch (e) { return ''; }
    })(),
    statusList: STATUS_LIST,
    spreadsheetUrl: (pertama && sesiSah_()) ? 'https://docs.google.com/spreadsheets/d/' + pertama.spreadsheetId + '/edit' : '',
    setupNeeded: brands.length === 0,
    butuhLogin: butuhLogin_(),
    versi: VERSI_APLIKASI,
  };
}

function getOverview() {
  butuhSesi_();
  return getBrands_().map(function (brand) {
    try {
      const sheet = getBrandSheet_(brand.name);
      const info = bacaHeaderInfo_(sheet);
      if (info.judul === -1) throw new Error('baris judul tab belum dikenali');
      const col = kolom_(info.header.map(teks_));
      const arsip = brandArsip_(brand.name);
      const jumlah = Math.max(sheet.getLastRow() - (info.judul + 1), 0);
      let total = 0;
      let baru = 0;
      if (jumlah > 0) {
        // Menghitung cukup 4 kolom: Nama, Email, Timestamp, Status. Membaca 42 kolom x 5.516 baris
        // pada setiap refresh itulah yang bikin satu klik terasa puluhan detik.
        const baca = function (k) {
          const i = col[k];
          if (i === undefined || i < 0) return null;
          return sheet.getRange(info.judul + 2, i + 1, jumlah, 1).getValues();
        };
        const nama = baca('nama');
        const email = baca('email');
        const waktu = baca('timestamp');
        const status = arsip ? null : baca('Status');
        for (let i = 0; i < jumlah; i++) {
          if (!teks_(nama && nama[i][0]) && !teks_(email && email[i][0])) continue;
          total++;
          const ts = waktu ? toDate_(waktu[i][0]) : null;
          if (status && masihBaru_(ts, normalisasiStatus_(teks_(status[i][0])))) baru++;
        }
      }
      return { brand: brand.name, sheet: sheet.getName(), total: total, baru: baru, contoh: contohBrand_(brand.name) };
    } catch (e) {
      return { brand: brand.name, sheet: brand.sheetName || '', total: 0, baru: 0, contoh: contohBrand_(brand.name), error: pesanIzin_(e) };
    }
  });
}

/** Brand buatan aplikasi (bukan form kantor) — ditandai di sidebar supaya tidak ketuker. */
function contohBrand_(namaBrand) {
  const n = teks_(namaBrand).toLowerCase();
  return CONTOH_BRAND.some(function (b) { return String(b).toLowerCase() === n; });
}

/**
 * Satu panggilan untuk membuka dashboard: data awal + jumlah per brand + isi halaman yang lagi dibuka.
 * Dipakai supaya muat pertama tidak bolak-balik ke server.
 */
function getInitData(halaman, brandAktif) {
  batalCache_();
  const bootstrap = getBootstrap();
  const hasil = { bootstrap: bootstrap, page: halaman || 'dashboard' };
  if (!sesiSah_()) return hasil;
  if (bootstrap.setupNeeded) return hasil;

  const daftar = getBrands_().map(function (b) { return b.name; });
  const brand = daftar.indexOf(brandAktif) !== -1 ? brandAktif : daftar[0];
  hasil.brand = brand;
  hasil.overview = getOverview();

  try {
    const total = hasil.overview.reduce(function (a2, o) { return a2 + (o.total || 0); }, 0);
    // Di bawah batas ini semua brand ikut dikirim sekali, jadi pindah menu/brand tidak perlu nunggu server.
    const bawaSemua = total > 0 && total <= 600;
    // Menu Dashboard tidak lagi membaca penuh tab brand aktif: gabungan dihitung dengan
    // kolom seperlunya, dan daftar pendaftar baru dibaca saat halamannya benar-benar dibuka.
    const butuhRows = hasil.page !== 'pengaturan' && hasil.page !== 'integrasi' && hasil.page !== 'dashboard';
    const dataAktif = butuhRows ? getBrandData(brand) : null;
    if (hasil.page === 'pengaturan' || hasil.page === 'integrasi') {
      hasil.konfigurasi = getKonfigurasi();
    } else if (hasil.page === 'laporan') {
      hasil.laporan = getLaporan(brand, '', '');
    } else if (hasil.page === 'dashboard') {
      hasil.dashboardGabungan = getDashboard('');
    } else {
      hasil.rowsData = dataAktif;
    }
    // Sekalian kirim data semua brand (tab-nya sudah terbaca di getOverview) supaya pindah menu dan
    // pindah brand tidak perlu panggilan server lagi, termasuk saat buka Pengaturan atau Laporan.
    if (bawaSemua) {
      hasil.rowsSemua = {};
      hasil.dashboardSemua = {};
      hasil.laporanPerBrand = {};
      daftar.forEach(function (nama) {
        let d = null;
        try { d = (nama === brand && dataAktif) ? dataAktif : getBrandData(nama); } catch (e) { return; }
        hasil.rowsSemua[nama] = d;
        try { hasil.dashboardSemua[nama] = getDashboard(nama, d); } catch (e) { /* dilewati */ }
        try { hasil.laporanPerBrand[nama] = getLaporan(nama, '', ''); } catch (e) { /* dilewati */ }
      });
      // Lingkup "Semua brand" juga disiapkan supaya buka Laporan tidak perlu nunggu lagi.
      try { hasil.laporanSemua = getLaporan(null, '', ''); } catch (e) { /* dilewati */ }
    }
  } catch (e) {
    hasil.galat = String(e.message || e);
  }
  return hasil;
}

/**
 * Baris ringan untuk rekap gabungan semua brand: hanya kolom yang benar-benar dipakai
 * tile/grafik yang dibaca. Membaca 42 kolom x ribuan baris per brand itulah yang membuat
 * Dashboard gabungan terasa lama, jadi kolomnya dibatasi dan sisanya diambil per baris
 * lewat Detail seperti biasa.
 */
function barisRingkas_(namaBrand) {
  const sheet = getBrandSheet_(namaBrand);
  const info = bacaHeaderInfo_(sheet);
  const arsip = brandArsip_(namaBrand);
  if (info.judul === -1) return { rows: [], arsip: arsip, judul: -1 };
  const header = info.header.map(teks_);
  const col = kolom_(header);
  const jumlah = Math.max(sheet.getLastRow() - (info.judul + 1), 0);
  const kunci = ['nama', 'email', 'timestamp', 'Status', 'minat', 'kampus', 'Follow-up Terakhir', 'Jadwal Wawancara', 'cv', 'portofolio'];
  const baca = {};
  kunci.forEach(function (k) {
    const i = col[k];
    if (i === undefined || i < 0 || jumlah === 0) { baca[k] = null; return; }
    baca[k] = sheet.getRange(info.judul + 2, i + 1, jumlah, 1).getValues();
  });
  const sel = function (k, i) { return baca[k] ? teks_(baca[k][i][0]) : ''; };
  const logMap = getFollowUpMap_();
  const keluar = [];

  for (let i = 0; i < jumlah; i++) {
    const nama = sel('nama', i);
    const email = sel('email', i);
    if (!nama && !email) continue;
    const ts = toDate_(baca.timestamp ? baca.timestamp[i][0] : '');
    const status = normalisasiStatus_(sel('Status', i));
    const contoh = contohByEmail_(email);
    const berkasPunya = ['cv', 'portofolio'].filter(function (k) { return !!baca[k]; }).length;
    const berkasAda = ['cv', 'portofolio'].filter(function (k) {
      return !!(sel(k, i) || (contoh && teks_(contoh[k])));
    }).length;
    keluar.push({
      row: info.judul + 2 + i,
      brand: namaBrand,
      nama: nama,
      email: email,
      status: status,
      minat: sel('minat', i) || (contoh ? teks_(contoh.minat) : ''),
      kampus: sel('kampus', i) || (contoh ? teks_(contoh.kampus) : ''),
      timestamp: ts ? tgl_(ts) : '-',
      iso: ts ? ts.toISOString() : '',
      baru: !arsip && masihBaru_(ts, status),
      arsip: arsip,
      jumlahFollowUp: (logMap[key_(namaBrand, email)] || []).length,
      followUpTerakhir: sel('Follow-up Terakhir', i),
      jadwal: sel('Jadwal Wawancara', i),
      berkas: { ada: berkasAda, total: contoh ? 2 : berkasPunya },
    });
  }
  return { rows: keluar, arsip: arsip, judul: info.judul };
}

/** Semua baris ringan dari semua brand; brand yang tidak bisa dibuka dilewati, tidak mematikan rekap. */
function barisSemuaBrand_() {
  const keluar = [];
  const dilewati = [];
  const merek = [];
  getBrands_().forEach(function (b) {
    merek.push(b.name);
    try {
      keluar.push.apply(keluar, barisRingkas_(b.name).rows);
    } catch (e) {
      dilewati.push(b.name + ': ' + pesanIzin_(e));
    }
  });
  return { rows: keluar, dilewati: dilewati, merek: merek };
}

function getDashboard(namaBrand, data) {
  butuhSesi_();
  const gabungan = !teks_(namaBrand);
  const hasil = data || (gabungan ? barisSemuaBrand_() : getBrandData(namaBrand));
  const rows = hasil.rows;
  const terbaru = gabungan
    ? rows.slice().sort(function (a, b) { return (b.iso || '') < (a.iso || '') ? -1 : 1; })
    : rows;
  return {
    stats: buildStats_(rows),
    gabungan: gabungan,
    jumlahBrand: gabungan ? (hasil.merek || []).length : 1,
    dilewati: hasil.dilewati || [],
    terbaru: terbaru.filter(function (r) { return r.status === 'Seleksi'; }).slice(0, 5).map(function (r) {
      return { row: r.row, brand: r.brand, nama: r.nama, minat: r.minat, status: r.status, timestamp: r.timestamp, baru: r.baru };
    }),
    perluTindakLanjut: gabungan
      ? rows.filter(function (r) { return !r.arsip && STATUS_WAWANCARA.indexOf(r.status) !== -1 && !r.jumlahFollowUp; }).length
      : ((hasil.arsip === undefined ? brandArsip_(namaBrand) : hasil.arsip) ? 0
        : rows.filter(function (r) { return STATUS_WAWANCARA.indexOf(r.status) !== -1 && !r.jumlahFollowUp; }).length),
    berkasBelumLengkap: rows.filter(function (r) { return STATUS_PROSES.indexOf(r.status) !== -1 && r.berkas.total && r.berkas.ada < r.berkas.total; }).length,
    sudahDijadwalkan: rows.filter(function (r) { return r.jadwal && STATUS_WAWANCARA.indexOf(r.status) !== -1; }).length,
    chartMinat: hitungPer_(rows, 'minat'),
    chartKampus: hitungPer_(rows, 'kampus'),
    trend: trendMinggu_(rows),
  };
}

/**
 * Menilai isi kolom "Link Sertifikat": file-nya masih ada atau tidak.
 * Link yang mati sengaja tidak ditampilkan supaya tidak ada tombol yang cuma bikin error.
 */
function berkasSertifikat_(link) {
  const teks = String(link || '').trim();
  if (!teks) return { ada: false, unduh: '', mati: false };
  const ketemu = teks.match(/\/d\/([^/?#]+)/) || teks.match(/[?&]id=([^&?#]+)/);
  if (!ketemu) return { ada: false, unduh: '', mati: true };
  try {
    const file = DriveApp.getFileById(ketemu[1]);
    if (file.isTrashed()) return { ada: false, unduh: '', mati: true };
    return { ada: true, unduh: 'https://drive.google.com/uc?export=download&id=' + ketemu[1], mati: false };
  } catch (e) {
    return { ada: false, unduh: '', mati: true };
  }
}

function getLaporan(namaBrand, dari, sampai) {
  butuhSesi_();
  // Nama brand kosong = laporan gabungan semua brand, ikut lingkup yang dipilih di sidebar.
  const gabungan = !namaBrand || namaBrand === 'SEMUA';
  const daftar = gabungan ? getBrands_().map(function (b) { return b.name; }) : [namaBrand];
  let rows = [];
  daftar.forEach(function (n) {
    try {
      rows = rows.concat(getBrandData(n).rows);
    } catch (e) { /* brand yang sheet-nya bermasalah dilewati, laporan tetap keluar */ }
  });
  rows.sort(function (a, b) { return String(b.iso || '').localeCompare(String(a.iso || '')); });
  rows = saringRentang_(rows, dari, sampai);
  const sertifikat = rows.filter(function (r) { return r.status === 'Lulus' || r.linkSertifikat; }).map(function (r) {
    const berkas = berkasSertifikat_(r.linkSertifikat);
    return {
      row: r.row,
      brand: r.brand,
      nama: r.nama,
      email: r.email,
      minat: r.minat,
      timestamp: r.timestamp,
      link: berkas.unduh,
      terkirim: berkas.ada,
      state: berkas.mati ? 'File hilang' : (berkas.ada ? 'Terkirim' : 'Belum dikirim'),
    };
  });
  return {
    rentang: { dari: dari || '', sampai: sampai || '' },
    lingkup: gabungan ? 'semua' : namaBrand,
    total: rows.length,
    stats: buildStats_(rows),
    perBulan: hitungPerBulan_(rows),
    perMinat: hitungPer_(rows, 'minat'),
    perKampus: hitungPer_(rows, 'kampus'),
    sertifikat: sertifikat,
    sudahTerkirim: sertifikat.filter(function (s) { return s.terkirim; }).length,
  };
}

/** Membuat berkas CSV di Google Drive dan mengembalikan tautannya. */
function eksporCsv(namaBrand, dari, sampai, nomorBaris) {
  butuhSesi_();
  const brand = findBrand_(namaBrand);
  const sheet = getBrandSheet_(namaBrand);
  const bagi = pisahTab_(nilaiSheet_(namaBrand, sheet));
  const header = bagi.header;
  const col = kolom_(header);
  const yangDipilih = {};
  (Array.isArray(nomorBaris) ? nomorBaris : []).forEach(function (n) { yangDipilih[Number(n)] = 1; });
  const adaFilter = Object.keys(yangDipilih).length > 0;
  const baris = bagi.rows.filter(function (row, i) {
    const ts = toDate_(isi_(row, col.timestamp));
    if (ts && !dalamRentang_(ts, dari, sampai)) return false;
    if (adaFilter && !yangDipilih[i + bagi.mulai]) return false;
    return true;
  });

  const isi = [header].concat(baris).map(function (row) {
    return row.map(function (v) {
      const s = v instanceof Date ? tgl_(v) : teks_(v);
      return '"' + s.replace(/"/g, '""') + '"';
    }).join(',');
  }).join('\r\n');

  const namaFile = 'Ekspor ' + brand.name + ' - ' + tgl_(new Date()).replace(/[:\/ ]/g, '-') + '.csv';
  const file = getFolderEkspor_().createFile(Utilities.newBlob('﻿' + isi, 'text/csv', namaFile));
  return { url: file.getUrl(), nama: namaFile, jumlahBaris: baris.length };
}

// Link web app yang sedang dipakai, supaya bisa disalin dari dalam dashboard.
function urlDashboard_() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return '';
  }
}

/**
 * Pemeriksaan singkat: apa yang bisa dibaca dan apa yang macet, ditampilkan di menu Pengaturan.
 */
function cekSistem() {
  butuhSesi_();
  const hasil = [];
  const catat = function (nama, ok, ket) { hasil.push({ nama: nama, ok: !!ok, ket: ket || '' }); };
  const brands = getBrands_();

  catat('Brand terdaftar', brands.length > 0, brands.length + ' brand: ' + brands.map(function (b) { return b.name; }).join(', '));

  brands.forEach(function (b) {
    try {
      const sheet = getBrandSheet_(b.name);
      const info = bacaHeaderInfo_(sheet);
      const header = info.header.map(teks_);
      const baris = Math.max(sheet.getLastRow() - (info.judul + 1), 0);
      const col = kolom_(header);
      const kurang = KOLOM_PENTING.filter(function (k) { return col[k] === -1; });
      catat(b.name + ' — tab "' + sheet.getName() + '"', info.judul !== -1,
        baris + ' pendaftar, ' + header.length + ' kolom' + (info.judul === -1 ? ' — tab ini belum ada baris judulnya'
          : (kurang.length ? ' — kolom utama belum lengkap: ' + kurang.join(', ') : ' — kolom utama lengkap')));
    } catch (e) {
      catat(b.name, false, String(e.message || e));
    }
  });

  // "Datanya kebaca tapi tidak bisa diedit" hampir selalu ada di sini: peran akses file-nya.
  brands.concat([{ name: 'file log (catatan follow-up)', spreadsheetId: teks_(PropertiesService.getScriptProperties().getProperty(PROP_LOG)) }])
    .forEach(function (b) {
      if (!b.spreadsheetId) return;
      try {
        const file = DriveApp.getFileById(b.spreadsheetId);
        let boleh = null;
        try { boleh = file.isEditor(); } catch (e) { boleh = null; }
        catat('Akses tulis — ' + b.name, boleh === true, boleh === true ? 'Editor · ' + file.getName()
          : (boleh === null ? 'Google tidak mau memberi tahu · ' + file.getName()
            : 'BARU VIEWER, tidak bisa menyimpan · ' + file.getName() + ' — bagikan ke ' + (akunAplikasi_() || 'akun yang menjalankan') + ' sebagai Editor'));
      } catch (e) {
        catat('Akses tulis — ' + b.name, false, pesanIzin_(e));
      }
    });

  const statusForm = brands.map(function (b) { return { name: b.name, f: cekFormBrand_(b) }; });
  const formNyambung = statusForm.filter(function (x) { return x.f.nyambung; });
  const formRusak = statusForm.filter(function (x) { return !x.f.nyambung; });
  catat('Google Form nyambung ke tab sekarang', brands.length > 0 && formNyambung.length === brands.length,
    formNyambung.length + ' dari ' + brands.length + ' brand' + (formRusak.length ? ' — ' + formRusak.map(function (x) { return x.name + ': ' + x.f.ket; }).join(' | ') : ''));

  try {
    catat('Spreadsheet log follow-up', true, getLogSs_().getName());
  } catch (e) {
    catat('Spreadsheet log follow-up', false, String(e.message || e));
  }
  try {
    catat('Folder sertifikat di Drive', true, getFolder_().getName());
  } catch (e) {
    catat('Folder sertifikat di Drive', false, String(e.message || e));
  }

  try {
    const b0 = brands[0];
    const header0 = nilaiSheet_(b0.name, getBrandSheet_(b0.name))[0] || [];
    const col0 = kolom_(header0);
    const penting = { nama: 'Nama', email: 'Email', hp: 'No. HP', kampus: 'Asal kampus', prodi: 'Program studi', minat: 'Bidang diminati', mulai: 'Rencana mulai' };
    const hilang = Object.keys(penting).filter(function (k) { return col0[k] === -1; }).map(function (k) { return penting[k]; });
    catat('Kolom form dikenali (dibaca dari tab ' + b0.name + ')', hilang.length === 0,
      hilang.length ? 'belum dikenali: ' + hilang.join(', ') + ' — cek judul kolomnya di spreadsheet' : Object.keys(penting).length + ' kolom utama aman');
  } catch (e) {
    catat('Kolom form dikenali', false, String(e.message || e));
  }

  return { versi: VERSI_APLIKASI, waktu: tgl_(new Date(), true), hasil: hasil };
}

/**
 * Mengecek apakah link form yang tersimpan benar-benar menulis ke tab yang dipakai sekarang.
 * Penting karena form dari percobaan lama masih menunjuk ke spreadsheet lama.
 */
function cekFormBrand_(b) {
  if (!b.formUrl) return { ada: false, nyambung: false, ket: 'belum ada form' };
  try {
    const form = b.formId ? FormApp.openById(b.formId) : FormApp.openByUrl(b.formUrl);
    const tujuan = sheetTanggapanForm_(form);
    if (!tujuan) return { ada: true, nyambung: false, ket: 'form ini belum menuliskan jawabannya ke spreadsheet mana pun — buka form-nya lalu sambungkan lewat menu "Integrasi Spreadsheet"' };
    const ss = tujuan.ss;
    const sheet = tujuan.sheet;
    const nyambung = ss.getId() === b.spreadsheetId && sheet.getName() === b.sheetName;
    return {
      ada: true,
      nyambung: nyambung,
      ket: nyambung ? '' : 'form ini menulis ke "' + ss.getName() + ' / ' + sheet.getName() + '", bukan tab yang dipakai dashboard',
    };
  } catch (e) {
    return { ada: true, nyambung: false, rusak: true, ket: 'form-nya sudah tidak ada di akun ini (kemungkinan sudah dihapus) — sambungkan form baru lewat menu Integrasi Spreadsheet' };
  }
}

const EMAIL_CONTOH = '@sample.invalid';

/**
 * Mengisi tab brand dengan beberapa pendaftar contoh supaya dashboard tidak kosong melompong.
 * Semua emailnya berakhiran @sample.invalid, jadi mudah dikenali dan bisa dihapus lagi.
 */
function isiDataContoh(namaBrand) {
  butuhSesi_();
  const sheet = getBrandSheet_(namaBrand, true);
  const nilai = nilaiSheet_(namaBrand, sheet);
  const header = pastikanKolomContoh_(sheet, namaBrand, nilai);
  if (!header.length) throw new Error('Tab ini belum punya baris judul.');
  // Baris contoh lama dibuang dulu, jadi tombol ini selalu mengganti dengan data contoh versi terbaru.
  const dihapus = buangBarisContoh_(sheet, nilai);

  const idx = CONTOH_BRAND.indexOf(namaBrand);
  if (idx === -1) throw new Error('Brand "' + namaBrand + '" bukan brand contoh, jadi tidak ada data contohnya. Isi tab ini lewat form-nya langsung.');
  const orang = CONTOH_ORANG.filter(function (o) { return o.brand === idx; })
    .sort(function (a, b) { return new Date(a.timestamp) - new Date(b.timestamp); });

  orang.forEach(function (o) { sheet.appendRow(barisContoh_(header, o)); });
  batalCache_(namaBrand);
  return { ok: true, jumlah: orang.length, brand: namaBrand, diganti: dihapus };
}

/** Menghapus baris yang emailnya @sample.invalid — data asli tidak tersentuh. */
function hapusDataContoh(namaBrand) {
  butuhSesi_();
  const sheet = getBrandSheet_(namaBrand, true);
  const dihapus = buangBarisContoh_(sheet, nilaiSheet_(namaBrand, sheet));
  batalCache_(namaBrand);
  return { ok: true, dihapus: dihapus };
}

function buangBarisContoh_(sheet, nilai) {
  const header = (nilai[0] || []).map(teks_);
  const col = kolom_(header);
  let dihapus = 0;
  for (let r = nilai.length; r >= 2; r--) {
    const email = teks_(isi_(nilai[r - 1], col.email)).toLowerCase();
    if (email.indexOf(EMAIL_CONTOH) !== -1) {
      sheet.deleteRow(r);
      dihapus++;
    }
  }
  return dihapus;
}

/**
 * Tab buatan form singkat tidak punya kolom asal kampus, jadi grafiknya kosong terus.
 * Kolomnya ditambahkan di baris judul, tapi hanya selama tab ini belum berisi jawaban form sungguhan.
 */
function pastikanKolomContoh_(sheet, namaBrand, nilai) {
  const header = (nilai[0] || []).map(teks_);
  if (!header.length) return header;
  const col = kolom_(header);
  if (col.kampus !== -1) return header;
  if (col.email === -1) return header;
  const adaAsli = nilai.slice(1).some(function (r) {
    const email = teks_(isi_(r, col.email)).toLowerCase();
    return !!email && email.indexOf(EMAIL_CONTOH) === -1;
  });
  if (adaAsli) return header;
  sheet.getRange(1, header.length + 1).setValue('Asal Sekolah/Kampus');
  header.push('Asal Sekolah/Kampus');
  return header;
}

// ---------- notifikasi pendaftar baru ----------
const PROP_NOTIF = 'NOTIF_PENDAFTAR_TERAKHIR';
const PROP_NOTIF_JAM = 'NOTIF_JAM_HARIAN';
const FUNGSI_NOTIF = 'kirimLaporanPendaftarBaru';

/** Pendaftar yang masuk setelah waktu `sejak`, diambil dari semua brand. */
function pendaftarSejak_(sejak) {
  const out = [];
  getBrands_().forEach(function (b) {
    let rows = [];
    try {
      rows = getBrandData(b.name).rows;
    } catch (e) {
      return; // brand bermasalah dilewati, jangan sampai email tidak terkirim
    }
    rows.forEach(function (r) {
      const ts = r.iso ? new Date(r.iso) : null;
      if (ts && ts.getTime() > sejak.getTime()) {
        out.push({ brand: b.name, nama: r.nama, email: r.email, tgl: r.timestamp, status: r.status, waktu: ts.getTime() });
      }
    });
  });
  return out.sort(function (a, b) { return a.waktu - b.waktu; });
}

/** Dipanggil trigger harian dan trigger "ada tanggapan form". Tidak pernah mengirim email kosong. */
function kirimLaporanPendaftarBaru() {
  SESI_TOKEN_PAKAI_ = sesiBuat_();
  const props = PropertiesService.getScriptProperties();
  const terakhir = toDate_(props.getProperty(PROP_NOTIF)) || new Date(0);
  const daftar = pendaftarSejak_(terakhir);
  if (!daftar.length) return { ok: true, jumlah: 0, dikirim: false };
  const ke = emailAktif_() || props.getProperty('EMAIL_ADMIN') || '';
  if (!ke) throw new Error('Email tujuan tidak terbaca. Jalankan fungsi ini dari akun Google yang mau dikirimi laporan.');
  const isi = daftar.map(function (d) {
    return '- ' + d.nama + ' (' + d.email + ') — ' + d.brand + ' — daftar ' + d.tgl + ' — status: ' + d.status;
  }).join('\n');
  MailApp.sendEmail({
    to: ke,
    subject: '[' + daftar.length + ' pendaftar baru] Dashboard Pendaftaran Magang',
    body: 'Pendaftar baru sejak ' + (terakhir.getTime() ? tgl_(terakhir, true) : 'cek pertama') + ':\n\n'
      + isi + '\n\nBuka dashboard: ' + urlDashboard_(),
  });
  props.setProperty(PROP_NOTIF, new Date(Math.max(terakhir.getTime(), daftar[daftar.length - 1].waktu)).toISOString());
  return { ok: true, jumlah: daftar.length, dikirim: true, ke: ke };
}

function triggerNotif_() {
  return ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === FUNGSI_NOTIF;
  });
}

/** Pengingat harian: satu email berisi semua pendaftar baru sejak cek terakhir. */
function pasangPengingatHarian(jam) {
  const jamKerja = Math.min(Math.max(Number(jam) || 7, 0), 23);
  triggerNotif_().forEach(function (t) {
    if (String(t.getTriggerSource()).indexOf('CLOCK') !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(FUNGSI_NOTIF).timeBased().atHour(jamKerja).everyDays(1).create();
  PropertiesService.getScriptProperties().setProperty(PROP_NOTIF_JAM, String(jamKerja));
  return { ok: true, jam: jamKerja };
}

/** Notifikasi seketika: tiap ada tanggapan form, email langsung dikirim. */
function pasangTriggerSubmitForm() {
  const sudah = {};
  triggerNotif_().forEach(function (t) {
    if (String(t.getTriggerSource()).indexOf('FORM') !== -1) sudah[t.getTriggerSourceId()] = true;
  });
  const dipasang = [];
  const gagal = [];
  getBrands_().forEach(function (b) {
    if (!b.formId || sudah[b.formId]) return;
    try {
      ScriptApp.newTrigger(FUNGSI_NOTIF).forForm(FormApp.openById(b.formId)).onFormSubmit().create();
      dipasang.push(b.name);
    } catch (e) {
      gagal.push(b.name + ': ' + (e.message || e));
    }
  });
  return { ok: true, dipasang: dipasang, gagal: gagal };
}

function hapusPengingatNotif() {
  const daftar = triggerNotif_();
  daftar.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  return { ok: true, dihapus: daftar.length };
}

function getKonfigurasi() {
  butuhSesi_();
  const props = PropertiesService.getScriptProperties();
  let logUrl = '', logNama = '', logBaris = 0;
  try {
    const logSs = SpreadsheetApp.openById(props.getProperty(PROP_LOG));
    logUrl = logSs.getUrl();
    logNama = logSs.getName();
    const sheet = logSs.getSheetByName(LOG_SHEET);
    if (sheet) logBaris = Math.max(sheet.getLastRow() - 1, 0);
  } catch (e) {
    logNama = 'belum dibuat';
  }
  const folderSertifikat = urlFolderSertifikat_();
  const hasil = {
    versi: VERSI_APLIKASI,
    urlDashboard: urlDashboard_(),
    statusList: STATUS_LIST,
    loginAktif: butuhLogin_(),
    loginUser: String(PropertiesService.getScriptProperties().getProperty(PROP_LOGIN_USER) || ''),
    folderSertifikat: folderSertifikat,
    logUrl: logUrl,
    logNama: logNama,
    logBaris: logBaris,
    fileKonfigurasi: logNama,
    integrasi: integrasiDenganStatus_(),
    brands: (function () {
      // Link form yang disimpan lewat menu Integrasi juga dipakai kartu Pengaturan,
      // supaya brand tidak kelihatan "belum ada form" padahal konfigurasinya sudah ada.
      const formKu = {};
      integrasiBacaSafing_().daftar.forEach(function (r) {
        const br = teks_(r.Brand);
        if (br && !formKu[br] && (teks_(r['Form ID']) || teks_(r['URL Google Form']))) {
          formKu[br] = { formId: teks_(r['Form ID']), formUrl: teks_(r['URL Google Form']) };
        }
      });
      return getBrands_().map(function (b) {
        const info = { name: b.name, spreadsheetId: b.spreadsheetId, sheetName: b.sheetName, formUrl: b.formUrl || '', url: '', spreadsheetName: '', sheetAda: false, jumlahBaris: 0, contoh: CONTOH_BRAND.indexOf(b.name) !== -1 };
        try {
          const ss = SpreadsheetApp.openById(b.spreadsheetId);
          info.spreadsheetName = ss.getName();
          info.url = ss.getUrl();
          const sheet = ss.getSheetByName(b.sheetName) || cariSheetForm_(ss);
          info.sheetAda = !!sheet;
          if (sheet) info.jumlahBaris = Math.max(sheet.getLastRow() - 1, 0);
        } catch (e) {
          info.galat = pesanIzin_(e);
        }
        const sumber = Object.assign({}, b, formKu[b.name] || {});
        info.formUrl = teks_(sumber.formUrl);
        try {
          info.form = cekFormBrand_(sumber);
        } catch (e) {
          info.form = { ada: false, nyambung: false, ket: 'tidak bisa diperiksa: ' + pesanIzin_(e) };
        }
        return info;
      });
    })(),
  };
  // Satu sumber yang tidak bisa dibuka tidak boleh mematikan seluruh halaman:
  // cukup ditandai, sisanya tetap tampil.
  if (KONFIG_GALAT_) hasil.peringatan = 'Tidak bisa mengambil data di konfigurasi ini: ' + KONFIG_GALAT_;
  if (!folderSertifikat) hasil.peringatan = (hasil.peringatan ? hasil.peringatan + ' — ' : '') + 'Folder sertifikat belum bisa dibuka.';
  return hasil;
}

/** Folder sertifikat dicari lewat Drive; kalau aplikasinya tidak boleh membuka Drive, jangan bikin halaman mati. */
function urlFolderSertifikat_() {
  try {
    return getFolder_().getUrl();
  } catch (e) {
    Logger.log('Folder sertifikat tidak bisa dibuka: ' + (e.message || e));
    return '';
  }
}

function hapusBrand(namaBrand) {
  butuhSesi_();
  setBrands_(getBrands_().filter(function (b) { return b.name !== namaBrand; }));
  return { ok: true };
}

/**
 * Tombol hapus di kartu brand: melepas brand dari dashboard DAN membuang kartu
 * konfigurasinya sekalian, supaya tidak ada lagi nama yang tertinggal.
 * Yang TIDAK disentuh: isi spreadsheet, tab jawaban, form aslinya. Data pendaftar selalu aman.
 */
function hapusBrandLengkap(namaBrand) {
  butuhSesi_();
  const merek = teks_(namaBrand);
  const brand = getBrands_().filter(function (b) { return b.name === merek; })[0];
  const kartu = integrasiList_().filter(function (r) { return teks_(r.Brand) === merek; });
  const dibuang = [];
  const gagal = [];
  kartu.forEach(function (r) {
    try {
      hapusIntegrasi(teks_(r.ID));
      dibuang.push(teks_(r['Nama Konfigurasi']));
    } catch (e) {
      gagal.push(teks_(r['Nama Konfigurasi']) + ': ' + pesanIzin_(e));
    }
  });
  if (brand) {
    hapusBrand(merek);
    batalCache_(merek);
  }
  BRAND_PUNYA_KONFIG = null;
  return {
    ok: true,
    brand: merek,
    adaBrand: !!brand,
    kartuDibuang: dibuang,
    gagal: gagal,
    tab: brand ? teks_(brand.sheetName) : '',
  };
}

/**
 * Membersihkan brand contoh: form CONTOH-nya dibuang ke Trash, tab jawabannya dihapus,
 * terus brandnya dilepas dari dashboard. Tab yang masih ada jawaban asli tidak pernah disentuh.
 */
function hapusFormDanTab(namaBrand) {
  butuhSesi_();
  if (CONTOH_BRAND.indexOf(namaBrand) === -1) {
    throw new Error('Tombol ini hanya untuk brand contoh (' + CONTOH_BRAND.join(', ') + ').');
  }
  const brand = findBrand_(namaBrand);
  const hasil = { brand: namaBrand, formDibuang: false, tabDihapus: '', peringatan: '' };

  let ss = null;
  try {
    ss = SpreadsheetApp.openById(brand.spreadsheetId);
  } catch (e) {
    throw new Error('Spreadsheet ' + namaBrand + ' tidak bisa dibuka: ' + e.message);
  }
  const sheet = ss.getSheetByName(brand.sheetName);
  if (sheet) {
    const nilai = sheet.getDataRange().getValues();
    const col = kolom_((nilai[0] || []).map(teks_));
    const barisAsli = nilai.slice(1).filter(function (r) {
      if (!r.some(function (v) { return String(v).trim() !== ''; })) return false;
      const email = col.email === -1 ? '' : teks_(r[col.email]).toLowerCase();
      return email.indexOf(EMAIL_CONTOH) === -1;
    }).length;
    if (barisAsli) {
      throw new Error('Tab "' + sheet.getName() + '" masih punya ' + barisAsli + ' baris jawaban asli, jadi tidak kuhapus. Kalau memang mau bersih-bersih, pindahkan dulu barisnya.');
    }
    const namaTab = sheet.getName();
    if (ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
      hasil.tabDihapus = namaTab;
    } else {
      // Spreadsheet tidak boleh tinggal tanpa tab, jadi tab-nya dikosongkan saja.
      sheet.clear();
      tulisHeader_(sheet, namaBrand);
      hasil.tabDihapus = namaTab + ' (dikosongkan, bukan dihapus, karena ini tab satu-satunya)';
    }
  }

  if (brand.formId) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === FUNGSI_NOTIF && String(t.getTriggerSourceId()) === String(brand.formId)) ScriptApp.deleteTrigger(t);
    });
    try {
      DriveApp.getFileById(brand.formId).setTrashed(true);
      hasil.formDibuang = true;
    } catch (e) {
      hasil.peringatan = 'Form-nya belum bisa dibuang otomatis (' + e.message + ') — hapus manual di Drive.';
    }
  }

  hapusBrand(namaBrand);
  batalCache_(namaBrand);
  return hasil;
}

/**
 * Membersihkan SEMUA brand contoh sekaligus: barisnya, tab jawabannya, form CONTOH-nya
 * (masuk Trash Drive, masih bisa dipulihkan 30 hari), kartunya di Integrasi, dan brandnya
 * dari sidebar. Brand kantor tidak pernah disentuh, dan brand contoh yang di tab-nya masih
 * ada jawaban asli ditolak hapus oleh hapusFormDanTab — itu disengaja.
 * Tanpa konfirmasi, fungsi ini cuma membuat daftar yang akan dibuang.
 */
function bersihkanContoh(konfirmasi) {
  butuhSesi_();
  const daftar = getBrands_().filter(function (b) { return CONTOH_BRAND.indexOf(b.name) !== -1; });
  const namaKartu = {};
  integrasiList_().forEach(function (r) {
    if (CONTOH_BRAND.indexOf(teks_(r.Brand)) === -1) return;
    (namaKartu[teks_(r.Brand)] = namaKartu[teks_(r.Brand)] || []).push(teks_(r['Nama Konfigurasi']));
  });
  const rencana = daftar.map(function (b) {
    return { brand: b.name, tab: teks_(b.sheetName), kartu: namaKartu[b.name] || [], punyaForm: !!(b.formId || b.formUrl) };
  });
  Object.keys(namaKartu).forEach(function (merek) {
    if (!daftar.some(function (b) { return b.name === merek; })) {
      rencana.push({ brand: merek, tab: '(tidak terdaftar sebagai brand)', kartu: namaKartu[merek], punyaForm: false });
    }
  });
  if (!konfirmasi) {
    return { ok: true, perluKonfirmasi: true, rencana: rencana, jumlah: rencana.length };
  }

  const dibersihkan = [];
  const gagal = [];
  rencana.forEach(function (r) {
    const buangKartu = function () {
      integrasiList_().filter(function (k) { return teks_(k.Brand) === r.brand; }).forEach(function (k) {
        try { hapusIntegrasi(k.ID); } catch (e) { /* kartunya memang sudah hilang */ }
      });
    };
    try {
      let h = { tabDihapus: '', formDibuang: false, peringatan: '' };
      if (getBrands_().some(function (b) { return b.name === r.brand; })) h = hapusFormDanTab(r.brand);
      buangKartu();
      dibersihkan.push({ brand: r.brand, tab: h.tabDihapus, formDibuang: !!h.formDibuang, peringatan: h.peringatan || '' });
    } catch (e) {
      buangKartu();
      gagal.push(r.brand + ': ' + pesanIzin_(e));
    }
  });
  batalCache_();
  BRAND_PUNYA_KONFIG = null;
  return { ok: true, dibersihkan: dibersihkan, gagal: gagal };
}

/**
 * Tab buatan form singkat tidak punya kolom kampus/prodi/NIM, jadi nilai itu hilang
 * waktu data contoh ditulis ke situ. Untuk baris contoh, nilainya diambil lagi dari
 * daftar CONTOH_ORANG lewat email — cuma tampilan, spreadsheet tidak ditulisi.
 */
var PETA_CONTOH_BY_EMAIL_ = null;

function contohByEmail_(email) {
  const key = teks_(email).toLowerCase();
  if (!key || key.indexOf(EMAIL_CONTOH) === -1) return null;
  if (!PETA_CONTOH_BY_EMAIL_) {
    PETA_CONTOH_BY_EMAIL_ = {};
    CONTOH_ORANG.forEach(function (o) { PETA_CONTOH_BY_EMAIL_[teks_(o.email).toLowerCase()] = o; });
  }
  return PETA_CONTOH_BY_EMAIL_[key] || null;
}

function nilaiKolom_(row, col, contoh, kunci) {
  const nilai = teks_(isi_(row, col[kunci]));
  if (nilai) return nilai;
  return contoh ? teks_(contoh[kunci]) : '';
}

/**
 * Mengabari satu pemagang lewat email. Penerima selalu diambil dari baris spreadsheet,
 * bukan dari yang dikirim browser, jadi aplikasi ini tidak bisa dipakai kirim ke sembarang alamat.
 */

/**
 * ---------- Konfigurasi integrasi Google Form + Spreadsheet respons ----------
 * Daftar sumber data disimpan di tab 'Integrasi' pada FILE LOG yang sudah ada:
 * beda file dari spreadsheet jawaban pendaftar, dan tidak menambah file baru ke Drive.
 */
const INTEGRASI_SHEET = 'Integrasi';
const INTEGRASI_COLS = [
  'ID', 'Nama Konfigurasi', 'Brand', 'Program', 'URL Google Form', 'Form ID', 'URL Edit Form',
  'URL Spreadsheet', 'Spreadsheet ID', 'Nama Sheet', 'Status', 'Status Koneksi',
  'Dibuat', 'Diubah', 'Terakhir Diuji', 'Terakhir Terhubung', 'Error Terakhir', 'Deskripsi',
];
const KOLOM_PENTING = ['nama', 'email', 'hp', 'kampus', 'prodi', 'minat', 'mulai'];

function sheetIntegrasi_(bikin) {
  const ss = getLogSs_();
  let sheet = ss.getSheetByName(INTEGRASI_SHEET);
  if (!sheet && bikin) {
    sheet = ss.insertSheet(INTEGRASI_SHEET);
    sheet.getRange(1, 1, 1, INTEGRASI_COLS.length).setValues([INTEGRASI_COLS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Dibaca tanpa bikin tab baru (jalur baca tidak boleh menulis); yang menulis pakai integrasiSimpan_.
 * Hasilnya disimpan sekali per eksekusi: satu halaman pernah membaca tab ini 4-5 kali dan itu
 * yang bikin tiap klik terasa lama. Salinannya yang diberikan, jadi pemanggil boleh mengubah datanya.
 */
var CACHE_INTEGRASI_ = null;
function integrasiList_() {
  if (CACHE_INTEGRASI_ === null) {
    const sheet = sheetIntegrasi_(false);
    if (!sheet) {
      CACHE_INTEGRASI_ = [];
    } else {
      const nilai = sheet.getDataRange().getValues();
      const header = (nilai[0] || []).map(teks_);
      CACHE_INTEGRASI_ = nilai.slice(1).map(function (row) {
        const rec = {};
        header.forEach(function (h, i) { rec[h] = teks_(row[i]); });
        return rec;
      }).filter(function (r) { return r.ID; });
    }
  }
  return CACHE_INTEGRASI_.map(function (r) { return Object.assign({}, r); });
}

/**
 * Membaca tab Integrasi lewat jalur yang tidak boleh menulis sama sekali: pointer file log
 * tidak dibuang, file log baru tidak dibuat, header tidak diperlebar. Dipakai jalur boot
 * (daftar brand) — di sana satu file yang tidak bisa dibuka tidak boleh mematikan halaman.
 * Kalau gagal, alasannya dikembalikan supaya bisa ditandai di layar.
 */
function integrasiBacaSafing_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_LOG);
  if (!id) return { daftar: [], galat: '' };
  let ss = null;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    return { daftar: [], galat: pesanIzin_(e) };
  }
  const sheet = ss.getSheetByName(INTEGRASI_SHEET);
  if (!sheet) return { daftar: [], galat: '' };
  const nilai = sheet.getDataRange().getValues();
  const header = (nilai[0] || []).map(teks_);
  const daftar = nilai.slice(1).map(function (row) {
    const rec = {};
    header.forEach(function (h, i) { rec[h] = teks_(row[i]); });
    return rec;
  }).filter(function (r) { return r.ID; });
  return { daftar: daftar, galat: '' };
}

function integrasiCari_(id) {
  return integrasiList_().filter(function (r) { return r.ID === teks_(id); })[0] || null;
}

/** Header tab Integrasi diperlebar dulu kalau ada kolom baru, supaya tab lama tidak ditolak Google. */
function headerIntegrasi_(sheet) {
  const ada = ((sheet.getDataRange().getValues() || [[]])[0] || []).map(teks_).filter(function (h) { return h; });
  const kurang = INTEGRASI_COLS.filter(function (h) { return ada.indexOf(h) === -1; });
  if (!kurang.length) return ada;
  const gabungan = ada.concat(kurang);
  const lebar = Math.max(sheet.getLastColumn(), 1);
  if (gabungan.length > lebar) sheet.insertColumnsAfter(lebar, gabungan.length - lebar);
  sheet.getRange(1, 1, 1, gabungan.length).setValues([gabungan]).setFontWeight('bold');
  return gabungan;
}

function integrasiSimpan_(rekaman) {
  batalCacheKonfigurasi_(); // daftar brand berubah, jangan sampai cache lama dipakai
  const sheet = sheetIntegrasi_(true);
  const header = headerIntegrasi_(sheet);
  const baris = header.map(function (h) {
    const nilai = rekaman[h];
    return nilai === undefined || nilai === null ? '' : String(nilai);
  });
  const nilai = sheet.getDataRange().getValues();
  for (let i = 1; i < nilai.length; i++) {
    if (teks_(nilai[i][0]) === rekaman.ID) {
      sheet.getRange(i + 1, 1, 1, baris.length).setValues([baris]);
      return rekaman;
    }
  }
  sheet.appendRow(baris);
  return rekaman;
}

function idIntegrasiBaru_() {
  return 'INT-' + Date.now().toString(36).toUpperCase() + '-' + Math.floor(Math.random() * 1296).toString(36).toUpperCase();
}

function ambilIdForm_(idAtauUrl) {
  const s = String(idAtauUrl || '').trim();
  const cocok = s.match(/\/forms\/d\/([^/?#]+)/) || s.match(/\/forms\/u\/\d+\/([^/?#]+)/) || s.match(/[?&]id=([^&?#]+)/);
  return cocok ? cocok[1] : '';
}

/**
 * Nama tab yang akan dipakai kalau penggunanya memilih "(otomatis)".
 * Pakai aturan yang sama dengan uji koneksi, supaya tab yang dibandingkan benar-benar tab yang dibaca.
 */
var TAB_OTOMATIS_CACHE_ = {};
function namaTabOtomatis_(idSs) {
  const id = teks_(idSs);
  if (!id) return '';
  if (TAB_OTOMATIS_CACHE_[id] !== undefined) return TAB_OTOMATIS_CACHE_[id];
  let nama = '';
  try {
    const ss = SpreadsheetApp.openById(id);
    const sheet = cariSheetForm_(ss) || ss.getSheets()[0];
    nama = sheet ? sheet.getName() : '';
  } catch (e) { nama = ''; }
  TAB_OTOMATIS_CACHE_[id] = nama;
  return nama;
}

/** Nama tab milik satu konfigurasi, termasuk yang dulu disimpan sebagai "(otomatis)". */
function tabMilik(rekaman) {
  return teks_(rekaman['Nama Sheet']) || namaTabOtomatis_(teks_(rekaman['Spreadsheet ID']));
}

/** Uji koneksi sungguhan: buka spreadsheet, pastikan sheet-nya ada, lalu baca judul kolomnya. */
function ujiKoneksiIntegrasi_(spreadsheetId, namaSheet) {
  const ss = SpreadsheetApp.openById(String(spreadsheetId).trim());
  const sheet = namaSheet ? ss.getSheetByName(String(namaSheet)) : (cariSheetForm_(ss) || ss.getSheets()[0]);
  if (!sheet) throw new Error('Sheet "' + namaSheet + '" tidak ada di spreadsheet "' + ss.getName() + '".');
  const info = bacaHeaderInfo_(sheet);
  const header = info.header.map(teks_);
  if (!header.length) {
    throw new Error('Tab "' + sheet.getName() + '" tidak dikenali baris judulnya (dicari sampai baris ke-20). Judul kolomnya harus memuat kata seperti Timestamp, Email, atau Nama Lengkap.');
  }
  const col = kolom_(header);
  return {
    spreadsheetNama: ss.getName(),
    sheetNama: sheet.getName(),
    kolom: header,
    judulBaris: info.judul + 1,
    dikenali: KOLOM_PENTING.filter(function (k) { return col[k] !== -1; }),
    hilang: KOLOM_PENTING.filter(function (k) { return col[k] === -1; }),
    baris: Math.max(sheet.getLastRow() - (info.judul + 1), 0),
  };
}

/** Sheet yang boleh dibaca dashboard — tab konfigurasi Integrasi dan tab tersembunyi tidak ikut ditawarkan. */
function daftarSheetTerbaca_(ss) {
  return ss.getSheets().filter(function (s) {
    return s.getName() !== INTEGRASI_SHEET && s.isSheetHidden && !s.isSheetHidden();
  }).map(function (s) {
    return { nama: s.getName(), baris: Math.max(s.getLastRow() - 1, 0) };
  });
}

/** Daftar sheet di sebuah spreadsheet, buat isian "Nama sheet respons" di form. Link form juga diterima. */
function daftarSheetIntegrasi(urlSpreadsheet) {
  butuhSesi_();
  const teksUrl = teks_(urlSpreadsheet);
  if (/\/forms\//i.test(teksUrl) && !/\/spreadsheets\//i.test(teksUrl)) {
    const dariForm = sumberDariForm_(teksUrl);
    const ss = SpreadsheetApp.openById(dariForm.idSs);
    return {
      ok: true,
      spreadsheet: dariForm.judulForm + ' → ' + ss.getName(),
      spreadsheetId: dariForm.idSs,
      dariForm: true,
      saranSheet: dariForm.namaSheet,
      sheets: daftarSheetTerbaca_(ss),
    };
  }
  const idSs = ambilIdSpreadsheet_(urlSpreadsheet);
  if (!idSs) throw new Error('Isi link spreadsheet atau link form-nya dulu.');
  const ss = SpreadsheetApp.openById(idSs);
  return {
    ok: true,
    spreadsheet: ss.getName(),
    spreadsheetId: ss.getId(),
    sheets: daftarSheetTerbaca_(ss),
  };
}

/**
 * Cukup tempel link Google Form: spreadsheet + tab jawabannya dibaca dari form-nya sendiri.
 * Ini yang dipakai waktu admin belum tahu link spreadsheet-nya.
 */
function sumberDariForm_(urlForm) {
  const idForm = ambilIdForm_(urlForm);
  if (!idForm) throw new Error('Link Google Form-nya tidak dikenali. Contoh: https://docs.google.com/forms/d/ID/edit');
  let form = null;
  try {
    form = FormApp.openById(idForm);
  } catch (e) {
    const izin = pesanIzin_(e);
    if (izin !== String((e && e.message) || e)) throw new Error(izin);
    throw new Error('Form itu tidak bisa dibuka aplikasi. Kalau form punya orang lain, bagikan dulu ke akun ' +
      (akunAplikasi_() || 'tempat aplikasi ini jalan') + ' (atau tempel link spreadsheet jawabannya saja).');
  }
  const tujuan = sheetTanggapanForm_(form);
  if (!tujuan) {
    throw new Error('Form "' + safeGet_(form, 'getTitle') + '" belum menuliskan jawabannya ke spreadsheet mana pun. Buka form → titik tiga → Pilih tujuan tanggapan → pilih spreadsheet.');
  }
  return {
    idForm: idForm,
    idSs: tujuan.ss.getId(),
    urlSs: tujuan.ss.getUrl(),
    namaSheet: tujuan.sheet.getName(),
    judulForm: safeGet_(form, 'getTitle'),
  };
}

function validasiIntegrasi_(data) {
  const nama = teks_(data.nama);
  const brand = teks_(data.brand);
  if (!nama) throw new Error('Nama konfigurasi wajib diisi.');
  if (!brand) throw new Error('Brand wajib diisi.');
  const urlFormAsli = teks_(data.urlForm);
  const urlSsUser = teks_(data.urlSpreadsheet);
  let dariForm = null;
  let peringatanForm = '';
  if (urlFormAsli) {
    try {
      dariForm = sumberDariForm_(urlFormAsli);
    } catch (e) {
      // Tanpa link spreadsheet kita benar-benar buta; kalau link spreadsheet ada, lanjut saja.
      if (!urlSsUser) throw e;
      peringatanForm = String(e.message || e);
    }
  }
  const urlSs = (dariForm && dariForm.urlSs) || urlSsUser;
  if (!urlSs) throw new Error('Tempel link Google Spreadsheet respons, atau link Google Form-nya saja (nanti spreadsheet + tabnya dicari otomatis).');
  const idSs = ambilIdSpreadsheet_(urlSs);
  if (!/^[A-Za-z0-9_-]{10,}$/.test(idSs)) {
    throw new Error('Link spreadsheet respons tidak dikenali. Tempel link lengkap, contoh: https://docs.google.com/spreadsheets/d/ID/edit');
  }
  if (urlFormAsli && !ambilIdForm_(urlFormAsli)) {
    throw new Error('Link Google Form-nya tidak dikenali. Contoh: https://docs.google.com/forms/d/ID/edit');
  }
  let peringatanPindah = '';
  if (data.pindah) {
    if (!dariForm) throw new Error('Menaruh jawaban form di tab sendiri cuma bisa kalau link Google Form-nya diisi dan bisa dibuka.');
    const pindah = pindahkanTujuanFormKe_(dariForm.idForm, idSs, teks_(data.id));
    dariForm.namaSheet = pindah.sheet;
    peringatanPindah = 'Jawaban form sekarang masuk ke tab "' + pindah.sheet + '"' + (pindah.peringatan ? ' — ' + pindah.peringatan : '');
  }
  const pilihanUser = data.pindah ? '' : teks_(data.namaSheet);
  const namaSheet = pilihanUser || (dariForm && dariForm.namaSheet) || '';
  // Pilihan "(otomatis: tab jawaban pertama)" ikut diterjemahkan jadi nama tab,
  // kalau tidak, dua konfigurasi yang sama persis bisa tersimpan terus tanpa ketahuan.
  const tabTerpilih = teks_(namaSheet) || namaTabOtomatis_(idSs);
  const peringatanTab = (pilihanUser && dariForm && pilihanUser !== dariForm.namaSheet)
    ? 'Tab "' + pilihanUser + '" kamu pilih manual, padahal form ini menulis jawabannya ke tab "' + dariForm.namaSheet + '".'
    : '';
  const dobel = tabTerpilih ? integrasiList_().filter(function (r) {
    return r['Spreadsheet ID'] === idSs && tabMilik(r) === tabTerpilih && r.ID !== teks_(data.id);
  })[0] : null;
  if (dobel && !data.paksa) {
    throw new Error('Tab "' + tabTerpilih + '" sudah dipakai konfigurasi "' + dobel['Nama Konfigurasi'] + '". ' +
      (dariForm
        ? 'Tekan "Pindahkan jawaban form" kalau kamu mau form ini punya tab sendiri, atau "Tetap buat" kalau memang mau baca tab yang sama.'
        : 'Tekan "Tetap buat" kalau kamu memang mau dua konfigurasi membaca tab yang sama.'));
  }
  const idFormSebenarnya = dariForm ? dariForm.idForm : ambilIdForm_(urlFormAsli);
  const urlEditForm = teks_(data.urlEdit) ||
    (idFormSebenarnya ? 'https://docs.google.com/forms/d/' + idFormSebenarnya + '/edit' : '');

  return {
    nama: nama, brand: brand, idSs: idSs, urlSs: urlSs, urlForm: urlFormAsli,
    idForm: idFormSebenarnya, urlEdit: urlEditForm, namaSheet: namaSheet,
    tabForm: (dariForm && dariForm.namaSheet) || '',
    peringatanForm: peringatanForm, peringatanTab: peringatanTab,
    peringatanPindah: peringatanPindah, dobel: !!dobel,
  };
}

/** Brand dari konfigurasi yang terbaca langsung didaftarkan ke sidebar, tidak perlu Sinkronkan dulu. */
function daftarkanBrand_(rekaman, uji) {
  const merek = teks_(rekaman.Brand);
  if (!merek) return { brand: '', baru: false };
  try {
    const ada = getBrands_().some(function (b) { return b.name === merek; });
    sambungkanBrand(merek, teks_(rekaman['Spreadsheet ID']), (uji && uji.sheetNama) || teks_(rekaman['Nama Sheet']));
    batalCache_(merek);
    return { brand: merek, baru: !ada };
  } catch (e) {
    return { brand: '', baru: false };
  }
}

function tambahIntegrasi(data) {
  butuhSesi_();
  const sah = validasiIntegrasi_(data || {});
  const sekarang = tgl_(new Date(), true);
  const rekaman = {
    ID: idIntegrasiBaru_(),
    'Nama Konfigurasi': sah.nama,
    Brand: sah.brand,
    Program: teks_(data.program),
    'URL Google Form': sah.urlForm,
    'Form ID': sah.idForm,
    'URL Edit Form': sah.urlEdit,
    'URL Spreadsheet': sah.urlSs,
    'Spreadsheet ID': sah.idSs,
    'Nama Sheet': sah.namaSheet,
    Status: teks_(data.status) === 'Nonaktif' ? 'Nonaktif' : 'Aktif',
    'Status Koneksi': 'Belum Diuji',
    Dibuat: sekarang,
    Diubah: sekarang,
    Deskripsi: teks_(data.deskripsi),
  };
  const hasil = jalankanUjiIntegrasi_(rekaman);
  const merek = hasil.error ? { brand: '', baru: false } : daftarkanBrand_(rekaman, hasil.uji);
  return {
    ok: true, id: rekaman.ID, integrasi: hasil.rekaman, uji: hasil.uji, error: hasil.error,
    peringatanTab: sah.peringatanTab, peringatanPindah: sah.peringatanPindah,
    brand: merek.brand, brandBaru: merek.baru,
  };
}

function ubahIntegrasi(id, data) {
  butuhSesi_();
  const lama = integrasiCari_(id);
  if (!lama) throw new Error('Konfigurasi ' + id + ' tidak ditemukan. Muat ulang halaman Integrasi.');
  const sah = validasiIntegrasi_(Object.assign({}, data, { id: id }));
  const rekaman = Object.assign({}, lama, {
    'Nama Konfigurasi': sah.nama,
    Brand: sah.brand,
    Program: teks_(data.program),
    'URL Google Form': sah.urlForm,
    'Form ID': sah.idForm,
    'URL Edit Form': sah.urlEdit,
    'URL Spreadsheet': sah.urlSs,
    'Spreadsheet ID': sah.idSs,
    'Nama Sheet': sah.namaSheet,
    Deskripsi: teks_(data.deskripsi),
    Diubah: tgl_(new Date(), true),
  });
  if (teks_(data.status) === 'Nonaktif') rekaman.Status = 'Nonaktif';
  else if (teks_(data.status) === 'Aktif') rekaman.Status = 'Aktif';
  const hasil = jalankanUjiIntegrasi_(rekaman);
  const merek = hasil.error ? { brand: '', baru: false } : daftarkanBrand_(rekaman, hasil.uji);
  return {
    ok: true, id: id, integrasi: hasil.rekaman, uji: hasil.uji, error: hasil.error,
    peringatanTab: sah.peringatanTab, peringatanPindah: sah.peringatanPindah,
    brand: merek.brand, brandBaru: merek.baru,
  };
}

/** Uji ulang satu konfigurasi dan catat hasilnya apa adanya — termasuk kalau gagal. */
function ujiIntegrasi(id) {
  butuhSesi_();
  const rekaman = integrasiCari_(id);
  if (!rekaman) throw new Error('Konfigurasi ' + id + ' tidak ditemukan.');
  const hasil = jalankanUjiIntegrasi_(rekaman);
  return { ok: !hasil.error, id: id, integrasi: hasil.rekaman, uji: hasil.uji, error: hasil.error };
}

/**
 * Inti "pindahkan jawaban form": setel ulang tujuan tanggapan form ke spreadsheet lain.
 * Kita tidak bisa menyebut nama tab-nya — Forms membuat sendiri tab baru bernama judul form.
 * Jawaban lama tetap tinggal di tab sebelumnya, tidak ada yang terhapus.
 */
function pindahkanTujuanFormKe_(idForm, idSs, kecualiId) {
  let form = null;
  try {
    form = FormApp.openById(idForm);
  } catch (e) {
    throw new Error('Form itu tidak bisa dibuka aplikasi. Kalau pemiliknya akun lain, bagikan dulu form-nya ke akun ini.');
  }
  const ss = SpreadsheetApp.openById(idSs);
  const sebelum = ss.getSheets().map(function (s) { return s.getName(); });
  const tujuanLama = sheetTanggapanForm_(form);
  const sheetLama = (tujuanLama && tujuanLama.ss.getId() === ss.getId()) ? '' : (tujuanLama ? tujuanLama.sheet.getName() : '');

  // Tautan lama dilepas dulu: tanpa itu, menyetel tujuan ke spreadsheet yang sama cuma
  // memakai ulang tab jawabannya dan tab baru tidak pernah terbentuk.
  try { form.removeDestination(); } catch (e) { /* belum tertaut, tidak apa-apa */ }
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const judul = safeGet_(form, 'getTitle');
  const anakBaru = ss.getSheets().filter(function (s) { return sebelum.indexOf(s.getName()) === -1; });
  const sheet = (anakBaru.length ? anakBaru[0] : null)
    || ss.getSheetByName(judul)
    || ss.getSheetByName(judul + ' (Jawaban Formulir)')
    || null;
  if (!sheet) {
    throw new Error('Form-nya sudah ditautkan ke "' + ss.getName() + '", tapi tab jawabannya tidak ketemu. Buka spreadsheet-nya, lihat tab paling baru.');
  }
  const pemakai = integrasiList_().filter(function (r) {
    return r.ID !== teks_(kecualiId) && r['Spreadsheet ID'] === ss.getId() && teks_(r['Nama Sheet']) === sheet.getName();
  })[0];
  if (!anakBaru.length && pemakai) {
    throw new Error('Forms memakai lagi tab "' + sheet.getName() + '" yang sudah dipakai konfigurasi "' + pemakai['Nama Konfigurasi'] +
      '". Ganti dulu judul form-nya (misal tambah nama brand), lalu ulangi.');
  }
  const peringatan = [];
  if (sheetLama) peringatan.push('Jawaban yang sudah masuk sebelum ini tetap di tab "' + sheetLama + '".');
  if (!anakBaru.length) peringatan.push('Forms memakai tab "' + sheet.getName() + '" yang sudah ada.');
  return {
    sheet: sheet.getName(), spreadsheet: ss.getName(), urlSs: ss.getUrl(), idSs: ss.getId(),
    formJudul: judul, sheetLama: sheetLama, peringatan: peringatan.join(' '),
  };
}

/** Tombol di kartu konfigurasi: pindahkan jawaban form-nya ke spreadsheet konfigurasi ini. */
function pindahkanTujuanForm(id) {
  butuhSesi_();
  const rekaman = integrasiCari_(id);
  if (!rekaman) throw new Error('Konfigurasi ' + id + ' tidak ditemukan.');
  const idForm = teks_(rekaman['Form ID']) || ambilIdForm_(rekaman['URL Google Form']);
  if (!idForm) throw new Error('Konfigurasi ini belum punya link Google Form. Tekan Edit, tempel link form-nya, simpan dulu.');
  if (!teks_(rekaman['Spreadsheet ID'])) throw new Error('Konfigurasi ini belum menunjuk ke spreadsheet mana pun.');
  const pindah = pindahkanTujuanFormKe_(idForm, teks_(rekaman['Spreadsheet ID']), teks_(rekaman.ID));
  rekaman['Nama Sheet'] = pindah.sheet;
  rekaman['URL Spreadsheet'] = pindah.urlSs;
  rekaman['Spreadsheet ID'] = pindah.idSs;
  rekaman['URL Google Form'] = rekaman['URL Google Form'] || ('https://docs.google.com/forms/d/' + idForm + '/edit');
  rekaman['Form ID'] = idForm;
  rekaman.Diubah = tgl_(new Date(), true);
  const hasil = jalankanUjiIntegrasi_(rekaman);
  return {
    ok: !hasil.error, id: teks_(rekaman.ID), formJudul: pindah.formJudul,
    spreadsheet: pindah.spreadsheet, sheetBaru: pindah.sheet, sheetLama: pindah.sheetLama,
    peringatan: pindah.peringatan, uji: hasil.uji, error: hasil.error,
  };
}

function jalankanUjiIntegrasi_(rekaman) {
  const sekarang = tgl_(new Date(), true);
  rekaman['Terakhir Diuji'] = sekarang;
  let uji = null;
  let error = '';
  try {
    uji = ujiKoneksiIntegrasi_(rekaman['Spreadsheet ID'], rekaman['Nama Sheet']);
    rekaman['Status Koneksi'] = 'Terhubung';
    rekaman['Nama Sheet'] = uji.sheetNama;
    rekaman['Terakhir Terhubung'] = sekarang;
    rekaman['Error Terakhir'] = '';
  } catch (e) {
    error = pesanIzin_(e);
    rekaman['Status Koneksi'] = 'Gagal Terhubung';
    rekaman['Terakhir Terhubung'] = '';
    rekaman['Error Terakhir'] = error;
  }
  try {
    integrasiSimpan_(rekaman);
  } catch (e) {
    // Hasil uji tetap dilaporkan walau tab Integrasi sedang tidak bisa ditulis.
    Logger.log('Gagal menyimpan hasil uji ' + teks_(rekaman.ID) + ': ' + (e.message || e));
  }
  return { rekaman: rekaman, uji: uji, error: error };
}

function setelStatusIntegrasi(id, aktif) {
  butuhSesi_();
  const rekaman = integrasiCari_(id);
  if (!rekaman) throw new Error('Konfigurasi ' + id + ' tidak ditemukan.');
  rekaman.Status = aktif ? 'Aktif' : 'Nonaktif';
  rekaman.Diubah = tgl_(new Date(), true);
  integrasiSimpan_(rekaman);
  return { ok: true, status: rekaman.Status };
}

function kolomIntegrasi(id) {
  butuhSesi_();
  const rekaman = integrasiCari_(id);
  if (!rekaman) throw new Error('Konfigurasi ' + id + ' tidak ditemukan.');
  const uji = ujiKoneksiIntegrasi_(rekaman['Spreadsheet ID'], rekaman['Nama Sheet']);
  return { ok: true, sheet: uji.sheetNama, kolom: uji.kolom, dikenali: uji.dikenali, hilang: uji.hilang, baris: uji.baris };
}

/**
 * Mendaftarkan satu integrasi sebagai brand yang dibaca dashboard.
 * Kegagalan satu sumber tidak boleh menghentikan sumber lainnya, jadi pemanggilnya yang mengumpulkan error.
 */
function sinkronIntegrasi(id) {
  butuhSesi_();
  const rekaman = integrasiCari_(id);
  if (!rekaman) throw new Error('Konfigurasi ' + id + ' tidak ditemukan.');
  if (rekaman.Status !== 'Aktif') return { ok: false, id: id, dilewati: true, alasan: 'konfigurasinya Nonaktif' };
  const hasilUji = jalankanUjiIntegrasi_(rekaman);
  if (hasilUji.error) return { ok: false, id: id, merek: rekaman.Brand, alasan: hasilUji.error };
  sambungkanBrand(rekaman.Brand, rekaman['Spreadsheet ID'], hasilUji.uji.sheetNama);
  batalCache_(rekaman.Brand);
  return { ok: true, id: id, merek: rekaman.Brand, sheet: hasilUji.uji.sheetNama, baris: hasilUji.uji.baris };
}

function sinkronSemuaIntegrasi() {
  butuhSesi_();
  const daftar = integrasiList_();
  const berhasil = [];
  const gagal = [];
  let dilewati = 0;
  daftar.forEach(function (r) {
    if (r.Status !== 'Aktif') { dilewati++; return; }
    try {
      const h = sinkronIntegrasi(r.ID);
      if (h.ok) berhasil.push(h.merek + ': ' + h.baris + ' baris');
      else gagal.push(r['Nama Konfigurasi'] + ': ' + (h.alasan || 'gagal'));
    } catch (e) {
      gagal.push(r['Nama Konfigurasi'] + ': ' + String(e.message || e));
    }
  });
  return { ok: true, jumlah: daftar.length, berhasil: berhasil, gagal: gagal, dilewati: dilewati };
}

/**
 * Spreadsheet yang sudah tersambung — dipakai tombol "Ambil semua tab" kalau penggunanya
 * tidak menempel link baru. Konfigurasi lebih dipercaya daripada daftar brand, karena di
 * sanalah spreadsheet + tab jawabannya dicatat lengkap.
 */
function daftarSpreadsheetTerhubung_() {
  const ids = [];
  integrasiList_().forEach(function (r) {
    const id = teks_(r['Spreadsheet ID']);
    if (id && ids.indexOf(id) === -1) ids.push(id);
  });
  if (ids.length) return ids;
  getBrands_().forEach(function (b) {
    const id = teks_(b.spreadsheetId);
    if (id && ids.indexOf(id) === -1) ids.push(id);
  });
  return ids;
}

/** Tab ini sudah diambil belum? Dicek ke konfigurasi DAN ke brand supaya brand lama tidak pernah tertimpa. */
function tabSudahDiambil_(idSs, namaTab) {
  const konfig = integrasiList_().filter(function (r) {
    return teks_(r['Spreadsheet ID']) === idSs && tabMilik(r) === namaTab;
  })[0];
  if (konfig) return 'sudah jadi konfigurasi "' + teks_(konfig['Nama Konfigurasi']) + '"';
  const merek = getBrands_().filter(function (b) { return b.name === namaTab; })[0];
  if (!merek) return '';
  if (teks_(merek.spreadsheetId) === idSs && teks_(merek.sheetName) === namaTab) return 'sudah dibaca dashboard';
  return 'brand "' + namaTab + '" sudah menunjuk ke tab lain, jadi tidak diubah';
}

/**
 * Semua tab di spreadsheet yang terhubung didaftarkan sekaligus: 1 tab = 1 brand,
 * namanya sama dengan nama tab. Yang bukan data pendaftar (kosong, tanpa kolom Nama/Email,
 * brand lama yang menunjuk tab lain) dilewati dan alasannya ikut dikembalikan.
 */
function ambilSemuaTab(data) {
  butuhSesi_();
  const link = teks_(data && data.urlSpreadsheet);
  let idSs = '';
  if (link) {
    idSs = ambilIdSpreadsheet_(link);
    if (!/^[A-Za-z0-9_-]{10,}$/.test(idSs)) {
      throw new Error('Link spreadsheet-nya tidak dikenali. Tempel link lengkap, contoh: https://docs.google.com/spreadsheets/d/ID/edit');
    }
  }
  const sumber = idSs ? [idSs] : daftarSpreadsheetTerhubung_();
  if (!sumber.length) throw new Error('Belum ada spreadsheet yang terhubung. Tempel link spreadsheet-nya di kotak di atas, atau buat satu konfigurasi dulu.');

  const masuk = [];
  const dilewati = [];
  const gagal = [];
  const diperiksa = [];
  sumber.forEach(function (id) {
    let ss = null;
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      gagal.push(id + ': ' + pesanIzin_(e));
      return;
    }
    diperiksa.push(ss.getName());
    daftarSheetTerbaca_(ss).forEach(function (t) {
      const namaTab = teks_(t.nama);
      if (!namaTab) return;
      const sudah = tabSudahDiambil_(id, namaTab);
      if (sudah) { dilewati.push({ tab: namaTab, alasan: sudah }); return; }
      // Tab jawaban form lama (Form Responses 1..22) bukan brand — brandnya bernama orangnya, bukan nomornya.
      if (/^(form responses?|respons formulir)/i.test(namaTab)) {
        dilewati.push({ tab: namaTab, alasan: 'tab jawaban form lama, bukan brand' });
        return;
      }
      if (!t.baris) { dilewati.push({ tab: namaTab, alasan: 'tabnya masih kosong' }); return; }
      let uji = null;
      try {
        uji = ujiKoneksiIntegrasi_(id, namaTab);
      } catch (e) {
        const sebab = pesanIzin_(e);
        // Tab tanpa baris judul yang dikenali itu "bukan data pendaftar", bukan kegagalan koneksi.
        if (sebab.indexOf('tidak dikenali baris judulnya') !== -1) {
          dilewati.push({ tab: namaTab, alasan: 'baris judulnya tidak dikenali dashboard' });
          return;
        }
        gagal.push(namaTab + ': ' + sebab);
        return;
      }
      if (uji.dikenali.indexOf('nama') === -1 || uji.dikenali.indexOf('email') === -1) {
        dilewati.push({ tab: namaTab, alasan: 'kolom Nama/Email tidak dikenali di tab ini' });
        return;
      }
      try {
        const hasil = tambahIntegrasi({
          nama: namaTab + ' (tab otomatis)',
          brand: namaTab,
          urlSpreadsheet: ss.getUrl(),
          namaSheet: namaTab,
          status: 'Aktif',
          deskripsi: 'Didaftarkan otomatis dari spreadsheet "' + ss.getName() + '" pada ' + tgl_(new Date(), true),
        });
        if (hasil.error) { gagal.push(namaTab + ': ' + hasil.error); return; }
        if (!hasil.brand) { gagal.push(namaTab + ': konfigurasinya tersimpan, tapi brand-nya gagal didaftarkan'); return; }
        masuk.push({ tab: namaTab, brand: hasil.brand, baris: uji.baris, spreadsheet: ss.getName() });
      } catch (e) {
        gagal.push(namaTab + ': ' + pesanIzin_(e));
      }
    });
  });
  batalCache_();
  return { ok: true, diperiksa: diperiksa, masuk: masuk, dilewati: dilewati, gagal: gagal };
}

/** Daftar integrasi + apakah brand-nya sudah dipakai dashboard sekarang. */
function integrasiDenganStatus_() {
  const dibaca = integrasiBacaSafing_();
  KONFIG_GALAT_ = dibaca.galat;
  const merek = getBrands_().map(function (b) { return b.name; });
  dibaca.daftar.forEach(function (r) { r.dipakai = merek.indexOf(r.Brand) !== -1; });
  return dibaca.daftar;
}

function daftarIntegrasi() {
  butuhSesi_();
  let namaFile = '';
  try {
    namaFile = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(PROP_LOG)).getName();
  } catch (e) { namaFile = KONFIG_GALAT_ || 'belum kebaca'; }
  return { ok: true, integrasi: integrasiDenganStatus_(), fileKonfigurasi: namaFile, peringatan: KONFIG_GALAT_ };
}

/**
 * Menghapus satu konfigurasi dari tab Integrasi. Brand-nya ikut dilepas dari dashboard
 * kalau tidak ada konfigurasi lain yang memakainya (brand contoh tidak pernah dilepas).
 */
function hapusIntegrasi(id) {
  butuhSesi_();
  const rekaman = integrasiCari_(id);
  const sheet = sheetIntegrasi_(false);
  if (!sheet) throw new Error('Tab Integrasi belum ada di file log, jadi tidak ada yang bisa dihapus.');
  const nilai = sheet.getDataRange().getValues();
  let barisDihapus = false;
  for (let i = 1; i < nilai.length; i++) {
    if (teks_(nilai[i][0]) === teks_(id)) {
      sheet.deleteRow(i + 1);
      batalCacheKonfigurasi_();
      barisDihapus = true;
      break;
    }
  }
  // Baris tidak ketemu (ID kembar/rusak, atau tab dibaca dari file log yang salah):
  // jangan berhenti di situ — brand-nya tetap dilepas supaya dashboard tidak macet selamanya.
  const merek = teks_((rekaman || {}).Brand);
  const sisa = merek ? integrasiList_().filter(function (r) { return teks_(r.Brand) === merek; }) : [];
  let brandDilepas = false;
  let peringatan = '';
  if (merek && !sisa.length) {
    // Brand yang tidak dipakai konfigurasi mana pun selalu dilepas — dulu brand contoh
    // dikecualikan, akibatnya kartunya sudah dihapus tapi namanya tetap nongol di Brand.
    hapusBrand(merek);
    batalCache_(merek);
    brandDilepas = true;
  } else if (sisa.length) {
    peringatan = 'brand ' + merek + ' masih dipakai ' + sisa.length + ' konfigurasi lain (' +
      sisa.map(function (r) { return teks_(r['Nama Konfigurasi']); }).join(', ') + ') — hapus kartunya satu-satu supaya brandnya ikut hilang';
  }
  if (!barisDihapus && !rekaman) throw new Error('Konfigurasi ' + teks_(id) + ' tidak ditemukan di tab Integrasi. Tekan "Muat ulang" lalu coba hapus lagi.');
  return {
    ok: true,
    dihapus: (rekaman && rekaman['Nama Konfigurasi']) || teks_(id),
    merek: merek,
    brandDilepas: brandDilepas,
    sisa: sisa.length,
    peringatan: peringatan,
    peringatan: barisDihapus ? '' : 'Barisnya tidak ketemu di tab Integrasi, jadi brand-nya saja yang dilepas.',
  };
}

/**
 * Konfigurasi kembar hasil menekan "Simpan" berkali-kali: brand + spreadsheet + tab sama.
 * Tanpa konfirmasi, fungsi ini cuma membuat daftar yang akan dibuang.
 * Yang dipertahankan selalu yang paling tua (barisnya paling atas), sisanya dihapus.
 */
function bersihkanDuplikatIntegrasi(konfirmasi) {
  butuhSesi_();
  const kelompok = {};
  integrasiList_().forEach(function (r) {
    const kunci = [teks_(r.Brand).toLowerCase(), teks_(r['Spreadsheet ID']), tabMilik(r).toLowerCase()].join('|');
    (kelompok[kunci] = kelompok[kunci] || []).push(r);
  });
  const rencana = [];
  Object.keys(kelompok).forEach(function (k) {
    const isi = kelompok[k];
    if (isi.length < 2) return;
    const sisa = isi.slice(1);
    rencana.push({
      dipertahankan: { id: teks_(isi[0].ID), nama: teks_(isi[0]['Nama Konfigurasi']) },
      dibuang: sisa.map(function (r) { return { id: teks_(r.ID), nama: teks_(r['Nama Konfigurasi']), brand: teks_(r.Brand), tab: tabMilik(r) }; }),
    });
  });
  if (!konfirmasi) return { ok: true, perluKonfirmasi: true, rencana: rencana, jumlah: rencana.reduce(function (a, g) { return a + g.dibuang.length; }, 0) };

  const dibuang = [];
  const gagal = [];
  rencana.forEach(function (g) {
    g.dibuang.forEach(function (d) {
      try {
        const h = hapusIntegrasi(d.id);
        dibuang.push({ id: d.id, nama: d.nama, brand: d.brand, tab: d.tab, brandDilepas: !!h.brandDilepas });
      } catch (e) {
        gagal.push(d.nama + ': ' + (e.message || e));
      }
    });
  });
  batalCacheKonfigurasi_();
  return { ok: true, dibuang: dibuang, gagal: gagal, kelompok: rencana.length };
}

/**
 * Isi satu kartu yang diklik (Total / baru / Diterima / belum di-follow up / dst).
 * Disaring dan dipotong di server: brand sebesar 5.516 baris tidak perlu dikirim semua
 * hanya untuk menampilkan 12 nama di bawah kartu.
 */
function sasaranBaris(opsi) {
  butuhSesi_();
  const o = opsi || {};
  const merek = teks_(o.brand);
  const batas = Math.max(Math.min(Number(o.batas) || 12, 40), 1);
  if (merek) {
    const satu = sasarSatuBrand_(merek, o, batas);
    return { ok: true, brand: merek, rows: satu.rows, jumlah: satu.jumlah, dibatasi: batas };
  }
  // Brand dikosongkan = cari di SEMUA brand. Ini yang dipakai Dashboard gabungan, supaya
  // datanya tidak perlu semuanya masuk browser cuma untuk menampilkan 12 nama.
  const keluar = [];
  const dilewati = [];
  let jumlah = 0;
  getBrands_().forEach(function (b) {
    let satu = null;
    try { satu = sasarSatuBrand_(b.name, o, batas); } catch (e) { dilewati.push(b.name + ': ' + pesanIzin_(e)); return; }
    jumlah += satu.jumlah;
    keluar.push.apply(keluar, satu.rows);
  });
  keluar.sort(function (a, b) { return (b.iso || '') < (a.iso || '') ? -1 : 1; });
  return { ok: true, brand: '', rows: keluar.slice(0, batas), jumlah: jumlah, dibatasi: batas, dilewati: dilewati };
}

/** Menyaring satu brand dan mengembalikan secuil baris yang cocok (kolom yang dibaca cuma yang perlu). */
function sasarSatuBrand_(merek, o, batas) {
  const sheet = getBrandSheet_(merek);
  const info = bacaHeaderInfo_(sheet);
  if (info.judul === -1) throw new Error('Baris judul tab "' + sheet.getName() + '" belum dikenali dashboard.');
  const header = info.header.map(teks_);
  const col = kolom_(header);
  const jumlah = Math.max(sheet.getLastRow() - (info.judul + 1), 0);
  const saring = teks_(o.status);
  const mnt = teks_(o.minat);
  const kata = teks_(o.kata).toLowerCase();
  const perlu = ['nama', 'email', 'timestamp', 'Status'];
  if (mnt || kata) perlu.push('minat');
  if (saring === 'belum-followup') perlu.push('Follow-up Terakhir');
  if (saring === 'dijadwalkan') perlu.push('Jadwal Wawancara');
  if (saring === 'berkas-kurang') perlu.push('cv', 'portofolio');
  const kolomBaca = {};
  perlu.forEach(function (k) {
    const i = col[k];
    if (i === undefined || i < 0 || jumlah === 0) { kolomBaca[k] = null; return; }
    kolomBaca[k] = sheet.getRange(info.judul + 2, i + 1, jumlah, 1).getValues();
  });
  const sel = function (k, i) { return kolomBaca[k] ? teks_(kolomBaca[k][i][0]) : ''; };
  const keluar = [];
  let cocok = 0;

  for (let i = jumlah - 1; i >= 0; i--) {
    const n = sel('nama', i);
    const e = sel('email', i);
    if (!n && !e) continue;
    const st = normalisasiStatus_(sel('Status', i));
    const ts = kolomBaca.timestamp ? toDate_(kolomBaca.timestamp[i][0]) : null;
    if (saring === 'baru') { if (!(ts && masihBaru_(ts, st))) continue; }
    else if (saring === 'belum-followup') { if (!(st === 'Seleksi' && !sel('Follow-up Terakhir', i))) continue; }
    else if (saring === 'dijadwalkan') { if (!(st === 'Seleksi' && sel('Jadwal Wawancara', i))) continue; }
    else if (saring === 'berkas-kurang') {
      const ada = [sel('cv', i), sel('portofolio', i)].filter(function (v) { return /^https?:|\/forms\/d|drive\.google/i.test(v); }).length;
      const total = [sel('cv', i), sel('portofolio', i)].filter(function (v) { return v; }).length;
      if (!(STATUS_PROSES.indexOf(st) !== -1 && total && ada < total)) continue;
    } else if (saring === 'gugur') { if (st !== 'Ditolak' && st !== 'Dikeluarkan') continue; }
    else if (saring && saring !== 'all') { if (st !== saring) continue; }
    if (mnt && sel('minat', i) !== mnt) continue;
    if (kata && (n + ' ' + e + ' ' + sel('minat', i)).toLowerCase().indexOf(kata) === -1) continue;
    cocok++;
    if (keluar.length < batas) keluar.push({
      row: info.judul + 2 + i, brand: merek, nama: n, status: st, minat: sel('minat', i),
      timestamp: ts ? tgl_(ts) : '-', iso: ts ? ts.toISOString() : '', baru: !!(ts && masihBaru_(ts, st)),
    });
  }
  return { rows: keluar, jumlah: cocok };
}

function getBrandData(namaBrand) {
  butuhSesi_();
  const sheet = getBrandSheet_(namaBrand);
  const bagi = pisahTab_(nilaiSheet_(namaBrand, sheet));
  const col = kolom_(bagi.header);
  const logMap = getFollowUpMap_();
  const arsip = brandArsip_(namaBrand);

  const rows = [];
  bagi.rows.forEach(function (row, i) {
    const email = teks_(isi_(row, col.email));
    const nama = teks_(isi_(row, col.nama));
    if (!email && !nama) return;
    const ts = toDate_(isi_(row, col.timestamp));
    const statusBaris = normalisasiStatus_(teks_(isi_(row, col.Status)));
    const contoh = contohByEmail_(email);
    rows.push({
      row: i + bagi.mulai,
      brand: namaBrand,
      nama: nama,
      email: email,
      minat: nilaiKolom_(row, col, contoh, 'minat'),
      kampus: nilaiKolom_(row, col, contoh, 'kampus'),
      prodi: nilaiKolom_(row, col, contoh, 'prodi'),
      mulai: nilaiKolom_(row, col, contoh, 'mulai'),
      hp: nilaiKolom_(row, col, contoh, 'hp'),
      status: statusBaris,
      timestamp: ts ? tgl_(ts) : '-',
      baru: !arsip && masihBaru_(ts, statusBaris),
      followUpTerakhir: tgl_(toDate_(isi_(row, col['Follow-up Terakhir']))),
      jumlahFollowUp: (logMap[key_(namaBrand, email)] || []).length,
      linkSertifikat: teks_(isi_(row, col['Link Sertifikat'])),
      jadwal: teks_(isi_(row, col['Jadwal Wawancara'])),
      dokumenDiperiksa: teks_(isi_(row, col['Dokumen Diperiksa'])),
      berkas: jumlahBerkas_(row, col, contoh),
      riwayat: riwayatList_(teks_(isi_(row, col['Riwayat Status']))),
      iso: ts ? ts.toISOString() : '',
    });
  });

  rows.reverse();
  return { rows: rows, stats: buildStats_(rows), arsip: arsip };
}

function getDetail(namaBrand, rowNum) {
  butuhSesi_();
  const sheet = getBrandSheet_(namaBrand);
  const nilai = nilaiSheet_(namaBrand, sheet);
  const header = bacaHeader_(sheet);
  const col = kolom_(header);
  const row = nilai[rowNum - 1] || [];
  const email = teks_(isi_(row, col.email));
  const ts = toDate_(isi_(row, col.timestamp));

  const jawaban = RINGKAS.filter(function (k) { return k !== 'email' && k !== 'nama'; }).map(function (k) {
    return { label: LABEL_RINGKAS[k] || label_(header, col[k]), nilai: nilaiKolom_(row, col, contohByEmail_(email), k) };
  }).filter(function (j) { return j.nilai; });

  const dipakai = {};
  if (col.timestamp !== -1) dipakai[col.timestamp] = true;
  RINGKAS.forEach(function (k) { if (col[k] !== -1) dipakai[col[k]] = true; });
  ADMIN_COLS.forEach(function (nama) { if (col[nama] !== -1) dipakai[col[nama]] = true; });

  const lainnya = header.map(function (judul, i) {
    return { label: String(judul || '').replace(/\s+/g, ' ').trim(), nilai: teks_(row[i]), i: i };
  }).filter(function (j) { return !dipakai[j.i] && j.nilai; });

  // Durasi sering ditulis di jawaban lain ("3 bulan"), bukan di kolom tanggalnya.
  const teksCadangan = jawaban.concat(lainnya).map(function (j) { return teks_(j.nilai); }).join(' | ');
  // Sel tanggal yang asli bertipe tanggal (bukan teks) tidak boleh dipaksa jadi teks dulu.
  const mulaiMentah = row[col.mulai];
  const jadwalMagang = tanggalMagang_(
    mulaiMentah instanceof Date ? mulaiMentah : nilaiKolom_(row, col, contohByEmail_(email), 'mulai'),
    teksCadangan, ts);

  return {
    row: rowNum,
    brand: namaBrand,
    nama: teks_(isi_(row, col.nama)),
    email: email,
    status: normalisasiStatus_(teks_(isi_(row, col.Status))),
    timestamp: ts ? tgl_(ts) : '-',
    masuk: jadwalMagang.masuk,
    berakhir: jadwalMagang.berakhir,
    masukPerkiraan: jadwalMagang.perkiraan,
    masukDariDaftar: jadwalMagang.dariDaftar,
    jadwalAsli: jadwalMagang.teks,
    linkSertifikat: teks_(isi_(row, col['Link Sertifikat'])),
    jadwal: teks_(isi_(row, col['Jadwal Wawancara'])),
    dokumenDiperiksa: teks_(isi_(row, col['Dokumen Diperiksa'])),
    emailTerakhir: teks_(isi_(row, col['Email Terakhir'])),
    berkas: jumlahBerkas_(row, col, contohByEmail_(email)),
    berkasList: daftarBerkas_(row, col, header, contohByEmail_(email)),
    riwayat: riwayatList_(teks_(isi_(row, col['Riwayat Status']))),
    jawaban: jawaban,
    jawabanLain: lainnya,
    history: getFollowUpMap_()[key_(namaBrand, email)] || [],
    statusList: STATUS_LIST,
  };
}

function updateStatus(namaBrand, rowNum, status) {
  butuhSesi_();
  const sheet = getBrandSheet_(namaBrand, true);
  const col = kolom_(bacaHeader_(sheet));
  const lama = normalisasiStatus_(teks_(bacaSel_(sheet, rowNum, col.Status, 'Status')));
  // Lewat normalisasiStatus_, jadi nama lama ("Ongoing", "Berjalan", "selesai", "do") tidak lagi
  // jatuh diam-diam ke Seleksi — sheet lama tetap terbaca tanpa perlu ditulisi ulang.
  const value = normalisasiStatus_(status);
  setelSel_(sheet, rowNum, col.Status, value, 'Status');
  if (lama !== value) catatRiwayat_(sheet, rowNum, 'Status: ' + lama + ' -> ' + value);
  return { ok: true, status: value, waktu: tgl_(new Date()) };
}

function setJadwalWawancara(namaBrand, rowNum, tanggal, jam, lokasi, catatan) {
  butuhSesi_();
  if (!teks_(tanggal)) throw new Error('Tanggal wawancara belum dipilih.');
  const tglBaru = toDate_(tanggal);
  if (!tglBaru) throw new Error('Format tanggal wawancara tidak dikenali.');
  const bagian = [tgl_(tglBaru, true), teks_(jam), teks_(lokasi), teks_(catatan)].filter(function (v) { return v; });
  const sheet = getBrandSheet_(namaBrand, true);
  const col = kolom_(bacaHeader_(sheet));
  setelSel_(sheet, rowNum, col['Jadwal Wawancara'], bagian.join(' | '), 'Jadwal Wawancara');
  catatRiwayat_(sheet, rowNum, 'Jadwal wawancara: ' + bagian.join(' | '));
  return { ok: true, jadwal: bagian.join(' | ') };
}

function periksaDokumen(namaBrand, rowNum, nilai) {
  butuhSesi_();
  const value = nilai === 'Sudah diperiksa' ? 'Sudah diperiksa' : '';
  const sheet = getBrandSheet_(namaBrand, true);
  const col = kolom_(bacaHeader_(sheet));
  setelSel_(sheet, rowNum, col['Dokumen Diperiksa'], value || 'Belum periksa', 'Dokumen Diperiksa');
  catatRiwayat_(sheet, rowNum, 'Dokumen: ' + (value || 'dikembalikan ke belum periksa'));
  return { ok: true, nilai: value || 'Belum periksa' };
}

function tambahFollowUp(namaBrand, rowNum, catatan, kirimEmail) {
  butuhSesi_();
  const pesan = teks_(catatan);
  if (!pesan) throw new Error('Catatan follow-up belum diisi.');

  const sheet = getBrandSheet_(namaBrand, true);
  const header = bacaHeader_(sheet);
  const col = kolom_(header);
  const row = sheet.getRange(rowNum, 1, 1, header.length).getValues()[0];

  const email = teks_(isi_(row, col.email));
  if (kirimEmail && !email) throw new Error('Baris ini tidak punya email, jadi catatannya tidak bisa dikirim sebagai email.');

  const now = new Date();
  getLogSheet_().appendRow([now, namaBrand, email, teks_(isi_(row, col.nama)), pesan]);
  setelSel_(sheet, rowNum, col['Follow-up Terakhir'], now, 'Follow-up Terakhir');

  if (kirimEmail) {
    const subjek = ('Info pendaftaran magang ' + namaBrand).slice(0, 100);
    MailApp.sendEmail({ to: email, subject: subjek, body: pesan });
    setelSel_(sheet, rowNum, col['Email Terakhir'], tgl_(now, true) + ' - ' + subjek, 'Email Terakhir');
    catatRiwayat_(sheet, rowNum, 'Email dikirim ke ' + email);
  }

  return { ok: true, waktu: tgl_(now), terkirim: kirimEmail ? email : '' };
}

/**
 * Dua penanda tangan per brand — MagangJogja, FreelanceJogja, dan SkillUp Jogja beda orang.
 * Ganti teks "Nama ..." dengan nama aslinya; brand yang tidak didaftarkan pakai default.
 */
const PENANDA_TANGAN_PER_BRAND = {
  'MagangJogja': {
    kiri: { nama: 'Nama pembimbing MagangJogja', jabatan: 'Pembimbing Lapangan MagangJogja' },
    kanan: { nama: 'Nama koordinator MagangJogja', jabatan: 'Koordinator MagangJogja' },
  },
  'FreelanceJogja': {
    kiri: { nama: 'Nama pembimbing FreelanceJogja', jabatan: 'Pembimbing Lapangan FreelanceJogja' },
    kanan: { nama: 'Nama koordinator FreelanceJogja', jabatan: 'Koordinator FreelanceJogja' },
  },
  'SkillUp Jogja': {
    kiri: { nama: 'Nama pembimbing SkillUp Jogja', jabatan: 'Pembimbing Lapangan SkillUp Jogja' },
    kanan: { nama: 'Nama koordinator SkillUp Jogja', jabatan: 'Koordinator SkillUp Jogja' },
  },
};

const PENANDA_TANGAN_DEFAULT = {
  kiri: { nama: 'Nama pembimbing lapangan', jabatan: 'Pembimbing Lapangan' },
  kanan: { nama: 'Nama koordinator program', jabatan: 'Koordinator Program Magang' },
};

function penandaTanganBrand_(namaBrand) {
  return PENANDA_TANGAN_PER_BRAND[namaBrand] || PENANDA_TANGAN_DEFAULT;
}

const BULAN_INDONESIA = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

/**
 * Periode di sertifikat harus terbaca seperti masa magang, bukan seperti jawaban kuis form.
 * Kalau isian "rencana mulai" sudah berupa rentang tanggal (ada tahunnya) dipakai apa adanya;
 * kalau tidak (misal "Segera, durasi 6 bulan") dihitung dari tanggal daftar jadi "Tahun 2026/2027".
 */
function periodeSertifikat_(isian, tanggalDaftar) {
  const teks = teks_(isian);
  if (/19|20\d{2}/.test(teks) && /-|s\/d|sampai/i.test(teks)) return teks;
  const awal = toDate_(isian) || tanggalDaftar || new Date();
  const tahun = awal.getFullYear();
  const bulan = awal.getMonth();
  const akhir = bulan >= 6 ? tahun + 1 : tahun;
  return akhir === tahun ? 'Tahun ' + tahun : 'Tahun ' + tahun + '/' + akhir;
}

function tanggalIndonesia_(d) {
  return d.getDate() + ' ' + BULAN_INDONESIA[d.getMonth()] + ' ' + d.getFullYear();
}

const WARNA_BIRU = '#1e3a8a';
const WARNA_EMAS = '#b45309';
const WARNA_TUAS = '#0f172a';
const WARNA_ABU = '#475569';

/** Kotak teks rata tengah; ukuran font/warna lewat TextStyle (pola resmi di guides/slides/editing-styling). */
function teksKotak_(slide, tulisan, x, y, lebar, tinggi, gaya) {
  const kotak = slide.insertTextBox(tulisan, x, y, lebar, tinggi);
  const teks = kotak.getText();
  const gayaTeks = teks.getTextStyle();
  gayaTeks.setFontFamily(gaya.font || 'Georgia');
  if (gaya.ukuran) gayaTeks.setFontSize(gaya.ukuran);
  if (gaya.tebal) gayaTeks.setBold(true);
  if (gaya.miring) gayaTeks.setItalic(true);
  gayaTeks.setForegroundColor(gaya.warna || WARNA_TUAS);
  teks.getParagraphs().forEach(function (p) {
    p.getRange().getParagraphStyle().setParagraphAlignment(gaya.rata || SlidesApp.ParagraphAlignment.CENTER);
  });
  return kotak;
}

/** Persegi tanpa isi, cuma garis tepi — dipakai buat bingkai dan ornamen sudut. */
function kotakGaris_(slide, x, y, lebar, tinggi, warna, tebal) {
  const b = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, lebar, tinggi);
  b.getFill().setTransparent();
  const bordir = b.getBorder();
  bordir.setWeight(tebal);
  bordir.getLineFill().setSolidFill(warna);
  return b;
}

/** Garis lurus horizontal tipis (bawah judul, bawah nama, tempat tanda tangan). Hiasan, jadi tidak boleh bikin kirim gagal. */
function garisHorizontal_(slide, x, y, lebar, warna, tebal) {
  try {
    const b = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, lebar, tebal);
    b.getFill().setSolidFill(warna);
    b.getBorder().setTransparent();
    return b;
  } catch (e) {
    return null;
  }
}

/**
 * Sertifikat disusun di Google Slides supaya bisa dikasih bingkai, ornamen sudut, dan watermark,
 * lalu diekspor PDF. HTML tidak bisa: Drive hanya mengekspor PDF dari file buatan Google.
 */
function buatBlobSertifikat_(isi) {
  const pres = SlidesApp.create('Sertifikat ' + isi.nama);
  const halaman = pres.getSlides()[0];
  const W = pres.getPageWidth();
  const H = pres.getPageHeight();
  const sx = W / 720;
  const sy = H / 540;
  const X = function (v) { return v * sx; };
  const Y = function (v) { return v * sy; };

  // Bingkai, ornamen, dan watermark cuma hiasan: kalau Google menolaknya, teks tetap jalan.
  try {
    kotakGaris_(halaman, X(14), Y(14), W - X(28), H - Y(28), WARNA_EMAS, 1.25);
    kotakGaris_(halaman, X(26), Y(26), W - X(52), H - Y(52), WARNA_BIRU, 3);
    [[34, 34], [W - 34 * sx - 26, 34], [34, H - 34 * sy - 26], [W - 34 * sx - 26, H - 34 * sy - 26]].forEach(function (sudut) {
      kotakGaris_(halaman, X(sudut[0]), Y(sudut[1]), X(26), Y(26), WARNA_EMAS, 1);
    });
    teksKotak_(halaman, 'MAGANG', X(0), Y(190), W, Y(160), { ukuran: 96, warna: '#f1f5f9', tebal: true });
  } catch (e) {
    // tanpa bingkai pun sertifikat tetap sah dikirim
  }

  // Kop.
  teksKotak_(halaman, String(isi.brand || '').toUpperCase(), X(60), Y(52), W - X(120), Y(30), { ukuran: 17, warna: WARNA_BIRU, tebal: true });
  teksKotak_(halaman, 'PROGRAM MAGANG & PENGALAMAN KERJA', X(60), Y(80), W - X(120), Y(18), { ukuran: 9, warna: '#64748b', tebal: true });
  garisHorizontal_(halaman, X(90), Y(104), W - X(180), WARNA_EMAS, 1.5);

  // Judul + nomor.
  teksKotak_(halaman, 'SERTIFIKAT', X(60), Y(126), W - X(120), Y(52), { ukuran: 34, warna: WARNA_TUAS, tebal: true });
  teksKotak_(halaman, 'Nomor: ' + (isi.nomor || '-'), X(60), Y(182), W - X(120), Y(18), { ukuran: 10, warna: WARNA_ABU });

  // Nama peserta.
  teksKotak_(halaman, 'Diberikan kepada', X(60), Y(212), W - X(120), Y(18), { ukuran: 11, warna: WARNA_ABU, miring: true });
  teksKotak_(halaman, isi.nama, X(60), Y(232), W - X(120), Y(42), { ukuran: 27, warna: WARNA_BIRU, tebal: true });
  garisHorizontal_(halaman, (W - X(240)) / 2, Y(280), X(240), '#94a3b8', 1);

  // Isi.
  teksKotak_(halaman,
    'Telah menyelesaikan program magang pada bidang ' + (isi.posisi || '-')
    + ' selama periode ' + (isi.periode || 'sesuai ketentuan program')
    + ', dan dinyatakan selesai pada ' + isi.tanggal + '.',
    X(80), Y(296), W - X(160), Y(52), { ukuran: 12, warna: '#1f2937' });

  // Dua blok tanda tangan sejajar.
  const lebarBlok = X(250);
  const kiriX = (W / 2) - lebarBlok - X(14);
  const kananX = (W / 2) + X(14);
  teksKotak_(halaman, 'Yogyakarta, ' + isi.tanggal, kiriX, Y(372), lebarBlok, Y(18), { ukuran: 10, warna: WARNA_ABU });
  garisHorizontal_(halaman, kiriX + X(40), Y(430), lebarBlok - X(80), '#334155', 1);
  teksKotak_(halaman, isi.penandaTangan.kiri.nama, kiriX, Y(436), lebarBlok, Y(18), { ukuran: 11, warna: WARNA_TUAS, tebal: true });
  teksKotak_(halaman, isi.penandaTangan.kiri.jabatan, kiriX, Y(454), lebarBlok, Y(16), { ukuran: 9, warna: WARNA_ABU });

  garisHorizontal_(halaman, kananX + X(40), Y(430), lebarBlok - X(80), '#334155', 1);
  teksKotak_(halaman, isi.penandaTangan.kanan.nama, kananX, Y(436), lebarBlok, Y(18), { ukuran: 11, warna: WARNA_TUAS, tebal: true });
  teksKotak_(halaman, isi.penandaTangan.kanan.jabatan, kananX, Y(454), lebarBlok, Y(16), { ukuran: 9, warna: WARNA_ABU });

  teksKotak_(halaman, 'Dokumen ini diterbitkan secara otomatis oleh sistem pendaftaran magang.', X(60), Y(492), W - X(120), Y(16), { ukuran: 8, warna: '#94a3b8' });

  pres.saveAndClose();
  const filePres = DriveApp.getFileById(pres.getId());
  const blob = filePres.getAs(MimeType.PDF).setName('Sertifikat - ' + isi.nama + ' (' + isi.brand + ').pdf');
  filePres.setTrashed(true);
  return blob;
}

const ROMAWI_BULAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];

/** Nomor sertifikat: urutan data / singkatan brand / bulan romawi / tahun. */
function nomorSertifikat_(namaBrand, rowNum, tanggal) {
  const singkatan = String(namaBrand).replace(/[^A-Z]/g, '').slice(0, 4) || 'MAG';
  const urut = ('00' + Number(rowNum)).slice(-3);
  return urut + '/' + singkatan + '/' + ROMAWI_BULAN[tanggal.getMonth()] + '/' + tanggal.getFullYear();
}

/**
 * Membuat PDF sertifikat dari jawaban form, menyimpannya di Drive, lalu email ke pemagang.
 * Hanya untuk yang berstatus Lulus; link sertifikat ditulis ke sheet.
 */
function kirimSertifikat(namaBrand, rowNum) {
  butuhSesi_();
  const sheet = getBrandSheet_(namaBrand, true);
  const header = bacaHeader_(sheet);
  const col = kolom_(header);
  const row = sheet.getRange(rowNum, 1, 1, header.length).getValues()[0];

  const nama = teks_(isi_(row, col.nama));
  const email = teks_(isi_(row, col.email));
  const minat = teks_(isi_(row, col.minat)) || 'Program Magang';
  if (!nama || !email) throw new Error('Nama atau email belum lengkap di baris ' + rowNum + ', sheet ' + sheet.getName() + '.');
  const statusSekarang = normalisasiStatus_(teks_(isi_(row, col.Status)));
  if (statusSekarang !== 'Lulus') throw new Error('Sertifikat hanya untuk yang berstatus Lulus. Status ' + nama + ' masih ' + statusSekarang + '.');

  const sekarang = new Date();
  const blob = buatBlobSertifikat_({
    nama: nama,
    brand: namaBrand,
    posisi: minat,
    tanggal: tanggalIndonesia_(sekarang),
    nomor: nomorSertifikat_(namaBrand, rowNum, sekarang),
    periode: periodeSertifikat_(teks_(isi_(row, col.mulai)), toDate_(isi_(row, col.timestamp)) || sekarang),
    penandaTangan: penandaTanganBrand_(namaBrand),
  });
  const file = getFolder_().createFile(blob);

  const link = file.getUrl();
  MailApp.sendEmail({
    to: email,
    subject: 'Sertifikat ' + namaBrand + ' - ' + nama,
    body:
      'Halo ' + nama + ',\n\n' +
      'Terima kasih sudah mengikuti ' + namaBrand + ' pada bidang ' + minat + '. ' +
      'Sertifikat sudah kami lampirkan pada email ini.\n\n' +
      'Arsip sertifikat: ' + link + '\n\n' +
      'Salam,\nAdmin ' + namaBrand,
    attachments: [file.getBlob()],
  });

  setelSel_(sheet, rowNum, col['Link Sertifikat'], link, 'Link Sertifikat');
  setelSel_(sheet, rowNum, col.Status, 'Lulus', 'Status');
  catatRiwayat_(sheet, rowNum, 'Sertifikat dikirim ke ' + email);
  batalCache_(namaBrand);
  return { ok: true, link: link, email: email };
}

/**
 * Menulis satu sel dengan pesan yang bisa dimengerti.
 * Tanpa ini, kolom yang tidak ketemu (indeks -1) bikin Google membuang error Inggris "Invalid value at ...".
 */
function setelSel_(sheet, rowNum, kolomNolBasis, nilai, namaKolom) {
  if (kolomNolBasis === undefined || kolomNolBasis < 0) {
    throw new Error('Kolom "' + namaKolom + '" belum ada di tab "' + sheet.getName() + '". Buka spreadsheet, jalankan ulang sekali, kolomnya dibuat otomatis.');
  }
  if (!(rowNum >= 2 && rowNum <= sheet.getMaxRows())) {
    throw new Error('Baris ' + rowNum + ' sudah tidak ada di tab "' + sheet.getName() + '". Muat ulang dashboardnya ya.');
  }
  sheet.getRange(rowNum, kolomNolBasis + 1).setValue(nilai);
}

/** Versi bacanya, biar error "Cannot invoke method getValue() on null" tidak muncul mentah. */
function bacaSel_(sheet, rowNum, kolomNolBasis, namaKolom) {
  if (kolomNolBasis === undefined || kolomNolBasis < 0) {
    throw new Error('Kolom "' + namaKolom + '" belum ada di tab "' + sheet.getName() + '".');
  }
  if (!(rowNum >= 2 && rowNum <= sheet.getMaxRows())) {
    throw new Error('Baris ' + rowNum + ' sudah tidak ada di tab "' + sheet.getName() + '". Muat ulang dashboardnya ya.');
  }
  return sheet.getRange(rowNum, kolomNolBasis + 1).getValue();
}

/* ---------------- bagian dalam ---------------- */

const CONTOH_HEADER = [
  'Timestamp',
  'Nama Lengkap',
  'Tahun Lahir',
  'NIM (Nomor Induk Mahasiswa)',
  'Email',
  'Jenis Kelamin',
  'No. HP Aktif yang bisa dihubungi dan chat WA',
  'Asal Sekolah/Kampus',
  'Program Studi',
  'Fakultas',
  'Nama Kota/Daerah tempat tinggal Anda saat ini',
  'Mengapa Anda ingin Magang/PKL disini?',
  'Jenis Magang apa yang Anda pilih ?',
  'Saya Siap dg Sistem Magang ini :',
  'Status Anda saat ini :',
  'Apakah Anda bisa membaca buku berbahasa Inggris kah?',
  'No. HP Aktif / WA Dosen atau Guru Pembimbing PKL/Magang beserta Nama dan Jabatan',
  'Program Magang/PKL apa yang Anda minati?',
  'Software design yang Anda kuasai',
  'Software editing video yang Anda kuasai',
  'Bahasa Pemrograman yang sudah Anda kuasai',
  'Materi Digital Marketing yang ingin dipraktikkan',
  'Apakah Anda ada alat kerja sendiri (LAPTOP) yang bisa dipakai selama Magang/PKL?',
  'Jika YA, alat apa yang Anda miliki?',
  'Kapan Anda berencana mulai Magang/PKL? beserta durasi magang',
  'Anda tahu info magang ini darimana?',
  'Silakan upload : CV, Scan KTP/KTM, Portofolio disini. Format pdf.',
  'Silakan upload : Portofolio (gambar jpg atau link youtube)',
  'Kegiatan Anda saat ini selain magang/pkl ada kah?',
  'Apakah Anda membutuhkan info kost/kontrakakan dekat kantor?',
  'Apakah Anda sudah berkeluarga?',
  'No. HP Aktif (WA) Wali/ Ortu peserta magang',
  'Sosial Media (Instagram) yang dimiliki',
];

const CONTOH_ORANG = [
  {
    brand: 0, timestamp: '2026-10-19T09:15:00+07:00', nama: 'Rina Kartika', statusAdmin: 'Diterima', tahunLahir: '2004', nim: '2025101234',
    email: 'rina.kartika@sample.invalid', hp: '081234567890', kampus: 'Universitas Amikom Yogyakarta',
    prodi: 'S-1 Teknik Informatika', fakultas: 'Fakultas Ilmu Komputer', kota: 'Sleman',
    alasan: 'Ingin merasakan kerja nyata di tim produk.', jenisMagang: 'Magang Kampus / Reguler',
    sistem: 'WFO (Work From Office)', statusDiri: 'Masih sekolah/masih kuliah', inggris: 'Saya Bisa',
    pembimbing: 'belum dapat dosen pembimbing', minat: 'UI/UX', cv: 'https://drive.google.com/file/d/CONTOH-CV-Rina/view',
    portofolio: 'https://www.figma.com/file/CONTOH-rina', mulai: '2 November 2026 - 2 Februari 2027 (3 Bulan)',
    sumber: 'Instagram', kegiatan: 'Asisten lab desain', kost: 'YA', keluarga: 'TIDAK',
    ortu: '081122223333 (Budi Santoso)', ig: '@rina.kartika', laptop: 'YA ADA', alat: 'Laptop/ Netbook',
  },
  {
    brand: 0, timestamp: '2026-10-18T15:40:00+07:00', nama: 'Bayu Prasetyo', statusAdmin: 'Diterima', tahunLahir: '2003', nim: '2023100987',
    email: 'bayu.pras@sample.invalid', hp: '085612341234', kampus: 'Universitas Gadjah Mada',
    prodi: 'S-1 Statistika', fakultas: 'FMIPA', kota: 'Bantul',
    alasan: 'Mau belajar data di dunia kerja.', jenisMagang: 'Magang Mandiri (magang atas inisatif diri sendiri, bukan kewajiban dari kampus)',
    sistem: 'WFO (Work From Office)', statusDiri: 'Masih sekolah/masih kuliah', inggris: 'Saya Bisa',
    pembimbing: '081298765432 (Dr. Sari, dosen Statistika)', minat: 'Programmer (Front End/ Backend)',
    jadwal: '03/11/2026 | 10.00 | Kantor 1', riwayat: '19/10/2026 10.05 - Data diperiksa, lanjut ke tahap seleksi\n20/10/2026 09.00 - Jadwal wawancara: 03/11/2026 | 10.00 | Kantor 1',
    cv: 'https://drive.google.com/file/d/CONTOH-CV-Bayu/view', mulai: '1 November 2026 - 1 Februari 2027 (3 Bulan)',
    sumber: 'Glints', kegiatan: 'Freelance data entry', kost: 'TIDAK', keluarga: 'TIDAK',
    ortu: '081311112222 (Suto Prasetyo)', ig: '@bayupras', laptop: 'YA ADA', alat: 'Laptop/ Netbook',
  },
  {
    brand: 0, timestamp: '2026-10-11T08:05:00+07:00', nama: 'Sinta Larasati', tahunLahir: '2002', nim: '2021100455',
    email: 'sinta.larasati@sample.invalid', hp: '087855556666', kampus: 'Universitas Negeri Yogyakarta',
    prodi: 'S-1 Desain Komunikasi Visual', fakultas: 'Fakultas Bahasa dan Seni', kota: 'Kulon Progo',
    alasan: 'Ingin portofolio desain yang dipakai orang banyak.', jenisMagang: 'Magang Kampus / Reguler',
    sistem: 'WFO (Work From Office)', statusDiri: 'Sudah lulus sekolah/kuliah dan belum bekerja', inggris: 'Saya Kurang Bisa (bisa sedikit-sedikit)',
    pembimbing: 'belum dapat dosen pembimbing', minat: 'Disainer Grafis', statusAdmin: 'Lulus',
    dokumen: 'Sudah diperiksa', riwayat: '15/10/2026 13.20 - Status: Seleksi -> Lulus', cv: 'https://drive.google.com/file/d/CONTOH-CV-Sinta/view',
    portofolio: 'https://drive.google.com/file/d/CONTOH-PDF-Sinta/view', mulai: '1 Oktober 2026 - 1 Januari 2027 (3 Bulan)',
    sumber: 'Instagram', kegiatan: 'Freelance desain logo', kost: 'YA', keluarga: 'TIDAK',
    ortu: '085299998888 (Rahman)', ig: '@sinta.dsgn', laptop: 'YA ADA', alat: 'Laptop yang sudah terinstal Corel dan Photoshop',
  },
  {
    brand: 0, timestamp: '2026-09-30T19:22:00+07:00', nama: 'Dimas Ardian', tahunLahir: '2004', nim: '2022100777',
    email: 'dimas.ardian@sample.invalid', hp: '082111223344', kampus: 'Universitas Ahmad Dahlan',
    prodi: 'D-3 Manajemen Informatika', fakultas: 'Fakultas Ilmu Komputer', kota: 'Kota Yogyakarta',
    alasan: 'Cari pengalaman sebelum lulus.', jenisMagang: 'Magang Mandiri (magang atas inisatif diri sendiri, bukan kewajiban dari kampus)',
    sistem: 'WFO (Work From Office)', statusDiri: 'Masih sekolah/masih kuliah', inggris: 'Saya Tidak bisa',
    pembimbing: 'belum dapat dosen pembimbing', minat: 'CS', statusAdmin: 'Ditolak', cv: 'https://drive.google.com/file/d/CONTOH-CV-Dimas/view',
    mulai: '20 Oktober 2026 - 20 Januari 2027 (3 Bulan)', sumber: 'Website', kegiatan: 'Aktif di BEM',
    kost: 'TIDAK', keluarga: 'TIDAK', ortu: '083811112222 (Sri Wahyuni)', ig: '@dimas.ardian', laptop: 'TIDAK ADA', alat: '-',
  },
  {
    brand: 1, timestamp: '2026-10-20T07:30:00+07:00', nama: 'Ayu Puspita', tahunLahir: '2003', nim: '2023100101',
    email: 'ayupuspita@sample.invalid', hp: '089529002944', kampus: 'Universitas Atma Jaya Yogyakarta',
    prodi: 'S-1 Ilmu Komunikasi', fakultas: 'Fakultas Sosial dan Humaniora', kota: 'Sleman',
    alasan: 'Suka menulis dan ingin belajar content planning.', jenisMagang: 'Magang Kampus / Reguler',
    sistem: 'WFO (Work From Office)', statusDiri: 'Masih sekolah/masih kuliah', inggris: 'Saya Bisa',
    pembimbing: '081234445555 (Nur Hidayati, M.I.Kom)', minat: 'Content Writer',
    cv: 'https://drive.google.com/file/d/CONTOH-CV-Ayu/view', portofolio: 'https://ayupuspita.blogspot.com',
    mulai: '3 November 2026 - 3 Februari 2027 (3 Bulan)', sumber: 'Instagram', kegiatan: 'Pengurus pers mahasiswa',
    kost: 'YA', keluarga: 'TIDAK', ortu: '081322223333 (Slamet)', ig: '@ayu.tulis', laptop: 'YA ADA', alat: 'Laptop/ Netbook',
  },
  {
    brand: 1, timestamp: '2026-10-17T11:02:00+07:00', nama: 'Fajar Nugroho', tahunLahir: '2001', nim: '-',
    email: 'fajar.nugroho@sample.invalid', hp: '081234509876', kampus: 'Universitas Kristen Duta Wacana',
    prodi: 'S-1 Desain Grafis', fakultas: 'Fakultas Desain dan Multimedia', kota: 'Kota Yogyakarta',
    alasan: 'Ingin fokus di video production.', jenisMagang: 'Magang Mandiri (magang atas inisatif diri sendiri, bukan kewajiban dari kampus)',
    sistem: 'WFO (Work From Office)', statusDiri: 'Sudah lulus dan sedang bekerja', inggris: 'Saya Bisa',
    pembimbing: 'belum dapat dosen pembimbing', minat: 'Videographer', statusAdmin: 'Lulus',
    cv: 'https://drive.google.com/file/d/CONTOH-CV-Fajar/view', portofolio: 'https://www.youtube.com/@CONTOH-fajar',
    mulai: '1 November 2026 - 1 Januari 2027 (2 Bulan)', sumber: 'Youtube', kegiatan: 'Freelance videographer',
    kost: 'TIDAK', keluarga: 'YA', ortu: '081123456789 (Agus)', ig: '@fajar.nug', laptop: 'YA ADA',
    alat: 'Laptop yang sudah terinstal Adobe Premiere Pro/Final Cut Pro/Adobe After Effect',
  },
  {
    brand: 2, timestamp: '2026-10-20T06:12:00+07:00', nama: 'Hendra Wijaya', statusAdmin: 'Dikeluarkan', tahunLahir: '2005', nim: '2025100222',
    email: 'hendra.wijaya@sample.invalid', hp: '085711223344', kampus: 'STIM AKNT',
    prodi: 'D-3 Manajemen Pemasaran', fakultas: 'Fakultas Ekonomi', kota: 'Bantul',
    alasan: 'Pengen belajar iklan dan social media.', jenisMagang: 'Magang Kampus / Reguler',
    sistem: 'WFO (Work From Office)', statusDiri: 'Masih sekolah/masih kuliah', inggris: 'Saya Kurang Bisa (bisa sedikit-sedikit)',
    pembimbing: 'belum dapat dosen pembimbing', minat: 'Digital Marketing',
    cv: 'https://drive.google.com/file/d/CONTOH-CV-Hendra/view', mulai: '10 November 2026 - 10 Februari 2027 (3 Bulan)',
    sumber: 'Instagram', kegiatan: 'Jualan online', kost: 'YA', keluarga: 'TIDAK',
    ortu: '085799887766 (Wijoyo)', ig: '@hendra.wjy', laptop: 'YA ADA', alat: 'Laptop/ Netbook',
  },
  {
    brand: 2, timestamp: '2026-10-14T13:55:00+07:00', nama: 'Nadia Rahman', statusAdmin: 'Dikeluarkan', tahunLahir: '2003', nim: '2023100333',
    email: 'nadia.r@sample.invalid', hp: '081277788899', kampus: 'Universitas Islam Indonesia',
    prodi: 'S-1 Manajemen', fakultas: 'Fakultas Ekonomi dan Bisnis', kota: 'Sleman',
    alasan: 'Ingin pengalaman marketing dan sales lapangan.', jenisMagang: 'Magang Kampus / Reguler',
    sistem: 'WFO (Work From Office)', statusDiri: 'Masih sekolah/masih kuliah', inggris: 'Saya Bisa',
    pembimbing: '081255554444 (Rina Marlina, SE, MM)', minat: 'Marketing dan Sales',
    jadwal: '05/11/2026 | 13.30 | https://meet.google.com/CONTOH-nadia', dokumen: 'Sudah diperiksa',
    cv: 'https://drive.google.com/file/d/CONTOH-CV-Nadia/view', mulai: '1 November 2026 - 1 Februari 2027 (3 Bulan)',
    sumber: 'Website', kegiatan: 'Part time barista', kost: 'TIDAK', keluarga: 'TIDAK',
    ortu: '081377776666 (Rahman)', ig: '@nadia.r', laptop: 'TIDAK ADA', alat: '-',
  },
];

// Peserta contoh tambahan supaya dashboard tidak sepi. Semua nama, kampus, NIM dan nomor di bawah ini palsu.
const NAMA_DEPAN_CONTOH = ['Alya', 'Bima', 'Cakra', 'Dewi', 'Eko', 'Fauzi', 'Gita', 'Hafiz', 'Indah', 'Joko', 'Kirana', 'Luthfi', 'Maya', 'Nabil', 'Oktarina', 'Putra', 'Rani', 'Salsa', 'Teguh', 'Umi', 'Vino', 'Winda', 'Yusuf', 'Zahra', 'Bagas', 'Cindy', 'Reza', 'Nadia'];
const NAMA_BELAKANG_CONTOH = ['Saputra', 'Wijaya', 'Lestari', 'Pratama', 'Hidayat', 'Nugroho', 'Rahayu', 'Setiawan', 'Maulida', 'Firmansyah'];
const PRODI_CONTOH = ['D-3 Manajemen Informatika', 'S-1 Teknik Informatika', 'S-1 Sistem Informasi', 'D-3 Akuntansi', 'S-1 DKV', 'D-3 Administrasi Bisnis', 'S-1 Ilmu Komunikasi', 'D-4 Teknik Elektro'];
const FAKULTAS_CONTOH = ['Fakultas Ilmu Komputer', 'Fakultas Ekonomi dan Bisnis', 'Fakultas Seni dan Desain', 'Fakultas Ilmu Sosial', 'Fakultas Vokasi'];
const SUMBER_CONTOH = ['Instagram', 'Teman/Kenalan', 'Kampus', 'LinkedIn', 'Tiktok'];
const CONTOH_FOLLOWUP_TAMBAHAN = [];

function isoContoh_(d) {
  const dua = function (n) { return ('0' + n).slice(-2); };
  return d.getFullYear() + '-' + dua(d.getMonth() + 1) + '-' + dua(d.getDate()) + 'T' + dua(d.getHours()) + ':' + dua(d.getMinutes()) + ':00+07:00';
}

/**
 * Tiap brand punya "kepribadian" sendiri: jumlah orang, sebaran status, bidang favorit,
 * asal kampus, dan tanggal daftar beda-beda. Kalau semua brand pakai pola geser yang sama,
 * grafik antar brand kelihatan identik dan datanya terasa palsu.
 */
const PROFIL_CONTOH = [
  {
    // MagangJogja — program paling tua, sudah banyak yang selesai. Total 12 orang.
    total: 12,
    status: ['Seleksi', 'Diterima', 'Diterima', 'Diterima', 'Lulus', 'Lulus', 'Ditolak', 'Seleksi'],
    minat: ['UI/UX', 'Programmer (Front End/ Backend)', 'Data Analyst', 'UI/UX', 'Programmer (Front End/ Backend)', 'UI/UX', 'Videographer', 'Admin Operasional'],
    kampus: ['Universitas Amikom Yogyakarta', 'Universitas Negeri Yogyakarta', 'Universitas Amikom Yogyakarta', 'Universitas Gadjah Mada', 'Politeknik Negeri Jogja', 'Universitas Amikom Yogyakarta', 'Universitas Atma Jaya Yogyakarta', 'Universitas Negeri Yogyakarta'],
    bulan: 8, hariMulai: 6, jarak: 4, jedaFollow: 3,
  },
  {
    // FreelanceJogja — baru dibuka, mayoritas masih seleksi, isinya anak desain & media. Total 17 orang.
    total: 17,
    status: ['Seleksi', 'Seleksi', 'Seleksi', 'Seleksi', 'Seleksi', 'Diterima', 'Seleksi', 'Diterima', 'Diterima', 'Ditolak', 'Seleksi', 'Seleksi', 'Diterima', 'Lulus', 'Lulus'],
    minat: ['Disainer Grafis', 'Videographer', 'Digital Marketing', 'Disainer Grafis', 'Copywriter', 'Videographer', 'Social Media Strategist', 'Disainer Grafis', 'Content Writer', 'Digital Marketing', 'Videographer', 'Disainer Grafis', 'Copywriter', 'Social Media Strategist', 'Digital Marketing'],
    kampus: ['Universitas Kristen Duta Wacana', 'Universitas Negeri Yogyakarta', 'Universitas Gadjah Mada', 'Universitas Ahmad Dahlan', 'Universitas Kristen Duta Wacana', 'Universitas Negeri Yogyakarta', 'Universitas Amikom Yogyakarta', 'Universitas Gadjah Mada', 'Universitas Teknologi Yogyakarta', 'Universitas Kristen Duta Wacana', 'Universitas Negeri Yogyakarta', 'Universitas Kristen Duta Wacana', 'Universitas Ahmad Dahlan', 'Universitas Gadjah Mada', 'Universitas Teknologi Yogyakarta'],
    bulan: 8, hariMulai: 24, jarak: 3, jedaFollow: 4,
  },
  {
    // SkillUp Jogja — pelamar ramai tapi banyak yang gugur. Total 18 orang.
    total: 18,
    status: ['Seleksi', 'Seleksi', 'Ditolak', 'Ditolak', 'Ditolak', 'Diterima', 'Diterima', 'Diterima', 'Seleksi', 'Dikeluarkan', 'Seleksi', 'Diterima', 'Ditolak', 'Seleksi', 'Lulus', 'Dikeluarkan'],
    minat: ['Digital Marketing', 'CS', 'Admin Operasional', 'Digital Marketing', 'Social Media Strategist', 'CS', 'Digital Marketing', 'Admin Operasional', 'CS', 'Digital Marketing', 'Admin Operasional', 'Social Media Strategist', 'CS', 'Digital Marketing', 'Admin Operasional', 'Content Writer'],
    kampus: ['Universitas Islam Indonesia', 'STIM AKNT', 'Universitas Teknologi Yogyakarta', 'Universitas Islam Indonesia', 'Politeknik Negeri Jogja', 'Universitas Teknologi Yogyakarta', 'Universitas Islam Indonesia', 'Universitas Amikom Yogyakarta', 'STIM AKNT', 'STIM AKNT', 'Politeknik Negeri Jogja', 'Universitas Islam Indonesia', 'STIM AKNT', 'Universitas Teknologi Yogyakarta', 'Universitas Islam Indonesia', 'Politeknik Negeri Jogja'],
    bulan: 9, hariMulai: 2, jarak: 3, jedaFollow: 3,
  },
];

function pilih_(daftar, i) {
  return daftar[i % daftar.length];
}

function tambahContohOrang_() {
  PROFIL_CONTOH.forEach(function (profil, b) {
    const punya = CONTOH_ORANG.filter(function (o) { return o.brand === b; }).length;
    // Jumlah orang per brand ditentukan profil.total; yang ditulis tangan sudah dihitung lebih dulu.
    const jumlah = Math.max(profil.total - punya, 0);
    for (let i = 0; i < jumlah; i++) {
      const status = pilih_(profil.status, i);
      const urut = punya + i;
      const nama = pilih_(NAMA_DEPAN_CONTOH, b * 7 + i) + ' ' + pilih_(NAMA_BELAKANG_CONTOH, i + b * 3);
      const email = nama.split(' ')[0].toLowerCase() + '.' + nama.split(' ')[1].toLowerCase() + b + urut + '@sample.invalid';
      const tanggal = new Date(2026, profil.bulan, profil.hariMulai + i * profil.jarak, 8 + (i % 9), (i * 7) % 60);
      const orang = {
        brand: b, timestamp: isoContoh_(tanggal), nama: nama, statusAdmin: status,
        tahunLahir: String(2001 + (urut % 5)), nim: '0000' + (1000 + b * 100 + urut),
        email: email, hp: '0000000000' + (b * 30 + urut),
        kampus: pilih_(profil.kampus, i),
        prodi: PRODI_CONTOH[(urut + b) % PRODI_CONTOH.length],
        fakultas: FAKULTAS_CONTOH[(urut + b) % FAKULTAS_CONTOH.length],
        kota: ['Sleman', 'Bantul', 'Kota Yogyakarta', 'Kulon Progo', 'Gunungkidul'][urut % 5],
        alasan: 'Ingin menambah pengalaman kerja sebelum lulus dan menerapkan ilmu kampus.',
        jenisMagang: urut % 2 ? 'Magang Mandiri (magang atas inisatif diri sendiri, bukan kewajiban dari kampus)' : 'Magang Kampus / Reguler',
        sistem: 'WFO (Work From Office)', statusDiri: 'Masih sekolah/masih kuliah',
        inggris: urut % 3 ? 'Saya Bisa' : 'Saya Kurang Bisa (bisa sedikit-sedikit)',
        pembimbing: 'belum dapat dosen pembimbing',
        minat: pilih_(profil.minat, i),
        cv: urut % 4 ? 'https://drive.google.com/file/d/CONTOH-CV-' + b + urut + '/view' : '',
        portofolio: urut % 3 === 0 ? 'https://drive.google.com/file/d/CONTOH-PF-' + b + urut + '/view' : '',
        mulai: ['1 November 2026 - 1 Februari 2027 (3 Bulan)', '2 November 2026 - 2 Januari 2027 (2 Bulan)', '16 Oktober 2026 - 16 Desember 2026 (2 Bulan)'][urut % 3],
        sumber: SUMBER_CONTOH[urut % SUMBER_CONTOH.length],
        kegiatan: urut % 4 ? 'Organisasi kampus' : 'Freelance',
        kost: urut % 2 ? 'YA' : 'TIDAK', keluarga: 'TIDAK',
        ortu: '000000000' + (b * 30 + urut) + ' (wali contoh)', ig: '@contoh.' + b + urut,
        laptop: 'YA ADA', alat: 'Laptop/ Netbook',
      };
      if (status !== 'Seleksi') orang.riwayat = '20/10/2026 09.00 - Status: Seleksi -> ' + status;
      if (status === 'Lulus' || urut % 5 === 1) orang.dokumen = 'Sudah diperiksa';
      if (status === 'Diterima' || status === 'Lulus' || status === 'Dikeluarkan') orang.jadwal = '0' + (3 + urut % 4) + '/11/2026 | 10.00 | Kantor';
      CONTOH_ORANG.push(orang);
      // Sebagian saja yang sudah dikontak, dan jaraknya beda per brand.
      if ((i + b) % profil.jedaFollow === 1) {
        CONTOH_FOLLOWUP_TAMBAHAN.push({
          brand: CONTOH_BRAND[b], email: email, ts: isoContoh_(new Date(2026, 9, 19 + (i % 2), 10, 0)),
          catatan: i % 2 ? 'Ditelepon, yang bersangkutan masih menunggu kepastian kampus.' : 'Sudah WA, berkas susulan dikirim lewat email.',
        });
      }
    }
  });
}

tambahContohOrang_();

/**
 * Contoh "berkas belum lengkap": sengaja 2 orang di SkillUp Jogja, brand lain semuanya lengkap
 * supaya angkanya jelas waktu didemoin. Yang dikosongkan CV-nya, portofolionya tetap ada.
 */
const BERKAS_KURANG_CONTOH = { 'SkillUp Jogja': 2 };
(function rapikanBerkasContoh_() {
  const sisa = {};
  Object.keys(BERKAS_KURANG_CONTOH).forEach(function (b) { sisa[b] = BERKAS_KURANG_CONTOH[b]; });
  CONTOH_ORANG.forEach(function (o) {
    const slug = String(o.email).split('@')[0];
    o.cv = 'https://drive.google.com/file/d/CONTOH-CV-' + slug + '/view';
    o.portofolio = 'https://drive.google.com/file/d/CONTOH-PF-' + slug + '/view';
    const merek = CONTOH_BRAND[o.brand] || '';
    if (sisa[merek] > 0 && o.statusAdmin === 'Seleksi') {
      o.cv = '';
      sisa[merek]--;
    }
  });
})();

const CONTOH_FOLLOWUP = [
  { brand: 'MagangJogja', email: 'bayu.pras@sample.invalid', ts: '2026-10-19T10:00:00+07:00', catatan: 'Sudah WA untuk konfirmasi jadwal seleksi tahap 1.' },
  { brand: 'SkillUp Jogja', email: 'nadia.r@sample.invalid', ts: '2026-10-16T08:30:00+07:00', catatan: 'Minta transkrip nilai lewat email.' },
  { brand: 'SkillUp Jogja', email: 'nadia.r@sample.invalid', ts: '2026-10-18T15:10:00+07:00', catatan: 'Transkrip sudah diterima, lanjut ke interview.' },
];
CONTOH_FOLLOWUP.push.apply(CONTOH_FOLLOWUP, CONTOH_FOLLOWUP_TAMBAHAN);

function barisContoh_(header, o) {
  const row = header.map(function () { return ''; });
  // row[-1] tidak menyimpan apa pun dan tidak error, jadi setiap indeks dicek dulu.
  const isi = function (i, nilai) { if (i !== -1 && i !== undefined) row[i] = nilai; };
  Object.keys(o).forEach(function (k) {
    if (k === 'brand' || k === 'statusAdmin') return;
    if (NAMA_KOLOM_ADMIN[k]) { isi(header.indexOf(NAMA_KOLOM_ADMIN[k]), o[k]); return; }
    isi(cariHeader_(header, k), o[k]);
  });
  isi(header.indexOf('Status'), o.statusAdmin || 'Seleksi');
  CONTOH_FOLLOWUP.forEach(function (f) {
    if (f.email === o.email) isi(header.indexOf('Follow-up Terakhir'), f.ts);
  });
  return row;
}

/**
 * Daftar brand yang dipakai dashboard. Sumbernya satu: yang tersimpan di properti.
 * Baris di tab Integrasi TIDAK otomatis jadi brand — dia baru masuk daftar lewat
 * "Simpan & Hubungkan" atau tombol "Ambil semua tab", supaya brand tidak bertambah sendiri.
 */
function getBrands_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_BRANDS);
  if (raw) {
    try {
      const data = JSON.parse(raw);
      if (data.length) return data;
    } catch (e) { /* dianggap belum ada, dicoba pulihkan */ }
  }
  pulihkanOtomatis_();
  try {
    return JSON.parse(props.getProperty(PROP_BRANDS) || "[]");
  } catch (e) {
    return [];
  }
}

function setBrands_(list) {
  PropertiesService.getScriptProperties().setProperty(PROP_BRANDS, JSON.stringify(list));
}

function findBrand_(namaBrand) {
  const brand = getBrands_().filter(function (b) { return b.name === namaBrand; })[0];
  if (!brand) throw new Error('Brand "' + namaBrand + '" belum terdaftar. Jalankan buatDataContoh atau sambungkanBrand dulu.');
  return brand;
}

function bukaSpreadsheet_(brand) {
  try {
    return SpreadsheetApp.openById(brand.spreadsheetId);
  } catch (e) {
    throw new Error('Tidak bisa membuka spreadsheet data untuk ' + brand.name + '. Cek akses akun kamu.');
  }
}

/**
 * Tab lain di file yang sama yang isinya masih ada dan belum dipakai brand lain.
 * Dipakai buat pulih sendiri setelah "Ganti form" memindahkan brand ke tab kosong.
 */
function tabPenghuniLama_(ss, namaBrand, sheetAktif) {
  const dipakai = getBrands_().filter(function (b) { return b.name !== namaBrand; }).map(function (b) { return b.sheetName; });
  let terbaik = null;
  let terbanyak = 0;
  ss.getSheets().forEach(function (s) {
    if (!sheetAktif || s.getName() === sheetAktif.getName()) return;
    if (dipakai.indexOf(s.getName()) !== -1) return;
    const baris = Math.max(s.getLastRow() - 1, 0);
    if (baris > terbanyak) { terbaik = s; terbanyak = baris; }
  });
  return terbaik;
}

// Merek mana yang tab-nya diatur lewat menu Integrasi — dihitung sekali per eksekusi.
var BRAND_PUNYA_KONFIG = null;
var KONFIG_GALAT_ = ''; // kenapa daftar konfigurasi tidak terbaca — dipakai layar supaya tidak mendadak kosong

/** Tab Integrasi berubah isi: buang hasil hitung yang lama sebelum ada yang membacanya lagi. */
function batalCacheKonfigurasi_() {
  BRAND_PUNYA_KONFIG = null;
  CACHE_INTEGRASI_ = null;
}

function brandPunyaKonfigurasi_(namaBrand) {
  if (BRAND_PUNYA_KONFIG === null) {
    BRAND_PUNYA_KONFIG = {};
    integrasiBacaSafing_().daftar.forEach(function (r) {
      if (teks_(r.Status) !== 'Nonaktif' && teks_(r.Brand)) BRAND_PUNYA_KONFIG[teks_(r.Brand).toLowerCase()] = true;
    });
  }
  return !!BRAND_PUNYA_KONFIG[teks_(namaBrand).toLowerCase()];
}

function getBrandSheet_(namaBrand, bolehTulis) {
  const brand = findBrand_(namaBrand);
  const ss = bukaSpreadsheet_(brand);
  let sheet = ss.getSheetByName(brand.sheetName);
  if (!sheet) {
    sheet = cariSheetForm_(ss);
    if (sheet) simpanNamaSheet_(namaBrand, sheet.getName());
  }
  if (!sheet) throw new Error('Sheet "' + brand.sheetName + '" tidak ditemukan di ' + ss.getName() + '.');
  if (!punyaHeader_(nilaiSheet_(namaBrand, sheet))) {
    // Sheet yang tercatat ternyata kosong: coba tab jawaban form yang lain di file yang sama.
    const alternatif = cariSheetForm_(ss);
    if (alternatif && alternatif.getName() !== sheet.getName()) {
      sheet = alternatif;
      simpanNamaSheet_(namaBrand, sheet.getName());
      batalCache_(namaBrand);
    }
    if (!punyaHeader_(nilaiSheet_(namaBrand, sheet))) tulisHeader_(sheet, namaBrand);
  }
  // Hanya di jalur baca: jangan pernah menulis status ke tab yang bukan asal barisnya.
  // Brand yang tab-nya diatur lewat menu Integrasi TIDAK boleh dilompati — tab form baru memang
  // kosong sampai ada yang submit, dan lompatan ini dulu yang membuat brand baca tab brand lain.
  if (!bolehTulis && !brandPunyaKonfigurasi_(namaBrand) && Math.max(sheet.getLastRow() - 1, 0) === 0) {
    const lama = tabPenghuniLama_(ss, namaBrand, sheet);
    if (lama) {
      sheet = lama;
      simpanNamaSheet_(namaBrand, lama.getName());
      batalCache_(namaBrand);
    }
  }
  if (bolehTulis) {
    ensureAdminCols_(sheet);
    batalCache_(namaBrand);
  }
  return sheet;
}

/**
 * File log dengan tab FollowUp paling berisi. Kalau ada lebih dari satu file bernama sama
 * (sisa proyek lama, atau file baru yang terlanjur dibuat saat pointer tidak kebaca),
 * yang kosong tidak dipilih — di sanalah catatan follow-up dan tab Integrasi yang asli.
 */
function cariFileLogLama_() {
  try {
    const files = DriveApp.getFilesByName(NAMA_LOG);
    let terbaik = null;
    let terbanyak = -1;
    while (files.hasNext()) {
      let ss = null;
      try {
        ss = SpreadsheetApp.openById(files.next().getId());
      } catch (e) { continue; }
      const followUp = ss.getSheetByName(LOG_SHEET);
      const integrasi = ss.getSheetByName(INTEGRASI_SHEET);
      const jumlah = followUp ? Math.max(followUp.getLastRow() - 1, 0) : (integrasi ? 1 : 0);
      if (jumlah > terbanyak) { terbaik = ss; terbanyak = jumlah; }
    }
    return terbaik;
  } catch (e) {
    return null;
  }
}

function getLogSs_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_LOG);
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // Pointer lama jangan langsung dibuang: file yang asli dicari lagi lewat namanya.
      // Membuangnya berarti menganggap tab Integrasi + catatan follow-up ikut hilang.
      const lama = cariFileLogLama_();
      if (lama) {
        props.setProperty(PROP_LOG, lama.getId());
        return lama;
      }
      props.deleteProperty(PROP_LOG);
    }
  }
  const ketemu = cariFileLogLama_();
  const ss = ketemu || SpreadsheetApp.create(NAMA_LOG);
  props.setProperty(PROP_LOG, ss.getId());
  return ss;
}

function getLogSheet_() {
  const ss = getLogSs_();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.getSheetByName('Sheet1') || ss.insertSheet(LOG_SHEET);
    sheet.setName(LOG_SHEET);
    sheet.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS]);
    formatHeader_(sheet);
  }
  return sheet;
}

function getFollowUpMap_() {
  let values = [];
  try {
    values = getLogSheet_().getDataRange().getValues();
  } catch (e) {
    return {};
  }
  values.shift();
  const map = {};
  values.forEach(function (row) {
    const key = key_(teks_(row[1]), teks_(row[2]));
    if (!map[key]) map[key] = [];
    map[key].push({ waktu: tgl_(toDate_(row[0])), catatan: teks_(row[4]) });
  });
  return map;
}

const BULAN_SINGKAT = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
const LABEL_BERKAS = {
  cv: 'CV, Scan KTP/KTM, portofolio PDF',
  portofolio: 'Portofolio gambar atau video',
};

function hitungPer_(rows, kunci) {
  const map = {};
  rows.forEach(function (r) {
    const label = teks_(r[kunci]) || 'Tidak diisi';
    map[label] = (map[label] || 0) + 1;
  });
  return Object.keys(map).map(function (label) {
    return { label: label, nilai: map[label] };
  }).sort(function (a, b) { return b.nilai - a.nilai; }).slice(0, 6);
}

function hitungPerBulan_(rows) {
  const map = {};
  rows.forEach(function (r) {
    if (!r.iso) return;
    const k = r.iso.slice(0, 7);
    map[k] = (map[k] || 0) + 1;
  });
  return Object.keys(map).sort().map(function (k) {
    const bagian = k.split('-');
    return { label: BULAN_SINGKAT[Number(bagian[1]) - 1] + ' ' + bagian[0], nilai: map[k] };
  });
}

function trendMinggu_(rows) {
  const kini = Date.now();
  const minggu = [];
  for (let i = JUMLAH_MINGGU_TREND - 1; i >= 0; i--) {
    minggu.push({ label: tgl_(new Date(kini - i * 7 * 86400000)).slice(0, 5), nilai: 0 });
  }
  rows.forEach(function (r) {
    if (!r.iso) return;
    const lalu = (kini - new Date(r.iso).getTime()) / (7 * 86400000);
    const i = Math.floor(lalu);
    if (lalu >= 0 && i < JUMLAH_MINGGU_TREND) minggu[JUMLAH_MINGGU_TREND - 1 - i].nilai++;
  });
  return minggu;
}

function saringRentang_(rows, dari, sampai) {
  if (!dari && !sampai) return rows;
  return rows.filter(function (r) {
    return !r.iso || dalamRentang_(new Date(r.iso), dari, sampai);
  });
}

function dalamRentang_(date, dari, sampai) {
  if (dari && date.getTime() < new Date(dari).getTime()) return false;
  if (sampai) {
    const akhir = new Date(sampai);
    akhir.setHours(23, 59, 59, 999);
    if (date.getTime() > akhir.getTime()) return false;
  }
  return true;
}

function jumlahBerkas_(row, col, contoh) {
  const kunci = ['cv', 'portofolio'];
  const isinya = function (k) {
    const nilai = teks_(isi_(row, col[k]));
    return nilai || (contoh ? teks_(contoh[k]) : '');
  };
  // Tab form singkat tidak punya kolom berkas: untuk baris contoh jumlahnya tetap 2,
  // untuk jawaban form asli jangan mengarang — totalnya ikut kolom yang benar-benar ada.
  const total = contoh ? kunci.length : kunci.filter(function (k) { return col[k] !== -1; }).length;
  const ada = kunci.filter(function (k) { return !!isinya(k); }).length;
  return { ada: ada, total: total };
}

function daftarBerkas_(row, col, header, contoh) {
  const out = [];
  ['cv', 'portofolio'].forEach(function (k) {
    const nilai = teks_(isi_(row, col[k])) || (contoh ? teks_(contoh[k]) : '');
    if (!nilai) return;
    const bagian = nilai.split(/[\n,]\s*/).filter(function (u) { return u; });
    bagian.forEach(function (url, i) {
      out.push({
        label: LABEL_BERKAS[k] + (bagian.length > 1 ? ' ' + (i + 1) : ''),
        url: url,
        judul: col[k] === -1 ? LABEL_BERKAS[k] : String(header[col[k]] || '').replace(/\s+/g, ' ').trim(),
      });
    });
  });
  return out;
}

function riwayatList_(isi) {
  return isi ? isi.split('\n').reverse() : [];
}

function catatRiwayat_(sheet, rowNum, pesan) {
  const col = kolom_(bacaHeader_(sheet));
  const lama = teks_(bacaSel_(sheet, rowNum, col['Riwayat Status'], 'Riwayat Status'));
  setelSel_(sheet, rowNum, col['Riwayat Status'], (lama ? lama + '\n' : '') + tgl_(new Date()) + ' - ' + pesan, 'Riwayat Status');
}

function getFolderEkspor_() {
  const it = DriveApp.getFoldersByName(FOLDER_EKSPOR);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_EKSPOR);
}

function buildStats_(rows) {
  const stats = { total: rows.length, baru: 0 };
  STATUS_LIST.forEach(function (s) { stats[s] = 0; });
  rows.forEach(function (r) {
    if (r.baru) stats.baru++;
    stats[r.status] = (stats[r.status] || 0) + 1;
  });
  return stats;
}

function ensureAdminCols_(sheet) {
  const info = bacaHeaderInfo_(sheet);
  if (info.judul === -1) return; // belum ketemu baris judulnya: jangan menulis di sembarang baris
  const header = info.header.map(teks_);
  const kurang = ADMIN_COLS.filter(function (nama) { return header.indexOf(nama) === -1; });
  // Kolom baru di ujung tab yang grid-nya sudah mentok DITOLAK Google (pelajaran dari tab Integrasi),
  // jadi grid-nya dilebarkan dulu sebelum ada yang menulis.
  const lebar = Math.max(sheet.getMaxColumns(), 1);
  const kebutuhan = header.length + kurang.length;
  if (kebutuhan > lebar) sheet.insertColumnsAfter(lebar, kebutuhan - lebar);
  let berubah = false;
  kurang.forEach(function (nama) {
    sheet.getRange(info.judul + 1, header.length + 1).setValue(nama);
    header.push(nama);
    berubah = true;
  });
  if (berubah) formatHeader_(sheet, info.judul + 1);
}

function formatHeader_(sheet, barisJudul) {
  const judul = barisJudul || 1;
  const last = Math.max(sheet.getLastColumn(), 1);
  sheet.getRange(judul, 1, 1, last).setFontWeight('bold').setBackground('#eef2ff').setVerticalAlignment('middle');
  sheet.setFrozenRows(judul);
}

function kolom_(header) {
  const map = {};
  Object.keys(ALIAS).forEach(function (k) { map[k] = cariHeader_(header, k); });
  ADMIN_COLS.forEach(function (nama) { map[nama] = header.indexOf(nama); });
  return map;
}

function cariHeader_(header, kunci) {
  const pola = ALIAS[kunci] || (kunci === 'status' ? 'Status' : kunci === 'followUpTerakhir' ? ADMIN_COLS[1] : null);
  if (!pola) return -1;
  if (typeof pola === 'string') return header.indexOf(pola);
  return header.map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); })
    .findIndex(function (h) { return pola.test(h); });
}

function label_(header, index) {
  return index === -1 ? '' : String(header[index] || '');
}

function isi_(row, index) {
  return index === -1 ? '' : row[index];
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_SERTIFIKAT);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_SERTIFIKAT);
}

function key_(brand, email) {
  return (brand || '').toLowerCase() + '|' + (email || '').toLowerCase();
}

function teks_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function toDate_(value) {
  if (value instanceof Date) return value;
  if (!value) return null;
  const parsed = new Date(String(value));
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Membaca tanggal dari jawaban form "Kapan Anda berencana mulai Magang/PKL? beserta durasi magang".
 * Contoh yang benar-benar ada: "16 Oktober 2026 - 16 Desember 2026 (2 Bulan)" dan
 * "2 November 2026 - 2 Februari 2027 (3 Bulan)". Kalau yang tertulis cuma satu tanggal
 * plus durasinya, tanggal berakhir dihitung dari durasi itu.
 * Tidak nebak: kalau bentuknya tidak dikenali, masuk/berakhir dibiarkan kosong dan
 * teks aslinya yang ditampilkan.
 */
const POLA_TANGGAL_ID = /(\d{1,2})\s*([A-Za-zĀ-ž]+)\.?\s*(\d{4})/;
// Orang menulis tanggal dengan angka juga: 12/10/2026, 12-10-2026, 12.10.2026, 2026-10-12.
// Urutan angka dipakai dd/bulan/tahun (cara Indonesia); yang yyyy-lebih-dulu dikenali terpisah.
const POLA_TANGGAL_NUM = /(^|[^\d])(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/;
const POLA_TANGGAL_ISO = /(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})/;

function tanggalNomor_(hari, bulan, tahun) {
  hari = Number(hari); bulan = Number(bulan); tahun = Number(tahun);
  if (tahun < 100) tahun += 2000;
  if (hari < 1 || hari > 31 || bulan < 1 || bulan > 12) return null;
  return new Date(tahun, bulan - 1, hari);
}

/** Tanggal pertama yang muncul di sebuah teks, apa pun bentuk penulisannya. */
function tanggalDariTeks_(teks) {
  const asli = String(teks || '');

  const iso = POLA_TANGGAL_ISO.exec(asli);
  if (iso) {
    const t = tanggalNomor_(iso[3], iso[2], iso[1]);
    if (t) return { tanggal: t, teks: iso[0], akhir: iso.index + iso[0].length };
  }

  const nama = POLA_TANGGAL_ID.exec(asli);
  const angka = POLA_TANGGAL_NUM.exec(asli);

  // Kalau dua-duanya ada, yang lebih dulu ditulis yang dipakai.
  if (nama && angka && angka.index < nama.index) {
    const t = tanggalNomor_(angka[2], angka[3], angka[4]);
    if (t) return { tanggal: t, teks: angka[0], akhir: angka.index + angka[0].length };
  }
  if (!nama) {
    if (!angka) return null;
    const t = tanggalNomor_(angka[2], angka[3], angka[4]);
    return t ? { tanggal: t, teks: angka[0], akhir: angka.index + angka[0].length } : null;
  }

  const hari = Number(nama[1]);
  const huruf = nama[2].charAt(0).toUpperCase() + nama[2].slice(1).toLowerCase();
  let bulan = BULAN_INDONESIA.indexOf(huruf);
  // Orang menulis "Okt", "Des", "Agu" juga: cocokkan dari tiga huruf pertama, tidak nebak yang lebih pendek.
  if (bulan === -1 && huruf.length >= 3) {
    const kecil = huruf.toLowerCase();
    bulan = BULAN_INDONESIA.findIndex(function (b) { return b.toLowerCase().indexOf(kecil) === 0; });
  }
  if (bulan === -1 || hari < 1 || hari > 31) return null;
  const tanggal = new Date(Number(nama[3]), bulan, hari);
  return { tanggal: tanggal, teks: nama[0], akhir: nama.index + nama[0].length };
}

/** Tanggal panjang berbahasa Indonesia: "12 Oktober 2026", bukan "12 October 2026". */
function tglIndonesia_(tanggal) {
  const pad = function (n) { return n < 10 ? '0' + n : String(n); };
  return pad(tanggal.getDate()) + ' ' + BULAN_INDONESIA[tanggal.getMonth()] + ' ' + tanggal.getFullYear();
}

/**
 * Berapa bulan durasi magang yang tertulis di sebuah teks jawaban.
 * Yang dikenali: "3 Bulan", "3 bln", "3-bulan", "1 tahun", "tiga bulan", "selama 6 bulan".
 * Angka di luar 1-36 atau kata yang tidak dikenal = tidak ditebak, hasilnya kosong.
 */
const ANGKA_KATA_ = { satu: 1, dua: 2, tiga: 3, empat: 4, lima: 5, enam: 6, tujuh: 7, delapan: 8, sembilan: 9, sepuluh: 10, sebelas: 11, 'dua belas': 12 };
function durasiBulan_(teks) {
  const asli = String(teks || '').toLowerCase().replace(/\s+/g, ' ');
  const angka = /(\d{1,2})\s*(?:-\s*)?(?:bulan|bln)/.exec(asli);
  if (angka) return Math.min(Math.max(Number(angka[1]), 1), 36);
  const kata = /(dua belas|satu|dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|sepuluh|sebelas)\s*(?:-\s*)?bulan/.exec(asli);
  if (kata) return ANGKA_KATA_[kata[1]];
  const tahun = /(\d{1,2})\s*(?:tahun|thn)/.exec(asli);
  if (tahun) return Math.min(Math.max(Number(tahun[1]) * 12, 1), 36);
  return 0;
}

/**
 * Tanggal masuk & berakhir dari jawaban form. Kalau kolomnya sendiri cuma berisi tanggal
 * mulai, durasi dicari di jawaban form yang lain (banyak form memisah "mulai" dan "durasi").
 * Tiga bentuk lain yang benar-benar ada di data kantor ikut ditangani: selnya berupa
 * tanggal betulan (bukan teks), penulisannya tanpa spasi ("8desember2026-30desember2026"),
 * dan jawabannya "Segera, durasi 6 bulan" — untuk yang terakhir ini tanggal daftar yang
 * jadi patokan dan itu ditandai supaya tidak dikira tertulis di form.
 */
function tanggalMagang_(jawabanMulai, jawabanLain, tanggalDaftar) {
  const tanggalAsli = jawabanMulai instanceof Date && !isNaN(jawabanMulai.getTime()) ? jawabanMulai : null;
  const teks = teks_(jawabanMulai);
  const hasil = {
    masuk: '', berakhir: '', perkiraan: false, dariLain: false, dariDaftar: false,
    teks: tanggalAsli ? tglIndonesia_(tanggalAsli) : teks,
  };

  let mula = tanggalAsli;
  let sisa = '';
  const pertama = tanggalAsli ? null : tanggalDariTeks_(teks);
  if (pertama) {
    mula = pertama.tanggal;
    sisa = teks.slice(pertama.akhir);
  } else if (!mula && tanggalDaftar instanceof Date && /segera|secepatnya|asap/i.test(teks)) {
    mula = tanggalDaftar;
    hasil.dariDaftar = true;
  }
  if (!mula) return hasil;

  hasil.masuk = tglIndonesia_(mula);
  const kedua = tanggalDariTeks_(sisa);
  if (kedua) {
    hasil.berakhir = tglIndonesia_(kedua.tanggal);
    return hasil;
  }
  let jumlah = durasiBulan_(teks);
  if (!jumlah && jawabanLain) {
    jumlah = durasiBulan_(jawabanLain);
    hasil.dariLain = jumlah > 0;
  }
  if (jumlah) {
    const akhir = new Date(mula.getFullYear(), mula.getMonth() + jumlah, mula.getDate());
    akhir.setDate(akhir.getDate() - 1);
    hasil.berakhir = tglIndonesia_(akhir);
    hasil.perkiraan = true;
  }
  return hasil;
}

/**
 * "Pendaftar baru" = mendaftar dalam N hari terakhir DAN masih berada di tahap Seleksi.
 * Orang yang sudah Diterima/Lulus/Ditolak/Dikeluarkan bukan "baru" lagi,
 * walaupun tanggal daftarmya masih muda.
 */
/** Yang belum ada keputusan dianggap pendaftar baru, tidak lagi dibatasi umur pendaftaran. */
function masihBaru_(ts, status) {
  return normalisasiStatus_(status) === 'Seleksi';
}

/**
 * Brand arsip = datanya tetap dibaca dan ikut dihitung di Laporan, tapi orangnya tidak
 * dianggap "pendaftar baru" / "perlu tindak lanjut". Dipakai untuk tab induk kantor yang
 * statusnya ditandai warna sel, bukan tulisan, jadi semua orang terbaca Seleksi.
 */
function brandArsip_(namaBrand) {
  const brand = getBrands_().filter(function (b) { return b.name === namaBrand; })[0];
  return !!(brand && brand.arsip);
}

function setelArsip(namaBrand, arsip) {
  butuhSesi_();
  const daftar = getBrands_();
  let ketemu = false;
  daftar.forEach(function (b) {
    if (b.name !== namaBrand) return;
    b.arsip = !!arsip;
    ketemu = true;
  });
  if (!ketemu) throw new Error('Brand "' + namaBrand + '" belum terdaftar.');
  setBrands_(daftar);
  batalCache_(namaBrand);
  return { ok: true, brand: namaBrand, arsip: !!arsip };
}

function normalisasiStatus_(value) {
  const kecil = String(value || '').trim().toLowerCase();
  const found = STATUS_LIST.filter(function (s) { return s.toLowerCase() === kecil; })[0];
  if (found) return found;
  return STATUS_LAMA[kecil] || 'Seleksi';
}

function tgl_(date, panjang) {
  if (!date) return '-';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), panjang ? 'dd MMMM yyyy' : 'dd/MM/yyyy HH.mm');
}

/**
 * Layar "Masuk" biasa: username + kata kunci, disimpan di properti script (tidak ada di HTML).
 * Ini penghalang pandangan, bukan pengganti kunci data — datanya tetap dijaga akun Google.
 */
const PROP_LOGIN_USER = 'LOGIN_USERNAME';
const PROP_LOGIN_KEY = 'LOGIN_KATA_KUNCI';
const PROP_LOGIN_OWNER = 'LOGIN_OWNER';

function hashKataKunci_(teks) {
  const byte = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(teks || ''), Utilities.Charset.UTF_8);
  return byte.map(function (b) { return (b + 256).toString(16).slice(-2); }).join('');
}

function butuhLogin_() {
  return !!PropertiesService.getScriptProperties().getProperty(PROP_LOGIN_KEY);
}

/** Reset kata kunci hanya untuk akun Google yang memasang login itu pertama kali. */
function akunPemilikLogin_() {
  const props = PropertiesService.getScriptProperties();
  const pemilik = String(props.getProperty(PROP_LOGIN_OWNER) || '');
  const aktif = String(emailAktif_() || '');
  return { boleh: !pemilik || !aktif || pemilik.toLowerCase() === aktif.toLowerCase(), email: aktif, pemilik: pemilik };
}

function login(username, kataKunci) {
  const props = PropertiesService.getScriptProperties();
  const kunci = props.getProperty(PROP_LOGIN_KEY) || '';
  if (!kunci) return { ok: true, belumDipasang: true };
  const userCocok = String(username || '').trim().toLowerCase() === String(props.getProperty(PROP_LOGIN_USER) || '').toLowerCase();
  const kunciCocok = hashKataKunci_(kataKunci) === kunci;
  if (!userCocok || !kunciCocok) return { ok: false };
  return { ok: true, token: sesiBuat_() };
}


/** Pasang atau ganti login. Kata kunci lama tidak diminta lagi — yang bisa membuka halaman ini sudah akun pemilik. */
function setLogin(username, kataKunci) {
  butuhSesi_();
  const props = PropertiesService.getScriptProperties();
  const user = String(username || '').trim();
  if (!user) throw new Error('Username belum diisi.');
  if (!kataKunci || String(kataKunci).length < 4) throw new Error('Kata kunci minimal 4 karakter.');
  const lama = props.getProperty(PROP_LOGIN_KEY) || '';
  props.setProperty(PROP_LOGIN_USER, user);
  props.setProperty(PROP_LOGIN_KEY, hashKataKunci_(kataKunci));
  if (!props.getProperty(PROP_LOGIN_OWNER)) props.setProperty(PROP_LOGIN_OWNER, emailAktif_());
  return { ok: true, username: user, dibuat: !lama };
}

/** Layar "Lupa kata kunci": cek dulu apakah yang membuka ini benar pemilik aplikasinya. */
function lupaKataKunci() {
  if (!butuhLogin_()) return { boleh: false, alasan: 'Belum ada login yang dipasang, jadi tidak ada yang perlu direset.' };
  const cek = akunPemilikLogin_();
  if (!cek.boleh) {
    return { boleh: false, alasan: 'Reset hanya bisa dari akun Google yang memasang login ini (' + cek.pemilik + '). Akun kamu sekarang: ' + (cek.email || 'tidak terbaca') + '.' };
  }
  return { boleh: true, email: cek.email };
}

/** Ganti kata kunci tanpa yang lama — cuma jalan kalau akun Google-nya memang pemilik aplikasi. */
function resetKataKunci(kataKunciBaru) {
  const cek = akunPemilikLogin_();
  if (!cek.boleh) throw new Error('Bukan akun pemilik, jadi kata kunci tidak bisa direset dari sini.');
  if (!kataKunciBaru || String(kataKunciBaru).length < 4) throw new Error('Kata kunci baru minimal 4 karakter.');
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_LOGIN_KEY, hashKataKunci_(kataKunciBaru));
  return { ok: true };
}

function lepasLogin(kataKunci) {
  const props = PropertiesService.getScriptProperties();
  const lama = props.getProperty(PROP_LOGIN_KEY) || '';
  if (!lama) return { ok: true, sudahKosong: true };
  if (hashKataKunci_(kataKunci) !== lama) throw new Error('Kata kunci tidak cocok.');
  props.deleteProperty(PROP_LOGIN_KEY);
  props.deleteProperty(PROP_LOGIN_USER);
  props.deleteProperty(PROP_LOGIN_OWNER);
  return { ok: true };
}

/** Jalan-jalan keluar kalau lupa kata kunci: buka editor Apps Script, pilih fungsi ini, Run. */
function resetLogin() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_LOGIN_KEY);
  props.deleteProperty(PROP_LOGIN_USER);
  props.deleteProperty(PROP_LOGIN_OWNER);
  Logger.log('Login dilepas. Kata kunci berikutnya dipasang lewat menu Pengaturan.');
}

/* ── Kunci lapis dua: sesi login (token) ────────────────────────────
   Aplikasi boleh dibuka "siapa saja di web" tanpa login Google, jadi kuncinya
   token yang diberikan server waktu login — bukan identitas akun. Fungsi data
   tetap mengecek sesinya sendiri, jadi panggilan langsung dari luar (tidak
   lewat panggilan()) otomatis tanpa token dan ditolak. */
const PROP_SESI = "SESI_TOKEN_";
const SESI_BASI_MS = 12 * 60 * 60 * 1000;
var SESI_TOKEN_PAKAI_ = "";

function sesiBuat_() {
  const token = String(Utilities.getUuid());
  PropertiesService.getScriptProperties().setProperty(PROP_SESI + token, String(Date.now()));
  return token;
}

function sesiHapus_(token) {
  const kunci = String(token || "");
  if (kunci) PropertiesService.getScriptProperties().deleteProperty(PROP_SESI + kunci);
}

function sesiSah_() {
  if (!butuhLogin_()) return true;
  if (!SESI_TOKEN_PAKAI_) return false;
  const disimpan = PropertiesService.getScriptProperties().getProperty(PROP_SESI + SESI_TOKEN_PAKAI_);
  return !!disimpan && Date.now() - Number(disimpan) <= SESI_BASI_MS;
}

function butuhSesi_() {
  if (!sesiSah_()) throw new Error("Sesi login belum aktif atau sudah basi. Masuk lewat halaman login dulu.");
  PropertiesService.getScriptProperties().setProperty(PROP_SESI + SESI_TOKEN_PAKAI_, String(Date.now()));
}

const RUTE_BEBAS = {
  getBootstrap: function (args) { return getBootstrap.apply(null, args); },
  getInitData: function (args) { return getInitData.apply(null, args); },
  kenapaBelumBisa: function (args) { return kenapaBelumBisa.apply(null, args); },
  login: function (args) { return login.apply(null, args); },
  lupaKataKunci: function (args) { return lupaKataKunci.apply(null, args); },
  resetKataKunci: function (args) { return resetKataKunci.apply(null, args); },
  keluar: function (args) { return keluar.apply(null, args); },
};

const RUTE_DATA = {
  ambilSemuaTab: function (args) { return ambilSemuaTab.apply(null, args); },
  bersihkanFileProyek: function (args) { return bersihkanFileProyek.apply(null, args); },
  buangTabSisa: function (args) { return buangTabSisa.apply(null, args); },
  buatDataContoh: function (args) { return buatDataContoh.apply(null, args); },
  buatFormBrand: function (args) { return buatFormBrand.apply(null, args); },
  buatFormContoh: function (args) { return buatFormContoh.apply(null, args); },
  cekAksesTulis: function (args) { return cekAksesTulis.apply(null, args); },
  cekSistem: function (args) { return cekSistem.apply(null, args); },
  daftarFileProyek: function (args) { return daftarFileProyek.apply(null, args); },
  daftarIntegrasi: function (args) { return daftarIntegrasi.apply(null, args); },
  daftarSheetIntegrasi: function (args) { return daftarSheetIntegrasi.apply(null, args); },
  daftarSisa: function (args) { return daftarSisa.apply(null, args); },
  daftarTabSisa: function (args) { return daftarTabSisa.apply(null, args); },
  eksporCsv: function (args) { return eksporCsv.apply(null, args); },
  getBrandData: function (args) { return getBrandData.apply(null, args); },
  getDashboard: function (args) { return getDashboard.apply(null, args); },
  getDetail: function (args) { return getDetail.apply(null, args); },
  getInitData: function (args) { return getInitData.apply(null, args); },
  getKonfigurasi: function (args) { return getKonfigurasi.apply(null, args); },
  getLaporan: function (args) { return getLaporan.apply(null, args); },
  getOverview: function (args) { return getOverview.apply(null, args); },
  hapusBrand: function (args) { return hapusBrand.apply(null, args); },
  hapusBrandLengkap: function (args) { return hapusBrandLengkap.apply(null, args); },
  hapusDataContoh: function (args) { return hapusDataContoh.apply(null, args); },
  hapusFormDanTab: function (args) { return hapusFormDanTab.apply(null, args); },
  hapusIntegrasi: function (args) { return hapusIntegrasi.apply(null, args); },
  bersihkanContoh: function (args) { return bersihkanContoh.apply(null, args); },
  bersihkanDuplikatIntegrasi: function (args) { return bersihkanDuplikatIntegrasi.apply(null, args); },
  isiDataContoh: function (args) { return isiDataContoh.apply(null, args); },
  kirimSertifikat: function (args) { return kirimSertifikat.apply(null, args); },
  kolomIntegrasi: function (args) { return kolomIntegrasi.apply(null, args); },
  pakaiSatuSpreadsheet: function (args) { return pakaiSatuSpreadsheet.apply(null, args); },
  periksaDokumen: function (args) { return periksaDokumen.apply(null, args); },
  pindahFileKeFolder: function (args) { return pindahFileKeFolder.apply(null, args); },
  pindahkanTujuanForm: function (args) { return pindahkanTujuanForm.apply(null, args); },
  sambungkanBrand: function (args) { return sambungkanBrand.apply(null, args); },
  sasaranBaris: function (args) { return sasaranBaris.apply(null, args); },
  setJadwalWawancara: function (args) { return setJadwalWawancara.apply(null, args); },
  setLogin: function (args) { return setLogin.apply(null, args); },
  setelArsip: function (args) { return setelArsip.apply(null, args); },
  setelStatusIntegrasi: function (args) { return setelStatusIntegrasi.apply(null, args); },
  sinkronIntegrasi: function (args) { return sinkronIntegrasi.apply(null, args); },
  sinkronSemuaIntegrasi: function (args) { return sinkronSemuaIntegrasi.apply(null, args); },
  tambahFollowUp: function (args) { return tambahFollowUp.apply(null, args); },
  tambahIntegrasi: function (args) { return tambahIntegrasi.apply(null, args); },
  ubahIntegrasi: function (args) { return ubahIntegrasi.apply(null, args); },
  ujiIntegrasi: function (args) { return ujiIntegrasi.apply(null, args); },
  updateStatus: function (args) { return updateStatus.apply(null, args); },
};

/** Satu-satunya pintu masuk dari halaman: nama aksi, token sesi, lalu argumennya. */
function panggilan(nama, token, args) {
  SESI_TOKEN_PAKAI_ = String(token || "");
  const bebas = RUTE_BEBAS[nama];
  const data = bebas || RUTE_DATA[nama];
  if (!data) throw new Error("Aksi tidak dikenal: " + nama);
  if (!bebas) butuhSesi_();
  try {
    return data(args || []);
  } catch (e) {
    // Pesan mentah Google cuma bilang "tidak memiliki izin" tanpa menyebut akun mana
    // yang harus ditambah dan file apa yang macet. Semua error dilewatkan pesanIzin_.
    throw new Error(pesanIzin_(e));
  }
}

/** Tombol Keluar: token dibuang, jadi sesi ini mati total di server. */
function keluar(token) {
  sesiHapus_(token);
  return { ok: true };
}

/* ── Pemulihan setelah proyek Apps Script dibuat ulang ─────────────────────
   Spreadsheet datanya TETAP yang lama. Brand diambil dari tab-tab yang sudah
   ada di file itu, dan file log lama dicari lagi lewat namanya supaya catatan
   follow-up serta konfigurasi Integrasi tidak dibuat dari nol. */
const SS_DATA_LAMA = "1_jyDo2ZLlhzqI2Qb9ydDPCv_Qmkb3NJH_uzv9Btc63s";

/** Dipakai getBrands_(): sekali saja per sesi eksekusi, diam-diam kalau gagal. */
function pulihkanOtomatis_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("PULIHKAN_COBA")) return;
  props.setProperty("PULIHKAN_COBA", String(Date.now()));
  try {
    pulihkanPengaturan();
  } catch (e) {
    Logger.log("Pemulihan otomatis tidak berhasil: " + (e.message || e));
  }
}

function pulihkanPengaturan() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(SS_DATA_LAMA);
  const daftar = [];
  ss.getSheets().forEach(function (sheet) {
    const nama = sheet.getName();
    if (!nama || nama.toLowerCase() === "integrasi" || nama.toLowerCase() === "log") return;
    daftar.push({ name: nama, spreadsheetId: ss.getId(), sheetName: nama, formId: "", formUrl: "" });
  });
  setBrands_(daftar);

  let log = "";
  const files = DriveApp.getFilesByName(NAMA_LOG);
  if (files.hasNext()) {
    const f = files.next();
    props.setProperty(PROP_LOG, f.getId());
    log = f.getName();
  }

  batalCache_();
  return {
    spreadsheet: ss.getName(),
    brand: daftar.map(function (b) { return b.name; }),
    log: log || "file log belum ketemu — akan dibuat saat ada follow-up pertama",
  };
}
