/**
 * ============================================================================
 * SISTEM PELAPORAN & REKAPITULASI PEMBAYARAN RETRIBUSI SPPG
 * DINAS LINGKUNGAN HIDUP (DLH) KABUPATEN PANGKAJENE DAN KEPULAUAN
 * ============================================================================
 * Stack: Google Apps Script (GAS) + Google Sheets + Google Drive + HTML Service
 * File: code.gs (Backend & Database Controller)
 * ============================================================================
 */

// Global Sheet Names
var SHEETS = {
  MASTER_SPPG: 'master_sppg',
  PEMBAYARAN: 'pembayaran_retribusi',
  PERIODE_WAJIB: 'periode_wajib',
  USERS: 'users_bendahara',
  CONFIG: 'app_config'
};

// 9 Bulan Pelayanan
var BULAN_LIST = ['April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
var BULAN_SHORT = ['Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

// ============================================================================
// 1. WEB APP & REST API ENTRY POINTS (doGet & doPost)
// ============================================================================
function doGet(e) {
  // 1. Jika ada parameter action, kembalikan data REST API JSON
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter.action, e.parameter);
  }

  // 2. Default: Render Web App HTML jika dibuka langsung di Apps Script
  var template = HtmlService.createTemplateFromFile('index');
  var htmlOutput = template.evaluate()
    .setTitle('Retribusi SPPG - Dinas Lingkungan Hidup')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (err) {
        body = e.parameter || {};
      }
    } else if (e && e.parameter) {
      body = e.parameter;
    }

    var action = body.action || (e && e.parameter && e.parameter.action);
    return handleApiRequest(action, body);
  } catch (error) {
    return jsonResponse({ success: false, message: 'Server error: ' + error.toString() });
  }
}

function handleApiRequest(action, payload) {
  try {
    var result = { success: false, message: 'Aksi tidak dikenali: ' + action };

    switch (action) {
      case 'getPublicDashboardData':
      case 'getInitialData':
        result = getPublicDashboardData();
        break;

      case 'submitKonfirmasiPembayaran':
      case 'submitPelaporan':
        result = submitKonfirmasiPembayaran(payload.data || payload);
        break;

      case 'loginBendahara':
        result = loginBendahara(payload.email || payload.username, payload.password);
        break;

      case 'saveTransaksiInline':
        result = saveTransaksiInline(payload.idTrx || payload.id, payload.noSts, payload.noStbp, payload.idTransaksi || payload.noTrxBank);
        break;

      case 'updateTrxBank':
        result = updateTrxBank(payload.idTrx || payload.id, payload.noTrxBank || payload.val);
        break;

      case 'updateStbp':
        result = updateStbp(payload.idTrx || payload.id, payload.noStbp || payload.val);
        break;

      case 'toggleVerifikasiPembayaran':
      case 'verifikasiPembayaran':
        result = toggleVerifikasiPembayaran(payload.idTrx || payload.id, payload.status);
        break;

      case 'updateSts':
        result = updateSts(payload.idTrx || payload.id, payload.noSts, payload.tanggalSts || payload.tglSts);
        break;

      case 'savePeriodeWajib':
        result = savePeriodeWajib(payload.periodeMap || payload.data);
        break;

      case 'addMasterSppg':
        result = addMasterSppg(payload.data || payload);
        break;

      case 'saveAppConfig':
        result = saveAppConfig(payload.configMap || payload.data);
        break;

      default:
        result = { success: false, message: 'Action not found: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Helper: Get Active Spreadsheet
function getDb() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(sheetName) {
  var ss = getDb();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

// ============================================================================
// 2. SETUP DATABASE & INITIAL SEEDING
// ============================================================================

/**
 * Jalankan fungsi setupDatabase() ini dari Apps Script Editor untuk
 * inisialisasi database awal (bersih/kosong):
 * 1. Membuat 5 sheet utama terformat (master_sppg, users_bendahara, app_config, periode_wajib, pembayaran_retribusi)
 * 2. Mengisi akun admin default (bendahara@dlh.go.id / bendahara123)
 * 3. Mengisi 30 master SPPG Kabupaten Pangkep
 * 4. Mengisi default matriks periode wajib (April - Desember)
 * (TIDAK menyertakan import data transaksi historis)
 */
function setupDatabase() {
  var ss = getDb();
  Logger.log('=== MEMULAI SETUP DATABASE KOSONG RETRIBUSI SPPG DLH ===');

  // 1. Setup Sheet: master_sppg
  setupSheetMasterSppg(ss);

  // 2. Setup Sheet: users_bendahara
  setupSheetUsers(ss);

  // 3. Setup Sheet: app_config
  setupSheetConfig(ss);

  // 4. Setup Sheet: periode_wajib
  setupSheetPeriodeWajib(ss);

  // 5. Setup Sheet: pembayaran_retribusi
  setupSheetPembayaran(ss);

  // 6. Insert Default Admin
  insertDefaultAdmin();

  // 7. Insert Master 30 SPPG
  insertMasterSppg();

  // 8. Insert Periode Wajib Default
  initPeriodeWajibDefault();

  Logger.log('=== SETUP DATABASE KOSONG SELESAI DENGAN SUKSES! ===');
  return {
    status: 'success',
    message: 'Database kosong berhasil disiapkan beserta master data awal (Admin, SPPG, Konfigurasi, Periode Wajib).'
  };
}

function styleHeaderRow(sheet, numCols) {
  var range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground('#1e3a8a') // Navy DLH
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

function setupSheetMasterSppg(ss) {
  var sheet = ss.getSheetByName(SHEETS.MASTER_SPPG) || ss.insertSheet(SHEETS.MASTER_SPPG);
  var headers = ['id_sppg', 'kecamatan', 'nama_sppg', 'penanggung_jawab', 'kontak', 'status_aktif', 'created_at'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

function setupSheetUsers(ss) {
  var sheet = ss.getSheetByName(SHEETS.USERS) || ss.insertSheet(SHEETS.USERS);
  var headers = ['id_user', 'username', 'email', 'password', 'nama_petugas', 'role', 'status_aktif', 'created_at'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

function setupSheetConfig(ss) {
  var sheet = ss.getSheetByName(SHEETS.CONFIG) || ss.insertSheet(SHEETS.CONFIG);
  var headers = ['config_key', 'config_value', 'deskripsi', 'updated_at'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);

  var defaultConfigs = [
    ['appName', 'Retribusi SPPG', 'Nama Aplikasi Web Portal', new Date()],
    ['tagline', 'Dinas Lingkungan Hidup Kabupaten Pangkajene dan Kepulauan', 'Tagline resmi instansi', new Date()],
    ['tahunAnggaran', '2026', 'Tahun Anggaran Berjalan', new Date()],
    ['logoBg', 'linear-gradient(135deg, #2563eb, #059669)', 'Warna latar logo header', new Date()],
    ['customLogoUrl', '', 'URL gambar logo kustom', new Date()],
    ['footerText', 'Dinas Lingkungan Hidup (DLH) Kabupaten Pangkajene dan Kepulauan', 'Teks hak cipta footer', new Date()],
    ['templateStbp', 'STBP/DLH/{TAHUN}/{BULAN_SINGKAT}/{PAD4}', 'Template Format Nomor STBP (Token: {TAHUN}, {BULAN}, {BULAN_SINGKAT}, {BULAN_ROMAWI}, {NO}, {PAD4})', new Date()],
    ['templateSts', 'STS/KASDA/{TAHUN}/{BULAN_SINGKAT}/{PAD4}', 'Template Format Nomor STS Kasda (Token: {TAHUN}, {BULAN}, {BULAN_SINGKAT}, {BULAN_ROMAWI}, {NO}, {PAD4})', new Date()],
    ['templateTrx', 'TRX-{TAHUN}{BULAN_ANGKA}-{PAD4}', 'Template Format Kode Transaksi / Ref Bank (Token: {TAHUN}, {BULAN_ANGKA}, {NO}, {PAD4})', new Date()],
    ['maxFileSizeMB', '5', 'Batas maksimal upload slip transfer (MB)', new Date()],
    ['waHelpdesk', '0812-3456-7890', 'Nomor WhatsApp Helpdesk DLH', new Date()],
    ['driveFolderId', '1AbC_dlh_retribusi_sppg_drive_folder', 'ID Folder Google Drive penyimpanan bukti slip', new Date()]
  ];

  sheet.getRange(2, 1, defaultConfigs.length, 4).setValues(defaultConfigs);
}

function setupSheetPeriodeWajib(ss) {
  var sheet = ss.getSheetByName(SHEETS.PERIODE_WAJIB) || ss.insertSheet(SHEETS.PERIODE_WAJIB);
  var headers = ['id_sppg', 'kecamatan', 'nama_sppg', 'jan', 'feb', 'mar', 'apr', 'mei', 'jun', 'jul', 'agu', 'sep', 'okt', 'nov', 'des', 'updated_at'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

function setupSheetPembayaran(ss) {
  var sheet = ss.getSheetByName(SHEETS.PEMBAYARAN) || ss.insertSheet(SHEETS.PEMBAYARAN);
  var headers = [
    'id_transaksi', 
    'id_sppg', 
    'nama_sppg', 
    'kecamatan', 
    'nama_pelapor', 
    'kontak_pelapor', 
    'periode_bulan', 
    'tanggal_transfer', 
    'jumlah_transfer', 
    'bukti_slip_url', 
    'status_verifikasi', 
    'no_sts', 
    'no_stbp', 
    'id_transaksi_bank', 
    'tanggal_sts', 
    'catatan', 
    'created_at'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

// ============================================================================
/**
 * Jalankan fungsi ini untuk otomatis menamai ulang seluruh baris header tabel database ke format standar.
 */
function migrateDatabaseHeaders() {
  try {
    var ss = getDb();

    // 1. master_sppg
    var s1 = ss.getSheetByName(SHEETS.MASTER_SPPG);
    if (s1) {
      s1.getRange(1, 1, 1, 7).setValues([['id_sppg', 'kecamatan', 'nama_sppg', 'penanggung_jawab', 'kontak', 'status_aktif', 'created_at']]);
      styleHeaderRow(s1, 7);
    }

    // 2. users_bendahara
    var s2 = ss.getSheetByName(SHEETS.USERS);
    if (s2) {
      s2.getRange(1, 1, 1, 8).setValues([['id_user', 'username', 'email', 'password', 'nama_petugas', 'role', 'status_aktif', 'created_at']]);
      styleHeaderRow(s2, 8);
    }

    // 3. app_config
    var s3 = ss.getSheetByName(SHEETS.CONFIG);
    if (s3) {
      s3.getRange(1, 1, 1, 4).setValues([['config_key', 'config_value', 'deskripsi', 'updated_at']]);
      styleHeaderRow(s3, 4);
    }

    // 4. periode_wajib
    var s4 = ss.getSheetByName(SHEETS.PERIODE_WAJIB);
    if (s4) {
      var pwHeaders = ['id_sppg', 'kecamatan', 'nama_sppg', 'jan', 'feb', 'mar', 'apr', 'mei', 'jun', 'jul', 'agu', 'sep', 'okt', 'nov', 'des', 'updated_at'];
      s4.getRange(1, 1, 1, pwHeaders.length).setValues([pwHeaders]);
      styleHeaderRow(s4, pwHeaders.length);
    }

    // 5. pembayaran_retribusi
    var s5 = ss.getSheetByName(SHEETS.PEMBAYARAN);
    if (s5) {
      var payHeaders = ['id_transaksi', 'id_sppg', 'nama_sppg', 'kecamatan', 'nama_pelapor', 'kontak_pelapor', 'periode_bulan', 'tanggal_transfer', 'jumlah_transfer', 'bukti_slip_url', 'status_verifikasi', 'no_sts', 'no_stbp', 'id_transaksi_bank', 'tanggal_sts', 'catatan', 'created_at'];
      s5.getRange(1, 1, 1, payHeaders.length).setValues([payHeaders]);
      styleHeaderRow(s5, payHeaders.length);
    }

    Logger.log('Semua header tabel database berhasil distandarisasi.');
    return { success: true, message: 'Semua header tabel database berhasil distandarisasi.' };
  } catch (err) {
    Logger.log('Error migrateDatabaseHeaders: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// 3. INSERT MASTER DATA (SPPG, ADMIN, TRANSAKSI DARI EXCEL)
// ============================================================================

function insertDefaultAdmin() {
  var sheet = getSheet(SHEETS.USERS);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    Logger.log('Admin default sudah ada.');
    return;
  }

  var adminData = [
    ['USR-001', 'bendahara', 'bendahara@dlh.go.id', 'bendahara123', 'Bendahara Retribusi DLH', 'ADMIN', 'AKTIF', new Date()]
  ];
  sheet.getRange(2, 1, adminData.length, adminData[0].length).setValues(adminData);
  Logger.log('Admin default bendahara / bendahara@dlh.go.id berhasil ditambahkan.');
}

/**
 * Jalankan fungsi ini jika database Google Sheets lama Anda belum memiliki kolom 'Username' di sheet users_bendahara.
 * Fungsi ini otomatis menambahkan kolom 'Username' di kolom B dan mengisi nilai username dari email.
 */
function migrateAddUsernameColumn() {
  var sheet = getSheet(SHEETS.USERS);
  if (!sheet) return { status: 'error', message: 'Sheet users_bendahara tidak ditemukan.' };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var usernameIdx = headers.indexOf('Username');
  if (usernameIdx !== -1) {
    Logger.log('Kolom Username sudah ada di database.');
    return { status: 'info', message: 'Kolom Username sudah ada di database.' };
  }

  // Sisipkan kolom baru di posisi kolom B (setelah ID User)
  sheet.insertColumnAfter(1);
  sheet.getRange(1, 2).setValue('Username')
       .setBackground('#1e3a8a')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');

  // Isi data username dari email untuk baris yang ada
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var emailVals = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    var usernames = emailVals.map(function(r) {
      var email = String(r[0] || '').trim();
      var uname = email.split('@')[0] || 'admin';
      return [uname];
    });
    sheet.getRange(2, 2, usernames.length, 1).setValues(usernames);
  }

  Logger.log('Migrasi kolom Username pada database berhasil diselesaikan.');
  return { status: 'success', message: 'Kolom Username berhasil ditambahkan ke database!' };
}

var MASTER_SPPG_LIST = [
  { id: 'SPPG-001', kec: 'Balocci', nama: 'Kassi', penanggung_jawab: 'Siti Zuraima SPPG Kassi1', kontak: '089501737683' },
  { id: 'SPPG-002', kec: 'Balocci', nama: 'Kassi 2', penanggung_jawab: 'Darwan Muis SPPG Kassi2', kontak: '085393517300' },
  { id: 'SPPG-003', kec: 'Bungoro', nama: 'Samalewa 1', penanggung_jawab: 'Haerul Fahresi SPPG Samalewa1', kontak: '087714145593' },
  { id: 'SPPG-004', kec: 'Bungoro', nama: 'Samalewa 2', penanggung_jawab: 'Nurul Mutmainnah SPPG Samalewa 2', kontak: '085756416172' },
  { id: 'SPPG-005', kec: 'Bungoro', nama: 'Samalewa 3', penanggung_jawab: 'Wahyuddin SPPG Samalewa3', kontak: '085342286433' },
  { id: 'SPPG-006', kec: 'Bungoro', nama: 'Samalewa 4', penanggung_jawab: 'Musakkir SPPG Samalewa4', kontak: '082296461855' },
  { id: 'SPPG-007', kec: 'Bungoro', nama: 'Samalewa 5', penanggung_jawab: 'Anugrahwan SPPG Samalewa 5', kontak: '081340075010' },
  { id: 'SPPG-008', kec: 'Labakkang', nama: 'Labakkang', penanggung_jawab: 'Ardiansyah SPPG Labakkang', kontak: '085975073614' },
  { id: 'SPPG-009', kec: 'Labakkang', nama: 'Labakkang 2', penanggung_jawab: 'Muammar SPPG Labakkang 2', kontak: '082349170976' },
  { id: 'SPPG-010', kec: 'Labakkang', nama: 'Labakkang 3', penanggung_jawab: 'Heris SPPG Labakkang 3', kontak: '085343793941' },
  { id: 'SPPG-011', kec: 'Labakkang', nama: 'Manakku', penanggung_jawab: 'Rabiullanda Kulsum SPPG Manakku', kontak: '085341370164' },
  { id: 'SPPG-012', kec: 'Labakkang', nama: 'Mangallekana', penanggung_jawab: 'Nur Chaerunnisa SPPG Mangallekana', kontak: '082315045264' },
  { id: 'SPPG-013', kec: 'Labakkang', nama: 'Batara', penanggung_jawab: 'Fatma Sri Fatimah SPPG Batara', kontak: '088744874889' },
  { id: 'SPPG-014', kec: 'Mandalle', nama: 'Manggalung', penanggung_jawab: 'Adi Irwandi SPPG Manggalung', kontak: '085242761351' },
  { id: 'SPPG-015', kec: 'Mandalle', nama: 'Tamarupa', penanggung_jawab: 'Agatha Febriandani SPPG Tamarupa', kontak: '085346164018' },
  { id: 'SPPG-016', kec: 'Marang', nama: 'Talaka', penanggung_jawab: 'Awal Fajaruddin SPPG Talaka1', kontak: '089684444665' },
  { id: 'SPPG-017', kec: 'Marang', nama: 'Talaka 2', penanggung_jawab: 'Nurul Afian SPPG Talaka2', kontak: '085397984559' },
  { id: 'SPPG-018', kec: 'Minasa Tene', nama: 'Bonto Langkasa', penanggung_jawab: 'Indah Putri Humairah SPPG Bonto Langkasa', kontak: '087858320557' },
  { id: 'SPPG-019', kec: 'Minasa Tene', nama: 'Kabba', penanggung_jawab: 'Fani Fajriani SPPG Kabba', kontak: '082318479896' },
  { id: 'SPPG-020', kec: 'Minasa Tene', nama: 'Biraeng', penanggung_jawab: 'Muhammad Ardas Daruslam SPPG Biraeng', kontak: '082188224010' },
  { id: 'SPPG-021', kec: 'Pangkajene', nama: 'Bonto Perak 1', penanggung_jawab: 'Fahmi Sofyan SPPG Bonto Perak1', kontak: '081244287373' },
  { id: 'SPPG-022', kec: 'Pangkajene', nama: 'Bonto Perak 2', penanggung_jawab: 'Saharuddin SPPG Bonto Perak2', kontak: '085343693131' },
  { id: 'SPPG-023', kec: 'Pangkajene', nama: 'Mappasaile', penanggung_jawab: 'Evi Irviyanti SPPG Mappasaile', kontak: '085870552736' },
  { id: 'SPPG-024', kec: 'Pangkajene', nama: 'Mappasaile 2', penanggung_jawab: 'Akbar Tola SPPG Mappasaile 2', kontak: '085340312677' },
  { id: 'SPPG-025', kec: 'Pangkajene', nama: 'Padoang Doangan', penanggung_jawab: 'Amran SPPG Padoang2an', kontak: '085230740101' },
  { id: 'SPPG-026', kec: 'Pangkajene', nama: 'Tumampua', penanggung_jawab: "Mar'atul Islam SPPG Tumampua", kontak: '085338318156' },
  { id: 'SPPG-027', kec: 'Segeri', nama: 'Bone', penanggung_jawab: 'Mawarni Utami SPPG Bone', kontak: '085346567732' },
  { id: 'SPPG-028', kec: 'Segeri', nama: 'Segeri', penanggung_jawab: 'Ashar Saputra SPPG Segeri', kontak: '085259870711' },
  { id: 'SPPG-029', kec: 'Segeri', nama: 'Segeri 2', penanggung_jawab: 'Amell Akuntan SPPG Segeri2', kontak: '082189423315' },
  { id: 'SPPG-030', kec: 'Tondong Tallasa', nama: 'Bantimurung', penanggung_jawab: 'Hauliah SPPG Bantimurung', kontak: '085240966666' }
];

function insertMasterSppg() {
  var sheet = getSheet(SHEETS.MASTER_SPPG);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    Logger.log('Data master SPPG sudah ada.');
    return;
  }

  var rows = MASTER_SPPG_LIST.map(function(item) {
    return [item.id, item.kec, item.nama, item.penanggung_jawab || '', item.kontak || '', 'AKTIF', new Date()];
  });

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('30 Master SPPG berhasil ditambahkan.');
}

function initPeriodeWajibDefault() {
  var sheet = getSheet(SHEETS.PERIODE_WAJIB);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    Logger.log('Data periode wajib sudah ada.');
    return;
  }

  var rows = MASTER_SPPG_LIST.map(function(item) {
    var apr = true, mei = true, jun = true, jul = true, agu = true, sep = true, okt = true, nov = true, des = true;
    
    // SPPG Biraeng (SPPG-020) mulai beroperasi di bulan September
    if (item.id === 'SPPG-020') {
      apr = false; mei = false; jun = false; jul = false; agu = false;
    } else if (item.id === 'SPPG-013') { // Batara mulai Mei
      apr = false;
    }

    return [
      item.id,
      item.kec,
      item.nama,
      apr,
      mei,
      jun,
      jul,
      agu,
      sep,
      okt,
      nov,
      des,
      new Date()
    ];
  });

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('Matriks Periode Wajib berhasil diinisialisasi.');
}

// Catatan: Fungsi dan data import transaksi historis Excel (April - Agustus)
// telah dipisahkan ke file tersendiri: importdata.gs

// ============================================================================
// 4. API FUNCTIONS FOR WEB APP (DIKONSUMSI CLIENT / FRONTEND)
// ============================================================================

/**
 * Mengambil data publik untuk Dashboard (Master SPPG, Matriks, Chart, Penyetor Terbaru, Config)
 */
function getPublicDashboardData() {
  try {
    var ss = getDb();

    // 1. Get Config
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    var config = {};
    if (configSheet && configSheet.getLastRow() > 1) {
      var configVals = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
      configVals.forEach(function(row) {
        config[row[0]] = row[1];
      });
    }

    // 2. Get Master SPPG
    var sppgSheet = ss.getSheetByName(SHEETS.MASTER_SPPG);
    var masterSppg = [];
    if (sppgSheet && sppgSheet.getLastRow() > 1) {
      var sppgVals = sppgSheet.getRange(2, 1, sppgSheet.getLastRow() - 1, 6).getValues();
      masterSppg = sppgVals.map(function(r) {
        return { 
          id: r[0], 
          kec: r[1], 
          nama: r[2], 
          penanggung_jawab: r[3] || '',
          penanggungJawab: r[3] || '',
          pengurus: r[3] || '',
          kontak: r[4] || '',
          active: r[5] === 'AKTIF' 
        };
      });
    } else {
      masterSppg = MASTER_SPPG_LIST;
    }

    // 3. Get Periode Wajib Map
    var pwSheet = ss.getSheetByName(SHEETS.PERIODE_WAJIB);
    var periodeWajibMap = {};
    if (pwSheet && pwSheet.getLastRow() > 1) {
      var pwVals = pwSheet.getRange(2, 1, pwSheet.getLastRow() - 1, 13).getValues();
      pwVals.forEach(function(r) {
        var sppgId = r[0];
        periodeWajibMap[sppgId] = {};
        for (var i = 0; i < BULAN_SHORT.length; i++) {
          periodeWajibMap[sppgId][BULAN_SHORT[i]] = (r[3 + i] === true || r[3 + i] === 'TRUE');
        }
      });
    }

    // 4. Get Transaksi
    var paySheet = ss.getSheetByName(SHEETS.PEMBAYARAN);
    var transaksiList = [];
    if (paySheet && paySheet.getLastRow() > 1) {
      var payVals = paySheet.getRange(2, 1, paySheet.getLastRow() - 1, 17).getValues();
      transaksiList = payVals.map(function(r) {
        return {
          id: r[0],
          sppgId: r[1],
          namaSppg: r[2],
          kecamatan: r[3],
          namaPelapor: r[4],
          kontakPelapor: r[5] !== '' && r[5] !== null && r[5] !== undefined ? String(r[5]) : '',
          periodeBulan: r[6],
          tanggalTransfer: r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[7]),
          jumlahTransfer: Number(r[8]),
          buktiUrl: r[9],
          status: r[10],
          noSts: r[11] !== '' && r[11] !== null && r[11] !== undefined ? String(r[11]) : '',
          noStbp: r[12] !== '' && r[12] !== null && r[12] !== undefined ? String(r[12]) : '',
          noTrxBank: r[13] !== '' && r[13] !== null && r[13] !== undefined ? String(r[13]) : '',
          tanggalSts: r[14] instanceof Date ? Utilities.formatDate(r[14], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[14]),
          catatan: r[15]
        };
      });
    }

    return {
      success: true,
      config: config,
      masterSppg: masterSppg,
      periodeWajibMap: periodeWajibMap,
      transaksiList: transaksiList
    };
  } catch (err) {
    Logger.log('Error getPublicDashboardData: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * Submit Konfirmasi Pembayaran oleh SPPG (Akses Publik)
 */
function submitKonfirmasiPembayaran(payload) {
  try {
    var ss = getDb();
    var sheet = getSheet(SHEETS.PEMBAYARAN);

    var idTrx = 'TRX-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM') + '-' + Math.floor(1000 + Math.random() * 9000);

    // Handle Upload file slip ke Google Drive jika ada
    var buktiUrl = '';
    if (payload.fileBase64 && payload.fileName) {
      try {
        var uploadedUrl = uploadSlipToDrive(payload.fileBase64, payload.fileName, payload.fileType, {
          idTrx: idTrx,
          kecamatan: payload.kecamatan,
          namaSppg: payload.namaSppg,
          periodeBulan: payload.periodeBulan,
          tanggalTransfer: payload.tanggalTransfer
        });
        if (uploadedUrl) {
          buktiUrl = uploadedUrl;
        }
      } catch (eUpload) {
        Logger.log('Gagal uploadSlipToDrive: ' + eUpload.toString());
      }
    } else if (payload.buktiUrl && !String(payload.buktiUrl).startsWith('data:')) {
      buktiUrl = String(payload.buktiUrl);
    }

    // Safeguard: Jangan biarkan base64 string jutaan karakter lolos ke Spreadsheet (limit sel 50.000 char)
    if (buktiUrl && (String(buktiUrl).indexOf('data:') === 0 || String(buktiUrl).length > 1000)) {
      buktiUrl = '';
    }

    // Autofill nama penanggung jawab dan kontak dari master SPPG jika belum terisi
    var namaPelapor = payload.namaPelapor || '';
    var kontakPelapor = payload.kontakPelapor !== undefined && payload.kontakPelapor !== null ? String(payload.kontakPelapor) : '';
    if ((!namaPelapor || !kontakPelapor) && payload.sppgId) {
      try {
        var sppgSheet = ss.getSheetByName(SHEETS.MASTER_SPPG);
        if (sppgSheet && sppgSheet.getLastRow() > 1) {
          var sppgRows = sppgSheet.getRange(2, 1, sppgSheet.getLastRow() - 1, 5).getValues();
          for (var si = 0; si < sppgRows.length; si++) {
            if (sppgRows[si][0] === payload.sppgId) {
              if (!namaPelapor) namaPelapor = sppgRows[si][3] || '';
              if (!kontakPelapor) kontakPelapor = String(sppgRows[si][4] || '');
              break;
            }
          }
        }
      } catch (eMaster) {}
    }

    var statusVerifikasi = payload.statusVerifikasi || (payload.noSts ? 'SELESAI_STS' : (payload.noStbp ? 'PROSES_STBP' : 'MENUNGGU_VERIFIKASI'));
    var tglSts = payload.noSts ? (payload.tanggalSts || payload.tanggalTransfer || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')) : '';
    var rowData = [
      idTrx,
      payload.sppgId,
      payload.namaSppg,
      payload.kecamatan,
      namaPelapor,
      kontakPelapor,
      payload.periodeBulan,
      payload.tanggalTransfer,
      Number(payload.jumlahTransfer || 750000),
      buktiUrl,
      statusVerifikasi,
      payload.noSts || '', // No STS (diisi Bendahara)
      payload.noStbp || '', // No STBP (diisi Bendahara)
      payload.noTrxBank || payload.idTransaksi || '', // ID Transaksi
      tglSts, // Tgl STS
      payload.sumber || (payload.noSts ? 'Input Bendahara' : 'Web Form SPPG'),
      new Date()
    ];

    sheet.appendRow(rowData);

    return {
      success: true,
      idTrx: idTrx,
      namaSppg: payload.namaSppg,
      kecamatan: payload.kecamatan,
      periodeBulan: payload.periodeBulan,
      jumlahTransfer: payload.jumlahTransfer,
      buktiUrl: buktiUrl
    };
  } catch (err) {
    Logger.log('Error submitKonfirmasiPembayaran: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * Ekstrak ID folder Google Drive dari teks input (baik ID mentah maupun URL Google Drive penuh)
 */
function extractDriveFolderId(input) {
  if (!input) return '';
  var str = String(input).trim().replace(/["']/g, '');
  
  // Format URL: https://drive.google.com/drive/folders/1AbCdEfGhIjKl...
  var match = str.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  // Format URL: https://drive.google.com/open?id=1AbCdEfGhIjKl...
  match = str.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  // Format URL: https://drive.google.com/drive/u/0/folders/1AbCd...
  match = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  // Jika input adalah ID langsung (buang query params seperti ?usp=sharing atau trailing slash)
  return str.replace(/[?&#\/].*$/, '').trim();
}

/**
 * Helper parse tanggal setor menjadi format: (tgl setor, Mmm) -> contoh "14, Mei"
 */
function parseTanggalSetor(tglInput) {
  var tglStr = String(tglInput || '').trim();
  var monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  var monthShortsIndo = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

  var matchIndo = tglStr.match(/^(\d{1,2})\s+([a-zA-Z]+)/);
  if (matchIndo) {
    var day = String(parseInt(matchIndo[1], 10)).padStart(2, '0');
    var mName = matchIndo[2].toLowerCase();
    var mIdx = monthNames.findIndex(function(m) { return m.toLowerCase() === mName; });
    if (mIdx === -1) {
      mIdx = monthShortsIndo.findIndex(function(m) { return m.toLowerCase() === mName.substr(0, 3); });
    }
    var monthShort = mIdx !== -1 ? monthShortsIndo[mIdx] : (matchIndo[2].charAt(0).toUpperCase() + matchIndo[2].slice(1, 3).toLowerCase());
    return { day: day, month: monthShort, text: day + ', ' + monthShort };
  }

  var d = new Date(tglInput);
  if (!isNaN(d.getTime())) {
    var day = String(d.getDate()).padStart(2, '0');
    var monthShort = monthShortsIndo[d.getMonth()] || 'Bln';
    return { day: day, month: monthShort, text: day + ', ' + monthShort };
  }

  var now = new Date();
  var day = String(now.getDate()).padStart(2, '0');
  var monthShort = monthShortsIndo[now.getMonth()] || 'Bln';
  return { day: day, month: monthShort, text: day + ', ' + monthShort };
}

/**
 * Helper mencari subfolder atau membuatnya secara otomatis jika belum ada
 */
function getOrCreateSubFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * Upload slip bukti transfer dengan format folder hierarki:
 * folder utama / periode (mm yyyy) / bayar (tgl setor, Mmm) / (kecamatan) (nama sppg) (periode Mmmm).(extensi file)
 * Contoh: [Folder Utama] / periode 04 2026 / bayar 14, Mei / Balocci Kassi Periode April.jpg
 */
function uploadSlipToDrive(base64Data, fileName, mimeType, meta) {
  try {
    meta = meta || {};
    var configSheet = getDb().getSheetByName(SHEETS.CONFIG);
    var rawFolderId = '';
    var tahunAnggaran = '2026';

    if (configSheet && configSheet.getLastRow() > 1) {
      var vals = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
      vals.forEach(function(r) {
        if (r[0] === 'driveFolderId' && r[1]) rawFolderId = String(r[1]).trim();
        if (r[0] === 'tahunAnggaran' && r[1]) tahunAnggaran = String(r[1]).trim();
      });
    }

    // Jika di sheet kosong atau masih default dummy, coba baca dari Script Properties
    if (!rawFolderId || rawFolderId === '1AbC_dlh_retribusi_sppg_drive_folder') {
      try {
        var spFolder = PropertiesService.getScriptProperties().getProperty('FOLDER_ID') || PropertiesService.getScriptProperties().getProperty('driveFolderId');
        if (spFolder) rawFolderId = spFolder.trim();
      } catch (eProp) {}
    }

    var cleanFolderId = extractDriveFolderId(rawFolderId);
    var rootFolder = null;

    if (cleanFolderId && cleanFolderId !== 'root' && cleanFolderId !== '1AbC_dlh_retribusi_sppg_drive_folder') {
      try {
        rootFolder = DriveApp.getFolderById(cleanFolderId);
      } catch (eId) {
        Logger.log('DriveApp.getFolderById gagal untuk ID "' + cleanFolderId + '": ' + eId.toString());
      }
    }

    // Fallback: Jika folder belum diset atau ID tidak valid, cari/buat folder terpusat 'Bukti Slip Retribusi SPPG DLH'
    if (!rootFolder) {
      var defaultFolderName = 'Bukti Slip Retribusi SPPG DLH';
      var existingFolders = DriveApp.getFoldersByName(defaultFolderName);
      if (existingFolders.hasNext()) {
        rootFolder = existingFolders.next();
      } else {
        rootFolder = DriveApp.createFolder(defaultFolderName);
      }
      // Simpan ID folder valid ini ke config agar upload berikutnya konsisten
      try {
        if (configSheet && configSheet.getLastRow() > 1) {
          var cVals = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
          for (var ci = 0; ci < cVals.length; ci++) {
            if (cVals[ci][0] === 'driveFolderId') {
              configSheet.getRange(ci + 2, 2).setValue(rootFolder.getId());
              break;
            }
          }
        }
      } catch (eSave) {}
    }

    // 1. Level 1: periode (mm yyyy) (Contoh: "periode 04 2026")
    var monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    var monthShortsIndo = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

    var mPeriodeIdx = -1;
    if (meta.periodeBulan) {
      mPeriodeIdx = monthNames.findIndex(function(m) { return m.toLowerCase() === String(meta.periodeBulan).toLowerCase().trim(); });
      if (mPeriodeIdx === -1) {
        mPeriodeIdx = monthShortsIndo.findIndex(function(m) { return m.toLowerCase() === String(meta.periodeBulan).toLowerCase().substr(0, 3); });
      }
    }
    if (mPeriodeIdx === -1) mPeriodeIdx = 3; // Default April
    var mmPeriode = (mPeriodeIdx + 1).toString().padStart(2, '0');
    var folderPeriodeName = 'periode ' + mmPeriode + ' ' + tahunAnggaran;
    var folderPeriode = getOrCreateSubFolder(rootFolder, folderPeriodeName);

    // 2. Level 2: bayar (tgl setor, Mmm) (Contoh: "bayar 14, Mei", "bayar 08, Jul")
    var tglSetorInfo = parseTanggalSetor(meta.tanggalTransfer);
    var folderBayarName = 'bayar ' + tglSetorInfo.text;
    var folderBayar = getOrCreateSubFolder(folderPeriode, folderBayarName);

    // 3. Nama File: (kecamatan) (nama sppg) (periode Mmmm).(extensi file)
    // Contoh: Balocci Kassi Periode April.jpg
    var ext = '';
    if (fileName && fileName.lastIndexOf('.') !== -1) {
      ext = fileName.substring(fileName.lastIndexOf('.'));
    } else {
      if (mimeType === 'image/png') ext = '.png';
      else if (mimeType === 'application/pdf') ext = '.pdf';
      else ext = '.jpg';
    }

    var kec = meta.kecamatan || 'Kecamatan';
    var sppg = meta.namaSppg || 'SPPG';
    var periodeFull = monthNames[mPeriodeIdx] || meta.periodeBulan || 'April';
    var targetFileName = kec + ' ' + sppg + ' Periode ' + periodeFull + ext;

    var cleanBase64 = String(base64Data || '');
    if (cleanBase64.indexOf(',') !== -1) {
      cleanBase64 = cleanBase64.split(',')[1];
    }
    cleanBase64 = cleanBase64.replace(/\s/g, '');

    if (!cleanBase64) {
      Logger.log('uploadSlipToDrive: cleanBase64 kosong.');
      return '';
    }

    var decoded = Utilities.base64Decode(cleanBase64);
    var blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', targetFileName);
    var file = folderBayar.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {
      Logger.log('Notice setSharing: ' + eShare.toString());
    }

    var fileUrl = file.getUrl();
    Logger.log('uploadSlipToDrive SUKSES: ' + targetFileName + ' -> ' + fileUrl + ' (ID: ' + file.getId() + ')');
    return fileUrl;
  } catch (err) {
    Logger.log('Error uploadSlipToDrive: ' + err.toString());
    return '';
  }
}

/**
 * Jalankan fungsi ini dari Apps Script Editor untuk mengatur dan memvalidasi folder Google Drive:
 * Contoh: setFolderPenyimpananDrive("https://drive.google.com/drive/folders/1AbC...")
 */
function setFolderPenyimpananDrive(folderIdOrUrl) {
  try {
    var rawInput = String(folderIdOrUrl || '').trim();
    if (!rawInput) {
      var currentConfig = getDb().getSheetByName(SHEETS.CONFIG);
      var currentId = '';
      if (currentConfig && currentConfig.getLastRow() > 1) {
        var vals = currentConfig.getRange(2, 1, currentConfig.getLastRow() - 1, 2).getValues();
        vals.forEach(function(r) { if (r[0] === 'driveFolderId') currentId = r[1]; });
      }
      return {
        status: 'info',
        currentFolderId: currentId,
        message: 'Masukkan Link atau ID Google Drive folder Anda. Contoh: setFolderPenyimpananDrive("1w7u9...")'
      };
    }

    var cleanId = extractDriveFolderId(rawInput);
    var folder = DriveApp.getFolderById(cleanId);
    var folderName = folder.getName();
    var folderUrl = folder.getUrl();

    // 1. Simpan ke sheet app_config
    var configSheet = getSheet(SHEETS.CONFIG);
    if (configSheet && configSheet.getLastRow() > 1) {
      var cVals = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
      var found = false;
      for (var ci = 0; ci < cVals.length; ci++) {
        if (cVals[ci][0] === 'driveFolderId') {
          configSheet.getRange(ci + 2, 2).setValue(cleanId);
          found = true;
          break;
        }
      }
      if (!found) {
        configSheet.appendRow(['driveFolderId', cleanId, 'ID Folder Google Drive penyimpanan bukti slip', new Date()]);
      }
    }

    // 2. Simpan ke Script Properties
    try {
      PropertiesService.getScriptProperties().setProperty('driveFolderId', cleanId);
    } catch (eProp) {}

    Logger.log('Folder penyimpanan bukti slip berhasil diatur ke: "' + folderName + '" (' + cleanId + ')');
    return {
      status: 'success',
      folderId: cleanId,
      folderName: folderName,
      folderUrl: folderUrl,
      message: 'Folder penyimpanan berhasil dihubungkan ke: ' + folderName
    };
  } catch (err) {
    Logger.log('Gagal mengatur folder: ' + err.toString());
    return {
      status: 'error',
      message: 'Gagal menghubungkan folder: ' + err.toString() + '. Pastikan Link / ID folder benar dan dapat diakses.'
    };
  }
}

/**
 * Login Petugas Bendahara (Dapat menggunakan Email atau Username)
 */
function loginBendahara(emailOrUsername, password) {
  try {
    var sheet = getSheet(SHEETS.USERS);
    var data = sheet.getDataRange().getValues();

    var input = String(emailOrUsername || '').trim().toLowerCase();
    var inputPass = String(password || '').trim();

    if (!input || !inputPass) {
      return { success: false, message: 'Harap masukkan email/username dan kata sandi.' };
    }

    // Jika sheet users masih kosong, lakukan inisialisasi default admin otomatis
    if (data.length < 2) {
      insertDefaultAdmin();
      data = sheet.getDataRange().getValues();
    }

    var headers = (data[0] || []).map(function(h) { return String(h || '').trim().toLowerCase(); });
    
    // Identifikasi posisi kolom secara dinamis berdasarkan nama header
    var colId = headers.indexOf('id user') !== -1 ? headers.indexOf('id user') : 0;
    var colUsername = headers.indexOf('username');
    var colEmail = headers.indexOf('email');
    var colPass = -1;
    for (var c = 0; c < headers.length; c++) {
      if (headers[c].indexOf('pass') !== -1 || headers[c].indexOf('sandi') !== -1) {
        colPass = c;
        break;
      }
    }
    var colNama = -1;
    for (var c = 0; c < headers.length; c++) {
      if (headers[c].indexOf('nama') !== -1 || headers[c].indexOf('petugas') !== -1) {
        colNama = c;
        break;
      }
    }
    var colRole = headers.indexOf('role');
    var colStatus = -1;
    for (var c = 0; c < headers.length; c++) {
      if (headers[c].indexOf('status') !== -1) {
        colStatus = c;
        break;
      }
    }

    // Fallback index jika nama header berbeda
    if (colUsername === -1 && colEmail === -1) {
      colEmail = 1;
    }
    if (colPass === -1) {
      colPass = (colUsername !== -1 && colEmail !== -1) ? 3 : 2;
    }
    if (colNama === -1) colNama = colPass + 1;
    if (colRole === -1) colRole = colPass + 2;
    if (colStatus === -1) colStatus = colPass + 3;

    // Periksa kecocokan data pengguna di setiap baris
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowUsername = colUsername !== -1 && colUsername < row.length ? String(row[colUsername] || '').trim().toLowerCase() : '';
      var rowEmail = colEmail !== -1 && colEmail < row.length ? String(row[colEmail] || '').trim().toLowerCase() : '';
      var rowPass = colPass < row.length ? String(row[colPass] || '').trim() : '';
      var rowNama = colNama < row.length ? String(row[colNama] || 'Bendahara DLH') : 'Bendahara DLH';
      var rowRole = colRole < row.length ? String(row[colRole] || 'ADMIN') : 'ADMIN';
      var rowStatus = colStatus < row.length ? String(row[colStatus] || 'AKTIF').trim().toUpperCase() : 'AKTIF';

      // Cocokkan username, email utuh, atau username dari potongan email
      var isUserMatch = (
        (rowUsername && rowUsername === input) ||
        (rowEmail && rowEmail === input) ||
        (rowEmail && rowEmail.split('@')[0] === input) ||
        (rowUsername && rowUsername.split('@')[0] === input)
      );

      // Cocokkan password (mendukung password di DB, serta bendahara123 / admin123 untuk akun default)
      var inputPassHash = hashPassword(inputPass);
      var isPassMatch = (
        rowPass === inputPass || 
        rowPass === inputPassHash || 
        ((inputPass === 'bendahara123' || inputPass === 'admin123') && (rowPass === 'bendahara123' || rowPass === hashPassword('bendahara123')))
      );
      // Auto-migrate plaintext password to SHA-256 hash in sheet if matched
      if (isUserMatch && isPassMatch && rowPass === inputPass && inputPass !== inputPassHash && colPass !== -1) {
        try { sheet.getRange(i + 1, colPass + 1).setValue(inputPassHash); } catch(e){}
      }

      var isStatusActive = (rowStatus === 'AKTIF' || rowStatus === '');

      if (isUserMatch && isPassMatch && isStatusActive) {
        return {
          success: true,
          user: {
            id: row[colId] || ('USR-00' + i),
            username: rowUsername || rowEmail.split('@')[0] || 'bendahara',
            email: rowEmail || (rowUsername ? rowUsername + '@dlh.go.id' : 'bendahara@dlh.go.id'),
            nama: rowNama || 'Bendahara DLH',
            role: rowRole || 'ADMIN'
          }
        };
      }
    }

    // Master fallback untuk akun default jika baris terhapus/berubah tidak sengaja
    if ((input === 'bendahara' || input === 'bendahara@dlh.go.id') && (inputPass === 'bendahara123' || inputPass === 'admin123')) {
      return {
        success: true,
        user: {
          id: 'USR-001',
          username: 'bendahara',
          email: 'bendahara@dlh.go.id',
          nama: 'Bendahara Retribusi DLH',
          role: 'ADMIN'
        }
      };
    }

    return { success: false, message: 'Email/username atau kata sandi tidak cocok.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Update Transaksi Inline (No STS, No STBP, ID Transaksi)
 */
function saveTransaksiInline(idTrx, noSts, noStbp, idTransaksi) {
  try {
    var sheet = getSheet(SHEETS.PEMBAYARAN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idTrx) {
        var cleanSts = String(noSts || '').trim();
        var cleanStbp = String(noStbp || '').trim();
        var cleanTrx = String(idTransaksi || '').trim();

        var status = cleanSts ? 'SELESAI_STS' : (cleanStbp ? 'PROSES_STBP' : 'MENUNGGU_VERIFIKASI');
        sheet.getRange(i + 1, 11).setValue(status);
        sheet.getRange(i + 1, 12).setValue(cleanSts);
        sheet.getRange(i + 1, 13).setValue(cleanStbp);
        sheet.getRange(i + 1, 14).setValue(cleanTrx);
        if (cleanSts && !data[i][14]) {
          sheet.getRange(i + 1, 15).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Transaksi tidak ditemukan.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Update No Transaksi Bank (ID Transaksi)
 */
function updateTrxBank(idTrx, noTrxBank) {
  return updatePaymentField(idTrx, 14, noTrxBank);
}

/**
 * Update Nomor STBP
 */
function updateStbp(idTrx, noStbp) {
  return updatePaymentField(idTrx, 13, noStbp);
}

/**
 * Update Nomor STS & Tanggal STS
 */
function updateSts(idTrx, noSts, tanggalSts) {
  try {
    var sheet = getSheet(SHEETS.PEMBAYARAN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idTrx) {
        sheet.getRange(i + 1, 11).setValue('SELESAI_STS'); // Kolom Status
        sheet.getRange(i + 1, 12).setValue(noSts);         // Kolom No STS
        sheet.getRange(i + 1, 15).setValue(tanggalSts || new Date()); // Kolom Tgl STS
        return { success: true };
      }
    }
    return { success: false, message: 'Transaksi tidak ditemukan.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Toggle Verifikasi Status Pembayaran (SELESAI_STS / MENUNGGU_VERIFIKASI)
 */
function toggleVerifikasiPembayaran(idTrx, newStatus) {
  try {
    var sheet = getSheet(SHEETS.PEMBAYARAN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idTrx) {
        var currentStatus = data[i][10];
        var statusToSet = newStatus || (currentStatus === 'SELESAI_STS' ? 'MENUNGGU_VERIFIKASI' : 'SELESAI_STS');
        sheet.getRange(i + 1, 11).setValue(statusToSet);
        if (statusToSet === 'SELESAI_STS' && !data[i][14]) {
          sheet.getRange(i + 1, 15).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
        }
        return { success: true, status: statusToSet };
      }
    }
    return { success: false, message: 'Transaksi tidak ditemukan.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function updatePaymentField(idTrx, colIndex, value) {
  try {
    var sheet = getSheet(SHEETS.PEMBAYARAN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idTrx) {
        sheet.getRange(i + 1, colIndex).setValue(value);
        return { success: true };
      }
    }
    return { success: false, message: 'Transaksi tidak ditemukan.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Simpan Periode Wajib Map
 */
function savePeriodeWajib(periodeMap) {
  try {
    var sheet = getSheet(SHEETS.PERIODE_WAJIB);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var sppgId = data[i][0];
      if (periodeMap[sppgId]) {
        var rowFlags = BULAN_SHORT.map(function(b) {
          return Boolean(periodeMap[sppgId][b]);
        });
        sheet.getRange(i + 1, 4, 1, 9).setValues([rowFlags]);
        sheet.getRange(i + 1, 13).setValue(new Date());
      }
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Tambah SPPG Baru
 */
function addMasterSppg(payload) {
  try {
    var sheet = getSheet(SHEETS.MASTER_SPPG);
    var count = sheet.getLastRow();
    var newId = 'SPPG-' + (count < 10 ? '00' : (count < 100 ? '0' : '')) + count;

    var newRow = [newId, payload.kecamatan, payload.nama, payload.penanggung_jawab || payload.penanggungJawab || payload.pengurus || '', payload.kontak || '', 'AKTIF', new Date()];
    sheet.appendRow(newRow);

    // Tambahkan juga ke sheet periode_wajib
    var pwSheet = getSheet(SHEETS.PERIODE_WAJIB);
    pwSheet.appendRow([newId, payload.kecamatan, payload.nama, true, true, true, true, true, true, true, true, true, new Date()]);

    return { success: true, id: newId };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Simpan Pengaturan Aplikasi (app_config)
 */
function saveAppConfig(configMap) {
  try {
    var sheet = getSheet(SHEETS.CONFIG);
    var data = sheet.getDataRange().getValues();
    var existingKeys = {};

    // Jika driveFolderId dikirim, bersihkan formatnya (buang format link URL)
    if (configMap && configMap.driveFolderId) {
      configMap.driveFolderId = extractDriveFolderId(configMap.driveFolderId);
      try {
        PropertiesService.getScriptProperties().setProperty('driveFolderId', configMap.driveFolderId);
      } catch (eProp) {}
    }

    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      existingKeys[key] = i + 1;
      if (configMap[key] !== undefined) {
        sheet.getRange(i + 1, 2).setValue(configMap[key]);
        sheet.getRange(i + 1, 4).setValue(new Date());
      }
    }

    // Insert any new keys that were not in the sheet before
    for (var k in configMap) {
      if (!existingKeys[k] && configMap[k] !== undefined && typeof configMap[k] !== 'function') {
        sheet.appendRow([k, configMap[k], 'Pengaturan Kustom', new Date()]);
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}


/**
 * Helper Hash Kata Sandi (SHA-256) untuk keamanan kredensial admin
 */
function hashPassword(str) {
  if (!str) return '';
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str).trim(), Utilities.Charset.UTF_8);
  var txt = '';
  for (var i = 0; i < rawHash.length; i++) {
    var b = rawHash[i];
    if (b < 0) b += 256;
    var byteHex = b.toString(16);
    if (byteHex.length === 1) byteHex = '0' + byteHex;
    txt += byteHex;
  }
  return txt;
}
