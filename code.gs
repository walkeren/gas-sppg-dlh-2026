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
  var headers = ['ID SPPG', 'Kecamatan', 'Nama SPPG', 'Status Aktif', 'Dibuat Pada'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

function setupSheetUsers(ss) {
  var sheet = ss.getSheetByName(SHEETS.USERS) || ss.insertSheet(SHEETS.USERS);
  var headers = ['ID User', 'Username', 'Email', 'Password', 'Nama Petugas', 'Role', 'Status Aktif', 'Created At'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

function setupSheetConfig(ss) {
  var sheet = ss.getSheetByName(SHEETS.CONFIG) || ss.insertSheet(SHEETS.CONFIG);
  var headers = ['Key', 'Value', 'Deskripsi', 'Updated At'];
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
  var headers = ['ID SPPG', 'Kecamatan', 'Nama SPPG', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember', 'Updated At'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

function setupSheetPembayaran(ss) {
  var sheet = ss.getSheetByName(SHEETS.PEMBAYARAN) || ss.insertSheet(SHEETS.PEMBAYARAN);
  var headers = [
    'ID Transaksi', 
    'ID SPPG', 
    'Nama SPPG', 
    'Kecamatan', 
    'Nama Pelapor', 
    'Kontak WhatsApp', 
    'Periode Bulan', 
    'Tanggal Transfer', 
    'Jumlah Transfer', 
    'Bukti Slip URL', 
    'Status Dokumen', 
    'No STS', 
    'No STBP', 
    'ID Transaksi Bank', 
    'Tanggal STS', 
    'Catatan / Ref Excel', 
    'Created At'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow(sheet, headers.length);
}

// ============================================================================
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
  { id: 'SPPG-001', kec: 'Balocci', nama: 'Kassi' },
  { id: 'SPPG-002', kec: 'Balocci', nama: 'Kassi 2' },
  { id: 'SPPG-003', kec: 'Bungoro', nama: 'Samalewa 1' },
  { id: 'SPPG-004', kec: 'Bungoro', nama: 'Samalewa 2' },
  { id: 'SPPG-005', kec: 'Bungoro', nama: 'Samalewa 3' },
  { id: 'SPPG-006', kec: 'Bungoro', nama: 'Samalewa 4' },
  { id: 'SPPG-007', kec: 'Bungoro', nama: 'Samalewa 5' },
  { id: 'SPPG-008', kec: 'Labakkang', nama: 'Labakkang' },
  { id: 'SPPG-009', kec: 'Labakkang', nama: 'Labakkang 2' },
  { id: 'SPPG-010', kec: 'Labakkang', nama: 'Labakkang 3' },
  { id: 'SPPG-011', kec: 'Labakkang', nama: 'Manakku' },
  { id: 'SPPG-012', kec: 'Labakkang', nama: 'Mangallekana' },
  { id: 'SPPG-013', kec: 'Labakkang', nama: 'Batara' },
  { id: 'SPPG-014', kec: 'Mandalle', nama: 'Manggalung' },
  { id: 'SPPG-015', kec: 'Mandalle', nama: 'Tamarupa' },
  { id: 'SPPG-016', kec: 'Marang', nama: 'Talaka' },
  { id: 'SPPG-017', kec: 'Marang', nama: 'Talaka 2' },
  { id: 'SPPG-018', kec: 'Minasa Tene', nama: 'Bonto Langkasa' },
  { id: 'SPPG-019', kec: 'Minasa Tene', nama: 'Kabba' },
  { id: 'SPPG-020', kec: 'Minasa Tene', nama: 'Biraeng' },
  { id: 'SPPG-021', kec: 'Pangkajene', nama: 'Bonto Perak 1' },
  { id: 'SPPG-022', kec: 'Pangkajene', nama: 'Bonto Perak 2' },
  { id: 'SPPG-023', kec: 'Pangkajene', nama: 'Mappasaile' },
  { id: 'SPPG-024', kec: 'Pangkajene', nama: 'Mappasaile 2' },
  { id: 'SPPG-025', kec: 'Pangkajene', nama: 'Padoang Doangan' },
  { id: 'SPPG-026', kec: 'Pangkajene', nama: 'Tumampua' },
  { id: 'SPPG-027', kec: 'Segeri', nama: 'Bone' },
  { id: 'SPPG-028', kec: 'Segeri', nama: 'Segeri' },
  { id: 'SPPG-029', kec: 'Segeri', nama: 'Segeri 2' },
  { id: 'SPPG-030', kec: 'Tondong Tallasa', nama: 'Bantimurung' }
];

function insertMasterSppg() {
  var sheet = getSheet(SHEETS.MASTER_SPPG);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    Logger.log('Data master SPPG sudah ada.');
    return;
  }

  var rows = MASTER_SPPG_LIST.map(function(item) {
    return [item.id, item.kec, item.nama, 'AKTIF', new Date()];
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
      var sppgVals = sppgSheet.getRange(2, 1, sppgSheet.getLastRow() - 1, 4).getValues();
      masterSppg = sppgVals.map(function(r) {
        return { id: r[0], kec: r[1], nama: r[2], active: r[3] === 'AKTIF' };
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
          kontakPelapor: r[5],
          periodeBulan: r[6],
          tanggalTransfer: r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[7]),
          jumlahTransfer: Number(r[8]),
          buktiUrl: r[9],
          status: r[10],
          noSts: r[11],
          noStbp: r[12],
          noTrxBank: r[13],
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
    var buktiUrl = payload.buktiUrl || '';
    if (payload.fileBase64 && payload.fileName) {
      buktiUrl = uploadSlipToDrive(payload.fileBase64, payload.fileName, payload.fileType, {
        idTrx: idTrx,
        kecamatan: payload.kecamatan,
        namaSppg: payload.namaSppg,
        periodeBulan: payload.periodeBulan,
        tanggalTransfer: payload.tanggalTransfer
      });
    }

    var rowData = [
      idTrx,
      payload.sppgId,
      payload.namaSppg,
      payload.kecamatan,
      payload.namaPelapor,
      payload.kontakPelapor,
      payload.periodeBulan,
      payload.tanggalTransfer,
      Number(payload.jumlahTransfer),
      buktiUrl,
      'MENUNGGU_VERIFIKASI',
      '', // No STS (diisi Bendahara)
      '', // No STBP (diisi Bendahara)
      payload.noTrxBank || payload.idTransaksi || '', // ID Transaksi
      '', // Tgl STS
      'Web Form SPPG',
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
 * [Folder Utama] / periode [mm yyyy] / bulan bayar [Bayar Mmm] / [Kecamatan] [Nama SPPG ] [Periode 'Month'].[extensi file]
 */
function uploadSlipToDrive(base64Data, fileName, mimeType, meta) {
  try {
    meta = meta || {};
    var configSheet = getDb().getSheetByName(SHEETS.CONFIG);
    var folderId = 'root';
    var tahunAnggaran = '2026';
    if (configSheet && configSheet.getLastRow() > 1) {
      var vals = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
      vals.forEach(function(r) {
        if (r[0] === 'driveFolderId' && r[1]) folderId = String(r[1]).trim();
        if (r[0] === 'tahunAnggaran' && r[1]) tahunAnggaran = String(r[1]).trim();
      });
    }

    var rootFolder;
    try {
      rootFolder = DriveApp.getFolderById(folderId);
    } catch (e) {
      rootFolder = DriveApp.getRootFolder();
    }

    // 1. Level 1: periode [mm yyyy] (Contoh: "periode 04 2026")
    var monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    var monthShortsIndo = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];

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

    // 2. Level 2: bulan bayar [Bayar Mmm] (Contoh: "bulan bayar Bayar Mei", "bulan bayar Bayar Jul")
    var payDate = meta.tanggalTransfer ? new Date(meta.tanggalTransfer) : new Date();
    if (isNaN(payDate.getTime())) payDate = new Date();
    var payMonthIdx = payDate.getMonth();
    var mmmBayar = monthShortsIndo[payMonthIdx] || 'Bln';
    var folderBulanBayarName = 'bulan bayar Bayar ' + mmmBayar;
    var folderBulanBayar = getOrCreateSubFolder(folderPeriode, folderBulanBayarName);

    // 3. Nama File: [Kecamatan] [Nama SPPG ] [Periode 'Month'].[extensi file]
    // Contoh: Balocci Kassi Periode 'April'.jpg
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
    var periodeFull = monthNames[mPeriodeIdx] || meta.periodeBulan || 'Bulan';
    var targetFileName = kec + ' ' + sppg + ' Periode \'' + periodeFull + '\'' + ext;

    var cleanBase64 = base64Data.split(',')[1] || base64Data;
    var decoded = Utilities.base64Decode(cleanBase64);
    var blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', targetFileName);
    var file = folderBulanBayar.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (err) {
    Logger.log('Error uploadSlipToDrive: ' + err.toString());
    return 'https://drive.google.com/thumbnail?id=demo_slip';
  }
}

/**
 * Login Petugas Bendahara (Dapat menggunakan Email atau Username)
 */
function loginBendahara(emailOrUsername, password) {
  try {
    var sheet = getSheet(SHEETS.USERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: 'Data pengguna tidak ditemukan.' };

    var headers = data[0];
    var hasUsernameCol = (headers[1] === 'Username');

    var input = String(emailOrUsername || '').trim().toLowerCase();
    var inputPass = String(password || '').trim();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowUsername = hasUsernameCol ? String(row[1] || '').trim().toLowerCase() : String(row[1] || '').split('@')[0].toLowerCase();
      var rowEmail = hasUsernameCol ? String(row[2] || '').trim().toLowerCase() : String(row[1] || '').trim().toLowerCase();
      var rowPass = hasUsernameCol ? String(row[3] || '').trim() : String(row[2] || '').trim();
      var rowNama = hasUsernameCol ? String(row[4] || '') : String(row[3] || '');
      var rowRole = hasUsernameCol ? String(row[5] || 'ADMIN') : String(row[4] || 'ADMIN');
      var rowStatus = hasUsernameCol ? String(row[6] || '').trim().toUpperCase() : String(row[5] || '').trim().toUpperCase();

      var isUserMatch = (rowUsername === input || rowEmail === input || rowEmail.split('@')[0] === input);
      if (isUserMatch && rowPass === inputPass && (rowStatus === 'AKTIF' || rowStatus === '')) {
        return {
          success: true,
          user: {
            id: row[0],
            username: rowUsername,
            email: rowEmail,
            nama: rowNama || 'Bendahara DLH',
            role: rowRole || 'ADMIN'
          }
        };
      }
    }
    return { success: false, message: 'Username/email atau kata sandi tidak cocok.' };
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

    var newRow = [newId, payload.kecamatan, payload.nama, 'AKTIF', new Date()];
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
