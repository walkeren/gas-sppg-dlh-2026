/**
 * ============================================================================
 * SISTEM PELAPORAN & REKAPITULASI PEMBAYARAN RETRIBUSI SPPG
 * DINAS LINGKUNGAN HIDUP (DLH) KABUPATEN PANGKAJENE DAN KEPULAUAN
 * ============================================================================
 * Stack: Google Apps Script (GAS) + Google Sheets + Google Drive + HTML Service
 * File: importdata.gs (Modul Khusus Import Data Transaksi Historis Excel)
 * ============================================================================
 */

/**
 * Data historis transaksi hasil ekstraksi dari file Excel (transaksi.xlsx)
 * beserta cell comments (Nomor STS, STBP, ID Transaksi / Ref Bank, Tanggal).
 */
var RAW_EXCEL_TRANSAKSI_DATA = [
  {"sppg_id": "SPPG-001", "nama": "Kassi", "kec": "Balocci", "bulan": "April", "tgl": "14 Mei 2026", "ref": "4635", "sts": "034", "stbp": "0084", "nom": 750000},
  {"sppg_id": "SPPG-001", "nama": "Kassi", "kec": "Balocci", "bulan": "Mei", "tgl": "08 Juli 2026", "ref": "5870", "sts": "084", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-001", "nama": "Kassi", "kec": "Balocci", "bulan": "Juni", "tgl": "28 Juli 2026", "ref": "8350", "sts": "116", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-002", "nama": "Kassi 2", "kec": "Balocci", "bulan": "April", "tgl": "14 Mei 2026", "ref": "8618", "sts": "035", "stbp": "0085", "nom": 750000},
  {"sppg_id": "SPPG-002", "nama": "Kassi 2", "kec": "Balocci", "bulan": "Mei", "tgl": "10 Juli 2026", "ref": "7839", "sts": "087", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-002", "nama": "Kassi 2", "kec": "Balocci", "bulan": "Juni", "tgl": "10 Ags 2026", "ref": "6417", "sts": "131", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-003", "nama": "Samalewa 1", "kec": "Bungoro", "bulan": "April", "tgl": "14 Mei 2026", "ref": "2626", "sts": "036", "stbp": "0086", "nom": 750000},
  {"sppg_id": "SPPG-003", "nama": "Samalewa 1", "kec": "Bungoro", "bulan": "Mei", "tgl": "2 JuLi 2026", "ref": "5122", "sts": "075", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-003", "nama": "Samalewa 1", "kec": "Bungoro", "bulan": "Juni", "tgl": "17 JuLi 2026", "ref": "0098", "sts": "103", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-004", "nama": "Samalewa 2", "kec": "Bungoro", "bulan": "April", "tgl": "14 Mei 2026", "ref": "3976", "sts": "037", "stbp": "0087", "nom": 750000},
  {"sppg_id": "SPPG-004", "nama": "Samalewa 2", "kec": "Bungoro", "bulan": "Mei", "tgl": "30 Juni 2026", "ref": "5456", "sts": "069", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-004", "nama": "Samalewa 2", "kec": "Bungoro", "bulan": "Juni", "tgl": "23 Juli 2026", "ref": "4717", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-005", "nama": "Samalewa 3", "kec": "Bungoro", "bulan": "April", "tgl": "16 Mei 2026", "ref": "0518", "sts": "041", "stbp": "0091", "nom": 750000},
  {"sppg_id": "SPPG-005", "nama": "Samalewa 3", "kec": "Bungoro", "bulan": "Mei", "tgl": "25 Juni 2026", "ref": "1936", "sts": "065", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-005", "nama": "Samalewa 3", "kec": "Bungoro", "bulan": "Juni", "tgl": "15 Juni 2026", "ref": "0259", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-006", "nama": "Samalewa 4", "kec": "Bungoro", "bulan": "April", "tgl": "19 Mei 2026", "ref": "1172", "sts": "045", "stbp": "", "nom": 743500},
  {"sppg_id": "SPPG-006", "nama": "Samalewa 4", "kec": "Bungoro", "bulan": "Mei", "tgl": "19 Juni 2026", "ref": "6424", "sts": "062", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-006", "nama": "Samalewa 4", "kec": "Bungoro", "bulan": "Juni", "tgl": "11 Juli 2026", "ref": "2527", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-007", "nama": "Samalewa 5", "kec": "Bungoro", "bulan": "April", "tgl": "03 Juni 2026", "ref": "2970", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-007", "nama": "Samalewa 5", "kec": "Bungoro", "bulan": "Mei", "tgl": "19 Juni 2026", "ref": "6887", "sts": "063", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-007", "nama": "Samalewa 5", "kec": "Bungoro", "bulan": "Juni", "tgl": "13 Juli 2026", "ref": "5266", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-008", "nama": "Labakkang", "kec": "Labakkang", "bulan": "April", "tgl": "16 Mei 2026", "ref": "1570", "sts": "042", "stbp": "0092", "nom": 750000},
  {"sppg_id": "SPPG-008", "nama": "Labakkang", "kec": "Labakkang", "bulan": "Mei", "tgl": "14 Juli 2026", "ref": "8472", "sts": "093", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-008", "nama": "Labakkang", "kec": "Labakkang", "bulan": "Juni", "tgl": "21 Jul 2026", "ref": "2419", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-008", "nama": "Labakkang", "kec": "Labakkang", "bulan": "Juli", "tgl": "21/8/2026", "ref": "1226", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-009", "nama": "Labakkang 2", "kec": "Labakkang", "bulan": "April", "tgl": "15 Mei 2026", "ref": "5299", "sts": "039", "stbp": "0089", "nom": 750000},
  {"sppg_id": "SPPG-009", "nama": "Labakkang 2", "kec": "Labakkang", "bulan": "Mei", "tgl": "30 Jul 2026", "ref": "2536", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-009", "nama": "Labakkang 2", "kec": "Labakkang", "bulan": "Juni", "tgl": "30 Jul 2026", "ref": "2537", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-010", "nama": "Labakkang 3", "kec": "Labakkang", "bulan": "April", "tgl": "29 Mei 2026 / 03 Juni 2026", "ref": "6780", "sts": "050", "stbp": "", "nom": 747500},
  {"sppg_id": "SPPG-010", "nama": "Labakkang 3", "kec": "Labakkang", "bulan": "Mei", "tgl": "17 Mei 2026", "ref": "4062", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-010", "nama": "Labakkang 3", "kec": "Labakkang", "bulan": "Juni", "tgl": "27 Juli 2026", "ref": "6240", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-010", "nama": "Labakkang 3", "kec": "Labakkang", "bulan": "Juli", "tgl": "27 Agus 2026", "ref": "7567", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-011", "nama": "Manakku", "kec": "Labakkang", "bulan": "April", "tgl": "13 Mei 2026", "ref": "6534", "sts": "027", "stbp": "0077", "nom": 750000},
  {"sppg_id": "SPPG-011", "nama": "Manakku", "kec": "Labakkang", "bulan": "Mei", "tgl": "08 Juli 2026", "ref": "5847", "sts": "082", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-011", "nama": "Manakku", "kec": "Labakkang", "bulan": "Juni", "tgl": "14 Juli 2026", "ref": "9153", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-012", "nama": "Mangallekana", "kec": "Labakkang", "bulan": "April", "tgl": "15 Mei 2026", "ref": "1973", "sts": "040", "stbp": "0090", "nom": 750000},
  {"sppg_id": "SPPG-012", "nama": "Mangallekana", "kec": "Labakkang", "bulan": "Mei", "tgl": "29 Juni 2026", "ref": "4680", "sts": "066", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-012", "nama": "Mangallekana", "kec": "Labakkang", "bulan": "Juni", "tgl": "27 Juli 2026", "ref": "3233", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-013", "nama": "Batara", "kec": "Labakkang", "bulan": "Mei", "tgl": "25 Juni 2026", "ref": "3870", "sts": "064", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-014", "nama": "Manggalung", "kec": "Mandalle", "bulan": "April", "tgl": "16 Mei 2026", "ref": "2076", "sts": "043", "stbp": "0093", "nom": 750000},
  {"sppg_id": "SPPG-014", "nama": "Manggalung", "kec": "Mandalle", "bulan": "Mei", "tgl": "07 Juli 2026", "ref": "", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-014", "nama": "Manggalung", "kec": "Mandalle", "bulan": "Juni", "tgl": "20 Juli 2026", "ref": "1010", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-015", "nama": "Tamarupa", "kec": "Mandalle", "bulan": "April", "tgl": "13 Mei 2026", "ref": "8308", "sts": "028", "stbp": "0078", "nom": 750000},
  {"sppg_id": "SPPG-015", "nama": "Tamarupa", "kec": "Mandalle", "bulan": "Mei", "tgl": "07 Juli 2026", "ref": "0991", "sts": "078", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-015", "nama": "Tamarupa", "kec": "Mandalle", "bulan": "Juni", "tgl": "16 Juli 2026", "ref": "8920", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-016", "nama": "Talaka", "kec": "Marang", "bulan": "April", "tgl": "18 Mei 2026", "ref": "", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-016", "nama": "Talaka", "kec": "Marang", "bulan": "Mei", "tgl": "16 Juli 2026", "ref": "1404", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-016", "nama": "Talaka", "kec": "Marang", "bulan": "Juni", "tgl": "16 Juli 2025", "ref": "9649", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-017", "nama": "Talaka 2", "kec": "Marang", "bulan": "April", "tgl": "13 Mei 2026", "ref": "1130", "sts": "029", "stbp": "0079", "nom": 750000},
  {"sppg_id": "SPPG-017", "nama": "Talaka 2", "kec": "Marang", "bulan": "Mei", "tgl": "05 JuLi 2026", "ref": "0724", "sts": "076", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-017", "nama": "Talaka 2", "kec": "Marang", "bulan": "Juni", "tgl": "24 Juli 2026", "ref": "2517", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-018", "nama": "Bonto Langkasa", "kec": "Minasa Tene", "bulan": "April", "tgl": "15 Mei 2026", "ref": "2679", "sts": "051", "stbp": "0099", "nom": 750000},
  {"sppg_id": "SPPG-018", "nama": "Bonto Langkasa", "kec": "Minasa Tene", "bulan": "Mei", "tgl": "", "ref": "", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-018", "nama": "Bonto Langkasa", "kec": "Minasa Tene", "bulan": "Juni", "tgl": "22 Juli 2026", "ref": "4561", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-019", "nama": "Kabba", "kec": "Minasa Tene", "bulan": "April", "tgl": "14 Mei 2026", "ref": "6036", "sts": "038", "stbp": "0088", "nom": 750000},
  {"sppg_id": "SPPG-019", "nama": "Kabba", "kec": "Minasa Tene", "bulan": "Mei", "tgl": "06 JuLi 2026", "ref": "4598", "sts": "077", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-019", "nama": "Kabba", "kec": "Minasa Tene", "bulan": "Juni", "tgl": "31 Juli 2026", "ref": "2407", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-021", "nama": "Bonto Perak 1", "kec": "Pangkajene", "bulan": "April", "tgl": "12 Mei 2026", "ref": "1100", "sts": "025", "stbp": "0075", "nom": 750000},
  {"sppg_id": "SPPG-021", "nama": "Bonto Perak 1", "kec": "Pangkajene", "bulan": "Mei", "tgl": "30 Juni 2026", "ref": "6903", "sts": "068", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-021", "nama": "Bonto Perak 1", "kec": "Pangkajene", "bulan": "Juni", "tgl": "10 Juli 2026", "ref": "4993", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-021", "nama": "Bonto Perak 1", "kec": "Pangkajene", "bulan": "Juli", "tgl": "20 JuLi 2026", "ref": "3912", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-022", "nama": "Bonto Perak 2", "kec": "Pangkajene", "bulan": "April", "tgl": "13 Mei 2026", "ref": "9043", "sts": "030", "stbp": "0080", "nom": 750000},
  {"sppg_id": "SPPG-022", "nama": "Bonto Perak 2", "kec": "Pangkajene", "bulan": "Mei", "tgl": "10 Juli 2026", "ref": "4991", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-022", "nama": "Bonto Perak 2", "kec": "Pangkajene", "bulan": "Juni", "tgl": "20 Juli 2026", "ref": "3912", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-023", "nama": "Mappasaile", "kec": "Pangkajene", "bulan": "April", "tgl": "13 Mei 2025", "ref": "1053", "sts": "031", "stbp": "0081", "nom": 750000},
  {"sppg_id": "SPPG-023", "nama": "Mappasaile", "kec": "Pangkajene", "bulan": "Mei", "tgl": "30 Juni 2026", "ref": "2039", "sts": "067", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-023", "nama": "Mappasaile", "kec": "Pangkajene", "bulan": "Juni", "tgl": "29 Juli 2026", "ref": "0608", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-024", "nama": "Mappasaile 2", "kec": "Pangkajene", "bulan": "April", "tgl": "16 Mei 2026", "ref": "4284", "sts": "044", "stbp": "0094", "nom": 750000},
  {"sppg_id": "SPPG-024", "nama": "Mappasaile 2", "kec": "Pangkajene", "bulan": "Mei", "tgl": "30 Juni 2026", "ref": "5314", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-024", "nama": "Mappasaile 2", "kec": "Pangkajene", "bulan": "Juni", "tgl": "23 Juli 2026", "ref": "7977", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-025", "nama": "Padoang Doangan", "kec": "Pangkajene", "bulan": "April", "tgl": "13 Mei 2026", "ref": "7536", "sts": "032", "stbp": "0082", "nom": 750000},
  {"sppg_id": "SPPG-025", "nama": "Padoang Doangan", "kec": "Pangkajene", "bulan": "Mei", "tgl": "30 Juni 2026", "ref": "0979", "sts": "061", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-025", "nama": "Padoang Doangan", "kec": "Pangkajene", "bulan": "Juni", "tgl": "30 Juli 2026", "ref": "2948", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-026", "nama": "Tumampua", "kec": "Pangkajene", "bulan": "April", "tgl": "13 Mei 2026", "ref": "6525", "sts": "033", "stbp": "0083", "nom": 750000},
  {"sppg_id": "SPPG-026", "nama": "Tumampua", "kec": "Pangkajene", "bulan": "Mei", "tgl": "30 Juni 2026", "ref": "0652", "sts": "070", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-026", "nama": "Tumampua", "kec": "Pangkajene", "bulan": "Juni", "tgl": "24 Juli 2026", "ref": "6610", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-026", "nama": "Tumampua", "kec": "Pangkajene", "bulan": "Juli", "tgl": "28 Agus 2026", "ref": "2088", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-027", "nama": "Bone", "kec": "Segeri", "bulan": "April", "tgl": "13 Mei 2026", "ref": "0581", "sts": "026", "stbp": "0076", "nom": 750000},
  {"sppg_id": "SPPG-027", "nama": "Bone", "kec": "Segeri", "bulan": "Mei", "tgl": "08 Juli 2026", "ref": "8584", "sts": "081", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-027", "nama": "Bone", "kec": "Segeri", "bulan": "Juni", "tgl": "08 Juli 2026", "ref": "8585", "sts": "082", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-028", "nama": "Segeri", "kec": "Segeri", "bulan": "April", "tgl": "20 Mei 2026", "ref": "0764", "sts": "046", "stbp": "0096", "nom": 750000},
  {"sppg_id": "SPPG-028", "nama": "Segeri", "kec": "Segeri", "bulan": "Mei", "tgl": "21 Juli 2026", "ref": "2499", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-028", "nama": "Segeri", "kec": "Segeri", "bulan": "Juni", "tgl": "21 Juli 2026", "ref": "5234", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-028", "nama": "Segeri", "kec": "Segeri", "bulan": "Juli", "tgl": "24 Agus 2026", "ref": "8393", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-029", "nama": "Segeri 2", "kec": "Segeri", "bulan": "April", "tgl": "20 Mei 2026", "ref": "4752", "sts": "047", "stbp": "0097", "nom": 750000},
  {"sppg_id": "SPPG-029", "nama": "Segeri 2", "kec": "Segeri", "bulan": "Mei", "tgl": "", "ref": "", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-029", "nama": "Segeri 2", "kec": "Segeri", "bulan": "Juni", "tgl": "23 Juli 2026", "ref": "8367", "sts": "", "stbp": "", "nom": 750000},
  {"sppg_id": "SPPG-029", "nama": "Segeri 2", "kec": "Segeri", "bulan": "Juli", "tgl": "10 Agus 2026", "ref": "1256", "sts": "", "stbp": "", "nom": 750000}
];

/**
 * Memasukkan seluruh data transaksi dari file Excel ke sheet 'pembayaran_retribusi'.
 * Urutan kolom database:
 * 1: ID Transaksi
 * 2: ID SPPG
 * 3: Nama SPPG
 * 4: Kecamatan
 * 5: Nama Pelapor
 * 6: Kontak WhatsApp
 * 7: Periode Bulan
 * 8: Tanggal Transfer
 * 9: Jumlah Transfer
 * 10: Bukti Slip URL
 * 11: Status Dokumen
 * 12: No STS
 * 13: No STBP
 * 14: ID Transaksi Bank
 * 15: Tanggal STS
 * 16: Catatan / Ref Excel
 * 17: Created At
 */
function insertHistoricalTransactions() {
  var sheet = getSheet(SHEETS.PEMBAYARAN);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    Logger.log('Data transaksi pembayaran sudah ada di database (' + (data.length - 1) + ' baris).');
    return {
      status: 'exists',
      count: data.length - 1,
      message: 'Data transaksi pembayaran sudah ada (' + (data.length - 1) + ' baris). Jika ingin menimpa, gunakan resetAndImportTransaksiExcel().'
    };
  }

  var bulanOrder = ['April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

  var rows = RAW_EXCEL_TRANSAKSI_DATA.map(function(item, idx) {
    var bIdx = bulanOrder.indexOf(item.bulan);
    var mmStr = (bIdx !== -1 ? bIdx + 4 : 4).toString().padStart(2, '0');
    var idTrx = 'TRX-2026' + mmStr + '-' + (idx + 1).toString().padStart(4, '0');
    var status = item.sts ? 'SELESAI_STS' : (item.stbp ? 'PROSES_STBP' : 'MENUNGGU_VERIFIKASI');
    var dummySlip = 'https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80';
    var tglSts = item.sts ? item.tgl : '';

    return [
      idTrx,
      item.sppg_id,
      item.nama,
      item.kec,
      'Pengurus ' + item.nama,
      '0812-3456-' + (1000 + idx),
      item.bulan,
      item.tgl || '2026-05-15',
      item.nom || 750000,
      dummySlip,
      status,
      item.sts || '',
      item.stbp || '',
      item.ref || (10000 + idx).toString(),
      tglSts,
      'Import Excel (Cell Comment)',
      new Date()
    ];
  });

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log(rows.length + ' data transaksi historis dari Excel berhasil diimpor ke sheet pembayaran_retribusi.');
  return {
    status: 'success',
    count: rows.length,
    message: rows.length + ' data transaksi historis Excel berhasil diimpor ke database.'
  };
}

/**
 * FUNGSI UTAMA: Mengimpor data transaksi awal dari file Excel (transaksi.xlsx)
 * Jalankan fungsi ini secara mandiri kapan saja setelah setupDatabase() selesai.
 */
function importDataTransaksiExcel() {
  return insertHistoricalTransactions();
}

/**
 * Alias fungsi import
 */
function insertTransaksiAgustus() {
  return insertHistoricalTransactions();
}

/**
 * Utility opsional: Membersihkan baris data transaksi lama dan mengimpor ulang data historis Excel
 */
function resetAndImportTransaksiExcel() {
  var sheet = getSheet(SHEETS.PEMBAYARAN);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  return insertHistoricalTransactions();
}
