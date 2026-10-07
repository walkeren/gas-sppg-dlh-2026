/**
 * ============================================================================
 * SISTEM PELAPORAN & REKAPITULASI PEMBAYARAN RETRIBUSI SPPG
 * DINAS LINGKUNGAN HIDUP (DLH) KABUPATEN PANGKAJENE DAN KEPULAUAN
 * ============================================================================
 * Stack: Google Apps Script (GAS) + Google Sheets + Google Drive + HTML Service
 * File: importdata.gs (Data Transaksi Historis Excel transaksi.xlsx & Kontak VCF)
 * ============================================================================
 */

/**
 * Dataset Transaksi Terverifikasi (Total: 88 Transaksi)
 * Sumber: transaksi.xlsx (Status TRUE = SELESAI_STS) & All Contacts.vcf
 */
var HISTORICAL_TRANSACTIONS_DATA = [
  {
    "id_transaksi": "TRX-202604-0001",
    "id_sppg": "SPPG-001",
    "nama_sppg": "Kassi",
    "kecamatan": "Balocci",
    "nama_pelapor": "Siti Zuraima SPPG Kassi1",
    "kontak_pelapor": "089501737683",
    "periode_bulan": "April",
    "tanggal_transfer": "14 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "034",
    "no_stbp": "0084",
    "id_transaksi_bank": "4635",
    "tanggal_sts": "14 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0002",
    "id_sppg": "SPPG-001",
    "nama_sppg": "Kassi",
    "kecamatan": "Balocci",
    "nama_pelapor": "Siti Zuraima SPPG Kassi1",
    "kontak_pelapor": "089501737683",
    "periode_bulan": "Mei",
    "tanggal_transfer": "08 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "084",
    "no_stbp": "",
    "id_transaksi_bank": "5870",
    "tanggal_sts": "08 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0003",
    "id_sppg": "SPPG-001",
    "nama_sppg": "Kassi",
    "kecamatan": "Balocci",
    "nama_pelapor": "Siti Zuraima SPPG Kassi1",
    "kontak_pelapor": "089501737683",
    "periode_bulan": "Juni",
    "tanggal_transfer": "28 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "116",
    "no_stbp": "",
    "id_transaksi_bank": "8350",
    "tanggal_sts": "28 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0004",
    "id_sppg": "SPPG-002",
    "nama_sppg": "Kassi 2",
    "kecamatan": "Balocci",
    "nama_pelapor": "Darwan Muis SPPG Kassi2",
    "kontak_pelapor": "085393517300",
    "periode_bulan": "April",
    "tanggal_transfer": "14 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "035",
    "no_stbp": "0085",
    "id_transaksi_bank": "8618",
    "tanggal_sts": "14 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0005",
    "id_sppg": "SPPG-002",
    "nama_sppg": "Kassi 2",
    "kecamatan": "Balocci",
    "nama_pelapor": "Darwan Muis SPPG Kassi2",
    "kontak_pelapor": "085393517300",
    "periode_bulan": "Mei",
    "tanggal_transfer": "10 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "087",
    "no_stbp": "",
    "id_transaksi_bank": "7839",
    "tanggal_sts": "10 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0006",
    "id_sppg": "SPPG-002",
    "nama_sppg": "Kassi 2",
    "kecamatan": "Balocci",
    "nama_pelapor": "Darwan Muis SPPG Kassi2",
    "kontak_pelapor": "085393517300",
    "periode_bulan": "Juni",
    "tanggal_transfer": "10 Ags 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "131",
    "no_stbp": "",
    "id_transaksi_bank": "6417",
    "tanggal_sts": "10 Ags 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0007",
    "id_sppg": "SPPG-003",
    "nama_sppg": "Samalewa 1",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Haerul Fahresi SPPG Samalewa1",
    "kontak_pelapor": "087714145593",
    "periode_bulan": "April",
    "tanggal_transfer": "14 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "036",
    "no_stbp": "0086",
    "id_transaksi_bank": "2626",
    "tanggal_sts": "14 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0008",
    "id_sppg": "SPPG-003",
    "nama_sppg": "Samalewa 1",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Haerul Fahresi SPPG Samalewa1",
    "kontak_pelapor": "087714145593",
    "periode_bulan": "Mei",
    "tanggal_transfer": "2 JuLi 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "075",
    "no_stbp": "",
    "id_transaksi_bank": "5122",
    "tanggal_sts": "2 JuLi 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0009",
    "id_sppg": "SPPG-003",
    "nama_sppg": "Samalewa 1",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Haerul Fahresi SPPG Samalewa1",
    "kontak_pelapor": "087714145593",
    "periode_bulan": "Juni",
    "tanggal_transfer": "17 JuLi 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "103",
    "no_stbp": "",
    "id_transaksi_bank": "0098",
    "tanggal_sts": "17 JuLi 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0010",
    "id_sppg": "SPPG-004",
    "nama_sppg": "Samalewa 2",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Nurul Mutmainnah SPPG Samalewa 2",
    "kontak_pelapor": "085756416172",
    "periode_bulan": "April",
    "tanggal_transfer": "14 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "037",
    "no_stbp": "0087",
    "id_transaksi_bank": "3976",
    "tanggal_sts": "14 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0011",
    "id_sppg": "SPPG-004",
    "nama_sppg": "Samalewa 2",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Nurul Mutmainnah SPPG Samalewa 2",
    "kontak_pelapor": "085756416172",
    "periode_bulan": "Mei",
    "tanggal_transfer": "30 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "069",
    "no_stbp": "",
    "id_transaksi_bank": "5456",
    "tanggal_sts": "30 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0012",
    "id_sppg": "SPPG-004",
    "nama_sppg": "Samalewa 2",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Nurul Mutmainnah SPPG Samalewa 2",
    "kontak_pelapor": "085756416172",
    "periode_bulan": "Juni",
    "tanggal_transfer": "23 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "4717",
    "tanggal_sts": "23 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0013",
    "id_sppg": "SPPG-005",
    "nama_sppg": "Samalewa 3",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Wahyuddin SPPG Samalewa3",
    "kontak_pelapor": "085342286433",
    "periode_bulan": "April",
    "tanggal_transfer": "16 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "041",
    "no_stbp": "0091",
    "id_transaksi_bank": "0518",
    "tanggal_sts": "16 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0014",
    "id_sppg": "SPPG-005",
    "nama_sppg": "Samalewa 3",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Wahyuddin SPPG Samalewa3",
    "kontak_pelapor": "085342286433",
    "periode_bulan": "Mei",
    "tanggal_transfer": "25 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "065",
    "no_stbp": "",
    "id_transaksi_bank": "1936",
    "tanggal_sts": "25 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0015",
    "id_sppg": "SPPG-005",
    "nama_sppg": "Samalewa 3",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Wahyuddin SPPG Samalewa3",
    "kontak_pelapor": "085342286433",
    "periode_bulan": "Juni",
    "tanggal_transfer": "15 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "0259",
    "tanggal_sts": "15 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0016",
    "id_sppg": "SPPG-006",
    "nama_sppg": "Samalewa 4",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Musakkir SPPG Samalewa4",
    "kontak_pelapor": "082296461855",
    "periode_bulan": "April",
    "tanggal_transfer": "19 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "045",
    "no_stbp": "",
    "id_transaksi_bank": "1172",
    "tanggal_sts": "19 Mei 2026",
    "catatan": "743.500 (kurang 6.500)"
  },
  {
    "id_transaksi": "TRX-202605-0017",
    "id_sppg": "SPPG-006",
    "nama_sppg": "Samalewa 4",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Musakkir SPPG Samalewa4",
    "kontak_pelapor": "082296461855",
    "periode_bulan": "Mei",
    "tanggal_transfer": "19 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "062",
    "no_stbp": "",
    "id_transaksi_bank": "6424",
    "tanggal_sts": "19 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0018",
    "id_sppg": "SPPG-006",
    "nama_sppg": "Samalewa 4",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Musakkir SPPG Samalewa4",
    "kontak_pelapor": "082296461855",
    "periode_bulan": "Juni",
    "tanggal_transfer": "11 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2527",
    "tanggal_sts": "11 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0019",
    "id_sppg": "SPPG-007",
    "nama_sppg": "Samalewa 5",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Anugrahwan SPPG Samalewa 5",
    "kontak_pelapor": "081340075010",
    "periode_bulan": "April",
    "tanggal_transfer": "03 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2970",
    "tanggal_sts": "03 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0020",
    "id_sppg": "SPPG-007",
    "nama_sppg": "Samalewa 5",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Anugrahwan SPPG Samalewa 5",
    "kontak_pelapor": "081340075010",
    "periode_bulan": "Mei",
    "tanggal_transfer": "19 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "063",
    "no_stbp": "",
    "id_transaksi_bank": "6887",
    "tanggal_sts": "19 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0021",
    "id_sppg": "SPPG-007",
    "nama_sppg": "Samalewa 5",
    "kecamatan": "Bungoro",
    "nama_pelapor": "Anugrahwan SPPG Samalewa 5",
    "kontak_pelapor": "081340075010",
    "periode_bulan": "Juni",
    "tanggal_transfer": "13 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "5266",
    "tanggal_sts": "13 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0022",
    "id_sppg": "SPPG-008",
    "nama_sppg": "Labakkang",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Ardiansyah SPPG Labakkang",
    "kontak_pelapor": "085975073614",
    "periode_bulan": "April",
    "tanggal_transfer": "16 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "042",
    "no_stbp": "0092",
    "id_transaksi_bank": "1570",
    "tanggal_sts": "16 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0023",
    "id_sppg": "SPPG-008",
    "nama_sppg": "Labakkang",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Ardiansyah SPPG Labakkang",
    "kontak_pelapor": "085975073614",
    "periode_bulan": "Mei",
    "tanggal_transfer": "14 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "093",
    "no_stbp": "",
    "id_transaksi_bank": "8472",
    "tanggal_sts": "14 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0024",
    "id_sppg": "SPPG-008",
    "nama_sppg": "Labakkang",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Ardiansyah SPPG Labakkang",
    "kontak_pelapor": "085975073614",
    "periode_bulan": "Juni",
    "tanggal_transfer": "21 Jul 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2419",
    "tanggal_sts": "21 Jul 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202607-0025",
    "id_sppg": "SPPG-008",
    "nama_sppg": "Labakkang",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Ardiansyah SPPG Labakkang",
    "kontak_pelapor": "085975073614",
    "periode_bulan": "Juli",
    "tanggal_transfer": "21/8/2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "1226",
    "tanggal_sts": "21/8/2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0026",
    "id_sppg": "SPPG-009",
    "nama_sppg": "Labakkang 2",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Muammar SPPG Labakkang 2",
    "kontak_pelapor": "082349170976",
    "periode_bulan": "April",
    "tanggal_transfer": "15 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "039",
    "no_stbp": "0089",
    "id_transaksi_bank": "5299",
    "tanggal_sts": "15 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0027",
    "id_sppg": "SPPG-009",
    "nama_sppg": "Labakkang 2",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Muammar SPPG Labakkang 2",
    "kontak_pelapor": "082349170976",
    "periode_bulan": "Mei",
    "tanggal_transfer": "30 Jul 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2536",
    "tanggal_sts": "30 Jul 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0028",
    "id_sppg": "SPPG-009",
    "nama_sppg": "Labakkang 2",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Muammar SPPG Labakkang 2",
    "kontak_pelapor": "082349170976",
    "periode_bulan": "Juni",
    "tanggal_transfer": "30 Jul 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2537",
    "tanggal_sts": "30 Jul 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0029",
    "id_sppg": "SPPG-010",
    "nama_sppg": "Labakkang 3",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Heris SPPG Labakkang 3",
    "kontak_pelapor": "085343793941",
    "periode_bulan": "April",
    "tanggal_transfer": "29 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "050",
    "no_stbp": "",
    "id_transaksi_bank": "6780",
    "tanggal_sts": "29 Mei 2026",
    "catatan": "Bukti TF: 29/05/2026 rek.koran 03/06/2026 747.500 (kurang 2.500)"
  },
  {
    "id_transaksi": "TRX-202605-0030",
    "id_sppg": "SPPG-010",
    "nama_sppg": "Labakkang 3",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Heris SPPG Labakkang 3",
    "kontak_pelapor": "085343793941",
    "periode_bulan": "Mei",
    "tanggal_transfer": "17 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "4062",
    "tanggal_sts": "17 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0031",
    "id_sppg": "SPPG-010",
    "nama_sppg": "Labakkang 3",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Heris SPPG Labakkang 3",
    "kontak_pelapor": "085343793941",
    "periode_bulan": "Juni",
    "tanggal_transfer": "27 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "6240",
    "tanggal_sts": "27 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202607-0032",
    "id_sppg": "SPPG-010",
    "nama_sppg": "Labakkang 3",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Heris SPPG Labakkang 3",
    "kontak_pelapor": "085343793941",
    "periode_bulan": "Juli",
    "tanggal_transfer": "27 Agus 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "7567",
    "tanggal_sts": "27 Agus 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0033",
    "id_sppg": "SPPG-011",
    "nama_sppg": "Manakku",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Rabiullanda Kulsum SPPG Manakku",
    "kontak_pelapor": "085341370164",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "027",
    "no_stbp": "0077",
    "id_transaksi_bank": "6534",
    "tanggal_sts": "13 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0034",
    "id_sppg": "SPPG-011",
    "nama_sppg": "Manakku",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Rabiullanda Kulsum SPPG Manakku",
    "kontak_pelapor": "085341370164",
    "periode_bulan": "Mei",
    "tanggal_transfer": "08 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "082",
    "no_stbp": "",
    "id_transaksi_bank": "5847",
    "tanggal_sts": "08 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0035",
    "id_sppg": "SPPG-011",
    "nama_sppg": "Manakku",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Rabiullanda Kulsum SPPG Manakku",
    "kontak_pelapor": "085341370164",
    "periode_bulan": "Juni",
    "tanggal_transfer": "14 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "9153",
    "tanggal_sts": "14 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0036",
    "id_sppg": "SPPG-012",
    "nama_sppg": "Mangallekana",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Nur Chaerunnisa SPPG Mangallekana",
    "kontak_pelapor": "082315045264",
    "periode_bulan": "April",
    "tanggal_transfer": "15 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "040",
    "no_stbp": "0090",
    "id_transaksi_bank": "1973",
    "tanggal_sts": "15 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0037",
    "id_sppg": "SPPG-012",
    "nama_sppg": "Mangallekana",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Nur Chaerunnisa SPPG Mangallekana",
    "kontak_pelapor": "082315045264",
    "periode_bulan": "Mei",
    "tanggal_transfer": "29 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "066",
    "no_stbp": "",
    "id_transaksi_bank": "4680",
    "tanggal_sts": "29 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0038",
    "id_sppg": "SPPG-012",
    "nama_sppg": "Mangallekana",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Nur Chaerunnisa SPPG Mangallekana",
    "kontak_pelapor": "082315045264",
    "periode_bulan": "Juni",
    "tanggal_transfer": "27 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "3233",
    "tanggal_sts": "27 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0039",
    "id_sppg": "SPPG-013",
    "nama_sppg": "Batara",
    "kecamatan": "Labakkang",
    "nama_pelapor": "Fatma Sri Fatimah SPPG Batara",
    "kontak_pelapor": "088744874889",
    "periode_bulan": "Mei",
    "tanggal_transfer": "25 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "064",
    "no_stbp": "",
    "id_transaksi_bank": "3870",
    "tanggal_sts": "25 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0040",
    "id_sppg": "SPPG-014",
    "nama_sppg": "Manggalung",
    "kecamatan": "Mandalle",
    "nama_pelapor": "Adi Irwandi SPPG Manggalung",
    "kontak_pelapor": "085242761351",
    "periode_bulan": "April",
    "tanggal_transfer": "16 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "043",
    "no_stbp": "0093",
    "id_transaksi_bank": "2076",
    "tanggal_sts": "16 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0041",
    "id_sppg": "SPPG-014",
    "nama_sppg": "Manggalung",
    "kecamatan": "Mandalle",
    "nama_pelapor": "Adi Irwandi SPPG Manggalung",
    "kontak_pelapor": "085242761351",
    "periode_bulan": "Mei",
    "tanggal_transfer": "07 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "",
    "tanggal_sts": "07 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0042",
    "id_sppg": "SPPG-014",
    "nama_sppg": "Manggalung",
    "kecamatan": "Mandalle",
    "nama_pelapor": "Adi Irwandi SPPG Manggalung",
    "kontak_pelapor": "085242761351",
    "periode_bulan": "Juni",
    "tanggal_transfer": "20 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "1010",
    "tanggal_sts": "20 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0043",
    "id_sppg": "SPPG-015",
    "nama_sppg": "Tamarupa",
    "kecamatan": "Mandalle",
    "nama_pelapor": "Agatha Febriandani SPPG Tamarupa",
    "kontak_pelapor": "085346164018",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "028",
    "no_stbp": "0078",
    "id_transaksi_bank": "8308",
    "tanggal_sts": "13 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0044",
    "id_sppg": "SPPG-015",
    "nama_sppg": "Tamarupa",
    "kecamatan": "Mandalle",
    "nama_pelapor": "Agatha Febriandani SPPG Tamarupa",
    "kontak_pelapor": "085346164018",
    "periode_bulan": "Mei",
    "tanggal_transfer": "07 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "078",
    "no_stbp": "",
    "id_transaksi_bank": "0991",
    "tanggal_sts": "07 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0045",
    "id_sppg": "SPPG-015",
    "nama_sppg": "Tamarupa",
    "kecamatan": "Mandalle",
    "nama_pelapor": "Agatha Febriandani SPPG Tamarupa",
    "kontak_pelapor": "085346164018",
    "periode_bulan": "Juni",
    "tanggal_transfer": "16 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "8920",
    "tanggal_sts": "16 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0046",
    "id_sppg": "SPPG-016",
    "nama_sppg": "Talaka",
    "kecamatan": "Marang",
    "nama_pelapor": "Awal Fajaruddin SPPG Talaka1",
    "kontak_pelapor": "089684444665",
    "periode_bulan": "April",
    "tanggal_transfer": "18 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "",
    "tanggal_sts": "18 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0047",
    "id_sppg": "SPPG-016",
    "nama_sppg": "Talaka",
    "kecamatan": "Marang",
    "nama_pelapor": "Awal Fajaruddin SPPG Talaka1",
    "kontak_pelapor": "089684444665",
    "periode_bulan": "Mei",
    "tanggal_transfer": "16 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "1404",
    "tanggal_sts": "16 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0048",
    "id_sppg": "SPPG-016",
    "nama_sppg": "Talaka",
    "kecamatan": "Marang",
    "nama_pelapor": "Awal Fajaruddin SPPG Talaka1",
    "kontak_pelapor": "089684444665",
    "periode_bulan": "Juni",
    "tanggal_transfer": "16 Juli 2025",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "9649",
    "tanggal_sts": "16 Juli 2025",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0049",
    "id_sppg": "SPPG-017",
    "nama_sppg": "Talaka 2",
    "kecamatan": "Marang",
    "nama_pelapor": "Nurul Afian SPPG Talaka2",
    "kontak_pelapor": "085397984559",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "029",
    "no_stbp": "0079",
    "id_transaksi_bank": "1130",
    "tanggal_sts": "13 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0050",
    "id_sppg": "SPPG-017",
    "nama_sppg": "Talaka 2",
    "kecamatan": "Marang",
    "nama_pelapor": "Nurul Afian SPPG Talaka2",
    "kontak_pelapor": "085397984559",
    "periode_bulan": "Mei",
    "tanggal_transfer": "05 JuLi 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "076",
    "no_stbp": "",
    "id_transaksi_bank": "0724",
    "tanggal_sts": "05 JuLi 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0051",
    "id_sppg": "SPPG-017",
    "nama_sppg": "Talaka 2",
    "kecamatan": "Marang",
    "nama_pelapor": "Nurul Afian SPPG Talaka2",
    "kontak_pelapor": "085397984559",
    "periode_bulan": "Juni",
    "tanggal_transfer": "24 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2517",
    "tanggal_sts": "24 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0052",
    "id_sppg": "SPPG-018",
    "nama_sppg": "Bonto Langkasa",
    "kecamatan": "Minasa Tene",
    "nama_pelapor": "Indah Putri Humairah SPPG Bonto Langkasa",
    "kontak_pelapor": "087858320557",
    "periode_bulan": "April",
    "tanggal_transfer": "15 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "051",
    "no_stbp": "0099",
    "id_transaksi_bank": "2679",
    "tanggal_sts": "15 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0053",
    "id_sppg": "SPPG-018",
    "nama_sppg": "Bonto Langkasa",
    "kecamatan": "Minasa Tene",
    "nama_pelapor": "Indah Putri Humairah SPPG Bonto Langkasa",
    "kontak_pelapor": "087858320557",
    "periode_bulan": "Mei",
    "tanggal_transfer": "15 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "",
    "tanggal_sts": "15 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0054",
    "id_sppg": "SPPG-018",
    "nama_sppg": "Bonto Langkasa",
    "kecamatan": "Minasa Tene",
    "nama_pelapor": "Indah Putri Humairah SPPG Bonto Langkasa",
    "kontak_pelapor": "087858320557",
    "periode_bulan": "Juni",
    "tanggal_transfer": "22 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "4561",
    "tanggal_sts": "22 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0055",
    "id_sppg": "SPPG-019",
    "nama_sppg": "Kabba",
    "kecamatan": "Minasa Tene",
    "nama_pelapor": "Fani Fajriani SPPG Kabba",
    "kontak_pelapor": "082318479896",
    "periode_bulan": "April",
    "tanggal_transfer": "14 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "038",
    "no_stbp": "0088",
    "id_transaksi_bank": "6036",
    "tanggal_sts": "14 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0056",
    "id_sppg": "SPPG-019",
    "nama_sppg": "Kabba",
    "kecamatan": "Minasa Tene",
    "nama_pelapor": "Fani Fajriani SPPG Kabba",
    "kontak_pelapor": "082318479896",
    "periode_bulan": "Mei",
    "tanggal_transfer": "06 JuLi 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "077",
    "no_stbp": "",
    "id_transaksi_bank": "4598",
    "tanggal_sts": "06 JuLi 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0057",
    "id_sppg": "SPPG-019",
    "nama_sppg": "Kabba",
    "kecamatan": "Minasa Tene",
    "nama_pelapor": "Fani Fajriani SPPG Kabba",
    "kontak_pelapor": "082318479896",
    "periode_bulan": "Juni",
    "tanggal_transfer": "31 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2407",
    "tanggal_sts": "31 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0058",
    "id_sppg": "SPPG-021",
    "nama_sppg": "Bonto Perak 1",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Fahmi Sofyan SPPG Bonto Perak1",
    "kontak_pelapor": "081244287373",
    "periode_bulan": "April",
    "tanggal_transfer": "12 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "025",
    "no_stbp": "0075",
    "id_transaksi_bank": "1100",
    "tanggal_sts": "12 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0059",
    "id_sppg": "SPPG-021",
    "nama_sppg": "Bonto Perak 1",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Fahmi Sofyan SPPG Bonto Perak1",
    "kontak_pelapor": "081244287373",
    "periode_bulan": "Mei",
    "tanggal_transfer": "30 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "068",
    "no_stbp": "",
    "id_transaksi_bank": "6903",
    "tanggal_sts": "30 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0060",
    "id_sppg": "SPPG-021",
    "nama_sppg": "Bonto Perak 1",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Fahmi Sofyan SPPG Bonto Perak1",
    "kontak_pelapor": "081244287373",
    "periode_bulan": "Juni",
    "tanggal_transfer": "10 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "4993",
    "tanggal_sts": "10 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202607-0061",
    "id_sppg": "SPPG-021",
    "nama_sppg": "Bonto Perak 1",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Fahmi Sofyan SPPG Bonto Perak1",
    "kontak_pelapor": "081244287373",
    "periode_bulan": "Juli",
    "tanggal_transfer": "20 JuLi 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "3912",
    "tanggal_sts": "20 JuLi 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0062",
    "id_sppg": "SPPG-022",
    "nama_sppg": "Bonto Perak 2",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Saharuddin SPPG Bonto Perak2",
    "kontak_pelapor": "085343693131",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "030",
    "no_stbp": "0080",
    "id_transaksi_bank": "9043",
    "tanggal_sts": "13 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0063",
    "id_sppg": "SPPG-022",
    "nama_sppg": "Bonto Perak 2",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Saharuddin SPPG Bonto Perak2",
    "kontak_pelapor": "085343693131",
    "periode_bulan": "Mei",
    "tanggal_transfer": "10 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "4991",
    "tanggal_sts": "10 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0064",
    "id_sppg": "SPPG-022",
    "nama_sppg": "Bonto Perak 2",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Saharuddin SPPG Bonto Perak2",
    "kontak_pelapor": "085343693131",
    "periode_bulan": "Juni",
    "tanggal_transfer": "20 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "3912",
    "tanggal_sts": "20 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0065",
    "id_sppg": "SPPG-023",
    "nama_sppg": "Mappasaile",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Evi Irviyanti SPPG Mappasaile",
    "kontak_pelapor": "085870552736",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2025",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "031",
    "no_stbp": "0081",
    "id_transaksi_bank": "1053",
    "tanggal_sts": "13 Mei 2025",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0066",
    "id_sppg": "SPPG-023",
    "nama_sppg": "Mappasaile",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Evi Irviyanti SPPG Mappasaile",
    "kontak_pelapor": "085870552736",
    "periode_bulan": "Mei",
    "tanggal_transfer": "30 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "067",
    "no_stbp": "",
    "id_transaksi_bank": "2039",
    "tanggal_sts": "30 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0067",
    "id_sppg": "SPPG-023",
    "nama_sppg": "Mappasaile",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Evi Irviyanti SPPG Mappasaile",
    "kontak_pelapor": "085870552736",
    "periode_bulan": "Juni",
    "tanggal_transfer": "29 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "0608",
    "tanggal_sts": "29 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0068",
    "id_sppg": "SPPG-024",
    "nama_sppg": "Mappasaile 2",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Akbar Tola SPPG Mappasaile 2",
    "kontak_pelapor": "085340312677",
    "periode_bulan": "April",
    "tanggal_transfer": "16 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "044",
    "no_stbp": "0094",
    "id_transaksi_bank": "4284",
    "tanggal_sts": "16 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0069",
    "id_sppg": "SPPG-024",
    "nama_sppg": "Mappasaile 2",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Akbar Tola SPPG Mappasaile 2",
    "kontak_pelapor": "085340312677",
    "periode_bulan": "Mei",
    "tanggal_transfer": "30 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "",
    "tanggal_sts": "30 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0070",
    "id_sppg": "SPPG-024",
    "nama_sppg": "Mappasaile 2",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Akbar Tola SPPG Mappasaile 2",
    "kontak_pelapor": "085340312677",
    "periode_bulan": "Juni",
    "tanggal_transfer": "23 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "7977",
    "tanggal_sts": "23 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0071",
    "id_sppg": "SPPG-025",
    "nama_sppg": "Padoang Doangan",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Amran SPPG Padoang2an",
    "kontak_pelapor": "085230740101",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "032",
    "no_stbp": "0082",
    "id_transaksi_bank": "7536",
    "tanggal_sts": "13 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0072",
    "id_sppg": "SPPG-025",
    "nama_sppg": "Padoang Doangan",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Amran SPPG Padoang2an",
    "kontak_pelapor": "085230740101",
    "periode_bulan": "Mei",
    "tanggal_transfer": "30 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "061",
    "no_stbp": "",
    "id_transaksi_bank": "0979",
    "tanggal_sts": "30 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0073",
    "id_sppg": "SPPG-025",
    "nama_sppg": "Padoang Doangan",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Amran SPPG Padoang2an",
    "kontak_pelapor": "085230740101",
    "periode_bulan": "Juni",
    "tanggal_transfer": "30 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2948",
    "tanggal_sts": "30 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0074",
    "id_sppg": "SPPG-026",
    "nama_sppg": "Tumampua",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Mar'atul Islam SPPG Tumampua",
    "kontak_pelapor": "085338318156",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "033",
    "no_stbp": "0083",
    "id_transaksi_bank": "6525",
    "tanggal_sts": "13 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0075",
    "id_sppg": "SPPG-026",
    "nama_sppg": "Tumampua",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Mar'atul Islam SPPG Tumampua",
    "kontak_pelapor": "085338318156",
    "periode_bulan": "Mei",
    "tanggal_transfer": "30 Juni 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "070",
    "no_stbp": "",
    "id_transaksi_bank": "0652",
    "tanggal_sts": "30 Juni 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0076",
    "id_sppg": "SPPG-026",
    "nama_sppg": "Tumampua",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Mar'atul Islam SPPG Tumampua",
    "kontak_pelapor": "085338318156",
    "periode_bulan": "Juni",
    "tanggal_transfer": "24 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "6610",
    "tanggal_sts": "24 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202607-0077",
    "id_sppg": "SPPG-026",
    "nama_sppg": "Tumampua",
    "kecamatan": "Pangkajene",
    "nama_pelapor": "Mar'atul Islam SPPG Tumampua",
    "kontak_pelapor": "085338318156",
    "periode_bulan": "Juli",
    "tanggal_transfer": "28 Agus 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2088",
    "tanggal_sts": "28 Agus 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0078",
    "id_sppg": "SPPG-027",
    "nama_sppg": "Bone",
    "kecamatan": "Segeri",
    "nama_pelapor": "Mawarni Utami SPPG Bone",
    "kontak_pelapor": "085346567732",
    "periode_bulan": "April",
    "tanggal_transfer": "13 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "026",
    "no_stbp": "0076",
    "id_transaksi_bank": "0581",
    "tanggal_sts": "13 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0079",
    "id_sppg": "SPPG-027",
    "nama_sppg": "Bone",
    "kecamatan": "Segeri",
    "nama_pelapor": "Mawarni Utami SPPG Bone",
    "kontak_pelapor": "085346567732",
    "periode_bulan": "Mei",
    "tanggal_transfer": "08 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "081",
    "no_stbp": "",
    "id_transaksi_bank": "8584",
    "tanggal_sts": "08 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0080",
    "id_sppg": "SPPG-027",
    "nama_sppg": "Bone",
    "kecamatan": "Segeri",
    "nama_pelapor": "Mawarni Utami SPPG Bone",
    "kontak_pelapor": "085346567732",
    "periode_bulan": "Juni",
    "tanggal_transfer": "08 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "082",
    "no_stbp": "",
    "id_transaksi_bank": "8585",
    "tanggal_sts": "08 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0081",
    "id_sppg": "SPPG-028",
    "nama_sppg": "Segeri",
    "kecamatan": "Segeri",
    "nama_pelapor": "Ashar Saputra SPPG Segeri",
    "kontak_pelapor": "085259870711",
    "periode_bulan": "April",
    "tanggal_transfer": "20 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "046",
    "no_stbp": "0096",
    "id_transaksi_bank": "0764",
    "tanggal_sts": "20 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0082",
    "id_sppg": "SPPG-028",
    "nama_sppg": "Segeri",
    "kecamatan": "Segeri",
    "nama_pelapor": "Ashar Saputra SPPG Segeri",
    "kontak_pelapor": "085259870711",
    "periode_bulan": "Mei",
    "tanggal_transfer": "21 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "2499",
    "tanggal_sts": "21 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0083",
    "id_sppg": "SPPG-028",
    "nama_sppg": "Segeri",
    "kecamatan": "Segeri",
    "nama_pelapor": "Ashar Saputra SPPG Segeri",
    "kontak_pelapor": "085259870711",
    "periode_bulan": "Juni",
    "tanggal_transfer": "21 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "5234",
    "tanggal_sts": "21 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202607-0084",
    "id_sppg": "SPPG-028",
    "nama_sppg": "Segeri",
    "kecamatan": "Segeri",
    "nama_pelapor": "Ashar Saputra SPPG Segeri",
    "kontak_pelapor": "085259870711",
    "periode_bulan": "Juli",
    "tanggal_transfer": "24 Agus 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "8393",
    "tanggal_sts": "24 Agus 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202604-0085",
    "id_sppg": "SPPG-029",
    "nama_sppg": "Segeri 2",
    "kecamatan": "Segeri",
    "nama_pelapor": "Amell Akuntan SPPG Segeri2",
    "kontak_pelapor": "082189423315",
    "periode_bulan": "April",
    "tanggal_transfer": "20 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "047",
    "no_stbp": "0097",
    "id_transaksi_bank": "4752",
    "tanggal_sts": "20 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202605-0086",
    "id_sppg": "SPPG-029",
    "nama_sppg": "Segeri 2",
    "kecamatan": "Segeri",
    "nama_pelapor": "Amell Akuntan SPPG Segeri2",
    "kontak_pelapor": "082189423315",
    "periode_bulan": "Mei",
    "tanggal_transfer": "15 Mei 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "",
    "tanggal_sts": "15 Mei 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202606-0087",
    "id_sppg": "SPPG-029",
    "nama_sppg": "Segeri 2",
    "kecamatan": "Segeri",
    "nama_pelapor": "Amell Akuntan SPPG Segeri2",
    "kontak_pelapor": "082189423315",
    "periode_bulan": "Juni",
    "tanggal_transfer": "23 Juli 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "8367",
    "tanggal_sts": "23 Juli 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  },
  {
    "id_transaksi": "TRX-202607-0088",
    "id_sppg": "SPPG-029",
    "nama_sppg": "Segeri 2",
    "kecamatan": "Segeri",
    "nama_pelapor": "Amell Akuntan SPPG Segeri2",
    "kontak_pelapor": "082189423315",
    "periode_bulan": "Juli",
    "tanggal_transfer": "10 Agus 2026",
    "jumlah_transfer": 750000,
    "bukti_slip_url": "https://images.unsplash.com/photo-1554224155-6726b3ff858f?w=600&auto=format&fit=crop&q=80",
    "status_verifikasi": "SELESAI_STS",
    "no_sts": "",
    "no_stbp": "",
    "id_transaksi_bank": "1256",
    "tanggal_sts": "10 Agus 2026",
    "catatan": "Import Excel transaksi.xlsx (Status: TRUE)"
  }
];

/**
 * Memasukkan seluruh 88 data transaksi historis ke sheet 'pembayaran_retribusi'
 * Mengikuti urutan snake_case header database standar:
 * [id_transaksi, id_sppg, nama_sppg, kecamatan, nama_pelapor, kontak_pelapor,
 *  periode_bulan, tanggal_transfer, jumlah_transfer, bukti_slip_url, status_verifikasi,
 *  no_sts, no_stbp, id_transaksi_bank, tanggal_sts, catatan, created_at]
 */
function insertHistoricalTransactions() {
  var sheet = getSheet(SHEETS.PEMBAYARAN);
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    Logger.log('Data transaksi pembayaran sudah ada di database (' + (data.length - 1) + ' baris).');
    return {
      status: 'exists',
      count: data.length - 1,
      message: 'Data transaksi pembayaran sudah ada (' + (data.length - 1) + ' baris). Gunakan resetAndImportTransaksiExcel() jika ingin mengimpor ulang.'
    };
  }

  var rows = HISTORICAL_TRANSACTIONS_DATA.map(function(item) {
    return [
      item.id_transaksi,
      item.id_sppg,
      item.nama_sppg,
      item.kecamatan,
      item.nama_pelapor,
      item.kontak_pelapor,
      item.periode_bulan,
      item.tanggal_transfer,
      item.jumlah_transfer,
      item.bukti_slip_url,
      item.status_verifikasi,
      item.no_sts,
      item.no_stbp,
      item.id_transaksi_bank,
      item.tanggal_sts,
      item.catatan,
      new Date()
    ];
  });

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log(rows.length + ' data transaksi historis berhasil diimpor ke sheet pembayaran_retribusi.');
  return {
    status: 'success',
    count: rows.length,
    message: rows.length + ' data transaksi historis Excel berhasil diimpor ke database.'
  };
}

/**
 * FUNGSI UTAMA: Mengimpor data transaksi Excel ke sheet pembayaran_retribusi
 */
function importDataTransaksiExcel() {
  return insertHistoricalTransactions();
}

/**
 * Alias import
 */
function insertTransaksiAgustus() {
  return insertHistoricalTransactions();
}

/**
 * Membersihkan isi tabel pembayaran_retribusi (kecuali baris header)
 */
function clearTabelTransaksi() {
  var sheet = getSheet(SHEETS.PEMBAYARAN);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  Logger.log('Tabel pembayaran_retribusi berhasil dikosongkan.');
  return { status: 'success', message: 'Tabel pembayaran_retribusi berhasil dikosongkan.' };
}

/**
 * Reset dan Import ulang dataset 88 transaksi historis Excel
 */
function resetAndImportTransaksiExcel() {
  clearTabelTransaksi();
  var res = insertHistoricalTransactions();
  updatePenanggungJawabDanKontakSppg();
  return res;
}

/**
 * ============================================================================
 * SINKRONISASI PENANGGUNG JAWAB & KONTAK SPPG KE SHEET master_sppg
 * ============================================================================
 * Menambahkan kolom 'penanggung_jawab' dan 'kontak' ke tabel master_sppg
 * lalu mencari data masing-masing penanggung jawab & kontak di sheet pembayaran_retribusi.
 */
function updatePenanggungJawabDanKontakSppg() {
  var ss = getDb();
  var sppgSheet = ss.getSheetByName(SHEETS.MASTER_SPPG);
  var paySheet = ss.getSheetByName(SHEETS.PEMBAYARAN);

  if (!sppgSheet) {
    Logger.log("Sheet master_sppg tidak ditemukan!");
    return { status: 'error', message: 'Sheet master_sppg tidak ditemukan.' };
  }

  // 1. Ekstrak data narahubung terbaru dari tabel pembayaran_retribusi
  var contactMap = {};
  if (paySheet && paySheet.getLastRow() > 1) {
    var payData = paySheet.getRange(2, 1, paySheet.getLastRow() - 1, 6).getValues();
    // Kolom pembayaran: [id_transaksi, id_sppg, nama_sppg, kecamatan, nama_pelapor, kontak_pelapor]
    payData.forEach(function(row) {
      var idSppg = String(row[1] || '').trim();
      var namaPelapor = String(row[4] || '').trim();
      var kontakPelapor = String(row[5] || '').trim();
      if (idSppg && (namaPelapor || kontakPelapor)) {
        contactMap[idSppg] = {
          penanggung_jawab: namaPelapor,
          kontak: kontakPelapor
        };
      }
    });
  }

  // Fallback dari HISTORICAL_TRANSACTIONS_DATA jika tabel pembayaran belum terisi
  if (Object.keys(contactMap).length === 0 && typeof HISTORICAL_TRANSACTIONS_DATA !== 'undefined') {
    HISTORICAL_TRANSACTIONS_DATA.forEach(function(item) {
      if (item.id_sppg && (item.nama_pelapor || item.kontak_pelapor)) {
        contactMap[item.id_sppg] = {
          penanggung_jawab: item.nama_pelapor,
          kontak: item.kontak_pelapor
        };
      }
    });
  }

  // Fallback kontak resmi untuk unit yang belum ada transaksi pembayaran
  if (!contactMap['SPPG-029']) {
    contactMap['SPPG-029'] = { penanggung_jawab: 'Amell Akuntan SPPG Segeri2', kontak: '082189423315' };
  }
  if (!contactMap['SPPG-030']) {
    contactMap['SPPG-030'] = { penanggung_jawab: 'Hauliah SPPG Bantimurung', kontak: '085240966666' };
  }

  // 2. Periksa & Perbarui Header master_sppg
  var lastRow = sppgSheet.getLastRow();
  var lastCol = Math.max(sppgSheet.getLastColumn(), 5);
  var headerValues = sppgSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var pjColIdx = headerValues.indexOf('penanggung_jawab') + 1;
  var kontakColIdx = headerValues.indexOf('kontak') + 1;

  if (pjColIdx === 0 || kontakColIdx === 0) {
    var newHeaders = ['id_sppg', 'kecamatan', 'nama_sppg', 'penanggung_jawab', 'kontak', 'status_aktif', 'created_at'];
    sppgSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    styleHeaderRow(sppgSheet, newHeaders.length);
    pjColIdx = 4;
    kontakColIdx = 5;
  }

  // 3. Isi nilai penanggung_jawab dan kontak untuk setiap baris SPPG
  var updatedCount = 0;
  if (lastRow > 1) {
    var numRows = lastRow - 1;
    var sppgIds = sppgSheet.getRange(2, 1, numRows, 1).getValues();
    var updateRange = sppgSheet.getRange(2, pjColIdx, numRows, 2);
    var newValues = [];

    for (var i = 0; i < numRows; i++) {
      var id = String(sppgIds[i][0] || '').trim();
      var info = contactMap[id] || { penanggung_jawab: '', kontak: '' };
      newValues.push([info.penanggung_jawab, info.kontak]);
      if (info.penanggung_jawab || info.kontak) updatedCount++;
    }

    updateRange.setValues(newValues);
    Logger.log('Berhasil mengisi ' + updatedCount + ' data penanggung jawab & kontak ke sheet master_sppg.');
  }

  return {
    status: 'success',
    updatedCount: updatedCount,
    totalRows: lastRow > 1 ? lastRow - 1 : 0,
    message: 'Kolom penanggungjawab dan kontak berhasil ditambahkan dan diisi dari tabel pembayaran.'
  };
}

/**
 * Alias fungsi untuk eksekusi fleksibel dari Google Apps Script Editor
 */
function syncPenanggungJawabDanKontakSppg() {
  return updatePenanggungJawabDanKontakSppg();
}

function isiPenanggungJawabDanKontak() {
  return updatePenanggungJawabDanKontakSppg();
}

/**
 * ============================================================================
 * SINKRONISASI & GENERATE BUKTI TRANSFER KE GOOGLE DRIVE
 * ============================================================================
 * Struktur Folder Sesuai Permintaan:
 * [Folder Utama] / periode (mm yyyy) / bayar (tgl setor, Mmm) / (kecamatan) (nama sppg) (periode Mmmm).(extensi file)
 * Contoh: [Folder Utama] / periode 04 2026 / bayar 14, Mei / Balocci Kassi Periode April.jpg
 * 
 * Jalankan fungsi ini dari Apps Script Editor:
 * - syncSemuaBuktiTransferKeDrive("LINK_ATAU_ID_FOLDER_DRIVE_ANDA")
 * atau jika folder sudah disimpan di Pengaturan, cukup:
 * - syncSemuaBuktiTransferKeDrive()
 */
function syncSemuaBuktiTransferKeDrive(folderIdOrUrl) {
  var ss = getDb();
  var paySheet = ss.getSheetByName(SHEETS.PEMBAYARAN);
  if (!paySheet || paySheet.getLastRow() < 2) {
    Logger.log("Tabel pembayaran kosong! Silakan jalankan resetAndImportTransaksiExcel() terlebih dahulu.");
    return { status: 'error', message: 'Tabel pembayaran_retribusi masih kosong. Jalankan resetAndImportTransaksiExcel() dahulu.' };
  }

  // 1. Tentukan Folder Utama
  var rawFolder = folderIdOrUrl;
  if (!rawFolder) {
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet && configSheet.getLastRow() > 1) {
      var cVals = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
      cVals.forEach(function(r) { if (r[0] === 'driveFolderId') rawFolder = r[1]; });
    }
  }
  if (!rawFolder) {
    try {
      rawFolder = PropertiesService.getScriptProperties().getProperty('driveFolderId');
    } catch (e) {}
  }

  var cleanFolderId = extractDriveFolderId(rawFolder);
  var rootFolder = null;
  if (cleanFolderId && cleanFolderId !== 'root' && cleanFolderId !== '1AbC_dlh_retribusi_sppg_drive_folder') {
    try {
      rootFolder = DriveApp.getFolderById(cleanFolderId);
    } catch (e) {
      Logger.log("Folder ID tidak ditemukan: " + e.toString());
    }
  }

  if (!rootFolder) {
    var defaultFolderName = 'Bukti Slip Retribusi SPPG DLH';
    var existingFolders = DriveApp.getFoldersByName(defaultFolderName);
    if (existingFolders.hasNext()) {
      rootFolder = existingFolders.next();
    } else {
      rootFolder = DriveApp.createFolder(defaultFolderName);
    }
    cleanFolderId = rootFolder.getId();
    setFolderPenyimpananDrive(cleanFolderId);
  }

  var monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  var monthShortsIndo = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

  var numRows = paySheet.getLastRow() - 1;
  var rows = paySheet.getRange(2, 1, numRows, 17).getValues();
  var updatedUrls = [];
  var filesCreatedCount = 0;

  for (var i = 0; i < numRows; i++) {
    var r = rows[i];
    var idTrx = r[0];
    var idSppg = r[1];
    var namaSppg = r[2];
    var kecamatan = r[3];
    var namaPelapor = r[4];
    var kontakPelapor = r[5];
    var periodeBulan = r[6];
    var tanggalTransfer = r[7];
    var nominal = r[8] || 750000;
    var currentUrl = r[9];
    var noSts = r[11];
    var noStbp = r[12];
    var idTrxBank = r[13];

    // Level 1: periode (mm yyyy)
    var mPeriodeIdx = monthNames.findIndex(function(m) { return m.toLowerCase() === String(periodeBulan).toLowerCase().trim(); });
    if (mPeriodeIdx === -1) {
      mPeriodeIdx = monthShortsIndo.findIndex(function(m) { return m.toLowerCase() === String(periodeBulan).toLowerCase().substr(0, 3); });
    }
    if (mPeriodeIdx === -1) mPeriodeIdx = 3;
    var mmPeriode = (mPeriodeIdx + 1).toString().padStart(2, '0');
    var folderPeriodeName = 'periode ' + mmPeriode + ' 2026';
    var folderPeriode = getOrCreateSubFolder(rootFolder, folderPeriodeName);

    // Level 2: bayar (tgl setor, Mmm)
    var tglSetorInfo = parseTanggalSetor(tanggalTransfer);
    var folderBayarName = 'bayar ' + tglSetorInfo.text;
    var folderBayar = getOrCreateSubFolder(folderPeriode, folderBayarName);

    // Level 3: (kecamatan) (nama sppg) (periode Mmmm).(extensi file)
    var targetFileName = kecamatan + ' ' + namaSppg + ' Periode ' + (monthNames[mPeriodeIdx] || periodeBulan) + '.jpg';

    // Cek apakah file sudah ada di folderBayar
    var existingFiles = folderBayar.getFilesByName(targetFileName);
    var fileUrl = '';
    if (existingFiles.hasNext()) {
      var existFile = existingFiles.next();
      fileUrl = existFile.getUrl();
    } else {
      // Buat file bukti digital resmi
      var svgContent = generateSlipSvgContent({
        id_transaksi: idTrx,
        id_sppg: idSppg,
        nama_sppg: namaSppg,
        kecamatan: kecamatan,
        periode_bulan: monthNames[mPeriodeIdx] || periodeBulan,
        tanggal_transfer: String(tanggalTransfer),
        jumlah_transfer: nominal,
        no_sts: noSts,
        no_stbp: noStbp,
        id_transaksi_bank: idTrxBank,
        nama_pelapor: namaPelapor,
        kontak_pelapor: kontakPelapor
      });

      var blob = Utilities.newBlob(svgContent, 'image/svg+xml', targetFileName);
      var newFile = folderBayar.createFile(blob);
      try {
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (eShare) {}
      fileUrl = newFile.getUrl();
      filesCreatedCount++;
    }

    updatedUrls.push([fileUrl]);
  }

  // Update kolom bukti_slip_url (Kolom J = 10)
  paySheet.getRange(2, 10, updatedUrls.length, 1).setValues(updatedUrls);

  Logger.log('Sinkronisasi selesai! ' + filesCreatedCount + ' file bukti transfer baru dibuat di folder "' + rootFolder.getName() + '"');
  return {
    status: 'success',
    folderName: rootFolder.getName(),
    folderUrl: rootFolder.getUrl(),
    filesCreated: filesCreatedCount,
    totalTransactions: numRows,
    message: 'Bukti transfer berhasil disinkronkan ke folder "' + rootFolder.getName() + '" (' + filesCreatedCount + ' file baru dibuat).'
  };
}

/**
 * Generator template SVG Bukti Slip Transfer Resmi DLH
 */
function generateSlipSvgContent(trx) {
  var sppg = trx.nama_sppg || 'SPPG';
  var kec = trx.kecamatan || 'Pangkep';
  var periode = trx.periode_bulan || 'April';
  var tgl = trx.tanggal_transfer || '-';
  var sts = trx.no_sts || '-';
  var stbp = trx.no_stbp || '-';
  var tranx = trx.id_transaksi_bank || '-';
  var penyetor = trx.nama_pelapor || 'Penanggung Jawab SPPG';
  var kontak = trx.kontak_pelapor || '-';

  return '<svg xmlns="http://www.w3.org/2000/svg" width="800" height="520" viewBox="0 0 800 520">' +
    '<defs>' +
    '<linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%">' +
    '<stop offset="0%" stop-color="#f8fafc"/>' +
    '<stop offset="100%" stop-color="#f1f5f9"/>' +
    '</linearGradient>' +
    '<linearGradient id="hdr" x1="0%" y1="0%" x2="100%" y2="0%">' +
    '<stop offset="0%" stop-color="#1e3a8a"/>' +
    '<stop offset="100%" stop-color="#047857"/>' +
    '</linearGradient>' +
    '</defs>' +
    '<rect width="800" height="520" fill="url(#bg)" rx="16" stroke="#cbd5e1" stroke-width="2"/>' +
    '<rect width="800" height="80" fill="url(#hdr)" rx="16 16 0 0"/>' +
    '<text x="400" y="32" fill="#ffffff" font-family="Arial, sans-serif" font-size="16" font-weight="bold" text-anchor="middle" letter-spacing="1">DINAS LINGKUNGAN HIDUP KABUPATEN PANGKAJENE DAN KEPULAUAN</text>' +
    '<text x="400" y="58" fill="#93c5fd" font-family="Arial, sans-serif" font-size="12" font-weight="bold" text-anchor="middle" letter-spacing="0.5">BUKTI PEMBAYARAN RETRIBUSI PERSAMPAHAN SPPG T.A. 2026</text>' +
    '<rect x="40" y="105" width="720" height="350" fill="#ffffff" rx="10" stroke="#e2e8f0" stroke-width="1"/>' +
    '<text x="70" y="145" fill="#64748b" font-family="Arial, sans-serif" font-size="13">Nama SPPG :</text>' +
    '<text x="210" y="145" fill="#0f172a" font-family="Arial, sans-serif" font-size="15" font-weight="bold">' + sppg + ' (Kec. ' + kec + ')</text>' +
    '<text x="70" y="180" fill="#64748b" font-family="Arial, sans-serif" font-size="13">Periode Pelayanan :</text>' +
    '<text x="210" y="180" fill="#047857" font-family="Arial, sans-serif" font-size="14" font-weight="bold">Bulan ' + periode + ' 2026</text>' +
    '<text x="70" y="215" fill="#64748b" font-family="Arial, sans-serif" font-size="13">Tanggal Setor :</text>' +
    '<text x="210" y="215" fill="#0f172a" font-family="Arial, sans-serif" font-size="13" font-weight="bold">' + tgl + '</text>' +
    '<text x="70" y="250" fill="#64748b" font-family="Arial, sans-serif" font-size="13">Penyetor / Kontak :</text>' +
    '<text x="210" y="250" fill="#0f172a" font-family="Arial, sans-serif" font-size="13">' + penyetor + ' (' + kontak + ')</text>' +
    '<line x1="70" y1="275" x2="730" y2="275" stroke="#e2e8f0" stroke-width="1"/>' +
    '<text x="70" y="310" fill="#64748b" font-family="Arial, sans-serif" font-size="13">Nomor STS :</text>' +
    '<text x="210" y="310" fill="#1e3a8a" font-family="monospace" font-size="13" font-weight="bold">' + sts + '</text>' +
    '<text x="420" y="310" fill="#64748b" font-family="Arial, sans-serif" font-size="13">Nomor STBP :</text>' +
    '<text x="530" y="310" fill="#0f172a" font-family="monospace" font-size="13">' + stbp + '</text>' +
    '<text x="70" y="345" fill="#64748b" font-family="Arial, sans-serif" font-size="13">ID Transaksi Bank :</text>' +
    '<text x="210" y="345" fill="#0f172a" font-family="monospace" font-size="13">' + tranx + '</text>' +
    '<rect x="70" y="375" width="660" height="60" fill="#ecfdf5" rx="8" stroke="#a7f3d0" stroke-width="1"/>' +
    '<text x="100" y="412" fill="#065f46" font-family="Arial, sans-serif" font-size="13" font-weight="bold">JUMLAH RETRIBUSI TERBAYAR :</text>' +
    '<text x="690" y="413" fill="#047857" font-family="monospace" font-size="18" font-weight="bold" text-anchor="end">Rp 750.000</text>' +
    '<rect x="570" y="125" width="160" height="36" fill="#dcfce7" rx="18" stroke="#86efac" stroke-width="1"/>' +
    '<text x="650" y="148" fill="#15803d" font-family="Arial, sans-serif" font-size="12" font-weight="bold" text-anchor="middle">LUNAS / VERIFIED</text>' +
    '<text x="400" y="490" fill="#94a3b8" font-family="Arial, sans-serif" font-size="11" text-anchor="middle">Dokumen digital ini merupakan bukti rekapitulasi sah sistem retribusi SPPG DLH Kab. Pangkep</text>' +
    '</svg>';
}

