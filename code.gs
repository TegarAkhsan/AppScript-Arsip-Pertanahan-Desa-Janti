/**
 * E-Arsip Desa Janti - Backend Logic
 * Built for Google Apps Script
 */

const CONFIG = {
  FOLDER_ID: '1KulfsciZbhNAnU_sE7UUblVmaGL-xhAW',
  SHEET_NAME: 'DataArsip',
  USERS_SHEET: 'Users',
  JENIS_DOKUMEN_SHEET: 'JenisDokumen',
  APP_TITLE: 'E-Arsip Desa Janti',
  // MASUKKAN ID SPREADSHEET ANDA DI SINI
  SPREADSHEET_ID: '1PhIAp30aTcOVPak4E1Nqs8vSSxcBIgE_Vo8hC5zui48' 
};

/**
 * Initialize Spreadsheet and Folder
 */
function getSS() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setup() {
  const ss = getSS();
  if (!ss) throw new Error("Spreadsheet tidak ditemukan! Pastikan SPREADSHEET_ID di CONFIG sudah benar.");
  
  // Setup DataArsip
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const requiredHeaders = [
    'Timestamp', 'Nomor Arsip', 'Nama Pemilik', 'Jenis Dokumen', 
    'No. Bidang/Persil', 'Luas', 'Tahun', 'Status', 'File ID', 'File URL',
    'Alamat', 'Kecamatan', 'Kabupaten', 'Keterangan'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders])
         .setFontWeight('bold')
         .setBackground('#10b981')
         .setFontColor('white');
    sheet.setFrozenRows(1);
  } else {
    // Check for missing headers and add them
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    requiredHeaders.forEach(h => {
      if (existingHeaders.indexOf(h) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h)
             .setFontWeight('bold')
             .setBackground('#10b981')
             .setFontColor('white');
      }
    });
  }

  // Setup Users
  let userSheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!userSheet) {
    userSheet = ss.insertSheet(CONFIG.USERS_SHEET);
    const userHeaders = ['Nama', 'Username', 'Role', 'Status'];
    userSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders])
             .setFontWeight('bold')
             .setBackground('#10b981')
             .setFontColor('white');
    
    // Add default users
    userSheet.appendRow(['Admin Desa', 'admin', 'Administrator', 'Aktif']);
    userSheet.appendRow(['Perangkat Desa 1', 'perangkat1', 'Perangkat', 'Aktif']);
    userSheet.appendRow(['Operator Arsip', 'operator', 'Operator', 'Aktif']);
  }

  // Setup Jenis Dokumen
  let jenisSheet = ss.getSheetByName(CONFIG.JENIS_DOKUMEN_SHEET);
  if (!jenisSheet) {
    jenisSheet = ss.insertSheet(CONFIG.JENIS_DOKUMEN_SHEET);
    jenisSheet.getRange(1, 1, 1, 2).setValues([['Kategori', 'Jenis Dokumen']])
              .setFontWeight('bold')
              .setBackground('#10b981')
              .setFontColor('white');
    
    const docTypes = [
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Hak Milik (SHM)'],
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Hak Guna Bangunan (HGB)'],
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Hak Guna Usaha (HGU)'],
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Hak Pakai (HP)'],
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Hak Pengelolaan (HPL)'],
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Hak Satuan Rumah Susun (SHSRS / Strata Title)'],
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Tanah Wakaf'],
      ['1. Dokumen Hak Atas Tanah', 'Sertifikat Hak Milik atas Satuan Rumah Susun (HMSRS)'],
      ['1. Dokumen Hak Atas Tanah', 'Girik'],
      ['1. Dokumen Hak Atas Tanah', 'Petok D'],
      ['1. Dokumen Hak Atas Tanah', 'Letter C'],
      ['1. Dokumen Hak Atas Tanah', 'Kohir'],
      ['1. Dokumen Hak Atas Tanah', 'Verponding Indonesia'],
      ['1. Dokumen Hak Atas Tanah', 'Verponding Barat'],
      ['1. Dokumen Hak Atas Tanah', 'Pipil'],
      ['1. Dokumen Hak Atas Tanah', 'Kekitir'],
      ['1. Dokumen Hak Atas Tanah', 'Rincik'],
      ['1. Dokumen Hak Atas Tanah', 'Tanah Yasan'],
      ['1. Dokumen Hak Atas Tanah', 'Tanah Gogol'],
      ['1. Dokumen Hak Atas Tanah', 'Eigendom Verponding'],
      ['1. Dokumen Hak Atas Tanah', 'Erfpacht'],
      ['1. Dokumen Hak Atas Tanah', 'Opstal'],
      ['1. Dokumen Hak Atas Tanah', 'Agrarisch Eigendom'],
      ['1. Dokumen Hak Atas Tanah', 'SKT (Surat Keterangan Tanah)'],
      ['1. Dokumen Hak Atas Tanah', 'SKGR (Surat Keterangan Ganti Rugi)'],
      ['1. Dokumen Hak Atas Tanah', 'Sporadik'],
      ['1. Dokumen Hak Atas Tanah', 'Surat Segel Tanah'],
      ['1. Dokumen Hak Atas Tanah', 'Surat Riwayat Tanah'],
      ['2. Dokumen Peralihan Hak', 'Akta Jual Beli (AJB)'],
      ['2. Dokumen Peralihan Hak', 'Perjanjian Pengikatan Jual Beli (PPJB)'],
      ['2. Dokumen Peralihan Hak', 'Akta Hibah'],
      ['2. Dokumen Peralihan Hak', 'Akta Tukar Menukar'],
      ['2. Dokumen Peralihan Hak', 'Akta Pembagian Hak Bersama (APHB)'],
      ['2. Dokumen Peralihan Hak', 'Akta Waris'],
      ['2. Dokumen Peralihan Hak', 'Surat Keterangan Waris (SKW)'],
      ['2. Dokumen Peralihan Hak', 'Akta Pembagian Waris'],
      ['2. Dokumen Peralihan Hak', 'Akta Pelepasan Hak'],
      ['2. Dokumen Peralihan Hak', 'Akta Inbreng'],
      ['2. Dokumen Peralihan Hak', 'Akta Pemasukan ke Dalam Perusahaan'],
      ['2. Dokumen Peralihan Hak', 'Akta Penggabungan/Pemisahan Bidang Tanah'],
      ['2. Dokumen Peralihan Hak', 'Akta Kuasa Menjual'],
      ['2. Dokumen Peralihan Hak', 'Surat Pernyataan Pelepasan Hak'],
      ['2. Dokumen Peralihan Hak', 'Surat Pernyataan Penguasaan Fisik Bidang Tanah'],
      ['2. Dokumen Peralihan Hak', 'Akta Pembatalan'],
      ['2. Dokumen Peralihan Hak', 'Akta Perdamaian Sengketa Tanah'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Buku Tanah'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Surat Ukur'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Gambar Situasi'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Peta Bidang Tanah'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Peta Pendaftaran'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'NIB (Nomor Identifikasi Bidang)'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Risalah Penelitian Data Yuridis'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Risalah Panitia A'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Dokumen PTSL'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Surat Pernyataan Batas'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Berita Acara Pengukuran'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Berita Acara Penetapan Batas'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Peta Zona Nilai Tanah'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Dokumen Konsolidasi Tanah'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Warkah Pertanahan'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Daftar Isian Pertanahan'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Sertifikat Pengganti'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Sertifikat Pemecahan'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Sertifikat Penggabungan'],
      ['3. Dokumen Pendaftaran dan Pengukuran Tanah', 'Sertifikat Pemisahan'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'KKPR (Kesesuaian Kegiatan Pemanfaatan Ruang)'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Izin Lokasi'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Izin Perubahan Penggunaan Tanah (IPPT)'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Persetujuan Bangunan Gedung (PBG)'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'SLF (Sertifikat Laik Fungsi)'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Site Plan'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Andalalin'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Dokumen AMDAL'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'UKL-UPL'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'RDTR (Rencana Detail Tata Ruang)'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'RTRW (Rencana Tata Ruang Wilayah)'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Persetujuan Substansi Tata Ruang'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Fatwa Tata Guna Tanah'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Rekomendasi Tata Ruang'],
      ['4. Dokumen Perizinan dan Tata Ruang', 'Izin Reklamasi'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'SPPT PBB'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'STTS PBB'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'BPHTB'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'SSPD BPHTB'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'NJOP'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'Surat Ketetapan Pajak'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'Bukti Pembayaran Pajak Daerah'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'Validasi Pajak'],
      ['5. Dokumen Pajak dan Retribusi Pertanahan', 'Surat Keterangan Nilai Tanah'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'APHT (Akta Pemberian Hak Tanggungan)'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Sertifikat Hak Tanggungan (SHT)'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Roya'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'SKMHT (Surat Kuasa Membebankan Hak Tanggungan)'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Fidusia terkait bangunan tertentu'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Perjanjian Kredit'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Cover Note Notaris'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Surat Persetujuan Kredit'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Cessie'],
      ['6. Dokumen Pembiayaan dan Jaminan Tanah', 'Subrogasi'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Gugatan Sengketa Tanah'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Putusan Pengadilan'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Penetapan Pengadilan'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Berita Acara Mediasi'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Surat Keberatan'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Memori Banding'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Memori Kasasi'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'PK (Peninjauan Kembali)'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Eksekusi Pengadilan'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Putusan PTUN Pertanahan'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Berita Acara Pemeriksaan Lapangan'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Surat Blokir Sertifikat'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Surat Sita Jaminan'],
      ['7. Dokumen Sengketa dan Penyelesaian Tanah', 'Surat Perdamaian'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Penetapan Lokasi (Penlok)'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Dokumen Pengadaan Tanah'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Daftar Nominatif'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Inventarisasi dan Identifikasi Bidang Tanah'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Berita Acara Ganti Kerugian'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Surat Pelepasan Hak'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Konsinyasi'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Dokumen Pengukuran Pengadaan Tanah'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Penilaian Appraisal Tanah'],
      ['8. Dokumen Pengadaan Tanah dan Pemerintah', 'Surat Keputusan Pengadaan Tanah'],
      ['9. Dokumen Pendukung Pertanahan', 'KTP Pemilik'],
      ['9. Dokumen Pendukung Pertanahan', 'KK'],
      ['9. Dokumen Pendukung Pertanahan', 'NPWP'],
      ['9. Dokumen Pendukung Pertanahan', 'Surat Nikah'],
      ['9. Dokumen Pendukung Pertanahan', 'Akta Kelahiran'],
      ['9. Dokumen Pendukung Pertanahan', 'Akta Kematian'],
      ['9. Dokumen Pendukung Pertanahan', 'Surat Kuasa'],
      ['9. Dokumen Pendukung Pertanahan', 'Surat Pernyataan'],
      ['9. Dokumen Pendukung Pertanahan', 'Surat Domisili'],
      ['9. Dokumen Pendukung Pertanahan', 'IMB Lama'],
      ['9. Dokumen Pendukung Pertanahan', 'Rekening Listrik/Air'],
      ['9. Dokumen Pendukung Pertanahan', 'Bukti Pembayaran Tanah'],
      ['9. Dokumen Pendukung Pertanahan', 'Kwitansi Jual Beli'],
      ['9. Dokumen Pendukung Pertanahan', 'Foto Lokasi Tanah'],
      ['9. Dokumen Pendukung Pertanahan', 'Patok Batas Tanah'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'Sertifikat Elektronik'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'HT-el (Hak Tanggungan Elektronik)'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'Dokumen Pendaftaran Elektronik'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'Peta Digital Pertanahan'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'Arsip Elektronik BPN'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'Validasi Elektronik Sertifikat'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'Informasi Zona Nilai Tanah Digital'],
      ['10. Dokumen Digital dan Sistem Elektronik Pertanahan', 'Dokumen Integrasi OSS-RBA']
    ];
    
    jenisSheet.getRange(2, 1, docTypes.length, 2).setValues(docTypes);
  }
}

function getJenisDokumen() {
  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.JENIS_DOKUMEN_SHEET);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  
  const categories = {};
  rows.forEach(r => {
    const cat = r[0];
    const item = r[1];
    if (!categories[cat]) categories[cat] = [];
    categories[cat].push(item);
  });
  
  return categories;
}

function doGet(e) {
  try {
    setup(); 
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle(CONFIG.APP_TITLE)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <div style="font-family: sans-serif; padding: 40px; text-align: center;">
        <h2 style="color: #dc2626;">Terjadi Kesalahan Sistem</h2>
        <p>${error.toString()}</p>
        <p style="color: #64748b;">Pastikan SPREADSHEET_ID sudah diisi di file code.gs</p>
      </div>
    `);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * Handle File Upload and Data Saving
 */
function uploadFile(data) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const ss = getSS();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    let fileId = '';
    let fileUrl = '';

    // Process File if exists
    if (data.file && data.file.includes('base64')) {
      const contentType = data.file.substring(5, data.file.indexOf(';'));
      const bytes = Utilities.base64Decode(data.file.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, data.fileName);
      const file = folder.createFile(blob);
      fileId = file.getId();
      fileUrl = file.getUrl();
    }
    
    // Save to Sheet
    const rowData = [
      new Date(),
      data.nomorArsip,
      data.namaPemilik,
      data.jenisDokumen,
      data.noBidang,
      data.luas,
      data.tahun,
      'Aktif',
      fileId,
      fileUrl,
      data.alamat || '-',
      data.kecamatan || '-',
      data.kabupaten || '-',
      data.keterangan || '-'
    ];
    
    sheet.appendRow(rowData);
    
    return { success: true, message: 'Data berhasil disimpan!', nomorArsip: data.nomorArsip };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateArchive(data) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === data.oldNomorArsip) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Data tidak ditemukan");
    
    // Update specific columns (excluding Timestamp and File Info if not changed)
    sheet.getRange(rowIndex, 2, 1, 6).setValues([[
      data.nomorArsip,
      data.namaPemilik,
      data.jenisDokumen,
      data.noBidang,
      data.luas,
      data.tahun
    ]]);
    
    sheet.getRange(rowIndex, 11, 1, 4).setValues([[
      data.alamat,
      data.kecamatan,
      data.kabupaten,
      data.keterangan
    ]]);

    return { success: true, message: 'Data berhasil diperbarui!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteArchive(nomorArsip) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === nomorArsip) {
        // Option 1: Mark as Deleted
        // sheet.getRange(i + 1, 8).setValue('Terhapus');
        
        // Option 2: Delete row
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data berhasil dihapus!' };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getLastNumber(sheet, year) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  
  for (let i = 1; i < data.length; i++) {
    const nr = data[i][1]; // Nomor Arsip column
    if (nr && nr.includes(`ARS-${year}-`)) {
      const num = parseInt(nr.split('-')[2]);
      if (num > max) max = num;
    }
  }
  return max;
}

/**
 * Fetch Dashboard Statistics
 */
function getDashboardData() {
  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  
  const now = new Date();
  const thisMonth = now.getMonth();
  const thisYear = now.getFullYear();
  
  const currentMonthData = rows.filter(r => {
    const d = new Date(r[0]);
    return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
  });
  
  const lastMonthData = rows.filter(r => {
    const d = new Date(r[0]);
    const lm = thisMonth === 0 ? 11 : thisMonth - 1;
    const ly = thisMonth === 0 ? thisYear - 1 : thisYear;
    return d.getMonth() === lm && d.getFullYear() === ly;
  });

  const getGrowth = (curr, last) => {
    if (last === 0) return curr > 0 ? `+${curr}` : '0%';
    const pct = Math.round(((curr - last) / last) * 100);
    return pct >= 0 ? `+${pct}%` : `${pct}%`;
  };

  const stats = {
    total: rows.length,
    totalGrowth: getGrowth(currentMonthData.length, lastMonthData.length),
    sertifikat: rows.filter(r => r[3].toString().toLowerCase().includes('sertifikat')).length,
    sertifikatGrowth: getGrowth(
      currentMonthData.filter(r => r[3].toString().toLowerCase().includes('sertifikat')).length,
      lastMonthData.filter(r => r[3].toString().toLowerCase().includes('sertifikat')).length
    ),
    suratKeterangan: rows.filter(r => r[3].toString().toLowerCase().includes('surat')).length,
    skGrowth: getGrowth(
      currentMonthData.filter(r => r[3].toString().toLowerCase().includes('surat')).length,
      lastMonthData.filter(r => r[3].toString().toLowerCase().includes('surat')).length
    ),
    tahunIni: rows.filter(r => r[6].toString() === thisYear.toString()).length,
    recent: rows.slice(-5).reverse().map(r => ({
      no: r[1],
      nama: r[2],
      jenis: r[3],
      tanggal: Utilities.formatDate(new Date(r[0]), "GMT+7", "dd MMM yyyy"),
      status: r[7]
    }))
  };
  
  return stats;
}

/**
 * Fetch All Archives
 */
function getArchives() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => {
      if (h === 'Timestamp') {
        obj[h] = Utilities.formatDate(new Date(r[i]), "GMT+7", "dd/MM/yyyy HH:mm");
      } else {
        obj[h] = r[i];
      }
    });
    return obj;
  });
}

function deleteArchive(id) {
  // Logic to delete file from drive and row from sheet if needed
}

/**
 * Get Report Data
 */
function getReportData() {
  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const rows = sheet.getDataRange().getValues().slice(1);
  const year = 2026; // Current display year
  
  const data2026 = rows.filter(r => r[6].toString() === year.toString());
  
  const monthlyTrend = Array(12).fill(0).map((_, i) => {
    const count = data2026.filter(r => new Date(r[0]).getMonth() === i).length;
    return { month: ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'][i], count: count };
  });

  const jenisSummary = {};
  data2026.forEach(r => {
    const jenis = r[3];
    jenisSummary[jenis] = (jenisSummary[jenis] || 0) + 1;
  });

  const total = data2026.length;
  const summaryList = Object.keys(jenisSummary).map(k => ({
    jenis: k,
    jumlah: jenisSummary[k],
    persentase: Math.round((jenisSummary[k] / total) * 100)
  }));

  return {
    total: total,
    avg: Math.round(total / (new Date().getMonth() + 1)),
    thisMonth: rows.filter(r => {
      const d = new Date(r[0]);
      return d.getMonth() === new Date().getMonth() && d.getFullYear() === new Date().getFullYear();
    }).length,
    summary: summaryList,
    trend: monthlyTrend.slice(0, new Date().getMonth() + 1)
  };
}

/**
 * Get User List
 */
function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

/**
 * Create 10 Dummy Data Entries
 */
function createDummyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  const dummyNames = ['Budi Santoso', 'Siti Aminah', 'Ahmad Wijaya', 'Dewi Lestari', 'Rudi Hartono', 'Rina Susanti', 'Hendra Gunawan', 'Maya Sari', 'Eko Prasetyo', 'Lusi Fitriani'];
  const dummyJenis = ['Sertifikat Tanah', 'Surat Keterangan Tanah', 'Sertifikat Tanah', 'Girik', 'Sertifikat Tanah', 'Surat Keterangan Tanah', 'Sertifikat Tanah', 'Girik', 'Sertifikat Tanah', 'Sertifikat Tanah'];
  
  const now = new Date();
  
  for (let i = 0; i < 10; i++) {
    const date = new Date(now);
    date.setDate(now.getDate() - i);
    
    const year = 2026;
    const lastNum = getLastNumber(sheet, year);
    const nextNum = (lastNum + 1).toString().padStart(4, '0');
    const nomorArsip = `ARS-${year}-${nextNum}`;
    
    const row = [
      date,
      nomorArsip,
      dummyNames[i],
      dummyJenis[i],
      `01${i}/45`,
      150 + (i * 20),
      year,
      'Aktif',
      'dummy-id-' + i,
      'https://drive.google.com/file/d/1KulfsciZbhNAnU_sE7UUblVmaGL-xhAW',
      'Jl. Raya Janti No. 12' + i + ', Desa Janti',
      'Waru',
      'Sidoarjo',
      'Sertifikat Hak Milik atas nama ' + dummyNames[i] + '. Dokumen asli tersimpan di lemari arsip nomor ' + (i+1) + '.'
    ];
    sheet.appendRow(row);
  }
  return "10 Dummy data created!";
}
