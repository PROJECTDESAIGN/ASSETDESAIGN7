// ============================================================
// Code.gs — ASSET PRO v2 · Google Apps Script Backend
// Koneksi ke Google Sheets untuk manajemen aset tetap
// dan stock opname. Deploy sebagai Web App (Execute as Me,
// Access: Anyone with the link).
// ============================================================

// ──────────────────────────────────────────────
// KONFIGURASI — ganti dengan ID Spreadsheet Anda
// ──────────────────────────────────────────────
const SPREADSHEET_ID = '1jAaI1W1sK4aUHBzNdO4deTEP5DSexHGDd1JQpZJ-_No';

// Nama file HTML di project Google Apps Script (tanpa ekstensi .html).
// Umumnya file yang dibuat dari template ini bernama "index".
const HTML_FILE_NAME = 'index';
const HTML_FILE_FALLBACK_NAME = 'index_1784811138558';

// Nama sheet (tab) di Spreadsheet
const SHEET_ASSETS     = 'Daftar Aset';
const SHEET_ACTIVE_SO  = 'Draft SO';
const SHEET_FINAL_SO   = 'Hasil SO Final';
const SHEET_PERGERAKAN = 'Pergerakan SO';

// ──────────────────────────────────────────────
// ENTRY POINT — Web App
// ──────────────────────────────────────────────
function doGet(e) {
  // Ambil parameter 'action' dari URL
  const action = e && e.parameter ? (e.parameter.action || '') : '';

  try {
    switch (action) {
      case 'getAssets':           return jsonResponse(getAssets());
      case 'getActiveSO':         return jsonResponse(getActiveSO());
      case 'getFinalSO':          return jsonResponse(getFinalSO());
      case 'getPergerakanSO':     return jsonResponse(getPergerakanSO());
      
      // Tanpa action, tampilkan aplikasi HTML. Endpoint API tetap memakai
      // parameter ?action=... atau method POST.
      default:
        return createAppHtml();
    }
  } catch (err) {
    // Jika HTML gagal dibuka, tetap kembalikan informasi error yang mudah
    // dibaca saat URL Web App dibuka langsung.
    return jsonResponse({ status: 'error', error: err.message });
  }
}

function createAppHtml() {
  let html;
  try {
    html = HtmlService.createHtmlOutputFromFile(HTML_FILE_NAME);
  } catch (firstError) {
    try {
      html = HtmlService.createHtmlOutputFromFile(HTML_FILE_FALLBACK_NAME);
    } catch (secondError) {
      throw new Error(
        'File HTML tidak ditemukan. Buat file index.html di project Apps Script, ' +
        'lalu tempel isi file HTML yang disediakan.'
      );
    }
  }
  return html
    .setTitle('Asset Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let payload = {};
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (parseError) {
    return jsonResponse({ success: false, error: 'Body request bukan JSON yang valid.' });
  }
  const action = payload.action || '';
  
  try {
    switch (action) {
      // Snapshot dipakai oleh frontend agar aset, draft SO, arsip SO, dan
      // URL spreadsheet selalu dimuat dalam satu request.
      case 'getAssets':
      case 'getAllAssets':
        return jsonResponse({ status: 'success', data: getSpreadsheetData() });

      // 2. Aset Master
      case 'addOrUpdateAsset':
      case 'saveAsset':
      case 'addAsset':
        return jsonResponse(addOrUpdateAssetGas(payload.data || payload.assetData));

      case 'updateAsset':
        return jsonResponse(updateAsset(payload.data || payload.assetData));

      case 'deleteAsset':
        return jsonResponse(deleteAsset(payload.id));

      // 3. Stock Opname (SO)
      case 'addSOItem':
        // Kompatibilitas dengan versi frontend lama yang mengirim seluruh
        // draft sebagai array.
        if (Array.isArray(payload.data)) {
          return jsonResponse(saveStockOpnameDraftGas(payload.data));
        }
        return jsonResponse(addSOItem(payload.data));

      case 'saveStockOpnameDraft':
        return jsonResponse(saveStockOpnameDraftGas(payload.data || []));

      case 'updateSOItem':
        return jsonResponse(updateSOItem(payload.data));

      case 'deleteSOItem':
        return jsonResponse(deleteSOItem(payload.id));

      case 'releaseSOReport':
        return jsonResponse(releaseSOReport(payload.data));

      case 'updateCompletedSO':
        return jsonResponse(updateCompletedSOGas(
          payload.id || (payload.data && payload.data.id),
          payload.periode || (payload.data && payload.data.periode),
          payload.lokasi || (payload.data && payload.data.lokasi)
        ));

      case 'deleteCompletedSO':
        return jsonResponse(deleteCompletedSOGas(
          payload.id || (payload.data && payload.data.id)
        ));

      case 'savePergerakanSO':
        return jsonResponse(savePergerakanSO(payload.data));

      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ──────────────────────────────────────────────
// HELPER — JSON response with CORS headers
// ──────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'MASUKKAN_ID_SPREADSHEET_ANDA_DI_SINI') {
    throw new Error('SPREADSHEET_ID belum diisi. Masukkan ID Google Spreadsheet pada Code.gs.');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ──────────────────────────────────────────────
// DAFTAR ASET TETAP
// ──────────────────────────────────────────────
// Header: ID | KodeAset | NamaAset | Barcode | Kategori | Lokasi | TanggalBeli | MasaManfaat | Status | Keterangan | FotoUrl

function getAssets() {
  const sheet = getSheet(SHEET_ASSETS);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({
    id:          String(r[0]),
    kodeAset:    String(r[1]),
    namaAset:    String(r[2]),
    barcode:     String(r[3]).replace(/\.0$/, ''),
    kategori:    String(r[4]),
    lokasi:      String(r[5]),
    tanggalBeli: String(r[6]),
    masaManfaat: String(r[7]),
    status:      String(r[8]),
    keterangan:  String(r[9]),
    fotoUrl:     convertDriveUrl(String(r[10] || '')),
  }));
}

// Nama lama ini dipertahankan agar deployment sebelumnya tetap kompatibel.
function getAssetsFromSheet() {
  return getAssets();
}

// ============================================================
// FUNGSI BRIDGE UNTUK iNDEX.HTML
// ============================================================
// Halaman menggunakan google.script.run secara langsung. Nama fungsi bridge
// berikut sengaja disamakan dengan yang dipanggil oleh frontend.

function getSpreadsheetData() {
  const ss = getSpreadsheet();
  return {
    success: true,
    assets: getAssets(),
    activeSO: getActiveSO(),
    completedSO: getFinalSO(),
    sheetUrl: ss.getUrl(),
  };
}

function getSpreadsheet() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'MASUKKAN_ID_SPREADSHEET_ANDA_DI_SINI') {
    throw new Error('SPREADSHEET_ID belum diisi. Masukkan ID Google Spreadsheet pada Code.gs.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function addOrUpdateAssetGas(data) {
  if (!data || !data.namaAset || !data.kodeAset) {
    return { error: 'Data aset tidak lengkap.' };
  }
  return data.id && data.id.indexOf('ast-') === 0
    ? updateAsset(data)
    : saveAsset(data);
}

function deleteAssetGas(id) {
  return deleteAsset(id);
}

// Menyimpan seluruh draft sebagai satu snapshot agar urutan operasi dari
// browser tidak menyebabkan data lama tertinggal di sheet.
function saveStockOpnameDraftGas(items) {
  const sheet = getSheet(SHEET_ACTIVE_SO);
  ensureHeader(sheet, ['ID','KodeBarang','NamaBarang','Lokasi','Qty','WaktuInput']);

  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  const rows = (items || []).map(item => [
    item.id || ('so-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7)),
    cleanBarcode(item.kodeBarang),
    item.namaBarang || '',
    item.lokasi || '',
    Number(item.qty) || 0,
    item.waktuInput || formatNow('yyyy-MM-dd HH:mm:ss'),
  ]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  return { success: true, jumlahItem: rows.length };
}

function saveStockOpnameGas(data) {
  if (!data || !Array.isArray(data.items) || data.items.length === 0) {
    return { error: 'Draft SO kosong.' };
  }

  const sheet = getSheet(SHEET_FINAL_SO);
  ensureHeader(sheet, ['ID','Periode','TanggalSelesai','Lokasi','KodeBarang','NamaBarang','Qty','WaktuInput']);
  const reportId = data.id || ('cso-' + Date.now());
  const tanggal = data.tanggalSelesai || formatNow('dd/MM/yyyy HH:mm');
  const rows = data.items.map(item => [
    reportId,
    data.periode || '',
    tanggal,
    data.lokasi || '',
    cleanBarcode(item.kodeBarang),
    item.namaBarang || '',
    Number(item.qty) || 0,
    item.waktuInput || '',
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);

  const draftSheet = getSheet(SHEET_ACTIVE_SO);
  ensureHeader(draftSheet, ['ID','KodeBarang','NamaBarang','Lokasi','Qty','WaktuInput']);
  if (draftSheet.getLastRow() > 1) {
    draftSheet.deleteRows(2, draftSheet.getLastRow() - 1);
  }
  return { success: true, reportId, jumlahItem: rows.length };
}

function updateCompletedSOGas(id, periode, lokasi) {
  const sheet = getSheet(SHEET_FINAL_SO);
  const range = sheet.getDataRange();
  const rows = range.getValues();
  let updated = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2).setValue(periode || '');
      sheet.getRange(i + 1, 4).setValue(lokasi || '');
      updated++;
    }
  }
  return updated ? { success: true } : { error: 'Arsip SO tidak ditemukan: ' + id };
}

function deleteCompletedSOGas(id) {
  const sheet = getSheet(SHEET_FINAL_SO);
  const rows = sheet.getDataRange().getValues();
  let deleted = 0;
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted ? { success: true, jumlahBaris: deleted } : { error: 'Arsip SO tidak ditemukan: ' + id };
}

function ensureHeader(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function formatNow(pattern) {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), pattern);
}

function saveAsset(data) {
  const sheet = getSheet(SHEET_ASSETS);
  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID','KodeAset','NamaAset','Barcode','Kategori','Lokasi','TanggalBeli','MasaManfaat','Status','Keterangan','FotoUrl']);
  }
  const id = 'ast-' + Date.now();
  sheet.appendRow([id, data.kodeAset, data.namaAset, data.barcode, data.kategori,
    data.lokasi, data.tanggalBeli, data.masaManfaat, data.status, data.keterangan, data.fotoUrl || '']);
  return { success: true, id };
}

function updateAsset(data) {
  const sheet = getSheet(SHEET_ASSETS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 1, 1, 11).setValues([[
        data.id, data.kodeAset, data.namaAset, data.barcode, data.kategori,
        data.lokasi, data.tanggalBeli, data.masaManfaat, data.status,
        data.keterangan, data.fotoUrl || '',
      ]]);
      return { success: true };
    }
  }
  return { error: 'Aset tidak ditemukan: ' + data.id };
}

function deleteAsset(id) {
  const sheet = getSheet(SHEET_ASSETS);
  const rows  = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Aset tidak ditemukan: ' + id };
}

// ──────────────────────────────────────────────
// DRAFT SO (STOCK OPNAME AKTIF)
// ──────────────────────────────────────────────
// Header: ID | KodeBarang | NamaBarang | Lokasi | Qty | WaktuInput

function getActiveSO() {
  const sheet = getSheet(SHEET_ACTIVE_SO);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => ({
    id:          String(r[0]),
    kodeBarang:  String(r[1]).replace(/\.0$/, ''),
    namaBarang:  String(r[2]),
    lokasi:      String(r[3]),
    qty:         Number(r[4]) || 0,
    waktuInput:  String(r[5]),
  }));
}

function addSOItem(data) {
  const sheet = getSheet(SHEET_ACTIVE_SO);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID','KodeBarang','NamaBarang','Lokasi','Qty','WaktuInput']);
  }
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  // Check if same barcode+lokasi already exists — if so, accumulate qty
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).replace(/\.0$/, '') === String(data.kodeBarang) &&
        String(rows[i][3]) === String(data.lokasi)) {
      const newQty = Number(rows[i][4]) + Number(data.qty);
      sheet.getRange(i + 1, 5).setValue(newQty);
      sheet.getRange(i + 1, 6).setValue(timestamp);
      return { success: true, id: String(rows[i][0]), accumulated: true };
    }
  }
  const id = 'so-' + Date.now();
  sheet.appendRow([id, data.kodeBarang, data.namaBarang, data.lokasi, data.qty, timestamp]);
  return { success: true, id };
}

function updateSOItem(data) {
  const sheet = getSheet(SHEET_ACTIVE_SO);
  const rows  = sheet.getDataRange().getValues();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 4).setValue(data.lokasi);
      sheet.getRange(i + 1, 5).setValue(data.qty);
      sheet.getRange(i + 1, 6).setValue(timestamp);
      return { success: true };
    }
  }
  return { error: 'Item SO tidak ditemukan: ' + data.id };
}

function deleteSOItem(id) {
  const sheet = getSheet(SHEET_ACTIVE_SO);
  const rows  = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Item SO tidak ditemukan: ' + id };
}

// ──────────────────────────────────────────────
// RILIS LAPORAN SO → Hasil SO Final
// ──────────────────────────────────────────────
// Menyalin semua item draft ke sheet Hasil SO Final, lalu bersihkan Draft SO.

function releaseSOReport(data) {
  // data: { periode, lokasi }
  const draftSheet = getSheet(SHEET_ACTIVE_SO);
  const finalSheet = getSheet(SHEET_FINAL_SO);

  const draftRows = draftSheet.getDataRange().getValues();
  const items = draftRows.slice(1); // exclude header
  if (items.length === 0) return { error: 'Draft SO kosong.' };

  // Ensure header in final sheet
  if (finalSheet.getLastRow() === 0) {
    finalSheet.appendRow(['ID','Periode','TanggalSelesai','Lokasi','KodeBarang','NamaBarang','Qty','WaktuInput']);
  }

  const tanggal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const reportId = 'cso-' + Date.now();

  items.forEach(row => {
    finalSheet.appendRow([
      reportId, data.periode, tanggal, data.lokasi,
      String(row[1]).replace(/\.0$/, ''), String(row[2]), Number(row[4]), String(row[5]),
    ]);
  });

  // Clear draft (keep header row)
  if (draftSheet.getLastRow() > 1) {
    draftSheet.deleteRows(2, draftSheet.getLastRow() - 1);
  }

  return { success: true, reportId, jumlahItem: items.length };
}

function getFinalSO() {
  const sheet = getSheet(SHEET_FINAL_SO);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  // Group by reportId
  const groups = {};
  rows.slice(1).forEach(r => {
    const rid = String(r[0]);
    if (!groups[rid]) {
      groups[rid] = {
        id: rid,
        periode: String(r[1]),
        tanggalSelesai: String(r[2]),
        lokasi: String(r[3]),
        items: [],
      };
    }
    groups[rid].items.push({
      kodeBarang: String(r[4]),
      namaBarang: String(r[5]),
      qty: Number(r[6]),
      waktuInput: String(r[7]),
    });
  });
  return Object.values(groups).map(g => ({
    ...g,
    jumlahItem: g.items.length,
    totalQty: g.items.reduce((s, i) => s + i.qty, 0),
    namaFileExcel: `SO_${g.id.replace('cso-','')}.xlsx`,
  }));
}

// ──────────────────────────────────────────────
// PERGERAKAN SO — Simpan hasil perbandingan final
// ──────────────────────────────────────────────
// Header: ID | Periode | TanggalFinalisasi | Lokasi | KodeBarang | NamaBarang | Satuan | QtyDraft | QtyFisik | Selisih | Status

function savePergerakanSO(data) {
  // data: { id, periode, tanggalFinalisasi, lokasi, rows: ComparisonRow[] }
  const sheet = getSheet(SHEET_PERGERAKAN);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID','Periode','TanggalFinalisasi','Lokasi','KodeBarang','NamaBarang','Satuan','QtyDraftSistem','QtyFisikImport','Selisih','Status']);
  }
  (data.rows || []).forEach(row => {
    sheet.appendRow([
      data.id, data.periode, data.tanggalFinalisasi, data.lokasi,
      row.kodeBarang, row.namaBarang, row.satuan,
      row.qtyDraft !== null ? row.qtyDraft : '',
      row.qtyFisik !== null ? row.qtyFisik : '',
      row.selisih, row.status,
    ]);
  });
  return { success: true, jumlahSKU: (data.rows || []).length };
}

function getPergerakanSO() {
  const sheet = getSheet(SHEET_PERGERAKAN);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const groups = {};
  rows.slice(1).forEach(r => {
    const id = String(r[0]);
    if (!groups[id]) {
      groups[id] = {
        id, periode: String(r[1]), tanggalFinalisasi: String(r[2]),
        lokasi: String(r[3]), rows: [],
      };
    }
    groups[id].rows.push({
      kodeBarang: String(r[4]),
      namaBarang: String(r[5]),
      satuan: String(r[6]),
      qtyDraft: r[7] !== '' ? Number(r[7]) : null,
      qtyFisik: r[8] !== '' ? Number(r[8]) : null,
      selisih: Number(r[9]) || 0,
      status: String(r[10]),
    });
  });
  return Object.values(groups).map(g => ({
    ...g,
    jumlahSKU: g.rows.length,
    totalSelisih: g.rows.reduce((s, r) => s + Math.abs(r.selisih || 0), 0),
  }));
}

// ──────────────────────────────────────────────
// UTILITY — Convert Google Drive share URL → thumbnail
// ──────────────────────────────────────────────
function convertDriveUrl(url) {
  if (!url) return '';
  const s = url.trim();
  if (s.includes('drive.google.com/thumbnail')) return s;
  const m1 = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) return `https://drive.google.com/thumbnail?id=${m1[1]}&sz=w800`;
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return `https://drive.google.com/thumbnail?id=${m2[1]}&sz=w800`;
  return s;
}

// ──────────────────────────────────────────────
// UTILITY — Bersihkan barcode dari trailing .0
// ──────────────────────────────────────────────
function cleanBarcode(raw) {
  return String(raw || '').trim().replace(/\.0+$/, '');
}
