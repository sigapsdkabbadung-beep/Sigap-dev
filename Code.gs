/**
 * ============================================================
 * SIGAP SD — Backend API (Google Apps Script)
 * Database: Google Spreadsheet
 * ============================================================
 */

// GANTI DENGAN SPREADSHEET ID ANDA (dari URL spreadsheet)
const SPREADSHEET_ID = 'GANTI_DENGAN_SPREADSHEET_ID_ANDA';

/**
 * Handle GET Request — Membaca data
 */
function doGet(e) {
  const table = e.parameter.table;
  const id = e.parameter.id;

  try {
    if (!table) return jsonResponse({ status: 'error', message: 'Parameter table wajib' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(table);

    if (!sheet) return jsonResponse({ status: 'error', message: `Tabel "${table}" tidak ditemukan` });

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return jsonResponse({ status: 'success', data: [] });

    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1);

    let result = rows.map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    }).filter(r => Object.values(r).some(v => v !== '' && v !== null && v !== undefined));

    if (id) {
      result = result.find(r => String(r.ID) === String(id)) || null;
    }

    return jsonResponse({ status: 'success', data: result });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * Handle POST Request — Tambah, Edit, Hapus, Update Setting
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const payload = JSON.parse(e.postData.contents);
    const { action, table, data, id } = payload;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(table);
    if (!sheet) throw new Error(`Tabel "${table}" tidak ditemukan`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    let message = '';

    // === UPDATE SETTING (khusus) ===
    if (action === 'updateSetting') {
      const key = data.Key;
      const value = data.Value;
      const rows = sheet.getDataRange().getValues();
      const rowIndex = rows.findIndex(r => String(r[0]) === String(key));

      if (rowIndex === -1) {
        sheet.appendRow([key, value]);
      } else {
        sheet.getRange(rowIndex + 2, 2).setValue(value);
      }
      message = 'Setting berhasil diperbarui';

    // === ADD ===
    } else if (action === 'add') {
      const newId = new Date().getTime();
      data.ID = newId;
      data['Created At'] = new Date().toISOString();
      data['Updated At'] = new Date().toISOString();
      if (data.Aktif === undefined) data.Aktif = 'Ya';

      const row = headers.map(h => data[h] !== undefined ? data[h] : '');
      sheet.appendRow(row);
      message = 'Data berhasil ditambahkan';

    // === UPDATE ===
    } else if (action === 'update') {
      const rows = sheet.getDataRange().getValues();
      const rowIndex = rows.findIndex(r => String(r[0]) === String(id));
      if (rowIndex === -1) throw new Error('Data tidak ditemukan');

      data['Updated At'] = new Date().toISOString();
      const oldRow = rows[rowIndex];
      const newRow = headers.map((h, i) => {
        if (h === 'ID' || h === 'Created At') return oldRow[i];
        if (data[h] !== undefined) return data[h];
        return oldRow[i];
      });

      sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([newRow]);
      message = 'Data berhasil diupdate';

    // === DELETE ===
    } else if (action === 'delete') {
      const rows = sheet.getDataRange().getValues();
      const rowIndex = rows.findIndex(r => String(r[0]) === String(id));
      if (rowIndex === -1) throw new Error('Data tidak ditemukan');

      sheet.deleteRow(rowIndex + 2);
      message = 'Data berhasil dihapus';

    } else {
      throw new Error('Action tidak dikenali');
    }

    return jsonResponse({ status: 'success', message });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper: Response JSON
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper: Setup awal — buat sheet jika belum ada (opsional, jalankan sekali)
 */
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const schemas = {
    'Sekolah': ['ID', 'NPSN', 'NamaSekolah', 'Status', 'Jenjang', 'Kecamatan', 'Desa', 'Alamat', 'KepalaSekolah', 'Telepon', 'Email', 'JumlahRombel', 'JumlahSiswa', 'KebutuhanGuru', 'Aktif', 'Created At', 'Updated At'],
    'Guru': ['ID', 'NIP', 'NamaGuru', 'JK', 'StatusPegawai', 'Pendidikan', 'Mapel', 'SekolahID', 'NoHP', 'Email', 'Aktif', 'Created At', 'Updated At'],
    'Kecamatan': ['ID', 'NamaKecamatan'],
    'TahunAjaran': ['ID', 'Tahun', 'Semester', 'Aktif'],
    'User': ['ID', 'Username', 'Password', 'Nama', 'Role', 'Aktif'],
    'Log': ['Tanggal', 'User', 'Aktivitas', 'IP'],
    'Setting': ['Key', 'Value'],
    'Rekap': ['Kecamatan', 'JumlahSekolah', 'JumlahGuru', 'PNS', 'PPPK', 'Honor', 'Kebutuhan'],
    'Backup': ['Tanggal', 'Keterangan']
  };

  for (const [name, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0e7ff');
      sheet.setFrozenRows(1);
    }
  }

  // Seed data kecamatan
  const kecSheet = ss.getSheetByName('Kecamatan');
  if (kecSheet.getLastRow() < 2) {
    const kecs = ['Abiansemal', 'Kuta', 'Kuta Selatan', 'Kuta Utara', 'Mengwi', 'Petang'];
    kecs.forEach((k, i) => kecSheet.appendRow([i + 1, k]));
  }

  // Seed setting tahun ajaran
  const setSheet = ss.getSheetByName('Setting');
  if (setSheet.getLastRow() < 2) {
    setSheet.appendRow(['TahunAjaran', '2024/2025 · Genap']);
  }

  SpreadsheetApp.getUi().alert('✅ Database berhasil di-setup!');
}