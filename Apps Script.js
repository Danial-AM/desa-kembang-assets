/**
 * Konfigurasi Sheet Google Sheets
 * Mendaftarkan tab yang dapat diakses, nama sheet asli, dan mode parsingnya.
 */
const SHEET_CONFIG = {
  'Beranda': { sheetName: 'Beranda', isKeyValue: true },
  'FaktaCepat': { sheetName: 'FaktaCepat', isKeyValue: false },
  'Wisata': { sheetName: 'Wisata', isKeyValue: false },
  'Produk': { sheetName: 'Produk', isKeyValue: false },
  'DataDesa': { sheetName: 'DataDesa', isKeyValue: true },
  'VisiMisi': { sheetName: 'VisiMisi', isKeyValue: false },
  'PotensiDesa': { sheetName: 'PotensiDesa', isKeyValue: false },
  'PetaBencana': { sheetName: 'PetaBencana', isKeyValue: true },
  'PerangkatDesa': { sheetName: 'PerangkatDesa', isKeyValue: false },
  'KontakDarurat': { sheetName: 'KontakDarurat', isKeyValue: false },
  'Kontak': { sheetName: 'Kontak', isKeyValue: true },
  'Dusun': { sheetName: 'Dusun', isKeyValue: false },
  'Bangunan': { sheetName: 'Bangunan', isKeyValue: false },
  'Evakuasi': { sheetName: 'Evakuasi', isKeyValue: false },
  'Sejarah': { sheetName: 'Sejarah', isKeyValue: true }
};

const CACHE_KEY_ALL = 'SHEET_DATA_ALL';
const CACHE_TIME_SEC = 300; // 5 menit (300 detik)

/**
 * Menangani request GET
 * Mengambil data dari sheet (Satu sheet atau semua sheet sekaligus)
 */
function doGet(e) {
  try {
    const tab = (e.parameter.tab || '').toLowerCase();
    const tabParam = e.parameter.tab || '';
    
    // Mengembalikan semua data sheet jika parameter tab=all
    if (tab === 'all') {
      const cache = CacheService.getScriptCache();
      const cachedData = cache.get(CACHE_KEY_ALL);
      
      // Gunakan cache jika masih tersedia
      if (cachedData) {
        return jsonResponse(JSON.parse(cachedData));
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const allData = {};
      
      for (const key in SHEET_CONFIG) {
        const config = SHEET_CONFIG[key];
        const actualSheetName = findSheetName(config.sheetName) || config.sheetName;
        const sheet = ss.getSheetByName(actualSheetName);
        
        if (sheet) {
          allData[key] = parseSheet(sheet, config);
        } else {
          allData[key] = config.isKeyValue ? {} : [];
        }
      }
      
      // Simpan ke cache selama 5 menit
      cache.put(CACHE_KEY_ALL, JSON.stringify(allData), CACHE_TIME_SEC);
      return jsonResponse(allData);
      
    } else if (tabParam) {
      // Mengembalikan data satu sheet berdasarkan parameter tab
      // Cari konfigurasi (case-insensitive)
      const configKey = Object.keys(SHEET_CONFIG).find(k => k.toLowerCase() === tabParam.toLowerCase());
      const config = configKey ? SHEET_CONFIG[configKey] : null;
      if (!config) {
        throw new Error('Tab tidak ditemukan dalam konfigurasi.');
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const actualSheetName = findSheetName(config.sheetName) || config.sheetName;
      const sheet = ss.getSheetByName(actualSheetName);
      
      if (!sheet) {
        throw new Error(`Sheet ${actualSheetName} tidak ditemukan.`);
      }
      
      const data = parseSheet(sheet, config);
      return jsonResponse(data);
      
    } else {
      throw new Error('Parameter tab tidak diberikan.');
    }
    
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.toString() });
  }
}

/**
 * Menangani request POST (Terutama untuk Formulir Kontak)
 */
function doPost(e) {
  try {
    let requestData;
    
    // Membaca body request sebagai JSON
    if (e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else {
      throw new Error('Data POST tidak valid atau kosong.');
    }
    
    const { nama, email, subjek, pesan } = requestData;
    
    if (!nama || !pesan) {
      throw new Error('Field nama dan pesan wajib diisi.');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Kontak-Masuk');
    
    // Buat sheet Kontak-Masuk jika belum ada
    if (!sheet) {
      sheet = ss.insertSheet('Kontak-Masuk');
      sheet.appendRow(['Timestamp', 'Nama', 'Email', 'Subjek', 'Pesan']);
    }
    
    // Tambahkan data pesan ke sheet
    const timestamp = new Date();
    sheet.appendRow([timestamp, nama.trim(), (email || '').trim(), (subjek || '').trim(), pesan.trim()]);
    
    return jsonResponse({ status: 'success', message: 'Pesan berhasil dikirim.' });
    
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.toString() });
  }
}

/**
 * Parsing data dari sheet berdasarkan konfigurasinya
 */
function parseSheet(sheet, config) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getDisplayValues(); // Gunakan getDisplayValues untuk menangani format tanggal/angka
  
  if (values.length === 0) return config.isKeyValue ? {} : [];
  
  if (config.isKeyValue) {
    // Mode Key-Value (Kolom 1: Kunci, Kolom 2: Nilai)
    const result = {};
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const key = String(row[0]).trim();
      if (key) {
        const value = row[1] !== undefined ? String(row[1]).trim() : '';
        result[key] = value;
      }
    }
    return result;
  } else {
    // Mode Array Object (Baris 1: Header, Baris selanjutnya: Data)
    const headers = values[0].map(h => String(h).trim());
    const result = [];
    
    // Cari index kolom StatusPublikasi (case-insensitive) jika ada
    const statusIndex = headers.findIndex(h => h.toLowerCase() === 'statuspublikasi');
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Lewati baris yang kosong semua
      if (row.every(cell => String(cell).trim() === '')) continue;
      
      // Terapkan filter StatusPublikasi jika kolom tersebut ada
      if (statusIndex !== -1) {
        const status = String(row[statusIndex]).trim().toLowerCase();
        // Skip baris jika status bukan ya, aktif, atau 1
        if (status !== 'ya' && status !== 'aktif' && status !== '1' && status !== 'true' && status !== 'publikasikan') {
          continue;
        }
      }
      
      const rowObj = {};
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        if (header) {
          rowObj[header] = row[j] !== undefined ? String(row[j]).trim() : '';
        }
      }
      result.push(rowObj);
    }
    
    return result;
  }
}

/**
 * Membuat response dalam format JSON
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mencari nama sheet yang sesuai tanpa mempedulikan huruf besar/kecil (case-insensitive)
 */
function findSheetName(requestedName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const lowerRequested = requestedName.toLowerCase();
  
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === lowerRequested) {
      return sheets[i].getName();
    }
  }
  return null;
}
