function doGet(e) {
  // Handle service worker request
  if (e && e.parameter && e.parameter.path === 'sw.js') {
    return ContentService.createTextOutput(getServiceWorkerCode())
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Menampilkan file index.html saat URL Web App dibuka
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Network Monitoring Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getServiceWorkerCode() {
  return `const CACHE_NAME = 'omc-enom-v1';
const ASSETS_TO_CACHE = [
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
  'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.min.js',
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    })
  );
});

self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request).then((response) => {
      return response || fetch(event.request);
    })
  );
});`;
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getDataDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); // Membuka file Spreadsheet aktif
    var sheet = ss.getSheetByName("OMC_R5");        // Mengambil data dari sheet bernama OMC_R5
    if (!sheet) throw new Error("Sheet 'OMC_R5' tidak ditemukan");
    return sheet.getDataRange().getDisplayValues(); // Mengambil semua data dalam bentuk teks (string)
  } catch (e) {
    Logger.log("Error getDataDashboard: " + e);
    throw e;
  }
}

function getDataSheet2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Ticket_SWFM"); // Ganti "Sheet2" dengan nama sheet Anda
    if (!sheet) throw new Error("Sheet 'Ticket_SWFM' tidak ditemukan");
    const data = sheet.getDataRange().getDisplayValues();
    Logger.log("Data Sheet2: " + data.length + " rows");
    return data;
  } catch(e) {
    Logger.log("Error getDataSheet2: " + e);
    throw e;
  }
}

function getDataSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("OMC_R5");
    if (!sheet) throw new Error("Sheet 'OMC_R5' tidak ditemukan");
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const headers = data[0];
    const nopIdx = headers.indexOf('NOP');
    const statusIdx = headers.indexOf('NE STATUS');
    const mbpIdx = headers.indexOf('MBP STATUS');
    const responsibleIdx = headers.indexOf('RESPONSIBLE');

    const summary = {};
    const nops = [];

    data.slice(1).forEach(row => {
      const nop = row[nopIdx] || 'UNKNOWN';
      if (!summary[nop]) {
        summary[nop] = { 
          'NOP': nop, 
          'MBP': 0, 
          'DOWN': 0, 
          'OTW': 0, 
          'MAINS FAIL': 0, 
          'TP': 0, 
          'TELKOM': 0, 
          'ENOM': 0 
        };
        nops.push(nop);
      }
      
      const status = String(row[statusIdx]).toUpperCase();
      const mbp = String(row[mbpIdx]).toUpperCase();
      const resp = String(row[responsibleIdx]).toUpperCase();

      if (status.includes('DOWN')) summary[nop]['DOWN']++;
      if (status.includes('MAINS FAIL')) summary[nop]['MAINS FAIL']++;
      if (mbp.includes('OTW')) summary[nop]['OTW']++;
      
      if (resp.includes('TP')) summary[nop]['TP']++;
      if (resp.includes('TELKOM')) summary[nop]['TELKOM']++;
      if (resp.includes('ENOM')) summary[nop]['ENOM']++;
      
      summary[nop]['MBP']++; // Total MBP in this NOP
    });

    const result = [['NOP', 'MBP', 'DOWN', 'OTW', 'MAINS FAIL', 'TP', 'TELKOM', 'ENOM']];
    nops.sort().forEach(nop => {
      const s = summary[nop];
      result.push([s['NOP'], s['MBP'], s['DOWN'], s['OTW'], s['MAINS FAIL'], s['TP'], s['TELKOM'], s['ENOM']]);
    });

    return result;
  } catch (e) {
    Logger.log("Error getDataSummary: " + e);
    throw e;
  }
}

function checkLogin(auth) {
  var adminUser = "admin";         // Username yang ditentukan
  var adminPass = "admin";   // Password yang ditentukan
  
  // Memeriksa apakah input dari user sama dengan kredensial di atas
  if (auth.username.trim() === adminUser && auth.password.trim() === adminPass) {
    return { success: true };
  } else {
    return { success: false, message: "Username atau Password salah!" };
  }
}
