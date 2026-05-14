function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || '';

  if (params.path === 'sw.js') {
    return ContentService.createTextOutput(getServiceWorkerCode())
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  if (action) {
    return handleApiRequest(action, params);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('OMC Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleApiRequest(action, params) {
  try {
    var payload;

    switch (action) {
      case 'getSummary':
        payload = getSummaryPayload(params);
        break;
      case 'getHourlyReport':
        payload = getHourlyReportPayload(params);
        break;
      case 'getMapPoints':
        payload = getMapPointsPayload(params);
        break;
      case 'getMeta':
        payload = getMetaPayload();
        break;
      default:
        payload = {
          success: false,
          message: 'Unknown action: ' + action,
          availableActions: ['getSummary', 'getHourlyReport', 'getMapPoints', 'getMeta']
        };
        break;
    }

    return createJsonOutput(payload);
  } catch (error) {
    return createJsonOutput({
      success: false,
      message: error.message || 'Unexpected server error'
    });
  }
}

function createJsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getServiceWorkerCode() {
  return `const CACHE_NAME = 'omc-dashboard-v2';
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
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE))
  );
});

self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request).then((response) => response || fetch(event.request))
  );
});`;
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getMetaPayload() {
  return {
    success: true,
    appUrl: getAppUrl(),
    updatedAt: getCurrentTimestamp(),
    sheetName: 'OMC_R5'
  };
}

function getSummaryPayload(params) {
  var rows = getDashboardRows();
  var filteredRows = filterRows(rows, params);
  var summaryRows = buildSummaryByNop(filteredRows);
  var metrics = buildGlobalMetrics(filteredRows);

  return {
    success: true,
    updatedAt: getCurrentTimestamp(),
    shiftTeam: params.shiftTeam || '-',
    filters: buildFilterOptions(rows),
    metrics: metrics,
    summary: summaryRows,
    data: filteredRows
  };
}

function getHourlyReportPayload(params) {
  var rows = getDashboardRows();
  var filteredRows = filterRows(rows, params);

  return {
    success: true,
    updatedAt: getCurrentTimestamp(),
    total: filteredRows.length,
    filters: buildFilterOptions(rows),
    data: filteredRows
  };
}

function getMapPointsPayload(params) {
  var rows = getDashboardRows();
  var filteredRows = filterRows(rows, params).filter(function(row) {
    return !!row.siteLatLong;
  });

  return {
    success: true,
    updatedAt: getCurrentTimestamp(),
    total: filteredRows.length,
    data: filteredRows.map(function(row) {
      return {
        siteId: row.siteId,
        siteName: row.siteName,
        nop: row.nop,
        siteClass: row.siteClass,
        neStatus: row.neStatus,
        severity: row.severity,
        mbpStatus: row.mbpStatus,
        responsible: row.responsible,
        jarakEta: row.jarakEta,
        remark: row.remark,
        pic: row.pic,
        siteLatLong: row.siteLatLong,
        teamLatLong: row.teamLatLong
      };
    })
  };
}

function getDashboardRows() {
  var sheet = getRequiredSheet_('OMC_R5');
  var values = sheet.getDataRange().getDisplayValues();

  if (!values || values.length < 2) {
    return [];
  }

  var headers = values[0].map(normalizeHeaderKey_);
  return values.slice(1)
    .filter(function(row) {
      return row.join('').trim() !== '';
    })
    .map(function(row) {
      return mapRowToObject_(headers, row);
    });
}

function buildSummaryByNop(rows) {
  var grouped = {};

  rows.forEach(function(row) {
    var nopKey = sanitizeText_(row.nop, 'UNKNOWN');
    if (!grouped[nopKey]) {
      grouped[nopKey] = {
        nop: nopKey,
        totalMbp: 0,
        mbpStandby: 0,
        mbpOtw: 0,
        mbpBackup: 0,
        mainsFail: 0,
        downEnom: 0,
        downTelkom: 0,
        downTp: 0,
        neDown: 0,
        responTsel: 0
      };
    }

    grouped[nopKey].totalMbp += 1;

    var mbpStatus = normalizeMbpStatus_(row.mbpStatus);
    var neStatus = normalizeNeStatus_(row.neStatus);
    var responsible = sanitizeText_(row.responsible).toUpperCase();

    if (mbpStatus === 'Standby') grouped[nopKey].mbpStandby += 1;
    if (mbpStatus === 'OTW') grouped[nopKey].mbpOtw += 1;
    if (mbpStatus === 'Backup') grouped[nopKey].mbpBackup += 1;
    if (neStatus === 'Mains Fail') grouped[nopKey].mainsFail += 1;

    if (neStatus === 'Down') {
      grouped[nopKey].neDown += 1;
      if (responsible.indexOf('ENOM') > -1) grouped[nopKey].downEnom += 1;
      if (responsible.indexOf('AKSES') > -1 || responsible.indexOf('TELKOM') > -1) grouped[nopKey].downTelkom += 1;
      if (responsible === 'TP') grouped[nopKey].downTp += 1;
    }
  });

  return Object.keys(grouped)
    .sort()
    .map(function(key) {
      return grouped[key];
    });
}

function buildGlobalMetrics(rows) {
  var metrics = {
    totalMbp: rows.length,
    mbpStandby: 0,
    mbpOtw: 0,
    mbpBackup: 0,
    mainsFail: 0,
    down: 0,
    neDown: 0,
    up: 0,
    downByClass: {
      Diamond: 0,
      Platinum: 0,
      Gold: 0,
      Silver: 0,
      Bronze: 0
    },
    responsible: {}
  };

  rows.forEach(function(row) {
    var mbpStatus = normalizeMbpStatus_(row.mbpStatus);
    var neStatus = normalizeNeStatus_(row.neStatus);
    var siteClass = sanitizeText_(row.siteClass);
    var responsible = sanitizeText_(row.responsible, 'UNKNOWN');

    if (mbpStatus === 'Standby') metrics.mbpStandby += 1;
    if (mbpStatus === 'OTW') metrics.mbpOtw += 1;
    if (mbpStatus === 'Backup') metrics.mbpBackup += 1;

    if (neStatus === 'Mains Fail') metrics.mainsFail += 1;
    if (neStatus === 'Down') {
      metrics.down += 1;
      metrics.neDown += 1;
      if (metrics.downByClass[siteClass] !== undefined) {
        metrics.downByClass[siteClass] += 1;
      }
    }
    if (neStatus === 'UP') metrics.up += 1;

    if (!metrics.responsible[responsible]) {
      metrics.responsible[responsible] = 0;
    }
    metrics.responsible[responsible] += 1;
  });

  return metrics;
}

function buildFilterOptions(rows) {
  return {
    nop: uniqueSortedValues_(rows, 'nop'),
    responsible: uniqueSortedValues_(rows, 'responsible'),
    siteClass: uniqueSortedValues_(rows, 'siteClass'),
    neStatus: uniqueSortedValues_(rows.map(function(row) {
      row.neStatus = normalizeNeStatus_(row.neStatus);
      return row;
    }), 'neStatus')
  };
}

function filterRows(rows, params) {
  return rows.filter(function(row) {
    if (params.nop && params.nop !== 'ALL' && sanitizeText_(row.nop) !== params.nop) {
      return false;
    }

    if (params.responsible && params.responsible !== 'ALL' && sanitizeText_(row.responsible) !== params.responsible) {
      return false;
    }

    if (params.siteClass && params.siteClass !== 'ALL' && sanitizeText_(row.siteClass) !== params.siteClass) {
      return false;
    }

    if (params.neStatus && params.neStatus !== 'ALL' && normalizeNeStatus_(row.neStatus) !== params.neStatus) {
      return false;
    }

    return true;
  });
}

function mapRowToObject_(headers, row) {
  var obj = {};
  headers.forEach(function(header, index) {
    obj[header] = row[index] || '';
  });

  return {
    siteId: sanitizeText_(obj.siteId),
    siteName: sanitizeText_(obj.siteName),
    nop: sanitizeText_(obj.nop),
    to: sanitizeText_(obj.to),
    siteClass: normalizeSiteClass_(obj.siteClass),
    tp: sanitizeText_(obj.tp),
    neStatus: normalizeNeStatus_(obj.neStatus),
    severity: sanitizeText_(obj.severity),
    mbpStatus: normalizeMbpStatus_(obj.mbpStatus),
    responsible: sanitizeText_(obj.responsible),
    alarmStart: sanitizeText_(obj.alarmStart),
    duration: sanitizeText_(obj.duration),
    remark: sanitizeText_(obj.remark),
    jarakEta: sanitizeText_(obj.jarakEta),
    kabupaten: sanitizeText_(obj.kabupaten),
    siteLatLong: sanitizeText_(obj.siteLatLong),
    teamLatLong: sanitizeText_(obj.teamLatLong),
    pic: sanitizeText_(obj.pic)
  };
}

function getRequiredSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' tidak ditemukan");
  }
  return sheet;
}

function normalizeHeaderKey_(header) {
  var value = String(header || '').trim();
  var directMap = {
    siteid: 'siteId',
    sitename: 'siteName',
    nop: 'nop',
    to: 'to',
    siteclass: 'siteClass',
    tp: 'tp',
    nestatus: 'neStatus',
    severity: 'severity',
    mbpstatus: 'mbpStatus',
    responsible: 'responsible',
    alarmstart: 'alarmStart',
    duration: 'duration',
    remark: 'remark',
    jaraketa: 'jarakEta',
    kabupaten: 'kabupaten',
    sitelatlong: 'siteLatLong',
    teamlatlong: 'teamLatLong',
    pic: 'pic'
  };

  var compact = value.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  if (directMap[compact]) {
    return directMap[compact];
  }

  var camel = value
    .replace(/[^a-zA-Z0-9]+(.)/g, function(match, chr) {
      return chr.toUpperCase();
    })
    .replace(/[^a-zA-Z0-9]/g, '');

  return camel ? camel.charAt(0).toLowerCase() + camel.slice(1) : '';
}

function normalizeNeStatus_(value) {
  var text = sanitizeText_(value).toUpperCase();

  if (text === 'SITE UP' || text === 'UP') return 'UP';
  if (text.indexOf('MAINS FAIL') > -1 || text.indexOf('PLN OFF') > -1) return 'Mains Fail';
  if (text.indexOf('DOWN') > -1 || text.indexOf('LOS') > -1) return 'Down';

  return sanitizeText_(value);
}

function normalizeMbpStatus_(value) {
  var text = sanitizeText_(value).toUpperCase();

  if (text === 'STANDBY') return 'Standby';
  if (text === 'OTW') return 'OTW';
  if (text === 'BACKUP') return 'Backup';
  if (text === 'LOS') return 'LOS';

  return sanitizeText_(value);
}

function normalizeSiteClass_(value) {
  var text = sanitizeText_(value).toLowerCase();
  if (!text) return '';
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function sanitizeText_(value, fallback) {
  if (value === null || value === undefined || value === '') {
    return fallback || '';
  }
  return String(value).trim();
}

function uniqueSortedValues_(rows, key) {
  var map = {};
  rows.forEach(function(row) {
    var value = sanitizeText_(row[key]);
    if (value) {
      map[value] = true;
    }
  });
  return Object.keys(map).sort();
}

function getCurrentTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
}

function checkLogin(auth) {
  var adminUser = 'admin';
  var adminPass = 'admin';

  if (auth.username.trim() === adminUser && auth.password.trim() === adminPass) {
    return { success: true };
  }

  return { success: false, message: 'Username atau Password salah!' };
}
