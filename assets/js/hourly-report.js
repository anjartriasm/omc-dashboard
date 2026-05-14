var hourlyFiltersBound = false;
var HOURLY_ROWS_PER_PAGE = 20;
var hourlyCurrentPage = 1;
var hourlyRowsCache = [];
var hourlyShowAllRows = false;

document.addEventListener('DOMContentLoaded', async function () {
  bindHourlyFilters();
  bindHourlyActions();
  await loadHourlyReport();
});

function bindHourlyFilters() {
  if (hourlyFiltersBound) return;

  ['hourlyFilterNop', 'hourlyFilterResponsible', 'hourlyFilterClass', 'hourlyFilterNeStatus'].forEach(function (id) {
    var element = document.getElementById(id);
    if (!element) return;
    element.addEventListener('change', function () {
      loadHourlyReport();
    });
  });

  hourlyFiltersBound = true;
}

function bindHourlyActions() {
  var prevButton = document.getElementById('hourlyPrevPage');
  var nextButton = document.getElementById('hourlyNextPage');
  var captureButton = document.getElementById('hourlyCaptureButton');
  var showAllButton = document.getElementById('hourlyShowAllButton');

  if (prevButton) {
    prevButton.addEventListener('click', function () {
      if (hourlyCurrentPage > 1) {
        hourlyCurrentPage -= 1;
        renderHourlyTable(hourlyRowsCache);
      }
    });
  }

  if (nextButton) {
    nextButton.addEventListener('click', function () {
      var totalPages = getTotalPages(hourlyRowsCache.length);
      if (hourlyCurrentPage < totalPages) {
        hourlyCurrentPage += 1;
        renderHourlyTable(hourlyRowsCache);
      }
    });
  }

  if (captureButton) {
    captureButton.addEventListener('click', function () {
      window.print();
    });
  }

  if (showAllButton) {
    showAllButton.addEventListener('click', function () {
      hourlyShowAllRows = !hourlyShowAllRows;
      showAllButton.textContent = hourlyShowAllRows ? 'Show Paginated Rows' : 'Show All Rows';
      renderHourlyTable(hourlyRowsCache);
    });
  }
}

async function loadHourlyReport() {
  setHourlyState('Loading hourly report...');

  var params = {
    nop: getFilterValue('hourlyFilterNop'),
    responsible: getFilterValue('hourlyFilterResponsible'),
    siteClass: getFilterValue('hourlyFilterClass'),
    neStatus: getFilterValue('hourlyFilterNeStatus')
  };

  var response;
  try {
    response = await window.OMCApi.getHourlyReportData(params);
  } catch (error) {
    console.error('Hourly request error:', error);
    setHourlyState('Failed to load hourly report. Please refresh the page.', 'error');
    renderHourlyCards([]);
    renderHourlyTable([]);
    return;
  }

  if (!response.success) {
    console.error('Failed to load hourly report:', response.message);
    setHourlyState('Failed to load hourly report. Please refresh the page.', 'error');
    renderHourlyCards([]);
    renderHourlyTable([]);
    return;
  }

  var rows = window.OMCTransformers.transformRows(response.data || []);
  var filters = response.filters || {};
  hourlyRowsCache = rows;
  hourlyCurrentPage = 1;

  window.OMCUtils.setText('hourlyUpdatedAt', response.updatedAt || '-');
  window.OMCUtils.setText('hourlyTotalRows', response.total || rows.length || 0);

  populateHourlyFilter('hourlyFilterNop', filters.nop || []);
  populateHourlyFilter('hourlyFilterResponsible', filters.responsible || []);
  populateHourlyFilter('hourlyFilterClass', filters.siteClass || []);
  populateHourlyFilter('hourlyFilterNeStatus', filters.neStatus || []);

  renderHourlyCards(rows);
  renderHourlyTable(rows);

  if (!rows.length) {
    setHourlyState('No data for selected filter.', 'empty');
  } else {
    clearHourlyState();
  }
}

function getFilterValue(elementId) {
  var element = document.getElementById(elementId);
  if (!element) return 'ALL';
  return element.value || 'ALL';
}

function populateHourlyFilter(elementId, values) {
  var select = document.getElementById(elementId);
  if (!select) return;

  var selectedValue = select.value;
  var firstOption = select.options.length > 0 ? select.options[0].outerHTML : '<option value="ALL">ALL</option>';
  select.innerHTML = firstOption;

  (values || []).forEach(function (value) {
    var option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });

  var hasSelected = Array.from(select.options).some(function (option) {
    return option.value === selectedValue;
  });

  if (hasSelected) {
    select.value = selectedValue;
  }
}

function renderHourlyCards(rows) {
  var container = document.getElementById('hourlyStatusCards');
  if (!container) return;

  var counters = {
    total: rows.length,
    down: 0,
    mainsFail: 0,
    backup: 0
  };

  rows.forEach(function (row) {
    if (row.neStatus === 'Down') counters.down += 1;
    if (row.neStatus === 'Mains Fail') counters.mainsFail += 1;
    if (row.mbpStatus === 'Backup') counters.backup += 1;
  });

  var cards = [
    { label: 'TOTAL', value: counters.total, className: 'scorecard scorecard--info' },
    { label: 'DOWN', value: counters.down, className: 'scorecard scorecard--danger' },
    { label: 'PLN OFF', value: counters.mainsFail, className: 'scorecard scorecard--warning' },
    { label: 'BACKUP', value: counters.backup, className: 'scorecard scorecard--success' }
  ];

  container.innerHTML = cards.map(function (card) {
    return [
      '<div class="' + card.className + '">',
      '<span class="scorecard__label">' + escapeHtml(card.label) + '</span>',
      '<strong class="scorecard__value">' + escapeHtml(card.value) + '</strong>',
      '</div>'
    ].join('');
  }).join('');
}

function renderHourlyTable(rows) {
  var table = document.getElementById('hourlyTable');
  if (!table) return;

  var tbody = table.querySelector('tbody');
  if (!tbody) return;

  if (!(rows || []).length) {
    tbody.innerHTML = '<tr><td colspan="16" class="empty-cell">No hourly data available.</td></tr>';
    updatePagination(0, 0, 0, 0);
    return;
  }

  var totalRows = rows.length;
  var totalPages = getTotalPages(totalRows);
  if (hourlyCurrentPage > totalPages) hourlyCurrentPage = totalPages;

  var rowsPerPage = hourlyShowAllRows ? totalRows : HOURLY_ROWS_PER_PAGE;
  var startIndex = hourlyShowAllRows ? 0 : (hourlyCurrentPage - 1) * HOURLY_ROWS_PER_PAGE;
  var endIndex = Math.min(startIndex + rowsPerPage, totalRows);
  var pageRows = rows.slice(startIndex, endIndex);

  tbody.innerHTML = pageRows.map(function (row) {
    return [
      '<tr>',
      '<td>' + safeCell(row.siteId) + '</td>',
      '<td>' + safeCell(row.siteName) + '</td>',
      '<td>' + safeCell(row.nop) + '</td>',
      '<td>' + safeCell(row.to) + '</td>',
      '<td>' + safeCell(row.siteClass) + '</td>',
      '<td>' + safeCell(row.tp) + '</td>',
      '<td>' + renderStatusBadge(row.neStatus) + '</td>',
      '<td>' + safeCell(row.severity) + '</td>',
      '<td>' + renderMbpBadge(row.mbpStatus) + '</td>',
      '<td>' + safeCell(row.responsible) + '</td>',
      '<td>' + safeCell(row.alarmStart) + '</td>',
      '<td>' + safeCell(row.duration) + '</td>',
      '<td class="hourly-remark-cell">' + safeCell(row.remark) + '</td>',
      '<td>' + safeCell(row.jarakEta) + '</td>',
      '<td>' + safeCell(row.kabupaten) + '</td>',
      '<td>' + safeCell(row.pic) + '</td>',
      '</tr>'
    ].join('');
  }).join('');

  updatePagination(startIndex, endIndex, totalRows, totalPages);
}

function updatePagination(startIndex, endIndex, totalRows, totalPages) {
  var info = document.getElementById('hourlyPaginationInfo');
  var indicator = document.getElementById('hourlyPageIndicator');
  var prevButton = document.getElementById('hourlyPrevPage');
  var nextButton = document.getElementById('hourlyNextPage');

  if (info) {
    if (!totalRows) {
      info.textContent = 'Showing 0 of 0';
    } else if (hourlyShowAllRows) {
      info.textContent = 'Showing all ' + totalRows + ' rows';
    } else {
      info.textContent = 'Showing ' + (startIndex + 1) + '–' + endIndex + ' of ' + totalRows;
    }
  }

  if (indicator) {
    indicator.textContent = hourlyShowAllRows ? 'All rows' : hourlyCurrentPage + ' / ' + Math.max(totalPages, 1);
  }

  if (prevButton) prevButton.disabled = hourlyShowAllRows || hourlyCurrentPage <= 1;
  if (nextButton) nextButton.disabled = hourlyShowAllRows || hourlyCurrentPage >= totalPages;
}

function getTotalPages(totalRows) {
  if (!totalRows) return 1;
  return Math.max(1, Math.ceil(totalRows / HOURLY_ROWS_PER_PAGE));
}

function renderStatusBadge(value) {
  var text = value || '-';
  var className = 'hourly-badge hourly-badge--info';

  if (text === 'UP') className = 'hourly-badge hourly-badge--success';
  else if (text === 'Mains Fail') className = 'hourly-badge hourly-badge--warning';
  else if (text === 'Down') className = 'hourly-badge hourly-badge--danger';

  return '<span class="' + className + '">' + escapeHtml(text) + '</span>';
}

function renderMbpBadge(value) {
  var text = value || '-';
  var className = 'hourly-badge hourly-badge--info';

  if (text === 'Standby') className = 'hourly-badge hourly-badge--info';
  else if (text === 'OTW') className = 'hourly-badge hourly-badge--warning';
  else if (text === 'Backup') className = 'hourly-badge hourly-badge--success';
  else if (text === 'LOS') className = 'hourly-badge hourly-badge--danger';

  return '<span class="' + className + '">' + escapeHtml(text) + '</span>';
}

function safeCell(value) {
  if (value === null || value === undefined || value === '') {
    return '-';
  }
  return escapeHtml(value);
}

function setHourlyState(message, kind) {
  var state = document.getElementById('hourlyState');
  if (!state) return;
  state.className = 'ui-state';
  if (kind === 'error') state.classList.add('ui-state--error');
  if (kind === 'empty') state.classList.add('ui-state--empty');
  state.textContent = message;
}

function clearHourlyState() {
  var state = document.getElementById('hourlyState');
  if (!state) return;
  state.className = 'ui-state is-hidden';
  state.textContent = '';
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
