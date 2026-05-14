document.addEventListener('DOMContentLoaded', async function () {
  bindHourlyFilters();
  await loadHourlyReport();
});

function bindHourlyFilters() {
  ['hourlyFilterNop', 'hourlyFilterResponsible', 'hourlyFilterClass', 'hourlyFilterNeStatus'].forEach(function (id) {
    var element = document.getElementById(id);
    if (!element) return;
    element.addEventListener('change', function () {
      loadHourlyReport();
    });
  });
}

async function loadHourlyReport() {
  setHourlyState('Loading hourly report...');

  var params = {
    nop: getFilterValue('hourlyFilterNop'),
    responsible: getFilterValue('hourlyFilterResponsible'),
    siteClass: getFilterValue('hourlyFilterClass'),
    neStatus: getFilterValue('hourlyFilterNeStatus')
  };

  var response = await window.OMCApi.getHourlyReportData(params);

  if (!response.success) {
    console.error('Failed to load hourly report:', response.message);
    setHourlyState('Failed to load hourly report. Please refresh the page.', 'error');
    renderHourlyCards([]);
    renderHourlyTable([]);
    return;
  }

  var rows = window.OMCTransformers.transformRows(response.data || []);
  var filters = response.filters || {};

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
    return;
  }

  tbody.innerHTML = (rows || []).map(function (row) {
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
      '<td>' + safeCell(row.remark) + '</td>',
      '<td>' + safeCell(row.jarakEta) + '</td>',
      '<td>' + safeCell(row.kabupaten) + '</td>',
      '<td>' + safeCell(row.pic) + '</td>',
      '</tr>'
    ].join('');
  }).join('');
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
