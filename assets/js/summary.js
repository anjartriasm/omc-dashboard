document.addEventListener('DOMContentLoaded', async function () {
  bindSummaryFilters();
  await loadSummary();
});

function bindSummaryFilters() {
  ['summaryFilterNop', 'summaryFilterResponsible', 'summaryFilterClass'].forEach(function (id) {
    var element = document.getElementById(id);
    if (!element) return;
    element.addEventListener('change', function () {
      loadSummary();
    });
  });
}

async function loadSummary() {
  setSummaryState('Loading summary data...');

  var params = {
    nop: getFilterValue('summaryFilterNop'),
    responsible: getFilterValue('summaryFilterResponsible'),
    siteClass: getFilterValue('summaryFilterClass')
  };

  var response = await window.OMCApi.getSummaryData(params);

  if (!response.success) {
    console.error('Failed to load summary data:', response.message);
    setSummaryState('Failed to load summary data. Please refresh the page.', 'error');
    renderScorecards('statusCards', []);
    renderScorecards('classCards', []);
    renderSummaryTable([]);
    return;
  }

  var metrics = response.metrics || {};
  var summaryRows = response.summary || [];
  var dataRows = window.OMCTransformers.transformRows(response.data || []);
  var filters = response.filters || {};

  window.OMCUtils.setText('summaryUpdatedAt', response.updatedAt || '-');
  window.OMCUtils.setText('summaryShiftTeam', response.shiftTeam || '-');

  renderScorecards('statusCards', window.OMCTransformers.buildStatusCards(metrics));
  renderScorecards('classCards', window.OMCTransformers.buildClassCards(metrics));
  renderSummaryTable(summaryRows);
  populateFilter('summaryFilterNop', filters.nop || []);
  populateFilter('summaryFilterResponsible', filters.responsible || []);
  populateFilter('summaryFilterClass', filters.siteClass || []);

  window.OMCCharts.renderPlaceholder('siteClassChart', 'Down by Site Class', metrics.downByClass || {});
  window.OMCCharts.renderPlaceholder('responsibleChart', 'Responsible', metrics.responsible || {});
  window.OMCCharts.renderPlaceholder('mbpStatusChart', 'MBP Metrics', {
    standby: metrics.mbpStandby || 0,
    otw: metrics.mbpOtw || 0,
    backup: metrics.mbpBackup || 0
  });
  window.OMCCharts.renderPlaceholder('neStatusChart', 'NE Metrics', {
    mainsFail: metrics.mainsFail || 0,
    down: metrics.down || 0,
    up: metrics.up || 0
  });
  window.OMCMap.renderPlaceholder(dataRows.filter(function (row) {
    return !!row.siteCoord;
  }));

  if (!summaryRows.length) {
    setSummaryState('No data for selected filter.', 'empty');
  } else {
    clearSummaryState();
  }
}

function getFilterValue(elementId) {
  var element = document.getElementById(elementId);
  if (!element) return 'ALL';
  return element.value || 'ALL';
}

function renderScorecards(containerId, items) {
  var container = document.getElementById(containerId);
  if (!container) return;

  if (!(items || []).length) {
    container.innerHTML = '<div class="ui-state ui-state--empty">No metrics available.</div>';
    return;
  }

  container.innerHTML = (items || []).map(function (item) {
    return [
      '<div class="' + item.className + '">',
      '<span class="scorecard__label">' + escapeHtml(item.label) + '</span>',
      '<strong class="scorecard__value">' + escapeHtml(item.value) + '</strong>',
      '</div>'
    ].join('');
  }).join('');
}

function renderSummaryTable(rows) {
  var table = document.getElementById('summaryTable');
  if (!table) return;

  var tbody = table.querySelector('tbody');
  if (!tbody) return;

  if (!(rows || []).length) {
    tbody.innerHTML = '<tr><td colspan="11" class="empty-cell">No summary data available.</td></tr>';
    return;
  }

  tbody.innerHTML = (rows || []).map(function (row) {
    return [
      '<tr>',
      '<td>' + safeCell(row.nop) + '</td>',
      '<td>' + safeCell(row.totalMbp, '0') + '</td>',
      '<td>' + safeCell(row.mbpStandby, '0') + '</td>',
      '<td>' + safeCell(row.mbpOtw, '0') + '</td>',
      '<td>' + safeCell(row.mbpBackup, '0') + '</td>',
      '<td>' + safeCell(row.mainsFail, '0') + '</td>',
      '<td>' + safeCell(row.downEnom, '0') + '</td>',
      '<td>' + safeCell(row.downTelkom, '0') + '</td>',
      '<td>' + safeCell(row.downTp, '0') + '</td>',
      '<td>' + safeCell(row.neDown, '0') + '</td>',
      '<td>' + safeCell(row.responTsel, '0') + '</td>',
      '</tr>'
    ].join('');
  }).join('');
}

function populateFilter(elementId, values) {
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

  if (selectedValue && Array.from(select.options).some(function (opt) { return opt.value === selectedValue; })) {
    select.value = selectedValue;
  }
}

function setSummaryState(message, kind) {
  var state = document.getElementById('summaryState');
  if (!state) return;
  state.className = 'ui-state';
  if (kind === 'error') state.classList.add('ui-state--error');
  if (kind === 'empty') state.classList.add('ui-state--empty');
  state.textContent = message;
}

function clearSummaryState() {
  var state = document.getElementById('summaryState');
  if (!state) return;
  state.className = 'ui-state is-hidden';
  state.textContent = '';
}

function safeCell(value, fallback) {
  var result = window.OMCUtils.safeText(value, fallback || '-');
  return escapeHtml(result);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
