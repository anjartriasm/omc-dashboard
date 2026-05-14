document.addEventListener('DOMContentLoaded', async function () {
  var response = await window.OMCApi.getSummaryData();

  if (!response.success) {
    console.error('Failed to load summary data:', response.message);
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
});

function renderScorecards(containerId, items) {
  var container = document.getElementById(containerId);
  if (!container) return;

  container.innerHTML = (items || []).map(function (item) {
    return [
      '<div class="' + item.className + '">',
      '<span class="scorecard__label">' + item.label + '</span>',
      '<strong class="scorecard__value">' + item.value + '</strong>',
      '</div>'
    ].join('');
  }).join('');
}

function renderSummaryTable(rows) {
  var table = document.getElementById('summaryTable');
  if (!table) return;

  var tbody = table.querySelector('tbody');
  if (!tbody) return;

  tbody.innerHTML = (rows || []).map(function (row) {
    return [
      '<tr>',
      '<td>' + (row.nop || '-') + '</td>',
      '<td>' + (row.totalMbp || 0) + '</td>',
      '<td>' + (row.mbpStandby || 0) + '</td>',
      '<td>' + (row.mbpOtw || 0) + '</td>',
      '<td>' + (row.mbpBackup || 0) + '</td>',
      '<td>' + (row.mainsFail || 0) + '</td>',
      '<td>' + (row.downEnom || 0) + '</td>',
      '<td>' + (row.downTelkom || 0) + '</td>',
      '<td>' + (row.downTp || 0) + '</td>',
      '<td>' + (row.neDown || 0) + '</td>',
      '<td>' + (row.responTsel || 0) + '</td>',
      '</tr>'
    ].join('');
  }).join('');
}

function populateFilter(elementId, values) {
  var select = document.getElementById(elementId);
  if (!select) return;

  (values || []).forEach(function (value) {
    var option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
}
