window.OMCCharts = {
  renderPlaceholder(elementId, title, data) {
    var element = document.getElementById(elementId);
    if (!element) return;

    element.innerHTML = [
      '<div style="width:100%;text-align:left;">',
      '<strong>' + title + '</strong>',
      '<pre style="margin-top:8px;font-size:12px;white-space:pre-wrap;">' +
        JSON.stringify(data, null, 2) +
      '</pre>',
      '</div>'
    ].join('');
  },

  renderMetricBars(elementId, data, options) {
    var element = document.getElementById(elementId);
    if (!element) return;

    var entries = Object.keys(data || {}).map(function (key) {
      return [key, Number(data[key] || 0)];
    }).filter(function (entry) {
      return entry[1] >= 0;
    });

    entries.sort(function (a, b) {
      return b[1] - a[1];
    });

    var maxItems = (options && options.maxItems) || entries.length;
    var selectedEntries = entries.slice(0, maxItems);
    var maxValue = selectedEntries.reduce(function (result, entry) {
      return Math.max(result, entry[1]);
    }, 0);

    if (!selectedEntries.length) {
      element.innerHTML = '<div class="ui-state ui-state--empty">' + escapeHtml((options && options.emptyLabel) || 'No metrics available.') + '</div>';
      return;
    }

    element.innerHTML = selectedEntries.map(function (entry) {
      var width = maxValue > 0 ? Math.max(6, Math.round((entry[1] / maxValue) * 100)) : 0;
      return [
        '<div class="metric-row">',
        '<div class="metric-row__top">',
        '<span class="metric-row__label">' + escapeHtml(entry[0]) + '</span>',
        '<strong class="metric-row__value">' + escapeHtml(entry[1]) + '</strong>',
        '</div>',
        '<div class="metric-row__bar"><span style="width:' + width + '%"></span></div>',
        '</div>'
      ].join('');
    }).join('');
  }
};

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
