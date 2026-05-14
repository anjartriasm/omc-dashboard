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
  }
};
