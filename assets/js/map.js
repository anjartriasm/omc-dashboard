window.OMCMap = {
  renderPlaceholder(rows) {
    var element = document.getElementById('summaryMap');
    var legend = document.getElementById('summaryMapLegend');
    if (!element) return;

    var total = (rows || []).length;
    element.innerHTML = '<div><strong>Map placeholder</strong><p>Total points: ' + total + '</p></div>';

    if (legend) {
      legend.innerHTML = [
        '<span>🔴 Down</span>',
        '<span>🟡 Mains Fail</span>',
        '<span>🟢 UP</span>'
      ].join(' ');
    }
  }
};
