var MAX_SITE_POINTS = 120;
var MAX_TEAM_POINTS = 80;

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
  },

  renderSummaryMap(rows) {
    var element = document.getElementById('summaryMap');
    var legend = document.getElementById('summaryMapLegend');
    if (!element) return;

    var sitePoints = (rows || []).filter(function (row) { return !!row.siteCoord; });
    var teamPoints = (rows || []).filter(function (row) { return !!row.teamCoord; });

    if (!sitePoints.length && !teamPoints.length) {
      element.innerHTML = [
        '<div class="map-fallback">',
        '<strong>Map data not available</strong>',
        '<p>Site and team coordinates are not available for current filter.</p>',
        '</div>'
      ].join('');
      if (legend) {
        legend.innerHTML = '<span class="map-legend__text">No coordinates found.</span>';
      }
      return;
    }

    var combined = sitePoints.map(function (row) { return row.siteCoord; }).concat(
      teamPoints.map(function (row) { return row.teamCoord; })
    );

    var bounds = getBounds(combined);
    var maxSitePoints = sitePoints.slice(0, MAX_SITE_POINTS);
    var maxTeamPoints = teamPoints.slice(0, MAX_TEAM_POINTS);

    element.innerHTML = [
      '<div class="map-canvas" role="img" aria-label="Site and team points map">',
      maxSitePoints.map(function (row) {
        var point = normalizePoint(row.siteCoord, bounds);
        return [
          '<span class="map-dot map-dot--site map-dot--' + normalizeStatusClass(row.neStatus) + '" style="left:',
          point.x,
          '%;top:',
          point.y,
          '%" title="Site ',
          escapeHtml(row.siteName || row.siteId || '-'),
          '"></span>'
        ].join('');
      }).join(''),
      maxTeamPoints.map(function (row) {
        var point = normalizePoint(row.teamCoord, bounds);
        return [
          '<span class="map-dot map-dot--team" style="left:',
          point.x,
          '%;top:',
          point.y,
          '%" title="Team ',
          escapeHtml(row.pic || '-'),
          '"></span>'
        ].join('');
      }).join(''),
      '</div>'
    ].join('');

    if (legend) {
      legend.innerHTML = [
        '<span><i class="map-key map-key--danger"></i>Down</span>',
        '<span><i class="map-key map-key--warning"></i>Mains Fail</span>',
        '<span><i class="map-key map-key--success"></i>UP</span>',
        '<span><i class="map-key map-key--team"></i>Team</span>',
        '<span class="map-legend__text">Site: ' + sitePoints.length + ' | Team: ' + teamPoints.length + '</span>'
      ].join('');
    }
  }
};

function getBounds(points) {
  var minLat = 90;
  var maxLat = -90;
  var minLng = 180;
  var maxLng = -180;

  (points || []).forEach(function (point) {
    if (!point) return;
    minLat = Math.min(minLat, point.lat);
    maxLat = Math.max(maxLat, point.lat);
    minLng = Math.min(minLng, point.lng);
    maxLng = Math.max(maxLng, point.lng);
  });

  return {
    minLat: minLat,
    maxLat: maxLat,
    minLng: minLng,
    maxLng: maxLng
  };
}

function normalizePoint(coord, bounds) {
  var latRange = Math.max(0.001, bounds.maxLat - bounds.minLat);
  var lngRange = Math.max(0.001, bounds.maxLng - bounds.minLng);
  var x = ((coord.lng - bounds.minLng) / lngRange) * 100;
  var y = (1 - ((coord.lat - bounds.minLat) / latRange)) * 100;

  return {
    x: clamp(x, 2, 98).toFixed(2),
    y: clamp(y, 3, 97).toFixed(2)
  };
}

function normalizeStatusClass(status) {
  if (status === 'Mains Fail') return 'warning';
  if (status === 'UP') return 'success';
  return 'danger';
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
