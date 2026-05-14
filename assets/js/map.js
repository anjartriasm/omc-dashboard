var MAX_SITE_POINTS = 120;
var MAX_TEAM_POINTS = 80;
var summaryLeafletMap = null;
var summaryLeafletLayer = null;

window.OMCMap = {
  renderPlaceholder(rows) {
    this.renderSummaryMap(rows || []);
  },

  renderSummaryMap(rows) {
    var element = document.getElementById('summaryMap');
    var legend = document.getElementById('summaryMapLegend');
    if (!element) return;

    var sitePoints = (rows || []).filter(function (row) { return !!row.siteCoord; }).slice(0, MAX_SITE_POINTS);
    var teamPoints = (rows || []).filter(function (row) { return !!row.teamCoord; }).slice(0, MAX_TEAM_POINTS);

    if (!sitePoints.length && !teamPoints.length) {
      teardownMap();
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

    if (!window.L) {
      teardownMap();
      element.innerHTML = [
        '<div class="map-fallback">',
        '<strong>Map library unavailable</strong>',
        '<p>Leaflet could not be loaded.</p>',
        '</div>'
      ].join('');
      if (legend) {
        legend.innerHTML = '<span class="map-legend__text">Unable to render map.</span>';
      }
      return;
    }

    element.innerHTML = '';
    ensureMap(element);
    summaryLeafletLayer.clearLayers();

    var bounds = [];

    sitePoints.forEach(function (row) {
      var coord = row.siteCoord;
      var status = window.OMCUtils.normalizeNeStatus(row.neStatus);
      var marker = window.L.circleMarker([coord.lat, coord.lng], {
        radius: 5,
        color: '#ffffff',
        weight: 1,
        fillColor: getSiteColor(status),
        fillOpacity: 0.95
      });
      marker.bindPopup([
        '<strong>Site:</strong> ' + escapeHtml(row.siteName || row.siteId || '-') + '<br>',
        '<strong>Status:</strong> ' + escapeHtml(status) + '<br>',
        '<strong>Class:</strong> ' + escapeHtml(row.siteClass || '-')
      ].join(''));
      marker.addTo(summaryLeafletLayer);
      bounds.push([coord.lat, coord.lng]);
    });

    teamPoints.forEach(function (row) {
      var coord = row.teamCoord;
      var marker = window.L.circleMarker([coord.lat, coord.lng], {
        radius: 4,
        color: '#1e3a8a',
        weight: 2,
        fillColor: '#bfdbfe',
        fillOpacity: 0.95
      });
      marker.bindPopup([
        '<strong>Team:</strong> ' + escapeHtml(row.pic || '-') + '<br>',
        '<strong>Site:</strong> ' + escapeHtml(row.siteName || row.siteId || '-') + '<br>',
        '<strong>ETA:</strong> ' + escapeHtml(row.jarakEta || '-')
      ].join(''));
      marker.addTo(summaryLeafletLayer);
      bounds.push([coord.lat, coord.lng]);
    });

    if (bounds.length === 1) {
      summaryLeafletMap.setView(bounds[0], 10);
    } else {
      summaryLeafletMap.fitBounds(bounds, { padding: [18, 18], maxZoom: 12 });
    }

    if (legend) {
      legend.innerHTML = [
        '<span><i class="map-key map-key--danger"></i>Down Site</span>',
        '<span><i class="map-key map-key--warning"></i>Mains Fail Site</span>',
        '<span><i class="map-key map-key--success"></i>UP Site</span>',
        '<span><i class="map-key map-key--team"></i>Team</span>',
        '<span class="map-legend__text">Site: ' + sitePoints.length + ' | Team: ' + teamPoints.length + '</span>'
      ].join('');
    }
  }
};

function ensureMap(element) {
  if (summaryLeafletMap) return;
  summaryLeafletMap = window.L.map(element, {
    zoomControl: true
  });

  window.L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
  }).addTo(summaryLeafletMap);

  summaryLeafletLayer = window.L.layerGroup().addTo(summaryLeafletMap);
}

function teardownMap() {
  if (!summaryLeafletMap) return;
  summaryLeafletMap.remove();
  summaryLeafletMap = null;
  summaryLeafletLayer = null;
}

function getSiteColor(status) {
  if (status === 'Mains Fail') return '#f59e0b';
  if (status === 'UP') return '#16a34a';
  return '#dc2626';
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
