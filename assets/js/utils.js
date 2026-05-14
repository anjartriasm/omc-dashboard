window.OMCUtils = {
  safeText(value, fallback) {
    if (value === null || value === undefined || value === '') {
      return fallback || '-';
    }
    return String(value).trim();
  },

  parseLatLong(value) {
    if (!value || typeof value !== 'string' || value.indexOf(',') === -1) {
      return null;
    }

    var parts = value.split(',');
    var lat = Number(String(parts[0]).trim());
    var lng = Number(String(parts[1]).trim());

    if (Number.isNaN(lat) || Number.isNaN(lng)) {
      return null;
    }

    return { lat: lat, lng: lng };
  },

  normalizeNeStatus(value) {
    var text = this.safeText(value, '').toUpperCase();

    if (text === 'UP' || text === 'SITE UP') return 'UP';
    if (text.indexOf('MAINS FAIL') > -1 || text.indexOf('PLN OFF') > -1) return 'Mains Fail';
    if (text.indexOf('DOWN') > -1 || text.indexOf('LOS') > -1) return 'Down';

    return this.safeText(value);
  },

  normalizeMbpStatus(value) {
    var text = this.safeText(value, '').toUpperCase();

    if (text === 'STANDBY') return 'Standby';
    if (text === 'OTW') return 'OTW';
    if (text === 'BACKUP') return 'Backup';
    if (text === 'LOS') return 'LOS';

    return this.safeText(value);
  },

  normalizeSiteClass(value) {
    var text = this.safeText(value, '');
    if (!text || text === '-') return '';
    return text.charAt(0).toUpperCase() + text.slice(1).toLowerCase();
  },

  getStatusBadgeClass(status) {
    var normalized = this.normalizeNeStatus(status);
    if (normalized === 'Down') return 'badge badge--danger';
    if (normalized === 'Mains Fail') return 'badge badge--warning';
    if (normalized === 'UP') return 'badge badge--success';
    return 'badge badge--neutral';
  },

  uniqueValues(rows, key) {
    var map = {};
    rows.forEach(function(row) {
      var value = row[key];
      if (value) map[value] = true;
    });
    return Object.keys(map).sort();
  },

  setText(elementId, value, fallback) {
    var el = document.getElementById(elementId);
    if (!el) return;
    el.textContent = this.safeText(value, fallback || '-');
  }
};
