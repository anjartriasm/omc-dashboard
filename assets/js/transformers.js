window.OMCTransformers = {
  transformRow(row) {
    return {
      siteId: row.siteId || '',
      siteName: row.siteName || '',
      nop: row.nop || '',
      to: row.to || '',
      siteClass: window.OMCUtils.normalizeSiteClass(row.siteClass),
      tp: row.tp || '',
      neStatus: window.OMCUtils.normalizeNeStatus(row.neStatus),
      severity: row.severity || '',
      mbpStatus: window.OMCUtils.normalizeMbpStatus(row.mbpStatus),
      responsible: row.responsible || '',
      alarmStart: row.alarmStart || '',
      duration: row.duration || '',
      remark: row.remark || '',
      jarakEta: row.jarakEta || '',
      kabupaten: row.kabupaten || '',
      siteLatLong: row.siteLatLong || '',
      teamLatLong: row.teamLatLong || '',
      pic: row.pic || '',
      siteCoord: window.OMCUtils.parseLatLong(row.siteLatLong || ''),
      teamCoord: window.OMCUtils.parseLatLong(row.teamLatLong || '')
    };
  },

  transformRows(rows) {
    return (rows || []).map(this.transformRow.bind(this));
  },

  buildStatusCards(metrics) {
    return [
      { label: 'PLN OFF', value: metrics.mainsFail || 0, className: 'scorecard scorecard--warning' },
      { label: 'DOWN', value: metrics.down || 0, className: 'scorecard scorecard--danger' },
      { label: 'NE DOWN', value: metrics.neDown || 0, className: 'scorecard scorecard--info' },
      { label: 'BACKUP', value: metrics.mbpBackup || 0, className: 'scorecard scorecard--success' }
    ];
  },

  buildClassCards(metrics) {
    var downByClass = (metrics && metrics.downByClass) || {};
    return [
      { label: 'DIAMOND', value: downByClass.Diamond || 0, className: 'scorecard scorecard--diamond' },
      { label: 'PLATINUM', value: downByClass.Platinum || 0, className: 'scorecard scorecard--platinum' },
      { label: 'GOLD', value: downByClass.Gold || 0, className: 'scorecard scorecard--gold' },
      { label: 'SILVER', value: downByClass.Silver || 0, className: 'scorecard scorecard--silver' },
      { label: 'BRONZE', value: downByClass.Bronze || 0, className: 'scorecard scorecard--bronze' }
    ];
  }
};
