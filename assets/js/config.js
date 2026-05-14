window.OMC_CONFIG = {
  appName: 'OMC Dashboard',
  refreshIntervalMs: 60000,
  timezone: 'Asia/Jakarta',
  appsScript: {
    baseUrl: 'https://script.google.com/macros/s/AKfycbxmrNlhlgaUGOkG3lAoS-K-bQRzlED8XyDGHW0ueYW0lQVwfvPn2nCJcrZgEjWSnlE6zw/exec',
    endpoints: {
      summary: '?action=getSummary',
      hourly: '?action=getHourlyReport',
      map: '?action=getMapPoints',
      meta: '?action=getMeta'
    }
  }
};
