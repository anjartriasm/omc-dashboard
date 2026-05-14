window.OMCApi = {
  async request(endpoint, queryParams) {
    var config = window.OMC_CONFIG;
    var url = config.appsScript.baseUrl + endpoint;

    if (queryParams && typeof queryParams === 'object') {
      var searchParams = new URLSearchParams(queryParams);
      var separator = url.indexOf('?') > -1 ? '&' : '?';
      url += separator + searchParams.toString();
    }

    try {
      var response = await fetch(url, {
        method: 'GET',
        headers: {
          Accept: 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error('HTTP ' + response.status);
      }

      return await response.json();
    } catch (error) {
      console.error('OMC API request failed:', error);
      return {
        success: false,
        message: error.message || 'Request failed',
        data: []
      };
    }
  },

  getMeta() {
    return this.request(window.OMC_CONFIG.appsScript.endpoints.meta);
  },

  getSummaryData(params) {
    return this.request(window.OMC_CONFIG.appsScript.endpoints.summary, params || {});
  },

  getHourlyReportData(params) {
    return this.request(window.OMC_CONFIG.appsScript.endpoints.hourly, params || {});
  },

  getMapData(params) {
    return this.request(window.OMC_CONFIG.appsScript.endpoints.map, params || {});
  }
};
