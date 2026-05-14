# omc-dashboard

Dashboard Monitoring OMC untuk Summary Dashboard dan Hourly Report berbasis Google Apps Script sebagai backend dan GitHub Pages sebagai frontend preview.

## Branch aktif pengembangan
- `feature/project-structure-starter`

## Halaman yang tersedia
- `index.html`
- `pages/summary.html`
- `pages/hourly-report.html`

## Struktur utama
- `Code.gs` — backend Apps Script
- `assets/css/` — styling global dan per halaman
- `assets/js/` — config, utils, API, transformers, summary, hourly report
- `pages/` — halaman dashboard

## Endpoint backend
Set URL Apps Script web app pada file `assets/js/config.js`.

Contoh action yang digunakan:
- `getMeta`
- `getSummary`
- `getHourlyReport`
- `getMapData`

## Preview frontend
Frontend dapat direview melalui GitHub Pages dari branch pengembangan atau melalui local server.

Contoh halaman:
- `/pages/summary.html`
- `/pages/hourly-report.html`

## Status saat ini
- Summary Dashboard: tersedia
- Hourly Report: tersedia
- Chart: placeholder
- Map: placeholder
