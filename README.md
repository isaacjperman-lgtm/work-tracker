# Work Tracker

A simple installable web app (PWA) for tracking work time on an Android phone and exporting a monthly invoice spreadsheet.

**Live app:** https://isaacjperman-lgtm.github.io/work-tracker/

## Features
- Start/stop stopwatch (keeps running if you close or reload the app)
- On stop, asks what you did and saves the date, start/end time and note
- Edit or delete entries
- Pick a month and export `Invoice-YYYY-MM.xlsx`:
  - one row per day worked, with multiple jobs joined on the same row
  - time worked as hours and minutes, plus decimal hours
  - bottom rows: total hours, pay rate, total pay (formulas)
- Works offline after the first load

## Install on Android
1. Open the live app link in Chrome.
2. Tap **⋮ → Add to Home Screen**.

Exported files are saved to your phone's Downloads folder.

## Data
Entries are stored in the browser's `localStorage` on your device. Clearing Chrome's site data erases them, so keep your exported files as backups.

## Configuration
Set your hourly pay in the app ("Pay per hour" field). It defaults to 11.75, is saved on the device, and is used in the summary and exports.

## Development
No build step. Serve the folder over HTTP and open it:

```
python3 -m http.server 8000
```

After changing cached files, bump `CACHE` in `sw.js` so installed copies update.

Spreadsheets are generated with [SheetJS](https://sheetjs.com) (`xlsx.full.min.js`, v0.18.5).

## Files
- `index.html` – UI
- `app.js` – stopwatch, storage, month aggregation, export
- `sw.js`, `manifest.webmanifest`, `icon-*.png` – PWA/offline support
