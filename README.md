# Project file map

Respirator Fitment attendance app (GitHub Pages).

| File | Purpose |
|------|---------|
| `index.html` | Page structure — Register / Find Me / Moderator tabs |
| `css/styles.css` | Look & layout (phone + laptop) |
| `js/config.js` | Apps Script `/exec` URL + moderator PIN (`2026`) |
| `js/app.js` | Behaviour: register, find, PIN gate, sheet editing, day filter, duplicate flags |
| `Code.gs` | Google Apps Script backend — paste into the Sheet’s Apps Script editor |

## Deploy notes

- Frontend: push to `main` → GitHub Pages updates automatically. Hard-refresh (Ctrl+F5).
- Backend: after changing `Code.gs`, in Apps Script: **Deploy → Manage deployments → Edit → New version → Deploy**.

## Behaviour notes

- Control number / ID is stored as **text** (leading zeros kept).
- New sign-ins cannot reuse an existing control number / ID.
- Existing duplicates in the sheet are **flagged** (yellow), not auto-deleted.
- Moderator view auto-refreshes from the spreadsheet (~20s) and on tab focus.
- After delete/save, the UI reloads from the sheet — no manual page refresh needed.
