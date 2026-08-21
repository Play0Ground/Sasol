# Project file map

Respirator Fitment attendance app (GitHub Pages).

| File | Purpose |
|------|---------|
| `index.html` | Page structure — Register / Find Me / Moderator tabs |
| `css/styles.css` | Look & layout (phone + laptop) |
| `js/config.js` | Apps Script `/exec` URL only (PIN is NOT here) |
| `js/app.js` | Behaviour: register, find, PIN gate, sheet editing, charts |
| `Code.gs` | Google Apps Script backend — paste into the Sheet’s Apps Script editor |

## Spreadsheet tabs

| Tab | Purpose |
|-----|---------|
| Register / Sheet1 / Attendance | Name, Control No/ID, Company, Date, Size, **Time Signed In** |
| `Config` | Session location / facilitator |
| **`Access`** | **Moderator PIN in cell B1** (label in A1). Created automatically on first API call. |

## Deploy notes

- Frontend: push to `main` → GitHub Pages. Hard-refresh (Ctrl+F5).
- Backend: after changing `Code.gs` → **Deploy → Manage deployments → Edit → New version → Deploy**.
- Change moderator PIN anytime in **Access!B1** — no website redeploy needed.

## Behaviour

- Sign-in stamps **date** (e.g. 21 August 2026) and **time** (HH:mm, Johannesburg) automatically.
- PIN is never shown on the website; it must match Access!B1.
- Size overview bars are simple counts (not ML).
