# Project file map

Respirator Fitment attendance app (GitHub Pages).

| File | Purpose |
|------|---------|
| `index.html` | Page structure only — Register / Find Me / Moderator tabs |
| `css/styles.css` | Look & layout (phone + laptop) |
| `js/config.js` | Apps Script `/exec` URL + moderator PIN (`2026`) |
| `js/app.js` | Behaviour: register, find, PIN gate, sheet editing |
| `Code.gs` | Google Apps Script backend — paste into the Sheet’s Apps Script editor |

## Bug fixed (Aug 2026)

PIN keyboard capture was listening globally, so typing a **control number** on Register was stolen as PIN digits (`2026`). PIN keys now only work while the **Moderator** tab is open and the PIN gate is showing.

## Deploy notes

- Frontend: push to `main` → GitHub Pages updates automatically.
- Backend: after changing `Code.gs`, Deploy → New version in Apps Script.
