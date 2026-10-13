# Lalitha Event Decors – Estimate Maker

Quotation, estimate and invoice app for Lalitha Event Decors (since 1985).

## Files
| File | What it is |
|---|---|
| `index.html` | The whole app |
| `manifest.json` | Lets phones install it as an app |
| `icon-180.png`, `icon-192.png`, `icon-512.png`, `icon-maskable-512.png` | App icons |
| `Code.gs` | Google Sheets + Google Calendar script (goes in Apps Script, not GitHub Pages) |
| `brand/` | Logo files: 1024px app icon, red 3D mark, gold 3D mark, white-background JPG |

## 1. Put the app online (GitHub Pages)
1. Create a repository, e.g. `lalitha-estimate`.
2. **Add file → Upload files** → upload everything in this folder (keep the file names) → **Commit changes**.
3. **Settings → Pages** → Branch `main`, folder `/ (root)` → **Save**.
4. After 1–2 minutes the app is at `https://<your-username>.github.io/lalitha-estimate/`.

To update later: upload the new files over the old ones (same names).

## 2. Install on phone / PC
- Android (Chrome): ⋮ → **Add to Home screen** / **Install app**.
- iPhone (Safari): Share → **Add to Home Screen**.
- PC (Chrome/Edge): ⋮ → **Cast, save and share → Install page as app**.
If you change the icon, remove the app and add it again.

## 3. Google Sheets storage (all devices share the same data)
1. Create a Google Sheet → **Extensions → Apps Script** → delete everything → paste `Code.gs`.
2. Change `SECRET` at the top to your own password → Save.
3. Choose function **authorizeCalendar** → **▶ Run** → allow access (once).
4. **Deploy → New deployment → Web app** → Execute as **Me**, Who has access **Anyone** → Deploy. Copy the URL.
5. In the app: **Settings → Sync & backup → Connection details → Change connection** → paste URL and password → **Save & connect**.

When you update `Code.gs`: paste it again (keep your SECRET) → **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. The URL stays the same.

The Sheet gets tabs: Estimates, Items, Payments, Clients, Catalog (editable), Packages, Photos, Settings.

## 4. Google Calendar
Confirmed bookings go to your Google Calendar automatically. To also send them to other Gmail calendars, add the addresses in **Settings → Sync & backup**. To use a separate calendar, put its ID in `CALENDAR_ID` in `Code.gs`. Turn off with `CALENDAR_SYNC = false`.
