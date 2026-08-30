# CareerOS

CareerOS is a personal job-search PWA designed to run on GitHub Pages with a Google Apps Script backend.

## Current V0.1

- Dashboard
- Jobs tracker
- Applications pipeline
- 36-company target list
- Contacts / networking drafts
- Master resume storage
- Local-first browser state
- Apps Script backend URL setting
- Mobile-responsive PWA shell

## GitHub Pages

1. Create a new repository, e.g. `CareerOS`.
2. Upload:
   - `index.html`
   - `manifest.json`
   - `Code.gs`
   - `README.md`
3. Enable GitHub Pages from the repository settings.
4. Open the published Pages URL.
5. Deploy `Code.gs` as an Apps Script Web App.
6. In CareerOS → Settings, paste the Apps Script `/exec` URL.

## Important

This V0.1 intentionally does NOT:
- store passwords
- automatically submit applications
- scrape LinkedIn/Naukri
- make unsupported resume claims
- send LinkedIn/email messages automatically

Those are later integrations/workflows.

## Data

The frontend currently keeps data in browser localStorage. The Apps Script backend provides the next integration layer for Google Sheets. Do not put passwords or sensitive credentials into the sheet.
