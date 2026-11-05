# EASY MATHS — Vercel Frontend

Frontend (index + admin) hosted on Vercel.
Backend = Google Apps Script Web App + Google Sheet.

## Setup Steps
1. Deploy `apps-script/Code.gs` as Web App
   (Execute as: Me, Who has access: Anyone).
2. Copy the `/exec` URL from Apps Script.
3. Open `public/config.js` and paste that URL into `EASY_MATHS_API_URL`.
4. Push this folder to GitHub.
5. Import the GitHub repo to Vercel → Deploy.