# QA Form Bug App - Netlify Deployment

This package converts the app from a long-running Render/Express server to a Netlify static site plus Netlify Functions.

## Architecture

Netlify static UI -> Netlify Function `/api/*` -> Google Apps Script -> Google Sheet + Google Drive.

## Important limits

- The browser upload limit is set to 5 MB per local file for stable serverless upload.
- For large videos, upload the video manually to Google Drive and paste/drag the Drive link into the upload box.
- You must update Google Apps Script with the included `google-apps-script/Code.gs` file.

## Files added/changed

- `netlify.toml`
- `netlify/functions/api.js`
- `public/app.js`
- `google-apps-script/Code.gs`
- `package.json`
- `.env.example`

## Google Apps Script update

1. Open your Google Sheet.
2. Go to `Extensions -> Apps Script`.
3. Replace the full script with `google-apps-script/Code.gs` from this package.
4. Click Save.
5. Go to `Deploy -> Manage deployments`.
6. Edit the web app deployment.
7. Choose `New version`.
8. Deploy.
9. Keep/copy the Web App URL.

## Netlify setup

In Netlify, create a new site from GitHub.

Use:

- Build command: `npm run build`
- Publish directory: `public`

The functions directory is already configured in `netlify.toml`.

Add these environment variables in Netlify:

```env
GOOGLE_APPS_SCRIPT_WEB_APP_URL=https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec
GOOGLE_SHEET_VIEW_URL=https://docs.google.com/spreadsheets/d/1JlJtBq3GlsEG1Rc9cwTLxcXDAwpvoX2C2Ld8fqEr6u0/edit?usp=sharing
GOOGLE_DRIVE_ROOT_FOLDER_ID=1VA4Awn12PKmMc1VIEdimK-qwYBieU0yC
```

## After deploy

Open your Netlify URL and test:

1. Save a form preset.
2. Submit one issue with a small image.
3. Check the Google Sheet.
4. Check the Google Drive folder.
5. Download Excel from the Issue Report tab.
