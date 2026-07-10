# Skylark Drones — Cards to Excel

Scan business cards on your phone → contact details land in **one shared Google Sheet** for the whole team, with the Company Sector filled in by AI web search.

- **Frontend:** React (Vite), hosted free on GitHub Pages, installable on phones ("Add to Home Screen")
- **Backend:** Google Apps Script bound to a Google Sheet — the Sheet *is* the shared database
- **AI:** Anthropic Claude reads each card and web-searches the company's sector; the API key stays server-side

Columns written to the sheet: Serial Number · Date Added · Name of Person · Name of Company · Company Sector · Phone Number · Email ID · Lead Owner.

---

## Setup (one time, ~15 minutes)

### 1. Create the shared Google Sheet
1. Go to [sheets.new](https://sheets.new) using the Skylark Google account that should own the data.
2. Name it e.g. **Skylark Cards to Excel**. (The "Contacts" and "Users" tabs are created automatically on first use.)

### 2. Add the backend script
1. In the Sheet: **Extensions → Apps Script**.
2. Delete any placeholder code and paste the full contents of [`apps-script/Code.gs`](apps-script/Code.gs).
3. **Project Settings (⚙) → Script properties → Add script property:**
   - Property: `ANTHROPIC_API_KEY`
   - Value: your key from [console.anthropic.com](https://console.anthropic.com) (starts with `sk-ant-`)
4. **Deploy → New deployment → Web app:**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Authorize when prompted, then **copy the Web app URL** (ends in `/exec`).

### 3. Point the frontend at your backend
Open `src/config.js` and paste the Web app URL:

```js
export const API_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
```

### 4. Push to GitHub and enable Pages
1. Create a new GitHub repository and push this folder to the `main` branch.
2. In the repo: **Settings → Pages → Source: GitHub Actions**.
3. The included workflow (`.github/workflows/deploy.yml`) builds and deploys automatically on every push. Your app URL will be `https://<username>.github.io/<repo-name>/`.

### 5. Install on phones
Open the Pages URL on each phone → browser menu → **Add to Home Screen**. It launches full-screen with the Skylark icon, camera and gallery included. Everyone signs up once and stays logged in on their device.

---

## Local development

```bash
npm install
npm run dev
```

## Updating the backend
After editing `Code.gs` in Apps Script, use **Deploy → Manage deployments → Edit (✏) → Version: New version** so the same URL serves the new code.

## Notes & limits
- **Costs:** GitHub Pages and Apps Script are free. Anthropic API usage is pay-as-you-go (a fraction of a rupee per card).
- **Security:** Passwords are stored as SHA-256 hashes in the hidden "Users" tab. This is lightweight team auth, not enterprise SSO — share the app URL only within Skylark. Keep the Google Sheet's own sharing restricted to the team.
- **Concurrency:** Serial numbers are protected by a script lock, so simultaneous uploads from different phones won't collide.
- **Sheet access:** The "Open the live Google Sheet" link in the app points at the real Sheet. Share the Sheet (view or edit) with the team from Google Sheets as you see fit.
