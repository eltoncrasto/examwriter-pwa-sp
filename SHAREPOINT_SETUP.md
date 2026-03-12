# SharePoint Sync – IT Admin Setup Guide
## Exam Writer PWA

This guide explains how to enable SharePoint sync in Exam Writer.
No server-side code is needed — everything runs in the browser using Microsoft's MSAL.js library.

---

## What you'll need
- An **Azure AD / Entra ID** account with permission to register applications
- A **SharePoint site** (e.g. `https://contoso.sharepoint.com/sites/Exams`) with a document library
- The deployed URL of Exam Writer (e.g. `https://yourschool.github.io/exam-writer-pwa/`)

---

## Step 1 — Register an Azure AD Application

1. Go to [https://portal.azure.com](https://portal.azure.com) and sign in.
2. Navigate to **Azure Active Directory → App registrations → New registration**.
3. Fill in:
   - **Name**: `Exam Writer PWA` (or any name)
   - **Supported account types**: *Accounts in this organizational directory only* (single tenant)
   - **Redirect URI**:
     - Platform: **Single-page application (SPA)**
     - URI: `https://yourschool.github.io/exam-writer-pwa/`
       *(must match exactly, including trailing slash)*
4. Click **Register**.
5. Copy the **Application (client) ID** — you'll need this.
6. Copy the **Directory (tenant) ID** — you'll need this too.

---

## Step 2 — Add API Permissions

1. In your new app registration, go to **API permissions → Add a permission**.
2. Choose **Microsoft Graph → Delegated permissions**.
3. Add these two permissions:
   - `Files.ReadWrite`
   - `Sites.ReadWrite.All`
4. Click **Add permissions**.
5. Click **Grant admin consent for [your org]** — this avoids candidates needing to consent individually.

---

## Step 3 — Configure the SharePoint Site

1. Create (or choose) a SharePoint site, e.g. `https://contoso.sharepoint.com/sites/Exams`.
2. Make sure the **document library** is accessible to the accounts that will sign in.
3. Decide on a root folder name, e.g. `ExamFiles`.
   Files will be saved as:
   ```
   ExamFiles/
   └── <CandidateID>/
       └── <centerNo>-<candidateId>-<examTitle>-<date>.txt
   ```

---

## Step 4 — Enter the Config in Exam Writer

1. Open Exam Writer in the browser.
2. Open the **Admin panel** (Ctrl+Shift+E, or tap the logo 5 times).
3. Enter the admin PIN (default: `0000`).
4. In the **SharePoint / OneDrive Sync** section, fill in:

   | Field | Value |
   |---|---|
   | Azure Client ID | *(from Step 1)* |
   | Azure Tenant ID | *(from Step 1)* |
   | SharePoint Site URL | `https://contoso.sharepoint.com/sites/Exams` |
   | Root Folder | `ExamFiles` *(or your chosen folder name)* |

5. Click **Save SP Config**.
6. Click **Connect SharePoint account** — a Microsoft login popup will appear.
7. Sign in with an account that has access to the SharePoint site.
8. The status pill in the toolbar should show **Connected: user@school.org**.

---

## Step 5 — Verify

- Type something in the editor, then click **☁ Sync** in the toolbar.
- Check the SharePoint site → Documents → `ExamFiles/` — the candidate's file should appear.
- Autosave will also upload to SharePoint every minute automatically.

---

## Notes for exam sessions

- Each device/candidate needs to sign in to SharePoint once (tokens are cached in `localStorage`).
- If the session is on a **shared/kiosk device**, use the admin panel's **Exit app / end session** button at the end — this clears auth state.
- If a candidate's ID changes mid-session, the next save will create a new subfolder automatically.
- SharePoint sync is **in addition to** local file saving — local save always happens first.
- If SharePoint upload fails (e.g. no internet), the status pill shows red. The local file is always safe.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| "SP: site not found" | Wrong Site URL or missing `Sites.ReadWrite.All` | Check URL; re-grant permissions |
| "SP: auth failed" | Token expired or wrong tenant | Re-sign-in via admin panel |
| "SP: upload failed" | Firewall / CORS / library locked | Check SharePoint library permissions |
| Login popup blocked | Browser popup blocker | Allow popups for the GitHub Pages URL |
| "MSAL library not loaded" | CDN blocked by network | Whitelist `alcdn.msauth.net` |

---

*Generated for Exam Writer PWA — SharePoint sync via Microsoft Graph API.*
