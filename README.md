# IC Project Report — Cloudflare Worker Setup Guide

## Overview
This system accepts IC Project Report form submissions, then simultaneously:
1. Creates a new item in a SharePoint List
2. Uploads photos to a SharePoint Document Library (in a dated subfolder)
3. Sends a formatted HTML confirmation email via Microsoft Graph

---

## Prerequisites
- A Cloudflare account with Workers enabled
- A Microsoft 365 / Azure AD tenant
- A SharePoint site with a configured List and Document Library

---

## Step 1 — Azure AD App Registrations

Two app registrations are needed: one for the **backend Worker** (server-to-server, client credentials) and one for the **admin SPA** (interactive user login). You can use a single registration for both if you prefer — see the note at the end of this section.

---

### 1a — Backend App Registration (Worker → SharePoint/Graph)

1. Go to **portal.azure.com → Azure Active Directory → App registrations → New registration**
2. Name: `IC Report Worker` · Account type: **Single tenant** · No redirect URI
3. Click **Register**

**API Permissions** (Application — not Delegated):

| API | Permission | Purpose |
|-----|-----------|---------|
| Microsoft Graph | `Sites.ReadWrite.All` | SharePoint list + file upload |
| Microsoft Graph | `Mail.Send` | Confirmation email |

Click **Grant admin consent** after adding both.

**Client Secret:** Go to **Certificates & secrets → New client secret** → copy the **Value** immediately.

**Copy from App Overview:**
- **Application (client) ID** → `AZURE_CLIENT_ID`
- **Directory (tenant) ID** → `AZURE_TENANT_ID`

---

### 1b — Admin SPA App Registration (Admin Viewer → Worker)

1. **New registration** (or continue using the same app from 1a — see note below)
2. Name: `IC Report Admin` · Account type: **Single tenant**
3. **Redirect URI:** Select platform **Single-page application (SPA)** and enter the URL where `admin.html` is hosted, e.g. `https://yourwebsite.com/admin.html`
4. Click **Register**

**Expose an API scope** (so the SPA can request a token the Worker will validate):

1. Go to **Expose an API → Add a scope**
2. If prompted for an Application ID URI, accept the default: `api://{ADMIN_CLIENT_ID}`
3. Scope name: `Reports.Read`
4. Who can consent: **Admins and users**
5. Save the scope — the full scope URI will be `api://{ADMIN_CLIENT_ID}/Reports.Read`

**Grant the SPA permission to the scope:**

1. Go to **API permissions → Add a permission → My APIs**
2. Select `IC Report Admin` (or `IC Report Worker` if using one app)
3. Add `Reports.Read`
4. Click **Grant admin consent**

**Copy from App Overview:**
- **Application (client) ID** → `ADMIN_CLIENT_ID`

> **Using one app registration:** If you want to keep a single registration, add the SPA redirect URI and `Expose an API` scope to the same app used in Step 1a. Set `ADMIN_CLIENT_ID` to the same value as `AZURE_CLIENT_ID` in both the Worker secrets and `admin.html`.

---

## Step 2 — SharePoint List Columns

Create a new SharePoint List (e.g. **"IC Project Reports"**) with these columns.
Column **internal names** must match exactly — set them on first creation.

| Column Name (Internal) | Type | Notes |
|------------------------|------|-------|
| `Title` | Single line | Built-in · used for project name |
| `Location` | Single line | City, Country |
| `ProjectDateFrom` | Date/Time | |
| `ProjectDateTo` | Date/Time | |
| `Introduction` | Multiple lines (Plain) | |
| `ChurchesParticipated` | Number | |
| `Localities` | Number | |
| `NationalParticipants` | Number | |
| `USAParticipants` | Number | |
| `OtherCountriesParticipants` | Number | |
| `TotalVisits` | Number | |
| `PeopleHeardGospel` | Number | |
| `ProfessionsOfFaith` | Number | |
| `Rededications` | Number | |
| `Baptisms` | Number | |
| `NewChurchesPlanted` | Number | |
| `Testimonies` | Multiple lines (Plain) | Stored as JSON array string |
| `TotalFundsSent` | Currency | |
| `SpentOnMaterials` | Currency | |
| `TicketsCost` | Currency | |
| `FuelCost` | Currency | |
| `AccommodationCost` | Currency | |
| `FoodCost` | Currency | |
| `FinancialHelpParticipants` | Currency | |
| `NumParticipantsHelp` | Number | |
| `RalliesExpenses` | Currency | |
| `RalliesDescription` | Single line | |
| `AdditionalExpenses` | Currency | |
| `AdditionalNeedDescription` | Single line | |
| `CoordinatorName` | Single line | |
| `CoordinatorEmail` | Single line | |
| `SubmittedAt` | Date/Time | |

---

## Step 3 — Deploy the Worker

### Install Wrangler CLI
```bash
npm install -g wrangler
wrangler login
```

### Create project
```bash
mkdir ic-report-worker && cd ic-report-worker
wrangler init --type=module
```

Copy `worker.js` into the project root. In `wrangler.toml`:
```toml
name = "ic-report-worker"
main = "worker.js"
compatibility_date = "2024-01-01"
```

### Set secrets
```bash
wrangler secret put AZURE_TENANT_ID        # Directory (tenant) ID
wrangler secret put AZURE_CLIENT_ID        # Backend app registration client ID
wrangler secret put AZURE_CLIENT_SECRET    # Backend app client secret value
wrangler secret put ADMIN_CLIENT_ID        # Admin SPA app registration client ID
                                           # (same as AZURE_CLIENT_ID if using one app)
wrangler secret put SHAREPOINT_SITE_URL    # https://yourorg.sharepoint.com/sites/yoursite
wrangler secret put SHAREPOINT_LIST_NAME   # IC Project Reports
wrangler secret put SHAREPOINT_FOLDER_PATH # /sites/yoursite/Shared Documents/IC Report Photos
wrangler secret put EMAIL_SENDER           # ic-reports@yourorg.com
wrangler secret put EMAIL_RECIPIENT        # admin@yourorg.com
wrangler secret put ALLOWED_ORIGIN         # https://yourwebsite.com
```

> For `SHAREPOINT_FOLDER_PATH`: use the server-relative path. Navigate to the folder in SharePoint and check the URL.

### Deploy
```bash
wrangler deploy
```

Wrangler outputs your Worker URL, e.g.: `https://ic-report-worker.your-subdomain.workers.dev`

---

## Step 4 — Configure the Form

In `form.html`, update the `WORKER_URL` constant:
```javascript
const WORKER_URL = "https://ic-report-worker.your-subdomain.workers.dev";
```

Embed or host `form.html` on your website.

---

## Step 5 — Configure the Admin Viewer

In `admin.html`, update the three constants at the top of the `<script>` block:
```javascript
const WORKER_URL      = "https://ic-report-worker.your-subdomain.workers.dev";
const TENANT_ID       = "your-directory-tenant-id";
const ADMIN_CLIENT_ID = "your-admin-spa-client-id";
```

Host `admin.html` at the same URL you registered as the redirect URI in Step 1b.

---

## Step 6 — Test

**Form submission:**
- [ ] New item appears in the SharePoint list with all fields populated
- [ ] Photos appear in the Document Library under a dated subfolder
- [ ] Confirmation email received at `EMAIL_RECIPIENT`
- [ ] Coordinator receives CC email if they provided their address
- [ ] Form shows a green success message

**Admin viewer:**
- [ ] `admin.html` loads and shows the Microsoft sign-in button
- [ ] Clicking sign-in opens the Microsoft login popup (or redirects)
- [ ] After sign-in, reports load in the left sidebar
- [ ] Clicking a report shows the full detail view with all testimonies flattened
- [ ] Sign out button clears the session

---

## Troubleshooting

| Symptom | Likely Cause |
|---------|-------------|
| `Token fetch failed 401` | Wrong `AZURE_CLIENT_ID` or `CLIENT_SECRET` |
| `List "..." not found` | `SHAREPOINT_LIST_NAME` doesn't match exactly |
| `Graph 403` on list write | `Sites.ReadWrite.All` not granted or missing admin consent |
| `sendMail failed 403` | `Mail.Send` not granted, or `EMAIL_SENDER` isn't a licensed M365 user |
| CORS error in browser | `ALLOWED_ORIGIN` must exactly match your site's origin (no trailing slash) |
| Admin: `Invalid audience` | `ADMIN_CLIENT_ID` in Worker secrets doesn't match the SPA app registration |
| Admin: `Invalid issuer` | `AZURE_TENANT_ID` in Worker secrets is wrong |
| Admin: popup blocked | Browser blocked the MSAL popup — the viewer falls back to redirect automatically |
| Admin: `AADSTS50011` | Redirect URI in admin.html doesn't match what's registered in Step 1b |
| Admin: consent required | User needs to consent to `Reports.Read` scope — grant admin consent in App Registration |
| Numbers saving as null | Column internal name mismatch — check SharePoint column settings |

---

## Security Notes
- All Azure credentials are stored as Cloudflare Worker **encrypted secrets** — never in code
- Admin endpoints validate Azure AD JWTs cryptographically (RSA-SHA256 + JWKS) — no shared secrets
- JWKS keys are cached for 1 hour per Worker isolate to avoid repeated Azure AD requests
- The `ALLOWED_ORIGIN` check blocks other sites from using your Worker's POST endpoint
- Consider adding a Cloudflare **Rate Limiting rule** on the Worker route to prevent form spam
- For production, enable **Bot Fight Mode** in Cloudflare on the Worker route
