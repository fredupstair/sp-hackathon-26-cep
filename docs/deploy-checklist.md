# CEP – Deploy Checklist

## ✅ Already done
- App Registration `CEP-Backend` created in Entra ID
- `user_impersonation` scope exposed (`api://<app-registration-client-id>`)
- Graph application permissions granted: `AiEnterpriseInteraction.Read.All`, `TeamsActivity.Send`, `Sites.Selected`
- Azure resources created in `<resource-group>` (westeurope): Storage Account, Queue `cep-ingest`, Application Insights, **Function App**
- Backend code deployed to `https://<function-app-name>.azurewebsites.net`
- Function App settings configured (all except `TeamsAppId`)

--- 

## ⏳ Step 1 – Deploy SPFx

```powershell
cd fe
npm run build       # heft test --clean --production && heft package-solution --production
```

1. Open **SharePoint Admin Center** → **More features** → **Apps** → **App Catalog**
2. Upload `fe/sharepoint/solution/cep.sppkg`
3. Check **"Make this solution available to all sites"** → Deploy

---

## ⏳ Step 2 – Approve API Access in SharePoint

1. **SharePoint Admin Center** → **Advanced** → **API access**
2. Approve the pending request:  
   `CEP-Backend` · scope `user_impersonation`
3. Once approved, the web part can obtain an Entra ID token to call the Function App

**Property pane value to set on the web part:**
```
Function App Base URL: https://<function-app-name>.azurewebsites.net
```

---

## ⏳ Step 3 – Create the Teams App for notifications

Teams notifications (badge earned, level-up, inactivity nudge) require a Teams App with the `activityTypes` declared in its manifest.

### 3a. Create the Teams App in Developer Portal

1. Open **https://dev.teams.microsoft.com** → **Apps** → **New app**
2. Fill in:
   - **Name:** `Copilot Engagement Program`
   - **App ID:** generate a new GUID (e.g. `[guid]::NewGuid()` in PowerShell) — **record this GUID, it is the `TeamsAppId`**
   - **Version:** `1.0.0`
   - **Short description:** `Gamify your Copilot usage`
   - **Developer / Company name:** your organization name
3. Under **App features** → **Activity feed notification**
4. Add the following `activityTypes`:

| Type | Description | templateText |
|------|-------------|--------------|
| `cepWelcome` | Welcome to CEP | `{actor} enrolled you in the Copilot Engagement Program!` |
| `cepLevelUp` | Level change | `You reached the {level} level! 🎉` |
| `cepBadgeEarned` | New badge | `You earned the "{badge}" badge!` |
| `cepInactivityNudge` | Inactivity reminder | `It has been {days} days since you last used Copilot – come back!` |
| `cepLeaderboardUpdate` | Leaderboard refresh | `The monthly leaderboard has been updated. You are #{rank}!` |

> **Quick MVP alternative:** use `systemDefault` instead of custom activity types — no dedicated manifest required, but notification text is less customizable.

### 3b. Add RSC permission (optional for MVP)
To send notifications without pre-installing the app per user, use the `TeamsActivity.Send` application permission (already granted) with the app installed at tenant level.

### 3c. Publish the Teams App
1. Developer Portal → **Publish** → **Publish to your org**
2. **Teams Admin Center** → **Manage apps** → search for `Copilot Engagement Program` → **Publish** (approval)
3. Optional but recommended: **Setup policies** → pre-install the app for all users (or a pilot group)

### 3d. Set TeamsAppId in the Function App

```powershell
# Replace <teams-app-id> with the App ID from step 3a
az functionapp config appsettings set `
    --name <function-app-name> `
    --resource-group <resource-group> `
    --settings TeamsAppId=<teams-app-id>
```

Also update `be/local.settings.json` for local development:
```json
"TeamsAppId": "<teams-app-id>"
```

---

## Quick reference

| Resource | Value |
|---------|--------|
| Function App URL | `https://<function-app-name>.azurewebsites.net` |
| App Registration ClientId | `<app-registration-client-id>` |
| App ID URI (SPFx resource) | `api://<app-registration-client-id>` |
| SharePoint Site | `https://<tenant>.sharepoint.com/sites/<site-name>` |
| Resource Group | `<resource-group>` (westeurope) |
| Tenant ID | `<tenant-id>` |
