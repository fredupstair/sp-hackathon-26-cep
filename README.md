# Copilot Engagement Program

**SharePoint Hackathon 2026 — submission [#145](https://github.com/SharePoint/sharepoint-hackathon/issues/145)**

> Driving daily Copilot usage through SharePoint extensibility, gamification, and agentic experiences.

---

## The Problem

Deploying Copilot licenses is easy. Driving sustained adoption is not. Users have no visibility into their own usage, no feedback loop to reward progress, and no reason to build a daily Copilot habit. Enablement teams cannot identify lagging users at scale without manual surveys.

---

## What it Does

**Copilot Engagement Program (CEP)** gamifies Microsoft 365 Copilot usage across an organization. Every prompt an employee sends in Word, Excel, PowerPoint, Outlook, Teams, OneNote, or Loop earns points, unlocks badges, and climbs a leaderboard — all without leaving SharePoint or Teams.

Users opt in voluntarily, granting explicit consent before any tracking begins. A persistent header strip on every SharePoint page shows live stats. A declarative Copilot agent inside Teams lets users query their rankings and stats in natural language.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         SharePoint Online (browser)                         │
│                                                                             │
│  ┌─────────────────────┐  ┌─────────────────────┐  ┌────────────────────┐  │
│  │   cepWelcome        │  │   cepLeaderboard    │  │     cepWins        │  │
│  │   Web Part          │  │   Web Part          │  │     Web Part       │  │
│  │                     │  │                     │  │                    │  │
│  │  • Enrollment wizard│  │  • Global ranking   │  │  • Submit wins     │  │
│  │  • Personal dashboard  │  • Dept / Team      │  │  • Screenshot upload  │
│  │  • Points & badges  │  │  • Podium + paging  │  │  • Bonus points    │  │
│  └──────────┬──────────┘  └──────────┬──────────┘  └─────────┬──────────┘  │
│             │                        │                        │             │
│  ┌──────────┴────────────────────────┴────────────────────────┴──────────┐  │
│  │                   cepBar  (SPFx Application Customizer)               │  │
│  │          Persistent header strip — live points · level · prompts      │  │
│  └───────────────────────────────────────────────────────────────────────┘  │
│                                      │  Entra ID token (EasyAuth)           │
└──────────────────────────────────────┼──────────────────────────────────────┘
                                       │
┌──────────────────────────────────────┼──────────────────────────────────────┐
│              Microsoft Teams         │                                      │
│  ┌───────────────────────────────┐   │                                      │
│  │         CeP Agent             │   │  Activity Feed Notifications         │
│  │  Declarative Copilot Agent    │   │  • Badge earned                      │
│  │  (TypeSpec + M365 Agents TK)  │   │  • Level up                          │
│  │  Natural language → API calls │   │  • Inactivity nudge                  │
│  └──────────────┬────────────────┘   │  • Leaderboard refresh               │
└─────────────────┼────────────────────┼──────────────────────────────────────┘
                  │                    │
┌─────────────────┼────────────────────┼──────────────────────────────────────┐
│                 │   Azure Functions (.NET 8 – Isolated Worker)              │
│                 │                    │                                      │
│  ┌──────────────┴──┐  ┌──────────────┴──┐  ┌─────────────┐  ┌───────────┐  │
│  │  EnrollmentApi  │  │     MeApi       │  │LeaderboardApi  │  AdminApi  │  │
│  │  POST join/leave│  │  GET summary    │  │  GET rankings│  │GET config │  │
│  └─────────────────┘  │  GET usage      │  └─────────────┘  │POST config│  │
│                       │  GET badges     │                   │GET status │  │
│                       │  POST wins      │                   └───────────┘  │
│                       │  GET suggest(AI)│                                   │
│                       └─────────────────┘                                   │
│                                                                             │
│  ╔═══════════════════════════════════════════════════════════════════════╗  │
│  ║  Background Processing                                                ║  │
│  ║                                                                       ║  │
│  ║  OrchestratorTimer (every minute)                                     ║  │
│  ║    └── reads CEP_Config + SyncState                                   ║  │
│  ║    └── if sync due: enqueues 1 msg/user ──────────────────────────┐  ║  │
│  ║                                                                    ▼  ║  │
│  ║  LeaderboardTimer (daily 08:00 UTC)         Azure Storage Queue       ║  │
│  ║    └── rebuilds CEP_Leaderboard                      │               ║  │
│  ║    └── awards MonthlyMaster badges                   ▼               ║  │
│  ║                                           IngestQueueWorker           ║  │
│  ║                                           (1 msg = 1 user)            ║  │
│  ║                                           1. fetch Graph AI history   ║  │
│  ║                                           2. count prompts per app    ║  │
│  ║                                           3. upsert Azure Table       ║  │
│  ║                                           4. update points + level    ║  │
│  ║                                           5. award badges             ║  │
│  ║                                           6. send Teams notification  ║  │
│  ╚═══════════════════════════════════════════════════════════════════════╝  │
└──────────────────────┬──────────────────────────────┬───────────────────────┘
                       │  Managed Identity              │
          ┌────────────┘                               └──────────────────┐
          ▼                                                               ▼
┌──────────────────────────────┐              ┌───────────────────────────────┐
│       Microsoft Graph        │              │         Data Stores           │
│                              │              │                               │
│  AI Interaction Export:      │              │  SharePoint Lists             │
│  GET /copilot/users/{id}/    │              │  ├── CEP_Users                │
│   interactionHistory/        │              │  ├── CEP_Leaderboard          │
│   getAllEnterpriseInteractions│              │  ├── CEP_BadgeCatalog         │
│                              │              │  ├── CEP_Config               │
│  Teams Activity Feed:        │              │  └── CEP_SyncState            │
│  POST /teamwork/             │              │                               │
│   sendActivityNotification   │              │  Azure Table Storage          │
│   ToRecipients               │              │  ├── cepActivityLog           │
│                              │              │  ├── cepEarnedBadges          │
│  Azure OpenAI (AI welcome)   │              │  └── cepCopilotWins           │
└──────────────────────────────┘              └───────────────────────────────┘
```

---

## Repository Structure

```
be/                     Azure Functions backend (.NET 8)
  Functions/            HTTP and timer/queue-triggered functions
  Services/             Shared services (Graph, SharePoint, engines, notifier)
  Models/               Data model types
  local.settings.json.example

fe/                     SPFx solution (web parts + app customizer)
  src/
    webparts/
      cepWelcome/       Enrollment wizard + personal dashboard
      cepLeaderboard/   Full paginated leaderboard
      cepWins/          Community wall for productivity wins
    extensions/         cepBar application customizer (header strip)
    services/           Shared API clients and helpers

fe-appcustomizer/       Standalone app customizer package (tenant-wide deployment)

CeP Agent/              Declarative Copilot Agent
  src/agent/            Agent definition
  appPackage/           Teams app manifest and adaptive cards
  tspconfig.yaml        TypeSpec configuration

sp/                     SharePoint provisioning
  templates/            PnP provisioning templates for CEP_* lists

deploy/                 Deployment scripts (PowerShell)
docs/                   PRD, tech spec, deploy checklist
```

---

## Components

| Component | Type | Description |
|---|---|---|
| `cepWelcome` | SPFx Web Part | Multi-step enrollment wizard (Welcome → Consent → Rules → Preferences) and personal dashboard showing points, level, per-app usage breakdown, badges, and a mini-leaderboard |
| `cepLeaderboard` | SPFx Web Part | Full paginated leaderboard — Global / Department / Team scopes, podium view, month selector |
| `cepWins` | SPFx Web Part | Community wall where users submit productivity stories with optional screenshots and earn bonus points |
| `cepBar` | SPFx Application Customizer | Tenant-wide persistent header strip showing live points, current level, prompt count, and daily Copilot tip |
| `CeP Agent` | M365 Declarative Copilot Agent | TypeSpec-defined agent deployed via Microsoft 365 Agents Toolkit; lets users ask "What's my rank?" or "Show me this month's leaderboard" in natural language directly inside Teams Copilot |

---

## Backend — Azure Functions (.NET 8)

The backend is a single **.NET 8 Isolated Worker** Function App registered in `be/CepFunctions.csproj`.

### Functions

| Function | Trigger | Responsibility |
|---|---|---|
| `OrchestratorTimer` | Timer (every minute) | Reads `CEP_Config` + `CEP_SyncState`; if a sync cycle is due, enqueues one message per enrolled user into the ingest queue; updates `SyncState` |
| `LeaderboardTimer` | Timer (daily 08:00 UTC) | Rebuilds materialized `CEP_Leaderboard` for all scopes; awards `MonthlyMaster` badges; records `LastLeaderboardRebuildUtc` |
| `IngestQueueWorker` | Storage Queue (1 msg = 1 user) | Fetches Graph AI interaction history with watermark pagination; aggregates prompt counts per app per day; upserts `cepActivityLog`; updates user points and level; evaluates and awards badges; fires Teams notifications |
| `EnrollmentApi` | HTTP POST | `join` — creates the user record with consent timestamp; `leave` — soft-deletes and excludes from all processing |
| `MeApi` | HTTP GET / POST | `summary`, `usage`, `badges`, `wins`, `suggestion` (AI-generated prompt tip) |
| `LeaderboardApi` | HTTP GET | Paginated rankings by scope and month |
| `AdminApi` | HTTP GET / POST | Config read/write, pipeline status, manual sync trigger (`POST /ops/sync`) |

### Shared Services

| Service | Description |
|---|---|
| `GraphClient` | Calls `aiInteractionHistory/getAllEnterpriseInteractions` using application identity (Managed Identity preferred); handles `@odata.nextLink` pagination and per-user watermarks |
| `SharePointClient` | CRUD against SharePoint lists via Microsoft Graph — used for `CEP_Users`, `CEP_Leaderboard`, `CEP_Config`, `CEP_SyncState`, `CEP_BadgeCatalog` |
| `PointsEngine` | Calculates `promptCount × PointsPerPrompt` (configurable) and maps total points to a level (Bronze / Silver / Gold) |
| `BadgeEngine` | Evaluates daily aggregates against badge criteria; upserts `cepEarnedBadges` idempotently |
| `TeamsNotifier` | Sends Activity Feed notifications via `POST /teamwork/sendActivityNotificationToRecipients` (batch ≤ 100) |

### Authentication

- **SPFx → Azure Functions**: Entra ID Bearer token validated by EasyAuth on the Function App. In local development, falls back to `X-User-Id` / `X-Admin-Key` headers.
- **Azure Functions → Microsoft Graph**: Application identity — **Managed Identity** (zero stored secrets) for production; App Registration with client secret for local dev.

### Why the queue-based fan-out pattern?

The Graph AI interaction export endpoint (`AiEnterpriseInteraction.Read.All`) only supports **application permissions** — the SPFx frontend cannot call it directly. All ingestion happens in the backend.

The `OrchestratorTimer → IngestQueue → IngestQueueWorker` fan-out lets the solution scale to 10,000+ enrolled users without hitting Function timeout limits: each queue message processes exactly one user in isolation.

---

## Data Architecture

### SharePoint Lists — configuration and ranked data

| List | Purpose |
|---|---|
| `CEP_Users` | Enrolled users with points, level, watermark, notification preferences, and `IsActive` flag |
| `CEP_Leaderboard` | Materialized monthly rankings pre-computed for Global / Department / Team scopes |
| `CEP_BadgeCatalog` | Admin-managed badge definitions — adding new badges requires zero code changes |
| `CEP_Config` | Runtime parameters: sync frequency, points-per-prompt multiplier, level thresholds, AI fallback flag |
| `CEP_SyncState` | Pipeline health: last run timestamps, processed user counts, error counts |

### Azure Table Storage — high-volume, ~€0.04 / GB / month

| Table | Purpose |
|---|---|
| `cepActivityLog` | Daily prompt aggregates per user per app (13-month retention); partition key = `AadUserId`, row key = `YYYYMMDD_AppKey` |
| `cepEarnedBadges` | One row per user–badge pair; idempotent upsert prevents duplicates |
| `cepCopilotWins` | User-submitted productivity stories with screenshot blob URLs and bonus points |

SharePoint is used for the tables that benefit from its built-in views and admin UX (config, leaderboard, catalog). Azure Table Storage handles the high-write-frequency tables that would otherwise hit SharePoint throttling limits.

---

## Gamification Model

### Levels

| Level | Monthly Points |
|---|---|
| Bronze | 0 – 499 |
| Silver | 500 – 1,499 |
| Gold | 1,500+ |

### Badges

| Badge | Criteria |
|---|---|
| First Steps | First prompt after enrollment |
| Cross-App Explorer | 3+ different Copilot apps used in one week |
| Weekly Warrior | 50+ prompts in a single week |
| Monthly Master | Top 10 on the monthly global leaderboard |
| Consistency King | 7 consecutive active days |

Badge definitions live in `CEP_BadgeCatalog` — admins can add new badges without touching the codebase.

---

## Privacy by Design

- **Opt-in only** — explicit informed consent (with a readable rules step) before any data is collected.
- **Data minimization** — only prompt count and app class are stored; prompt content is never captured or stored.
- **No direct list access** — all UI calls go through Entra ID–protected Azure Functions; enrolled users cannot read each other's raw data.
- **Managed Identity** — no stored secrets for Microsoft Graph access in production.
- **One-click leave** — soft delete via `POST /enrollment/leave` immediately excludes the user from all processing and leaderboards.
- **Transparent model** — the leaderboard shows name, department, points and level only; per-app breakdowns are visible only to the user themselves.

---

## Local Development

### Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download)
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [Azurite](https://learn.microsoft.com/azure/storage/common/storage-use-azurite) (local Storage emulator)
- Node.js 18+ and [SPFx toolchain](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment) (`gulp`, `@microsoft/generator-sharepoint`)

### Backend

```powershell
cd be
copy local.settings.json.example local.settings.json
# Fill in TenantId, ClientId, ClientSecret, SpSiteId, ListId_* values
dotnet build
func start   # or use the VS Code task "func: 4"
```

The Functions host starts at `http://localhost:7071`. For local auth bypass, pass `X-User-Id` and `X-Admin-Key` headers (values configured in `local.settings.json`).

### Frontend

```powershell
cd fe
npm install
gulp serve   # launches the SPFx workbench with live reload
```

Point the web part property pane to `http://localhost:7071` as the Function App base URL.

### Required `local.settings.json` values

| Key | Description |
|---|---|
| `TenantId` | Entra ID tenant GUID |
| `ClientId` | App Registration client ID (needs `AiEnterpriseInteraction.Read.All`, `TeamsActivity.Send`, `Sites.Selected`) |
| `ClientSecret` | App Registration secret (local only — use Managed Identity in Azure) |
| `SpSiteId` | GUID of the SharePoint site hosting the CEP lists |
| `SpSiteUrl` | Full URL of the SharePoint site |
| `ListId_*` | GUIDs of each SharePoint list (run `deploy/get-sp-ids.ps1` to retrieve them) |
| `IngestQueueName` | Azure Storage Queue name (default: `cep-ingest`) |
| `TeamsAppId` | GUID of the Teams App used for Activity Feed notifications |

---

## Deployment

A full step-by-step checklist is in [docs/deploy-checklist.md](docs/deploy-checklist.md). The high-level sequence is:

1. **Provision SharePoint lists** — run the PnP template in `sp/` against your target site.
2. **Create App Registration** — run `deploy/create-app-reg.ps1`; grant `AiEnterpriseInteraction.Read.All`, `TeamsActivity.Send`, `Sites.Selected`.
3. **Deploy Azure resources** — run `deploy/deploy-azure.ps1` (creates Storage Account, Queue, Application Insights, Function App in `rg-cep-hackathon`).
4. **Deploy backend** — `dotnet publish --configuration Release` → deploy the output to the Function App, or use `deploy/deploy-now.ps1`.
5. **Deploy SPFx** — `gulp bundle --ship && gulp package-solution --ship` → upload `cep.sppkg` to the App Catalog → approve the API access request in SharePoint Admin Center.
6. **Deploy CeP Agent** — use Microsoft 365 Agents Toolkit (`m365agents.yml`) to package and publish the declarative agent.
7. **Configure Teams App** — create the Teams App in Developer Portal with the required `activityTypes`, publish to org, and set `TeamsAppId` in Function App settings.

---

## Technologies Used

`SharePoint Framework (SPFx)` · `Azure Functions .NET 8` · `Azure Table Storage` · `Azure Storage Queue` · `Microsoft Graph API` · `Copilot aiInteractionHistory API` · `Azure OpenAI Service` · `Microsoft 365 Agents Toolkit` · `TypeSpec` · `Microsoft Teams Activity Feed` · `Entra ID / Managed Identity` · `Application Insights`

---

## Video Demo

[https://youtu.be/2YKaxMCoer4](https://youtu.be/2YKaxMCoer4)

---

## Author

**Federico Porceddu** — [LinkedIn](https://www.linkedin.com/in/federicoporceddu/) · [GitHub](https://github.com/fredupstair)
