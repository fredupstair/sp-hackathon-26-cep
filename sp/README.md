# SharePoint Provisioning – Copilot Engagement Program

Contiene i template **PnP PowerShell** per creare le liste SharePoint usate dal sistema CEP.

---

## Struttura

```
sp/
├── templates/
│   └── cep-lists.xml        # Template PnP con tutte e 6 le liste + site columns
├── seed-data/               # (opzionale) dati di test
├── deploy.ps1               # Script di deployment
└── README.md
```

---

## Liste create dal template

| Lista | URL interna | Descrizione |
|-------|-------------|-------------|
| `CEP_Users` | `Lists/CEP_Users` | Utenti iscritti al programma |
| `CEP_ActivityLog` | `Lists/CEP_ActivityLog` | Aggregati giornalieri di utilizzo per utente/app |
| `CEP_Leaderboard` | `Lists/CEP_Leaderboard` | Classifica mensile materializzata (Global/Department/Team) |
| `CEP_Badges` | `Lists/CEP_Badges` | Badge conquistati dagli utenti |
| `CEP_Config` | `Lists/CEP_Config` | Configurazione runtime (editabile da admin) |
| `CEP_SyncState` | `Lists/CEP_SyncState` | Stato e log delle run di sincronizzazione |

---

## Prerequisiti

1. **PnP.PowerShell 2.x+**
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   ```
2. Sito SharePoint Online già esistente (Communication Site o Team Site).
3. Utente con permessi **Site Owner** (o Site Collection Administrator).

---

## Deployment

### Autenticazione interattiva (browser)

```powershell
cd sp
.\deploy.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/CopilotEngagement"
```

### Device Login (headless / CI)

```powershell
.\deploy.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/CopilotEngagement" -UseDeviceLogin
```

### Dry-run (WhatIf)

```powershell
.\deploy.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/CopilotEngagement" -WhatIf
```

### Solo applicare il template PnP (senza lo script wrapper)

```powershell
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/CopilotEngagement" -Interactive
Invoke-PnPSiteTemplate -Path .\templates\cep-lists.xml -Handlers Fields,Lists
```

---

## Colonne create (Site Columns – Gruppo "CEP Columns")

### CEP_Users
| Colonna interna | Tipo | Note |
|-----------------|------|------|
| `CEP_UserEmail` | Text | Indexed |
| `CEP_UserPrincipalName` | Text | Indexed |
| `CEP_AadUserId` | Text | Indexed – chiave tecnica primaria |
| `CEP_Department` | Text | Indexed |
| `CEP_Team` | Text | Indexed |
| `CEP_EnrollmentDate` | DateTime | |
| `CEP_CurrentLevel` | Choice | Explorer / Practitioner / Master |
| `CEP_TotalPoints` | Number | |
| `CEP_MonthlyPoints` | Number | |
| `CEP_IsActive` | Boolean | Default: true |
| `CEP_LastActivityDate` | DateTime | Indexed |
| `CEP_IsEngagementNudgesEnabled` | Boolean | Default: true |
| `CEP_LastSyncWatermarkUtc` | DateTime | Watermark per ingestion incrementale |
| `CEP_LastNudgeSentUtc` | DateTime | Anti-spam notifiche |
| `CEP_LastLeaderboardNotifiedMonth` | Text | Idempotenza notifiche leaderboard |

### CEP_ActivityLog
| Colonna interna | Tipo | Note |
|-----------------|------|------|
| `CEP_Log_AadUserId` | Text | Indexed |
| `CEP_Log_UserEmail` | Text | |
| `CEP_Log_UsageDate` | DateTime (DateOnly) | Indexed |
| `CEP_Log_AppKey` | Choice | Word/Excel/PPT/Outlook/Teams/OneNote/Loop/BizChat/WebChat/M365App/Forms/SharePoint/Whiteboard |
| `CEP_Log_PromptCount` | Number | |
| `CEP_Log_PointsEarned` | Number | |
| `CEP_Log_MonthKey` | Text | `YYYY-MM`, Indexed |

### CEP_Leaderboard
| Colonna interna | Tipo | Note |
|-----------------|------|------|
| `CEP_LB_UserEmail` | Text | |
| `CEP_LB_DisplayName` | Text | |
| `CEP_LB_Department` | Text | |
| `CEP_LB_Team` | Text | |
| `CEP_LB_AadUserId` | Text | |
| `CEP_LB_MonthlyPoints` | Number | |
| `CEP_LB_Rank` | Number | Standard competition ranking |
| `CEP_LB_Level` | Choice | Explorer/Practitioner/Master |
| `CEP_LB_Scope` | Choice | Global/Department/Team, Indexed |
| `CEP_LB_MonthKey` | Text | `YYYY-MM`, Indexed |

### CEP_Badges
| Colonna interna | Tipo | Note |
|-----------------|------|------|
| `CEP_Badge_UserEmail` | Text | |
| `CEP_Badge_AadUserId` | Text | Indexed |
| `CEP_Badge_BadgeKey` | Text | Indexed – chiave univoca per deduplicazione |
| `CEP_Badge_BadgeType` | Choice | FirstSteps/CrossAppExplorer/WeeklyWarrior/MonthlyMaster/ConsistencyKing |
| `CEP_Badge_EarnedDate` | DateTime | |
| `CEP_Badge_Description` | Note | |
| `CEP_Badge_MonthKey` | Text | Opzionale per badge mensili |

### CEP_Config
| Colonna interna | Tipo | Default |
|-----------------|------|---------|
| `CEP_Cfg_SyncFrequency` | Text | `daily` |
| `CEP_Cfg_PointsPerPrompt` | Number | `1` |
| `CEP_Cfg_LevelThresholdSilver` | Number | `500` |
| `CEP_Cfg_LevelThresholdGold` | Number | `1500` |
| `CEP_Cfg_InactivityDaysForNudge` | Number | `3` |
| `CEP_Cfg_LeaderboardRefreshNotifEnabled` | Boolean | `true` |
| `CEP_Cfg_MaxUsersPerIngestionBatch` | Number | `50` |
| `CEP_Cfg_TimeZone` | Text | `UTC` |

> Il template inserisce automaticamente un record "Default" con i valori sopra.

### CEP_SyncState
| Colonna interna | Tipo | Note |
|-----------------|------|------|
| `CEP_Sync_LastSuccessfulRunUtc` | DateTime | |
| `CEP_Sync_LastRunStatus` | Choice | Success/Failure/Running |
| `CEP_Sync_LastRunCorrelationId` | Text | |
| `CEP_Sync_LastRunSummary` | Note | |
| `CEP_Sync_UsersProcessed` | Number | |
| `CEP_Sync_PromptsIngested` | Number | |

---

## Permessi consigliati

Dopo il provisioning, configurare i permessi sulle liste dal sito SharePoint:

| Lista | Service Principal / Managed Identity | Utenti finali |
|-------|--------------------------------------|---------------|
| CEP_Users | Full Control | Nessun accesso diretto (solo via API) |
| CEP_ActivityLog | Full Control | Nessun accesso diretto |
| CEP_Leaderboard | Full Control | Read (solo colonne pubbliche) |
| CEP_Badges | Full Control | Read (solo propri record via API) |
| CEP_Config | Full Control | Nessun accesso / Read per admin |
| CEP_SyncState | Full Control | Read per admin |

---

## Re-apply (idempotente)

Il template PnP è idempotente: rieseguire `deploy.ps1` sullo stesso sito non duplica liste o colonne già esistenti.
