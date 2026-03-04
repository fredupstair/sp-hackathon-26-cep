<#
.SYNOPSIS
    Crea le risorse Azure per CEP e deploya la Function App.

    Risorse create:
      - Resource Group
      - Storage Account (+ queue cep-ingest)
      - Application Insights
      - Function App  (Consumption, dotnet-isolated, .NET 8, Windows)

    Poi esegue:
      - dotnet publish (Release)
      - Zip deploy tramite Azure CLI

.PARAMETER SubscriptionId
    GUID della sottoscrizione Azure.

.PARAMETER TenantId
    GUID del tenant Entra ID.

.PARAMETER Location
    Regione Azure (default: westeurope).

.PARAMETER ResourceGroupName
    Nome resource group (default: rg-cep-hackathon).

.PARAMETER BaseName
    Prefisso usato per tutti i nomi di risorsa (default: cephackathon).
    Deve essere alfanumerico lowercase, max ~16 caratteri.

.PARAMETER AppRegClientId
    ClientId dell'App Registration CEP-Backend (output di create-app-reg.ps1).

.PARAMETER AppRegClientSecret
    Client secret dell'App Registration CEP-Backend.

.PARAMETER SpSiteId
    Site ID SharePoint nel formato Graph (<host>,<siteId>,<webId>).
    Ottenibile con get-sp-ids.ps1.

.PARAMETER SpSiteUrl
    URL completo del sito SharePoint CEP.

.PARAMETER ListId_Users
.PARAMETER ListId_ActivityLog
.PARAMETER ListId_Leaderboard
.PARAMETER ListId_Badges
.PARAMETER ListId_Config
.PARAMETER ListId_SyncState
    GUID delle liste SharePoint CEP_* (output di get-sp-ids.ps1).
    Se non forniti, vengono lasciati vuoti e stampato un reminder.

.PARAMETER TeamsAppId
    App ID Teams per le notifiche (opzionale, può essere configurato dopo).

.PARAMETER SkipDeploy
    Se specificato, crea solo le risorse Azure senza effettuare il deploy del codice.

.EXAMPLE
    # Crea risorse + deploy completo
    .\deploy-azure.ps1 `
        -SubscriptionId "00000000-0000-0000-0000-000000000000" `
        -TenantId       "00000000-0000-0000-0000-000000000000" `
        -AppRegClientId "00000000-0000-0000-0000-000000000000" `
        -AppRegClientSecret "secret" `
        -SpSiteId "tenant.sharepoint.com,<siteGuid>,<webGuid>" `
        -SpSiteUrl "https://tenant.sharepoint.com/sites/Copilot-Engagement-Program"

.EXAMPLE
    # Solo infrastruttura, deploy manuale dopo
    .\deploy-azure.ps1 -SubscriptionId "..." -TenantId "..." -SkipDeploy
#>
param(
    [Parameter(Mandatory)][string]$SubscriptionId,
    [Parameter(Mandatory)][string]$TenantId,
    [string]$Location          = "westeurope",
    [string]$ResourceGroupName = "rg-cep-hackathon",
    [string]$BaseName          = "cephackathon",

    # App Registration (da create-app-reg.ps1)
    [string]$AppRegClientId     = "",
    [string]$AppRegClientSecret = "",

    # SharePoint (da get-sp-ids.ps1)
    [string]$SpSiteId         = "",
    [string]$SpSiteUrl        = "",
    [string]$ListId_Users         = "",
    [string]$ListId_ActivityLog   = "",
    [string]$ListId_Leaderboard   = "",
    [string]$ListId_Badges        = "",
    [string]$ListId_Config        = "",
    [string]$ListId_SyncState     = "",

    # Teams
    [string]$TeamsAppId = "",

    [switch]$SkipDeploy
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── Helpers ────────────────────────────────────────────────────────────────
function Info  { param([string]$m) Write-Host "  $m" -ForegroundColor Cyan }
function Ok    { param([string]$m) Write-Host "  ✓ $m" -ForegroundColor Green }
function Warn  { param([string]$m) Write-Host "  ⚠ $m" -ForegroundColor Yellow }
function Fatal { param([string]$m) Write-Host "  ✗ $m" -ForegroundColor Red; exit 1 }
function Section { param([string]$m) Write-Host ""; Write-Host "=== $m ===" -ForegroundColor Magenta }

function Invoke-AzSilent {
    param([string[]]$Arguments)
    $out = az @Arguments 2>$null
    if ($LASTEXITCODE -ne 0) {
        $err = az @Arguments 2>&1 | Select-Object -Last 3
        Fatal "az $($Arguments[0..2] -join ' ') ... failed: $err"
    }
    return $out
}

# ─── Derived names ──────────────────────────────────────────────────────────
$storageAccountName  = ($BaseName -replace '[^a-z0-9]', '')
if ($storageAccountName.Length -gt 24) { $storageAccountName = $storageAccountName.Substring(0, 24) }
$appInsightsName     = "appi-$BaseName"
$functionAppName     = "func-$BaseName"
$queueName           = "cep-ingest"
$repoRoot            = Split-Path $PSScriptRoot -Parent
$beDir               = Join-Path $repoRoot "be"
$publishDir          = Join-Path $beDir "bin\Release\net8.0\publish"
$zipPath             = Join-Path $env:TEMP "cep-functions-deploy.zip"

# ─── Banner ─────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Magenta
Write-Host "║       CEP – Azure Resources + Function Deploy       ║" -ForegroundColor Magenta
Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Magenta
Write-Host ""
Write-Host "  Subscription : $SubscriptionId"
Write-Host "  Tenant       : $TenantId"
Write-Host "  Location     : $Location"
Write-Host "  Rg           : $ResourceGroupName"
Write-Host "  Storage      : $storageAccountName"
Write-Host "  Function App : $functionAppName"
Write-Host "  App Insights : $appInsightsName"
Write-Host ""

# ─── 0. Prerequisites ───────────────────────────────────────────────────────
Section "0. Checking prerequisites"

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    Fatal "Azure CLI non trovato. Installa da: https://aka.ms/installazurecli"
}
Ok "Azure CLI found"

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Fatal "dotnet SDK non trovato. Installa da: https://dot.net"
}
Ok "dotnet SDK found"

# ─── 1. Login + subscription ────────────────────────────────────────────────
Section "1. Azure CLI login"

$account = az account show --output json 2>$null | ConvertFrom-Json
if (-not $account -or $account.tenantId -ne $TenantId) {
    Info "Eseguo login nel tenant $TenantId..."
    az login --tenant $TenantId --output none
    $account = az account show --output json | ConvertFrom-Json
}
Ok "Loggato come: $($account.user.name)"

az account set --subscription $SubscriptionId --output none
Ok "Sottoscrizione attiva: $SubscriptionId"

# ─── 2. Resource Group ──────────────────────────────────────────────────────
Section "2. Resource Group"

$rgExists = az group exists --name $ResourceGroupName --output tsv 2>$null
if ($rgExists -eq "true") {
    Warn "Resource group '$ResourceGroupName' già esistente – riuso."
} else {
    Info "Creo resource group '$ResourceGroupName' in '$Location'..."
    Invoke-AzSilent @("group","create","--name",$ResourceGroupName,"--location",$Location,"--output","none") | Out-Null
    Ok "Resource group creato"
}

# ─── 3. Storage Account + Queue ─────────────────────────────────────────────
Section "3. Storage Account + Queue"

$stExists = az storage account show --name $storageAccountName --resource-group $ResourceGroupName --output tsv --query "name" 2>$null
if ($stExists -eq $storageAccountName) {
    Warn "Storage account '$storageAccountName' già esistente – riuso."
} else {
    Info "Creo storage account '$storageAccountName'..."
    Invoke-AzSilent @("storage","account","create",
        "--name",$storageAccountName,
        "--resource-group",$ResourceGroupName,
        "--location",$Location,
        "--sku","Standard_LRS",
        "--kind","StorageV2",
        "--allow-blob-public-access","false",
        "--min-tls-version","TLS1_2",
        "--output","none") | Out-Null
    Ok "Storage account creato"
}

# Recupera la connection string per la queue
$storageConnStr = (az storage account show-connection-string `
    --name $storageAccountName `
    --resource-group $ResourceGroupName `
    --query connectionString --output tsv 2>$null).Trim()

# Crea la queue
$queueExists = az storage queue exists --name $queueName --connection-string $storageConnStr --query exists --output tsv 2>$null
if ($queueExists -eq "true") {
    Warn "Queue '$queueName' già esistente – skip."
} else {
    Info "Creo queue '$queueName'..."
    az storage queue create --name $queueName --connection-string $storageConnStr --output none 2>$null
    Ok "Queue '$queueName' creata"
}

# ─── 4. Application Insights ────────────────────────────────────────────────
Section "4. Application Insights"

# Ensure extension is present before any app-insights command
Info "Checking application-insights CLI extension..."
az extension add --name application-insights --only-show-errors 2>$null | Out-Null
Ok "application-insights extension ready"

$appInsightsJson = $null
try {
    $appInsightsJson = az monitor app-insights component show `
        --app $appInsightsName `
        --resource-group $ResourceGroupName `
        --output json 2>$null | ConvertFrom-Json
} catch { <# not found – will create below #> }

if ($appInsightsJson -and $appInsightsJson.name) {
    Warn "App Insights '$appInsightsName' già esistente – riuso."
    $aiConnStr = $appInsightsJson.connectionString
} else {
    Info "Creo Application Insights '$appInsightsName'..."
    $aiResult = az monitor app-insights component create `
        --app $appInsightsName `
        --resource-group $ResourceGroupName `
        --location $Location `
        --application-type web `
        --output json 2>$null | ConvertFrom-Json
    $aiConnStr = $aiResult.connectionString
    Ok "App Insights creato"
}

if (-not $aiConnStr) {
    Warn "Connection string App Insights non recuperata – impostala manualmente."
    $aiConnStr = ""
}

# ─── 5. Function App (Consumption, Windows, dotnet-isolated, .NET 8) ────────
Section "5. Function App"

$funcExists = az functionapp show --name $functionAppName --resource-group $ResourceGroupName --query name --output tsv 2>$null
if ($funcExists -eq $functionAppName) {
    Warn "Function App '$functionAppName' già esistente – aggiorno solo le impostazioni."
} else {
    Info "Creo Function App '$functionAppName' (Consumption, .NET 8 isolated)..."
    Invoke-AzSilent @("functionapp","create",
        "--name",$functionAppName,
        "--resource-group",$ResourceGroupName,
        "--storage-account",$storageAccountName,
        "--consumption-plan-location",$Location,
        "--runtime","dotnet-isolated",
        "--runtime-version","8",
        "--functions-version","4",
        "--os-type","Windows",
        "--output","none") | Out-Null
    Ok "Function App creata"
}

# ─── 6. App Settings ────────────────────────────────────────────────────────
Section "6. App Settings"

$settings = [ordered]@{
    "FUNCTIONS_WORKER_RUNTIME"               = "dotnet-isolated"
    "APPLICATIONINSIGHTS_CONNECTION_STRING"  = $aiConnStr
    "AzureWebJobsStorage"                    = $storageConnStr
    "IngestQueueName"                        = $queueName

    "TenantId"                               = $TenantId
    "ClientId"                               = $AppRegClientId
    "ClientSecret"                           = $AppRegClientSecret

    "SpSiteId"                               = $SpSiteId
    "SpSiteUrl"                              = $SpSiteUrl
    "ListId_Users"                           = $ListId_Users
    "ListId_ActivityLog"                     = $ListId_ActivityLog
    "ListId_Leaderboard"                     = $ListId_Leaderboard
    "ListId_Badges"                          = $ListId_Badges
    "ListId_Config"                          = $ListId_Config
    "ListId_SyncState"                       = $ListId_SyncState

    "TeamsAppId"                             = $TeamsAppId
}

# Build the --settings argument list: KEY=VALUE ...
$settingsArgs = @()
foreach ($kv in $settings.GetEnumerator()) {
    $settingsArgs += "$($kv.Key)=$($kv.Value)"
}

Info "Applico $($settingsArgs.Count) impostazioni alla Function App..."
# az functionapp config appsettings set accepts --settings as multiple KEY=VALUE
$azArgs = @("functionapp","config","appsettings","set",
    "--name",$functionAppName,
    "--resource-group",$ResourceGroupName,
    "--settings") + $settingsArgs + @("--output","none")
Invoke-AzSilent $azArgs | Out-Null
Ok "App Settings configurati"

# Segnala impostazioni vuote
$emptyKeys = $settings.GetEnumerator() | Where-Object { [string]::IsNullOrWhiteSpace($_.Value) } | Select-Object -ExpandProperty Key
if ($emptyKeys) {
    Warn "Le seguenti impostazioni sono vuote e vanno completate:"
    foreach ($k in $emptyKeys) { Write-Host "      - $k" -ForegroundColor Yellow }
    Write-Host "    Usa: az functionapp config appsettings set --name $functionAppName --resource-group $ResourceGroupName --settings KEY=VALUE" -ForegroundColor DarkGray
}

# ─── 7. Build + Publish + Deploy ────────────────────────────────────────────
if ($SkipDeploy) {
    Warn "SkipDeploy specificato – salto build e deploy. Risorse Azure pronte."
} else {
    Section "7. dotnet publish"

    Info "Eseguo 'dotnet publish' in Release dal progetto be/..."
    Push-Location $beDir
    dotnet publish --configuration Release --output $publishDir
    if ($LASTEXITCODE -ne 0) { Fatal "dotnet publish fallito." }
    Pop-Location
    Ok "Publish completato in: $publishDir"

    Section "8. Zip + Deploy"

    # Comprimi il contenuto della publish dir (non la cartella stessa)
    Info "Creo archivio zip: $zipPath"
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
    Compress-Archive -Path "$publishDir\*" -DestinationPath $zipPath -Force
    $zipSizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
    Ok "Zip creato ($zipSizeMB MB)"

    Info "Deploy in corso su '$functionAppName'..."
    az functionapp deployment source config-zip `
        --name $functionAppName `
        --resource-group $ResourceGroupName `
        --src $zipPath `
        --output none
    if ($LASTEXITCODE -ne 0) { Fatal "Zip deploy fallito." }
    Ok "Deploy completato"

    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
}

# ─── 9. Output riepilogo ────────────────────────────────────────────────────
$funcUrl = (az functionapp show `
    --name $functionAppName `
    --resource-group $ResourceGroupName `
    --query "hostNames[0]" --output tsv 2>$null).Trim()

$functionAppBaseUrl = "https://$funcUrl"

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║                   DEPLOY COMPLETATO                  ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "  Resource Group : $ResourceGroupName"
Write-Host "  Function App   : $functionAppName"
Write-Host "  URL            : $functionAppBaseUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "=== Copia in local.settings.json (per sviluppo locale) ===" -ForegroundColor Yellow
Write-Host ""
Write-Host "  `"AzureWebJobsStorage`": `"$storageConnStr`","
Write-Host "  `"APPLICATIONINSIGHTS_CONNECTION_STRING`": `"$aiConnStr`","
Write-Host ""
Write-Host "=== SPFx – Property Pane 'Function App Base URL' ===" -ForegroundColor Yellow
Write-Host ""
Write-Host "  $functionAppBaseUrl" -ForegroundColor Cyan
Write-Host ""
Write-Host "=== fe/config/package-solution.json – webApiPermissionRequests ===" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Assicurati che 'resource' sia il display name dell'App Registration" -ForegroundColor DarkGray
Write-Host "  (default 'CEP-Backend') oppure l'App ID URI (api://<clientId>)" -ForegroundColor DarkGray
Write-Host ""
if ($emptyKeys -and @($emptyKeys).Count -gt 0) {
    Write-Host "=== ⚠ AZIONI RIMANENTI ===" -ForegroundColor Yellow
    Write-Host "  Esegui i seguenti script se non già fatto:" -ForegroundColor Yellow
    Write-Host "    .\create-app-reg.ps1   → ottieni ClientId, ClientSecret, App ID URI"
    Write-Host "    .\get-sp-ids.ps1       → ottieni SpSiteId e ListId_*"
    Write-Host "  Poi ri-lancia questo script con tutti i parametri, oppure:" -ForegroundColor DarkGray
    Write-Host "    az functionapp config appsettings set \" -ForegroundColor DarkGray
    Write-Host "        --name $functionAppName --resource-group $ResourceGroupName \" -ForegroundColor DarkGray
    Write-Host "        --settings ClientId=xxx ClientSecret=yyy SpSiteId=zzz ..." -ForegroundColor DarkGray
    Write-Host ""
}
