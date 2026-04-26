<#
.SYNOPSIS
    Imposta le Tenant Storage Entities di SharePoint per la soluzione CEP.

    Le chiavi scritte sono:
      - CEP_FunctionAppBaseUrl  -> URL base della Function App Azure
      - CEP_FunctionAppClientId -> Client ID dell'App Registration CEP-Backend
      - CEP_DashboardPageUrl    -> URL pagina Dashboard CEP (letta dall'App Customizer CepBar)
      - CEP_OptinPageUrl        -> URL pagina iscrizione CEP (letta dall'App Customizer CepBar)

    Le Tenant Properties (Storage Entities) sono leggibili da qualsiasi
    site collection tramite l'endpoint REST:
      GET /_api/web/GetStorageEntity('<key>')

    Devono essere SCRITTE dal sito di amministrazione SharePoint
    (https://<tenant>-admin.sharepoint.com).

.PARAMETER AdminSiteUrl
    URL del sito di amministrazione SharePoint.
    Es: https://<TENANT>-admin.sharepoint.com

.PARAMETER FunctionAppBaseUrl
    URL base della Azure Function App (senza trailing slash).
    Default: https://func-cephackathon.azurewebsites.net

.PARAMETER FunctionAppClientId
    Client ID (appId) dell'App Registration CEP-Backend in Entra ID.
    Default: 0cb84638-30db-4bbd-936f-54a599840aec

.PARAMETER DashboardPageUrl
    URL assoluto della pagina Dashboard CEP (pulsante "CEP Portal" nella barra).

.PARAMETER OptinPageUrl
    URL assoluto della pagina di iscrizione CEP.

.PARAMETER PnpClientId
    Client ID dell'app PnP per l'autenticazione interattiva.
    Default: dd024163-fc77-44ee-8389-1ff0b5e8da2a

.EXAMPLE
    .\set-tenant-properties.ps1 -AdminSiteUrl "https://<TENANT>-admin.sharepoint.com"

.EXAMPLE
    .\set-tenant-properties.ps1 `
        -AdminSiteUrl        "https://contoso-admin.sharepoint.com" `
        -FunctionAppBaseUrl  "https://func-cep-prod.azurewebsites.net" `
        -FunctionAppClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -DashboardPageUrl    "https://contoso.sharepoint.com/sites/CEP/SitePages/CEP-Dashboard.aspx" `
        -OptinPageUrl        "https://contoso.sharepoint.com/sites/CEP/SitePages/CEP-Onboarding.aspx"
#>
param(
    [string]$AdminSiteUrl = "https://<TENANT>-admin.sharepoint.com",

    [string]$FunctionAppBaseUrl   = "https://func-cephackathon.azurewebsites.net",
    [string]$FunctionAppClientId  = "0cb84638-30db-4bbd-936f-54a599840aec",

    [string]$DashboardPageUrl = "https://<TENANT>.sharepoint.com/sites/Copilot-Engagement-Program/SitePages/CePWelcome.aspx",
    [string]$OptinPageUrl     = "https://<TENANT>.sharepoint.com/sites/Copilot-Engagement-Program/SitePages/CePWelcome.aspx",

    [string]$PnpClientId = "dd024163-fc77-44ee-8389-1ff0b5e8da2a"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── 1. Verifica modulo PnP PowerShell ────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "Il modulo PnP.PowerShell non è installato. Eseguire: Install-Module PnP.PowerShell -Scope CurrentUser"
    exit 1
}

# Helper: Connect-PnPOnline senza -ClientId se non specificato
function Connect-Pnp([string]$Url) {
    if ($PnpClientId) {
        Connect-PnPOnline -Url $Url -ClientId $PnpClientId -Interactive
    } else {
        Connect-PnPOnline -Url $Url -Interactive
    }
}

# ── 2. Connessione all'admin per ricavare l'App Catalog URL ───────────────────
Write-Host "`n[1/3] Connessione a $AdminSiteUrl ..."
Connect-Pnp $AdminSiteUrl

Write-Host "`n[2/3] Preparazione App Catalog ..."
$appCatalogUrl = Get-PnPTenantAppCatalogUrl
if (-not $appCatalogUrl) {
    Write-Error "Impossibile trovare il Tenant App Catalog. Assicurati che esista un App Catalog nel tenant."
    exit 1
}
Write-Host "  App Catalog: $appCatalogUrl"

# Fix per Access Denied su Set-PnPStorageEntity (https://github.com/pnp/powershell/issues/5240):
# custom script deve essere abilitato sull'App Catalog prima di scrivere le storage entities.
Write-Host "  Abilitazione custom script sull'App Catalog ..."
Set-PnPSite -Identity $appCatalogUrl -NoScriptSite:$false
Set-PnPSite -Identity $appCatalogUrl -DenyAddAndCustomizePages:$false

# Riconnessione all'App Catalog per scrivere le Storage Entities (stesso tenant → token riusato, nessun nuovo login)
Connect-Pnp $appCatalogUrl

# ── 3. Scrittura e verifica Storage Entities ──────────────────────────────────
$properties = @(
    @{ Key = "CEP_FunctionAppBaseUrl";  Value = $FunctionAppBaseUrl;  Description = "CEP - Azure Function App base URL" },
    @{ Key = "CEP_FunctionAppClientId"; Value = $FunctionAppClientId; Description = "CEP - App Registration Client ID (Entra ID)" },
    @{ Key = "CEP_DashboardPageUrl";    Value = $DashboardPageUrl;    Description = "CEP - Dashboard page URL (CepBar customizer)" },
    @{ Key = "CEP_OptinPageUrl";        Value = $OptinPageUrl;        Description = "CEP - Opt-in page URL (CepBar customizer)" }
)

Write-Host "`n[3/3] Scrittura Tenant Properties ..."
foreach ($prop in $properties) {
    Set-PnPStorageEntity `
        -Key         $prop.Key `
        -Value       $prop.Value `
        -Description $prop.Description

    try {
        $entity = Get-PnPStorageEntity -Key $prop.Key
        if ($entity.Value -eq $prop.Value) {
            Write-Host "  ✓ $($prop.Key) = $($prop.Value)"
        } else {
            Write-Warning "  ⚠ $($prop.Key): valore atteso '$($prop.Value)', trovato '$($entity.Value)'"
        }
    } catch {
        Write-Host "  ✓ $($prop.Key) scritto (verifica lettura non disponibile)"
    }
}

Write-Host "`nTenant Properties impostate con successo.`n"
Disconnect-PnPOnline
