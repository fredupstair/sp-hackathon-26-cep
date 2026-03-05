<#
.SYNOPSIS
    Imposta le Tenant Properties di SharePoint per la soluzione CEP.

    Le due chiavi scritte sono:
      - CEP_FunctionAppBaseUrl   → URL base della Function App Azure
      - CEP_FunctionAppClientId → Client ID dell'App Registration CEP-Backend

    Le Tenant Properties (Storage Entities) sono leggibili da qualsiasi
    site collection tramite l'endpoint REST:
      GET /_api/web/GetStorageEntity('<key>')

    Devono essere SCRITTE dal sito di amministrazione SharePoint
    (https://<tenant>-admin.sharepoint.com).

.PARAMETER AdminSiteUrl
    URL del sito di amministrazione SharePoint.
    Es: https://federicoporceddumvp-admin.sharepoint.com

.PARAMETER FunctionAppBaseUrl
    URL base della Azure Function App (senza trailing slash).
    Default: https://func-cephackathon.azurewebsites.net

.PARAMETER FunctionAppClientId
    Client ID (appId) dell'App Registration CEP-Backend in Entra ID.
    Default: 0cb84638-30db-4bbd-936f-54a599840aec

.PARAMETER PnpClientId
    Non più necessario (ora si usa il modulo ufficiale Microsoft.Online.SharePoint.PowerShell).
    Mantenuto per retrocompatibilità con eventuali chiamate precedenti, viene ignorato.

.EXAMPLE
    .\set-tenant-properties.ps1 `
        -AdminSiteUrl   "https://federicoporceddumvp-admin.sharepoint.com" `
        -PnpClientId    "dd024163-fc77-44ee-8389-1ff0b5e8da2a"

.EXAMPLE
    # Sovrascrivere i valori di default
    .\set-tenant-properties.ps1 `
        -AdminSiteUrl        "https://contoso-admin.sharepoint.com" `
        -FunctionAppBaseUrl  "https://func-cep-prod.azurewebsites.net" `
        -FunctionAppClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -PnpClientId         "dd024163-fc77-44ee-8389-1ff0b5e8da2a"
#>
param(
    [Parameter(Mandatory)][string]$AdminSiteUrl,

    [string]$FunctionAppBaseUrl   = "https://func-cephackathon.azurewebsites.net",
    [string]$FunctionAppClientId  = "0cb84638-30db-4bbd-936f-54a599840aec",

    # Non più necessario – ignorato, mantenuto per retrocompatibilità
    [string]$PnpClientId = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── 1. Verifica modulo PnP PowerShell ────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "Il modulo PnP.PowerShell non è installato. Eseguire: Install-Module PnP.PowerShell -Scope CurrentUser"
    exit 1
}

# ── 2. Connessione all'admin per ricavare l'App Catalog URL ───────────────────
Write-Host "`n[1/3] Connessione a $AdminSiteUrl ..."
Connect-PnPOnline -Url $AdminSiteUrl -ClientId $PnpClientId -Interactive

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

# Riconnessione all'App Catalog per scrivere le Storage Entities
Disconnect-PnPOnline
Connect-PnPOnline -Url $appCatalogUrl -ClientId $PnpClientId -Interactive

# ── 3. Scrittura e verifica Storage Entities ──────────────────────────────────
$properties = @(
    @{ Key = "CEP_FunctionAppBaseUrl";  Value = $FunctionAppBaseUrl;  Description = "CEP – Azure Function App base URL" },
    @{ Key = "CEP_FunctionAppClientId"; Value = $FunctionAppClientId; Description = "CEP – App Registration Client ID (Entra ID)" }
)

Write-Host "`n[3/3] Scrittura Tenant Properties ..."
foreach ($prop in $properties) {
    Set-PnPStorageEntity `
        -Key         $prop.Key `
        -Value       $prop.Value `
        -Description $prop.Description

    $entity = Get-PnPStorageEntity -Key $prop.Key
    if ($entity.Value -eq $prop.Value) {
        Write-Host "  ✓ $($prop.Key) = $($prop.Value)"
    } else {
        Write-Warning "  ⚠ $($prop.Key): valore atteso '$($prop.Value)', trovato '$($entity.Value)'"
    }
}

Write-Host "`nTenant Properties impostate con successo.`n"
Disconnect-PnPOnline
