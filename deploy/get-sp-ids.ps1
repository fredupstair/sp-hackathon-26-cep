<#
.SYNOPSIS
    Recupera i GUID del sito SharePoint e di tutte le liste CEP_*
    e li stampa nel formato pronto per local.settings.json.

.EXAMPLE
    .\get-sp-ids.ps1 -SiteUrl "https://federicoporceddumvp.sharepoint.com/sites/Copilot-Engagement-Program" `
                     -ClientId "dd024163-fc77-44ee-8389-1ff0b5e8da2a"
#>
param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$Tenant
)

if (-not $Tenant) {
    # Derive <prefix>.onmicrosoft.com from the SharePoint URL host
    $prefix = ([System.Uri]$SiteUrl).Host.Split('.')[0]   # e.g. "federicoporceddumvp"
    $Tenant = "$prefix.onmicrosoft.com"
}

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -Tenant $Tenant

# Site ID (Graph format: <hostname>,<site-collection-id>,<web-id>)
$web  = Get-PnPWeb -Includes Id
$site = Get-PnPSite -Includes Id
$graphSiteId = "$($web.Url.Split('/')[2]),$($site.Id.ToString('D')),$($web.Id.ToString('D'))"

Write-Host ""
Write-Host "=== Copy into local.settings.json / Azure Function App Settings ===" -ForegroundColor Yellow
Write-Host ""
Write-Host "  `"SpSiteId`": `"$graphSiteId`","

$lists = @("CEP_Users","CEP_ActivityLog","CEP_Leaderboard","CEP_Badges","CEP_Config","CEP_SyncState")
$keyMap = @{
    "CEP_Users"       = "ListId_Users"
    "CEP_ActivityLog" = "ListId_ActivityLog"
    "CEP_Leaderboard" = "ListId_Leaderboard"
    "CEP_Badges"      = "ListId_Badges"
    "CEP_Config"      = "ListId_Config"
    "CEP_SyncState"   = "ListId_SyncState"
}

foreach ($listName in $lists) {
    try {
        $list = Get-PnPList -Identity $listName
        Write-Host "  `"$($keyMap[$listName])`": `"$($list.Id.ToString('D'))`","
    } catch {
        Write-Warning "List '$listName' not found – run sp/deploy.ps1 first."
    }
}

Write-Host ""
Write-Host "Also needed:" -ForegroundColor Yellow
Write-Host "  TenantId: $(Get-PnPTenantId)"
