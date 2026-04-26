<#
.SYNOPSIS
    Applies the CEP PnP provisioning template to a SharePoint Online site.

.DESCRIPTION
    Uses PnP.PowerShell 3.1.0+ to create the 6 SharePoint lists required by the
    Copilot Engagement Program:
    CEP_Users, CEP_ActivityLog, CEP_Leaderboard, CEP_Badges, CEP_Config, CEP_SyncState.

    PnP.PowerShell 3.x retired the shared "PnP Management Shell" Azure AD app,
    so -DeviceLogin and -Interactive both require an explicit -ClientId.
    When -ClientId is not provided the script falls back to -UseWebLogin
    (legacy cookie-based auth that opens a browser window – no app registration needed).

.PARAMETER SiteUrl
    URL of the target SharePoint Online site.
    Example: https://contoso.sharepoint.com/sites/CopilotEngagement

.PARAMETER TemplatePath
    Path to the PnP XML template file.
    Default: .\templates\cep-lists.xml

.PARAMETER ClientId
    Azure AD (Entra ID) App Registration Client ID.
    Required for device-code and interactive login.
    The app needs Sites.FullControl.All (or Sites.Selected) with admin consent.
    If omitted, the script uses -UseWebLogin instead.

.PARAMETER Tenant
    Tenant hostname or GUID (e.g. "contoso.sharepoint.com" or "contoso.onmicrosoft.com").
    Auto-derived from -SiteUrl when not specified (recommended).
    Required by PnP.PowerShell 3.x when using -ClientId.

.PARAMETER Interactive
    Use browser-based interactive login (requires WebView2 and -ClientId).
    Ignored when -ClientId is not supplied.

.EXAMPLE
    # Recommended: device-code login with your own app registration
    .\deploy.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/CopilotEngagement" `
                 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    # No app registration needed – opens a browser window (legacy web login)
    .\deploy.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/CopilotEngagement"

.EXAMPLE
    # Browser-based login (requires WebView2)
    .\deploy.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/CopilotEngagement" `
                 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Interactive

.EXAMPLE
    # Dry-run (no changes applied)
    .\deploy.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/CopilotEngagement" -WhatIf

.NOTES
    Requires : PnP.PowerShell 3.1.0+
    Install  : Install-Module PnP.PowerShell -RequiredVersion 3.1.0 -Scope CurrentUser
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$TemplatePath = ".\templates\cep-lists.xml",

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── 0. Prerequisites ──────────────────────────────────────────────────────────
Write-Host "`n=== Copilot Engagement Program – SharePoint Provisioning ===" -ForegroundColor Cyan

$pnpModule = Get-Module -ListAvailable -Name PnP.PowerShell |
             Sort-Object Version -Descending |
             Select-Object -First 1
if (-not $pnpModule) {
    Write-Error "PnP.PowerShell module not found.`nInstall  : Install-Module PnP.PowerShell -RequiredVersion 3.1.0 -Scope CurrentUser"
}

$pnpVersion = $pnpModule.Version
if ($pnpVersion -lt [Version]"3.1.0") {
    Write-Warning "PnP.PowerShell $pnpVersion detected – 3.1.0+ is required. Upgrade with: Update-Module PnP.PowerShell"
} else {
    Write-Host "[OK] PnP.PowerShell $pnpVersion" -ForegroundColor Green
}

$templateFile = Resolve-Path $TemplatePath -ErrorAction Stop
Write-Host "[OK] Template : $templateFile" -ForegroundColor Green

# ── 1. Connect ────────────────────────────────────────────────────────────────
Write-Host "`n[1/3] Connecting to $SiteUrl ..." -ForegroundColor Yellow

# Derive tenant from SiteUrl when not explicitly provided
# e.g. https://contoso.sharepoint.com/sites/... → contoso.onmicrosoft.com
if (-not $Tenant -and $ClientId) {
    $spHost = ([Uri]$SiteUrl).Host                        # <TENANT>.sharepoint.com
    $subdomain = $spHost.Split('.')[0]                    # <TENANT>
    $Tenant = "$subdomain.onmicrosoft.com"
    Write-Host "      Tenant : $Tenant (auto-derived from SiteUrl)" -ForegroundColor DarkGray
}

if ($ClientId -and $Interactive) {
    Write-Host "      Method : Interactive / browser (WebView2)" -ForegroundColor Gray
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -Tenant $Tenant
} elseif ($ClientId) {
    Write-Host "      Method : Device code login (open the URL above and enter the code)" -ForegroundColor Gray
    Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $ClientId -Tenant $Tenant
} else {
    Write-Host "      Method : Web login / cookie-based (a browser window will open)" -ForegroundColor Gray
    Write-Host "      Tip    : pass -ClientId <entra-app-id> to use device-code login instead" -ForegroundColor DarkGray
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
}

# ── 2. Verify site ────────────────────────────────────────────────────────────
Write-Host "`n[2/3] Verifying site ..." -ForegroundColor Yellow
$web = Get-PnPWeb -Includes Title,Url,CurrentUser
Write-Host "[OK] Site  : $($web.Title)" -ForegroundColor Green
Write-Host "     URL   : $($web.Url)" -ForegroundColor Green
if ($web.CurrentUser) {
    Write-Host "     User  : $($web.CurrentUser.LoginName)" -ForegroundColor Green
}

# ── 3. Apply template ─────────────────────────────────────────────────────────
Write-Host "`n[3/3] Applying PnP template ..." -ForegroundColor Yellow

if ($PSCmdlet.ShouldProcess($SiteUrl, "Invoke-PnPSiteTemplate")) {
    Invoke-PnPSiteTemplate -Path $templateFile
    Write-Host "[OK] Template applied successfully!" -ForegroundColor Green
}

# ── 4. Summary ────────────────────────────────────────────────────────────────
Write-Host "`n=== CEP lists ===" -ForegroundColor Cyan
$cepLists = @("CEP_Users","CEP_ActivityLog","CEP_Leaderboard","CEP_Badges","CEP_Config","CEP_SyncState")
foreach ($listName in $cepLists) {
    try {
        $list = Get-PnPList -Identity $listName -ErrorAction Stop
        Write-Host "  [OK] $($list.Title.PadRight(22)) $($list.ItemCount) items  (ID: $($list.Id))" -ForegroundColor Green
    } catch {
        Write-Host "  [!!] $($listName.PadRight(22)) NOT FOUND – check provisioning log above" -ForegroundColor Red
    }
}

Write-Host "`n=== Provisioning complete ===" -ForegroundColor Cyan
Write-Host "Site    : $SiteUrl"
Write-Host "Template: $templateFile"
Write-Host ""
