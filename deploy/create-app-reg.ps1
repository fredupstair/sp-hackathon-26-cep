<#
.SYNOPSIS
    Crea (o aggiorna) l'App Registration Entra ID per CEP backend.
    - Permessi applicativi Graph: AiEnterpriseInteraction.Read.All, TeamsActivity.Send, Sites.Selected
    - Grant admin consent automatico
    - Crea un client secret (validità 1 anno)
    - Assegna Sites.Selected al sito SharePoint CEP via PnP
    - Stampa i valori pronti per local.settings.json / Azure Function App Settings

.PARAMETER TenantId
    GUID del tenant Entra ID (es. 69d000be-840d-4db5-be28-0965a4fd72fc)

.PARAMETER SiteUrl
    URL del sito SharePoint CEP

.PARAMETER AppName
    Nome dell'app registration (default: "CEP-Backend")

.PARAMETER PnpClientId
    ClientId dell'app PnP già registrata (usata per connettersi a SharePoint con -Interactive)

.EXAMPLE
    .\create-app-reg.ps1 `
        -TenantId   "69d000be-840d-4db5-be28-0965a4fd72fc" `
        -SiteUrl    "https://federicoporceddumvp.sharepoint.com/sites/Copilot-Engagement-Program" `
        -PnpClientId "dd024163-fc77-44ee-8389-1ff0b5e8da2a"
#>
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$SiteUrl,
    [string]$AppName     = "CEP-Backend",
    [string]$PnpClientId = "dd024163-fc77-44ee-8389-1ff0b5e8da2a"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Info  { param([string]$m) Write-Host "  $m" -ForegroundColor Cyan }
function Ok    { param([string]$m) Write-Host "  ✓ $m" -ForegroundColor Green }
function Warn  { param([string]$m) Write-Host "  ⚠ $m" -ForegroundColor Yellow }
function Fatal { param([string]$m) Write-Host "  ✗ $m" -ForegroundColor Red; exit 1 }

function Invoke-Az {
    param([string[]]$Args)
    # Redirect stderr to avoid warnings polluting JSON stdout
    $result = az @Args 2>$null
    if ($LASTEXITCODE -ne 0) {
        $errDetail = az @Args 2>&1 | Where-Object { $_ -notmatch '^WARNING' } | Select-Object -First 3
        Fatal "az $($Args -join ' ') failed: $errDetail"
    }
    # Strip any non-JSON preamble lines (e.g. // comments sometimes emitted by az CLI)
    $jsonLines = $result | Where-Object { $_ -notmatch '^\s*//' }
    return $jsonLines | ConvertFrom-Json
}

# ---------------------------------------------------------------------------
# 0. Prerequisites
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "=== CEP – Create App Registration ===" -ForegroundColor Magenta
Write-Host ""

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    Fatal "Azure CLI non trovato. Installa da https://aka.ms/installazurecli"
}

# ---------------------------------------------------------------------------
# 1. Login Azure CLI
# ---------------------------------------------------------------------------
Info "Checking Azure CLI login..."
$account = az account show --output json 2>$null | ConvertFrom-Json
if (-not $account -or $account.tenantId -ne $TenantId) {
    Info "Logging in to tenant $TenantId..."
    az login --tenant $TenantId --allow-no-subscriptions --output none 2>$null
    $account = Invoke-Az @("account","show")
}
Ok "Logged in as $($account.user.name)"

# ---------------------------------------------------------------------------
# 2. Resolve Microsoft Graph permission IDs
#    az --query filters client-side so no huge JSON parse in PowerShell
# ---------------------------------------------------------------------------
Info "Resolving Microsoft Graph permission IDs..."

function Get-GraphRoleId {
    param([string]$RoleName)
    $id = az ad sp show --id "00000003-0000-0000-c000-000000000000" `
            --query "appRoles[?value=='$RoleName' && contains(allowedMemberTypes,'Application')].id | [0]" `
            --output tsv 2>$null
    if ([string]::IsNullOrWhiteSpace($id)) {
        Warn "Permission '$RoleName' not found – add it manually in Entra portal"
        return $null
    }
    Ok "  $RoleName = $id"
    return $id.Trim()
}

$idAiInteraction  = Get-GraphRoleId "AiEnterpriseInteraction.Read.All"
$idTeamsActivity  = Get-GraphRoleId "TeamsActivity.Send"
$idSitesSelected  = Get-GraphRoleId "Sites.Selected"

# ---------------------------------------------------------------------------
# 3. Create (or find) the App Registration
# ---------------------------------------------------------------------------
Info "Looking for existing app registration '$AppName'..."
$existingApps = az ad app list --display-name $AppName --output json 2>$null | ConvertFrom-Json
$app = $existingApps | Where-Object { $_.displayName -eq $AppName } | Select-Object -First 1

if ($app) {
    Warn "App '$AppName' already exists (appId: $($app.appId)) – reusing it."
} else {
    Info "Creating app registration '$AppName'..."
    $app = Invoke-Az @("ad","app","create","--display-name",$AppName,"--sign-in-audience","AzureADMyOrg")
    Ok "Created app: $($app.appId)"
}

$appId     = $app.appId
$appObjId  = $app.id

# ---------------------------------------------------------------------------
# 4. Add required Graph application permissions
# ---------------------------------------------------------------------------
Info "Setting required Graph permissions..."

$resourceAccess = @()
foreach ($roleId in @($idAiInteraction, $idTeamsActivity, $idSitesSelected)) {
    if ($roleId) { $resourceAccess += @{ id = $roleId; type = "Role" } }
}

$requiredAccess = @{
    resourceAppId  = "00000003-0000-0000-c000-000000000000"
    resourceAccess = $resourceAccess
} | ConvertTo-Json -Depth 5 -Compress

$requiredAccessJson = "[" + $requiredAccess + "]"

# Write to temp file to avoid shell escaping issues
$tmpFile = [System.IO.Path]::GetTempFileName() + ".json"
$requiredAccessJson | Set-Content -Path $tmpFile -Encoding UTF8

az ad app update --id $appObjId --required-resource-accesses "@$tmpFile" --output none 2>$null | Out-Null
Remove-Item $tmpFile -Force
Ok "Permissions configured"

# ---------------------------------------------------------------------------
# 5. Ensure service principal exists
# ---------------------------------------------------------------------------
Info "Ensuring service principal..."
$sp = az ad sp show --id $appId --output json 2>$null | ConvertFrom-Json
if (-not $sp) {
    $sp = Invoke-Az @("ad","sp","create","--id",$appId)
    Ok "Service principal created: $($sp.id)"
} else {
    Ok "Service principal exists: $($sp.id)"
}
$spObjId = $sp.id

# ---------------------------------------------------------------------------
# 6. Admin consent (grant application permissions)
# ---------------------------------------------------------------------------
Info "Granting admin consent..."

$graphSpId = (az ad sp show --id "00000003-0000-0000-c000-000000000000" --query id --output tsv 2>$null).Trim()

foreach ($roleId in @($idAiInteraction, $idTeamsActivity, $idSitesSelected)) {
    if (-not $roleId) { continue }
    # Use REST via az rest for appRoleAssignments (more reliable than az ad)
    $body = @{
        principalId = $spObjId
        resourceId  = $graphSpId
        appRoleId   = $roleId
    } | ConvertTo-Json -Compress

    $result = az rest --method POST `
        --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$spObjId/appRoleAssignments" `
        --headers "Content-Type=application/json" `
        --body $body 2>&1

    if ($LASTEXITCODE -eq 0) {
        Ok "Granted role $roleId"
    } else {
        # 409 = already granted, that's fine
        if ($result -match "Permission being assigned already exists" -or $result -match "already exists") {
            Warn "Role $roleId already granted – skipping"
        } else {
            Warn "Could not auto-grant $roleId : $result"
            Warn "Grant manually: Entra Portal → App registrations → API permissions → Grant admin consent"
        }
    }
}

# ---------------------------------------------------------------------------
# 7. Create client secret (1 year)
# ---------------------------------------------------------------------------
Info "Creating client secret (valid 1 year)..."
$secretResult = Invoke-Az @("ad","app","credential","reset",
    "--id",$appObjId,
    "--years","1",
    "--append",
    "--display-name","CEP-Local-$(Get-Date -Format 'yyyyMMdd')")

$clientSecret = $secretResult.password
Ok "Client secret created"

# ---------------------------------------------------------------------------
# 8. Grant Sites.Selected on the SharePoint site via PnP
# ---------------------------------------------------------------------------
Info "Granting Sites.Selected on $SiteUrl..."
$prefix = ([System.Uri]$SiteUrl).Host.Split('.')[0]
$pnpTenant = "$prefix.onmicrosoft.com"

try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $PnpClientId -Tenant $pnpTenant
    Grant-PnPAzureADAppSitePermission `
        -AppId $appId `
        -DisplayName $AppName `
        -Site $SiteUrl `
        -Permissions Write
    Ok "Sites.Selected granted (Write) on $SiteUrl"
} catch {
    Warn "Could not auto-grant Sites.Selected via PnP: $_"
    Warn "Grant manually:"
    Warn "  Grant-PnPAzureADAppSitePermission -AppId $appId -DisplayName $AppName -Site $SiteUrl -Permissions Write"
}

# ---------------------------------------------------------------------------
# 9. Output configuration values
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "=== Copy into local.settings.json / Azure Function App Settings ===" -ForegroundColor Yellow
Write-Host ""
Write-Host "  `"TenantId`":     `"$TenantId`","
Write-Host "  `"ClientId`":     `"$appId`","
Write-Host "  `"ClientSecret`": `"$clientSecret`","
Write-Host ""
Write-Host "  (SpSiteId and ListId_* → run .\get-sp-ids.ps1 if not yet done)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "IMPORTANT – also:" -ForegroundColor Yellow
Write-Host "  1. In Entra portal, verify API permissions show 'Granted' (green tick)."
Write-Host "  2. If AiEnterpriseInteraction.Read.All requires specific licensing, verify in tenant admin."
Write-Host "  3. TeamsActivity.Send requires the Teams App to be installed for target users."
Write-Host ""
