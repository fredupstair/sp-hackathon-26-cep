# Get the Teams catalog app ID for the CEP app
# Uses the CEP app registration (client credentials) which has Graph app permissions

$tenantId   = "69d000be-840d-4db5-be28-0965a4fd72fc"
$clientId   = "0cb84638-30db-4bbd-936f-54a599840aec"
$manifestId = "1336459f-9891-4da1-8f8f-300a2368c01b"

# Read client secret from local.settings.json
$settings = Get-Content "$PSScriptRoot\..\be\local.settings.json" | ConvertFrom-Json
$clientSecret = $settings.Values.ClientSecret

# Get app-only token
$body = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = "https://graph.microsoft.com/.default"
}
$token = (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body).access_token
$headers = @{ Authorization = "Bearer $token" }

# Query catalog for our app
try {
    $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?`$filter=externalId eq '$manifestId'" -Headers $headers
    if ($resp.value -and $resp.value.Count -gt 0) {
        $catalogId = $resp.value[0].id
        Write-Host "Teams Catalog App ID: $catalogId" -ForegroundColor Green
        Write-Host "Manifest ID (externalId): $manifestId"
        Write-Host "Display Name: $($resp.value[0].displayName)"
        Write-Host ""
        Write-Host "Set this in Azure App Settings:" -ForegroundColor Yellow
        Write-Host "  TeamsAppId = $catalogId"
    } else {
        Write-Host "App not found by externalId. Listing org apps:" -ForegroundColor Yellow
        $all = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?`$filter=distributionMethod eq 'organization'" -Headers $headers
        foreach ($app in $all.value) {
            Write-Host "  ID: $($app.id) | ExternalId: $($app.externalId) | Name: $($app.displayName)"
        }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "The app registration may need AppCatalog.Read.All permission." -ForegroundColor Yellow
    Write-Host "Add it with:" -ForegroundColor Yellow
    Write-Host "  az ad app permission add --id $clientId --api 00000003-0000-0000-c000-000000000000 --api-permissions e12dae10-5a57-4571-a376-990bdd3e1dfc=Role"
    Write-Host "  az ad app permission admin-consent --id $clientId"
}
