# run-local.ps1
# Avvia Azure Functions in locale usando func.exe (winget install).
# Prerequisito: Azurite in esecuzione  (npx azurite  o  Docker)

$func = "C:\Program Files\Microsoft\Azure Functions Core Tools\func.exe"
$projectDir = $PSScriptRoot

Write-Host "Starting Azurite (Storage Emulator) in background..." -ForegroundColor Cyan
$azurite = Start-Process -FilePath "npx" -ArgumentList "azurite --silent --location $env:TEMP\azurite" `
    -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue

Write-Host "Starting Azure Functions host..." -ForegroundColor Green
Push-Location $projectDir
& $func start --dotnet-isolated-debug
Pop-Location

if ($azurite -and !$azurite.HasExited) { Stop-Process -Id $azurite.Id -Force }
