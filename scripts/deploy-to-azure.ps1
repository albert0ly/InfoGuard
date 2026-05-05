# deploy-to-azure.ps1
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Starting Azure Deployment Process" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

$ErrorActionPreference = "Stop"
$resourceGroup = "rg-InfoGuard-dev"
$appName = "app-InfoGuard-backend-dev01"
$appUrl = "https://$appName.azurewebsites.net"

# Step 1: Build
Write-Host "`n[1/9] Building production bundle..." -ForegroundColor Yellow
npm run build:prod
if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed!" -ForegroundColor Red
    exit 1
}

# Step 2: Create deployment package
Write-Host "`n[2/9] Creating deployment package..." -ForegroundColor Yellow
Remove-Item app.zip -ErrorAction SilentlyContinue
Remove-Item deploy-temp -Recurse -ErrorAction SilentlyContinue

New-Item -ItemType Directory -Path "deploy-temp" -Force | Out-Null

# Copy CONTENTS of dist to root of deploy-temp
Copy-Item "dist\*" -Destination "deploy-temp\" -Recurse
Copy-Item "assets" -Destination "deploy-temp\assets" -Recurse

# Create production-only package.json
Write-Host "  Creating production package.json..." -ForegroundColor Gray
$packageJson = Get-Content "package.json" -Raw | ConvertFrom-Json

# Create production package.json with proper structure
$prodPackageJson = @{
    name = $packageJson.name
    version = $packageJson.version
    private = $packageJson.private
    scripts = @{
        start = "node middletier.js"
    }
    dependencies = $packageJson.dependencies
}

# Add engines field - check if it exists in original, otherwise use default
if ($packageJson.PSObject.Properties['engines'] -and $packageJson.engines) {
    $prodPackageJson['engines'] = $packageJson.engines
} else {
    # Default engines specification for Azure
    $prodPackageJson['engines'] = @{
        node = ">=18.0.0"
    }
}

# Convert to JSON and save
$prodPackageJson | ConvertTo-Json -Depth 10 | Set-Content "deploy-temp\package.json" -Encoding UTF8

Copy-Item "package-lock.json" -Destination "deploy-temp\"

# Create zip with 7-Zip
Write-Host "  Creating zip file..." -ForegroundColor Gray
& "C:\Program Files\7-Zip\7z.exe" a -tzip app.zip .\deploy-temp\* -r | Out-Null
Remove-Item deploy-temp -Recurse -Force

$zipSize = (Get-Item app.zip).Length / 1MB
Write-Host "  Created app.zip ($([math]::Round($zipSize, 2)) MB)" -ForegroundColor Green

# Step 3: Disable build on Azure
Write-Host "`n[3/9] Configuring Azure to skip build..." -ForegroundColor Yellow
az webapp config appsettings set `
    --resource-group $resourceGroup `
    --name $appName `
    --settings SCM_DO_BUILD_DURING_DEPLOYMENT=false | Out-Null

Write-Host "  Build disabled on Azure" -ForegroundColor Green

# Step 4: Stop the app
Write-Host "`n[4/9] Stopping Azure App Service..." -ForegroundColor Yellow
az webapp stop --resource-group $resourceGroup --name $appName | Out-Null
Write-Host "  App stopped" -ForegroundColor Green

# Step 5: Clean up old files via Kudu API
Write-Host "`n[5/9] Cleaning up old deployment files..." -ForegroundColor Yellow

try {
    # Get publishing credentials
    Write-Host "  Getting deployment credentials..." -ForegroundColor Gray
    $credsJson = az webapp deployment list-publishing-credentials `
        --resource-group $resourceGroup `
        --name $appName `
        --query "{username:publishingUserName, password:publishingPassword}" `
        -o json
    
    $creds = $credsJson | ConvertFrom-Json
    $pair = "$($creds.username):$($creds.password)"
    $encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
    $headers = @{ Authorization = "Basic $encodedCreds" }
    
    $kuduUrl = "https://$appName.scm.azurewebsites.net"
    
    # Get list of items in wwwroot
    Write-Host "  Fetching current files..." -ForegroundColor Gray
    $items = Invoke-RestMethod -Uri "$kuduUrl/api/vfs/site/wwwroot/" -Headers $headers -Method GET -TimeoutSec 30
    
    # Delete each item except node_modules (will be recreated by npm install)
    $deletedCount = 0
    foreach ($item in $items) {
        if ($item.name -ne "node_modules") {
            $itemPath = "$kuduUrl/api/vfs/site/wwwroot/$($item.name)/"
            Write-Host "    Deleting: $($item.name)" -ForegroundColor Gray
            try {
                Invoke-RestMethod -Uri $itemPath -Headers $headers -Method DELETE -TimeoutSec 30 | Out-Null
                $deletedCount++
            } catch {
                Write-Host "    Warning: Could not delete $($item.name)" -ForegroundColor Yellow
            }
        }
    }
    
    Write-Host "  Deleted $deletedCount items" -ForegroundColor Green
    
} catch {
    Write-Host "  Warning: Could not clean up via API" -ForegroundColor Yellow
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "  Continuing with deployment anyway..." -ForegroundColor Gray
}

# Step 6: Deploy to Azure
Write-Host "`n[6/9] Deploying to Azure..." -ForegroundColor Yellow
Write-Host "  This may take a minute..." -ForegroundColor Gray

az webapp deploy `
    --resource-group $resourceGroup `
    --name $appName `
    --src-path app.zip `
    --type zip
    
if ($LASTEXITCODE -ne 0) {
    Write-Host "  Deployment failed!" -ForegroundColor Red
    exit 1
}

Write-Host "  Deployment successful" -ForegroundColor Green

# Step 7: Ensure npm install runs
Write-Host "`n[7/9] Ensuring dependencies will be installed..." -ForegroundColor Yellow
az webapp config appsettings set `
    --resource-group $resourceGroup `
    --name $appName `
    --settings SCM_DO_BUILD_DURING_DEPLOYMENT=true | Out-Null

Write-Host "  npm install will run on restart" -ForegroundColor Green

# Step 8: Set startup command
Write-Host "`n[8/9] Configuring startup command..." -ForegroundColor Yellow
az webapp config set `
    --resource-group $resourceGroup `
    --name $appName `
    --startup-file "node middletier.js" | Out-Null

Write-Host "  Startup command set to: node middletier.js" -ForegroundColor Green

# Step 9: Start the app
Write-Host "`n[9/9] Starting Azure App Service..." -ForegroundColor Yellow
az webapp start --resource-group $resourceGroup --name $appName | Out-Null
Write-Host "  App started" -ForegroundColor Green

# Wait for startup
Write-Host "`nWaiting for app to start..." -ForegroundColor Yellow
Start-Sleep -Seconds 15

# Health check with retry
Write-Host "`nPerforming health check..." -ForegroundColor Yellow
$healthCheckPassed = $false

for ($i = 1; $i -le 5; $i++) {
    Write-Host "  Attempt $i of 5..." -ForegroundColor Gray
    try {
        $response = Invoke-WebRequest -Uri $appUrl -TimeoutSec 10 -UseBasicParsing
        if ($response.StatusCode -eq 200) {
            $healthCheckPassed = $true
            Write-Host "  [OK] App is responding (200 OK)" -ForegroundColor Green
            break
        }
    } catch {
        if ($i -lt 5) {
            Write-Host "  Failed, waiting 10 seconds before retry..." -ForegroundColor Gray
            Start-Sleep -Seconds 10
        } else {
            Write-Host "  [FAIL] App not responding" -ForegroundColor Red
        }
    }
}

if (-not $healthCheckPassed) {
    Write-Host "`n[FAIL] App failed health check after 5 attempts" -ForegroundColor Red
    Write-Host "  This may be normal if the app takes longer to start" -ForegroundColor Yellow
    Write-Host "  Check logs for details" -ForegroundColor Yellow
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Deployment Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nApp URL: $appUrl" -ForegroundColor White
Write-Host "Kudu Console: https://$appName.scm.azurewebsites.net/" -ForegroundColor White

# Offer to view logs
Write-Host "`nWould you like to view live logs now? (Y/N): " -ForegroundColor Yellow -NoNewline
$userResponse = Read-Host

if ($userResponse -eq "Y" -or $userResponse -eq "y") {
    Write-Host "`nOpening log stream (Press Ctrl+C to exit)..." -ForegroundColor Cyan
    az webapp log tail --resource-group $resourceGroup --name $appName
} else {
    Write-Host "`nTo view logs later, run:" -ForegroundColor Cyan
    Write-Host "  az webapp log tail --resource-group $resourceGroup --name $appName" -ForegroundColor White
}