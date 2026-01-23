# Script for Azure AD Admin to create Service Principal
# This creates a service principal for CI/CD automation

param(
    [Parameter(Mandatory=$false)]
    [string]$ServicePrincipalName = "sp-outlook-addin-terraform",
    
    [Parameter(Mandatory=$false)]
    [string]$SubscriptionId
)

Write-Host "Creating Service Principal for Outlook Add-in CI/CD" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Gray

# Login check
$context = az account show 2>$null
if (-not $context) {
    Write-Host "Please login to Azure first..." -ForegroundColor Yellow
    az login
}

# Get subscription ID if not provided
if (-not $SubscriptionId) {
    $SubscriptionId = az account show --query id -o tsv
    Write-Host "Using current subscription: $SubscriptionId" -ForegroundColor Green
}

Write-Host "`nCreating service principal: $ServicePrincipalName" -ForegroundColor Cyan

# Create the service principal
$sp = az ad sp create-for-rbac `
    --name $ServicePrincipalName `
    --role Contributor `
    --scopes "/subscriptions/$SubscriptionId" `
    --years 2 | ConvertFrom-Json

if (-not $sp) {
    Write-Host "❌ Failed to create service principal" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Service principal created successfully!" -ForegroundColor Green

# Prepare the output
$output = @{
    clientId = $sp.appId
    clientSecret = $sp.password
    subscriptionId = $SubscriptionId
    tenantId = $sp.tenant
}

# Display the credentials
Write-Host "`n" -NoNewline
Write-Host "=" * 60 -ForegroundColor Green
Write-Host "SERVICE PRINCIPAL CREDENTIALS" -ForegroundColor Green
Write-Host "=" * 60 -ForegroundColor Green
Write-Host ""
Write-Host "Add these as GitHub Secrets:" -ForegroundColor Yellow
Write-Host ""
Write-Host "ARM_CLIENT_ID:       " -NoNewline -ForegroundColor Cyan
Write-Host $output.clientId -ForegroundColor White
Write-Host "ARM_CLIENT_SECRET:   " -NoNewline -ForegroundColor Cyan
Write-Host $output.clientSecret -ForegroundColor White
Write-Host "ARM_SUBSCRIPTION_ID: " -NoNewline -ForegroundColor Cyan
Write-Host $output.subscriptionId -ForegroundColor White
Write-Host "ARM_TENANT_ID:       " -NoNewline -ForegroundColor Cyan
Write-Host $output.tenantId -ForegroundColor White
Write-Host ""
Write-Host "=" * 60 -ForegroundColor Green

# Save to file
$outputFile = "service-principal-credentials.json"
$output | ConvertTo-Json | Out-File -FilePath $outputFile -Encoding UTF8

Write-Host "`n✅ Credentials saved to: $outputFile" -ForegroundColor Green
Write-Host "⚠️  IMPORTANT: Keep this file secure and delete after adding to GitHub!" -ForegroundColor Yellow

# Create instructions file
$instructions = @"
GitHub Secrets Setup Instructions
==================================

1. Go to your GitHub repository
2. Click Settings → Secrets and variables → Actions
3. Click "New repository secret"
4. Add each of these secrets:

Secret Name: ARM_CLIENT_ID
Value: $($output.clientId)

Secret Name: ARM_CLIENT_SECRET
Value: $($output.clientSecret)

Secret Name: ARM_SUBSCRIPTION_ID
Value: $($output.subscriptionId)

Secret Name: ARM_TENANT_ID
Value: $($output.tenantId)

5. After adding all secrets, DELETE this file and service-principal-credentials.json

Service Principal Details:
- Name: $ServicePrincipalName
- App ID: $($output.clientId)
- Subscription: $($output.subscriptionId)
- Role: Contributor
- Valid for: 2 years

To test the service principal:
az login --service-principal -u $($output.clientId) -p $($output.clientSecret) --tenant $($output.tenantId)
az account show
"@

$instructions | Out-File -FilePath "github-secrets-instructions.txt" -Encoding UTF8

Write-Host "`n📄 Instructions saved to: github-secrets-instructions.txt" -ForegroundColor Cyan
Write-Host "`n⚠️  Security Reminder:" -ForegroundColor Red
Write-Host "   1. Add these credentials to GitHub Secrets immediately" -ForegroundColor White
Write-Host "   2. Delete both credential files after setup" -ForegroundColor White
Write-Host "   3. Never commit these files to git" -ForegroundColor White