param(
  [Parameter(Mandatory = $true)][string]$ResourceGroup,
  [Parameter(Mandatory = $true)][string]$AppName
)

$ErrorActionPreference = 'Stop'

# Configure app settings for server-side build and runtime
$settings = @(
  'SCM_DO_BUILD_DURING_DEPLOYMENT=true',
  'WEBSITE_RUN_FROM_PACKAGE=0',
  'NODE_ENV=production',
  'WEBSITE_NODE_DEFAULT_VERSION=~24'
)
az webapp config appsettings set -g $ResourceGroup -n $AppName --settings $settings | Out-Null

# Ensure startup command uses npm start
az webapp config set -g $ResourceGroup -n $AppName --startup-file "npm start" | Out-Null

Write-Host "One-time setup applied to $AppName in $ResourceGroup."