param(
  [Parameter(Mandatory = $true)][string]$ResourceGroup,
  [Parameter(Mandatory = $true)][string]$AppName
)

$ErrorActionPreference = 'Stop'

# Ensure correct app settings for run-from-package (no remote build)
$rfpSettings = @(
  'SCM_DO_BUILD_DURING_DEPLOYMENT=false',
  'WEBSITE_RUN_FROM_PACKAGE=1',
  'NODE_ENV=production',
  'WEBSITE_NODE_DEFAULT_VERSION=~24'
)
az webapp config appsettings set -g $ResourceGroup -n $AppName --settings $rfpSettings | Out-Null
# Set startup command to run Node directly (avoid npm start/Oryx)
az webapp config set -g $ResourceGroup -n $AppName --startup-file "node middletier.js" | Out-Null

# Move working directory to repo root (script is in scripts/)
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

# Clean previous package
$zipPath = Join-Path $repoRoot 'deploy.zip'
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

# Build production
npm run build:prod

# Verify dist exists
$dist = Join-Path $repoRoot 'dist'
if (!(Test-Path $dist)) {
  throw 'Build did not produce dist folder.'
}

# Copy package files and install production dependencies into dist for self-contained run-from-package
$pkg = Join-Path $repoRoot 'package.json'
$lock = Join-Path $repoRoot 'package-lock.json'
if (Test-Path $pkg) { Copy-Item $pkg $dist -Force }
if (Test-Path $lock) { Copy-Item $lock $dist -Force }

# Clean any previous node_modules in dist
$distNodeModules = Join-Path $dist 'node_modules'
if (Test-Path $distNodeModules) { Remove-Item $distNodeModules -Recurse -Force }

# Install production dependencies into dist
npm ci --omit=dev --prefix $dist

# Remove package files from dist to avoid remote build detection
$pkgInDist = Join-Path $dist 'package.json'
$lockInDist = Join-Path $dist 'package-lock.json'
if (Test-Path $pkgInDist) { Remove-Item $pkgInDist -Force }
if (Test-Path $lockInDist) { Remove-Item $lockInDist -Force }

# Optional: remove source maps to shrink package
Get-ChildItem $dist -Filter *.map -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue

# Create zip using .NET ZipFile for compatibility
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($dist, $zipPath, [System.IO.Compression.CompressionLevel]::Fastest, $false)

# Deploy the package (run-from-package)
az webapp deployment source config-zip -g $ResourceGroup -n $AppName --src $zipPath

Write-Host "Deployed run-from-package to $AppName in $ResourceGroup."