param(
  [Parameter(Mandatory = $true)][string]$ResourceGroup,
  [Parameter(Mandatory = $true)][string]$AppName,
  [string]$EnvFile = '.env.production',
  [switch]$SkipAppSettings
)

$ErrorActionPreference = 'Stop'

function Get-EnvSettingsFromFile {
  param([Parameter(Mandatory=$true)][string]$Path)

  if (!(Test-Path $Path)) { throw "Env file not found: $Path" }

  $map = @{}
  $lines = Get-Content -Path $Path -Encoding UTF8
  foreach ($raw in $lines) {
    $line = $raw.Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    if ($line.StartsWith('#')) { continue }

    # Support leading 'export '
    if ($line.StartsWith('export ')) { $line = $line.Substring(7).Trim() }

    # Match KEY=VALUE
    if ($line -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$') {
      $key = $Matches[1]
      $val = $Matches[2]

      # Remove inline comments only if preceded by space (do not break URLs)
      # Disabled by default to avoid stripping secrets containing '#'
      # if ($val -match '^(.*?)(\s+#.*)$') { $val = $Matches[1] }

      $val = $val.Trim()
      # Strip surrounding quotes
      if ($val.Length -ge 2 -and $val.StartsWith('"') -and $val.EndsWith('"')) {
        $val = $val.Substring(1, $val.Length - 2)
      } elseif ($val.Length -ge 2 -and $val.StartsWith("'") -and $val.EndsWith("'")) {
        $val = $val.Substring(1, $val.Length - 2)
      }

      # Store; last definition wins
      $map[$key] = $val
    }
  }

  # Convert to KEY=VALUE string array
  $out = @()
  foreach ($k in $map.Keys) { $out += ("{0}={1}" -f $k, $map[$k]) }
  return ,$out
}

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

# Copy package files for Kudu to restore dependencies and run npm start
$pkg = Join-Path $repoRoot 'package.json'
$lock = Join-Path $repoRoot 'package-lock.json'
if (Test-Path $pkg) { Copy-Item $pkg $dist -ErrorAction SilentlyContinue }
if (Test-Path $lock) { Copy-Item $lock $dist -ErrorAction SilentlyContinue }

# Optional: remove source maps to shrink package
Get-ChildItem $dist -Filter *.map -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue

# Create zip using .NET ZipFile for compatibility
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($dist, $zipPath, [System.IO.Compression.CompressionLevel]::Fastest, $false)

# Configure App Settings (Azure environment variables) from .env.production
if (-not $SkipAppSettings) {
  $settings = Get-EnvSettingsFromFile -Path (Join-Path $repoRoot $EnvFile)

  # Ensure NODE_ENV present; default to production if missing
  if (-not ($settings -match '^NODE_ENV=')) { $settings += 'NODE_ENV=production' }

  # Configure for Oryx build on server (install node_modules) and run from extracted files
  $settings = $settings | Where-Object { $_ -notmatch '^WEBSITE_RUN_FROM_PACKAGE=' -and $_ -notmatch '^SCM_DO_BUILD_DURING_DEPLOYMENT=' }
  $settings += 'WEBSITE_RUN_FROM_PACKAGE=0'
  $settings += 'SCM_DO_BUILD_DURING_DEPLOYMENT=true'

  az webapp config appsettings set -g $ResourceGroup -n $AppName --settings $settings | Out-Null

  # Ensure startup command is npm start (reads package.json -> node dist/middletier.js)
  az webapp config set -g $ResourceGroup -n $AppName --startup-file "npm start" | Out-Null
}

# Deploy using config-zip (Oryx will build and start app)
az webapp deployment source config-zip -g $ResourceGroup -n $AppName --src $zipPath

# Optional health check
try {
  $healthUrl = "https://$AppName.azurewebsites.net/health"
  $resp = Invoke-WebRequest -UseBasicParsing -Uri $healthUrl -TimeoutSec 30
  Write-Host "Health $($resp.StatusCode): $healthUrl"
} catch {
  Write-Host "Health check failed or endpoint not reachable."
}
