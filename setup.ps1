<#
  One-click environment setup for E2E DMS.
  Safe to run any number of times: anything already installed is skipped.

  Installs (via winget, if missing): Node.js LTS, Google Chrome
  Then installs: npm packages (package.json) and the Playwright Chromium browser.
  Creates .env from .env.example and resources/credentials.json from
  resources/credentials.example.json if they do not exist (both are kept out of Git).

  Usage:  double-click setup.bat   or   powershell -ExecutionPolicy Bypass -File setup.ps1
#>
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

function Step($msg) { Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Ok($msg)   { Write-Host "    [OK] $msg" -ForegroundColor Green }

function Refresh-Path {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
}

function Ensure-WingetPackage($command, $wingetId, $label) {
    if (Get-Command $command -ErrorAction SilentlyContinue) { Ok "$label found"; return }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "$label is missing and winget is not available. Install $label manually and re-run."
    }
    Write-Host "    Installing $label ..."
    winget install --id $wingetId -e --silent --accept-package-agreements --accept-source-agreements
    Refresh-Path
    if (-not (Get-Command $command -ErrorAction SilentlyContinue)) {
        throw "$label installed but not on PATH yet. Close this window and run setup again."
    }
    Ok "$label installed"
}

# ---------- System software ----------
Step 'Checking Node.js'
Ensure-WingetPackage 'node' 'OpenJS.NodeJS.LTS' 'Node.js'

Step 'Checking Google Chrome (playwright.config.ts uses channel: chrome)'
$chromePaths = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)
if ($chromePaths | Where-Object { Test-Path $_ }) {
    Ok 'Google Chrome found'
} else {
    Write-Host '    Installing Google Chrome ...'
    winget install --id Google.Chrome -e --silent --accept-package-agreements --accept-source-agreements
    Ok 'Google Chrome installed'
}

# ---------- Project packages ----------
Step 'Installing npm packages'
if (Test-Path 'package-lock.json') { npm ci } else { npm install }
if ($LASTEXITCODE -ne 0) { throw 'npm install failed' }
Ok 'npm packages installed'

Step 'Installing Playwright browsers'
npx playwright install chromium
if ($LASTEXITCODE -ne 0) { throw 'playwright install failed' }
Ok 'Playwright browsers installed'

Step 'Caching Playwright MCP server (configured in .mcp.json)'
npx -y @playwright/mcp@latest --version
if ($LASTEXITCODE -ne 0) { throw 'Playwright MCP download failed' }
Ok 'Playwright MCP ready'

# ---------- Config ----------
Step 'Checking .env'
if (Test-Path '.env') {
    Ok '.env exists'
} else {
    Copy-Item '.env.example' '.env'
    Write-Host '    Created .env from .env.example - open it and fill in the values.' -ForegroundColor Yellow
}

Step 'Checking resources/credentials.json'
if (Test-Path 'resources/credentials.json') {
    Ok 'credentials.json exists'
} else {
    Copy-Item 'resources/credentials.example.json' 'resources/credentials.json'
    Write-Host '    Created resources/credentials.json from the example - fill in the DMS logins.' -ForegroundColor Yellow
}

Write-Host "`nSetup complete. Run tests with:  npm test" -ForegroundColor Green
