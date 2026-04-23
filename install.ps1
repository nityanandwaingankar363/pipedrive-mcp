#Requires -Version 5.1
# ===================================================================
#  Pipedrive MCP Installer for Claude Desktop — Aventis Advisors
# ===================================================================
# USAGE:
#   1. Right-click the PowerShell icon, choose "Run as administrator".
#   2. cd to the folder where this script lives (e.g. Downloads).
#   3. Run:  .\install.ps1
# -------------------------------------------------------------------

$ErrorActionPreference = "Stop"

# ---------- Helpers -----------------------------------------------

function Write-Step($msg) { Write-Host ""; Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok  ($msg) { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Write-Info($msg) { Write-Host "  $msg" -ForegroundColor White }
function Write-Warn($msg) { Write-Host "  [!]  $msg" -ForegroundColor Yellow }

function Test-Cmd($name) {
    try { Get-Command $name -ErrorAction Stop | Out-Null; return $true } catch { return $false }
}

function Update-SessionPath {
    # Merge machine + user PATH, plus uv's default install dir
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [Environment]::GetEnvironmentVariable("Path", "User") + ";" +
                "$env:USERPROFILE\.local\bin"
}

function ConvertTo-HashtableRecursive($Object) {
    if ($null -eq $Object) { return @{} }
    if ($Object -is [hashtable]) {
        $new = @{}
        foreach ($k in $Object.Keys) { $new[$k] = ConvertTo-HashtableRecursive $Object[$k] }
        return $new
    }
    if ($Object -is [PSCustomObject]) {
        $hash = @{}
        foreach ($p in $Object.PSObject.Properties) {
            $hash[$p.Name] = ConvertTo-HashtableRecursive $p.Value
        }
        return $hash
    }
    if ($Object -is [array]) {
        return ,@($Object | ForEach-Object { ConvertTo-HashtableRecursive $_ })
    }
    return $Object
}

function Assert-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host ""
        Write-Host "ERROR: This script must be run as administrator." -ForegroundColor Red
        Write-Host "Close this window, right-click PowerShell, choose 'Run as administrator'," -ForegroundColor Red
        Write-Host "then re-run this script." -ForegroundColor Red
        exit 1
    }
}

# ---------- Banner ------------------------------------------------

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Pipedrive MCP Installer for Claude Desktop"       -ForegroundColor Cyan
Write-Host "  Aventis Advisors"                                 -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

Assert-Admin

# ---------- Configuration ----------------------------------------

$RepoUrl          = "https://github.com/nityanandwaingankar363/pipedrive-mcp.git"
$InstallPath      = "$env:USERPROFILE\pipedrive-mcp"
$CompanyDomain    = "aventis-advisors"
$ClaudeConfigPath = "$env:APPDATA\Claude\claude_desktop_config.json"

# ---------- Step 1: Python 3.12+ ---------------------------------

Write-Step "Checking Python"
$pythonOk = $false
if (Test-Cmd python) {
    $ver = (python --version 2>&1) -replace '^Python ', ''
    $parts = $ver.Split('.')
    if ($parts.Count -ge 2) {
        $major = [int]$parts[0]; $minor = [int]$parts[1]
        if ($major -eq 3 -and $minor -ge 12) {
            Write-Ok "Python $ver already installed"
            $pythonOk = $true
        } else {
            Write-Warn "Python $ver is too old (need 3.12+). Will install 3.12."
        }
    }
}
if (-not $pythonOk) {
    Write-Info "Installing Python 3.12 via winget..."
    winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    Update-SessionPath
    if (-not (Test-Cmd python)) {
        Write-Host "Python installed but not yet on PATH in this window." -ForegroundColor Red
        Write-Host "Close this PowerShell window, reopen as administrator, and re-run install.ps1." -ForegroundColor Red
        exit 1
    }
    Write-Ok "Python installed"
}

# ---------- Step 2: Git ------------------------------------------

Write-Step "Checking Git"
if (Test-Cmd git) {
    Write-Ok "Git already installed"
} else {
    Write-Info "Installing Git via winget..."
    winget install -e --id Git.Git --silent --accept-source-agreements --accept-package-agreements
    Update-SessionPath
    if (-not (Test-Cmd git)) {
        Write-Host "Git installed but not yet on PATH. Restart PowerShell as admin and re-run." -ForegroundColor Red
        exit 1
    }
    Write-Ok "Git installed"
}

# ---------- Step 3: uv -------------------------------------------

Write-Step "Checking uv"
if (Test-Cmd uv) {
    Write-Ok "uv already installed"
} else {
    Write-Info "Installing uv..."
    & powershell -ExecutionPolicy ByPass -Command "irm https://astral.sh/uv/install.ps1 | iex" | Out-Null
    Update-SessionPath
    if (-not (Test-Cmd uv)) {
        Write-Host "uv installed but not yet on PATH. Restart PowerShell as admin and re-run." -ForegroundColor Red
        exit 1
    }
    Write-Ok "uv installed"
}

# ---------- Step 4: Clone / update repo --------------------------

Write-Step "Getting the Pipedrive MCP code"
if (Test-Path $InstallPath) {
    Write-Info "Folder already exists at $InstallPath -- pulling latest changes..."
    Push-Location $InstallPath
    git pull origin main 2>&1 | Out-Null
    Pop-Location
    Write-Ok "Updated $InstallPath"
} else {
    Write-Info "Cloning to $InstallPath..."
    git clone $RepoUrl $InstallPath 2>&1 | Out-Null
    Write-Ok "Cloned to $InstallPath"
}

# ---------- Step 5: Python deps ----------------------------------

Write-Step "Installing Python dependencies (this takes 1-2 minutes)"
Push-Location $InstallPath
uv sync 2>&1 | Out-Null
Pop-Location
Write-Ok "Dependencies installed"

# ---------- Step 6: Prompt for API token -------------------------

Write-Step "Pipedrive API token"
Write-Info "Get your personal API token (each teammate needs their own):"
Write-Info "  1. Sign in at aventis-advisors.pipedrive.com"
Write-Info "  2. Profile picture (top-right) -> Personal preferences"
Write-Info "  3. API tab -> Generate new token -> copy the value"
Write-Host ""
$secureToken = Read-Host "Paste your Pipedrive API token" -AsSecureString
$token = [System.Net.NetworkCredential]::new("", $secureToken).Password
if ([string]::IsNullOrWhiteSpace($token)) {
    Write-Host "ERROR: No token entered. Aborting." -ForegroundColor Red
    exit 1
}
Write-Ok "Token captured"

# ---------- Step 7: Write .env -----------------------------------

Write-Step "Writing .env configuration"
$envContent = @"
# Generated by install.ps1

HOST=127.0.0.1
PORT=8152
TRANSPORT=stdio
CONTAINER_MODE=false

PIPEDRIVE_API_TOKEN=$token
PIPEDRIVE_COMPANY_DOMAIN=$CompanyDomain

PIPEDRIVE_BASE_URL=https://api.pipedrive.com/v2
PIPEDRIVE_TIMEOUT=30
PIPEDRIVE_RETRY_ATTEMPTS=3
PIPEDRIVE_RETRY_BACKOFF=0.5
VERIFY_SSL=true
PIPEDRIVE_LOG_REQUESTS=false
PIPEDRIVE_LOG_RESPONSES=false

# Read-only mode: set to true to hide create/update/delete tools from Claude.
PIPEDRIVE_READ_ONLY=false

PIPEDRIVE_FEATURE_PERSONS=true
PIPEDRIVE_FEATURE_ORGANIZATIONS=true
PIPEDRIVE_FEATURE_DEALS=true
PIPEDRIVE_FEATURE_LEADS=true
PIPEDRIVE_FEATURE_ITEM_SEARCH=true
PIPEDRIVE_FEATURE_ACTIVITIES=true
FEATURE_CONFIG_PATH=

LOG_LEVEL=INFO
"@
$envPath = Join-Path $InstallPath ".env"
$envContent | Out-File -FilePath $envPath -Encoding UTF8
Write-Ok ".env written at $envPath"

# ---------- Step 8: Merge Claude Desktop config ------------------

Write-Step "Updating Claude Desktop config"
$configDir = Split-Path $ClaudeConfigPath
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    Write-Info "Created $configDir"
}

$config = $null
if (Test-Path $ClaudeConfigPath) {
    try {
        $raw = Get-Content $ClaudeConfigPath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            $config = @{}
            Write-Info "Existing config was empty -- starting fresh"
        } else {
            $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
            $config = ConvertTo-HashtableRecursive $parsed
            Write-Info "Existing config loaded (other MCP servers preserved)"
        }
    } catch {
        $backupPath = "$ClaudeConfigPath.broken"
        Write-Warn "Existing config is not valid JSON. Backing up to:"
        Write-Warn "  $backupPath"
        Copy-Item $ClaudeConfigPath $backupPath -Force
        $config = @{}
    }
} else {
    $config = @{}
    Write-Info "No existing config found -- creating a new one"
}

if (-not $config.ContainsKey("mcpServers") -or $null -eq $config.mcpServers) {
    $config.mcpServers = @{}
}
if ($config.mcpServers -isnot [hashtable]) {
    $config.mcpServers = ConvertTo-HashtableRecursive $config.mcpServers
}

$config.mcpServers.pipedrive = @{
    command = "uv"
    args    = @("--directory", $InstallPath, "run", "server.py")
}

$config | ConvertTo-Json -Depth 10 | Set-Content -Path $ClaudeConfigPath -Encoding UTF8
Write-Ok "Pipedrive entry written to $ClaudeConfigPath"

# ---------- Done -------------------------------------------------

Write-Host ""
Write-Host "==================================================" -ForegroundColor Green
Write-Host "  Installation complete!"                           -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Fully quit Claude Desktop (system tray -> right-click Claude icon -> Quit)."
Write-Host "  2. Reopen Claude Desktop."
Write-Host "  3. In a new regular chat, try this:"
Write-Host ""
Write-Host "       Using Pipedrive, show me my 5 most recent deals." -ForegroundColor White
Write-Host ""
Write-Host "If anything goes wrong, contact Nitya:"
Write-Host "  nityanand.waingankar@aventis-advisors.com"
Write-Host ""
