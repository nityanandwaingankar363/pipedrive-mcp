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

# Runs a native command (git, uv, winget, etc.) without triggering PowerShell's
# "NativeCommandError" behaviour on stderr writes. Checks $LASTEXITCODE and
# throws on non-zero. Pass the command as a script block.
function Invoke-NativeCommand {
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $ScriptBlock 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "$Description failed (exit code $LASTEXITCODE). See output above for details."
        }
    } finally {
        $ErrorActionPreference = $prev
    }
}

# ---------- Banner ------------------------------------------------

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Pipedrive MCP Installer for Claude Desktop"       -ForegroundColor Cyan
Write-Host "  Aventis Advisors"                                 -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

Assert-Admin

# ---------- Check: Claude Desktop must be fully quit --------------

$claudeProc = Get-Process -Name "Claude" -ErrorAction SilentlyContinue
if ($claudeProc) {
    Write-Host ""
    Write-Warn "Claude Desktop is currently running."
    Write-Warn "It must be fully quit before we continue -- otherwise file locks will"
    Write-Warn "block the installation."
    Write-Host ""
    Write-Info "How to fully quit Claude Desktop:"
    Write-Info "  1. Look at the system tray (bottom-right of screen, near the clock)."
    Write-Info "     Click the small up-arrow (^) if you don't see the Claude icon."
    Write-Info "  2. Right-click the Claude icon."
    Write-Info "  3. Click 'Quit' (not just closing the window)."
    Write-Host ""
    $answer = Read-Host "Press Enter when Claude Desktop is fully closed (or type 'cancel' to abort)"
    if ($answer -eq 'cancel') {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 0
    }
    Start-Sleep -Seconds 1
    if (Get-Process -Name "Claude" -ErrorAction SilentlyContinue) {
        Write-Warn "Claude Desktop still appears to be running. Proceeding anyway, but"
        Write-Warn "you may encounter file-lock errors below. If so, fully quit Claude"
        Write-Warn "Desktop and re-run this script."
    } else {
        Write-Ok "Claude Desktop is fully quit."
    }
}

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
    Invoke-NativeCommand "winget install Python" {
        winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    }
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
    Invoke-NativeCommand "winget install Git" {
        winget install -e --id Git.Git --silent --accept-source-agreements --accept-package-agreements
    }
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

# If the folder exists but isn't a real git repo (e.g. half-deleted from a
# previous failed run), remove it so we can re-clone from scratch.
if ((Test-Path $InstallPath) -and (-not (Test-Path (Join-Path $InstallPath ".git")))) {
    Write-Warn "Folder exists at $InstallPath but is not a git repo (corrupt from a previous run). Removing..."
    try {
        Remove-Item -Recurse -Force $InstallPath -ErrorAction Stop
    } catch {
        Write-Host "Could not remove $InstallPath -- files may still be locked by Claude Desktop." -ForegroundColor Red
        Write-Host "Fully quit Claude Desktop and re-run this script." -ForegroundColor Red
        throw
    }
}

if (-not (Test-Path $InstallPath)) {
    Write-Info "Cloning to $InstallPath..."
    Invoke-NativeCommand "git clone" { git clone --quiet $RepoUrl $InstallPath }
    Write-Ok "Cloned to $InstallPath"
} else {
    Write-Info "Folder already exists at $InstallPath -- pulling latest changes..."
    $pullFailed = $false
    Push-Location $InstallPath
    try {
        Invoke-NativeCommand "git pull" { git pull origin main --quiet }
    } catch {
        $pullFailed = $true
        Write-Warn "git pull failed. Will try a fresh clone as recovery."
    } finally {
        Pop-Location
    }

    if ($pullFailed) {
        try {
            Remove-Item -Recurse -Force $InstallPath -ErrorAction Stop
            Invoke-NativeCommand "git clone" { git clone --quiet $RepoUrl $InstallPath }
            Write-Ok "Re-cloned to $InstallPath"
        } catch {
            Write-Host "Could not recover. The local folder is still in a bad state." -ForegroundColor Red
            Write-Host "Fully quit Claude Desktop, manually delete $InstallPath, then re-run this script." -ForegroundColor Red
            throw
        }
    } else {
        Write-Ok "Updated $InstallPath"
    }
}

# ---------- Step 5: Python deps ----------------------------------

Write-Step "Installing Python dependencies (this takes 1-2 minutes)"
Push-Location $InstallPath
try {
    Invoke-NativeCommand "uv sync" { uv sync }
} finally {
    Pop-Location
}
Write-Ok "Dependencies installed"

# ---------- Step 6: Prompt for API token -------------------------

Write-Step "Pipedrive API token"
Write-Info "You need a personal Pipedrive API token. Each teammate uses their own --"
Write-Info "do NOT share tokens, and do NOT reuse someone else's."
Write-Host ""
Write-Info "How to get your token:"
Write-Info "  1. In your web browser, open:  https://aventis-advisors.pipedrive.com"
Write-Info "  2. Sign in to Pipedrive."
Write-Info "  3. Click your profile picture in the top-right corner."
Write-Info "  4. In the dropdown, click 'Personal preferences'."
Write-Info "  5. Click the 'API' tab (may be under a 'More' menu on narrow windows)."
Write-Info "  6. If an existing token is shown, click 'Revoke' first, then 'Generate new token'."
Write-Info "     Otherwise, just click 'Generate new token'."
Write-Info "  7. Copy the long string of letters and numbers it shows you."
Write-Host ""
Write-Warn "SECURITY: treat this token like a password. When you paste it below,"
Write-Warn "nothing will appear on your screen -- that is intentional, not a bug."
Write-Warn "Just paste (Ctrl+V or right-click) and press Enter."
Write-Host ""
$secureToken = Read-Host "Paste your Pipedrive API token here" -AsSecureString
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
# Write as UTF-8 WITHOUT BOM (the BOM breaks some parsers / tools).
[System.IO.File]::WriteAllText($envPath, $envContent, [System.Text.UTF8Encoding]::new($false))
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
        # Use .NET ReadAllText: handles UTF-8 BOM correctly (strips it).
        $raw = [System.IO.File]::ReadAllText($ClaudeConfigPath)
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

$jsonOut = $config | ConvertTo-Json -Depth 10
# Write as UTF-8 WITHOUT BOM. Claude Desktop's parser rejects files that start
# with the UTF-8 BOM (0xEF 0xBB 0xBF), producing "Unexpected token" errors.
[System.IO.File]::WriteAllText($ClaudeConfigPath, $jsonOut, [System.Text.UTF8Encoding]::new($false))
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
