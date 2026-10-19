<#
.SYNOPSIS
    Installs the Teams MCP server for use with Claude Code.

.DESCRIPTION
    - Checks for Python 3.10+ and installs via winget if missing
    - Creates an isolated venv at %LOCALAPPDATA%\mcp-teams-server
    - Installs the mcp-teams-server wheel and dependencies
    - Registers the MCP server with Claude Code (claude mcp add)
    - First use will trigger a one-time device-code authentication in the browser

.NOTES
    Run from the repository root (where pyproject.toml lives).
    Requires: winget (for auto-install of Python), Claude Code CLI (claude)

.EXAMPLE
    .\install-teams-mcp.ps1 -TeamsAppId <app-guid> -TeamsTenantId <tenant-guid>
#>

param(
    # Your Azure Entra app registration. Pass them here, or set TEAMS_APP_ID and
    # TEAMS_APP_TENANT_ID in the environment. See doc/MS-Teams-setup.md.
    [string]$TeamsAppId = $env:TEAMS_APP_ID,
    [string]$TeamsTenantId = $env:TEAMS_APP_TENANT_ID,
    [string]$WheelPath = "",
    [string]$PythonPath = "",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$InstallDir = "$env:LOCALAPPDATA\mcp-teams-server"
$VenvDir = "$InstallDir\.venv"
$PythonExe = "$VenvDir\Scripts\python.exe"
$MinPythonMajor = 3
$MinPythonMinor = 10
$MaxPythonMinor = 12

# Resolve script directory (fallback to current dir if PSScriptRoot is empty)
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
} else {
    $scriptDir = (Get-Location).Path
}

# --- Azure Entra App Registration ---
# Supplied by -TeamsAppId / -TeamsTenantId or the matching env vars. No IDs are
# baked into this script: a tenant ID is a durable pointer to one organisation.
if (-not $TeamsAppId) {
    throw "TeamsAppId is required. Pass -TeamsAppId <guid> or set `$env:TEAMS_APP_ID. See doc/MS-Teams-setup.md."
}
if (-not $TeamsTenantId) {
    throw "TeamsTenantId is required. Pass -TeamsTenantId <guid> or set `$env:TEAMS_APP_TENANT_ID."
}
$TeamsAppType     = "SingleTenant"
$TeamsAuthMode    = "delegated"

# ---------------------------------------------------------------
# Helper: Write colored status messages
# ---------------------------------------------------------------
function Write-Step([string]$msg) { Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-Ok([string]$msg)   { Write-Host "   OK: $msg" -ForegroundColor Green }
function Write-Warn([string]$msg) { Write-Host "   WARN: $msg" -ForegroundColor Yellow }
function Write-Err([string]$msg)  { Write-Host "   ERROR: $msg" -ForegroundColor Red }

# Helper: Run a native command without ErrorActionPreference=Stop killing
# the script when the command writes to stderr.
function Invoke-Native {
    param([scriptblock]$Command)
    $prevPref = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $Command 2>&1 | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) {
                Write-Host "   $_" -ForegroundColor DarkGray
            } else {
                Write-Host "   $_"
            }
        }
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prevPref
    }
}

# ---------------------------------------------------------------
# Step 0: Check Windows Long Path support
# ---------------------------------------------------------------
Write-Step "Checking Windows long path support..."

$longPathEnabled = $false
try {
    $val = Get-ItemPropertyValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -ErrorAction Stop
    if ($val -eq 1) { $longPathEnabled = $true }
} catch {
    # Key doesn't exist or access denied - treat as disabled
}

if (-not $longPathEnabled) {
    Write-Warn "Windows long path support is NOT enabled."
    Write-Host "   The msgraph SDK has deeply nested paths that exceed the 260-char limit."
    Write-Host "   Attempting to enable it now (requires admin)..."

    # Try to enable it - will fail silently if not admin
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -Type DWord -ErrorAction Stop
        Write-Ok "Long paths enabled successfully."
        Write-Warn "You must CLOSE this PowerShell window and open a new one for the change to take effect."
        Write-Host "   Then re-run this script."
        exit 0
    } catch {
        Write-Err "Could not enable long paths (admin required)."
        Write-Host ""
        Write-Host "   Run this command in an ADMIN PowerShell:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host '   Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -Type DWord' -ForegroundColor White
        Write-Host ""
        Write-Host "   Then CLOSE and reopen PowerShell, and re-run this script."
        exit 1
    }
}
else {
    Write-Ok "Long paths enabled."
}

# ---------------------------------------------------------------
# Step 1: Check / Install Python
# ---------------------------------------------------------------
Write-Step "Checking Python installation..."

$pythonCmd = $null

if ($PythonPath -and (Test-Path $PythonPath)) {
    $pythonCmd = $PythonPath
    Write-Ok "Using specified Python: $pythonCmd"
} else {
    $candidateList = @()

    # Prioritize 64-bit Python 3.12 (known compatible binary wheels on Windows)
    if (Test-Path "C:\Program Files\Python312\python.exe") {
        $candidateList += "C:\Program Files\Python312\python.exe"
    }
    try {
        $pyExe = (& py -3.12 -c "import sys; print(sys.executable)" 2>$null).Trim()
        if ($pyExe -and (Test-Path $pyExe) -and ($candidateList -notcontains $pyExe)) {
            $candidateList += $pyExe
        }
    } catch {}

    $candidateList += @("python", "python3", "py")

    # Check locations in order of preference
    foreach ($candidate in $candidateList) {
        try {
            $ver = & $candidate --version 2>&1
            if ($ver -match "Python (\d+)\.(\d+)") {
                $major = [int]$Matches[1]
                $minor = [int]$Matches[2]
                if ($major -eq $MinPythonMajor -and $minor -ge $MinPythonMinor -and $minor -le $MaxPythonMinor) {
                    $pythonCmd = $candidate
                    Write-Ok "Found $ver ($candidate)"
                    break
                } else {
                    Write-Warn "$candidate is version $ver (need $MinPythonMajor.$MinPythonMinor - $MinPythonMajor.$MaxPythonMinor)"
                }
            }
        } catch {
            # candidate not found, continue
        }
    }
}

if (-not $pythonCmd) {
    Write-Warn "Python $MinPythonMajor.$MinPythonMinor+ not found on PATH."
    Write-Step "Installing Python via winget..."

    # Check if winget is available
    try {
        $null = Get-Command winget -ErrorAction Stop
    } catch {
        Write-Err "winget is not available. Please install Python $MinPythonMajor.$MinPythonMinor+ manually from https://www.python.org/downloads/"
        Write-Host "After installing, re-run this script."
        exit 1
    }

    # Install latest Python 3 via winget
    Write-Host "   Installing Python 3.12 via winget (this may take a minute)..."
    $exitCode = Invoke-Native { winget install --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements --silent }

    if ($exitCode -ne 0) {
        Write-Err "winget install failed (exit code $exitCode)."
        Write-Host "   Please install Python manually from https://www.python.org/downloads/"
        exit 1
    }

    # Refresh PATH so we can find the newly installed python
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

    # Re-check
    foreach ($candidate in @("python", "python3", "py")) {
        try {
            $ver = & $candidate --version 2>&1
            if ($ver -match "Python (\d+)\.(\d+)") {
                $major = [int]$Matches[1]
                $minor = [int]$Matches[2]
                if ($major -gt $MinPythonMajor -or ($major -eq $MinPythonMajor -and $minor -ge $MinPythonMinor)) {
                    $pythonCmd = $candidate
                    Write-Ok "Installed $ver ($candidate)"
                    break
                }
            }
        } catch {}
    }

    if (-not $pythonCmd) {
        Write-Err "Python still not found after install. You may need to restart your terminal or add Python to PATH manually."
        exit 1
    }
}

# ---------------------------------------------------------------
# Step 2: Find or build the wheel
# ---------------------------------------------------------------
Write-Step "Locating wheel package..."

if ($WheelPath -and (Test-Path $WheelPath)) {
    # User provided explicit path
} else {
    # Look in dist/ relative to this script, then current directory
    $candidates = @(
        (Join-Path $scriptDir "dist\mcp_teams_server-*.whl"),
        (Join-Path $scriptDir "mcp_teams_server-*.whl"),
        "dist\mcp_teams_server-*.whl",
        "mcp_teams_server-*.whl"
    )
    foreach ($pattern in $candidates) {
        $found = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) {
            $WheelPath = $found.FullName
            break
        }
    }
}

# If no wheel found, build it from source
if (-not $WheelPath -or -not (Test-Path $WheelPath)) {
    $pyprojectPath = Join-Path $scriptDir "pyproject.toml"
    if (Test-Path $pyprojectPath) {
        Write-Warn "No wheel found. Building from source..."
        Invoke-Native { & $pythonCmd -m pip install --upgrade build --quiet } | Out-Null
        Push-Location $scriptDir
        $buildExit = Invoke-Native { & $pythonCmd -m build --wheel }
        Pop-Location

        if ($buildExit -ne 0) {
            Write-Err "Wheel build failed."
            exit 1
        }

        $found = Get-ChildItem -Path (Join-Path $scriptDir "dist\mcp_teams_server-*.whl") -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) {
            $WheelPath = $found.FullName
        }
    }
}

if (-not $WheelPath -or -not (Test-Path $WheelPath)) {
    Write-Err "Could not find or build mcp_teams_server wheel (.whl) file."
    Write-Host "   Run this script from the teams-MCP-local directory, or pass -WheelPath <path>"
    exit 1
}

Write-Ok "Using wheel: $WheelPath"

# ---------------------------------------------------------------
# Step 3: Create venv and install
# ---------------------------------------------------------------
Write-Step "Creating virtual environment at $VenvDir..."

# Helper: Remove a directory even when it contains paths > 260 chars.
# Uses robocopy to mirror an empty folder (which purges all contents),
# then removes the now-empty directory normally.
function Remove-LongPathDir([string]$dir) {
    if (-not (Test-Path $dir)) { return }
    $tempRoot = [System.IO.Path]::GetTempPath()
    $emptyDir = Join-Path $tempRoot "mcp-empty-$(Get-Random)"
    New-Item -ItemType Directory -Path $emptyDir -Force | Out-Null
    try {
        robocopy $emptyDir $dir /purge /s /njh /njs /ndl /nc /ns /nfl | Out-Null
        Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
    } finally {
        try { Remove-Item $emptyDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}

if ((Test-Path $VenvDir) -and -not $Force) {
    Write-Warn "Venv already exists. Use -Force to recreate."
    $response = Read-Host "   Reinstall into existing venv? (y/N)"
    if ($response -ne "y") {
        Write-Host "   Skipping venv creation, will upgrade package."
    } else {
        Remove-LongPathDir $VenvDir
        & $pythonCmd -m venv $VenvDir
    }
} elseif ((Test-Path $VenvDir) -and $Force) {
    Remove-LongPathDir $VenvDir
    & $pythonCmd -m venv $VenvDir
} else {
    & $pythonCmd -m venv $VenvDir
}

if (-not (Test-Path $PythonExe)) {
    Write-Err "Failed to create virtual environment."
    exit 1
}

Write-Ok "Virtual environment ready."

Write-Step "Installing mcp-teams-server and dependencies..."
Invoke-Native { & $PythonExe -m pip install --upgrade pip --quiet } | Out-Null
$pipExit = Invoke-Native { & $PythonExe -m pip install --upgrade --force-reinstall $WheelPath "mcp<2" "mcp[cli]<2" }

if ($pipExit -ne 0) {
    Write-Err "pip install failed."
    exit 1
}

Write-Ok "Package installed."

# ---------------------------------------------------------------
# Step 4: Verify import
# ---------------------------------------------------------------
Write-Step "Verifying installation..."

$verifyExit = Invoke-Native { & $PythonExe -c "import mcp_teams_server; print('mcp_teams_server loaded OK')" }

if ($verifyExit -ne 0) {
    Write-Err "Verification failed - package did not import cleanly."
    exit 1
}

Write-Ok "All imports verified."

# ---------------------------------------------------------------
# Step 5: Register with Claude Code
# ---------------------------------------------------------------
Write-Step "Registering MCP server with Claude Code..."

# Check if claude CLI is available
$claudeAvailable = $false
try {
    $null = Get-Command claude -ErrorAction Stop
    $claudeAvailable = $true
} catch {
    # claude not found
}

if (-not $claudeAvailable) {
    Write-Warn "Claude Code CLI not found on PATH."
    Write-Host "   After installing Claude Code, run:"
    Write-Host ""
    Write-Host "   claude mcp add teams-mcp -e TEAMS_APP_ID=$TeamsAppId -e TEAMS_APP_TENANT_ID=$TeamsTenantId -e TEAMS_APP_TYPE=$TeamsAppType -e TEAMS_AUTH_MODE=$TeamsAuthMode -- `"$PythonExe`" -m mcp_teams_server" -ForegroundColor White
    Write-Host ""
    exit 0
}

# Remove existing entry if present (ignore errors)
Invoke-Native { claude mcp remove teams-mcp --scope user } | Out-Null

# Add the MCP server with env vars at user scope (available globally, not just this project)
Invoke-Native { claude mcp add teams-mcp --scope user -e "TEAMS_APP_ID=$TeamsAppId" -e "TEAMS_APP_TENANT_ID=$TeamsTenantId" -e "TEAMS_APP_TYPE=$TeamsAppType" -e "TEAMS_AUTH_MODE=$TeamsAuthMode" -- "$PythonExe" -m mcp_teams_server }

Write-Ok "MCP server registered."

# ---------------------------------------------------------------
# Step 6: Authenticate with Microsoft (one-time device-code flow)
# ---------------------------------------------------------------
Write-Step "Authenticating with Microsoft..."
Write-Host "   A browser window will open and a dialog will show the login code."
Write-Host "   Go to https://microsoft.com/devicelogin and enter the code."
Write-Host ""

$authScript = "import os; os.environ.update({'TEAMS_APP_ID':'$TeamsAppId','TEAMS_APP_TENANT_ID':'$TeamsTenantId','TEAMS_APP_TYPE':'$TeamsAppType','TEAMS_AUTH_MODE':'$TeamsAuthMode'}); from mcp_teams_server.auth import create_credential, DELEGATED_SCOPES; from mcp_teams_server.config import BotConfiguration; config = BotConfiguration(); cred = create_credential(config); cred.get_token(*DELEGATED_SCOPES, enable_cae=True); print('Authentication successful - token cached.')"
$authExit = Invoke-Native { & $PythonExe -c $authScript }

if ($authExit -ne 0) {
    Write-Warn "Authentication was skipped or failed."
    Write-Host "   You can authenticate later by running:"
    Write-Host ""
    Write-Host "   & `"$PythonExe`" -c `"import os; os.environ.update({'TEAMS_APP_ID':'$TeamsAppId','TEAMS_APP_TENANT_ID':'$TeamsTenantId','TEAMS_APP_TYPE':'$TeamsAppType','TEAMS_AUTH_MODE':'$TeamsAuthMode'}); from mcp_teams_server.auth import create_credential, DELEGATED_SCOPES; from mcp_teams_server.config import BotConfiguration; cred = create_credential(BotConfiguration()); cred.get_token(*DELEGATED_SCOPES, enable_cae=True); print('Done')`"" -ForegroundColor White
    Write-Host ""
} else {
    Write-Ok "Authentication cached. Claude Code will use it silently."
}

# ---------------------------------------------------------------
# Step 7: Summary
# ---------------------------------------------------------------
Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  Teams MCP Server installed successfully!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Install location:  $InstallDir"
Write-Host "  Python:            $PythonExe"
Write-Host "  MCP server name:   teams-mcp"
Write-Host ""
Write-Host "  Verify with:  claude mcp list" -ForegroundColor Cyan
Write-Host ""
