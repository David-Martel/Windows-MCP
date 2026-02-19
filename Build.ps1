#Requires -Version 7.0
<#
.SYNOPSIS
    Unified build, test, and release script for Windows-MCP.
.DESCRIPTION
    Orchestrates Python (uv/ruff/pytest) and Rust (maturin/cargo via CargoTools)
    builds, testing, linting, version tagging, and GitHub releases.
.PARAMETER Action
    The build action to perform. Default: Build
.PARAMETER Release
    Create a GitHub release after building. Requires gh CLI.
.PARAMETER BumpVersion
    Version bump type for tagging: major, minor, patch. Default: patch
.PARAMETER NativeOnly
    Only build the Rust native extension.
.PARAMETER PythonOnly
    Only run Python build/test/lint steps.
.PARAMETER SkipTests
    Skip running tests.
.PARAMETER SkipLint
    Skip linting step.
.PARAMETER Verbose
    Show detailed output.
.EXAMPLE
    .\Build.ps1                    # Full build (Python + Rust)
    .\Build.ps1 -Action Test       # Run all tests
    .\Build.ps1 -Action Lint       # Lint only
    .\Build.ps1 -Action Native     # Build Rust extension only
    .\Build.ps1 -Action Release -BumpVersion minor  # Tag + release
    .\Build.ps1 -Action Clean      # Clean build artifacts
    .\Build.ps1 -Action Check      # Preflight checks (cargo check + ruff)
#>
[CmdletBinding()]
param(
    [ValidateSet('Build', 'Test', 'Lint', 'Native', 'Release', 'Clean', 'Check', 'All')]
    [string]$Action = 'Build',

    [switch]$Release,

    [ValidateSet('major', 'minor', 'patch')]
    [string]$BumpVersion = 'patch',

    [switch]$NativeOnly,
    [switch]$PythonOnly,
    [switch]$SkipTests,
    [switch]$SkipLint
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProjectRoot = $PSScriptRoot
$NativeDir = Join-Path $ProjectRoot 'native'
$SrcDir = Join-Path $ProjectRoot 'src'
$TestDir = Join-Path $ProjectRoot 'tests'

# ──────────────────────────────────────────────────────────────────────
#  CargoTools Integration
# ──────────────────────────────────────────────────────────────────────

$CargoToolsAvailable = $false
try {
    Import-Module CargoTools -ErrorAction Stop
    $CargoToolsAvailable = $true
    Write-Host "[+] CargoTools module loaded" -ForegroundColor Green
} catch {
    Write-Host "[!] CargoTools not available -- using direct cargo/maturin" -ForegroundColor Yellow
}

# ──────────────────────────────────────────────────────────────────────
#  Helpers
# ──────────────────────────────────────────────────────────────────────

function Write-Phase {
    param([string]$Name, [string]$Status = 'Starting')
    $color = switch ($Status) {
        'Starting' { 'Cyan' }
        'Done'     { 'Green' }
        'Skipped'  { 'Yellow' }
        'Failed'   { 'Red' }
        default    { 'White' }
    }
    Write-Host "`n== $Name [$Status] ==" -ForegroundColor $color
}

function Invoke-Step {
    param([string]$Name, [scriptblock]$Block)
    Write-Phase $Name 'Starting'
    try {
        & $Block
        if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
            Write-Phase $Name 'Failed'
            throw "$Name failed with exit code $LASTEXITCODE"
        }
        Write-Phase $Name 'Done'
    } catch {
        Write-Phase $Name 'Failed'
        throw
    }
}

function Get-ProjectVersion {
    $pyproject = Join-Path $ProjectRoot 'pyproject.toml'
    $content = Get-Content $pyproject -Raw
    if ($content -match 'version\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return '0.0.0'
}

function Set-ProjectVersion {
    param([string]$NewVersion)
    $pyproject = Join-Path $ProjectRoot 'pyproject.toml'
    $content = Get-Content $pyproject -Raw
    $content = $content -replace '(version\s*=\s*)"[^"]+"', "`$1`"$NewVersion`""
    Set-Content $pyproject $content -NoNewline

    # Also update native/Cargo.toml if it exists
    $cargoToml = Join-Path $NativeDir 'Cargo.toml'
    if (Test-Path $cargoToml) {
        $cargo = Get-Content $cargoToml -Raw
        $cargo = $cargo -replace '(version\s*=\s*)"[^"]+"', "`$1`"$NewVersion`""
        Set-Content $cargoToml $cargo -NoNewline
    }
}

function Get-BumpedVersion {
    param([string]$Current, [string]$Bump)
    $parts = $Current.Split('.')
    $major = [int]$parts[0]
    $minor = if ($parts.Count -gt 1) { [int]$parts[1] } else { 0 }
    $patch = if ($parts.Count -gt 2) { [int]$parts[2] } else { 0 }

    switch ($Bump) {
        'major' { $major++; $minor = 0; $patch = 0 }
        'minor' { $minor++; $patch = 0 }
        'patch' { $patch++ }
    }
    return "$major.$minor.$patch"
}

# ──────────────────────────────────────────────────────────────────────
#  Build Steps
# ──────────────────────────────────────────────────────────────────────

function Step-PythonSync {
    Invoke-Step 'Python Sync' {
        Push-Location $ProjectRoot
        try { uv sync --extra dev } finally { Pop-Location }
    }
}

function Step-PythonLint {
    Invoke-Step 'Python Lint (zero-warning enforcement)' {
        Push-Location $ProjectRoot
        try {
            # Check for lint errors -- no --fix, CI must see failures
            uv run ruff check .
            # Check formatting
            uv run ruff format --check .
        } finally { Pop-Location }
    }
}

function Step-PythonTest {
    Invoke-Step 'Python Tests' {
        Push-Location $ProjectRoot
        try {
            uv run python -m pytest tests/ -v --tb=short
        } finally { Pop-Location }
    }
}

function Step-NativeBuild {
    if (-not (Test-Path $NativeDir)) {
        Write-Phase 'Native Build' 'Skipped'
        Write-Host "  No native/ directory found" -ForegroundColor Yellow
        return
    }

    Invoke-Step 'Native Build (Rust workspace -- warnings=deny)' {
        $savedRustFlags = $env:RUSTFLAGS
        $savedPyo3Python = $env:PYO3_PYTHON
        try {
            # PYO3_PYTHON must point to the venv Python for correct linking
            $venvPython = Join-Path $ProjectRoot '.venv' 'Scripts' 'python.exe'
            if (Test-Path $venvPython) {
                $env:PYO3_PYTHON = $venvPython
            }

            if ($CargoToolsAvailable) {
                Write-Host "  Using CargoTools Invoke-CargoWrapper" -ForegroundColor Cyan
                $env:RUSTFLAGS = '-D warnings'
                Invoke-CargoWrapper -Command 'build' `
                    -AdditionalArgs @('--release', '--workspace') `
                    -WorkingDirectory $NativeDir
            } else {
                Write-Host "  Using direct cargo build (warnings=deny)" -ForegroundColor Yellow
                Push-Location $NativeDir
                try {
                    $env:RUSTC_WRAPPER = ''
                    $env:RUSTFLAGS = '-D warnings'
                    cargo build --release --workspace
                } finally { Pop-Location }
            }
        } finally {
            $env:RUSTFLAGS = $savedRustFlags
            $env:PYO3_PYTHON = $savedPyo3Python
        }
    }
}

function Find-NativeArtifact {
    <#
    .SYNOPSIS
    Searches cargo target directories for a named build artifact.
    .DESCRIPTION
    Checks (in order): CARGO_TARGET_DIR env, native/target/.
    Returns the first found path or $null.
    .PARAMETER Name
    The artifact filename (e.g. 'windows_mcp_core.dll', 'wmcp-worker.exe').
    #>
    param([string]$Name)
    $searchPaths = @()

    # 1. CARGO_TARGET_DIR environment variable (highest priority)
    if ($env:CARGO_TARGET_DIR) {
        $searchPaths += Join-Path $env:CARGO_TARGET_DIR 'release' $Name
    }

    # 2. Local native/target/ (standard cargo location)
    $searchPaths += Join-Path $NativeDir 'target' 'release' $Name

    foreach ($path in $searchPaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    return $null
}

function Step-NativeInstall {
    if (-not (Test-Path $NativeDir)) {
        Write-Phase 'Native Install' 'Skipped'
        return
    }

    Invoke-Step 'Native Install (workspace artifacts to venv)' {
        $venvSitePackages = Join-Path $ProjectRoot '.venv' 'Lib' 'site-packages'
        $venvScripts = Join-Path $ProjectRoot '.venv' 'Scripts'
        if (-not (Test-Path $venvSitePackages)) {
            Write-Host "  venv not found at $venvSitePackages -- run 'uv sync' first" -ForegroundColor Yellow
            return
        }

        $installed = 0

        # 1. PyO3 extension: windows_mcp_core.dll -> .pyd
        $pyo3Dll = Find-NativeArtifact 'windows_mcp_core.dll'
        if ($pyo3Dll) {
            $dest = Join-Path $venvSitePackages 'windows_mcp_core.pyd'
            Copy-Item $pyo3Dll $dest -Force
            $sizeKB = [math]::Round((Get-Item $pyo3Dll).Length / 1024, 0)
            Write-Host "  Installed windows_mcp_core.pyd (${sizeKB} KB)" -ForegroundColor Green
            $installed++

            # Verify the module loads
            $testResult = & (Join-Path $venvScripts 'python.exe') -c "import windows_mcp_core; print('OK')" 2>&1
            if ($testResult -eq 'OK') {
                Write-Host "  Import verification: OK" -ForegroundColor Green
            } else {
                Write-Host "  Import verification: FAILED -- $testResult" -ForegroundColor Red
            }
        } else {
            Write-Host "  windows_mcp_core.dll not found" -ForegroundColor Yellow
        }

        # 2. FFI DLL: windows_mcp_ffi.dll -> site-packages (for ctypes)
        $ffiDll = Find-NativeArtifact 'windows_mcp_ffi.dll'
        if ($ffiDll) {
            $srcPkg = Join-Path $ProjectRoot 'src' 'windows_mcp'
            $dest = Join-Path $srcPkg 'windows_mcp_ffi.dll'
            Copy-Item $ffiDll $dest -Force
            $sizeKB = [math]::Round((Get-Item $ffiDll).Length / 1024, 0)
            Write-Host "  Installed windows_mcp_ffi.dll (${sizeKB} KB)" -ForegroundColor Green
            $installed++
        }

        # 3. Worker binary: wmcp-worker.exe -> venv/Scripts
        $workerExe = Find-NativeArtifact 'wmcp-worker.exe'
        if ($workerExe) {
            $dest = Join-Path $venvScripts 'wmcp-worker.exe'
            Copy-Item $workerExe $dest -Force
            $sizeKB = [math]::Round((Get-Item $workerExe).Length / 1024, 0)
            Write-Host "  Installed wmcp-worker.exe (${sizeKB} KB)" -ForegroundColor Green
            $installed++
        }

        if ($installed -eq 0) {
            Write-Host "  No artifacts found. Run 'Build.ps1 -Action Native' to compile first" -ForegroundColor Yellow
        } else {
            Write-Host "  Installed $installed artifact(s)" -ForegroundColor Green
        }
    }
}

function Step-NativeCheck {
    if (-not (Test-Path $NativeDir)) { return }

    Invoke-Step 'Native Check (cargo check + clippy -- workspace)' {
        if ($CargoToolsAvailable) {
            Invoke-CargoWrapper -Command 'check' `
                -AdditionalArgs @('--workspace') `
                -WorkingDirectory $NativeDir
            Invoke-CargoWrapper -Command 'clippy' `
                -AdditionalArgs @('--workspace', '--all-targets', '--', '-D', 'warnings') `
                -WorkingDirectory $NativeDir
        } else {
            Push-Location $NativeDir
            try {
                $env:RUSTC_WRAPPER = ''
                cargo check --workspace
                cargo clippy --workspace --all-targets -- -D warnings
            } finally { Pop-Location }
        }
    }
}

function Step-NativeTest {
    if (-not (Test-Path $NativeDir)) { return }

    Invoke-Step 'Native Tests (Rust workspace)' {
        if ($CargoToolsAvailable) {
            Invoke-CargoWrapper -Command 'test' `
                -AdditionalArgs @('--workspace') `
                -WorkingDirectory $NativeDir
        } else {
            Push-Location $NativeDir
            try {
                $env:RUSTC_WRAPPER = ''
                cargo test --workspace
            } finally { Pop-Location }
        }
    }
}

function Step-Clean {
    Invoke-Step 'Clean' {
        # Python
        $pycacheDirs = Get-ChildItem -Path $ProjectRoot -Recurse -Directory -Filter '__pycache__' -ErrorAction SilentlyContinue
        foreach ($dir in $pycacheDirs) {
            Remove-Item $dir.FullName -Recurse -Force
            Write-Host "  Removed $($dir.FullName)" -ForegroundColor Gray
        }

        $eggInfo = Get-ChildItem -Path $ProjectRoot -Recurse -Directory -Filter '*.egg-info' -ErrorAction SilentlyContinue
        foreach ($dir in $eggInfo) {
            Remove-Item $dir.FullName -Recurse -Force
        }

        # Rust
        $nativeTarget = Join-Path $NativeDir 'target'
        if (Test-Path $nativeTarget) {
            Remove-Item $nativeTarget -Recurse -Force
            Write-Host "  Removed $nativeTarget" -ForegroundColor Gray
        }

        # .pytest_cache
        $pytestCache = Join-Path $ProjectRoot '.pytest_cache'
        if (Test-Path $pytestCache) {
            Remove-Item $pytestCache -Recurse -Force
        }

        Write-Host "  Clean complete" -ForegroundColor Green
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Version Tagging & GitHub Release
# ──────────────────────────────────────────────────────────────────────

function Step-VersionTag {
    param([string]$Bump = 'patch')

    Invoke-Step "Version Tag ($Bump)" {
        $current = Get-ProjectVersion
        $new = Get-BumpedVersion $current $Bump
        Write-Host "  $current -> $new" -ForegroundColor Cyan

        Set-ProjectVersion $new

        Push-Location $ProjectRoot
        try {
            git add pyproject.toml
            if (Test-Path (Join-Path $NativeDir 'Cargo.toml')) {
                git add (Join-Path $NativeDir 'Cargo.toml')
            }
            git commit -m "chore: bump version to $new"
            git tag -a "v$new" -m "Release v$new"
            Write-Host "  Tagged v$new" -ForegroundColor Green
        } finally { Pop-Location }
    }
}

function Step-GitHubRelease {
    Invoke-Step 'GitHub Release' {
        $gh = Get-Command gh -ErrorAction SilentlyContinue
        if (-not $gh) {
            Write-Host "  gh CLI not found -- install from https://cli.github.com" -ForegroundColor Yellow
            return
        }

        $version = Get-ProjectVersion
        $tag = "v$version"

        Push-Location $ProjectRoot
        try {
            # Push tag to remote
            git push origin main --tags

            # Create release with auto-generated notes
            gh release create $tag `
                --title "Windows-MCP $tag" `
                --generate-notes `
                --latest
            Write-Host "  Released $tag on GitHub" -ForegroundColor Green
        } finally { Pop-Location }
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Action Dispatch
# ──────────────────────────────────────────────────────────────────────

$startTime = Get-Date

Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host "  Windows-MCP Build System" -ForegroundColor Cyan
Write-Host "  Version: $(Get-ProjectVersion)" -ForegroundColor Cyan
Write-Host "  Action:  $Action" -ForegroundColor Cyan
Write-Host "  CargoTools: $(if ($CargoToolsAvailable) { 'Yes' } else { 'No' })" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

switch ($Action) {
    'Build' {
        Step-PythonSync
        if (-not $PythonOnly -and -not $SkipLint) { Step-PythonLint }
        if (-not $PythonOnly) { Step-NativeBuild; Step-NativeInstall }
        if (-not $SkipTests -and -not $NativeOnly) { Step-PythonTest }
    }
    'Test' {
        if (-not $NativeOnly) { Step-PythonTest }
        if (-not $PythonOnly) { Step-NativeTest }
    }
    'Lint' {
        Step-PythonLint
        if (-not $PythonOnly) { Step-NativeCheck }
    }
    'Native' {
        Step-NativeBuild
        Step-NativeInstall
        if (-not $SkipTests) { Step-NativeTest }
    }
    'Check' {
        Step-PythonLint
        Step-NativeCheck
    }
    'Clean' {
        Step-Clean
    }
    'Release' {
        # Full build + test + tag + release
        Step-PythonSync
        Step-PythonLint
        Step-NativeBuild
        Step-NativeInstall
        Step-PythonTest
        Step-NativeTest
        Step-VersionTag -Bump $BumpVersion
        Step-GitHubRelease
    }
    'All' {
        Step-PythonSync
        Step-PythonLint
        Step-NativeBuild
        Step-NativeInstall
        Step-PythonTest
        Step-NativeTest
    }
}

if ($Release -and $Action -ne 'Release') {
    Step-VersionTag -Bump $BumpVersion
    Step-GitHubRelease
}

$elapsed = (Get-Date) - $startTime
Write-Host "`n=====================================================" -ForegroundColor Green
Write-Host "  Build complete in $([math]::Round($elapsed.TotalSeconds, 1))s" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
