param(
    [string]$PythonExe = "",
    [string]$PythonVersion = "3.12",
    [string]$VenvPath = ".venv-build",
    [switch]$NoVenv,
    [switch]$RecreateVenv
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Exe,
        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Args
    )

    & $Exe @Args
    if ($LASTEXITCODE -ne 0) {
        $argsText = ($Args -join " ").Trim()
        throw "Command failed with exit code ${LASTEXITCODE}: $Exe $argsText"
    }
}

function Get-PythonMajorMinor {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonCmd
    )

    $version = (& $PythonCmd -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')").Trim()
    if ($LASTEXITCODE -ne 0 -or -not $version) {
        throw "Could not determine Python version from: $PythonCmd"
    }
    return $version
}

function Get-VenvMajorMinor {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VenvPath
    )

    $cfgPath = Join-Path $VenvPath "pyvenv.cfg"
    if (-not (Test-Path $cfgPath)) {
        return $null
    }

    $versionLine = Get-Content $cfgPath | Where-Object { $_ -match "^\s*version\s*=" } | Select-Object -First 1
    if (-not $versionLine) {
        return $null
    }

    $rawVersion = (($versionLine -split "=", 2)[1]).Trim()
    if ($rawVersion -match "^(\d+\.\d+)") {
        return $Matches[1]
    }

    return $null
}

function Resolve-BasePython {
    param(
        [string]$OverrideExe,
        [string]$Version
    )

    if ($OverrideExe) {
        return $OverrideExe
    }

    if (Get-Command py -ErrorAction SilentlyContinue) {
        try {
            $resolved = (& py "-$Version" -c "import sys; print(sys.executable)").Trim()
            if ($LASTEXITCODE -eq 0 -and $resolved -and (Test-Path $resolved)) {
                Write-Host "Using Python $Version from py launcher: $resolved"
                return $resolved
            }
        } catch {
        }
    }

    Write-Host "Falling back to 'python' from PATH."
    return "python"
}

function Ensure-Pip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonCmd
    )

    $pipReady = $false
    try {
        & $PythonCmd -m pip --version *> $null
        $pipReady = ($LASTEXITCODE -eq 0)
    } catch {
        $pipReady = $false
    }

    if ($pipReady) {
        return
    }

    Write-Host "pip missing in build environment. Bootstrapping with ensurepip..."
    Invoke-Checked $PythonCmd -m ensurepip --upgrade
    Invoke-Checked $PythonCmd -m pip --version
}

$basePython = Resolve-BasePython -OverrideExe $PythonExe -Version $PythonVersion
$basePythonVersion = Get-PythonMajorMinor $basePython
Write-Host "Build interpreter: $basePython (Python $basePythonVersion)"

if (-not $NoVenv) {
    if ([System.IO.Path]::IsPathRooted($VenvPath)) {
        $venvPath = $VenvPath
    } else {
        $venvPath = Join-Path $root $VenvPath
    }
    $venvPython = Join-Path $venvPath "Scripts\python.exe"
    $venvVersion = Get-VenvMajorMinor $venvPath
    $needsRefresh = $RecreateVenv

    if (-not $needsRefresh -and $venvVersion -and ($venvVersion -ne $basePythonVersion)) {
        $needsRefresh = $true
        Write-Host "Refreshing venv due to version mismatch ($venvVersion -> $basePythonVersion)."
    }

    if ($needsRefresh -and (Test-Path $venvPath)) {
        Write-Host "Refreshing build environment: $venvPath"
        Invoke-Checked $basePython -m venv --clear $venvPath
    }

    if (-not (Test-Path $venvPython)) {
        Write-Host "Creating build environment: $venvPath"
        Invoke-Checked $basePython -m venv $venvPath
    }

    $pythonCmd = $venvPython
    if (-not (Test-Path $pythonCmd)) {
        throw "Build Python executable not found: $pythonCmd"
    }
} else {
    $pythonCmd = $basePython
}

Ensure-Pip $pythonCmd
Invoke-Checked $pythonCmd -m pip install --upgrade pip
Invoke-Checked $pythonCmd -m pip install -r requirements-build.txt

if (Test-Path "build") {
    Remove-Item -Recurse -Force "build"
}
if (Test-Path "dist") {
    Remove-Item -Recurse -Force "dist"
}

Invoke-Checked $pythonCmd -m PyInstaller --noconfirm --clean POtrol.spec

$exePath = Join-Path $root "dist\POtrol.exe"
if (-not (Test-Path $exePath)) {
    throw "Build finished without expected output: $exePath"
}

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $exePath"
