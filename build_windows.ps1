param(
    [switch]$SkipVenv
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

if (-not $SkipVenv) {
    if (-not (Test-Path ".venv\Scripts\python.exe")) {
        if (Get-Command py -ErrorAction SilentlyContinue) {
            py -3.11 -m venv .venv
        } elseif (Get-Command python -ErrorAction SilentlyContinue) {
            python -m venv .venv
        } else {
            throw "Kein Python-Launcher gefunden (weder 'py' noch 'python')."
        }
    }
}

$PythonExe = ".\.venv\Scripts\python.exe"
if (-not (Test-Path $PythonExe)) {
    throw "Python-venv nicht gefunden: $PythonExe"
}

& $PythonExe -m pip install --upgrade pip
& $PythonExe -m pip install -r requirements.txt
& $PythonExe -m pip install pyinstaller

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name KopfnotenTool `
    app.py

Write-Host ""
Write-Host "Build fertig: $ProjectRoot\dist\KopfnotenTool.exe"
Write-Host "Installer-Skript: $ProjectRoot\installer\KopfnotenTool.iss"
