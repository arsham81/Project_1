param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'ZCMCPlanCompare'),
    [string]$DesktopRoot = [Environment]::GetFolderPath('Desktop'),
    [switch]$NoLaunch
)
$ErrorActionPreference = 'Stop'
try {
    $python = $null
    foreach ($command in @('py', 'python')) {
        if (Get-Command $command -ErrorAction SilentlyContinue) {
            try {
                $probe = & $command -c 'import sys; print(sys.executable if sys.version_info >= (3,10) else "")' 2>$null
                if ($LASTEXITCODE -eq 0 -and $probe -and (Test-Path $probe)) {
                    $python = [string]$probe
                    break
                }
            } catch {}
        }
    }
    if (-not $python) {
        Write-Host 'Python 3.10+ is required. Install Python and then run install.bat again.'
        Write-Host 'In the Python installer, select Add Python to PATH.'
        Start-Process 'https://www.python.org/downloads/windows/'
        exit 1
    }
    New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
    foreach ($name in @('app.py','desktop.py','report_export.py','requirements-desktop.txt','icon.ico','README-Windows.md')) {
        $source = Join-Path $PSScriptRoot $name
        $destination = Join-Path $InstallRoot $name
        if ([IO.Path]::GetFullPath($source) -ne [IO.Path]::GetFullPath($destination)) {
            Copy-Item -LiteralPath $source -Destination $destination -Force
        }
    }
    $venv = Join-Path $InstallRoot '.venv'
    $runtime = Join-Path $venv 'Scripts\python.exe'
    if (-not (Test-Path $runtime)) {
        & $python -m venv $venv
        if ($LASTEXITCODE -ne 0) { throw 'Could not create Python environment.' }
    }
    & $runtime -m pip install -r (Join-Path $InstallRoot 'requirements-desktop.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed. Check the Internet connection and try again.' }
    $pythonw = Join-Path $venv 'Scripts\pythonw.exe'
    $entry = Join-Path $InstallRoot 'desktop.py'
    # Validate imports and Tk before creating a shortcut that launches the application.
    & $runtime -c 'import tkinter, openpyxl, matplotlib, PIL; r=tkinter.Tk(); r.withdraw(); r.destroy()'
    if ($LASTEXITCODE -ne 0) { throw 'Required application components could not be loaded.' }
    New-Item -ItemType Directory -Force -Path $DesktopRoot | Out-Null
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut((Join-Path $DesktopRoot 'ZCMC Plan Compare.lnk'))
    $shortcut.TargetPath = $pythonw
    $shortcut.Arguments = '"' + $entry + '"'
    $shortcut.WorkingDirectory = $InstallRoot
    $shortcut.IconLocation = (Join-Path $InstallRoot 'icon.ico')
    $shortcut.Description = 'Compare two Excel plans and export Excel and PDF charts'
    $shortcut.Save()
    Write-Host 'Installed. Use the ZCMC Plan Compare icon on your Desktop.' -ForegroundColor Green
    if (-not $NoLaunch) { Start-Process -FilePath $pythonw -ArgumentList ('"'+$entry+'"') -WorkingDirectory $InstallRoot }
} catch {
    Write-Host ('Installation failed: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
