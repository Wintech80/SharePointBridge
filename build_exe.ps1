# Builds dist\SharePointBridge.exe (single file, windowed app with a system tray icon).
#   powershell -ExecutionPolicy Bypass -File build_exe.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
python -m pip install -r requirements.txt pyinstaller | Out-Null
python -m PyInstaller --noconfirm --onefile --noconsole --name SharePointBridge `
    --collect-all wsgidav --collect-all cheroot --collect-all msal `
    --hidden-import graph --hidden-import provider --hidden-import register `
    --hidden-import housekeeping --hidden-import winutil --hidden-import tray `
    --hidden-import pystray._win32 `
    main.py
Write-Host "Built: $PSScriptRoot\dist\SharePointBridge.exe"
