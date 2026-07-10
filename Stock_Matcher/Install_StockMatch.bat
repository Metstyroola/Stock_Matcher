@echo off
setlocal enabledelayedexpansion
title StockMatch Tyroola Installer
cls

echo.
echo  ============================================================
echo   StockMatch Tyroola Installer
echo   Tyre supplier stock matching app
echo  ============================================================
echo.

set "DIR=%LOCALAPPDATA%\StockMatch_Tyroola"
set "SOURCE=%~dp0"
set "DESKTOP=%USERPROFILE%\Desktop"
set "STARTMENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs"

echo  Source  : %SOURCE%
echo  Install : %DIR%
echo.

:: ── Check Node.js ────────────────────────────────────────────────────────
echo  [1/3] Checking Node.js...
where node >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Node.js is required to run StockMatch.
    echo.
    set /p DL="  Open nodejs.org to download it now? [Y/N]: "
    if /i "!DL!"=="Y" (
        start https://nodejs.org/en/download
        echo.
        echo  Install Node.js LTS, then run this installer again.
        pause & exit /b 0
    )
    echo.
    echo  ERROR: Cannot continue without Node.js.
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('node --version 2^>nul') do set "NODE_VER=%%v"
echo  [OK] Node.js !NODE_VER!

:: ── Check source files ────────────────────────────────────────────────────
echo  [2/3] Checking files...
if not exist "%SOURCE%stock_matcher_app.html" (
    echo  ERROR: stock_matcher_app.html not found in %SOURCE%
    echo  Run this installer from the StockMatch folder.
    pause & exit /b 1
)
if not exist "%SOURCE%server.js" (
    echo  ERROR: server.js not found in %SOURCE%
    pause & exit /b 1
)
echo  [OK] All required files found

:: ── Download JS libraries ─────────────────────────────────────────────────
echo  [3/4] Downloading required libraries...
if not exist "%DIR%\lib" mkdir "%DIR%\lib"

if exist "%DIR%\lib\xlsx.full.min.js" (
    echo  [OK] xlsx library already downloaded
) else (
    echo  Downloading xlsx.full.min.js...
    powershell -NoProfile -Command ^
      "[Net.ServicePointManager]::SecurityProtocol='Tls12';" ^
      "try{(New-Object Net.WebClient).DownloadFile('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js','%DIR%\lib\xlsx.full.min.js');Write-Host '[OK] xlsx downloaded'}catch{Write-Host '[WARN] xlsx download failed - trying alternate...';try{(New-Object Net.WebClient).DownloadFile('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js','%DIR%\lib\xlsx.full.min.js');Write-Host '[OK] xlsx downloaded (alternate)'}catch{Write-Host '[WARN] xlsx download failed - app may not parse Excel files'}}" 2>nul
)

if exist "%DIR%\lib\papaparse.min.js" (
    echo  [OK] papaparse library already downloaded
) else (
    echo  Downloading papaparse.min.js...
    powershell -NoProfile -Command ^
      "[Net.ServicePointManager]::SecurityProtocol='Tls12';" ^
      "try{(New-Object Net.WebClient).DownloadFile('https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js','%DIR%\lib\papaparse.min.js');Write-Host '[OK] papaparse downloaded'}catch{Write-Host '[WARN] papaparse download failed'}" 2>nul
)
if not exist "%DIR%" mkdir "%DIR%"

copy /Y "%SOURCE%stock_matcher_app.html" "%DIR%\stock_matcher_app.html" >nul
echo  [OK] App HTML installed

copy /Y "%SOURCE%server.js"             "%DIR%\server.js"              >nul
echo  [OK] Server installed

copy /Y "%SOURCE%StockMatch.bat"        "%DIR%\StockMatch.bat"         >nul
echo  [OK] Launcher installed

if exist "%SOURCE%.env.example" (
    if not exist "%DIR%\.env" (
        copy /Y "%SOURCE%.env.example"  "%DIR%\.env.example"           >nul
    )
)

:: ── Install xlsx package for server ──────────────────────────────────────
if "%NODE_INST%"=="1" (
    if not exist "%DIR%\node_modules\xlsx" (
        echo  Installing xlsx for server Drive reading...
        cd /d "%DIR%"
        call npm install xlsx --save --no-audit --no-fund >nul 2>&1
        if exist "%DIR%\node_modules\xlsx" (
            echo  [OK] xlsx installed
        ) else (
            echo  [WARN] xlsx install failed - Drive XLSX reading may not work
        )
    ) else (
        echo  [OK] xlsx already installed
    )
)
powershell -NoProfile -Command ^
  "$ws=New-Object -ComObject WScript.Shell;" ^
  "$sc=$ws.CreateShortcut('%DESKTOP%\StockMatch Tyroola.lnk');" ^
  "$sc.TargetPath='%DIR%\StockMatch.bat';" ^
  "$sc.WorkingDirectory='%DIR%';" ^
  "$sc.IconLocation='%SystemRoot%\System32\SHELL32.dll,14';" ^
  "$sc.Description='StockMatch - Tyre supplier matching';" ^
  "$sc.WindowStyle=1;" ^
  "$sc.Save()" >nul 2>&1
echo  [OK] Desktop shortcut created

:: Start menu
if not exist "%STARTMENU%\StockMatch Tyroola" mkdir "%STARTMENU%\StockMatch Tyroola" >nul 2>&1
powershell -NoProfile -Command ^
  "$ws=New-Object -ComObject WScript.Shell;" ^
  "$sc=$ws.CreateShortcut('%STARTMENU%\StockMatch Tyroola\StockMatch.lnk');" ^
  "$sc.TargetPath='%DIR%\StockMatch.bat';" ^
  "$sc.WorkingDirectory='%DIR%';" ^
  "$sc.IconLocation='%SystemRoot%\System32\SHELL32.dll,14';" ^
  "$sc.Save()" >nul 2>&1
echo  [OK] Start Menu shortcut created

:: ── Done ──────────────────────────────────────────────────────────────────
echo.
echo  ============================================================
echo   Installation complete!
echo  ============================================================
echo.
echo   How to open StockMatch:
echo     Double-click "StockMatch" on your Desktop
echo.
echo   What happens when you open it:
echo     1. A terminal window opens (keep it open)
echo     2. Chrome/Edge opens with the app at http://localhost:3000
echo     3. Use the app in the browser
echo     4. Close the terminal to stop
echo.
echo   To set up live BigQuery refresh (one time):
echo     Run: gcloud auth application-default login
echo.

set /p LAUNCH="  Open StockMatch now? [Y/N]: "
if /i "%LAUNCH%"=="Y" (
    start "" "%DIR%\StockMatch.bat"
)
echo.
timeout /t 3 >nul
exit /b 0
