@echo off
setlocal enabledelayedexpansion
title StockMatch Tyroola
cls

set "DIR=%LOCALAPPDATA%\StockMatch_Tyroola"
set "PORT=3000"
set "URL=http://localhost:%PORT%"

echo.
echo  ============================================================
echo   StockMatch Tyroola - Starting...
echo  ============================================================
echo.

:: ── Check Node.js ────────────────────────────────────────────────────────
where node >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Node.js not installed.
    echo  Download from: https://nodejs.org
    pause & exit /b 1
)

:: ── Check server.js ───────────────────────────────────────────────────────
if not exist "%DIR%\server.js" (
    echo  ERROR: server.js not found at %DIR%
    echo  Run Install_StockMatch.bat first.
    pause & exit /b 1
)

:: ── Kill anything already on port 3000 ──────────────────────────────────
echo  Checking port 3000...
powershell -NoProfile -Command ^
  "try { $p=Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue; if($p){ Stop-Process -Id $p.OwningProcess -Force -ErrorAction SilentlyContinue; Write-Host '  Stopped existing process on port 3000' } } catch {}"
timeout /t 1 >nul 2>&1

:: ── Check if server already running ──────────────────────────────────────
curl -s "%URL%/api/health" >nul 2>&1
if not errorlevel 1 (
    echo  Server already running - opening browser...
    goto :openbrowser
)

:: ── Start server in background ────────────────────────────────────────────
echo  Starting server...
cd /d "%DIR%"
start /b "" node server.js > "%DIR%\server.log" 2>&1

:: Wait for server to start (up to 10 seconds)
echo  Waiting for server to start...
set "TRIES=0"
:waitloop
timeout /t 1 >nul 2>&1
set /a TRIES+=1
curl -s "%URL%/api/health" >nul 2>&1
if not errorlevel 1 goto :serverready
if !TRIES! lss 10 goto :waitloop

echo.
echo  Server failed to start. Check %DIR%\server.log
echo.
type "%DIR%\server.log"
pause & exit /b 1

:serverready
echo  [OK] Server running at %URL%

:openbrowser
echo  Opening StockMatch in browser...
echo.

:: Try Chrome first, then Edge, then default browser
set "CHROME="
for %%p in (
    "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    "%LocalAppData%\Google\Chrome\Application\chrome.exe"
) do (
    if exist %%p set "CHROME=%%p"
)

if defined CHROME (
    echo  Opening in Chrome: %URL%
    start "" %CHROME% --app="%URL%" --window-size=1440,900
    goto :done
)

:: Try Edge
set "EDGE="
for %%p in (
    "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
    "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
) do (
    if exist %%p set "EDGE=%%p"
)

if defined EDGE (
    echo  Opening in Edge: %URL%
    start "" %EDGE% --app="%URL%" --window-size=1440,900
    goto :done
)

:: Fallback to default browser
echo  Opening in default browser: %URL%
start "" "%URL%"

:done
echo.
echo  ============================================================
echo   StockMatch is running at %URL%
echo   Close this window to stop the server.
echo  ============================================================
echo.

:: Keep window open (server runs in background)
echo  Press Ctrl+C or close this window to stop StockMatch.
echo.
node "%DIR%\server.js"
