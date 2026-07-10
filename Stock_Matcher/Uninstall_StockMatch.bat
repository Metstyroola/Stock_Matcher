@echo off
setlocal
title StockMatch Tyroola Uninstaller
cls

echo.
echo  ============================================================
echo   StockMatch Uninstaller
echo  ============================================================
echo.
echo  This will remove StockMatch from your computer.
echo.
echo  The following will be REMOVED:
echo    - App files in: %LOCALAPPDATA%\StockMatch_Tyroola
echo    - Desktop shortcut: StockMatch
echo    - Desktop shortcut: StockMatch Server
echo    - Start Menu folder: StockMatch
echo.
echo  The following will be KEPT:
echo    - Your SC snapshot file (sc_au_snapshot.json in Drive)
echo    - Your settings and API keys (stored in browser)
echo.

set /p CONFIRM="  Are you sure you want to uninstall? [Y/N]: "
if /i not "%CONFIRM%"=="Y" (
    echo  Uninstall cancelled.
    timeout /t 2 >nul
    exit /b 0
)

echo.
echo  Removing...

set "INSTALL_DIR=%LOCALAPPDATA%\StockMatch_Tyroola"
set "DESKTOP=%USERPROFILE%\Desktop"
set "STARTMENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs"

:: Remove install folder
if exist "%INSTALL_DIR%" (
    rd /s /q "%INSTALL_DIR%" >nul 2>&1
    if exist "%INSTALL_DIR%" (
        echo  WARNING: Could not fully remove %INSTALL_DIR%
        echo  Please delete it manually.
    ) else (
        echo  [OK] App files removed
    )
) else (
    echo  [OK] App folder not found (already removed)
)

:: Remove desktop shortcuts
if exist "%DESKTOP%\StockMatch Tyroola.lnk" (
    del "%DESKTOP%\StockMatch Tyroola.lnk" >nul 2>&1
    echo  [OK] Desktop shortcut removed
)
if exist "%DESKTOP%\StockMatch Tyroola Server.lnk" (
    del "%DESKTOP%\StockMatch Tyroola Server.lnk" >nul 2>&1
    echo  [OK] Server shortcut removed
)

:: Remove Start Menu
if exist "%STARTMENU%\StockMatch Tyroola" (
    rd /s /q "%STARTMENU%\StockMatch Tyroola" >nul 2>&1
    echo  [OK] Start Menu entries removed
)

echo.
echo  ============================================================
echo   Uninstall complete. StockMatch has been removed.
echo  ============================================================
echo.
pause
exit /b 0
