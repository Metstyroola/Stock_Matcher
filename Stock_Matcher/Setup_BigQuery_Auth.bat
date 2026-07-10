@echo off
setlocal enabledelayedexpansion
title StockMatch - Google Auth Setup
cls

echo.
echo  ============================================================
echo   StockMatch - Google Authentication Setup
echo  ============================================================
echo.
echo  This signs you in to Google Cloud via gcloud.
echo  A browser window will open.
echo  Sign in with: emeterio@tyroola.com
echo.

:: Find gcloud
set "GCLOUD="
for %%p in (
    "%LOCALAPPDATA%\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd"
    "%PROGRAMFILES%\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd"
    "%PROGRAMFILES(X86)%\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd"
    "%USERPROFILE%\AppData\Local\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd"
) do ( if exist %%p set "GCLOUD=%%p" )

if not defined GCLOUD (
    where gcloud >nul 2>&1
    if not errorlevel 1 set "GCLOUD=gcloud"
)

if not defined GCLOUD (
    echo  ERROR: Google Cloud SDK not found.
    echo  Install from: https://cloud.google.com/sdk/docs/install
    pause & exit /b 1
)
echo  [OK] gcloud: %GCLOUD%
echo.

:: Set project
call %GCLOUD% config set project heroic-ruler-198603 >nul 2>&1
call %GCLOUD% config set account emeterio@tyroola.com >nul 2>&1

:: Sign in with gcloud auth login (uses Google's own verified OAuth client)
echo  Signing in... (browser will open)
echo.
call %GCLOUD% auth login --account=emeterio@tyroola.com

if errorlevel 1 (
    echo  ERROR: Login failed.
    pause & exit /b 1
)
echo.
echo  [OK] Signed in successfully.

:: Verify BQ access
echo.
echo  Verifying BigQuery access...
call %GCLOUD% auth print-access-token >nul 2>&1
if errorlevel 1 (
    echo  [WARN] Could not get token - try restarting StockMatch
) else (
    echo  [OK] Token obtained - BigQuery access ready
)

echo.
echo  ============================================================
echo   Done! Restart StockMatch and click Live BQ Refresh AU
echo  ============================================================
echo.
pause
