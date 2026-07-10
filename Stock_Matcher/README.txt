StockMatch - Tyre Supplier Stock Matching App
=============================================
Version: 3.0  |  Built for Tyroola

QUICK START
-----------
1. Install Node.js from https://nodejs.org (LTS version)
2. Run Install_StockMatch.bat
3. Run Setup_BigQuery_Auth.bat  (one time only)
4. Double-click "StockMatch Tyroola" on your desktop


WHAT EACH FILE DOES
-------------------
stock_matcher_app.html  The app (opens in Chrome/Edge)
server.js               Local server - handles BigQuery
Install_StockMatch.bat  Installs app + creates shortcut
StockMatch.bat          Opens the app (via shortcut)
Setup_BigQuery_Auth.bat Connect to BigQuery (one time)
Uninstall_StockMatch.bat Remove the app
.env.example            Config template (rename to .env)


HOW IT WORKS
------------
Frontend (Chrome/Edge at http://localhost:3000)
  - Upload stock file Excel
  - Display matching results
  - Download results as CSV

Backend (server.js running locally)
  - Receives requests from the browser
  - Queries BigQuery server-side
  - Returns results to the browser

BigQuery (heroic-ruler-198603)
  - suppliers_check_dashboard_table
  - Queried server-side ONLY (never from browser)


BIGQUERY AUTH (one-time setup)
------------------------------
Run Setup_BigQuery_Auth.bat OR run this in Command Prompt:

  gcloud auth application-default login

Sign in with emeterio@tyroola.com
Credentials are saved automatically.


UPDATING THE APP
----------------
Download the new stock_matcher_app.html and server.js
Copy both to: %LOCALAPPDATA%\StockMatch_Tyroola\
Restart StockMatch.


INSTALL LOCATION
----------------
%LOCALAPPDATA%\StockMatch_Tyroola\
