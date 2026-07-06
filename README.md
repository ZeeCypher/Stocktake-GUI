# Stocktake App

This is a simple stocktake app for IT inventory.

It lets people check stock, update quantities, retire items, and export a CSV. Managers can unlock extra tools to add, edit, delete, and restore items.

## Files

| File | What it does |
| --- | --- |
| `server.js` | Runs the app and saves the data |
| `index.html` | The web page users see |
| `stocktake-data.json` | The database file |
| `logo.png` | The logo at the top of the page |

## Before Publishing To GitHub

Do not commit local secrets, certificates, generated reports, or real stocktake data.

This repo ignores:

- `run-stocktake-report.sh` and backup copies of that script
- `Certificate/` and certificate/private key files such as `.pfx`, `.p12`, `.pem`, `.key`, `.cer`, and `.crt`
- `reports/` and log files
- `stocktake-data.json`
- `.env` files

If any of those files were already committed, remove them from Git without deleting your local copies:

```bash
git rm --cached -r Certificate reports stocktake-data.json
git rm --cached run-stocktake-report.sh
```

Then commit the removal and the `.gitignore` update.

If this repository was ever pushed with a real certificate password, PFX file, tenant/client details, sender mailbox, or recipient list, treat those values as exposed. Revoke the old certificate credential in Entra ID, create a new certificate, and use a new PFX password before making the repo public.

## How To Run

Open PowerShell in this folder:

```powershell
cd "C:\Users\ZaidAlshorogy\Documents\Stocktake GUI"
node server.js
```

Or use the npm command:

```powershell
npm start
```

Then open:

```text
http://localhost:3000
```

If port `3000` is busy:

```powershell
$env:PORT=3001; node server.js
```

Then open:

```text
http://localhost:3001
```

## PINs

Main app PIN:

```text
1234
```

Manager PIN:

```text
4321
```

The main PIN opens the page.

The manager PIN unlocks add, edit, delete, and restore.

To change the PINs without editing the code:

```powershell
$env:STOCKTAKE_PIN="1111"
$env:STOCKTAKE_MANAGER_PIN="9999"
node server.js
```

For production, set `STOCKTAKE_PIN` and `STOCKTAKE_MANAGER_PIN` on the server instead of relying on the example defaults in the code or README.

## Where The Data Saves

All data saves in:

```text
stocktake-data.json
```

That file stores:

- Current stock items
- Deleted items bin
- Quantities
- Locations
- Notes
- Movement history
- Who changed something
- When it changed

Back up `stocktake-data.json` if you want to keep the stock data safe.

## Monthly Email Report

The app can create a polished monthly stocktake report and email it from the Linux VM on the 4th of every month.

The monthly email includes:

- A short summary in the email body
- Summary tiles for active items, low stock, out of stock, retired items, and categories
- A styled Excel workbook with filters, frozen headings, column widths, and colour-coded stock status
- Excel attachments saved in `reports/`

Use:

- `Pipelines/scripts/Send-StocktakeReport-MgGraph.ps1`
- `run-stocktake-report.sh`
- `install-monthly-stocktake-cron.sh`

The script:

- Reads `stocktake-data.json`
- Builds a styled Excel workbook in `reports/`
- Authenticates to Microsoft Graph with the certificate PFX on the VM
- Sends the report from the shared mailbox

To install the monthly schedule, run this on the Linux VM:

```bash
cd /home/zaialsadmin/StocktakeGUI
chmod +x install-monthly-stocktake-cron.sh
./install-monthly-stocktake-cron.sh
```

The installer schedules the report for 8:00 AM Sydney time on the 4th of every month:

```text
0 * * * * [ "$(TZ=Australia/Sydney date +\%d\%H\%M)" = "040800" ] && /home/zaialsadmin/StocktakeGUI/run-stocktake-report.sh >> /home/zaialsadmin/StocktakeGUI/reports/monthly-stocktake-cron.log 2>&1 # stocktake-monthly-report
```

To use a different local day/time, set `LOCAL_DAY_HOUR_MINUTE` before running the installer:

```bash
LOCAL_DAY_HOUR_MINUTE="041430" ./install-monthly-stocktake-cron.sh
```

Edit the VM run script:

```text
run-stocktake-report.sh
```

Set:

```text
-TenantID
-ClientID
-CertificateFilePath
-CertificatePassword
-Sender
-Recipient
```

Keep that file local only. It contains environment-specific values and may contain a certificate password.

Example local VM script:

```bash
#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"
mkdir -p reports

pwsh -NoProfile -File "./Pipelines/scripts/Send-StocktakeReport-MgGraph.ps1" \
  -TenantID "your-tenant-id" \
  -ClientID "your-app-client-id" \
  -CertificateFilePath "/home/your-user/StocktakeGUI/Certificate/StocktakeReportMailer.pfx" \
  -CertificatePassword "your-local-pfx-password" \
  -Sender "stocktake-shared-mailbox@example.com" \
  -Recipient "manager@example.com"
```

Create it on the VM:

```bash
nano run-stocktake-report.sh
chmod 700 run-stocktake-report.sh
```

The script is ignored by Git.

## Certificate Setup

The monthly email report uses Microsoft Graph application authentication with a certificate. Only the public certificate should be uploaded to Entra ID. The private PFX must stay on the VM and must not be committed to GitHub.

On the VM, create a local certificate folder:

```bash
mkdir -p Certificate
chmod 700 Certificate
```

Create a self-signed certificate and export both files:

```bash
openssl req -x509 -newkey rsa:2048 -sha256 -days 730 -nodes \
  -keyout Certificate/StocktakeReportMailer.key \
  -out Certificate/StocktakeReportMailer.cer \
  -subj "/CN=StocktakeReportMailer"

openssl pkcs12 -export \
  -out Certificate/StocktakeReportMailer.pfx \
  -inkey Certificate/StocktakeReportMailer.key \
  -in Certificate/StocktakeReportMailer.cer

chmod 600 Certificate/StocktakeReportMailer.*
```

When prompted during the PFX export, choose a strong password. Put that password only in the local `run-stocktake-report.sh` file or a secure secret store.

In Microsoft Entra ID:

1. Open **App registrations**.
2. Open the app used for the stocktake mailer, or create a new one.
3. Copy the **Application (client) ID** and **Directory (tenant) ID** into your local run script.
4. Go to **Certificates & secrets**.
5. Upload `Certificate/StocktakeReportMailer.cer`.
6. Go to **API permissions**.
7. Add Microsoft Graph **Application** permission `Mail.Send`.
8. Grant admin consent.
9. Restrict the app so it can only send as the stocktake shared mailbox.

Copy only the local files to the VM. Do not push them:

```text
Certificate/StocktakeReportMailer.pfx
Certificate/StocktakeReportMailer.cer
Certificate/StocktakeReportMailer.key
run-stocktake-report.sh
```

The Azure app needs Microsoft Graph **Application** permission:

- `Mail.Send`

Grant admin consent and restrict the app so it can only send from the stocktake shared mailbox.

## What Normal Users Can Do

Normal users can:

- View stock
- Search and filter stock
- Add or remove 1 from quantity
- Retire or restore an item
- Export a CSV

## What Managers Can Do

Managers can also:

- Add new items
- Edit items
- Move items to the deleted items bin
- Restore deleted items from the bin

## Simple Explanation Of `server.js`

`server.js` is the backend. It runs on Node.js.

The backend is the part that:

- Opens the website
- Checks the PIN
- Reads the database file
- Saves changes to the database file
- Sends data to the browser

Important code in `server.js`:

| Code | Beginner explanation |
| --- | --- |
| `const http = require("node:http")` | Loads Node's web server tool |
| `const fs = require("node:fs")` | Loads Node's file tool, used to read and write the database |
| `const port = Number(process.env.PORT || 3000)` | Uses port 3000 unless another port is set |
| `const appPin = ...` | Stores the main PIN |
| `const managerPin = ...` | Stores the manager PIN |
| `const databaseFile = ...` | Points to `stocktake-data.json` |
| `defaultDatabase()` | Creates an empty database structure |
| `readDatabase()` | Opens `stocktake-data.json` and reads the stock data |
| `writeDatabase()` | Saves changes back into `stocktake-data.json` |
| `addLog()` | Adds a line to the movement history |
| `sanitizeItem()` | Cleans item data before saving it |
| `handleAction()` | Handles changes like save, delete, restore, plus, minus, and retire |
| `isAuthenticated()` | Checks if the user entered the main PIN |
| `isManagerAuthenticated()` | Checks if the manager PIN is unlocked |
| `sendJson()` | Sends data back to the browser |
| `server.listen()` | Starts the app so people can open it in a browser |

## Server Routes

Routes are the web addresses the browser talks to.

| Route | What it does |
| --- | --- |
| `/` | Opens the app page |
| `/api/login` | Checks the main PIN |
| `/api/logout` | Logs out |
| `/api/state` | Sends the latest stock data to the page |
| `/api/action` | Saves stock changes |
| `/api/manager-login` | Checks the manager PIN |
| `/api/manager-logout` | Locks manager tools again |

## Simple Explanation Of `index.html`

`index.html` is the frontend. It is the page people see and click.

It contains:

- HTML for the page layout
- CSS for the colours and design
- JavaScript for buttons and live updates

Important code in `index.html`:

| Code | Beginner explanation |
| --- | --- |
| `<style>` | Contains the page design and colours |
| `<header>` | The top part with the logo, user name, and buttons |
| `<table>` | Shows the stock items |
| `<aside>` | Shows the manager panel, movement log, and deleted items bin |
| `<script>` | Contains the browser JavaScript |
| `let items = []` | Stores the current stock list in the browser |
| `let deletedItems = []` | Stores the deleted bin list in the browser |
| `let managerUnlocked = false` | Tracks whether manager tools are unlocked |
| `const els = {...}` | Keeps references to page elements like buttons and inputs |
| `startApp()` | Starts the page and loads data from the server |
| `refreshState()` | Gets the latest database data from the server |
| `render()` | Redraws the page after data changes |
| `renderTable()` | Draws the stock table |
| `renderLog()` | Draws the movement history |
| `renderBin()` | Draws the deleted items bin |
| `unlockManager()` | Sends the manager PIN to the server |
| `saveItem()` | Sends a new or edited item to the server |
| `handleRowAction()` | Handles row buttons like plus, minus, edit, retire, and delete |
| `handleBinAction()` | Handles restore from the deleted items bin |
| `apiAction()` | Sends stock changes to `/api/action` |
| `getActor()` | Gets the user name from the top of the page |

## How The App Works

The basic flow is:

1. Someone opens the app in the browser.
2. `server.js` shows the PIN screen.
3. After the correct PIN, the browser loads `index.html`.
4. `index.html` asks `/api/state` for the latest stock data.
5. The user clicks a button or saves an item.
6. The browser sends the change to `/api/action`.
7. `server.js` updates `stocktake-data.json`.
8. The page refreshes and shows the latest data.

In simple words: the browser is the screen, the server is the brain, and `stocktake-data.json` is the memory.

## Deleted Items Bin

Delete does not completely destroy the item.

When a manager deletes an item:

- The item is removed from the active stock table.
- The item is moved into `deletedItems` inside `stocktake-data.json`.
- The app records who deleted it.
- The app records when it was deleted.

A manager can restore the item from the bin.

## Movement Log

The movement log records stock changes.

It records:

- What changed
- The item name
- The user name typed at the top
- The IP address that made the change
- The time of the change

The user name is typed by the person using the page. It is useful for tracking, but it is not a secure login account.
The IP address is detected by the server and shown as read-only at the top of the page.

## Logo

The logo file is:

```text
logo.png
```

To change the logo, replace `logo.png` with another image using the same file name.

The logo line in `index.html` is:

```html
<img class="brandLogo" src="logo.png" alt="Company logo">
```

The logo size is controlled by `.brandLogo` in the CSS.

## If Someone Asks What This Is

You can say:

```text
This is a small internal stocktake app. The page is built in index.html, the Node.js server is server.js, and the shared data is saved in stocktake-data.json. Users enter a PIN to open the page, and managers use a second PIN to add, edit, delete, or restore stock items. All changes are saved to the JSON database and shown in the movement log.
```
