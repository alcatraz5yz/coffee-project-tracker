# PCS Kaffee Dashboard

Local dashboard for tracking coffee machine builds, samples, approbation evidence, certifications, and critical documents.

## Requirements

- **Node.js** (v18+) — [nodejs.org](https://nodejs.org)
- **Git**

No Python needed.

## Setup — macOS

```sh
git clone https://github.com/alcatraz5yz/coffee-project-tracker.git
cd coffee-project-tracker
npm install
node server.js
```

## Setup — Windows 11

```powershell
git clone https://github.com/alcatraz5yz/coffee-project-tracker.git
cd coffee-project-tracker
git checkout windows
npm install
node server.js
```

Then open **http://localhost:8090** in your browser.

## First launch (Windows)

### 1. Map the network drive

Before scanning, the PCS project folder must be mapped as a network drive:

1. Open **File Explorer**
2. Right-click **"This PC"** → **"Map network drive"**
3. Choose drive letter `P:`
4. Enter the network path (ask IT if unsure), e.g. `\\server\PCS`
5. Check **"Reconnect at sign-in"**
6. Click **Finish**

### 2. Place the archive Excel on your Desktop

Copy `PCS_Archiv_Muster.xlsx` to your Desktop (`C:\Users\<you>\Desktop\`).

### 3. First scan

1. Click **"Neu scannen"** — scans all project folders on the network drive
2. Click **"Archiv sync"** — loads archive locations from the Excel file

> The scanner only reads file/folder names and a few Excel cells. Nothing is uploaded or modified.

## Troubleshooting (Windows)

If `npm install` fails on `better-sqlite3`, install build tools first:

```powershell
npm install --global windows-build-tools
```

Then retry `npm install`.
