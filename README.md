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

## First launch

1. Make sure your network drive (e.g. `P:\PCS`) is mapped
2. Click **"Neu scannen"** to scan your PCS project folders
3. Place `PCS_Archiv_Muster.xlsx` on your Desktop, then click **"Archiv sync"**

## Troubleshooting (Windows)

If `npm install` fails on `better-sqlite3`, install build tools first:

```powershell
npm install --global windows-build-tools
```

Then retry `npm install`.
