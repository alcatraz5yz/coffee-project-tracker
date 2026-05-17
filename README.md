# PCS Kaffee Dashboard

Local dashboard for tracking coffee machine builds, samples, approbation evidence, certifications, and critical documents.

## Requirements

- **Node.js** (v18+) — [nodejs.org](https://nodejs.org)
- **Git**

No Python needed.

## Setup

```sh
git clone https://github.com/alcatraz5yz/coffee-project-tracker.git
cd coffee-project-tracker
npm install
node server.js
```

Then open **http://localhost:8090** in your browser.

## Branches

| Branch | Platform |
|--------|----------|
| `main` | macOS |
| `windows` | Windows 11 |

Use `git checkout windows` on a Windows machine.

## First launch

1. Click **"Neu scannen"** to scan your PCS project folders
2. Click **"Archiv sync"** to load archive locations from `PCS_Archiv_Muster.xlsx` (place it on your Desktop)
3. Make sure your network drive (e.g. `P:\PCS`) is mapped before scanning

## Troubleshooting (Windows)

If `npm install` fails on `better-sqlite3`, install build tools first:

```powershell
npm install --global windows-build-tools
```

Then retry `npm install`.
