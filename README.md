# PCS Coffee Dashboard

Local PCS dashboard for tracking coffee machine builds, samples, approbation evidence, report versions, blockers, and critical documents.

## Run

Open `index.html` directly, or serve it locally:

```sh
cd ~/Desktop/coffee-project-tracker
python3 -m http.server 8090
```

Then open `http://localhost:8090`.

For iPad access on the same Wi-Fi, use your Mac's local IP address:

```sh
ipconfig getifaddr en0
```

Then open `http://<mac-ip>:8090` on the iPad.

## Edit Data

Edit `data.js` to add more machine projects.

## Company Deployment

For a secure company version, use SharePoint/Teams document links plus company login. Do not rely on direct `P:\...` drive links as the production mechanism.

See `SECURITY_PLAN.md`.
