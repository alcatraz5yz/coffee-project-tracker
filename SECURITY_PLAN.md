# Secure PCS Dashboard Deployment Plan

## Recommendation

Use Microsoft 365 as the approval-friendly path:

- Host documents in SharePoint or Teams document libraries.
- Host the PCS dashboard as an internal web app behind company login.
- Use Microsoft Entra ID for authentication.
- Keep access controlled by existing SharePoint groups.
- Store tracker data in a real backend later, not in `data.js`.

This is safer and easier to approve than exposing a mapped drive such as `P:\PCS\1157` directly to a website.

## Why Not Direct P: Drive Links

Browser support for `P:\...` or `file://` links is inconsistent and often blocked by security policy. It also makes access control harder to audit from the web app.

The P drive can still remain the working source during transition, but production links should become SharePoint or internal HTTPS links.

## Current Prototype

The PCS prototype has three document concepts:

- `localDocumentRoot`: local Mac demo path through `evidence-1157/`.
- `companyDriveRoot`: visible reference path, currently `P:\PCS\1157\`.
- `sharePointRoot`: production target base URL.

Switch production behavior in `config.js`:

```js
const trackerConfig = {
  documentMode: "sharepoint",
  localDocumentRoot: "evidence-1157/",
  companyDriveRoot: "P:\\PCS\\1157\\",
  sharePointRoot: "https://company.sharepoint.com/sites/coffee-projects/Shared%20Documents/PCS/1157/"
};
```

## Production Checklist

- Move or sync `P:\PCS\1157` documents into a controlled SharePoint library.
- Replace the placeholder SharePoint URL in `config.js`.
- Restrict site access to the project/certification groups.
- Use read-only access for most users and edit rights only for owners.
- Keep version history enabled for evidence documents.
- Add audit logging through Microsoft 365.
- Move status edits from browser `localStorage` into a database with user/time history.

## Best First Company Pilot

Start with EF1157 only:

- One SharePoint document library folder: `PCS/1157`.
- One internal tracker page.
- One Approbation checklist.
- Read access for viewers, edit access for Approbation/PCS owners.

After that works, add EF1234 and the rest of the machine projects.
