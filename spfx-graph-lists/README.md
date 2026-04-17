# spfx-graph-lists

SPFx React web part (`GraphLists`) that lists all lists in the current SharePoint site using Microsoft Graph via `this.context.msGraphClientFactory.getClient('3')`.

## API flow used

1. Resolve current site ID from hostname + current web path:
   - `GET /sites/{hostname}:{current-web-path}`
2. List lists:
   - `GET /sites/{site-id}/lists?$select=id,displayName,list`

## Prerequisites

- Node.js version supported by this solution (`>=22.14.0 <23.0.0`)
- SharePoint Online tenant + app catalog

## Build and run locally

```bash
npm install
gulp build
gulp serve
```

## Package

```bash
gulp build
```

Generated package:

`sharepoint/solution/spfx-graph-lists.sppkg`

## Deploy

1. Upload `.sppkg` to tenant app catalog.
2. Deploy the app.
3. Add **GraphLists** web part to a modern SharePoint page.

## Approve Microsoft Graph API permission

This solution requests:

- Resource: `Microsoft Graph`
- Scope: `Sites.Read.All`

Approval steps:

1. Deploy the solution package to app catalog.
2. Open **SharePoint Admin Center** -> **Advanced** -> **API access**.
3. Find pending request for `Microsoft Graph - Sites.Read.All`.
4. Approve the request.
5. Re-open the SharePoint page and use the **GraphLists** web part.

## Notes

- Output shows: `displayName` and `template` (from `list.template` when available).
