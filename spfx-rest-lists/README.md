# spfx-rest-lists

SPFx React web part (`RestLists`) that lists all lists in the current SharePoint site using SharePoint REST API via `this.context.spHttpClient`.

## API used

`/_api/web/lists?$select=Title,Id,BaseTemplate,Hidden,ItemCount&$orderby=Title`

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

`sharepoint/solution/spfx-rest-lists.sppkg`

## Deploy

1. Upload `.sppkg` to tenant app catalog.
2. Deploy the app.
3. Add **RestLists** web part to a modern SharePoint page.

## Notes

- This solution does **not** request Microsoft Graph permissions.
- Output shows: `Title`, `ItemCount`, `Hidden`, `BaseTemplate`.
