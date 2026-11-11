# Armodilo Tools

One site that holds the internal Armodilo tools. Each tool lives in its own folder:

| Folder | Tool | Came from |
| --- | --- | --- |
| `sales-revenue/` | Sales Revenue (shareholder dashboard) | `salesrevenue` repo |
| `monthly-summary/` | Monthly Summary | `monthlysummary` repo |
| `product-allocation/` | Product Allocation | `productallocation` repo |
| `ci-builder/` | Commercial Invoice Builder | `armo-ci-builder` repo |

`index.html` is the home page with the sign-in, and `nav.js` is the dark bar at the top of every page.

## Sign-in

There is one sign-in, on the home page. It signs in to each tool's Apps Script with the same
username and password and saves a session for every tool that accepts them. Tools that reject
the sign-in show as "No access with this sign-in". The tools' own sign-in forms are hidden;
a tool page with no session sends the visitor back to the home page. **Sign out** (on the home
page or inside any tool) signs out of everything.

For one sign-in to open every tool, the same username and password must exist in each tool's
Apps Script. The data is still protected by each Apps Script, which checks the session before
returning anything. The Invoice Builder has no sign-in.

## Updating a tool

Open the tool's folder on GitHub, choose **Add file → Upload files**, and upload the new
`index.html` (and for the Invoice Builder, `config.js`, `assets.js`, `ci-core.js`).
A freshly generated `index.html` needs one line directly after `<body>` so the top bar shows:

```html
<script src="../nav.js"></script>
```

## Adding a tool

1. Create a new folder with the tool's `index.html` (including the line above).
2. Add it to the `TOOLS` list at the top of `nav.js`.
3. Add it to the `TOOLS` list in the home `index.html`.

## Publishing

Settings → Pages → Deploy from a branch → `main` / root.
The Google Sheets / Apps Script connections are unchanged; each tool still calls the same Apps Script URL.
