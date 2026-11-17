# Landing CTA Hero (`landing-cta`)

A custom SharePoint Framework (SPFx `1.24.1`) **web part**—a reusable, no-code-editable building block for SharePoint Online Communication Sites.

## Why Has This Been Built

SharePoint's built-in **Editorial Card** and standard sections cannot stretch edge-to-edge across the browser, and native full-width web parts (*Image*, *Hero*) lack combined control over a responsive heading, subtitle, contrast overlay, and customizable CTA button. This component enables true full-width placement (`supportsFullBleed: true`) inside a SharePoint **Full-width section** while keeping normal SharePoint navigation and page editing intact.

## Features

- **Full-width placement:** Supports SharePoint **Full-width sections** (`supportsFullBleed: true`) for edge-to-edge rendering.
- **Editable content:** Heading, multiline subheading, CTA button text, validated `https://` or site-relative link, and new-tab toggle.
- **Background & overlay:** Built-in image picker (browse **Site Assets** / **Documents**, upload to **Site Assets**, or paste a URL), fallback hex background color, and adjustable dark overlay opacity (`0–90%`).
- **Layout & typography:** Hero height presets (`viewport`, `tall`, `medium`, `compact`), left/center alignment, and configurable font sizes, weights, italics, colors, and button padding/radius.
- **Safety:** Rejects `javascript:`, `data:`, `http:`, and protocol-relative (`//`) URLs.

## Prerequisites & Packages

- **Node.js:** `v22.14+` (`22.x`), `24.x`, or `26.x` (see `engines` in `package.json`).
- **Global packages:** None required. All build, lint, test, and packaging tools (`@rushstack/heft`, TypeScript, ESLint, Jest, SPFx build rig) are installed locally via `npm install`.

## Local Testing

```zsh
cd landing-cta
npm install
npx heft trust-dev-cert   # one-time: trusts the local localhost dev certificate
export SPFX_SERVE_TENANT_DOMAIN="<your-tenant>.sharepoint.com"
npm start
```

1. When SharePoint opens, click **Load debug scripts**.
2. Edit a test page, add a **Full-width section**, and insert **Landing CTA Hero**.

## Build & Admin Deployment

Build, test, and package the `.sppkg`:

```zsh
cd landing-cta
npm run build
```

This outputs `sharepoint/solution/landing-cta.sppkg` with all client-side assets bundled (`includeClientSideAssets: true`).

### SharePoint Admin Steps

1. Open the **SharePoint Admin Center** → **More features** → **Apps** → **Manage apps** (or a site's **Apps for SharePoint** if a Site Collection App Catalog is enabled).
2. Upload `sharepoint/solution/landing-cta.sppkg`.
3. Enable the app (select **Only enable this app** to scope it to specific sites, then add it via **Site contents → New → App** on the target site).
4. On the target Communication Site page, place **Landing CTA Hero** inside a **Full-width section** and **Republish**.
