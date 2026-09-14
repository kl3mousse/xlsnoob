# xlsNoob

A tiny Excel add-in for macOS users who want a couple of missing quality-of-life actions without carrying a bigger toolbelt.

## Current features

- Add, search, edit, reorder, group, and open favorite workbooks
- Open the containing SharePoint or OneDrive folder for the current workbook or any saved favorite
- View a lightweight info popup with source and license details
- Run a local development ribbon backed by `https://localhost:3000`
- Build a production bundle hosted on GitHub Pages and package it as a macOS installer

## Install on your Mac

1. Quit Excel.
2. Download the latest macOS installer from [GitHub Releases](https://github.com/kl3mousse/xlsnoob/releases/latest).
3. Open the downloaded `.pkg` file and follow the installation steps.
4. Reopen Excel. The `xlsNoob` tab should appear in the ribbon.

The installer places the manifest in Excel's add-in folder. The add-in itself is served securely from GitHub Pages, so production updates usually arrive automatically.

## Architecture

This add-in uses the Office Web Add-in model:

- `manifest.xml` defines the custom Excel ribbon buttons
- the add-in loads static HTML/JS from a public HTTPS URL
- the production add-in is hosted on GitHub Pages
- local development uses a separate manifest that points to `https://localhost:3000`

## Local development

```bash
npm install
npm test
npm start
```

`npm start` generates `.dev/manifest.xml`, starts the local HTTPS server, and opens Excel with the development add-in. Look for the `xlsNoob Dev` tab in the ribbon.

- Changes to task pane HTML, TypeScript, or assets are rebuilt by the development server.
- Changes to ribbon buttons are defined by the manifest. Run `npm stop`, quit Excel, then run `npm start` again to reload them.
- Local changes are served from `https://localhost:3000` and do not affect the production GitHub Pages site.

If Excel keeps an older development ribbon after restarting, run `npm stop`, delete `7af0ae9d-6054-4726-a230-e83c7ca10cff.manifest.xml` from `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`, and run `npm start` again. Do not delete `xlsNoob.xml`; that is the installed production add-in.

## GitHub Pages deployment

This repository includes `.github/workflows/pages.yml`, which builds `dist/` and deploys it automatically to GitHub Pages on pushes to `main`.

## Publishing an installer

Update the versions in `package.json` and `manifest.xml`, validate the release, then create and push a matching tag:

```bash
npm test
npm run build
npx office-addin-manifest validate manifest.xml
git tag v1.0.0
git push origin v1.0.0
```

GitHub Actions builds `xlsNoob-1.0.0-macOS.pkg` and publishes it on the repository's Releases page.
