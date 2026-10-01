# Upload with GEEK — SharePoint launcher

This SPFx 1.22.2 command set opens the Azure-hosted GEEK page in a new tab. It is visible only in DORK Documents (`dd8ac8ef-3be9-4fb8-b7cb-6ac0f56d2b03`) when no rows are selected. It makes no API calls, requests no additional API permissions, and does not allocate numbers. It leaves SharePoint's Create or Upload button intact.

The feature registers a command-bar extension on document libraries (template 101), with runtime visibility restricted to the exact DORK library. Install the app only on the DORK site. There is no tenant-wide instance registration, custom script enablement, or domain isolation.

## Installation

1. Download `https://thankful-sky-0cc1b0210.6.azurestaticapps.net/downloads/geek-launcher.sppkg` after the Azure deployment succeeds.
2. Open the SharePoint admin center with the administrative account. Under More features > Apps, open the app catalog / Manage apps. Upload the package and enable/deploy it. Review the external asset domain: `thankful-sky-0cc1b0210.6.azurestaticapps.net`. The package contains no API permission requests.
3. On the DORK site, use Site contents > New > App / Add an app and add geek-launcher-client-side-solution from your organization. UI labels can vary between app catalog experiences.
4. Refresh Documents with no file selected. Select Upload with GEEK; it should open geek.html in a new tab. Browsing does not consume a document number.

Do not check a tenant-wide deployment option or add this app to unrelated sites. If the app catalog is not yet available, stop at that point and establish it through the supported SharePoint admin flow. NoScript does not need changing (`requiresCustomScript: false`). The package is a launcher, not another Entra application.

## Build and hosting

Use Node 22.14–22.x, then `npm ci` and `npm run build` in this directory. Run `node scripts/stage-geek-launcher.cjs` from the repository root afterward. It stages release assets and the base64-encoded install package for Azure deployment. The root webpack build decodes that package into `/downloads/geek-launcher.sppkg` and copies launcher assets under `/sharepoint/geek-launcher/`. Runtime code and manifests stay Azure-hosted; SharePoint holds only the installed extension registration/package.

Compiled assets are checked in so the ordinary NERD deployment does not install a second build toolchain. For extension changes, increment both solution and component versions, rebuild, stage, and update the app catalog package. GEEK page changes require only the existing GitHub deployment.

## Verification

SPFx production compilation, lint, and package validation passed. Root regression tests execute the actual staged AMD bundle with a mock SharePoint host to verify library scope, row-selection visibility, safe navigation, and event cleanup. The package manifest points to Azure and contains no localhost URL or embedded runtime assets. Tenant installation and button rendering remain pending until the package is installed in DORK.
