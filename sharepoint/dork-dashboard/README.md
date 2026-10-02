# DORK Dashboard

A read-only SPFx web part for the DORK Home page. JavaScript assets are hosted in the existing Azure Static Web App. SharePoint REST calls use SPHttpClient and the current user's site access. No application permissions, secrets, new Entra registration or popup authentication are requested.

## What it shows

- Lifecycle counts for numbered documents; selecting a bar filters the register.
- Review health for Active documents only: past review date, due today / approaching within 30 days, good standing beyond 30 days, and missing or invalid review date. Other lifecycles are excluded. Dates are compared as calendar dates rather than shifted by browser timezone.
- A searchable classification reference sourced from `docs/drafts/tldr-search-index.md` and `classification-index.md`. Search ranks lexical matches; it does not use AI, assign metadata, or assume a product name determines a classification. Unknown terms need editorial additions to `classification.ts`.
- A compact selectable calendar of Active review dates and open Needs target dates. Completed and Cancelled Needs are excluded.
- Open Needs ranked by priority and target date. New remains awaiting discussion; Planned means agreed work.
- A document register with full selectable, copyable URLs. Copy uses the clipboard API when available and falls back to selecting the URL for Ctrl+C / Command+C. URLs grant no access and should be recopied after a rename or move.

The loader follows SharePoint pagination, resolves columns by display name and reports failures rather than displaying false healthy counts. Charts reflect only documents the signed-in user can access. This component never allocates numbers, updates files, writes metadata or changes Needs statuses. GEEK and NERD remain independent.

## Build

Node 22.14 or later within Node 22.

```
npm ci
npm run build
cd ../..
node scripts/stage-dork-dashboard.cjs
npm run build:azure
```

Commit the source, locked dependencies, staged Azure assets and base64 install package. The existing root deployment workflow publishes the assets and decodes the download package. Rebuild and restage whenever dashboard source changes; increment package-solution version for installed upgrades. Hashed previous bundles can remain available so installed cached manifests continue to load.

## Install on DORK

After Azure deployment succeeds, download `/downloads/dork-dashboard.sppkg` from the existing Static Web App. Upload the package to the SharePoint app catalog, enable it and add **DORK Dashboard** to the DORK site. The package is site installed (not tenant-wide) and requests no new API permissions.

Edit Home, add the **DORK Dashboard** web part in a wide one-column section and save as a draft. Validate live counts, date selection, search and copy links before removing duplicate native Register/Needs/Calendar sections or publishing. Do not change the existing NERD or GEEK registrations or authorization headers.

## Checks completed / still required

Nine local tests cover date boundaries, missing dates, unexpected lifecycle values, term search, encoded filenames, pagination, loading failures, REST mapping and component interactions. Root NERD/GEEK regression tests also run unchanged. Production build and package generation pass.

Installed and enabled on DORK on 2026-10-02. Home remains a saved draft. Live checks verified four Draft documents, five open Needs, taxonomy labels, people names, classification search, chart filtering, calendar date selection and clipboard copy (including an ampersand in the filename). Duplicate native dashboard sections were removed from the draft.

Review health has no live Active documents yet; its date boundaries and missing-date handling are covered by local tests. Access is inherited from the current user’s SharePoint requests; comparison with a lower-access account and an actual phone remain acceptance checks. Responsive layout uses the dashboard container width as well as viewport media queries so narrow SharePoint columns can stack safely.
