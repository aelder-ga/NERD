# DORK deployment kit — 0.1 development preview

This toolkit prepares NERD, GEEK, the GEEK launcher and DORK dashboard for a separate school district. It is not yet a fully automated or externally validated installer. RTSD's live installation is unchanged.

## Included
- ChatGPT-assisted installation instructions.
- Strict, secret-free district configuration and isolated source renderer.
- SharePoint schema provisioning scaffold.
- District build entry point rebuilding both SPFx packages before Azure assets.
- Tests for configuration rejection, isolation and preserved proxy security checks.

## Constraints
Commercial Microsoft cloud, one installation per tenant, site `/sites/DORK`, Documents URL `/Shared Documents`. Runtime applications stay Azure-hosted; SharePoint holds files, metadata and extension registrations. Government/national clouds and migration of existing populated sites are unsupported in this preview.

## Prepare source
Fill out a copy of district.example.json with actual identities and resource IDs. Never put secrets in it.

```sh
node --test deployment/configure.test.cjs
node deployment/configure.cjs /path/to/district.json /path/to/new-district-source
```

In the generated directory, using Node 22:

```sh
npm ci --ignore-scripts --no-audit
node build-district.cjs
```

Build output is `dist`, managed API source is `api`, and generated Office manifest is `manifest-azure.xml`. Generated SPFx packages are in each component's `sharepoint/solution` folder and staged Azure downloads. The preview deliberately omits RTSD's deployment workflow: do not accidentally deploy to its Azure app. Set up a district-owned repository and Azure deployment connection after review.

## Outstanding before a supported release
Automated Entra registration/consent, Azure infrastructure and secret handling, neutral Word/Excel template generation and bindings, Framework bootstrap, dashboard/view provisioning, reproducible release archive and a full installation in a second tenant. Provisioning scripts require PowerShell/tenant validation; local renderer tests alone do not validate them.

Follow START-HERE-CHATGPT.md for the end-to-end checklist. Do not distribute RTSD documentation, administrative audit reports, users, registry rows or credentials with the software. Review product licensing and bundled third-party licenses before external redistribution.
