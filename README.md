# NERD

NERD is the Office metadata pane for DORK. The frontend and SharePoint proxy are hosted entirely in the existing Azure Static Web App **nerd-dork**:

https://thankful-sky-0cc1b0210.6.azurestaticapps.net

## Automatic deployment

The GitHub Actions workflow deploys pushes to `main`, or a manual **Run workflow**, to the existing Azure app. It installs the locked dependencies with Node 22, builds `dist`, and uploads the frontend and `api` together. It does not provision a new Azure app.

One-time connection:

1. In Azure, open **nerd-dork → Manage deployment token** and copy the token.
2. In this repository, open **Settings → Secrets and variables → Actions → New repository secret**.
3. Name the secret `AZURE_STATIC_WEB_APPS_API_TOKEN` and paste the token as its value.
4. Open **Actions → Deploy NERD to Azure → Run workflow**, selecting `main`.

Never commit deployment tokens or paste them into chat. The secret selects the destination Azure app; use the token belonging to nerd-dork.

## Working baseline

Initial source: `NERD-Azure-SignIn-Fix.zip`, Library version 3. Metadata saves were confirmed live by changing Function to Firewall and reopening the document.

Preserve `X-NERD-SharePoint-Authorization` in the frontend and `x-nerd-sharepoint-authorization` in the API proxy. Azure Static Web Apps can replace the ordinary Authorization header. The proxy forwards the dedicated caller header as SharePoint's Authorization header.

The production build uses relative `/api` requests. `server.js` and the localhost manifest are retained only for local development and are not deployed by this workflow.

## Local build

```sh
npm ci --ignore-scripts --no-audit --no-fund
npm run build:azure
npm run manifest:azure -- https://thankful-sky-0cc1b0210.6.azurestaticapps.net
```

The centrally deployed Office manifest is managed separately. Website deployments do not automatically replace that manifest.

The fallback sign-in obtains Graph and SharePoint tokens in one Office dialog. It requests consent for both APIs, uses a refresh token only inside the dialog to obtain the second access token, and keeps any additional consent or MFA in the same window. Access tokens are cached in pane memory; refresh tokens are never sent to the pane.

Authentication regression checks: `node --test tests/auth-flow.test.cjs`.

NERD enables automatic pane opening after successfully loading a document in the DORK Documents library. The setting is saved inside that document. Existing explicit opt-outs are preserved; the pane checkbox allows turning it off or back on. Unsupported clients hide the option, and settings-save failures do not block metadata saves.

Existing documents need one manual NERD opening to be tagged. Templates must also be tagged to give newly created documents this behavior from their first opening. The existing manifest already contains the required TaskpaneId; no manifest update is needed.

Next planned changes: fix template fields so the document page reflects saved metadata, and enable automatic opening in those templates.
