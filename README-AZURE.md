# NERD on Azure

This package includes the current NERD source, a prebuilt frontend in `dist`, and managed Azure Functions in `api`. It preserves the latest metadata, shared-lock, filename, and taxonomy save fixes. No client secret is required. Document templates are a separate step.

## 1. Create the hosting resource

In the Azure portal, search **Static Web Apps**, open it, and choose **Create**.

- Subscription: your active Free Trial subscription.
- Resource group: create `DORK`.
- Name: `nerd-dork`.
- Plan type: **Free**.
- Region for API/staging: **East US 2**, if available; otherwise a nearby available US region.
- Deployment source: **Other** (we will deploy the included build with Microsoft's CLI).

Review the resource and create it. Do not upgrade the trial. This deployment does not require a standalone App Service plan, VM, database, or separately created Function App.

Once created, open the Static Web App Overview and copy its HTTPS URL. Keep it for the next steps. Confirm the app belongs to the district's existing Entra directory, tenant ID `1169a3a9-3860-4a0b-80aa-410be007cde5`.

## 2. Deploy the prebuilt package

Extract this ZIP into a new folder, leaving your original NERD folder intact.

Open PowerShell in the extracted folder and run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Deploy-NERD.ps1
```

The script asks for the Static Web App URL and then its deployment token. In Azure, find **Manage deployment token** on the Static Web App Overview (or the corresponding deployment-token menu). Copy that token into the script's hidden local prompt. Do not send it in chat. The token authorizes publishing to this app; the script passes it to Microsoft's Static Web Apps CLI and removes its temporary environment value afterward. Node/npm must be available, as for the original NERD project. The CLI is downloaded from npm.

The script deploys `dist` and `api` to production and generates `manifest-azure.xml` using the real Azure URL. This is a one-time upload from your computer; afterward NERD runs on Azure with no local server.

Open `https://YOUR-AZURE-HOST/api/health`. Expected: `{"status":"ok","service":"NERD API"}`. Opening `/taskpane.html` outside Word may show an Office initialization/authentication message; that isn't an end-to-end add-in test.

## 3. Update the existing Entra app registration

Use the existing NERD app registration with client ID `f8371eb2-8758-48cf-8fa5-2f9b696d89c4`, in the district tenant.

In **Authentication**, add these redirect URIs under **Single-page application** (replace the host with the exact Azure host):

- `brk-multihub://YOUR-AZURE-HOST`
- `https://YOUR-AZURE-HOST/taskpane.html`

Keep the existing localhost entries during the migration. Do not create a new client secret or enable implicit grants. Existing delegated Graph and SharePoint permissions are used unchanged. The manifest ID and Entra client ID are different identifiers; don't substitute one for the other.

## 4. Deploy the Office add-in

Use the generated `manifest-azure.xml`, not the original development manifest. Deploy initially to Abraham's account through Microsoft 365 admin center **Settings > Integrated apps > Upload custom apps**, choosing the Office add-in XML path offered in your tenant. Verify the pane loads in a SharePoint-hosted document and saves metadata and a rename before wider rollout.

The manifest has `Office.AutoShowTaskpaneWithDocument`. Automatic first opening from a new SharePoint document also requires the template to contain the add-in's document extension parts. The uploaded Standard template did not have those parts; hosting alone does not finish that requirement. We will prepare the templates after this hosted add-in is verified.

## Costs and trial

Free hosting plus managed consumption APIs is the intended configuration. Exact costs remain unmeasured. Trial credit is not a permanent free-service entitlement. Record usage and costs under this resource group; compare them with ongoing pricing before upgrading or moving to reseller billing. Trial services stop at expiry/credit exhaustion unless upgraded.

## Build future changes

After editing source, use a supported current Node release compatible with the locked dependencies (the included frontend was built with Node 24):

```text
npm ci
npm run build:azure
```

Then run the deployment script again. The Azure Functions runtime is Node 22. The original development build, manifest, and local `server.js` are retained for local development; they are excluded from the deployed frontend.

## Validation performed

- Production webpack build succeeded. Webpack reported advisory bundle-size warnings.
- API tests covered missing authorization, wrong tenant/site, path traversal, unsupported methods, client-header filtering, redirect blocking, and success/error response compatibility.
- Manifest generation checked as valid XML with the real-origin substitution pattern.
- The deployed Azure runtime, district sign-in, Word web behavior, and actual SharePoint save/rename still require verification after deployment.
