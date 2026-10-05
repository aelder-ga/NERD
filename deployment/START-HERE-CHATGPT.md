# DORK — Documentation Organization & Reference Key: guided installation

Upload this file to ChatGPT and say: “Read this and help me install DORK for our organization.”

## Instructions for the assisting agent
This is an installation task, not a request merely to summarize this file. Inspect the kit's README and scripts before taking action. Tell the administrator which files, connections and execution capabilities you can actually access. Uploading this file does not give access to their computer, Microsoft tenant or Azure subscription. Request the remaining kit files when required; do not pretend to have executed commands.

Start by collecting organization name, commercial Microsoft tenant, intended SharePoint hostname, Azure subscription/resource group and deployment administrator. Ask whether this is a fresh installation. Version 0.1 supports one dedicated /sites/DORK installation per tenant; never retrofit a populated site automatically. Keep deployment progress in an installation checklist with evidence and outstanding checks.

1. Inspect the source and prerequisites. Use Node 22.14 or later in the Node 22 family and current compatible PnP PowerShell. Check Microsoft documentation for current setup commands and consent requirements.
2. Prepare a dedicated private SharePoint site at /sites/DORK with a Documents library whose URL is /Shared Documents. Review intended permissions. Obtain a separately registered provisioning app and an administrator with site and term-store rights.
3. Review and run Provision-SharePoint.ps1 only after authorization for that tenant/site. Review the schema and taxonomy with the organization before creating resources. This script is an initial scaffold, not a completed deployment.
4. Register the NERD frontend and a separate numbering service in the organization's own tenant. No RTSD registrations or credentials may be reused. Determine and document required consent, redirect URIs and selected-list grants. Use the application's actual scope requests rather than claiming narrower scopes work without testing.
5. Create an Azure Static Web App. Capture its HTTPS origin, then complete district.json from district.example.json using the provisioned Documents library ID. Configuration contains no secrets.
6. Run configure.cjs to generate a NEW directory outside the original source. Build SharePoint components for this district before building Azure. Never deploy previously compiled packages from another tenant.
7. Configure the numbering registry, Graph site ID, app identities and secret directly in Azure application settings. Do not ask for passwords, tokens or client-secret values in chat. Start with numbering disabled.
8. Create and inspect organization-neutral Word/Excel templates with the NERD content-control/defined-name contract and add-in auto-open binding. Install the newly generated Office manifest and SharePoint packages through the tenant's administrative deployment mechanisms. Upload templates and bind content types. These template tasks are not yet automated in version 0.1.
9. Create the organization's real Framework document and explicitly reserve DORK-0001 against its unique document identity. For a fresh registry verify allocation arithmetic and next-number behavior before enabling numbering. Never create fake numbered documents just to test feasibility, reset the registry, recycle IDs or blindly copy another tenant's bindings.
10. Configure Document Register views, dashboard page, Documentation Needs and navigation; test permissions and metadata persistence in Word, Excel and GEEK. Mark documents Draft until team review. Verify one save assigns one ID, repeat save retains it, rename retains identity, all metadata survives reopening, and unauthorized users cannot write.

Act autonomously on reversible preparation within the administrator's authorized scope. Present concrete settings and permission changes before asking for consequential approvals. Never make unrelated tenant changes. Report failed or untested steps honestly. Do not describe the system as ready for production until external-tenant acceptance tests pass.

Keep DORK's reusable knowledge separate from working records. Preserve the nine domains and organization-approved functions; platforms and collections are attributes. Template, Visual, Policy, Report, Register and Form may be registered through GEEK without a DORK layout. Do not introduce Runbook as a replacement for Procedure.
