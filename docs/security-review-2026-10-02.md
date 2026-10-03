# NERD / GEEK / DORK security review

Review started 2 October 2026; verification continued 3 October UTC. Baseline: `4bf375cb8d6527db517b4508734c21efc7aa7a20`.

## Scope and evidence

- Reviewed current frontend, Azure Functions, numbering authorization, standalone taxonomy editor, SharePoint dashboard and launcher, deployment workflows, and dependency manifests.
- Inspected the main branch's 41 commits and scanned 245 unique text blobs (about 7.25 MB) from 145 historical paths for common credential patterns. Five matches were synthetic test fixtures; no real credential was identified by that scan. Seven historical binary PNG blobs were unavailable through the connector. This custom scan is supplemented by a full-history Gitleaks workflow; its result must be checked separately.
- Reviewed packaged SharePoint component source and newly generated manifests. This is not a forensic scan of every tenant file, every GitHub ref, or all Microsoft/vendor SDK code.
- All 93 application tests passed, including rejection tests for foreign Graph destinations, hostile upload destinations, repeated paging, non-metadata SharePoint operations, unsupported query options, missing tokens, and redirect handling. Canonical metadata routes preserve `$select`/`$expand` values after encoding.
- Azure production build passed. Both SharePoint Framework 1.23.2 packages built with Node 22, including the dashboard's existing test suite.
- `npm audit --omit=dev` reported zero known vulnerabilities in all three projects: root, dashboard, and launcher. Full development audits still report advisories; they are not clean.
- The existing live DORK dashboard loaded as the regular district account. Its Site access panel displayed Abraham Elder as owner and no listed site members or visitors. This limited panel observation is not proof that all individual item permissions or external links are restricted.

## Fixes

| Area | Change |
| --- | --- |
| Graph requests | Validate the HTTPS Graph v1.0 destination before obtaining or attaching a delegated token; reject redirects and repeated/unbounded pagination. |
| Uploads | Validate Microsoft HTTPS upload-session destinations; reject redirects; retain the no-bearer upload behavior. |
| SharePoint proxy | Restrict the bridge to the configured DORK library's metadata operations and user resolution; reject `$batch`, arbitrary endpoints, non-JSON content types, and unsupported methods. Preserve the dedicated authorization header. |
| Local development proxy | Reuse the same Azure handler and bind to loopback instead of maintaining a weaker second proxy. |
| Browser protections | Add Content Security Policy, stricter standalone GEEK/terms policies, no-store on authentication pages, and cache-busting frontend script URLs. Preserve Office embedding and authentication compatibility. |
| Supply chain | Pin checkout, Node setup, Azure deployment, Gitleaks, and CodeQL actions to verified commits; avoid persisting checkout credentials. |
| Dependencies | Upgrade Office tooling and SharePoint Framework to supported versions, refresh locks, and build replacement SharePoint packages. |
| Ongoing checks | Add push/PR/weekly secret scanning, CodeQL security-extended queries, production dependency audit, regression tests, and weekly Dependabot updates. |

## Existing controls verified in source

The numbering API checks the delegated caller's edit permission on the specific document before using application credentials for the registry. Registry credentials are not sent to the browser. The proxy forwards only the dedicated SharePoint header and rejects foreign origins and redirects. Auth uses PKCE and state/nonce validation. Dynamic document and dashboard values are escaped or rendered as text, and Excel values are protected against formula interpretation. The launcher requests no additional API permissions.

## Remaining work and limits

1. **Tenant least privilege:** current delegated scopes include Graph `Sites.ReadWrite.All`, SharePoint `AllSites.Write`, and `TermStore.ReadWrite.All`. Delegated user permissions still apply, but these grants are broader than DORK. A selected-resource migration needs real Entra and SharePoint grant verification and functional testing; it was not silently applied in this review.
2. **Classification and sharing:** `Sensitive` is a label, not enforced file security. Review library/item permissions, organization/anonymous sharing links, external sharing, and taxonomy curator permissions. No actual data exposure was established by this review.
3. **Development tooling advisories:** upstream `braces` and `node-forge` currently have no patched npm release. Office development/debugging tooling also retains older nested `adm-zip`; SPFx development tooling retains older `qs`/`uuid` paths. These are outside the audited production dependency sets, but build machines still need trustworthy inputs and restricted credentials. Do not use `npm audit fix --force` to substitute incompatible toolchains merely to clear counters.
4. **Office SDK compatibility:** Office pages retain CSP `unsafe-eval` for SDK compatibility. Standalone GEEK and terms pages use a stricter script policy. This review does not establish that all external Microsoft SDKs can operate under a no-eval policy.
5. **Administrative settings:** complete verification of Azure secret rotation/expiration, Entra consents and application grants, Conditional Access, SharePoint external sharing, GitHub branch protection, and tenant-wide audit/alert configuration needs administrator access. No claim is made that these are configured correctly.
6. **Deployment:** Azure changes deploy through the existing workflow. Upgraded dashboard package `1.0.4.0` and launcher `1.0.2.0` must also replace the installed SharePoint app-catalog packages before their framework upgrade is considered live. Their rebuilt assets remain hosted exclusively in the existing Azure Static Web App.

No destructive live tests, document metadata changes, numbering allocations, taxonomy additions, permission changes, or secret rotation were used as part of this review. Check the security workflow and deployment results for the exact remediation commit before treating this change as deployed.

## Final verification

Application remediation commit: `86d19c95574a20a38c1a9674216c6cb65ee57561`.

- [Azure deployment](https://github.com/aelder-ga/NERD/actions/runs/37087294566) completed successfully.
- [Security review](https://github.com/aelder-ga/NERD/actions/runs/37087294595) passed all three jobs. Full-history Gitleaks scanned 63 reachable commits with no leaks found. CodeQL security-extended produced no findings; no finding suppression was added. Its initial request-forgery finding prompted reconstruction of outbound proxy operations from fixed routes and encoded values. The security check reads SARIF driver and extension metadata and rejects actionable findings.
- Production audits passed for all three projects, and all 93 application tests passed in CI.
- Live GEEK responses supplied the configured CSP, no-store, nosniff, no-referrer and HSTS headers. The cache-busted frontend loaded. GEEK authenticated as the regular account and listed all four existing numbered files; Collections & Systems authenticated and read the existing ten collections under its stricter policy.
- Live SharePoint and numbering APIs returned 401 without a token. The live proxy rejected foreign destinations and `$batch` with 403, and a permitted metadata GET with a deliberately invalid token returned 401. These checks did not write document data.
- SharePoint's app detail page still lists dashboard catalog version `1.0.3.0`; replacement packages `1.0.4.0` (dashboard) and `1.0.2.0` (launcher) are built and available in the existing Azure site's `/downloads/` folder. This review does not claim the catalog replacements are installed.

The application fixes are deployed; administrative verification and app-catalog replacement remain open. The public repository contains no document body data added by this review. A clean scan does not certify the tenant or eliminate the remaining concerns above.
