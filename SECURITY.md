# Security

NERD, GEEK, and the DORK dashboard are internal district tools hosted in Azure and SharePoint. Public frontend files are not an authorization boundary: Microsoft validates delegated access to documents, and the numbering service separately checks the caller's edit permission before using its registry credentials.

## Reporting a vulnerability

Contact the district's IT owner privately. Do not put credentials, access tokens, document content, student information, or exploit details in a public GitHub issue. This repository is public.

## Controls

- Keep application secrets in Azure configuration and deployment credentials in GitHub Actions secrets. Never embed them in frontend code or packages.
- Preserve `X-NERD-SharePoint-Authorization` in both frontend and API. The proxy ignores the ambient `Authorization` header, restricts destinations and metadata operations, and rejects redirects.
- Graph bearer tokens are only sent to Microsoft Graph v1.0. Upload sessions must use supported Microsoft HTTPS destinations and do not receive Graph bearer tokens.
- Treat all metadata as untrusted. Render text safely; escape XML and HTML; write Excel metadata as text rather than formulas.
- Pin workflow actions to reviewed commits. Run secret scanning, CodeQL, dependency checks, and regression tests on changes and weekly. Review Dependabot updates before merging.
- Use supported Node 22 for SharePoint Framework builds. Production dependency audits do not replace reviews of bundled code or development tools.

## Administrative responsibilities

The `Classification` field is descriptive metadata, not an access-control rule. Sensitive material requires appropriate SharePoint permissions and sharing restrictions. Lifecycle changes also do not automatically change permissions.

Review Entra consents, selected-resource grants, SharePoint group membership and external sharing, Azure secret expiration and rotation, and GitHub branch/repository protections. Broad delegated grants are bounded by the signed-in user's access but still deserve a least-privilege review. Do not remove permissions without verifying the affected feature and replacement grant.

See [the October 2026 review](docs/security-review-2026-10-02.md) for the tested scope and remaining work. A scan is not a guarantee against all vulnerabilities.
