# Word silent sign-in

NERD first reuses its valid tab-session tokens, then attempts MSAL token acquisition with Word's account hint. Word on the web can fall back from the MSAL cache to `ssoSilent` with `Office.auth.getAuthContext().userPrincipalName`. Desktop retains silent host token acquisition. The existing paired Graph/SharePoint Office dialog remains the fallback when silent sign-in is unavailable.

## Existing Entra app registration

Application/client ID: `f8371eb2-8758-48cf-8fa5-2f9b696d89c4`.

Under Authentication, Single-page application, retain existing redirects and ensure these entries are present:

- `brk-multihub://thankful-sky-0cc1b0210.6.azurestaticapps.net`
- `https://thankful-sky-0cc1b0210.6.azurestaticapps.net/taskpane.html`
- `https://thankful-sky-0cc1b0210.6.azurestaticapps.net/sso-redirect.html`

The last page runs the MSAL v5 redirect bridge, without loading Word or starting NERD. It is served with Cache-Control: no-store. Do not add Cross-Origin-Opener-Policy headers to this page.

Use the existing delegated Graph and SharePoint permissions. If deploying to other district users, confirm tenant admin consent for those existing permissions. Do not broaden permissions for this change.

Reopen a DORK document in Word on the web after registration is saved. Success is NERD loading and a verified metadata save without a sign-in popup. A valid session-cache hit alone does not prove fresh silent SSO. Consent, MFA, expired sessions and browser privacy controls can still require the dialog.

Sources:
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in
- https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/redirect-bridge.md
