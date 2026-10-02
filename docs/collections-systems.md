# Collections and Systems manager

`/terms.html` is hosted in the existing nerd-dork Azure Static Web App. It uses the existing NERD Microsoft application and browser MSAL cache, with silent sign-in first and a user-clicked sign-in fallback. No new registration, app-only credential or API scope was added. Existing NERD and GEEK authentication/proxy behavior is unchanged.

The manager discovers the DORK term group and only exposes its Collection and System / Platform term sets. It lists current entries and descriptions, searches labels and aliases, and adds a top-level term with an optional description. Duplicate checks read fresh data, include nested terms and aliases, and compare case, spacing and Unicode-normalized names. It does not rename or delete terms, modify controlled Domain/Function/Document Type values, allocate document IDs, or update documents.

Microsoft Graph enforces the signed-in user's existing term-store write permissions. The page does not grant contributor rights. A read-only account may browse but receives an explicit permission message if an addition is rejected. Graph/server duplicate conflicts are reported. Uncertain write outcomes are not automatically retried; refresh and search before another attempt.

NERD and GEEK read these same term sets when connecting. Reopen or reconnect those tools to load additions. GEEK has a management link; DORK Home navigation provides the site entry point.

Validation: five service tests cover set restrictions, duplicate aliases, nested terms, paging, permission rejection and uncertain-write handling. The full root regression suite and Azure production build pass. Live list/search and duplicate checks should be verified after deployment. A successful real addition and a lower-access account's rejection remain tenant acceptance tests; do not populate production with fake test terms.
