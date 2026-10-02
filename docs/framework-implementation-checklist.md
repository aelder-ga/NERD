# DORK Framework implementation checklist

Reviewed 2026-10-02 against the uploaded DORK-0001 DORK Framework, the earlier framework draft, deployed source at 6435ed9f36e4b5f5800082576fa9ff52ecae6fe4, and user-confirmed tenant tests.

This audit separates source implementation, user-confirmed behavior, planned features and unverified tenant configuration. It does not certify all SharePoint settings. The uploaded Framework copy is the document evidence; no direct inspection of a newer SharePoint copy was performed.

## Current Framework coverage

| Current section | Commitment | Evidence and remaining work |
| --- | --- | --- |
| 1 | Useful, authoritative, maintainable knowledge | SharePoint is the document store. Usefulness, ownership and maintenance remain operating practices to assess with staff. |
| 2–3 | Six dimensions, purpose-based Domain/Function, controlled taxonomy | Selection and persistence have passed user tests. Full tenant vocabulary has not been compared term by term. Source sorts terms alphabetically, which differs from the Framework's locked Domain order. |
| 4 | Eleven document types; suitable formats | Word templates and Reference Sheet exist; GEEK supports native files with general DORK Document fallback. Dedicated content types/templates for every type are not promised by the current text. Inventory actual term/type/template coverage. |
| 5 | Permanent IDs on metadata save; Framework 0001; stable IDs on rename | Numbering, prefix rename and existing-ID retention implemented and user-confirmed. Framework reservation bound to registry row 4. Latest allocation-boundary repair is deployed and regression-tested; the next fresh allocation expected to be 0005 is not yet tenant-confirmed. Opening a document alone does not allocate. An allocation can remain reserved after a later save failure; retry the same file. Moves, copies, reuploads and recovery still require validation/admin reconciliation. |
| 6 | Owner, responsibility, classification and supporting metadata | Owner and Classification validation confirmed. Metadata save/reopen confirmed for tested Word and Excel fields; platform removal persisted. Classification labels do not enforce access permissions. Optional-field clearing has automated coverage, but not every field/type has a tenant acceptance test. |
| 7 | Draft, owner release review, review dates, supersession and history | Draft default and date persistence implemented. Owner release review is currently a human process: no authorization/approval gate demonstrated. No managed Superseded By relationship or lifecycle navigation implemented. Retention/version settings unverified. |
| 8 | NERD Word/Excel; GEEK intake; Azure hosting and authentication | User-confirmed Word auto-open, stable saves/rename and silent sign-in; Excel embedded fields and reopened metadata; GEEK upload, ID/name/metadata verification and SharePoint launcher. All application hosting remains Azure. Excel automatic opening is not established. GEEK editing does not synchronize embedded Office fields; close Office before editing externally. |
| 9 | Completed working records stored separately with context | Policy described, but destination, permissions and staff workflow not verified. Do not allocate a DORK ID to every completed checklist merely because its source template has one. |
| 10 | Registry, document register, Needs list and dashboard | Numbering registry operational. Registry is not a staff-facing document register or Needs backlog. Register views, staff Needs list, dashboard, calendar and charts remain planned. |
| 11 | Framework maintenance and structural-change governance | Operating practice not verified. Related documentation still calls 0003 planned, although the user has now populated Microsoft Administration Portals. Update that wording in the authoritative Word document. |

## Source discrepancies to address

- [ ] Preserve the specified Domain order instead of alphabetically sorting the Domain picker; confirm the intended order against tenant terms.
- [ ] Check Excel Reference content-type routing. Current source special-cases Reference + .xlsx only in GEEK intake; NERD maps the plain Reference term to DORK Reference Doc, while Reference Sheet maps correctly. Confirm the actual term and content type on 0003 before treating its current result as broken.
- [ ] Reconcile taxonomy inventories. Current Framework lists eleven types including Runbook; the older draft lists fourteen including Form, Report, Policy and Register but no Runbook. Decide the authoritative vocabulary rather than silently dropping or adding terms.
- [ ] Compare all Domain/Function labels and relationships against the live term store, including Data Governance versus earlier Governance terminology.
- [ ] Change the 0003 related-document entry from planned to Microsoft Administration Portals, DORK-0003. Clean up the Standard Template footer if desired.

## Earlier commitments that need an explicit decision

The earlier draft is preserved in [framework.md](drafts/framework.md). These commitments were not fully carried into the current uploaded Framework. Their omission is not evidence that the user approved abandoning them.

| Earlier commitment | Current disposition / decision needed |
| --- | --- |
| Owner OR Responsible | Current Framework and NERD require Owner. Confirm this deliberate policy change; no current-text mismatch remains. |
| Substantive Last Updated, distinct from Last Reviewed | Not implemented as a separate field/process. SharePoint Modified is not a substitute for substantive revision tracking. Decide whether to retain this requirement. |
| Managed Related Documents and Superseded By; stable links through moves/archive | Current text describes manual relationships; managed association and move/archive acceptance remain open. |
| Overdue review indication and event-driven reviews | Dates exist; overdue views/notifications and trigger practices remain open. |
| Two monthly department meetings: development and review | Cadence not verified. Retain, revise or retire explicitly. |
| Knowledge concentration assessment and staff Documentation Needed intake | Needs list is planned; risk assessment and operating practice remain open. |
| Owner approval for Active; extra review for Sensitive/high-impact material | Current Framework has owner checking before release; enforcement and evidence remain open. |
| IncidentIQ references and routing recurring gaps into the backlog | Not verified. This can be a staff process and does not require an API integration. |
| Search across title, headings, content and metadata | Tenant indexing and representative searches not tested in this audit. |
| Version/audit history, Sensitive permission boundaries, archival, export and recovery | Platform capability is insufficient evidence; configuration and recovery acceptance remain open. |
| CTO/delegated governance and controlled taxonomy changes | Authority and change process not verified. |
| Classification Index and TLDR searchable reference | Drafts exist; final review/publication still open. |

## Next implementation order

- [x] Word authoring and metadata workflow accepted for tested scenarios.
- [x] Excel synchronization and reopen persistence accepted for tested scenarios.
- [x] GEEK upload, metadata, ID/name verification and launcher accepted.
- [x] Framework assigned DORK-0001; procedure 0002, portal reference 0003 and floorplans 0004 retained.
- [ ] Confirm next fresh allocation is 0005 when creating a useful document; do not create a throwaway numbered file.
- [ ] Resolve the source discrepancies and taxonomy decisions above.
- [ ] Build the staff document register as SharePoint views and create the Documentation Needed list, with a simple staff intake.
- [ ] Add review-due/overdue views and agree release approval, supersession and relationship handling.
- [ ] Verify search, Sensitive access, version/audit history, archival and recovery.
- [ ] Publish the classification index/TLDR and record governance practices.
- [ ] Build the dashboard over authoritative SharePoint data: search/filter discovery, due reviews, ownership, Needs, lifecycle, calendar and charts.

## Baseline to preserve

- Existing Azure Static Web App nerd-dork and automatic GitHub deployment; all application hosting remains Azure.
- X-NERD-SharePoint-Authorization in both frontend and API proxy.
- Stable Word startup/auto-open, silent authentication and removal of redundant Connect button.
- Tested metadata verification, required Owner/Classification checks, permanent IDs, filename renames and reopened persistence.
- User-installed templates in DORK-Templates.
- Excel field synchronization and metadata persistence; GEEK remains distinct from embedded Office synchronization.

See also [DORK follow-ups](dork-follow-ups.md). Completion applies to the tested scenario, not every content type, permission boundary or tenant setting.
