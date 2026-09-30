# DORK Framework implementation checklist

Reviewed 2026-09-30 against DORK_Framework_v0.2.docx (Library version 2, modified 2026-09-25), current NERD source, and user-reported acceptance tests.

This is a requirements audit, not certification that the tenant is fully configured. "Tested" means the specific behavior was confirmed; platform capability alone is not evidence of configuration. Governance requirements need an operating practice as well as any supporting software.

## Framework coverage

| Framework sections | Promise | Current evidence / gap | Acceptance work |
| --- | --- | --- | --- |
| 1, 1.1, 20, 32 | Useful knowledge; low burden; sufficient documentation | Partial. Silent authentication, automatic pane opening and page refresh tested. Content usefulness and burden need ongoing staff review. | Assess real workflows and remove redundant entry; judge documents by their purpose. |
| 2 | Platform-independent Framework; separate platform instructions | Framework exists; operational instructions remain planned. | Create the DORK operations reference sheet separately from the Framework. |
| 3–5 | Classify by purpose; one Domain and controlled Function taxonomy | Domain/Function selection and persistence tested. Full taxonomy equivalence not audited term by term. | Compare all tenant terms with Framework; document controlled-change authority. |
| 6–7 | Fourteen document types; suitable creation standards and templates | Partial. Five Word templates regenerated; Reference Sheet restored. Current content-type mapping covers six types. Word template page refresh tested on Standard. | Inventory all 14 types and choose an appropriate creation format/workflow for each; test each mapped template. Preserve authoritative Records and native Visual formats. |
| 8–9 | Expandable Systems/Platforms and Collections; multiple or none | Multi-value saving verified. Creation/expansion governance and zero-value behavior need verification. | Test empty and multi-value cases; define who adds controlled terms. |
| 10–12 | Optional Audience and Tags; Classification reflects sensitivity | Selection and persistence tested, including multiple tags. Audience must not imply access rights. | Verify optional-field clearing and tag expansion; verify actual Sensitive access controls separately. |
| 13–14 | Ownership, responsibility, backup and source; appropriate required metadata | People and Source persistence tested. **Mismatch:** NERD requires Owner; Framework permits Owner OR Responsible. | Resolve policy deliberately; verify role-based ownership expectations and required fields by document type. |
| 14.1 | Permanent document identifier | Not implemented. Templates contain an ID slot; NERD can read an existing ID. No allocator or numbering prefix. | Reserve DORK-0001 for Framework; allocate unique permanent IDs, preserve across rename/move, and settle filename prefix convention. Do not use library row IDs as permanent IDs. |
| 14.2 | Distinguish substantive Last Updated from accuracy Last Reviewed | Review dates persist. Substantive-change tracking is not implemented; SharePoint Modified alone is insufficient evidence. | Define and implement both meanings; allow review confirmation without substantive revision. |
| 15 | Draft default; Active, Superseded and Retired; replacement identification; historical retention | Draft default and lifecycle field implemented. Superseded By relationship and normal-navigation exclusions not implemented/verified. | Add replacement association and lifecycle views; retain historical documents and IDs. |
| 16–17 | Contextual reviews, overdue indication and event-driven review | Dates save correctly. Due/overdue presentation and event-triggered workflow not implemented. | Keep overdue documents Active while visibly marking Overdue Review; support varying/no review intervals and earlier reviews after relevant events. |
| 18 | Low-friction Documentation Needed backlog | Not implemented as a DORK operational feature. Repository engineering follow-ups are a different backlog. | Create a staff-facing intake with optional assignment; use in development meetings; handle urgent needs immediately. |
| 19 | Reduce single-person knowledge risk | Operating practice not verified. | Identify high-risk knowledge and assign practical documentation work. |
| 21, 23 | Useful relationships and stable discovery through rename/move/reclassification/archive | Title-to-filename rename tested. Stable relationship management and move/archive behavior not verified. | Add managed Related Documents and Superseded By associations; test links after rename, move and archival. |
| 22 | Search titles, headings, content and metadata; filters; Systems searchable without duplicate tags | SharePoint search configuration and indexing not verified. Dashboard not built. | Test searches and filtered views against real documents, including system-only metadata and content matches. |
| 24 | Two recurring departmental meetings monthly; development and maintenance agendas | Operating practice not verified. | Establish cadence covering backlog, gaps, assignments, due/overdue reviews and significant revisions. Urgent work must not wait. |
| 25 | Staff creation; scoped assignments; significant review; Owner approves Active | No owner-approval gate demonstrated in NERD. Lifecycle selection alone does not prove approval. | Define/enforce activation approval and additional review for Sensitive/high-impact content; allow minor corrections without undue overhead. Do not treat overdue review as an employee-performance metric. |
| 26 | IncidentIQ references authoritative DORK knowledge; recurring gaps feed backlog | Practice/integration not verified. Framework does not require a full API integration. | Establish stable ID/link references and backlog routing; avoid duplicating authoritative content into tickets. |
| 27 | SharePoint authoritative record; appropriate metadata | Save/read-back verification and reopened persistence tested for tested fields. | Complete dimension coverage across document types, optional clearing and lifecycle transitions. |
| 28 | Version/audit history, archival, Sensitive security, files/visuals, export/recovery, identity integration, maintainability | Azure hosting and identity flow tested. Other tenant capabilities/configurations remain unverified. | Verify versioning, audit evidence, permission boundaries, archival and a recoverable export/restore path. |
| 29 | CTO governance/delegation; controlled taxonomy; useful requirements | Operating practice not verified. | Record authorities, delegation and change process; distinguish controlled taxonomy from expandable terms and flexible tags. |
| 30–31 | Classification Index and TLDR Search Only Classification Reference | Current artifacts and completeness not verified in this audit. | Locate/create and validate both references; ensure common search terms aid discovery without overriding purpose-based classification. |

## Next implementation order

- [ ] Resolve Owner-versus-Responsible requirement and activation approval policy.
- [ ] Implement permanent numbering, reserving DORK-0001 for Framework. [Allocation and recovery plan](permanent-numbering.md) prepared; tenant registry/schema/permissions inspection is required before activation.
- [ ] Finish the document-type/template inventory and remaining tenant metadata round-trip tests. Automated request/read-back coverage now includes clearing all nine optional fields, stale-value rejection and invalid review dates (2026-09-30).
- [ ] Confirm title font preservation in Word: explicit per-control font preservation implemented and regression-tested (2026-09-30); tenant visual acceptance remains open. Investigate the observed Word save/reconnect warning separately. Rename concurrency is a hypothesis, not a confirmed cause.
- [ ] Implement review confirmation, contextual schedules and overdue views.
- [ ] Implement Superseded By / Related Documents and lifecycle navigation.
- [ ] Create the Documentation Needed backlog and DORK operations reference sheet.
- [ ] Verify tenant search, Sensitive permissions, version/audit history and export/recovery.
- [ ] Validate classification references and establish the governance/meeting practices.
- [ ] Build the dashboard after documentation is populated: search/filter discovery, due/overdue reviews, ownership, backlog, lifecycle and relationships. Reuse authoritative SharePoint data.

## Confirmed baseline to preserve

- Automatic deployment to the existing Azure Static Web App nerd-dork; application remains entirely Azure-hosted.
- Preserve X-NERD-SharePoint-Authorization in frontend requests and the API proxy.
- Tested Word experience: automatic pane opening, silent sign-in without popups, stable startup and removal of redundant Connect button.
- Tested metadata persistence/read-back, required Classification/Owner validation, title-to-filename rename, and Standard content-type agreement.
- New Word templates use NERD page controls; user uploaded/repointed templates in the DORK-Templates library.
- Reference Sheet template restored and repointed. Live NERD Excel synchronization is not implemented; current add-in is Word-only.

See also [DORK follow-ups](dork-follow-ups.md). Checkboxes remain open until their acceptance work is evidenced; do not infer completion from the platform supporting a feature.
