# Using and Maintaining DORK

> DRAFT • v0.1 • Prepared 2026-09-30 • Intended supporting operations reference  
> Based on Framework v0.2 and the current NERD implementation. No permanent ID assigned.

**Proposed classification:** DATA / Data Governance • Document Type: Procedure. Final classification remains open.

## Start with the useful work

Ask “Why does this document exist?” Write enough for another authorized person to operate, support or understand the subject without unreasonable dependence on your memory. Search DORK before creating a duplicate. If useful knowledge is missing but you cannot document it now, capture the need for follow-up.

The [Framework](framework.md) defines governance. This guide explains the working process. Features labeled planned below are not available merely because the Framework requires them.

## Create a document

1. Choose the Document Type that matches its purpose. A Procedure explains steps; a Checklist verifies completion; a Reference supports lookup; a Standard states required specifications. A Guide primarily instructs nontechnical users.
2. In the DORK Documents library, use the appropriate configured content type/template. The current template library is **DORK-Templates**, separate from Documents. Create a new document from the template rather than editing the reusable template.
3. Give it a clear, descriptive title. Do not invent a DORK number: automatic numbering is not enabled, and 0001 is reserved for the Framework.
4. Use NERD in Word to set the relevant metadata. Save Metadata stores and verifies SharePoint fields; Word's own save remains responsible for document content.
5. Confirm success. If Save reports an error, the operation is not fully verified; do not assume that all fields failed or all fields succeeded. Preserve your content and follow the error guidance.
6. Obtain the required review before treating the document as Active.

Keep Visuals in their suitable formats and signed/finalized Records in their authoritative forms. Metadata can be maintained separately. Current NERD is Word-only; the restored Reference Sheet template does not have live Excel synchronization.

## Choose metadata by purpose

| Field | How to use it |
| --- | --- |
| Title | Describe the useful subject/task; avoid vague names such as “New Document.” |
| Domain / Function | Select one of each. Function must belong to Domain. Use the [Classification Index](classification-index.md) or [TLDR](tldr-search-index.md). |
| Document Type | Match the document's purpose and structure. Not all 14 Framework types have verified creation workflows yet. |
| System / Platform | Select directly applicable technologies; multiple or none. A passing mention is insufficient. |
| Collection | Select relevant program/event/work groupings; multiple or none. Collections cross Domains. |
| Classification | Base Public / Internal / Sensitive on the information contained in the document. Selecting Sensitive is not proof that permissions have been restricted. |
| Audience | Intended readers, optionally multiple. Audience does not grant access. |
| Owner | Ultimately accountable. **Current NERD requires Owner.** Framework's Owner-or-Responsible wording remains under review. |
| Responsible | Maintains and reviews the document. |
| Secondary | Optional maintenance/review backup. |
| Information Source | Identify the authority supplying the information, where useful. |
| Tags | Supplemental terms not already captured/searchable. Do not repeat all Systems or classifications as Tags. |
| Last Reviewed | Date someone confirmed accuracy; not automatically the date of every save. |
| Next Review | Context-based latest validation date, where a scheduled review is useful. No universal annual requirement. |
| Document ID | Permanent assigned identifier once allocation is enabled; do not use a Documents library row number. |

Use role context where useful, without pretending the current people picker supports role-only assignments. Automatically known information should not be retyped solely to fill a field.

### Current save behavior

NERD requires Title, Domain, Function, Document Type, Classification, Lifecycle and Owner. Successful saves verify the submitted dimensions against SharePoint. Optional selections may be removed and saved as empty values; automated coverage exists, while tenant clearing acceptance remains open.

Changing Title can rename the file. Word's top-left filename may not refresh until you close/reopen. Title font preservation was deployed on 2026-09-30 and awaits a visual acceptance check. An observed Word save/reconnect warning is still under investigation.

Use NERD-managed page metadata controls in the regenerated templates. The page is a representation of SharePoint metadata, not a separate place to maintain conflicting values. Avoid deleting those controls when editing content.

## Review and update

1. Confirm that settings, steps, ownership and source information are still accurate.
2. If accurate without revision, record Last Reviewed and set Next Review only when appropriate. Do not rewrite a document to prove that a review occurred.
3. If inaccurate, correct it. Significant changes require review before becoming authoritative; authorized minor corrections may be made directly.
4. Review earlier after significant configuration/process changes, upgrades, vendor changes, incidents, failed procedures, regulatory changes, discovered inaccuracies or key staff reassignment.
5. Reassess responsibility when the people maintaining the document change.

Last Updated means a substantive change. SharePoint Modified may also reflect routine saves and metadata edits. Automatic substantive-change tracking is not implemented.

An overdue document remains Active; it should be visibly marked Overdue Review. Automated overdue views/indicators are planned. A blank Next Review is appropriate only where context supports no schedule, not as a way to conceal an overdue obligation.

## Lifecycle and approval

| Status | Use it when | Required handling |
| --- | --- | --- |
| Draft | Creating or substantially revising content | Treat as work in progress. |
| Active | Current authoritative content | Owner approves activation; sensitive/high-impact material may need additional review. |
| Superseded | Another document replaces it | Identify the replacement; retain history and the old permanent ID. |
| Retired | It no longer applies operationally | Retain history; record useful retirement context. A replacement is not required merely because it is retired. |

The current lifecycle selector does not enforce an Owner approval workflow. Selecting Active is not evidence of approval. The approval method must be settled before this guide is published as an operating procedure.

### Superseded documents

Prepare and approve the replacement first. Record which document replaces the old one, including its permanent ID when numbering is available. Mark the old item Superseded, retain its history and ID, and check relevant references. A replacement document receives its own ID; it does not inherit the old document's number.

Managed Superseded By relationships and lifecycle navigation exclusions are not yet implemented/verified. Until configured, the department needs an agreed interim way to record replacement associations. Do not imply changing the lifecycle field alone completes the relationship or hides the document.

### Retired documents

Use Retired when the system, process or information is no longer operationally applicable and no replacement is needed. Preserve historical evidence. Superseding or retiring an item does not mean deleting it.

## Find and relate information

Search using familiar words, the title, a system name or the operational task. Classification knowledge should help narrow results, not be a prerequisite for finding them.

Use managed relationships for useful associations such as Procedure → Checklist or Reference → Restoration Procedure once implemented. Verify links after moves/archival; do not assume copied or moved files retain the same identity. Search indexing, stable relationships and filtered views still need tenant verification.

Incident IQ tickets should reference authoritative DORK information rather than copy it. A recurring knowledge gap can feed Documentation Needed. A full API integration is not required to follow this practice.

## Documentation Needed and department meetings

Proposed minimal intake: a short statement of what is missing, why it matters, and assignment if known. Helpful context can include the affected system or an existing document. Avoid making every metadata dimension mandatory for a backlog entry.

The operational backlog is planned; GitHub engineering follow-ups are a separate development list. Review needs, new systems, knowledge gaps and assignments during Documentation Development. Review due/overdue items and significant revisions during Documentation Maintenance. The Framework calls for two departmental meetings monthly; urgent needs do not wait for meetings.

Review status alone is not an employee performance measure. Normal supervisory expectations still apply to reasonable assignments.

## Maintain templates and protect records

Keep reusable templates in DORK-Templates and maintain each Documents content type's template reference. Before removing old templates, inventory all references, including the Reference Sheet. Test a new document from each changed template before treating the rollout as complete.

Do not alter finalized evidence just to add a metadata cover. Handle Sensitive content through approved access protections, and verify actual access. Versioning, audit evidence and recovery/export configuration still require tenant checks.

## Before this guide becomes Active

- Resolve Owner/Responsible requirements and the activation approval method.
- Complete tenant tests for all supported types and cleared optional fields.
- Confirm numbering, replacement relationships and review workflows.
- Confirm Sensitive permissions and search/navigation behavior.
- Choose the guide's final document type, owner/responsible and review approach.
- Convert using the appropriate DORK template and verify metadata/page consistency.
