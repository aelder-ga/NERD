# DORK Documentation Governance Framework

> WORKING DRAFT • v0.3 candidate • Prepared 2026-09-30  
> Based on DORK_Framework_v0.2.docx, Library version 2, modified 2026-09-25.  
> **DORK-0001 is reserved for the Framework; it has not been allocated by the software.**

**Proposed classification:** DATA / Data Governance • Document Type: Standard. This is a recommendation pending adoption.

Rockaway Township School District • Technology Department  
DORK: Documentation Organization & Reference Key

This editorial working draft carries forward the v0.2 principles and requirements, links the full taxonomy to its companion index, and adds a clearly separated publication review appendix. It is not an approved replacement for v0.2. Requirements describe the intended system; implementation progress belongs in the [implementation checklist](../framework-implementation-checklist.md).

## 1. Purpose

DORK establishes how the Technology Department organizes and maintains operational documentation. It makes reliable information easy to find and reduces dependence on undocumented individual knowledge.

Technology documentation must provide staff with enough information to operate and support a system or process without unreasonable dependence on individual memory. Documentation is created where it provides meaningful operational value; every Technology activity does not require a document.

### 1.1 Usability

DORK must remain easier to use than to avoid. Required metadata, review processes and administrative steps must justify their operational value. Information reliably generated, inherited or inferred by the platform should not require manual entry.

## 2. Platform Independence

The platform implements DORK; it does not define it. This Framework remains independent of the selected software. Platform-specific instructions belong in separate supporting documentation.

## 3. Governing Classification Principle

**Why does this document exist?**

Classify by operational purpose, not simply by a technology mentioned. Domain and Function locate the document; other fields provide context and additional discovery paths.

Disabling an employee account after separation is IDENTITY / Account Management. Auditing former employees' access is CYBERSECURITY / Audit & Controls. Testing accommodations data belongs in DATA; testing wireless readiness belongs in NETWORK / Wireless. Both may use the Standardized Testing Collection.

## 4. Classification Model

Each item has enough structured information to identify what it is, where it belongs and who is responsible. Field requirements may vary by Document Type.

| Category | Purpose |
| --- | --- |
| Required | Necessary classification or control information |
| Automatic / Defaulted | Information generated or assigned whenever practical |
| As Needed | Useful context or improved discovery |

### 4.1 Domain

Each document has one Domain.

| Readability group | Domains |
| --- | --- |
| Technical Infrastructure | NETWORK, SERVER, ENDPOINT |
| Access, Information & Security | IDENTITY, DATA, CYBERSECURITY |
| Operational & Business | FACILITIES, BUSINESS |
| Public-Facing | PUBLIC |

These group names are not another classification level.

## 5. Functions

Each document has one Function belonging to its Domain. The [Classification Index](classification-index.md) contains all 62 controlled Function names and descriptions from v0.2 and forms the Function reference for this working draft. Changes to these names require controlled review.

## 6. Document Types

| Document Type | Purpose |
| --- | --- |
| Settings Configuration | Required or intended system settings |
| Procedure | Steps for completing a task |
| Reference | Information maintained for lookup and understanding |
| Visual | Visual representation of systems or locations |
| Checklist | Items requiring verification or completion |
| Template | Reusable structure for creating documents |
| Form | Structured document for collecting information |
| Report | Documented findings or analysis |
| Record | Evidence of an activity or transaction |
| Policy | Rules governing required behavior |
| Standard | Required technical or operational specifications |
| Guide | Instructions for nontechnical users |
| Register | Maintained structured list or inventory |
| Plan | Intended future or contingency actions |

A Procedure explains work; a Checklist verifies completion. A Record preserves evidence, including signed agreements, purchase orders, approvals and completed forms. Guides primarily support users outside Technology. Visuals include diagrams, maps, layouts and flowcharts.

## 7. Document Creation Standards

Use a consistent structure suited to each type. Templates should simplify creation and understanding, without unnecessary work.

Procedures contain the information needed to perform the task. Settings Configurations use structured values, selections, checkboxes and expected states where practical. References support quick lookup. Checklists contain concise, actionable verification items.

Keep Visuals in suitable native formats. Keep Records in their authoritative form. Maintain metadata around these artifacts when editing the artifact would be inappropriate. Supporting DORK standards maintain detailed templates.

## 8. System / Platform

Identify technologies to which the document directly applies. Multiple values or none are permitted. A mention alone does not justify a value.

The list is controlled but expandable. Examples include Google Workspace, Google Admin, Google Drive, Entra ID, Intune, Microsoft 365, Windows, ChromeOS, iPadOS, Meraki, Duo, Genesis, Clever, Frontline, Systems 3000, GoGuardian, Gaggle, Thrillshare, Webex, Incident IQ, Raptor, NJSLEDS and CRDC.

System values must be discoverable without duplicating them as Tags.

## 9. Collections

Collections group documents for an activity, event, program or operational purpose, across Domains. They do not replace Domain or Function. Multiple Collections or none are permitted; Collections expand as needed.

Examples: Standardized Testing, Student Enrollment, Staff Onboarding, Staff Offboarding, New School Year, Emergency School Closure, Disaster Recovery, Chromebook Program, Board Meetings and User Management.

## 10. Audience

Audience is optional and may contain multiple values: Technology, Administration, Staff, Students, Parents, Public, Vendor and Government Agency.

Audience does not determine access or sensitivity.

## 11. Classification

| Classification | Meaning |
| --- | --- |
| Public | Approved for unrestricted public access |
| Internal | Intended for authorized district personnel |
| Sensitive | Requires additional protection |

Classification reflects the information contained in the item, not merely its subject.

## 12. Tags

Tags are optional supplemental search terms for context not already captured by DORK or the document. Examples include MFA, VLAN, ransomware, fiber, UPS, Chromebook repair, NJSLA and FERPA.

Do not manually reproduce Domain, Function, System / Platform or other searchable information as Tags.

## 13. Responsibility

| Role | Meaning |
| --- | --- |
| Owner | Ultimately accountable for the documentation |
| Responsible | Maintains and reviews the documentation |
| Secondary | Optional backup to Responsible |
| Source | Authority providing information used in the documentation |

Assignments may identify a role followed by the current individual when useful. Sources include district departments, schools, vendors and government agencies. Technology faithfully documents supplied information; the Source remains responsible for its accuracy and timeliness.

## 14. Document Control

Usually required control information comprises Document ID, Title, Domain, Function, Document Type, Classification, **Owner or Responsible**, and Status. Appropriate requirements may vary by type.

Additional fields include System / Platform, Collection, Audience, Tags, Secondary, Source, Version, Last Reviewed, Next Review, Superseded By, Document Location and Related Documents. Generate creation/modification dates and useful defaults whenever practical.

### 14.1 Document ID

Each item receives a permanent identifier, generated automatically where supported. DORK-0001 is reserved for the Framework by the project decision recorded on 2026-09-30. Other documents must not receive that number. The allocation design is maintained separately in [Permanent Numbering](../permanent-numbering.md).

### 14.2 Last Updated and Last Reviewed

Last Updated records the most recent substantive change and is generated automatically when possible. Last Reviewed records confirmation of continued accuracy. A review need not produce a revision. An ordinary platform modification timestamp does not by itself establish a substantive change.

## 15. Lifecycle

| Status | Meaning |
| --- | --- |
| Draft | Being created or substantially revised |
| Active | Current authoritative documentation |
| Superseded | Replaced by newer documentation |
| Retired | No longer operationally applicable |

New documents default to Draft where practical. Superseded documents identify their replacement. Superseded and Retired documents remain available historically but are excluded from normal navigation.

## 16. Review and DORK Maintenance

Review according to context, not a universal interval. Consider change frequency, consequences of inaccuracy and events likely to reveal a needed update. Stable infrastructure may justify longer intervals; finalized historical Records may need no scheduled review.

Next Review is the latest date by which an Active document should be validated. Operational changes may require earlier attention.

An overdue document remains Active and is clearly marked Overdue Review until reviewed. Accuracy confirmation can occur without revision.

## 17. Event-Driven Review

Review affected documents after significant configuration changes, major upgrades, vendor changes, cybersecurity incidents, failed procedures, regulatory or process changes, discovered inaccuracies, or loss/reassignment of key staff. Significant operational changes include updating affected documentation. Minor changes should not create unnecessary work.

## 18. Documentation Needed

Maintain a low-friction backlog of important missing information and undocumented individual knowledge. Intake requires only enough information to identify the need and assign it when appropriate. Review it during Documentation Development; address urgent needs without waiting.

## 19. Institutional Knowledge Risk

Significant operational knowledge held by one person creates documentation risk. Capture it when practical to reduce dependence on that person.

## 20. Documentation Sufficiency

Begin with the document's purpose. Depth and quantity should reflect operational value, preserved knowledge and reduced dependence on memory, rather than a universal document quota.

## 21. Relationships

Maintain direct associations where they provide useful context, preferably through the platform rather than fragile manually embedded links.

Examples: Configuration → Procedure; Procedure → Checklist; Server Reference → Restoration Procedure; Account Provisioning → Rostering Process; Incident Response → Emergency Communications; Superseded Document → Replacement Document.

## 22. Navigation and Discovery

Users must be able to discover information without knowing its classification. Search titles, headings, content, Domain/Function, type, Systems, Collections, Tags and relevant metadata where supported. Provide useful filtered navigation. Tags supplement discovery; they are not the primary search mechanism.

**Always check the DORK.**

## 23. Navigational Integrity

Renaming, moving, reclassifying, updating or archiving should not unnecessarily break relationships or navigation. Use platform-managed identifiers and relationships where practical.

## 24. Department Operating Rhythm

Hold two recurring departmental meetings each month, incorporating DORK as standing agenda items alongside other departmental business.

### 24.1 Documentation Development

Cover Documentation Needed, new documentation, new systems/processes, knowledge gaps and assignments.

### 24.2 Documentation Review

Cover due and overdue reviews, inaccurate documentation and significant revisions. Urgent issues do not wait for a meeting.

## 25. Creation, Assignment, and Approval

Technology staff may create useful documentation. The CTO and appropriate supervisors may assign creation, revision or review with reasonable scope, quality and completion expectations.

Keep administrative work minimal. Authorized staff may correct minor errors directly. Review significant changes before they become authoritative. **The Owner approves Active status.** Sensitive/high-impact content may need additional review.

DORK review status alone is not an employee performance measure. Normal supervisory expectations still apply to assigned responsibilities.

## 26. Help Desk Integration

DORK is the authoritative Technology operational knowledge source. Incident IQ tickets may reference its document IDs or links. Repeated or missing knowledge may generate Documentation Needed entries. Prefer referencing authoritative information over copying it into tickets.

## 27. Authoritative DORK Record

The platform maintains the authoritative metadata appropriate to each item: ID, Title, Domain, Function, type, Systems, Collections, Audience, Classification, Tags, Owner, Responsible, Secondary, Source, Version, Status, Created, Last Updated, Last Reviewed, Next Review, Superseded By, Location and Related Documents.

A field's presence in the model does not make it mandatory for all items. Metadata need not appear inside every underlying artifact.

## 28. Platform Requirements

The platform must support:

- Low-friction creation, templates or structured creation.
- Structured metadata, controlled values, optional/required fields and practical automation/defaults.
- Full-text, title, heading and metadata search; useful filtered navigation.
- Managed relationships and stable identifiers or links.
- Version history, audit history, archival and practical export/recovery.
- Review dates, contextual schedules and overdue identification.
- Sensitive access controls; suitable support for files and Visuals.
- Documentation Needed tracking.

The platform must remain maintainable without excessive custom development and minimize duplicate entry. Use district identity/security integration where practical. Adding useful documentation must remain substantially easier than administering the system.

## 29. DORK Administration

The CTO initially maintains governance and administration; routine administration may be delegated as DORK matures.

Domain, Function and Document Type changes are controlled. Systems and Collections expand as needed; Tags remain flexible. Add requirements only when their operational value justifies user effort.

## 30. Classification Index

Maintain a compact Domain | Function | Short Description reference. The complete [Classification Index](classification-index.md) accompanies this draft.

## 31. TLDR: Search-Only Classification Reference

Maintain common terms associated with classifications, including synonyms users actually search for. The [TLDR reference](tldr-search-index.md) assists discovery; it does not determine classification. Purpose remains the governing question.

## 32. Framework Maintenance

DORK evolves with operational needs. Change controlled taxonomy deliberately; expand other values through routine administration. Evaluate every new field, requirement and process against its practical value and burden.

Success depends on information quality and staff willingness to maintain it.

## Publication review appendix — unresolved, not policy changes

- Resolve the existing tension between “Owner or Responsible” in section 14 and Owner approval in section 25. Current NERD requires Owner. This draft retains both v0.2 provisions pending a deliberate decision.
- Confirm classification and document type for the Framework and supporting documents. Only the Framework's reserved number is decided; do not preassign IDs to the others.
- Confirm the authoritative Framework file before enabling numbering.
- Check implementation evidence separately; this draft does not mark unmet requirements complete.
- Review proposed TLDR vocabulary against actual staff searches.
- Before conversion, assemble the complete taxonomy with the Framework or retain a controlled companion reference so section 5 remains complete.

### Editorial changes from v0.2

This candidate reorganizes wording into Markdown, moves the complete Function reference to its companion file, incorporates the agreed Framework number reservation, clarifies substantive updates versus routine timestamps, and records unresolved policy questions. No approval rule, taxonomy name, required-owner policy or universal review interval has been newly approved.
