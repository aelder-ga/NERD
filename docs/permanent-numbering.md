# Permanent DORK numbering

Implementation plan prepared 2026-09-30. Number assignment is **not enabled**.

## Contract

- Reserve DORK-0001 for the authoritative Framework. Identify that file explicitly; never infer it from a filename.
- Assign one permanent identifier to each new authoritative document, automatically on first successful metadata save.
- Render identifiers as DORK-0002, DORK-0003, etc.; four digits are a minimum, not a maximum.
- Title edits, classification changes, review and retirement do not allocate another number.
- A new document copied from a template or existing document needs its own identifier, even if copied metadata contains an old identifier.
- Keep the authoritative identifier in SharePoint metadata. NERD's page control displays it; users do not type it.
- Proposed filename: DORK-0002 Document Title.docx. Keep Title as the readable title; construct the prefix from the assigned identifier, avoiding duplicated prefixes.
- Gaps after failed or abandoned creation are acceptable. Never recycle numbers.

## Storage and concurrency

Use a dedicated SharePoint registry inside DORK, separate from the Documents library. The document library's item ID is not a permanent document identifier.

A proposed allocator uses the registry's server-assigned item sequence, with an offset reserving 0001. This requires a newly initialized, non-recreated registry and an existing-document inventory before choosing the offset. The final allocation must be collision-checked against imported/existing IDs. Restoring or migrating the registry must preserve mappings and the allocation high-water mark.

Registry schema:

| Field | Purpose |
| --- | --- |
| DocumentKey (indexed, unique text) | Identity of the actual file, independent of its displayed title/path |
| DorkId (indexed, unique text) | Permanent public identifier |
| AllocationState | Reserved / Applied; retry diagnostics without recycling |
| CurrentLocation | Recovery/discovery aid; not the document identity |

The Framework mapping is explicitly initialized to DORK-0001. Registry writes must be protected from ordinary modification/deletion. A frontend-only counter or unrestricted registry would not enforce this contract; use a trusted Azure allocator with least-privilege registry access if delegated permissions cannot enforce it. Keep frontend and API hosted on Azure. Do not add infrastructure until the tenant's existing capabilities and permission model have been inspected.

SharePoint supports unique column constraints; these must be deliberately configured:
https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/fb223131-3f8d-4590-857d-a334efa05f4b

## Save/recovery sequence

1. Resolve the actual file identity and existing document ID. Validate the ID's registry binding; copied template metadata is not ownership evidence.
2. Find an existing registry reservation by unique DocumentKey. Reuse it for retries.
3. If absent, create a reservation through the trusted allocator. On a confirmed uniqueness race, read the winning reservation. Do not treat unrelated network/authorization errors as successful allocation.
4. Persist and read back the assigned DorkId in the registry before returning it. A crash between reservation and ID write must be recoverable from the reservation.
5. Save metadata, including DorkId, to the authoritative document. If this fails, retain the reservation and retry with the same ID.
6. Rename using the assigned ID plus title. If rename fails, retain the ID and report a partial result; do not allocate again.
7. Verify document ID, metadata and filename; refresh the page control.
8. Mark the registry mapping Applied. Failure here must not undo an already verified document save; recover on retry.
9. Keep current location updated on supported moves and retain mappings for superseded/retired documents.

Do not assume a move is a copy or vice versa. Confirm file identity behavior in supported DORK workflows before activation. Native SharePoint Document ID services provide a separate stable-link mechanism and must be inspected before duplicating that functionality; they do not establish that the required exact DORK-0001 sequence is already configured:
https://support.microsoft.com/en-us/sharepoint/admin/enable-and-configure-unique-document-ids

## Activation gates

- [ ] Inventory existing IDs and locate the authoritative Framework.
- [ ] Inspect native Document ID configuration and document identity across rename/copy/move.
- [ ] Provision and verify document ID metadata plus registry constraints and permissions.
- [ ] Implement the trusted allocator and recovery protocol; configure its Azure identity only if required.
- [ ] Test concurrent saves for one file and for different files.
- [ ] Test ambiguous network failure, retries, reservation-only failure and rename failure.
- [ ] Test copies of templates and numbered documents receive new IDs.
- [ ] Test rename/move/reclassification/archive retain the intended ID.
- [ ] Verify Framework receives 0001 and ordinary creation can never receive it.
- [ ] Verify restore/migration cannot recycle numbers.
- [ ] Enable allocation and filename prefixes only after these gates pass.

Available connector operations currently do not expose list/column provisioning or the required registry permission setup. No tenant configuration or numbering migration was performed in this change.
