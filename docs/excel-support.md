# Excel support

The same NERD add-in supports Word and Excel through manifest version 1.0.2.0. Update the existing add-in deployment with manifest-azure.xml; retain the existing add-in identity. No second app registration or new Azure app is required.

Excel requires ExcelApi 1.11. NERD uses the existing SharePoint metadata schema, authentication, protected proxy header, and permanent numbering registry. An assigned number uses the workbook's identity, not its filename. Renames preserve the workbook extension. Excel worksheet changes and embedded SharePoint server properties are saved before SharePoint metadata is written and verified, and before the file is renamed. No worksheet metadata writes occur during startup; reopening only hydrates the pane. Word retains its previously verified save sequence.

NERD writes a dedicated `DORK Metadata` worksheet with a version marker in A1 and the metadata fields in columns A and B. It does not activate that sheet or modify working worksheets. If an unrelated sheet already uses this name, NERD refuses to overwrite it. Metadata values are literal text, including titles beginning with formula characters. The known DORK Reference Sheet Document Control layout also receives its numbered title, Domain, Function, System / Platform, Collection, and Classification. NERD checks the template headings and labels before writing these cells, preserving merged ranges and formatting. An unrecognized Document Control layout blocks a successful field-update report. Checklist controls are not yet implemented.

## Acceptance still required

Automated checks cover host dispatch, metadata sheet collisions, literal text, and save ordering. They do not establish live Excel authentication, synchronization, or auto-open behavior.

Use the next intentional DORK Reference Sheet for live acceptance, rather than a disposable numbered test workbook. Confirm all dimensions in SharePoint after saving, the same number after another save, title rename without losing NERD, unchanged working-sheet formulas/data, and automatic opening after closing and reopening. Do not copy an identifier into a new file to reuse an allocation. Preserve DORK-0002 for the planned documentation procedure.

Embedded server-property synchronization uses the workbook’s own content-type schema to map internal field names to XML elements. It updates taxonomy term GUIDs, people, multichoice values, dates, and scalar fields, along with the core title and content-type identity. Stale duplicate custom-property aliases for managed fields are removed; unrelated properties and XML are preserved. This addresses a workbook that retained populated worksheet cells while blank embedded properties overwrote library metadata. Live reopen acceptance is still required.
