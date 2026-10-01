/* global Excel */
const sheetName = 'DORK Metadata';
const marker = 'NERD workbook metadata v1';
// Text literals keep titles, tags and sources from becoming spreadsheet formulas.
export function excelText(value) {
    const text = String(value ?? '');
    return /^[=+\-@']/.test(text) ? `'${text}` : text;
}
// Only the known Reference Sheet layout is eligible for cover writes.
export async function syncReferenceCover(context, values) {
    const cover = context.workbook.worksheets.getItemOrNullObject('Document Control');
    cover.load('isNullObject');
    await context.sync();
    if (cover.isNullObject) return 0;
    const signature = cover.getRange('A1:A14');
    signature.load('values');
    await context.sync();
    const expected = {0: 'ROCKAWAY TOWNSHIP SCHOOL DISTRICT | TECHNOLOGY',
        2: 'DORK REFERENCE SHEET', 8: 'DOCUMENT CONTROL', 9: 'Domain',
        10: 'Function', 11: 'System / Platform', 12: 'Collection', 13: 'Classification'};
    if (!Object.entries(expected).every(([row, label]) => signature.values[Number(row)][0] === label)) {
        throw new Error('The Document Control sheet does not match the DORK Reference Sheet template. Its visible fields were not updated.');
    }
    const documentId = String(values.DORK_DocumentId || '').trim();
    const title = documentId
        ? String(values.DORK_Title || '').replace(/^(?:DORK-\d+\s+)+/i, '')
        : String(values.DORK_Title || '');
    const fields = {D10: values.DORK_Domain, D11: values.DORK_Function,
        D12: values.DORK_System, D13: values.DORK_Collection, D14: values.DORK_Classification,
        A17: `${documentId} ${title}`.trim(),
        A24: 'Metadata maintained through NERD. Reference data begins on the next sheet.'};
    for (const [address, value] of Object.entries(fields)) {
        const cell = cover.getRange(address);
        cell.numberFormat = [['@']];
        cell.values = [[excelText(value)]];
    }
    await context.sync();
    return Object.keys(fields).length;
}
export async function syncWorkbookMetadata(values, excel = Excel) {
    return excel.run(async context => {
        let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load('isNullObject');
        await context.sync();
        if (!sheet.isNullObject) {
            const identity = sheet.getRange('A1');
            identity.load('values');
            await context.sync();
            if (identity.values[0][0] !== marker) {
                throw new Error('A sheet named DORK Metadata already exists and is not a NERD metadata sheet. Rename that sheet before retrying.');
            }
        } else {
            sheet = context.workbook.worksheets.add(sheetName);
            sheet.getRange('A1').values = [[marker]];
        }
        const coverUpdates = await syncReferenceCover(context, values);
        const entries = Object.entries(values);
        const rows = entries.map(([tag, value]) => [tag.replace(/^DORK_/, ''), excelText(value)]);
        const range = sheet.getRange(`A3:B${rows.length + 2}`);
        range.numberFormat = rows.map(() => ['@', '@']);
        range.values = rows;
        sheet.getRange('A2:B2').values = [['Metadata field', 'Value']];
        sheet.getRange('A2:B2').format.font.bold = true;
        sheet.getRange('A:A').format.columnWidth = 150;
        sheet.getRange('B:B').format.columnWidth = 300;
        range.format.wrapText = true;
        range.format.autofitRows();
        // Do not activate the metadata sheet or change the user's active worksheet.
        await context.sync();
        return entries.length + coverUpdates;
    });
}
