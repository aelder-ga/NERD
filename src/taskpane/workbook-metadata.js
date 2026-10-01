/* global Excel */
const sheetName = 'DORK Metadata';
const marker = 'NERD workbook metadata v1';
// Text literals keep titles, tags and sources from becoming spreadsheet formulas.
export function excelText(value) {
    const text = String(value ?? '');
    return /^[=+\-@']/.test(text) ? `'${text}` : text;
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
        return entries.length;
    });
}
