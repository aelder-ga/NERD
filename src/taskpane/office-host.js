/* global Office, Word, Excel */
import { syncDocumentMetadata } from './document-metadata';
import { syncWorkbookMetadata } from './workbook-metadata';
export function assertSupportedHost(office = Office) {
    if (office.context.host === office.HostType.Excel) {
        if (!office.context.requirements.isSetSupported('ExcelApi', '1.11')) {
            throw new Error('NERD requires ExcelApi 1.11. Open this workbook in Excel for the web or an updated Excel desktop app.');
        }
    } else if (office.context.host !== office.HostType.Word) {
        throw new Error('Open NERD in Word or Excel.');
    }
}
export async function syncHostMetadata(values) {
    assertSupportedHost();
    return Office.context.host === Office.HostType.Excel
        ? syncWorkbookMetadata(values) : syncDocumentMetadata(values);
}
export async function saveHostDocument(office = Office, word = null, excel = null) {
    assertSupportedHost(office);
    if (office.context.host === office.HostType.Excel) {
        return (excel || Excel).run(async context => {
            context.workbook.save('Save');
            await context.sync();
        });
    }
    return (word || Word).run(async context => {
        context.document.save();
        await context.sync();
    });
}
