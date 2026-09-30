/* global Word */
const metadataFontProperties = ['name', 'size', 'bold', 'italic', 'color', 'underline'];
// These tags form the contract between NERD and the Word templates.
export async function syncDocumentMetadata(values, word = Word) {
    return word.run(async context => {
        const controls = context.document.contentControls;
        controls.load('items/tag,items/text,items/cannotEdit');
        await context.sync();
        const changed = controls.items.filter(control =>
            Object.prototype.hasOwnProperty.call(values, control.tag) &&
            control.text !== values[control.tag]);
        if (!changed.length) return 0;
        // Each occurrence has its own formatting (cover title versus metadata table).
        // Word for the web can reset it when replacing content-control text.
        for (const control of changed) control.font.load(metadataFontProperties);
        const locked = changed.filter(control => control.cannotEdit);
        await context.sync();
        const formats = changed.map(control => Object.fromEntries(
            metadataFontProperties.map(property => [property, control.font[property]])
                .filter(([property, value]) => value != null && value !== '' &&
                    (property !== 'size' || value > 0))));
        try {
            for (const control of locked) control.cannotEdit = false;
            await context.sync();
            for (let index = 0; index < changed.length; index++) {
                const control = changed[index];
                control.insertText(values[control.tag], 'Replace');
                control.font.set(formats[index]);
            }
            await context.sync();
        } finally {
            for (const control of locked) control.cannotEdit = true;
            await context.sync();
        }
        return changed.length;
    });
}
