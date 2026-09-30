/* global Word */
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
        const locked = changed.filter(control => control.cannotEdit);
        for (const control of locked) control.cannotEdit = false;
        await context.sync();
        try {
            for (const control of changed) control.insertText(values[control.tag], 'Replace');
            await context.sync();
        } finally {
            for (const control of locked) control.cannotEdit = true;
            await context.sync();
        }
        return changed.length;
    });
}
