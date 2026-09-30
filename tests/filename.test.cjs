const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.js'), 'utf8');
const block = source.slice(source.indexOf('function buildDesiredFilename()'), source.indexOf('async function verifySavedMetadata('));
function filename(title, loadedTitle, name) {
    const context = {document: {getElementById: () => ({value: title})}, currentDriveItem: {name}, getFieldValue: () => loadedTitle};
    vm.runInNewContext(block + '\nglobalThis.result = buildDesiredFilename();', context);
    return context.result;
}
test('Unchanged title preserves a copied filename instead of colliding with the original', () => {
    assert.equal(filename('Test Standard', 'Test Standard', 'Test-Standard-Repaired.docx'), 'Test-Standard-Repaired.docx');
    assert.equal(filename(' Test Standard ', 'Test Standard', 'Test Standard (1).docx'), 'Test Standard (1).docx');
});
test('Explicit title change still renames and preserves the file extension', () => {
    assert.equal(filename('Firewall Guide', 'Test Standard', 'Test-Standard-Repaired.docx'), 'Firewall Guide.docx');
});
test('Invalid title remains rejected', () => {
    assert.throws(() => filename('Bad/Name', 'Old', 'Old.docx'), /filename characters/);
});
