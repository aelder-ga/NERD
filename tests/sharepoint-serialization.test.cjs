const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.js'), 'utf8');
const block = source.slice(source.indexOf('function taxonomyMultiValue('), source.indexOf('function buildDesiredFilename('));
const context = {};
vm.runInNewContext(block + '\nglobalThis.multi = taxonomyMultiValue; globalThis.date = sharePointDateValue;', context);
test('File-item multi taxonomy uses semicolon separated label/GUID pairs', () => {
    const values = [{label: 'Meraki', id: 'bb046161-49cc-41bd-a459-5667175920d4'}, {label: 'Webex', id: '0069972e-67f1-4c5e-99b6-24ac5c90b7c9'}];
    assert.equal(context.multi(values), 'Meraki|bb046161-49cc-41bd-a459-5667175920d4;Webex|0069972e-67f1-4c5e-99b6-24ac5c90b7c9');
    assert.equal(context.multi([values[0]]), 'Meraki|bb046161-49cc-41bd-a459-5667175920d4');
    assert.equal(context.multi([]), '');
});
test('Dates use DORK form format and preserve calendar days and clearing', () => {
    for (const [value, expected] of [['2026-09-30','9/30/2026'], ['2026-10-01','10/1/2026'], ['2028-02-29','2/29/2028'], ['','']]) {
        assert.equal(context.date(value), expected);
    }
    assert.throws(() => context.date('2026-02-29'), /calendar date/);
    assert.throws(() => context.date('2026-13-01'), /calendar date/);
    assert.throws(() => context.date('9/30/2026'), /YYYY-MM-DD/);
});
test('Both review date form fields call the serializer', () => {
    for (const id of ['last-reviewed', 'next-review']) {
        assert.ok(source.includes('sharePointDateValue(document.getElementById("' + id + '").value)'));
    }
});
