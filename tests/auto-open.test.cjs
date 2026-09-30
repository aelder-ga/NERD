const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../src/taskpane/auto-open.js'), 'utf8').replace('export async function', 'async function');
function fixture(initial, supported = true, fails = false) {
    let saved = initial, saves = 0;
    const control = {}, container = {hidden: true}, status = {};
    const settings = {get: () => saved, set: (_, value) => {saved = value;}, remove: () => {saved = undefined;}, saveAsync: callback => {saves++;callback({status: fails ? 'failed' : 'ok'});}};
    const context = {Office: {context: {document: {settings}, requirements: {isSetSupported: () => supported}}, AsyncResultStatus: {Failed: 'failed'}}, document: {getElementById: id => ({'nerd-auto-open': control, 'nerd-auto-open-control': container, 'nerd-auto-open-status': status})[id]}};
    vm.runInNewContext(source + '\nglobalThis.configure = configureAutoOpen;', context);
    return {context, control, container, status, value: () => saved, saves: () => saves};
}
test('Enables and saves auto-open for an untagged DORK document', async () => {
    const f = fixture(); await f.context.configure();
    assert.equal(f.value(), true); assert.equal(f.saves(), 1); assert.equal(f.control.checked, true);
});
test('Preserves an explicit opt-out and existing enabled setting without rewriting', async () => {
    for (const value of [true, false]) {
        const f = fixture(value); await f.context.configure();
        assert.equal(f.value(), value); assert.equal(f.control.checked, value); assert.equal(f.saves(), 0);
    }
});
test('User can disable and re-enable automatic opening', async () => {
    const f = fixture(true); await f.context.configure();
    f.control.checked = false; await f.control.onchange(); assert.equal(f.value(), false);
    f.control.checked = true; await f.control.onchange(); assert.equal(f.value(), true);
});
test('Unsupported clients do not write document settings', async () => {
    const f = fixture(undefined, false); await f.context.configure();
    assert.equal(f.saves(), 0); assert.equal(f.container.hidden, true);
});
test('Settings save failure restores the previous value and does not throw', async () => {
    const f = fixture(undefined, true, true); await f.context.configure();
    assert.equal(f.value(), undefined); assert.equal(f.control.checked, false); assert.match(f.status.textContent, /could not save/);
});
