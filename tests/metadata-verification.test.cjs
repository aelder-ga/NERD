const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.js'), 'utf8');
const block = source.slice(source.indexOf('async function verifySavedMetadata('), source.indexOf('/* UI */'));
const fields = {
    Title: 'Test', Domain: 'NETWORK|id1', Function: 'NETWORK:Firewall|id2', 'Document Type': 'Guide|id3',
    'System / Platform': ['Meraki', 'Webex'], Collection: ['Operations'], Tags: ['Voice', 'Firewall'],
    Audience: ['Staff', 'IT'], Classification: 'Internal', Lifecycle: 'Draft',
    'Information Source': 'Vendor documentation', 'Last Reviewed': '2026-09-30T00:00:00Z', 'Next Review': '2027-09-30T00:00:00Z',
    OwnerLookupId: '6', ResponsibleLookupId: '7', SecondaryLookupId: '8', ContentTypeId: '0x010100ABC123',
};
async function verify(overrides = {}, contentType) {
    let hydrated = false;
    const input = {title: 'Test', classification: 'Internal', lifecycle: 'Draft', 'information-source': 'Vendor documentation', 'last-reviewed': '2026-09-30', 'next-review': '2027-09-30'};
    const context = {
        currentDriveItem: {id: 'file', name: 'Copy.docx'}, currentListItem: {}, currentDocumentFields: {},
        getDriveItemById: async () => ({id: 'file', name: 'Copy.docx'}),
        getCurrentListItem: async () => ({fields: {...fields, ...overrides}, contentType}),
        document: {getElementById: id => ({value: input[id]})},
        requireColumn: name => ({name}), normalizeName: value => String(value || '').toLowerCase().replace(/[^a-z0-9]/g, ''),
        getSelectedTerm: id => ({label: {domain: 'NETWORK', function: 'Firewall', 'document-type': 'Guide'}[id]}),
        pickerSelections: {'system-platform': [{label: 'Webex'}, {label: 'Meraki'}], collection: [{label: 'Operations'}], tags: [{label: 'Firewall'}, {label: 'Voice'}], audience: [{label: 'IT'}, {label: 'Staff'}]},
        personSelections: {owner: {sharePointLookupId: '6'}, responsible: {sharePointLookupId: '7'}, secondary: {sharePointLookupId: '8'}},
        hydrateAllControls: async () => {hydrated = true;},
    };
    vm.runInNewContext(block + '\nfunction getFieldValue(name) {return currentDocumentFields[name];}\nglobalThis.verify = verifySavedMetadata;', context);
    await context.verify({id: '0x010100ABC'}, 'Copy.docx');
    return hydrated;
}
test('All dimensions verify despite taxonomy order and SharePoint function path', async () => {
    assert.equal(await verify(), true);
});
test('Each lost or changed dimension blocks saved-and-verified', async () => {
    for (const [name, value, label] of [
        ['System / Platform', ['Meraki'], 'System / Platform'], ['Collection', [], 'Collection'], ['Tags', ['Voice'], 'Tags'],
        ['Audience', ['Staff'], 'Audience'], ['OwnerLookupId', '99', 'Owner'], ['ResponsibleLookupId', null, 'Responsible'], ['SecondaryLookupId', '99', 'Secondary'],
        ['Information Source', 'Wrong', 'Information Source'], ['Last Reviewed', null, 'Last Reviewed'], ['Next Review', '2027-10-01', 'Next Review'],
        ['Domain', 'Wrong', 'Domain'], ['Function', 'Voice', 'Function'], ['Document Type', 'Standard', 'Document Type'],
        ['Classification', 'Public', 'Classification'], ['Lifecycle', 'Approved', 'Lifecycle'], ['Title', 'Wrong', 'Title'], ['ContentTypeId', '', 'Content Type'],
    ]) {
        await assert.rejects(() => verify({[name]: value}), error => error.message.includes(label), name);
    }
});
test('Clearing selections requires SharePoint to remove every previous value', () => {
    const context = {normalizeName: value => String(value).toLowerCase()};
    vm.runInNewContext(source.slice(source.indexOf('function sameValueSet('), source.indexOf('/* FIELD PARSING */')) + '\nglobalThis.check = verifyLabelSet;', context);
    const failures = [];
    context.check(failures, 'Tags', ['Old tag'], []);
    assert.deepEqual(failures, ['Tags']);
    const empty = [];
    context.check(empty, 'Tags', [], []);
    assert.deepEqual(empty, []);
});

test('Content type verifies from the documented Graph listItem property when fields omit the ID', async () => {
    assert.equal(await verify({ContentTypeId: undefined}, {id: '0x010100ABC123', name: 'DORK Guide'}), true);
});
test('An incorrect listItem content type fails even when the field bag contains the expected ID', async () => {
    await assert.rejects(() => verify({}, {id: '0x010100DEF', name: 'DORK Standard'}), error => error.message.includes('Content Type') && error.message.includes('DORK Standard'));
});
