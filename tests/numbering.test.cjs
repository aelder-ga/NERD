'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { identifier, documentKey, reserve } = require('../api/numbering/allocator');
function storage() {
 const rows = new Map(); let sequence = 0;
 return { rows,
 async findFramework() { return structuredClone([...rows.values()].find(row => row.fields.DorkID === 'DORK-0001') || null); },
 async find(key) { return rows.has(key) ? structuredClone(rows.get(key)) : null; },
 async create(fields) { if (rows.has(fields.DocumentKey)) throw Error('Unique constraint'); const row = { id: String(++sequence), fields: structuredClone(fields) }; rows.set(fields.DocumentKey, row); return structuredClone(row); },
 async patch(row, fields) { Object.assign(rows.get(row.fields.DocumentKey).fields, fields); } };
}
test('sequence reserves 0001 and supports more than four digits', () => {
 assert.equal(identifier('1', 1), 'DORK-0002'); assert.equal(identifier('10000', 1), 'DORK-10001');
 assert.throws(() => identifier('0', 1)); assert.throws(() => identifier('1', 0));
});
test('concurrent requests share one reservation and different files get distinct numbers', async () => {
 const store = storage(); const rows = await Promise.all(Array.from({ length: 10 }, () => reserve(store, 'a', '/a', 1)));
 assert.equal(store.rows.size, 1); assert.ok(rows.every(row => row.fields.DorkID === 'DORK-0002'));
 assert.equal((await reserve(store, 'b', '/b', 1)).fields.DorkID, 'DORK-0003');
});
test('ambiguous committed create and patch are recovered; rename retains number', async () => {
 const store = storage(), create = store.create, patch = store.patch;
 store.create = async fields => { await create(fields); throw Error('Lost response'); };
 store.patch = async (row, fields) => { await patch(row, fields); throw Error('Lost response'); };
 assert.equal((await reserve(store, 'a', '/a', 1)).fields.DorkID, 'DORK-0002');
 assert.equal((await reserve(store, 'a', '/renamed', 1)).fields.DorkID, 'DORK-0002');
});
test('uncommitted failures do not report allocation success', async () => {
 const store = storage(); store.create = async () => { throw Error('Forbidden'); };
 await assert.rejects(reserve(store, 'a', '/a', 1), /Forbidden/);
});
test('Framework mapping retained and unexpected sequence fails closed', async () => {
 const store = storage(); store.rows.set('f', { id: '7', fields: { DocumentKey: 'f', DorkID: 'DORK-0001' } });
 assert.equal((await reserve(store, 'f', '/f', 1)).fields.DorkID, 'DORK-0001');
 store.rows.set('bad', { id: '8', fields: { DocumentKey: 'bad', DorkID: 'DORK-0010' } });
 await assert.rejects(reserve(store, 'bad', '/bad', 1), /sequence mismatch/);
});
test('file identity validates GUID; missing caller token denied while disabled', async () => {
 assert.throws(() => documentKey('site', "x' or 1 eq 1"));
 const handler = require('../api/numbering'); const ctx = {};
 await handler(ctx, { headers: {}, body: {} }); assert.equal(ctx.res.status, 401);
});
test('read-only caller cannot reach application credential exchange', async () => {
 const handler = require('../api/numbering'), saved = { ...process.env }, oldFetch = global.fetch;
 const g = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
 Object.assign(process.env, { NERD_NUMBERING_ENABLED: 'true', NERD_NUMBERING_SITE_ID: `rocktwpnet.sharepoint.com,${g},${g}`, NERD_NUMBERING_REGISTRY_LIST_ID: g, NERD_NUMBERING_DOCUMENTS_LIST_ID: g, NERD_NUMBERING_TENANT_ID: g, NERD_NUMBERING_CLIENT_ID: g, NERD_NUMBERING_CLIENT_SECRET: 'test-only', NERD_NUMBERING_SEQUENCE_OFFSET: '1' });
 let calls = 0;
 global.fetch = async url => { calls++; assert.match(url, /EffectiveBasePermissions$/); return { ok: true, json: async () => ({ Low: '1', High: '0' }) }; };
 try { const ctx = {}; await handler(ctx, { headers: { 'x-nerd-sharepoint-authorization': 'Bearer test' }, body: { itemId: 14, action: 'reserve' } }); assert.equal(ctx.res.status, 403); assert.equal(calls, 1); }
 finally { global.fetch = oldFetch; for (const key of Object.keys(process.env)) if (!(key in saved)) delete process.env[key]; Object.assign(process.env, saved); }
});

test('late Framework binding consumes no ordinary number; retries preserve all assigned IDs', async () => {
 const store = storage();
 for (const [key, expected] of [['a', 'DORK-0002'], ['b', 'DORK-0003'], ['c', 'DORK-0004']]) {
  assert.equal((await reserve(store, key, '/'+key, 1)).fields.DorkID, expected);
 }
 const framework = await store.create({ DocumentKey: 'framework', DorkID: 'DORK-0001' });
 assert.equal(framework.id, '4');
 const rows = await Promise.all(Array.from({length: 8}, () => reserve(store, 'next', '/next', 1)));
 assert.ok(rows.every(row => row.fields.DorkID === 'DORK-0005'));
 assert.equal((await reserve(store, 'later', '/later', 1)).fields.DorkID, 'DORK-0006');
 for (const [key, expected] of [['a', 'DORK-0002'], ['b', 'DORK-0003'], ['c', 'DORK-0004'], ['framework', 'DORK-0001'], ['next', 'DORK-0005']]) {
  assert.equal((await reserve(store, key, '/renamed', 1)).fields.DorkID, expected);
 }
});
test('already assigned original mapping after Framework remains permanent', async () => {
 const store = storage();
 await store.create({DocumentKey: 'framework', DorkID: 'DORK-0001'});
 await store.create({DocumentKey: 'old', DorkID: 'DORK-0003'});
 assert.equal((await reserve(store, 'old', '/old', 1)).fields.DorkID, 'DORK-0003');
});
test('Framework lookup failure creates no reservation', async () => {
 const store = storage(); store.findFramework = async () => { throw Error('Unavailable'); };
 await assert.rejects(reserve(store, 'new', '/new', 1), /Unavailable/);
 assert.equal(store.rows.size, 0);
});
