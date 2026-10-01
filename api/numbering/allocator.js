'use strict';
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function documentKey(siteId, uniqueId) {
  if (!guid.test(uniqueId || '')) throw new Error('Invalid server file identity.');
  return `${siteId}:${uniqueId.toLowerCase()}`;
}
function identifier(itemId, offset) {
  const number = Number(itemId) + offset;
  if (!Number.isSafeInteger(Number(itemId)) || Number(itemId) < 1 || !Number.isSafeInteger(offset) || offset < 1 || !Number.isSafeInteger(number)) {
    throw new Error('Invalid registry sequence.');
  }
  return `DORK-${String(number).padStart(4, '0')}`;
}
// Storage supplies atomic unique DocumentKey / DorkId constraints. No in-memory counter.
async function reserve(store, key, location, offset) {
  let row = await store.find(key);
  if (!row) {
    try { row = await store.create({ Title: key, DocumentKey: key, AllocationState: 'Reserved', CurrentLocation: location }); }
    catch (error) {
      // A failed create may have committed. Read back before deciding to fail.
      row = await store.find(key);
      if (!row) throw error;
    }
  }
  if (row.fields.DocumentKey !== key) throw new Error('Registry identity mismatch.');
  const expected = identifier(row.id, offset);
  // 0001 can only be explicitly provisioned by an administrator.
  if (row.fields.DorkId === 'DORK-0001') return row;
  if (row.fields.DorkId && row.fields.DorkId !== expected) throw new Error('Registry sequence mismatch; administrator recovery required.');
  if (!row.fields.DorkId) {
    try { await store.patch(row, { DorkId: expected }); }
    catch (error) {
      const committed = await store.find(key);
      if (committed?.fields.DorkId !== expected) throw error;
    }
    row = await store.find(key);
    if (row?.fields.DorkId !== expected) throw new Error('Registry assignment was not verified.');
  }
  return row;
}
module.exports = { guid, documentKey, identifier, reserve };
