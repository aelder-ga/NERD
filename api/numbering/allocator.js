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
// Storage supplies atomic unique DocumentKey / DorkID constraints. No in-memory counter.
async function reserve(store, key, location, offset) {
  // The explicitly bound Framework occupies a registry row but not an ordinary
  // sequence number. Resolve its boundary before creating any reservation.
  const framework = await store.findFramework();
  const boundary = framework ? Number(framework.id) : null;
  if (framework && (!Number.isSafeInteger(boundary) || boundary < 1 || framework.fields.DorkID !== 'DORK-0001')) {
    throw new Error('Invalid Framework registry boundary.');
  }
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
  const original = identifier(row.id, offset);
  const expected = boundary !== null && Number(row.id) > boundary
    ? identifier(Number(row.id) - 1, offset) : original;
  // 0001 can only be explicitly provisioned by an administrator.
  if (row.fields.DorkID === 'DORK-0001') return row;
  if (row.fields.DorkID && row.fields.DorkID !== expected && row.fields.DorkID !== original) throw new Error('Registry sequence mismatch; administrator recovery required.');
  if (!row.fields.DorkID) {
    try { await store.patch(row, { DorkID: expected }); }
    catch (error) {
      const committed = await store.find(key);
      if (committed?.fields.DorkID !== expected) throw error;
    }
    row = await store.find(key);
    if (row?.fields.DorkID !== expected) throw new Error('Registry assignment was not verified.');
  }
  return row;
}
module.exports = { guid, documentKey, identifier, reserve };
