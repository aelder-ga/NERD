'use strict';
const { guid, documentKey, reserve } = require('./allocator');
const SP = 'https://rocktwpnet.sharepoint.com/sites/DORK';
const GRAPH = 'https://graph.microsoft.com/v1.0';
function problem(status, message) { return Object.assign(new Error(message), { status }); }
async function request(url, options = {}) {
  const response = await fetch(url, { ...options, redirect: 'error', signal: AbortSignal.timeout(20000) });
  const data = await response.json().catch(() => null);
  if (!response.ok) throw problem(response.status, 'The numbering service request failed.');
  return data;
}
module.exports = async function (context, req) {
  const reply = (status, body) => { context.res = { status, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }, body }; };
  const authorization = req.headers?.['x-nerd-sharepoint-authorization'];
  if (typeof authorization !== 'string' || !/^Bearer \S+$/i.test(authorization)) return reply(401, { error: 'Missing SharePoint bearer token.' });
  // Deployment prepares the integration; activation requires completed inventory and acceptance gates.
  if (process.env.NERD_NUMBERING_ENABLED !== 'true') return reply(200, { enabled: false });
  try {
    const env = process.env;
    const site = env.NERD_NUMBERING_SITE_ID;
    const registry = env.NERD_NUMBERING_REGISTRY_LIST_ID;
    const library = env.NERD_NUMBERING_DOCUMENTS_LIST_ID;
    const tenant = env.NERD_NUMBERING_TENANT_ID;
    const client = env.NERD_NUMBERING_CLIENT_ID;
    const secret = env.NERD_NUMBERING_CLIENT_SECRET;
    const offset = Number(env.NERD_NUMBERING_SEQUENCE_OFFSET);
    if (!site?.startsWith('rocktwpnet.sharepoint.com,') || ![registry, library, tenant, client].every(x => guid.test(x || '')) || !secret || !Number.isSafeInteger(offset) || offset < 1) throw problem(503, 'Numbering configuration is incomplete.');
    let input = req.body;
    if (typeof input === 'string') { try { input = JSON.parse(input); } catch { throw problem(400, 'Invalid JSON.'); } }
    if (!input || !['reserve', 'applied'].includes(input.action) || !/^\d{1,10}$/.test(String(input.itemId)) || Number(input.itemId) < 1) throw problem(400, 'Invalid numbering request.');
    const endpoint = `${SP}/_api/web/lists(guid'${library}')/items(${Number(input.itemId)})`;
    const spHeaders = { Authorization: authorization, Accept: 'application/json;odata=nometadata' };
    // SharePoint validates the delegated token and calculates permissions on the actual document.
    const permissionsRaw = await request(`${endpoint}/EffectiveBasePermissions`, { headers: spHeaders });
    const permissions = permissionsRaw?.d?.EffectiveBasePermissions || permissionsRaw?.d || permissionsRaw;
    if (!/^\d+$/.test(String(permissions?.Low)) || (BigInt(permissions.Low) & 4n) === 0n) throw problem(403, 'Document edit permission is required.');
    const raw = await request(`${endpoint}?$select=Id,DORK_x0020_ID,File/UniqueId,File/ServerRelativeUrl&$expand=File`, { headers: spHeaders });
    const item = raw?.d || raw;
    if (Number(item?.Id) !== Number(input.itemId) || !item.File?.ServerRelativeUrl?.toLowerCase().startsWith('/sites/dork/shared documents/')) throw problem(403, 'The item is not a DORK document.');
    const key = documentKey(site, item.File.UniqueId);
    // Store a relative recovery location. Never silently truncate the column's 255-character limit.
    const location = item.File.ServerRelativeUrl;
    if (location.length > 255) throw problem(409, 'Registry CurrentLocation must support this document path before numbering.');
    const token = await request(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
      method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: client, client_secret: secret, grant_type: 'client_credentials', scope: 'https://graph.microsoft.com/.default' }).toString(),
    });
    if (!token?.access_token) throw problem(502, 'Numbering authentication failed.');
    const root = `${GRAPH}/sites/${encodeURIComponent(site)}/lists/${registry}/items`;
    const headers = { Authorization: `Bearer ${token.access_token}`, 'Content-Type': 'application/json' };
    const store = {
      async findFramework() {
        const query = new URLSearchParams({ '$expand': 'fields', '$filter': "fields/DorkID eq 'DORK-0001'", '$top': '2' });
        const result = await request(`${root}?${query}`, { headers });
        if (!Array.isArray(result?.value) || result.value.length !== 1) throw problem(409, 'The Framework registry binding must be restored before allocating numbers.');
        return result.value[0];
      },
      async find(value) {
        const query = new URLSearchParams({ '$expand': 'fields', '$filter': `fields/DocumentKey eq '${value}'`, '$top': '2' });
        const result = await request(`${root}?${query}`, { headers });
        if (!Array.isArray(result?.value) || result.value.length > 1) throw problem(409, 'Registry uniqueness could not be verified.');
        return result.value[0] || null;
      },
      create(fields) { return request(root, { method: 'POST', headers, body: JSON.stringify({ fields }) }); },
      patch(row, fields) {
        const etag = row.eTag || row['@odata.etag'];
        if (!etag) throw problem(502, 'Registry did not return a concurrency token.');
        return request(`${root}/${row.id}/fields`, { method: 'PATCH', headers: { ...headers, 'If-Match': etag }, body: JSON.stringify(fields) });
      },
    };
    const row = input.action === 'reserve' ? await reserve(store, key, location, offset) : await store.find(key);
    if (!row?.fields?.DorkID) throw problem(409, 'No verified number reservation exists.');
    if (input.action === 'applied') {
      if (item.DORK_x0020_ID !== row.fields.DorkID) throw problem(409, 'Document number has not been saved.');
      await store.patch(row, { AllocationState: 'Applied', CurrentLocation: location });
    }
    reply(200, { enabled: true, documentId: row.fields.DorkID });
  } catch (error) {
    // Never expose upstream token responses, credentials, document metadata or request bodies.
    const status = [400, 401, 403, 409, 503].includes(error.status) ? error.status : 502;
    reply(status, { error: status === 502 ? 'Numbering could not be completed. Retry uses the same reservation.' : error.message });
  }
};
