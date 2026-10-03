'use strict';
// SharePoint validates the caller's delegated token. No application credentials.
module.exports = async function (context, req) {
  const reply = (status, body) => {
    context.res = { status, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }, body };
  };
  // SWA can replace Authorization with its internal function token.
  // Only forward the caller token from this dedicated header.
  const authorization = req.headers?.["x-nerd-sharepoint-authorization"];
  if (typeof authorization !== 'string' || !/^Bearer \S+$/i.test(authorization)) {
    return reply(401, { error: 'Missing SharePoint bearer token.' });
  }
  let input = req.body;
  if (typeof input === 'string') {
    try { input = JSON.parse(input); } catch { return reply(400, { error: 'Invalid JSON.' }); }
  }
  if (!input || typeof input !== 'object' || Array.isArray(input)) {
    return reply(400, { error: 'A JSON request is required.' });
  }
  if (Buffer.byteLength(JSON.stringify(input), 'utf8') > 1024 * 1024) {
    return reply(413, { error: 'Request exceeds 1 MB.' });
  }
  let target;
  try {
    if (typeof input.url !== 'string') throw new Error();
    target = new URL(input.url);
    const decodedPath = decodeURIComponent(target.pathname);
    if (target.origin !== 'https://rocktwpnet.sharepoint.com' || target.username || target.password ||
        !target.pathname.toLowerCase().startsWith('/sites/dork/_api/') ||
        !decodedPath.toLowerCase().startsWith('/sites/dork/_api/') ||
        /[\\]/.test(decodedPath) || decodedPath.split('/').some(segment => segment === '..' || segment === '.') ||
        /%2f|%5c|%25/i.test(target.pathname)) {
      return reply(403, { error: 'NERD may only access the DORK SharePoint REST API.' });
    }
  } catch { return reply(400, { error: 'Invalid SharePoint URL.' }); }
  const method = String(input.method || 'POST').toUpperCase();
  if (!['GET', 'POST'].includes(method)) return reply(405, { error: 'Unsupported SharePoint method.' });
  // This is a metadata bridge, not a general REST tunnel (including $batch).
  const library = process.env.NERD_NUMBERING_DOCUMENTS_LIST_ID || 'dd8ac8ef-3be9-4fb8-b7cb-6ac0f56d2b03';
  if (!/^[0-9a-f-]{36}$/i.test(library)) return reply(503, { error: 'DORK library configuration is invalid.' });
  const path = target.pathname.toLowerCase();
  const itemRoot = `/sites/dork/_api/web/lists(guid'${library.toLowerCase()}')/items(`;
  const itemPath = path.startsWith(itemRoot) ? path.slice(itemRoot.length) : '';
  const allowed = method === 'POST'
    ? path === '/sites/dork/_api/web/ensureuser' || /^\d+\)\/validateupdatelistitem$/.test(itemPath)
    : /^\/sites\/dork\/_api\/web\/getuserbyid\(\d+\)$/.test(path) || /^\d+\)\/file(?:\/properties)?$/.test(itemPath);
  if (!allowed) return reply(403, { error: 'This operation is not part of the DORK metadata bridge.' });
  const headers = { Accept: 'application/json;odata=nometadata' };
  for (const [key, value] of Object.entries(input.headers || {})) {
    const normalized = key.toLowerCase();
    if (['accept', 'content-type'].includes(normalized) &&
        (typeof value !== 'string' || !/^application\/json(?:\s*;[^\r\n]*)?$/i.test(value))) {
      return reply(400, { error: 'Only JSON SharePoint requests are supported.' });
    }
    if (['content-type', 'if-match'].includes(normalized) && typeof value === 'string' && !/[\r\n]/.test(value)) {
      headers[normalized] = value;
    }
  }
  headers.Authorization = authorization;
  const body = input.body;
  if (method === 'GET' && body != null) return reply(400, { error: 'GET requests cannot have a body.' });
  if (body != null && !headers['content-type']) headers['content-type'] = 'application/json;odata=nometadata';
  try {
    const upstream = await fetch(target.href, {
      method, headers,
      body: body == null ? undefined : typeof body === 'string' ? body : JSON.stringify(body),
      redirect: 'error', signal: AbortSignal.timeout(20000),
    });
    const text = await upstream.text();
    let data = text || null;
    if (text) { try { data = JSON.parse(text); } catch {} }
    if (!upstream.ok) {
      const diagnostics = {};
      for (const name of ['www-authenticate', 'x-ms-diagnostics', 'sprequestguid', 'request-id']) {
        const value = upstream.headers.get(name);
        if (value) diagnostics[name] = value;
      }
      return reply(upstream.status, { error: 'SharePoint request failed.', status: upstream.status,
        details: { response: data, diagnostics, method, endpoint: target.pathname } });
    }
    return reply(upstream.status === 204 ? 200 : upstream.status, { ok: true, status: upstream.status, data });
  } catch (error) {
    // Don't log tokens, request bodies, URLs, or upstream metadata.
    return reply(error.name === 'TimeoutError' ? 504 : 502, { error: 'The SharePoint service could not be reached. Try again.' });
  }
};
