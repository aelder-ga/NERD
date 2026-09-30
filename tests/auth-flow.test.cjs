const assert = require('node:assert/strict');
const {test} = require('node:test');
const vm = require('node:vm');
const fs = require('node:fs');
const {webcrypto} = require('node:crypto');
const path = require('node:path');
const authSource = fs.readFileSync(path.join(__dirname, '../src/auth/auth.js'), 'utf8');
const paneSource = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.js'), 'utf8');
const origin = 'https://thankful-sky-0cc1b0210.6.azurestaticapps.net';
const storageKey = 'nerd-dialog-pkce';

async function runAuth(search, saved, responses = []) {
    const storage = new Map(saved ? [[storageKey, JSON.stringify(saved)]] : []);
    const requests = [], messages = [], redirects = [];
    let ready;
    const context = {
        URL, URLSearchParams, TextEncoder, crypto: webcrypto,
        btoa: text => Buffer.from(text, 'binary').toString('base64'),
        window: {location: {origin, search, replace: url => redirects.push(new URL(url))}},
        history: {replaceState() {}}, document: {getElementById: () => ({textContent: ''})},
        sessionStorage: {getItem: k => storage.get(k), setItem: (k,v) => storage.set(k,v), removeItem: k => storage.delete(k)},
        Office: {onReady: callback => {ready = callback;}, context: {ui: {messageParent: (message, options) => messages.push({data: JSON.parse(message), options})}}},
        fetch: async (url, options) => {
            requests.push(new URLSearchParams(options.body));
            assert.ok(responses.length, 'Unexpected token request');
            const data = responses.shift();
            return {ok: !data.error, json: async () => data};
        },
    };
    vm.runInNewContext(authSource, context);
    await ready();
    return {storage, requests, messages, redirects};
}
const savedState = () => ({nonce: 'pane-nonce', state: 'state', resource: 'graph', verifier: 'verifier', tokens: {}, started: Date.now()});
const token = (name, refresh) => ({access_token: name, expires_in: 3600, ...(refresh ? {refresh_token: refresh} : {})});

test('One authorization requests consent for both APIs and PKCE', async () => {
    const result = await runAuth('?resource=graph&nonce=pane-nonce');
    assert.equal(result.redirects.length, 1);
    const params = result.redirects[0].searchParams;
    assert.match(params.get('scope'), /graph.microsoft.com\/User.Read/);
    assert.match(params.get('scope'), /sharepoint.com\/AllSites.Write/);
    assert.match(params.get('scope'), /offline_access/);
    assert.equal(params.get('code_challenge_method'), 'S256');
    assert.ok(params.get('state'));
    assert.equal(result.messages.length, 0);
});

test('One dialog obtains both resource tokens without a second navigation', async () => {
    const result = await runAuth('?code=code&state=state', savedState(), [token('graph-token', 'refresh-placeholder'), token('sharepoint-token')]);
    assert.equal(result.redirects.length, 0);
    assert.equal(result.requests.length, 2);
    assert.equal(result.requests[0].get('grant_type'), 'authorization_code');
    assert.equal(result.requests[1].get('grant_type'), 'refresh_token');
    assert.equal(result.requests[1].get('scope'), 'https://rocktwpnet.sharepoint.com/AllSites.Write');
    assert.equal(result.messages[0].data.tokens.graph.token, 'graph-token');
    assert.equal(result.messages[0].data.tokens.sharepoint.token, 'sharepoint-token');
    assert.equal(result.messages[0].options.targetOrigin, origin);
    assert.ok(!JSON.stringify(result.messages).includes('refresh-placeholder'));
    assert.equal(result.storage.size, 0);
});

test('Additional consent stays in the same dialog with fresh PKCE state', async () => {
    const result = await runAuth('?code=code&state=state', savedState(), [token('graph-token', 'refresh-placeholder'), {error: 'interaction_required', claims: '{"access_token":{}}'}]);
    assert.equal(result.redirects.length, 1);
    const next = JSON.parse(result.storage.get(storageKey));
    assert.equal(next.resource, 'sharepoint');
    assert.notEqual(next.state, 'state');
    assert.equal(next.tokens.graph.token, 'graph-token');
    assert.equal(result.redirects[0].searchParams.get('claims'), '{"access_token":{}}');
    assert.equal(result.messages.length, 0);
    const completion = await runAuth('?code=next&state=' + next.state, next, [token('sharepoint-token')]);
    assert.equal(completion.requests.length, 1);
    assert.equal(completion.messages[0].data.tokens.graph.token, 'graph-token');
    assert.equal(completion.messages[0].data.tokens.sharepoint.token, 'sharepoint-token');
});

test('Invalid or expired callback state cannot exchange or disclose tokens', async () => {
    for (const [saved, state] of [[savedState(), 'wrong'], [{...savedState(), started: Date.now() - 700000}, 'state']]) {
        const result = await runAuth('?code=code&state=' + state, saved, []);
        assert.equal(result.requests.length, 0);
        assert.equal(result.messages.length, 0);
        assert.equal(result.storage.size, 0);
    }
});

test('Concurrent Graph and SharePoint calls use one dialog and cache both tokens', async () => {
    let callback, dialogUrl, dialogCount = 0, closed = 0;
    const handlers = {};
    const context = {
        URL, crypto: webcrypto, console: {warn() {}}, setTimeout: () => 1, clearTimeout() {},
        window: {location: {origin, href: origin + '/taskpane.html'}},
        currentUser: null,
        msalInstance: {acquireTokenSilent: async () => {throw new Error('Silent unavailable');}},
        Office: {AsyncResultStatus: {Succeeded: 'ok'}, EventType: {DialogMessageReceived: 'message', DialogEventReceived: 'event'}, context: {ui: {
            displayDialogAsync: (url, options, handler) => {dialogCount++; dialogUrl = new URL(url); callback = handler;},
        }}},
    };
    const authBlock = paneSource.split('/* AUTHENTICATION */')[1].split('/* GRAPH */')[0];
    vm.runInNewContext(authBlock + '\nglobalThis.acquire = acquireToken;', context);
    const graph = context.acquire(['User.Read']);
    const sharepoint = context.acquire(['https://rocktwpnet.sharepoint.com/AllSites.Write']);
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(dialogCount, 1);
    callback({status: 'ok', value: {close() {closed++;}, addEventHandler(name, handler) {handlers[name] = handler;}}});
    handlers.message({origin, message: JSON.stringify({
        nonce: dialogUrl.searchParams.get('nonce'),
        tokens: {
            graph: {token: 'graph-token', expires: Date.now() + 3600000},
            sharepoint: {token: 'sharepoint-token', expires: Date.now() + 3600000},
        },
    })});
    assert.equal(await graph, 'graph-token');
    assert.equal(await sharepoint, 'sharepoint-token');
    assert.equal(closed, 1);
    assert.equal(await context.acquire(['User.Read']), 'graph-token');
    assert.equal(await context.acquire(['https://rocktwpnet.sharepoint.com/AllSites.Write']), 'sharepoint-token');
    assert.equal(dialogCount, 1);
});
