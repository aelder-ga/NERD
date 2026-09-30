const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.js'), 'utf8');
const block = source.split('/* AUTHENTICATION */')[1].split('const dialogCacheKey')[0];
function setup({platform = 'OfficeOnline', cached = false, hostError = false, ssoError = false} = {}) {
    const calls = [];
    const context = {
        console: {warn() {}}, currentUser: null, dialogTokens: new Map(),
        Office: {PlatformType: {OfficeOnline: 'OfficeOnline'}, context: {platform}, auth: {getAuthContext: async () => {if (hostError) throw Error('Unsupported'); return {userPrincipalName: 'aelder@rocktwp.net'};}}},
        msalInstance: {
            getAccountByUsername: name => ({username: name}),
            acquireTokenSilent: async request => {calls.push({kind:'cache',request}); if(cached) return {accessToken:'cached'}; throw {errorCode:'no_account_error'};},
            ssoSilent: async request => {calls.push({kind:'sso',request}); if(ssoError) throw {errorCode:'interaction_required'}; return {accessToken:'silent'};},
        },
        acquireDialogToken: async () => {calls.push({kind:'dialog'}); return 'fallback';},
    };
    vm.runInNewContext(block+'\nglobalThis.acquire = acquireToken;', context);
    return {context,calls};
}
test('Word web uses Office identity for silent Graph and SharePoint requests without a dialog', async () => {
    const {context,calls} = setup();
    for (const scopes of [['User.Read'],['https://rocktwpnet.sharepoint.com/AllSites.Write']]) assert.equal(await context.acquire(scopes),'silent');
    assert.equal(calls.filter(x=>x.kind==='dialog').length,0);
    for(const call of calls.filter(x=>x.kind==='sso')) assert.equal(call.request.loginHint,'aelder@rocktwp.net');
});
test('MSAL cached token avoids a new SSO request', async () => {
    const {context,calls} = setup({cached:true});
    assert.equal(await context.acquire(['User.Read']),'cached');
    assert.deepEqual(calls.map(x=>x.kind),['cache']);
});
test('Interaction required retains the working dialog fallback', async () => {
    const {context,calls} = setup({ssoError:true});
    assert.equal(await context.acquire(['User.Read']),'fallback');
    assert.deepEqual(calls.map(x=>x.kind),['cache','sso','dialog']);
});
test('Desktop and unavailable Office identity retain the existing fallback', async () => {
    for(const options of [{platform:'PC'},{hostError:true}]) {
        const {context,calls} = setup(options);
        assert.equal(await context.acquire(['User.Read']),'fallback');
        assert.deepEqual(calls.map(x=>x.kind),['cache','dialog']);
    }
});

test('Silent resource tokens are reused through Save and reacquired near expiry', async () => {
    const {context,calls} = setup({cached:true});
    context.msalInstance.acquireTokenSilent = async request => {
        calls.push({kind:'cache',request});
        return {accessToken:'resource-token',expiresOn:new Date(Date.now()+3600000)};
    };
    const scopes = ['https://rocktwpnet.sharepoint.com/AllSites.Write'];
    assert.equal(await context.acquire(scopes),'resource-token');
    assert.equal(await context.acquire(scopes),'resource-token');
    assert.equal(calls.length,1);
    context.dialogTokens.get('sharepoint').expires = Date.now()+60000;
    assert.equal(await context.acquire(scopes),'resource-token');
    assert.equal(calls.length,2);
});
