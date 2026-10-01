const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const source = fs.readFileSync(require('node:path').join(__dirname, '../src/taskpane/document-location.js'), 'utf8').replace(/export /g, '');
function setup() {
 const context = { URL }; vm.runInNewContext(source + '\nglobalThis.api={rememberDocumentLocation,resolveDocumentLocation};', context);
 const values = new Map(); const storage = { getItem: k => values.get(k), setItem: (k,v) => values.set(k,v) };
 return { ...context.api, storage };
}
const oldUrl='https://example.com/sites/DORK/Shared%20Documents/Old.docx';
const item={id:'stable-id',webUrl:'https://example.com/sites/DORK/Shared%20Documents/DORK-0002%20New.docx',sharepointIds:{listId:'library'}};
const missing=()=>{throw Object.assign(Error('Missing'),{status:404});};
test('reload after rename resolves known old URL by stable identity',async()=>{
 const api=setup(); api.rememberDocumentLocation(api.storage,'site','library',oldUrl,item);
 const resolved=await api.resolveDocumentLocation({storage:api.storage,siteId:'site',libraryId:'library',url:oldUrl,getByPath:missing,getById:async id=>{assert.equal(id,item.id);return item;}});
 assert.equal(resolved.id,item.id);
});
test('copy or reused filename resolves actual path first, never the original mapping',async()=>{
 const api=setup();api.rememberDocumentLocation(api.storage,'site','library',oldUrl,item);
 const copy={...item,id:'copy-id'};
 const result=await api.resolveDocumentLocation({storage:api.storage,siteId:'site',libraryId:'library',url:oldUrl,getByPath:async()=>copy,getById:()=>assert.fail('Must use actual path')});
 assert.equal(result.id,'copy-id');
});
test('unknown paths and authorization failures never recover through unrelated files',async()=>{
 const api=setup();api.rememberDocumentLocation(api.storage,'site','library',oldUrl,item);
 for(const [url,status] of [[oldUrl,403],['https://example.com/copy.docx',404]]) {
 await assert.rejects(api.resolveDocumentLocation({storage:api.storage,siteId:'site',libraryId:'library',url,getByPath:()=>{throw Object.assign(Error('Denied'),{status});},getById:()=>assert.fail('Must not fallback')}));
 }
});
test('cached identity outside configured library is rejected',async()=>{
 const api=setup();api.rememberDocumentLocation(api.storage,'site','library',oldUrl,item);
 await assert.rejects(api.resolveDocumentLocation({storage:api.storage,siteId:'site',libraryId:'library',url:oldUrl,getByPath:missing,getById:async()=>({...item,sharepointIds:{listId:'other'}})}),/does not belong/);
});
