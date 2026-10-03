'use strict';
const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const {graphUrl, uploadUrl} = require('../src/security/requests');
const proxy = require('../api/sharepoint');
const origin = 'https://rocktwpnet.sharepoint.com/sites/DORK/_api/';
const library = "web/lists(guid'dd8ac8ef-3be9-4fb8-b7cb-6ac0f56d2b03')/items(14)";
test('Graph token destinations reject foreign hosts, userinfo, HTTP and encoded escapes',()=>{
 assert.equal(graphUrl('https://graph.microsoft.com/v1.0/me?$select=id'),'https://graph.microsoft.com/v1.0/me?$select=id');
 for(const url of ['https://evil.example/v1.0/me','https://graph.microsoft.com.evil.example/v1.0/me','https://user@graph.microsoft.com/v1.0/me','http://graph.microsoft.com/v1.0/me','https://graph.microsoft.com/beta/me','https://graph.microsoft.com/v1.0/%2fme','https://graph.microsoft.com/v1.0/me#fragment']) assert.throws(()=>graphUrl(url));
});
test('Graph paging cannot send tokens to attacker-controlled destinations or loop forever',async()=>{
 const source=fs.readFileSync(require('node:path').join(__dirname,'../src/taskpane/taskpane.js'),'utf8');
 const block=source.split('/* GRAPH */')[1].split('/* NERD API')[0];
 const calls=[];const ctx={graphUrl,acquireToken:async()=>{calls.push('token');return 'test-only';},GRAPH_SCOPES:[],fetch:async(url,options)=>{calls.push({url,options});return {ok:true,text:async()=>JSON.stringify({value:[], '@odata.nextLink':'https://evil.example/page'})};}};
 vm.runInNewContext(block+'\nglobalThis.all=graphGetAll;',ctx);
 await assert.rejects(ctx.all('https://graph.microsoft.com/v1.0/me'),/only send/);
 assert.equal(calls.filter(c=>c==='token').length,1);
 assert.equal(calls[1].options.redirect,'error');
 ctx.fetch=async()=>({ok:true,text:async()=>JSON.stringify({value:[],'@odata.nextLink':'https://graph.microsoft.com/v1.0/me'})});
 await assert.rejects(ctx.all('https://graph.microsoft.com/v1.0/me'),/paging/);
});
test('Upload sessions reject non-Microsoft destinations before sending file bytes',()=>{
 assert.equal(uploadUrl('https://rocktwpnet.sharepoint.com/session?token=placeholder'),'https://rocktwpnet.sharepoint.com/session?token=placeholder');
 for(const url of ['https://upload.example/file','https://sharepoint.com.evil.example/file','http://rocktwpnet.sharepoint.com/file','https://user@rocktwpnet.sharepoint.com/file','https://rocktwpnet.sharepoint.com:444/file']) assert.throws(()=>uploadUrl(url));
});
test('Proxy never fetches without a token or for arbitrary methods, hosts, paths, batches or non-JSON bodies',async()=>{
 const old=global.fetch;let calls=0;global.fetch=async()=>{calls++;throw Error('must not run');};
 try {
  for(const [url,method,headers,token] of [
   [origin+library+'/File','GET',{},null],
   ['https://evil.example/_api/web','GET',{},'Bearer test-only'],
   [origin+'web','GET',{},'Bearer test-only'],
   [origin+'$batch','POST',{'Content-Type':'multipart/mixed'},'Bearer test-only'],
   [origin+library+'/ValidateUpdateListItem','DELETE',{},'Bearer test-only'],
   [origin+library+'/ValidateUpdateListItem','POST',{'Content-Type':'text/plain'},'Bearer test-only'],
   ['https://rocktwpnet.sharepoint.com/sites/OTHER/_api/web','GET',{},'Bearer test-only'],
   [origin+"web/lists(guid'00000000-0000-0000-0000-000000000000')/items(14)/ValidateUpdateListItem",'POST',{},'Bearer test-only'],
  ]) {const ctx={};await proxy(ctx,{headers:token?{'x-nerd-sharepoint-authorization':token}:{},body:{url,method,headers}});assert.ok(ctx.res.status>=400);}
  assert.equal(calls,0);
 } finally {global.fetch=old;}
});
test('Metadata proxy preserves only the dedicated user token and rejects redirects',async()=>{
 const old=global.fetch;const calls=[];
 global.fetch=async(url,options)=>{calls.push({url,options});return {ok:true,status:200,text:async()=>'{}'};};
 try {
  for(const [path,method] of [['web/ensureuser','POST'],['web/GetUserById(3)','GET'],[library+'/File/Properties','GET'],[library+'/File?$select=CheckOutType','GET'],[library+'/ValidateUpdateListItem','POST']]) {
   const ctx={};await proxy(ctx,{headers:{'x-nerd-sharepoint-authorization':'Bearer test-only',authorization:'Bearer wrong'},body:{url:origin+path,method,headers:{Authorization:'Bearer wrong','X-HTTP-Method':'DELETE'},body:method==='POST'?{}:null}});assert.equal(ctx.res.status,200);
  }
  assert.equal(calls.length,5);for(const c of calls){assert.equal(c.options.headers.Authorization,'Bearer test-only');assert.equal(c.options.headers['X-HTTP-Method'],undefined);assert.equal(c.options.redirect,'error');}
 } finally {global.fetch=old;}
});
test('Proxy network errors never expose bearer tokens or request bodies',async()=>{
 const old=global.fetch;global.fetch=async()=>{throw Error('sensitive-upstream-placeholder');};
 try {const ctx={};await proxy(ctx,{headers:{'x-nerd-sharepoint-authorization':'Bearer test-only'},body:{url:origin+library+'/File',method:'GET'}});assert.equal(ctx.res.status,502);assert.ok(!JSON.stringify(ctx.res).includes('sensitive-upstream-placeholder'));}finally{global.fetch=old;}
});
