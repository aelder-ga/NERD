const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const context={};vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../src/taskpane/document-metadata.js'),'utf8').replace('export async function','async function')+'\nglobalThis.refresh=syncDocumentMetadata;',context);
function control(tag,text,cannotEdit=false) {return {tag,text,cannotEdit,insertText(value,location){assert.equal(this.cannotEdit,false);assert.equal(location,'Replace');this.text=value;}};}
test('Updates every matching occurrence, joins provided multivalues and preserves body and locks',async()=>{
 const items=[control('DORK_Title','Old',true),control('DORK_Title','Old'),control('DORK_System','Entra ID'),control('DORK_Owner','Old Owner'),control('Body','User content')];
 let syncs=0;const word={run:fn=>fn({document:{contentControls:{items,load(){}}},sync:async()=>syncs++})};
 assert.equal(await context.refresh({DORK_Title:'New',DORK_System:'Entra ID; Webex',DORK_Owner:''},word),4);
 assert.deepEqual(items.map(x=>x.text),['New','New','Entra ID; Webex','','User content']);assert.equal(items[0].cannotEdit,true);
 assert.equal(await context.refresh({DORK_Title:'New'},word),0);assert.equal(syncs,5);
});
test('Restores locks and reports Word write failures without claiming an update',async()=>{
 const item=control('DORK_Title','Old',true);item.insertText=()=>{throw Error('Read only');};
 const word={run:fn=>fn({document:{contentControls:{items:[item],load(){}}},sync:async()=>{}})};
 await assert.rejects(context.refresh({DORK_Title:'New'},word),/Read only/);assert.equal(item.cannotEdit,true);
});
