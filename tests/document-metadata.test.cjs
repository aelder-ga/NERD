const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const context={};vm.runInNewContext(fs.readFileSync(path.join(__dirname,'../src/taskpane/document-metadata.js'),'utf8').replace('export async function','async function')+'\nglobalThis.refresh=syncDocumentMetadata;',context);
function control(tag,text,cannotEdit=false) {return {tag,text,cannotEdit,font:{name:'Aptos',size:11,bold:false,italic:false,color:'#000000',underline:'None',load(){},set(values){Object.assign(this,values);}},insertText(value,location){assert.equal(this.cannotEdit,false);assert.equal(location,'Replace');this.text=value;}};}
test('Updates every matching occurrence, joins provided multivalues and preserves body and locks',async()=>{
 const items=[control('DORK_Title','Old',true),control('DORK_Title','Old'),control('DORK_System','Entra ID'),control('DORK_Owner','Old Owner'),control('Body','User content')];
 let syncs=0;const word={run:fn=>fn({document:{contentControls:{items,load(){}}},sync:async()=>syncs++})};
 assert.equal(await context.refresh({DORK_Title:'New',DORK_System:'Entra ID; Webex',DORK_Owner:''},word),4);
 assert.deepEqual(items.map(x=>x.text),['New','New','Entra ID; Webex','','User content']);assert.equal(items[0].cannotEdit,true);
 assert.equal(await context.refresh({DORK_Title:'New'},word),0);assert.equal(syncs,6);
});
test('Restores locks and reports Word write failures without claiming an update',async()=>{
 const item=control('DORK_Title','Old',true);item.insertText=()=>{throw Error('Read only');};
 const word={run:fn=>fn({document:{contentControls:{items:[item],load(){}}},sync:async()=>{}})};
 await assert.rejects(context.refresh({DORK_Title:'New'},word),/Read only/);assert.equal(item.cannotEdit,true);
});

test('Preserves separate cover and table fonts when Word resets formatting on replacement',async()=>{
 const cover=control('DORK_Title','Old',true), table=control('DORK_Title','Old');
 Object.assign(cover.font,{name:'Aptos Display',size:32,bold:true,color:'#17365D'});
 for(const item of [cover,table]) {
   const insert=item.insertText;
   item.insertText=function(value,location){insert.call(this,value,location);Object.assign(this.font,{name:'Calibri',size:11,bold:false,color:'#000000'});};
 }
 const word={run:fn=>fn({document:{contentControls:{items:[cover,table],load(){}}},sync:async()=>{}})};
 await context.refresh({DORK_Title:'Renamed'},word);
 assert.equal(cover.font.size,32);assert.equal(cover.font.name,'Aptos Display');assert.equal(cover.font.bold,true);assert.equal(cover.font.color,'#17365D');
 assert.equal(table.font.size,11);assert.equal(table.font.bold,false);assert.equal(cover.cannotEdit,true);
});
test('Does not assign mixed-font null properties or invalid mixed size',async()=>{
 const item=control('DORK_Source','Old');Object.assign(item.font,{name:null,size:0,bold:null});
 item.font.set=values=>{assert.equal(Object.hasOwn(values,'name'),false);assert.equal(Object.hasOwn(values,'size'),false);assert.equal(Object.hasOwn(values,'bold'),false);};
 const word={run:fn=>fn({document:{contentControls:{items:[item],load(){}}},sync:async()=>{}})};
 await context.refresh({DORK_Source:'New'},word);
});
test('Repairs an already-small standalone cover title without enlarging table metadata or changing text',async()=>{
 const cover=control('DORK_Title','Already current',true),table=control('DORK_Title','Already current');
 cover.parentTableCellOrNullObject={isNullObject:true,load(){}};
 table.parentTableCellOrNullObject={isNullObject:false,load(){}};
 cover.insertText=()=>assert.fail('Unchanged title must not be replaced');
 const word={run:fn=>fn({document:{contentControls:{items:[cover,table],load(){}}},sync:async()=>{}})};
 assert.equal(await context.refresh({DORK_Title:'Already current'},word),1);
 assert.equal(cover.font.size,32);assert.equal(table.font.size,11);assert.equal(cover.cannotEdit,true);
});
