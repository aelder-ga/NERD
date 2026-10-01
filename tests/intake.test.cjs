const {test}=require('node:test'); const assert=require('node:assert/strict'); const vm=require('node:vm'); const fs=require('node:fs');
const source=fs.readFileSync(require('node:path').join(__dirname,'../src/taskpane/taskpane.js'),'utf8');
const save=source.substring(source.indexOf('async function saveMetadata()'),source.indexOf('function validateRequiredMetadata()'));
function setup(locked=false) {
 const events=[];const ctx={intakeMode:true,isSaving:false,currentDriveItem:{name:'Old.pdf'},currentListItem:{id:'1'},currentDocumentUrl:'url',dorkSite:{id:'s'},dorkDocumentsLibrary:{id:'l'},window:{sessionStorage:{}},document:{getElementById:id=>id==='title'?{value:'New'}:{disabled:false}},console,
 setSaveStatus(){},validateRequiredMetadata(){},buildDesiredFilename:()=> 'New.pdf',materializeNewTags:async()=>{},resolvePeople:async()=>{},resolveDesiredContentType:async()=>null,buildSharePointFormValues:()=>[],assertIntakeFileClosed:async()=>{events.push('closed');if(locked)throw Error('locked');},requestNumbering:async()=>{events.push('reserve');return null;},validateUpdateListItem:async v=>events.push(v[0]?.FieldName==='FileLeafRef'?'rename':'metadata'),verifySavedMetadata:async()=>events.push('verify'),rememberDocumentLocation(){},refreshIntakeFiles:async()=>events.push('list'),renderIntakeFiles(){},getErrorMessage:e=>e.message};
 vm.runInNewContext(save+'\nglobalThis.save=saveMetadata;',ctx);return {ctx,events};
}
test('GEEK saves and verifies PDF metadata without an Office host or document writes',async()=>{const {ctx,events}=setup();await ctx.save();assert.deepEqual(events,['closed','reserve','metadata','verify','rename','verify','list']);});
test('An open file cannot allocate a number or write metadata through intake',async()=>{const {ctx,events}=setup(true);ctx.console={error(){}};await ctx.save();assert.deepEqual(events,['closed']);});
test('Intake loading and selection contain no numbering or Office writes',()=>{const block=source.split('/* BROWSER INTAKE:')[1];assert.doesNotMatch(block,/requestNumbering|saveHostDocument|syncHostMetadata|Office\./);});
test('Intake preserves arbitrary extensions and extensionless files',()=>{const block=source.substring(source.indexOf('function buildDesiredFilename()'),source.indexOf('async function verifySavedMetadata('));for(const name of ['x.pdf','x.png','x']){const ctx={intakeMode:true,currentDriveItem:{name},document:{getElementById:()=>({value:'Updated'})},getFieldValue:()=> 'Old'};vm.runInNewContext(block+'\nglobalThis.filename=buildDesiredFilename;',ctx);assert.equal(ctx.filename(),'Updated'+(name.includes('.')?name.slice(name.lastIndexOf('.')):''));}});
const contentTypeBlock=source.substring(source.indexOf('async function resolveDesiredContentType()'),source.indexOf('function buildSharePointFormValues('));
test('Uploaded Visual uses a matching content type or general DORK Document, never default Procedure',async()=>{
 for(const types of [[{id:'general',name:'DORK Document'},{id:'procedure',name:'DORK Procedure'}],[{id:'visual',name:'DORK Visual'},{id:'general',name:'DORK Document'}]]){
  const ctx={intakeMode:true,currentDriveItem:{name:'map.pdf'},contentTypeMap:{},dorkSite:{id:'s'},dorkDocumentsLibrary:{id:'l'},GRAPH_ROOT:'graph',getSelectedTerm:()=>({label:'Visual'}),normalizeName:v=>String(v).toLowerCase().replace(/[^a-z0-9]/g,''),graphGetAll:async()=>types};
  vm.runInNewContext(contentTypeBlock+'\nglobalThis.resolve=resolveDesiredContentType;',ctx);assert.equal((await ctx.resolve()).id,types[0].id);
 }
});
test('Missing suitable intake content type fails instead of retaining a mismatched default',async()=>{
 const ctx={intakeMode:true,currentDriveItem:{name:'map.pdf'},contentTypeMap:{},dorkSite:{id:'s'},dorkDocumentsLibrary:{id:'l'},GRAPH_ROOT:'graph',getSelectedTerm:()=>({label:'Visual'}),normalizeName:v=>String(v).toLowerCase().replace(/[^a-z0-9]/g,''),graphGetAll:async()=>[{id:'p',name:'DORK Procedure'}]};
 vm.runInNewContext(contentTypeBlock+'\nglobalThis.resolve=resolveDesiredContentType;',ctx);await assert.rejects(ctx.resolve(),/DORK Document/);
});
const selectionBlock=source.substring(source.indexOf('async function selectIntakeFile('),source.indexOf('async function assertIntakeFileClosed('));
test('First selection of a new upload preserves the latest entered draft, not upload-time values',async()=>{
 const latest={values:{lifecycle:'Active'},title:'My Floorplans'},older={values:{lifecycle:'Draft'}};let applied;
 const elements={};const ctx={isSaving:false,connecting:false,intakeUploading:false,currentDriveItem:null,currentListItem:null,currentDocumentFields:null,currentDocumentUrl:null,intakeDefaults:older,intakeUploadedIds:new Set(['file']),dorkSite:{id:'s'},dorkDocumentsLibrary:{id:'l'},GRAPH_ROOT:'graph',captureIntakeDefaults:()=>latest,document:{getElementById:id=>elements[id] ||= {}},setSaveStatus(){},graphGet:async()=>({id:'file',file:{},name:'Floorplans.pdf',webUrl:'url'}),getCurrentListItem:async()=>({fields:{}}),hydrateAllControls:async()=>{},getFieldValue:()=>'',applyIntakeDefaults:async(v,title)=>{applied={v,title};},enableSave(){}};
 vm.runInNewContext(selectionBlock+'\nglobalThis.select=selectIntakeFile;',ctx);await ctx.select({id:'1'});assert.equal(applied.v,latest);assert.equal(applied.title,true);
 ctx.currentDriveItem=null;ctx.getFieldValue=()=> 'DORK-0004';applied=null;await ctx.select({id:'1'});assert.equal(applied,null,'Saved metadata is authoritative on reopen');
});
