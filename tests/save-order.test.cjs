const {test}=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync(require('node:path').join(__dirname,'../src/taskpane/taskpane.js'),'utf8');
const block=source.substring(source.indexOf('async function saveMetadata()'),source.indexOf('function validateRequiredMetadata()'));
test('Word fields and content save finish before rename; no Word writes follow rename',async()=>{
 const events=[],button={disabled:false};
 const ctx={isSaving:false,currentDriveItem:{id:'file',name:'Old.docx'},currentListItem:{id:'15'},currentDocumentUrl:'old',dorkSite:{id:'site'},dorkDocumentsLibrary:{id:'list'},window:{sessionStorage:{}},document:{getElementById:id=>id==='title'?{value:'New'}:button},console,
 setSaveStatus(){},validateRequiredMetadata(){},materializeNewTags:async()=>{},resolvePeople:async()=>{},resolveDesiredContentType:async()=>null,requestNumbering:async()=>null,buildSharePointFormValues:()=>[],buildDesiredFilename:()=> 'New.docx',validateUpdateListItem:async fields=>events.push(fields[0]?.FieldName==='FileLeafRef'?'rename':'metadata'),verifySavedMetadata:async()=>events.push('verify'),refreshDocumentMetadata:async()=>{events.push('fields');return true;},rememberDocumentLocation(){},Word:{run:async fn=>fn({document:{save:()=>events.push('word-save')},sync:async()=>{}})},getErrorMessage:e=>e.message};
 vm.runInNewContext(block+'\nglobalThis.save=saveMetadata;',ctx);await ctx.save();
 assert.ok(events.indexOf('fields')<events.indexOf('word-save'));assert.ok(events.indexOf('word-save')<events.indexOf('rename'));assert.deepEqual(events.slice(events.indexOf('rename')+1),['verify']);
});
