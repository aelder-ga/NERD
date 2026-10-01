const {test}=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
const source=fs.readFileSync(require('node:path').join(__dirname,'../src/taskpane/taskpane.js'),'utf8');
const block=source.substring(source.indexOf('async function saveMetadata()'),source.indexOf('function validateRequiredMetadata()'));
test('Word fields and content save finish before rename; no Word writes follow rename',async()=>{
 const events=[],button={disabled:false};
 const ctx={isSaving:false,currentDriveItem:{id:'file',name:'Old.docx'},currentListItem:{id:'15'},currentDocumentUrl:'old',dorkSite:{id:'site'},dorkDocumentsLibrary:{id:'list'},window:{sessionStorage:{}},Office:{context:{host:'Word'},HostType:{Excel:'Excel'}},document:{getElementById:id=>id==='title'?{value:'New'}:button},console,
 syncWorkbookServerProperties:async()=>events.push('server-properties'),dorkColumns:{},personSelections:{},requireColumn:name=>({name}),setSaveStatus(){},validateRequiredMetadata(){},materializeNewTags:async()=>{},resolvePeople:async()=>{},resolveDesiredContentType:async()=>null,requestNumbering:async()=>null,buildSharePointFormValues:()=>[],buildDesiredFilename:()=> 'New.docx',validateUpdateListItem:async fields=>events.push(fields[0]?.FieldName==='FileLeafRef'?'rename':'metadata'),verifySavedMetadata:async()=>events.push('verify'),refreshDocumentMetadata:async()=>{events.push('fields');return true;},rememberDocumentLocation(){},saveHostDocument:async()=>events.push('word-save'),getErrorMessage:e=>e.message};
 vm.runInNewContext(block+'\nglobalThis.save=saveMetadata;',ctx);await ctx.save();
 assert.ok(events.indexOf('fields')<events.indexOf('word-save'));assert.ok(events.indexOf('word-save')<events.indexOf('rename'));assert.deepEqual(events.slice(events.indexOf('rename')+1),['verify']);
});

test('Excel saves workbook before SharePoint metadata and verification; no workbook edits follow verification',async()=>{
 const events=[],button={disabled:false};
 const ctx={isSaving:false,currentDriveItem:{id:'file',name:'Old.xlsx'},currentListItem:{id:'15'},currentDocumentUrl:'old',dorkSite:{id:'site'},dorkDocumentsLibrary:{id:'list'},window:{sessionStorage:{}},Office:{context:{host:'Excel'},HostType:{Excel:'Excel'}},document:{getElementById:id=>id==='title'?{value:'New'}:button},console,
 syncWorkbookServerProperties:async()=>events.push('server-properties'),dorkColumns:{},personSelections:{},requireColumn:name=>({name}),setSaveStatus(){},validateRequiredMetadata(){},materializeNewTags:async()=>{},resolvePeople:async()=>{},resolveDesiredContentType:async()=>null,requestNumbering:async()=>null,buildSharePointFormValues:()=>[],buildDesiredFilename:()=> 'New.xlsx',validateUpdateListItem:async fields=>events.push(fields[0]?.FieldName==='FileLeafRef'?'rename':'metadata'),verifySavedMetadata:async()=>events.push('verify'),refreshDocumentMetadata:async()=>{events.push('fields');return true;},rememberDocumentLocation(){},saveHostDocument:async()=>events.push('workbook-save'),getErrorMessage:e=>e.message};
 vm.runInNewContext(block+'\nglobalThis.save=saveMetadata;',ctx);await ctx.save();
 assert.deepEqual(events,['fields','server-properties','workbook-save','metadata','verify','rename','verify']);
});
