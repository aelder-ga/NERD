const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
function moduleFrom(path, imports={}) {
 const code=fs.readFileSync(require('node:path').join(__dirname,'../src/taskpane/',path),'utf8').replace(/^import .*;$/gm,'').replace(/export /g,'');
 const ctx={...imports}; vm.runInNewContext(code+'\nglobalThis.api={'+(path==='office-host.js'?'assertSupportedHost,saveHostDocument':'excelText,syncWorkbookMetadata,syncReferenceCover')+'};',ctx); return ctx.api;
}
const api=moduleFrom('workbook-metadata.js');
function fixture(existingMarker, exists=true) {
 const ranges={},events=[],sheet={isNullObject:!exists,load(){},getRange(address){return ranges[address] ||= {values:address==='A1'?[[existingMarker]]:[],load(){},format:{font:{},autofitRows(){}}};}};
 const context={workbook:{worksheets:{getItemOrNullObject(name){return name==='Document Control'?{isNullObject:true,load(){}}:sheet;},add(name){events.push(name);return sheet;}}},sync:async()=>{}};
 return {ranges,events,excel:{run:fn=>fn(context)}};
}
test('Excel sync creates only a dedicated metadata sheet, safely writes literal values and does not activate it',async()=>{
 const f=fixture(null,false); await api.syncWorkbookMetadata({DORK_Title:'=HYPERLINK("bad")',DORK_Tags:'one; two',DORK_Owner:''},f.excel);
 assert.deepEqual(f.events,['DORK Metadata']);assert.equal(f.ranges.A1.values[0][0],'NERD workbook metadata v1');assert.equal(f.ranges['A3:B5'].values[0][1],'\'=HYPERLINK("bad")');assert.equal(f.ranges['A3:B5'].values[2][1],'');
});
test('Existing business sheet with matching name is never overwritten',async()=>{
 const f=fixture('Business information'); await assert.rejects(api.syncWorkbookMetadata({DORK_Title:'New'},f.excel),/already exists/);assert.equal(f.ranges.A1.values[0][0],'Business information');assert.equal(f.ranges['A3:B3'],undefined);
});
test('Marked sheet can be refreshed without duplicate sheets; dangerous Excel text remains literal',async()=>{
 const f=fixture('NERD workbook metadata v1');await api.syncWorkbookMetadata({DORK_Title:'New'},f.excel);assert.equal(f.events.length,0);assert.equal(f.ranges['A3:B3'].values[0][1],'New');
 for(const text of ['=1+1','+1','-1','@SUM(A1)',"'quoted"]) assert.equal(api.excelText(text),"'"+text);
});
test('Host save dispatches to Excel and Word; unsupported Excel fails before writing',async()=>{
 const host=moduleFrom('office-host.js'); const events=[]; const office={HostType:{Excel:'Excel',Word:'Word'},context:{host:'Excel',requirements:{isSetSupported:()=>true}}};
 await host.saveHostDocument(office,null,{run:fn=>fn({workbook:{save:value=>events.push(value)},sync:async()=>events.push('sync')})});assert.deepEqual(events,['Save','sync']);
 office.context.host='Word';await host.saveHostDocument(office,{run:fn=>fn({document:{save:()=>events.push('Word')},sync:async()=>{}})});assert.equal(events[2],'Word');
 office.context.host='Excel';office.context.requirements.isSetSupported=()=>false;await assert.rejects(host.saveHostDocument(office),/ExcelApi 1.11/);
});

test('Reference cover replaces repeated-save title and every visible dimension while preserving formatting and data sheets', async()=>{
 const labels=Array.from({length:14},()=>['']);
 for(const [row,text] of Object.entries({0:'ROCKAWAY TOWNSHIP SCHOOL DISTRICT | TECHNOLOGY',2:'DORK REFERENCE SHEET',8:'DOCUMENT CONTROL',9:'Domain',10:'Function',11:'System / Platform',12:'Collection',13:'Classification'})) labels[row][0]=text;
 const ranges={},cover={isNullObject:false,load(){},getRange(address){return ranges[address] ||= {values:address==='A1:A14'?labels:[],load(){}};}};
 const context={workbook:{worksheets:{getItemOrNullObject(name){assert.equal(name,'Document Control');return cover;}}},sync:async()=>{}};
 const values={DORK_DocumentId:'DORK-0003 ',DORK_Title:'Microsoft Administration Portals',DORK_Domain:'IDENTITY',DORK_Function:'Administrative Authorization',DORK_System:'Entra ID; Intune; Microsoft 365; Windows',DORK_Collection:'',DORK_Classification:'Internal'};
 await api.syncReferenceCover(context,values);
 assert.equal(ranges.A17.values[0][0],'DORK-0003 Microsoft Administration Portals');
 for(const [address,key] of Object.entries({D10:'DORK_Domain',D11:'DORK_Function',D12:'DORK_System',D13:'DORK_Collection',D14:'DORK_Classification'})) assert.equal(ranges[address].values[0][0],values[key]);
 values.DORK_Title='DORK-0003 Microsoft Administration Portals';await api.syncReferenceCover(context,values);assert.equal(ranges.A17.values[0][0],'DORK-0003 Microsoft Administration Portals');
 values.DORK_Title='=literal title';await api.syncReferenceCover(context,values);assert.equal(ranges.A17.values[0][0],'DORK-0003 =literal title');
 assert.equal(ranges.A17.format,undefined);
 labels[2][0]='Business document';await assert.rejects(api.syncReferenceCover(context,values),/does not match/);
});
