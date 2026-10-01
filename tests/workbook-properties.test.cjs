const {test}=require('node:test');const assert=require('node:assert/strict');const vm=require('node:vm');const fs=require('node:fs');
const {DOMParser,XMLSerializer}=require('@xmldom/xmldom');
const source=fs.readFileSync(require('node:path').join(__dirname,'../src/taskpane/workbook-properties.js'),'utf8').replace(/export /g,'');
const ctx={DOMParser,XMLSerializer};vm.runInNewContext(source+'\nglobalThis.api={rewriteServerProperties,syncWorkbookServerProperties};',ctx);
const schema=`<ct:contentTypeSchema xmlns:ct="http://schemas.microsoft.com/office/2006/metadata/contentType" xmlns:ma="http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes" xmlns:xsd="http://www.w3.org/2001/XMLSchema" ma:contentTypeName="DORK Procedure"><xsd:schema targetNamespace="urn:site"><xsd:element name="taxonomy-key" ma:taxonomy="true" ma:taxonomyFieldName="Domain"/><xsd:element name="Platforms" ma:taxonomy="true" ma:taxonomyFieldName="System"/><xsd:element name="Owner" ma:internalName="Owner" ma:list="UserInfo"/><xsd:element name="Audience" ma:internalName="Audience"/><xsd:element name="Reviewed" ma:internalName="Reviewed" ma:format="DateOnly" nillable="true"/><xsd:element name="Classification" ma:internalName="Classification"/><xsd:element name="DORK_ID" ma:internalName="DORK_ID" nillable="true"/></xsd:schema></ct:contentTypeSchema>`;
const properties=`<p:properties xmlns:p="http://schemas.microsoft.com/office/2006/metadata/properties" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><documentManagement><taxonomy-key xmlns="urn:site"/><Classification xmlns="urn:site"/><Unrelated xmlns="urn:other">Keep me</Unrelated></documentManagement></p:properties>`;
test('Embedded metadata synchronizes taxonomy GUIDs, multiple terms, owner, dates, ID, and content type without changing unrelated XML',()=>{
 const fields=Object.entries({Domain:'IDENTITY|00000000-0000-0000-0000-000000000001',System:'Entra & ID|00000000-0000-0000-0000-000000000002;Intune|00000000-0000-0000-0000-000000000003',Owner:JSON.stringify([{Key:'i:0#.f|membership|user@example.test'}]),Audience:'Technology;#Administration',Reviewed:'10/1/2026',Classification:'Internal',DORK_ID:'DORK-0003'}).map(([FieldName,FieldValue])=>({FieldName,FieldValue}));
 const out=ctx.api.rewriteServerProperties(properties,schema,fields,{Owner:'Example Owner'},{id:'type-reference',name:'DORK Reference Sheet'});
 const xml=new DOMParser().parseFromString(out.properties,'application/xml');const texts=tag=>Array.from(xml.getElementsByTagName(tag)).map(n=>n.textContent);
 assert.deepEqual(texts('pc:TermName'),['IDENTITY','Entra & ID','Intune']);assert.equal(xml.getElementsByTagName('AccountId')[0].textContent,'i:0#.f|membership|user@example.test');assert.equal(xml.getElementsByTagName('Reviewed')[0].textContent,'2026-10-01T00:00:00');assert.equal(xml.getElementsByTagName('DORK_ID')[0].textContent,'DORK-0003');assert.equal(xml.getElementsByTagName('Unrelated')[0].textContent,'Keep me');assert.match(out.schema,/DORK Reference Sheet/);
 const again=ctx.api.rewriteServerProperties(out.properties,out.schema,fields,{Owner:'Example Owner'},{id:'type-reference',name:'DORK Reference Sheet'});assert.equal(again.properties,out.properties);
 const cleared=ctx.api.rewriteServerProperties(out.properties,out.schema,fields.map(f=>({...f,FieldValue:''})));assert.doesNotMatch(cleared.properties,/Example Owner|Internal|DORK-0003|2026-10-01/);
});
test('Invalid taxonomy is rejected before server property XML can be applied',()=>assert.throws(()=>ctx.api.rewriteServerProperties(properties,schema,[{FieldName:'Domain',FieldValue:'IDENTITY|invalid'}]),/invalid term/));
test('Excel removes only conflicting metadata aliases and updates the real workbook title and content type',async()=>{
 const events=[],custom={items:[{key:'Domain',delete:()=>events.push('delete Domain')},{key:'Document Type',delete:()=>events.push('delete alias')},{key:'Unrelated',delete:()=>events.push('WRONG')}],load(){},add:(key,value)=>events.push([key,value])};
 const empty={items:[],load(){}};const context={workbook:{customXmlParts:empty,properties:{custom}},sync:async()=>{}};
 await ctx.api.syncWorkbookServerProperties([{FieldName:'Domain',FieldValue:'x'},{FieldName:'Document_x0020_Type',FieldValue:'y'},{FieldName:'Title',FieldValue:'A title'}],{type:{name:'Document_x0020_Type',displayName:'Document Type'}},{},{id:'reference',name:'DORK Reference Sheet'},{run:fn=>fn(context)});
 assert.deepEqual(events,['delete Domain','delete alias',['ContentTypeId','reference']]);assert.equal(context.workbook.properties.title,'A title');
});
test('Excel discovers prefixed server XML parts and synchronizes them without altering other parts',async()=>{
 const written=[],part=(xml,label)=>({getXml:()=>({value:xml}),setXml:value=>written.push([label,value])});
 const parts={items:[part(properties,'properties'),part(schema,'schema'),part('<Other xmlns="urn:other">unchanged</Other>','other')],load(){}};
 const context={workbook:{customXmlParts:parts,properties:{custom:{items:[],load(){},add(){}}}},sync:async()=>{}};
 await ctx.api.syncWorkbookServerProperties([{FieldName:'Classification',FieldValue:'Internal'}],{},{},{id:'reference',name:'DORK Reference Sheet'},{run:fn=>fn(context)});
 assert.deepEqual(written.map(v=>v[0]),['properties','schema']);assert.match(written[0][1],/Internal/);assert.match(written[1][1],/DORK Reference Sheet/);
 parts.items.push(part(properties,'duplicate'));await assert.rejects(ctx.api.syncWorkbookServerProperties([],{}, {},null,{run:fn=>fn(context)}),/ambiguous/);
});
