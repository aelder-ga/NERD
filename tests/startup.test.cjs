const {test} = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname,'../src/taskpane/taskpane.js'),'utf8');
const block = source.split('/* STARTUP */')[1].split('Office.onReady')[0];
function setup(fail=false) {
    const elements = Object.fromEntries(['connect-nerd','metadata-form','save-metadata'].map(id=>[id,{hidden:false,disabled:false}]));
    const events=[];
    const context={document:{getElementById:id=>elements[id]},console:{error(){}},msalConfig:{},GRAPH_ROOT:'graph',DORK_HOST:'host',DORK_SITE_PATH:'site',SHAREPOINT_SCOPES:['sharepoint'],
        createNestablePublicClientApplication:async()=>({}),graphGet:async()=>({displayName:'Test'}),
        acquireToken:async()=>events.push('sharepoint'),
        initializeTaxonomy:async()=>events.push('taxonomy'),initializeLibrarySchema:async()=>events.push('schema'),
        initializeCurrentDocument:async()=>{assert.equal(elements['metadata-form'].hidden,true);if(fail)throw Error('Discovery failed');events.push('document');elements['metadata-form'].hidden=false;elements['save-metadata'].disabled=false;},
        setConnectionStatus:value=>events.push(value),setSaveStatus(){},getErrorMessage:error=>error.message};
    vm.runInNewContext(block+'\nglobalThis.connect=connectToDork;',context);
    return {context,elements,events};
}
test('Startup checks SharePoint before loading fields and hides redundant connection button',async()=>{
    const {context,elements,events}=setup();await context.connect();
    assert.ok(events.indexOf('sharepoint')<events.indexOf('document'));
    assert.equal(elements['connect-nerd'].hidden,true);
    assert.equal(elements['metadata-form'].hidden,false);
    assert.equal(elements['save-metadata'].disabled,false);
});
test('Failed startup keeps Save unavailable and provides a working retry button',async()=>{
    const {context,elements,events}=setup(true);await context.connect();
    assert.equal(elements['metadata-form'].hidden,true);
    assert.equal(elements['save-metadata'].disabled,true);
    assert.equal(elements['connect-nerd'].hidden,false);
    assert.equal(elements['connect-nerd'].disabled,false);
    assert.equal(elements['connect-nerd'].textContent,'Retry connection');
    assert.ok(events.includes('Connection failed: Discovery failed'));
});
