const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../src/taskpane/taskpane.js'), 'utf8');
function request(optional = {}) {
    const input = {title:'Test', classification:'Internal', lifecycle:'Draft', 'information-source':'', 'last-reviewed':'', 'next-review':'', ...optional};
    const context = {
        document:{getElementById:id=>({value:input[id]})}, requireColumn:name=>({name}),
        getSelectedTerm:id=>({label:id,id:`guid-${id}`}),
        pickerSelections:{'system-platform':[],collection:[],tags:[],audience:[]},
        personSelections:{owner:{userPrincipalName:'owner@example.org'},responsible:null,secondary:null},
    };
    vm.runInNewContext(source.slice(source.indexOf('function buildSharePointFormValues('),source.indexOf('function buildDesiredFilename(')) + '\nglobalThis.build=buildSharePointFormValues;',context);
    return Object.fromEntries(context.build({id:'0x0101'}).map(item=>[item.FieldName,item.FieldValue]));
}
test('Clearing optional metadata explicitly sends empty fields rather than omitting them',()=>{
    const fields=request();
    for(const name of ['System / Platform','Collection','Tags','Audience','Responsible','Secondary','Information Source','Last Reviewed','Next Review']) {
        assert.ok(Object.hasOwn(fields,name),name);assert.equal(fields[name],'',name);
    }
    assert.equal(fields.Owner,'[{"Key":"i:0#.f|membership|owner@example.org"}]');
    assert.equal(fields.ContentTypeId,'0x0101');
    assert.equal(fields.Domain,'domain|guid-domain');
});
test('Review dates preserve calendar days including leap days and reject impossible dates',()=>{
    assert.equal(request({'last-reviewed':'2028-02-29','next-review':'2029-01-01'})['Last Reviewed'],'2/29/2028');
    assert.equal(request({'next-review':'2029-01-01'})['Next Review'],'1/1/2029');
    for(const date of ['2027-02-29','2026-04-31','2026-13-01','9/30/2026']) {
        assert.throws(()=>request({'last-reviewed':date}),/Review date/);
    }
});
