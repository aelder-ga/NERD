'use strict';
const fs=require('node:fs');const path=require('node:path');
const GUID=/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function validate(c){
 const keys=['organization','tenantId','clientId','sharepointHost','azureOrigin','documentsListId'];
 if(Object.keys(c).some(k=>!keys.includes(k)))throw Error('Unknown configuration key; never put secrets in configuration.');
 for(const k of ['tenantId','clientId','documentsListId'])if(!GUID.test(c[k]||''))throw Error('Invalid '+k);
 if(!/^[a-z0-9][a-z0-9-]*\.sharepoint\.com$/.test(c.sharepointHost||''))throw Error('Use a commercial Microsoft SharePoint hostname.');
 if(!/^https:\/\/[a-z0-9][a-z0-9.-]*\.azurestaticapps\.net$/.test(c.azureOrigin||''))throw Error('Use an Azure Static Web Apps HTTPS origin without a trailing slash.');
 if(typeof c.organization!=='string'||!c.organization.trim()||c.organization.length>100||/[<>"'&\r\n\\]/.test(c.organization))throw Error('Use a plain organization name without markup or quotes.');
 return c;
}
const exclusions=new Set(['node_modules','.git','dist','lib','temp','release','azure','solution','deployment']);
const roots=['src','api','assets','scripts','tests','sharepoint'];
const rootFiles=['package.json','package-lock.json','babel.config.json','webpack.azure.js','manifest.xml','staticwebapp.config.json'];
function configure(source,destination,c){
 validate(c);source=path.resolve(source);destination=path.resolve(destination);
 if(destination===source||destination.startsWith(source+path.sep))throw Error('Output must be outside the source tree.');
 if(fs.existsSync(destination))throw Error('Output already exists; choose a new directory.');
 const mapping=[['https://rocktwpnet', 'https://'+c.sharepointHost+'/'],['rocktwpnet.sharepoint.com',c.sharepointHost],['1169a3a9-3860-4a0b-80aa-410be007cde5',c.tenantId],['f8371eb2-8758-48cf-8fa5-2f9b696d89c4',c.clientId],['dd8ac8ef-3be9-4fb8-b7cb-6ac0f56d2b03',c.documentsListId],['https://thankful-sky-0cc1b0210.6.azurestaticapps.net',c.azureOrigin],['Rockaway Township School District',c.organization]];
 const files=[];
 function copy(relative){const full=path.join(source,relative);const s=fs.lstatSync(full);if(s.isSymbolicLink())throw Error('Symlink rejected: '+relative);
  if(s.isDirectory()){for(const name of fs.readdirSync(full)){if(!exclusions.has(name))copy(path.join(relative,name));}return;}
  if(/\.sppkg(?:\.base64)?$/.test(relative)||/(?:serve|deploy-azure)\.json$/.test(relative))return;
  let bytes=fs.readFileSync(full);
  if(/\.(?:js|cjs|ts|json|html|xml|css|scss|ya?ml|md)$/.test(relative)){
   let t=bytes.toString('utf8');
   // Resource detection originally used a partial tenant hostname. Fix it before full host replacement.
   t=t.replaceAll('startsWith("https://rocktwpnet")','startsWith("https://'+c.sharepointHost+'/")');
   for(const [from,to] of mapping.slice(1))t=t.split(from).join(to);
   if(relative==='src/taskpane/taskpane.js')t=t.replaceAll('http://localhost:3001','');
   bytes=Buffer.from(t);
  }
  const output=path.join(destination,relative);fs.mkdirSync(path.dirname(output),{recursive:true});fs.writeFileSync(output,bytes);files.push(relative);
 }
 try{for(const name of [...rootFiles,...roots])copy(name);
 fs.writeFileSync(path.join(destination,'district.json'),JSON.stringify(c,null,2)+'\n');
 fs.writeFileSync(path.join(destination,'build-district.cjs'),`const {execFileSync}=require('node:child_process');\nconst npm=process.platform==='win32'?'npm.cmd':'npm';\nfor(const name of ['geek-launcher','dork-dashboard']){execFileSync(npm,['ci','--ignore-scripts','--no-audit'],{cwd:'sharepoint/'+name,stdio:'inherit'});execFileSync(npm,['run','build'],{cwd:'sharepoint/'+name,stdio:'inherit'});execFileSync(process.execPath,['scripts/stage-'+name+'.cjs'],{stdio:'inherit'});}\nexecFileSync(npm,['run','build:azure'],{stdio:'inherit'});\nexecFileSync(process.execPath,['scripts/create-manifest.cjs',require('./district.json').azureOrigin],{stdio:'inherit'});\n`);
 return files;
 }catch(e){fs.rmSync(destination,{recursive:true,force:true});throw e;}
}
module.exports={validate,configure};
if(require.main===module){const [config,destination]=process.argv.slice(2);if(!config||!destination)throw Error('Usage: node deployment/configure.cjs district.json NEW_OUTPUT_FOLDER');const files=configure(path.resolve(__dirname,'..'),destination,JSON.parse(fs.readFileSync(config,'utf8')));console.log('Prepared '+files.length+' source files. No deployment performed.');}
