const fs=require('node:fs');const path=require('node:path');
const root=path.join(__dirname,'..','sharepoint','dork-dashboard');const target=path.join(root,'azure');fs.mkdirSync(target,{recursive:true});
for(const folder of ['release/assets','release/manifests'])for(const name of fs.readdirSync(path.join(root,folder)))fs.copyFileSync(path.join(root,folder,name),path.join(target,name));
const manifest=JSON.parse(fs.readFileSync(path.join(target,'d451ec37-7370-42a6-906b-b985e1db43e6.manifest.json'),'utf8'));
if(manifest.loaderConfig.internalModuleBaseUrls.some(url=>!url.startsWith('https://thankful-sky-0cc1b0210.6.azurestaticapps.net/sharepoint/dork-dashboard/')))throw Error('Dashboard assets must stay in the existing Azure Static Web App.');
fs.writeFileSync(path.join(root,'dork-dashboard.sppkg.base64'),fs.readFileSync(path.join(root,'sharepoint/solution/dork-dashboard.sppkg')).toString('base64')+'\n');
console.log('Staged dashboard assets and package for Azure.');
