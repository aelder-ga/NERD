const fs = require('node:fs');
const path = require('node:path');
const root = path.join(__dirname, '..', 'sharepoint', 'geek-launcher');
const target = path.join(root, 'azure');
fs.mkdirSync(target, { recursive: true });
for (const folder of ['release/assets', 'release/manifests']) {
 for (const name of fs.readdirSync(path.join(root, folder))) fs.copyFileSync(path.join(root, folder, name), path.join(target, name));
}
const manifest = JSON.parse(fs.readFileSync(path.join(target, '99c231a1-a31c-48ed-b6a5-919d89235101.manifest.json'), 'utf8'));
if (manifest.loaderConfig.internalModuleBaseUrls.some(url => !url.startsWith('https://thankful-sky-0cc1b0210.6.azurestaticapps.net/'))) throw Error('Launcher assets must be hosted in Azure.');
fs.writeFileSync(path.join(root, 'geek-launcher.sppkg.base64'), fs.readFileSync(path.join(root, 'sharepoint/solution/geek-launcher.sppkg')).toString('base64') + '\n');
console.log('Staged Azure launcher assets and install package.');
