const fs = require('fs');
const path = require('path');
const origin = new URL(process.argv[2]);
if (origin.protocol !== 'https:' || origin.username || origin.password || origin.pathname !== '/' || origin.search || origin.hash) {
  throw new Error('Supply the HTTPS origin only, such as https://your-app.azurestaticapps.net');
}
const root = path.resolve(__dirname, '..');
let xml = fs.readFileSync(path.join(root, 'manifest.xml'), 'utf8');
xml = xml.replaceAll('https://localhost:3000', origin.origin)
  .replace('https://www.contoso.com/help', 'https://rocktwpnet.sharepoint.com/sites/DORK')
  .replace('https://www.contoso.com', origin.origin)
  .replace('<ProviderName>Contoso</ProviderName>', '<ProviderName>Rockaway Township School District</ProviderName>')
  .replace('<Version>1.0.0.0</Version>', '<Version>1.0.2.0</Version>')
  .replace('A template to get started.', 'Edit and verify DORK document metadata.')
  .replace('Get started with your sample add-in!', 'Get started with NERD')
  .replace('Commands Group', 'DORK')
  .replaceAll('Show Task Pane', 'Open NERD')
  .replace('Your sample add-in loaded successfully.', 'NERD loaded successfully.')
  .replace('Click to Show a Taskpane', 'Open the NERD metadata pane');
fs.writeFileSync(path.join(root, 'manifest-azure.xml'), xml);
if (fs.existsSync(path.join(root, 'dist'))) fs.writeFileSync(path.join(root, 'dist', 'manifest-azure.xml'), xml);
console.log('Created manifest-azure.xml');
console.log('Add SPA redirect URI: brk-multihub://' + origin.host);
console.log('Add SPA redirect URI: ' + origin.origin + '/taskpane.html');
console.log('Add SPA redirect URI: ' + origin.origin + '/auth.html');
