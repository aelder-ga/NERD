import { PublicClientApplication } from '@azure/msal-browser';
import { TermManager, SET_NAMES, normalized, label, description } from './service';
const scopes = ['User.Read','User.ReadBasic.All','Sites.ReadWrite.All','TermStore.ReadWrite.All'];
const msal = new PublicClientApplication({auth:{clientId:'f8371eb2-8758-48cf-8fa5-2f9b696d89c4',authority:'https://login.microsoftonline.com/1169a3a9-3860-4a0b-80aa-410be007cde5',redirectUri:window.location.origin+'/sso-redirect.html'},cache:{cacheLocation:'localStorage'}});
const $ = id => document.getElementById(id);
let manager, entries = [], busy = false, connected = false;
const hint = new URLSearchParams(window.location.hash.slice(1)).get('login_hint');
if (hint) history.replaceState(null,'',window.location.pathname);
function status(message, error=false) { $('status').textContent=message; $('status').className=error?'error':''; }
function controls() { $('entry-fields').disabled = !connected || busy; $('refresh').disabled = !connected || busy; $('signin').disabled=busy; }
async function token(interactive=false) {
 const accounts = msal.getAllAccounts();
 const account = hint ? accounts.find(a=>a.username?.toLowerCase()===hint.toLowerCase()) : msal.getActiveAccount() || (accounts.length===1 ? accounts[0] : null);
 const request={scopes,...(account?{account}:{}),...(hint?{loginHint:hint}:{})};
 let result;
 try {result=await msal.acquireTokenSilent(request);} catch (_) {
  try {result=await msal.ssoSilent(request);} catch (_) {
   if (!interactive) { $('signin').hidden=false; throw Error('Select Sign in to DORK to continue.'); }
   result=await msal.acquireTokenPopup(request);
  }
 }
 msal.setActiveAccount(result.account); return result.accessToken;
}
async function request(url, options={}) {
 const accessToken=await token();
 const response=await fetch(url,{...options,headers:{Authorization:`Bearer ${accessToken}`,'Content-Type':'application/json'}});
 const body=await response.json().catch(()=>({}));
 if (!response.ok) {const error=Error(body.error?.message || `Microsoft returned ${response.status}.`); error.status=response.status; throw error;}
 return body;
}
function render() {
 const query=normalized($('search').value); const visible=entries.filter(t=>[...(t.labels||[]).map(l=>l.name),description(t)].some(v=>normalized(v).includes(query)));
 $('count').textContent=`${visible.length} of ${entries.length} entries`;
 $('entries').replaceChildren();
 for (const term of visible) {
  const li=document.createElement('li'), title=document.createElement('strong'), notes=document.createElement('p');
  title.textContent=label(term); notes.textContent=description(term); li.append(title); if(notes.textContent)li.append(notes); $('entries').append(li);
 }
 if (!visible.length) {const li=document.createElement('li');li.textContent='No matching entries.';$('entries').append(li);}
}
async function refresh() {entries=await manager.list($('kind').value); render();}
async function run(action) {if(busy)return;busy=true;controls();try{await action();}catch(e){status(e.message,true);}finally{busy=false;controls();}}
async function connect(interactive) {
 await token(interactive);
 const me=await request('https://graph.microsoft.com/v1.0/me?$select=displayName');
 const site=await request('https://graph.microsoft.com/v1.0/sites/rocktwpnet.sharepoint.com:/sites/DORK');
 manager=new TermManager(request,site.id); await manager.initialize(); await refresh(); connected=true;
 $('signin').hidden=true; $('identity').textContent=`Connected as ${me.displayName}`; status('Ready. Search existing entries before adding one.');
}
$('signin').onclick=()=>run(()=>connect(true));
$('search').oninput=render;
$('kind').onchange=()=>run(async()=>{entries=[];render();status('Loading entries…');await refresh();status('Ready.');});
$('refresh').onclick=()=>run(async()=>{await refresh();status('List refreshed.');});
$('entry-form').onsubmit=event=>{event.preventDefault();if(!connected)return;void run(async()=>{
 const kind=$('kind').value;status('Checking existing entries…');
 const created=await manager.create(kind,$('name').value,$('description').value);
 entries.push(created); render(); $('name').value='';$('description').value='';
 status(`${label(created)} added to ${SET_NAMES[kind]}. Reopen or reconnect NERD/GEEK to load the new entry.`);
 });};
void run(async()=>{await msal.initialize();await connect(false);});
