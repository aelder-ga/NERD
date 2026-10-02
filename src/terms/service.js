const SET_NAMES = { collection: 'Collection', system: 'System / Platform' };
const normalized = value => String(value || '').normalize('NFKC').trim().replace(/\s+/g, ' ').toLocaleLowerCase('en-US');
const label = term => term.labels?.find(item => item.isDefault)?.name || term.labels?.[0]?.name || '';
const description = term => term.descriptions?.find(item => item.languageTag?.toLowerCase() === 'en-us')?.description || term.descriptions?.[0]?.description || '';
class TermManager {
 constructor(request, siteId) { this.request = request; this.base = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/termStore`; this.sets = {}; }
 async all(url) {
  const items = [], seen = new Set();
  while (url) {
   const location = new URL(url);
   if (location.origin !== 'https://graph.microsoft.com' || !decodeURIComponent(location.pathname).startsWith(decodeURIComponent(new URL(this.base).pathname) + '/') || seen.has(url)) throw Error('Unexpected term-store results page.');
   seen.add(url); const result = await this.request(url);
   if (!Array.isArray(result.value)) throw Error('Unexpected term-store response.');
   items.push(...result.value); url = result['@odata.nextLink'];
  }
  return items;
 }
 async initialize() {
  const groups = await this.all(this.base + '/groups');
  const matches = groups.filter(group => normalized(group.displayName) === 'dork');
  if (matches.length !== 1) throw Error('A unique DORK term group was not found.');
  const sets = await this.all(`${this.base}/groups/${encodeURIComponent(matches[0].id)}/sets`);
  for (const [kind, name] of Object.entries(SET_NAMES)) {
   const found = sets.filter(set => (set.localizedNames || []).some(item => normalized(item.name) === normalized(name)) || normalized(set.displayName) === normalized(name));
   if (found.length !== 1) throw Error(`A unique ${name} term set was not found.`);
   this.sets[kind] = found[0].id;
  }
 }
 endpoint(kind) {
  if (!Object.hasOwn(SET_NAMES, kind) || !this.sets[kind]) throw Error('Choose Collection or System / Platform.');
  return `${this.base}/sets/${encodeURIComponent(this.sets[kind])}`;
 }
 async list(kind) {
  const base = this.endpoint(kind), terms = [], queue = [base + '/children'], seen = new Set();
  while (queue.length) {
   const batch = await Promise.all(queue.splice(0, 4).map(url => this.all(url)));
   for (const term of batch.flat()) {
    if (seen.has(term.id)) continue;
    seen.add(term.id); terms.push(term); queue.push(`${base}/terms/${encodeURIComponent(term.id)}/children`);
   }
  }
  return terms.sort((a,b) => label(a).localeCompare(label(b), undefined, {sensitivity:'base'}));
 }
 async create(kind, name, notes) {
  const endpoint = this.endpoint(kind); const cleaned = String(name || '').normalize('NFKC').trim().replace(/\s+/g, ' ');
  if (!cleaned || cleaned.length > 255 || /[;<>|\x00-\x1f]/.test(cleaned)) throw Error('Enter a name of 1–255 characters without semicolons, angle brackets, pipes or control characters.');
  const cleanedNotes = String(notes || '').trim();
  if (cleanedNotes.length > 1000) throw Error('Keep the description to 1,000 characters or fewer.');
  const terms = await this.list(kind);
  const duplicate = terms.find(term => (term.labels || []).some(item => normalized(item.name) === normalized(cleaned)));
  if (duplicate) throw Error(`Already exists: ${label(duplicate)}. Use the existing entry.`);
  let created;
  try {
   created = await this.request(endpoint + '/children', {method:'POST', body:JSON.stringify({labels:[{name:cleaned, languageTag:'en-US', isDefault:true}], ...(cleanedNotes ? {descriptions:[{description:cleanedNotes, languageTag:'en-US'}]} : {})})});
  } catch (error) {
   if (error.status === 403) throw Error('Your account can view these entries but does not have permission to add them. Ask a DORK term-store maintainer.');
   if (error.status === 409) throw Error('This entry was added by someone else. Refresh the list before trying again.');
   if (error.status >= 500 || !error.status) throw Error('The result could not be confirmed. Refresh and search for the name before trying again.');
   throw error;
  }
  if (!created?.id) throw Error('The result could not be confirmed. Refresh and search for the name before trying again.');
  return created;
 }
}
module.exports = {TermManager, SET_NAMES, normalized, label, description};
