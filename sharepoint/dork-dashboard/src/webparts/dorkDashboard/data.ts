import { DocumentEntry, NeedEntry, safeUrl, fileUrl } from './model';
interface Row { [key: string]: any; }
export interface DashboardData {
  documents: DocumentEntry[]; needs: NeedEntry[]; documentError: string; needsError: string;
  documentsUrl: string; needsUrl: string; newNeedUrl: string; loadedAt: Date;
}
type ReadJson = (url: string) => Promise<Row>;
const DOCUMENTS_ID = 'dd8ac8ef-3be9-4fb8-b7cb-6ac0f56d2b03';
const text = (v: unknown): string => typeof v === 'string' ? v : '';
function term(v: unknown): string {
  if (typeof v === 'string') return v.replace(/^\d+;#/, '').split('|')[0];
  return v && typeof v === 'object' ? text((v as Row).Label) : '';
}
export async function allRows(url: string, read: ReadJson, site: string): Promise<Row[]> {
  const rows: Row[] = [], visited = new Set<string>();
  while (url) {
    const checked = safeUrl(url, site);
    if (!checked || !new URL(checked).pathname.startsWith(`${new URL(site).pathname}/_api/`)) throw new Error('Unexpected SharePoint paging location.');
    if (visited.has(checked)) throw new Error('SharePoint returned a repeated results page.');
    visited.add(checked);
    const page = await read(checked);
    const entries = page.value || page.d?.results;
    if (!Array.isArray(entries)) throw new Error('SharePoint returned an unexpected results format.');
    rows.push(...entries);
    url = page['@odata.nextLink'] || page['odata.nextLink'] || page.d?.__next || '';
  }
  return rows;
}
export async function loadData(site: string, read: ReadJson): Promise<DashboardData> {
  const result: DashboardData = { documents: [], needs: [], documentError: '', needsError: '', documentsUrl: `${site}/Shared%20Documents/Forms/AllItems.aspx`, needsUrl: '', newNeedUrl: '', loadedAt: new Date() };
  async function fields(base: string): Promise<{ [key: string]: string }> {
    const values = await allRows(`${base}/fields?$select=Title,InternalName&$top=500`, read, site);
    const map: { [key: string]: string } = {};
    values.forEach(field => { map[text(field.Title).toLowerCase()] = text(field.InternalName); });
    return map;
  }
  const error = (e: unknown): string => e instanceof Error ? e.message : 'Could not load SharePoint data.';
  await Promise.all([
    (async () => {
      try {
        const base = `${site}/_api/web/lists(guid'${DOCUMENTS_ID}')`;
        const map = await fields(base);
        const required = ['dork id', 'domain', 'function', 'document type', 'owner', 'lifecycle', 'next review'];
        const missing = required.filter(name => !map[name]);
        if (missing.length) throw new Error(`Documents columns missing: ${missing.join(', ')}.`);
        const rows = await allRows(`${base}/items?$select=*,FileRef,FileLeafRef,FSObjType,FieldValuesAsText&$expand=FieldValuesAsText&$top=500`, read, site);
        result.documents = rows.filter(row => row.FSObjType !== 1 && text(row[map['dork id']]).trim()).map(row => {
          const display = row.FieldValuesAsText || {};
          const get = (title: string): string => text(display[map[title]]) || term(row[map[title]]);
          return { id: text(row[map['dork id']]), title: text(row.Title) || text(row.FileLeafRef), url: fileUrl(text(row.FileRef), site),
            domain: get('domain'), fn: get('function').replace(/^[^:]+:/, ''), type: get('document type'), owner: get('owner'),
            lifecycle: get('lifecycle'), review: text(row[map['next review']]), modified: text(row.Modified) };
        }).sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
      } catch (e) { result.documentError = error(e); }
    })(),
    (async () => {
      try {
        const lists = await allRows(`${site}/_api/web/lists?$select=Id,Title,DefaultViewUrl,DefaultNewFormUrl,DefaultDisplayFormUrl&$filter=Hidden eq false&$top=500`, read, site);
        const list = lists.find(item => text(item.Title).toLowerCase() === 'documentation needs');
        if (!list) throw new Error('Documentation Needs was not found or is not accessible.');
        result.needsUrl = safeUrl(text(list.DefaultViewUrl), site);
        result.newNeedUrl = safeUrl(text(list.DefaultNewFormUrl), site);
        const displayUrl = safeUrl(text(list.DefaultDisplayFormUrl), site);
        const base = `${site}/_api/web/lists(guid'${list.Id}')`;
        const map = await fields(base);
        const missing = ['details', 'need type', 'priority', 'status', 'assigned to', 'target date'].filter(name => !map[name]);
        if (missing.length) throw new Error(`Needs columns missing: ${missing.join(', ')}.`);
        const rows = await allRows(`${base}/items?$select=*,FieldValuesAsText&$expand=FieldValuesAsText&$top=500`, read, site);
        result.needs = rows.map(row => ({ id: Number(row.Id), title: text(row.Title), details: text(row[map.details]),
          url: displayUrl ? `${displayUrl}${displayUrl.includes('?') ? '&' : '?'}ID=${Number(row.Id)}` : '', kind: text(row[map['need type']]),
          priority: text(row[map.priority]), status: text(row[map.status]), assigned: text(row.FieldValuesAsText?.[map['assigned to']]),
          assignedId: Number(row[`${map['assigned to']}Id`]) || 0, target: text(row[map['target date']]) }));
      } catch (e) { result.needsError = error(e); }
    })()
  ]);
  result.loadedAt = new Date();
  return result;
}
