export interface DocumentEntry {
  id: string; title: string; url: string; domain: string; fn: string; type: string;
  owner: string; lifecycle: string; review: string; modified: string;
}
export interface NeedEntry {
  id: number; title: string; details: string; url: string; kind: string;
  priority: string; status: string; assigned: string; assignedId: number; target: string;
}
export interface CalendarEvent { date: string; title: string; url: string; kind: 'review' | 'need'; }
export const DOMAIN_ORDER = ['NETWORK', 'SERVER', 'ENDPOINT', 'IDENTITY', 'DATA', 'CYBERSECURITY', 'FACILITIES', 'BUSINESS', 'PUBLIC'];
export const lower = (value: string): string => (value || '').trim().toLowerCase();
export function dayKey(value: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:$|T)/.exec(value || '');
  if (!match) return '';
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day ? match[0].slice(0, 10) : '';
}
export function todayKey(now: Date = new Date()): string {
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}
export function isActive(doc: DocumentEntry): boolean { return lower(doc.lifecycle) === 'active'; }
export function isOpen(need: NeedEntry): boolean { return !['completed', 'cancelled'].includes(lower(need.status)); }
export function dueReviews(docs: DocumentEntry[], today: string): DocumentEntry[] {
  return docs.filter(doc => isActive(doc) && dayKey(doc.review) !== '' && dayKey(doc.review) <= today)
    .sort((a, b) => dayKey(a.review).localeCompare(dayKey(b.review)) || a.id.localeCompare(b.id, undefined, { numeric: true }));
}
export function sortNeeds(needs: NeedEntry[]): NeedEntry[] {
  const ranks: { [key: string]: number } = { urgent: 0, high: 1, normal: 2, low: 3 };
  return [...needs].sort((a, b) => (ranks[lower(a.priority)] ?? 2) - (ranks[lower(b.priority)] ?? 2)
    || (dayKey(a.target) || '9999').localeCompare(dayKey(b.target) || '9999') || b.id - a.id);
}
export function calendarEvents(docs: DocumentEntry[], needs: NeedEntry[]): CalendarEvent[] {
  return [
    ...docs.filter(doc => isActive(doc) && dayKey(doc.review)).map(doc => ({ date: dayKey(doc.review), title: `${doc.id} · ${doc.title}`, url: doc.url, kind: 'review' as const })),
    ...needs.filter(need => isOpen(need) && dayKey(need.target)).map(need => ({ date: dayKey(need.target), title: need.title, url: need.url, kind: 'need' as const }))
  ].sort((a, b) => a.date.localeCompare(b.date) || a.title.localeCompare(b.title));
}
export function escapeHtml(value: unknown): string {
  return String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] || c));
}
export function safeUrl(value: string, siteUrl: string): string {
  try { const url = new URL(value, siteUrl); return value && url.protocol === 'https:' && url.origin === new URL(siteUrl).origin ? url.href : ''; }
  catch { return ''; }
}
export const REVIEW_LABELS = ['Past review date', 'Due today / approaching', 'Good standing', 'Missing review date'];
export function reviewStatus(doc: DocumentEntry, today: string): string {
 if (!isActive(doc)) return 'Not active';
 const date=dayKey(doc.review); if (!date) return REVIEW_LABELS[3]; if (date<today) return REVIEW_LABELS[0];
 const limit=new Date(`${today}T00:00:00Z`);limit.setUTCDate(limit.getUTCDate()+30);
 return date<=limit.toISOString().slice(0,10)?REVIEW_LABELS[1]:REVIEW_LABELS[2];
}
export function lifecycleCounts(docs: DocumentEntry[]): {label:string;count:number}[] {
 const labels=['Draft','Active','Superseded','Retired'];
 docs.forEach(d=>{const l=d.lifecycle.trim()||'Missing lifecycle';if(!labels.some(v=>lower(v)===lower(l)))labels.push(l);});
 return labels.map(label=>({label,count:docs.filter(d=>lower(d.lifecycle||'Missing lifecycle')===lower(label)).length}));
}
export function reviewCounts(docs:DocumentEntry[],today:string):{label:string;count:number}[]{return REVIEW_LABELS.map(label=>({label,count:docs.filter(d=>reviewStatus(d,today)===label).length}));}
export function fileUrl(fileRef:string,site:string):string{return fileRef.startsWith('/')?safeUrl(fileRef.split('/').map(s=>encodeURIComponent(s)).join('/'),site):'';}
