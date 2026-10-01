// Session-only aliases survive a taskpane reload. They are not embedded in templates/copies.
const STORAGE_KEY = "nerd.document-locations.v1";
let memory = {};
function readLocations(storage) {
    try { const value = JSON.parse(storage.getItem(STORAGE_KEY) || "{}"); return value && typeof value === "object" && !Array.isArray(value) ? value : {}; }
    catch { return memory; }
}
function locationKey(siteId, libraryId, url) {
    const parsed = new URL(url);
    return `${siteId}|${libraryId}|${parsed.origin}${decodeURIComponent(parsed.pathname)}`;
}
export function rememberDocumentLocation(storage, siteId, libraryId, url, item) {
    if (!item?.id || String(item.sharepointIds?.listId || "").toLowerCase() !== libraryId.toLowerCase()) return;
    const entries = readLocations(storage);
    for (const alias of [url, item.webUrl].filter(Boolean)) {
        entries[locationKey(siteId, libraryId, alias)] = item.id;
    }
    memory = Object.fromEntries(Object.entries(entries).slice(-40));
    try { storage.setItem(STORAGE_KEY, JSON.stringify(memory)); } catch { /* No tokens or document metadata in cache. */ }
}
export async function resolveDocumentLocation({ storage, siteId, libraryId, url, getByPath, getById }) {
    let item;
    try { item = await getByPath(); }
    catch (error) {
        // Authorization, network and other failures must never redirect to a cached document.
        if (error.status !== 404) throw error;
        const id = readLocations(storage)[locationKey(siteId, libraryId, url)];
        if (typeof id !== "string" || !id) throw error;
        item = await getById(id);
    }
    if (String(item?.sharepointIds?.listId || "").toLowerCase() !== libraryId.toLowerCase()) {
        throw new Error("The resolved file does not belong to the DORK Documents library.");
    }
    rememberDocumentLocation(storage, siteId, libraryId, url, item);
    return item;
}
