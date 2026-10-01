import { PublicClientApplication, createNestablePublicClientApplication } from "@azure/msal-browser";
import { uploadIntakeFile } from "./intake-upload";
import { configureAutoOpen } from "./auto-open";
import { assertSupportedHost, syncHostMetadata, saveHostDocument } from "./office-host";
import { syncWorkbookServerProperties } from "./workbook-properties";
import { rememberDocumentLocation, resolveDocumentLocation } from "./document-location";

/* global document, Office */

const intakeMode = document.body.dataset.nerdMode === "intake";

const msalConfig = {
    auth: {
        clientId: "f8371eb2-8758-48cf-8fa5-2f9b696d89c4",
        redirectUri: window.location.origin + "/sso-redirect.html",
        authority: "https://login.microsoftonline.com/1169a3a9-3860-4a0b-80aa-410be007cde5",
    },
};

const GRAPH_SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Sites.ReadWrite.All",
    "TermStore.ReadWrite.All",
];

const SHAREPOINT_SCOPES = [
    "https://rocktwpnet.sharepoint.com/AllSites.Write",
];

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DORK_HOST = "rocktwpnet.sharepoint.com";
const DORK_SITE_PATH = "/sites/DORK";
const DORK_SITE_URL = `https://${DORK_HOST}${DORK_SITE_PATH}`;
const API_BASE = typeof NERD_API_BASE !== "undefined"
    ? NERD_API_BASE
    : "http://localhost:3001";

let msalInstance = null;
let currentUser = null;
let dorkSite = null;
let dorkTermGroup = null;
let dorkDocumentsLibrary = null;
let currentDocumentUrl = null;
let currentDriveItem = null;
let currentListItem = null;
let currentDocumentFields = null;
let isSaving = false;

const dorkTermSets = {};
const dorkColumns = {};

const taxonomyData = {
    domain: [],
    function: [],
    "document-type": [],
    "system-platform": [],
    collection: [],
    tags: [],
};

const pickerSelections = {
    "system-platform": [],
    collection: [],
    tags: [],
    audience: [],
};

const personSelections = {
    owner: null,
    responsible: null,
    secondary: null,
};

const personSearchTimers = {};

const contentTypeMap = {
    procedure: "DORK Procedure",
    guide: "DORK Guide",
    plan: "DORK Plan",
    reference: "DORK Reference Doc",
    referencedoc: "DORK Reference Doc",
    referencesheet: "DORK Reference Sheet",
    standard: "DORK Standard",
};

/* STARTUP */

let connecting = false;
let browserInteractive = false;
const browserLoginHint = intakeMode ? new URLSearchParams(window.location.hash.slice(1)).get("login_hint") : null;
if (intakeMode && browserLoginHint) history.replaceState(null, "", window.location.pathname + window.location.search);
async function connectToDork(interactive = false) {
    if (connecting) return;
    connecting = true;
    browserInteractive = interactive;
    const connectButton = document.getElementById("connect-nerd");
    connectButton.disabled = true;
    connectButton.hidden = true;
    document.getElementById("metadata-form").hidden = true;
    document.getElementById("save-metadata").disabled = true;
    setSaveStatus("", "");

    try {
        setConnectionStatus("Connecting to DORK...");

        if (!msalInstance) {
            if (intakeMode) {
                msalInstance = new PublicClientApplication({...msalConfig, cache: {cacheLocation: "localStorage"}});
                await msalInstance.initialize();
                const accounts = msalInstance.getAllAccounts();
                const hinted = browserLoginHint && accounts.find(account => account.username?.toLowerCase() === browserLoginHint.toLowerCase());
                if (hinted || (!browserLoginHint && accounts.length === 1)) msalInstance.setActiveAccount(hinted || accounts[0]);
            } else {
                msalInstance = await createNestablePublicClientApplication(msalConfig);
            }
        }

        currentUser = await graphGet(
            `${GRAPH_ROOT}/me?$select=id,displayName,userPrincipalName,mail`
        );

        dorkSite = await graphGet(
            `${GRAPH_ROOT}/sites/${DORK_HOST}:${DORK_SITE_PATH}`
        );

        await acquireToken(SHAREPOINT_SCOPES);
        setConnectionStatus("Loading DORK metadata...");

        await Promise.all([
            initializeTaxonomy(),
            initializeLibrarySchema(),
        ]);

        if (intakeMode && currentDriveItem) {
            document.getElementById("metadata-form").hidden = false;
            document.getElementById("save-metadata").disabled = false;
            setConnectedStatus();
        } else if (intakeMode) await initializeIntake();
        else await initializeCurrentDocument();
    } catch (error) {
        console.error("NERD initialization failed:", error);

        setConnectionStatus(`Connection failed: ${getErrorMessage(error)}`, true);
        connectButton.textContent = intakeMode ? "Sign in to DORK" : "Retry connection";
        connectButton.hidden = false;

        setSaveStatus(getErrorMessage(error), "error");
    } finally {
        connecting = false;
        browserInteractive = false;
        document.getElementById("connect-nerd").disabled = false;
    }
}

function startInterface() {
    initializeInterface();
    document.getElementById("connect-nerd").onclick = () => connectToDork(true);
    if (intakeMode) {
        const button = document.getElementById("connect-nerd");
        button.textContent = "Sign in to DORK";
        button.hidden = false;
        connectToDork();
    } else connectToDork();
}
if (intakeMode) startInterface();
else Office.onReady(startInterface);

/* AUTHENTICATION */

async function getOfficeLoginHint() {
    if (intakeMode) return currentUser?.userPrincipalName || browserLoginHint || null;
    try {
        const context = await Office.auth?.getAuthContext?.();
        if (context?.userPrincipalName) return context.userPrincipalName;
    } catch (_) { /* Older hosts can lack the Office authentication context. */ }
    return currentUser?.userPrincipalName || null;
}

async function acquireToken(scopes) {
    const resource = scopes[0].startsWith("https://rocktwpnet") ? "sharepoint" : "graph";
    const cached = dialogTokens.get(resource);
    if (!intakeMode && cached && cached.expires > Date.now() + 120000) return cached.token;
    const loginHint = await getOfficeLoginHint();
    const account = loginHint
        ? msalInstance.getAccountByUsername?.(loginHint) || msalInstance.getAllAccounts?.().find(candidate => candidate.username?.toLowerCase() === loginHint.toLowerCase())
        : msalInstance.getActiveAccount?.();
    const request = { scopes, ...(loginHint ? {loginHint} : {}), ...(account ? {account} : {}) };

    try {
        const result = await msalInstance.acquireTokenSilent(request);
        rememberSilentToken(resource, result);
        return result.accessToken;
    } catch (silentError) {
        if (intakeMode) {
            let result;
            try {
                result = await msalInstance.ssoSilent(request);
            } catch (_) {
                if (!browserInteractive) {
                    const button = document.getElementById("connect-nerd");
                    button.textContent = "Sign in to DORK";
                    button.hidden = false;
                    throw new Error("Microsoft requires sign-in. Select Sign in to DORK to continue.");
                }
                result = await msalInstance.acquireTokenPopup(request);
            }
            msalInstance.setActiveAccount(result.account);
            rememberSilentToken(resource, result);
            return result.accessToken;
        }
        // Word on the web needs the host identity to reuse the Entra session.
        const isWeb = Office.context.platform === (Office.PlatformType?.OfficeOnline || "OfficeOnline");
        if (isWeb && loginHint && typeof msalInstance.ssoSilent === "function") {
            try {
                const result = await msalInstance.ssoSilent(request);
                rememberSilentToken(resource, result);
                return result.accessToken;
            } catch (ssoError) {
                console.warn("Word silent sign-in unavailable:", ssoError.errorCode || ssoError.code || "unknown");
            }
        }
        console.warn("Silent token acquisition unavailable:", silentError.errorCode || silentError.code || "unknown");
        return acquireDialogToken(scopes);
    }
}

function rememberSilentToken(resource, result) {
    const expires = result.expiresOn?.getTime();
    if (result.accessToken && Number.isFinite(expires) && expires > Date.now() + 120000) {
        // Reuse the resource token until it approaches expiry; never store refresh tokens.
        dialogTokens.set(resource, {token: result.accessToken, expires});
    }
}

const dialogCacheKey = "nerd-dialog-access-tokens-v1";
const validDialogToken = value => typeof value?.token === "string" && value.token.length > 0 &&
    Number.isFinite(value.expires) && value.expires > Date.now() + 120000;
const dialogTokens = new Map();
// Tab-session storage survives pane reloads; refresh tokens remain in the dialog.
try {
    const saved = JSON.parse(sessionStorage.getItem(dialogCacheKey) || "null");
    if (saved && ["graph", "sharepoint"].every(name => validDialogToken(saved[name]))) {
        for (const name of ["graph", "sharepoint"]) dialogTokens.set(name, saved[name]);
    } else {
        sessionStorage.removeItem(dialogCacheKey);
    }
} catch (_) { /* Storage can be unavailable in restricted Office clients. */ }
let activeAuthDialog = null;
function acquireDialogToken(scopes) {
    const resource = scopes[0].startsWith("https://rocktwpnet") ? "sharepoint" : "graph";
    const cached = dialogTokens.get(resource);
    if (cached && cached.expires > Date.now() + 120000) return Promise.resolve(cached.token);
    if (activeAuthDialog) return activeAuthDialog.then(() => acquireDialogToken(scopes));
    activeAuthDialog = new Promise((resolve, reject) => {
        const nonce = crypto.randomUUID();
        const url = new URL("auth.html", window.location.href);
        url.searchParams.set("resource", resource);
        url.searchParams.set("nonce", nonce);
        if (currentUser?.userPrincipalName) url.searchParams.set("login_hint", currentUser.userPrincipalName);
        Office.context.ui.displayDialogAsync(url.href, {height:60, width:35, displayInIframe:false}, result => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(result.error.message + " Use Retry connection to try again.")); return;
            }
            const dialog = result.value;
            let settled = false;
            const timer = setTimeout(() => { settled = true; dialog.close(); reject(new Error("Sign-in timed out. Use Retry connection to try again.")); }, 300000);
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
                if (settled) return;
                if (event.origin && event.origin !== window.location.origin) return;
                let data;
                try { data = JSON.parse(event.message); } catch (_) { return; }
                if (data.nonce !== nonce) return;
                settled = true;
                clearTimeout(timer); dialog.close();
                if (data.error) { reject(new Error(data.error)); return; }
                const tokens = data.tokens;
                if (!tokens || !["graph", "sharepoint"].every(name =>
                    validDialogToken(tokens[name]))) {
                    reject(new Error("Invalid sign-in response.")); return;
                }
                for (const name of ["graph", "sharepoint"]) dialogTokens.set(name, tokens[name]);
                try { sessionStorage.setItem(dialogCacheKey, JSON.stringify(tokens)); } catch (_) {}
                resolve(tokens[resource].token);
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                if (settled) return;
                settled = true;
                clearTimeout(timer); reject(new Error("Sign-in window closed. Use Retry connection to try again."));
            });
        });
    }).finally(() => { activeAuthDialog = null; });
    return activeAuthDialog;
}

/* GRAPH */

async function graphRequest(url, options = {}) {
    const token = await acquireToken(GRAPH_SCOPES);

    const response = await fetch(url, {
        ...options,
        headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json",
            ...(options.body
                ? { "Content-Type": "application/json" }
                : {}),
            ...(options.headers || {}),
        },
    });

    const text = await response.text();
    let body = null;

    if (text) {
        try {
            body = JSON.parse(text);
        } catch {
            body = text;
        }
    }

    if (!response.ok) {
        const error = new Error(
            `Microsoft Graph returned ${response.status}: ${
                typeof body === "string" ? body : JSON.stringify(body)
            }`
        );
        error.status = response.status;
        throw error;
    }

    return body;
}

function graphGet(url, headers = {}) {
    return graphRequest(url, { method: "GET", headers });
}

async function graphGetAll(url, headers = {}) {
    const values = [];
    let nextUrl = url;

    while (nextUrl) {
        const result = await graphGet(nextUrl, headers);

        if (Array.isArray(result?.value)) {
            values.push(...result.value);
        }

        nextUrl = result?.["@odata.nextLink"] || null;
    }

    return values;
}

/* NERD API / SHAREPOINT REST BRIDGE */

async function sharePointRequest(
    relativeUrl,
    method = "POST",
    body = null,
    headers = {}
) {
    const token = await acquireToken(SHAREPOINT_SCOPES);

    const targetUrl = relativeUrl.startsWith("http")
        ? relativeUrl
        : `${DORK_SITE_URL}${relativeUrl}`;

    const response = await fetch(`${API_BASE}/api/sharepoint`, {
        method: "POST",
        headers: {
            "X-NERD-SharePoint-Authorization": `Bearer ${token}`,
            Accept: "application/json",
            "Content-Type": "application/json",
        },
        body: JSON.stringify({
            url: targetUrl,
            method,
            headers,
            body,
        }),
    });

    const result = await response.json().catch(() => ({}));

    if (!response.ok) {
        const details = result?.details
            ? typeof result.details === "string"
                ? result.details
                : JSON.stringify(result.details)
            : "";

        throw new Error(
            `${result?.error || "SharePoint request failed."}${
                details ? ` ${details}` : ""
            }`
        );
    }

    return result.data;
}

async function ensureSharePointUser(person) {
    const login = person.userPrincipalName || person.mail;

    if (!login) {
        throw new Error(
            `NERD cannot resolve ${person.displayName} to SharePoint.`
        );
    }

    const result = await sharePointRequest(
        "/_api/web/ensureuser",
        "POST",
        { logonName: `i:0#.f|membership|${login}` }
    );

    const id = result?.Id ?? result?.id ?? result?.d?.Id;
    const loginName =
        result?.LoginName ??
        result?.loginName ??
        result?.d?.LoginName;

    if (!id) {
        throw new Error(
            `SharePoint did not return a user ID for ${person.displayName}.`
        );
    }

    return {
        id: String(id),
        loginName: loginName || `i:0#.f|membership|${login}`,
    };
}

async function validateUpdateListItem(formValues) {
    if (!currentListItem?.id) {
        throw new Error(
            "NERD does not have a current SharePoint list item."
        );
    }

    const itemEndpoint =
        `/_api/web/lists(guid'${dorkDocumentsLibrary.id}')` +
        `/items(${currentListItem.id})`;

    const properties = await sharePointRequest(
        `${itemEndpoint}/File/Properties`,
        "GET"
    );

    const fileProperties = properties?.d || properties || {};

    const sharedLockId =
        fileProperties.vti_x005f_sourcecontrollockid ||
        fileProperties.vti_sourcecontrollockid ||
        null;

    const result = await sharePointRequest(
        `${itemEndpoint}/ValidateUpdateListItem`,
        "POST",
        {
            formValues,
            bNewDocumentUpdate: false,
            ...(sharedLockId ? { sharedLockId } : {}),
        }
    );

    const rawResults =
        result?.value ||
        result?.d?.ValidateUpdateListItem ||
        [];

    const results = Array.isArray(rawResults)
        ? rawResults
        : rawResults?.results || [];

    const failures = results.filter(
        (entry) =>
            entry?.HasException === true ||
            entry?.ErrorMessage
    );

    if (failures.length) {
        throw new Error(
            failures
                .map(
                    (entry) =>
                        `${entry.FieldName || "Field"}: ${
                            entry.ErrorMessage ||
                            "SharePoint rejected the value."
                        }`
                )
                .join(" | ")
        );
    }

    return result;
}

/* TAXONOMY */

async function initializeTaxonomy() {
    const groups = await graphGetAll(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}/termStore/groups`
    );

    dorkTermGroup = groups.find(
        (group) => normalizeName(group.displayName) === "dork"
    );

    if (!dorkTermGroup) {
        throw new Error("The DORK term group was not found.");
    }

    const sets = await graphGetAll(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/termStore/groups/${encodeURIComponent(dorkTermGroup.id)}/sets`
    );

    for (const set of sets) {
        dorkTermSets[normalizeName(getTermSetName(set))] = set;
    }

    const [domains, documentTypes, systems, collections, tags] =
        await Promise.all([
            getTopLevelTerms(requireTermSet("Domain").id),
            getTopLevelTerms(requireTermSet("Document Type").id),
            getTopLevelTerms(requireTermSet("System / Platform").id),
            getTopLevelTerms(requireTermSet("Collection").id),
            getTopLevelTerms(requireTermSet("Tags").id),
        ]);

    requireTermSet("Function");

    taxonomyData.domain = domains;
    taxonomyData["document-type"] = documentTypes;
    taxonomyData["system-platform"] = systems;
    taxonomyData.collection = collections;
    taxonomyData.tags = tags;

    populateSingleTermSelect("domain", domains, "Select a domain");
    populateSingleTermSelect(
        "document-type",
        documentTypes,
        "Select a document type"
    );

    renderTaxonomyPicker("system-platform");
    renderTaxonomyPicker("collection");
    renderTaxonomyPicker("tags");
    resetFunctionSelect();
}

function getTermSetName(set) {
    const names = set.localizedNames || [];

    return (
        names.find((name) =>
            String(name.languageTag || "")
                .toLowerCase()
                .startsWith("en")
        )?.name ||
        names[0]?.name ||
        set.displayName ||
        ""
    );
}

function requireTermSet(name) {
    const set = dorkTermSets[normalizeName(name)];

    if (!set) {
        throw new Error(`The DORK term set "${name}" was not found.`);
    }

    return set;
}

async function getTopLevelTerms(setId) {
    const terms = await graphGetAll(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/termStore/sets/${encodeURIComponent(setId)}/children`
    );

    return sortTerms(terms);
}

async function getChildTerms(setId, termId) {
    const terms = await graphGetAll(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/termStore/sets/${encodeURIComponent(setId)}` +
        `/terms/${encodeURIComponent(termId)}/children`
    );

    return sortTerms(terms);
}

function sortTerms(terms) {
    return [...terms].sort((a, b) =>
        getTermLabel(a).localeCompare(
            getTermLabel(b),
            undefined,
            { sensitivity: "base" }
        )
    );
}

function getTermLabel(term) {
    return (
        term?.labels?.find((label) => label.isDefault)?.name ||
        term?.labels?.[0]?.name ||
        ""
    );
}

async function createTagTerm(label) {
    const tagsSet = requireTermSet("Tags");

    const created = await graphRequest(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/termStore/sets/${encodeURIComponent(tagsSet.id)}/children`,
        {
            method: "POST",
            body: JSON.stringify({
                labels: [
                    {
                        languageTag: "en-US",
                        name: label,
                        isDefault: true,
                    },
                ],
            }),
        }
    );

    return created;
}

/* LIBRARY SCHEMA */

async function initializeLibrarySchema() {
    const lists = await graphGetAll(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        "/lists?$select=id,name,displayName,webUrl,list"
    );

    const libraries = lists.filter(
        (list) => list.list?.template === "documentLibrary"
    );

    dorkDocumentsLibrary =
        libraries.find(
            (library) =>
                normalizeName(library.displayName) === "documents"
        ) ||
        libraries.find(
            (library) =>
                normalizeName(library.name) === "documents"
        ) ||
        libraries.find((library) =>
            String(library.webUrl || "")
                .toLowerCase()
                .includes("/shared documents")
        ) ||
        null;

    if (!dorkDocumentsLibrary) {
        throw new Error("The DORK Documents library was not found.");
    }

    const columns = await graphGetAll(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/lists/${encodeURIComponent(dorkDocumentsLibrary.id)}/columns`
    );

    for (const column of columns) {
        if (column.displayName) {
            dorkColumns[normalizeName(column.displayName)] = column;
        }

        if (column.name) {
            dorkColumns[normalizeName(column.name)] = column;
        }
    }

    [
        "Title",
        "Domain",
        "Function",
        "Document Type",
        "System / Platform",
        "Collection",
        "Tags",
        "Classification",
        "Lifecycle",
        "Audience",
        "Owner",
        "Responsible",
        "Secondary",
        "Information Source",
        "Last Reviewed",
        "Next Review",
    ].forEach(requireColumn);

    populateChoiceSelect(
        "classification",
        requireColumn("Classification")
    );
    populateChoiceSelect("lifecycle", requireColumn("Lifecycle"));
    initializeAudience(requireColumn("Audience"));
}

function requireColumn(name) {
    const column = dorkColumns[normalizeName(name)];

    if (!column) {
        throw new Error(`The DORK column "${name}" was not found.`);
    }

    return column;
}

/* CURRENT DOCUMENT */

async function initializeCurrentDocument() {
    assertSupportedHost();
    currentDocumentUrl = await getCurrentDocumentUrl();

    if (!currentDocumentUrl) {
        throw new Error(
            "Save this document to DORK before editing its metadata."
        );
    }

    const relativePath =
        getDorkLibraryRelativePath(currentDocumentUrl);

    currentDriveItem = await resolveDocumentLocation({
        storage: window.sessionStorage,
        siteId: dorkSite.id,
        libraryId: dorkDocumentsLibrary.id,
        url: currentDocumentUrl,
        getByPath: () => getDriveItemByPath(relativePath),
        getById: getDriveItemById,
    });
    currentListItem = await getCurrentListItem(currentDriveItem);
    currentDocumentFields = currentListItem.fields || {};

    await hydrateAllControls();
    // Excel startup must be read-only for worksheet metadata: a partial or stale
    // library response must never erase a previously saved workbook cover.
    if (Office.context.host !== Office.HostType.Excel) await refreshDocumentMetadata();
    await configureAutoOpen();
    document.getElementById("metadata-form").hidden = false;
    enableSave();
    setConnectedStatus();
}

function getCurrentDocumentUrl() {
    return new Promise((resolve, reject) => {
        Office.context.document.getFilePropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                reject(
                    new Error(
                        result.error?.message ||
                        "Office could not determine the current document location."
                    )
                );
                return;
            }

            resolve(String(result.value?.url || "").trim());
        });
    });
}

function getDorkLibraryRelativePath(documentUrl) {
    const parsed = new URL(documentUrl);

    if (parsed.hostname.toLowerCase() !== DORK_HOST) {
        throw new Error(
            "The current document is not stored in the Rockaway Township SharePoint tenant."
        );
    }

    const decodedPath = decodeURIComponent(parsed.pathname);

    const markers = [
        "/sites/DORK/Shared Documents/",
        "/sites/DORK/Documents/",
    ];

    const marker = markers.find((candidate) =>
        decodedPath.toLowerCase().startsWith(candidate.toLowerCase())
    );

    if (!marker) {
        throw new Error(
            "The current document is not stored in the DORK Documents library."
        );
    }

    return decodedPath.substring(marker.length);
}

function encodeGraphPath(path) {
    return path.split("/").map(encodeURIComponent).join("/");
}

async function getDriveItemByPath(path) {
    return graphGet(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/drive/root:/${encodeGraphPath(path)}` +
        "?$select=id,name,webUrl,parentReference,sharepointIds"
    );
}

async function getDriveItemById(id) {
    return graphGet(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/drive/items/${encodeURIComponent(id)}` +
        "?$select=id,name,webUrl,parentReference,sharepointIds"
    );
}

async function getCurrentListItem(driveItem) {
    const itemId = driveItem.sharepointIds?.listItemId;

    if (!itemId) {
        throw new Error(
            "The current document did not expose a SharePoint list item ID."
        );
    }

    return graphGet(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/lists/${encodeURIComponent(dorkDocumentsLibrary.id)}` +
        `/items/${encodeURIComponent(itemId)}?$expand=fields`
    );
}

/* HYDRATION */

async function hydrateAllControls() {
    setTextValue("title", getFieldValue("Title") || String(currentDriveItem?.name || "").replace(/\.[^.]+$/, ""));
    setChoiceValue(
        "classification",
        getFieldValue("Classification")
    );
    setChoiceValue("lifecycle", getFieldValue("Lifecycle") || "Draft");
    setTextValue(
        "information-source",
        getFieldValue("Information Source")
    );
    setDateValue("last-reviewed", getFieldValue("Last Reviewed"));
    setDateValue("next-review", getFieldValue("Next Review"));

    setPickerSelectionsByLabels(
        "audience",
        choiceLabelsFromField(getFieldValue("Audience"))
    );

    await hydrateTaxonomy();

    await Promise.all([
        hydratePerson("owner", "Owner"),
        hydratePerson("responsible", "Responsible"),
        hydratePerson("secondary", "Secondary"),
    ]);
}

async function refreshDocumentMetadata() {
    const value = id => document.getElementById(id)?.value || "";
    const choice = id => value(id) ? document.getElementById(id)?.selectedOptions?.[0]?.textContent?.trim() || value(id) : "";
    const labels = id => pickerSelections[id].map(item => item.label).join("; ");
    const person = id => personSelections[id]?.displayName || "";
    const type = choice("document-type");
    const values = {
        DORK_DocumentId: "", DORK_Title: value("title") || "[Document Title]", DORK_Domain: choice("domain"),
        DORK_Function: choice("function"), DORK_DocumentType: type,
        DORK_TypeHeading: `DORK ${type.toUpperCase()}`,
        DORK_System: labels("system-platform"), DORK_Collection: labels("collection"),
        DORK_Tags: labels("tags"), DORK_Classification: value("classification"),
        DORK_Audience: labels("audience"), DORK_Lifecycle: value("lifecycle"),
        DORK_Owner: person("owner"), DORK_Responsible: person("responsible"),
        DORK_Secondary: person("secondary"), DORK_InformationSource: value("information-source"),
        DORK_LastReviewed: value("last-reviewed"), DORK_NextReview: value("next-review"),
    };
    // Read an assigned identifier if the library already provides one. Never invent
    // a permanent cross-library ID from a library-local item number.
    const idColumn = Object.values(dorkColumns).find(column =>
        ["documentid", "dorkid"].includes(normalizeName(column.displayName).replace(/[^a-z0-9]/g, "")));
    const documentId = idColumn && currentDocumentFields?.[idColumn.name];
    if (documentId) values.DORK_DocumentId = `${documentId} `;
    try {
        await syncHostMetadata(values);
        return true;
    } catch (error) {
        console.warn("NERD document field refresh failed:", getErrorMessage(error));
        if (Office.context.host === Office.HostType.Excel) throw error;
        return false;
    }
}

function getFieldValue(displayName) {
    const column = requireColumn(displayName);
    return currentDocumentFields?.[column.name];
}

function setTextValue(id, value) {
    const element = document.getElementById(id);

    if (element) {
        element.value = value == null ? "" : String(value);
    }
}

function setDateValue(id, value) {
    setTextValue(id, value ? String(value).substring(0, 10) : "");
}

function setChoiceValue(id, value) {
    const select = document.getElementById(id);
    if (!select) return;

    const wanted = normalizeName(value);

    const option = Array.from(select.options).find(
        (candidate) =>
            normalizeName(candidate.value) === wanted ||
            normalizeName(candidate.textContent) === wanted
    );

    select.value = option?.value || "";
}

async function hydrateTaxonomy() {
    const domainLabel =
        taxonomyLabelsFromField(getFieldValue("Domain"))[0] || "";

    setTermSelectByLabel("domain", domainLabel);

    if (domainLabel) {
        await updateFunctions();

        const functionLabels =
            taxonomyLabelsFromField(getFieldValue("Function"));

        const functionLabel =
            functionLabels
                .map((label) =>
                    String(label || "")
                        .split(/[;:>\\/]/)
                        .pop()
                        .trim()
                )
                .find(Boolean) || "";

        setTermSelectByLabel("function", functionLabel);
    }

    setTermSelectByLabel(
        "document-type",
        taxonomyLabelsFromField(getFieldValue("Document Type"))[0] ||
            String(currentListItem?.contentType?.name || "").replace(/^DORK\s+/i, "")
    );

    setPickerSelectionsByLabels(
        "system-platform",
        taxonomyLabelsFromField(getFieldValue("System / Platform"))
    );
    setPickerSelectionsByLabels(
        "collection",
        taxonomyLabelsFromField(getFieldValue("Collection"))
    );
    setPickerSelectionsByLabels(
        "tags",
        taxonomyLabelsFromField(getFieldValue("Tags"))
    );
}

async function hydratePerson(pickerId, displayName) {
    const column = requireColumn(displayName);

    const lookupValue =
        currentDocumentFields?.[`${column.name}LookupId`];

    const lookupId =
        Array.isArray(lookupValue) ? lookupValue[0] : lookupValue;

    if (!lookupId) {
        clearPerson(pickerId, false);
        return;
    }

    let display = currentDocumentFields?.[column.name];
    if (Array.isArray(display)) display = display[0];

    let resolvedUser = null;

    try {
        resolvedUser = await sharePointRequest(
            `/_api/web/GetUserById(${encodeURIComponent(String(lookupId))})`,
            "GET"
        );
    } catch (error) {
        console.warn(
            `Unable to resolve SharePoint user ${lookupId}:`,
            error
        );
    }

    const resolvedDisplayName =
        resolvedUser?.Title ??
        resolvedUser?.title ??
        resolvedUser?.d?.Title ??
        resolvedUser?.d?.title ??
        display ??
        `SharePoint User ${lookupId}`;

    const resolvedEmail =
        resolvedUser?.Email ??
        resolvedUser?.email ??
        resolvedUser?.d?.Email ??
        resolvedUser?.d?.email ??
        "";

    const resolvedLoginName =
        resolvedUser?.LoginName ??
        resolvedUser?.loginName ??
        resolvedUser?.d?.LoginName ??
        resolvedUser?.d?.loginName ??
        "";

    const person = {
        displayName: String(resolvedDisplayName),
        sharePointLookupId: String(lookupId),
        userPrincipalName: String(
            resolvedEmail ||
            extractMembershipLogin(resolvedLoginName) ||
            ""
        ),
        mail: String(resolvedEmail || ""),
        loginName: String(resolvedLoginName || ""),
        existing: true,
    };

    personSelections[pickerId] = person;
    renderSelectedPerson(pickerId);
}
/* SINGLE SELECTS */

function populateSingleTermSelect(id, terms, placeholder) {
    const select = document.getElementById(id);
    if (!select) return;

    select.innerHTML = "";

    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = placeholder;
    select.appendChild(empty);

    for (const term of terms) {
        const label = getTermLabel(term);
        if (!label) continue;

        const option = document.createElement("option");
        option.value = term.id;
        option.textContent = label;
        option.dataset.termLabel = label;
        select.appendChild(option);
    }

    select.disabled = false;
}

function setTermSelectByLabel(id, label) {
    const select = document.getElementById(id);
    if (!select) return;

    const wanted = normalizeName(label);

    const option = Array.from(select.options).find(
        (candidate) =>
            normalizeName(
                candidate.dataset.termLabel || candidate.textContent
            ) === wanted
    );

    select.value = option?.value || "";
}

function getSelectedTerm(id) {
    const select = document.getElementById(id);
    if (!select) return null;

    const option = select.options[select.selectedIndex];
    if (!option?.value) return null;

    return {
        id: option.value,
        label: option.dataset.termLabel || option.textContent,
    };
}

async function updateFunctions() {
    const domain = getSelectedTerm("domain");

    if (!domain) {
        resetFunctionSelect();
        return;
    }

    const select = document.getElementById("function");
    select.disabled = true;
    select.innerHTML =
        '<option value="">Loading functions...</option>';

    const functionSet = requireTermSet("Function");
    const parents = await getTopLevelTerms(functionSet.id);

    const parent = parents.find(
        (term) =>
            normalizeName(getTermLabel(term)) ===
            normalizeName(domain.label)
    );

    if (!parent) {
        throw new Error(
            `No Function parent was found for Domain "${domain.label}".`
        );
    }

    const children = await getChildTerms(functionSet.id, parent.id);
    taxonomyData.function = children;

    populateSingleTermSelect(
        "function",
        children,
        "Select a function"
    );
}

function resetFunctionSelect() {
    taxonomyData.function = [];

    const select = document.getElementById("function");
    if (!select) return;

    select.innerHTML =
        '<option value="">Select a domain first</option>';
    select.disabled = true;
}

function populateChoiceSelect(id, column) {
    const choices = column.choice?.choices;

    if (!Array.isArray(choices)) {
        throw new Error(
            `"${column.displayName}" is not a SharePoint Choice column.`
        );
    }

    const select = document.getElementById(id);
    select.innerHTML = "";

    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent =
        `Select ${column.displayName.toLowerCase()}`;
    select.appendChild(empty);

    for (const choice of choices) {
        const option = document.createElement("option");
        option.value = choice;
        option.textContent = choice;
        select.appendChild(option);
    }

    select.disabled = false;
}

/* COMPACT MULTI-PICKERS */

function initializeAudience(column) {
    const choices = column.choice?.choices;

    if (!Array.isArray(choices)) {
        throw new Error(
            `"${column.displayName}" is not a SharePoint Choice column.`
        );
    }

    taxonomyData.audience = choices.map((choice) => ({
        id: choice,
        label: choice,
        isChoice: true,
    }));

    renderTaxonomyPicker("audience");
}

function getPickerSource(id) {
    const values = taxonomyData[id] || [];

    return values.map((value) => {
        if (value.label) return value;

        return {
            id: value.id,
            label: getTermLabel(value),
            term: value,
        };
    });
}

function renderTaxonomyPicker(id, filter = "") {
    const menu = document.getElementById(`${id}-menu`);
    if (!menu) return;

    const query = normalizeSearchText(filter);

    const selectedIds = new Set(
        pickerSelections[id].map((item) => String(item.id))
    );

    let items = getPickerSource(id);

    if (query) {
        items = items.filter((item) =>
            normalizeSearchText(item.label).includes(query)
        );
    }

    menu.innerHTML = "";

    for (const item of items) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "nerd-picker-option";

        if (selectedIds.has(String(item.id))) {
            button.classList.add("is-selected");
        }

        button.textContent = item.label;

        button.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            togglePickerSelection(id, item);
        });

        menu.appendChild(button);
    }

    if (id === "tags" && query) {
        const exactExisting = getPickerSource("tags").some(
            (item) =>
                normalizeName(item.label) === normalizeName(filter)
        );

        const exactSelected = pickerSelections.tags.some(
            (item) =>
                normalizeName(item.label) === normalizeName(filter)
        );

        if (!exactExisting && !exactSelected) {
            const createButton = document.createElement("button");
            createButton.type = "button";
            createButton.className = "nerd-picker-create";
            createButton.textContent = `Add "${filter.trim()}"`;

            createButton.addEventListener("click", (event) => {
                event.preventDefault();
                event.stopPropagation();
                addPendingTag(filter.trim());
            });

            menu.appendChild(createButton);
        }
    }

    if (!menu.children.length) {
        const empty = document.createElement("div");
        empty.className = "nerd-picker-empty";
        empty.textContent = "No matches";
        menu.appendChild(empty);
    }

    renderPickerChips(id);
}

function togglePickerSelection(id, item) {
    const index = pickerSelections[id].findIndex(
        (selected) => String(selected.id) === String(item.id)
    );

    if (index >= 0) {
        pickerSelections[id].splice(index, 1);
    } else {
        pickerSelections[id].push({
            id: item.id,
            label: item.label,
            term: item.term || null,
            isChoice: item.isChoice || false,
            isNew: item.isNew || false,
        });
    }

    renderTaxonomyPicker(
        id,
        document.getElementById(`${id}-search`)?.value || ""
    );
}

function addPendingTag(label) {
    const clean = String(label || "").trim();
    if (!clean) return;

    pickerSelections.tags.push({
        id: `new:${clean.toLowerCase()}`,
        label: clean,
        isNew: true,
        term: null,
    });

    const search = document.getElementById("tags-search");
    if (search) search.value = "";

    renderTaxonomyPicker("tags");
    closePicker("tags");
}

function renderPickerChips(id) {
    const container = document.getElementById(`${id}-selected`);
    if (!container) return;

    container.innerHTML = "";

    for (const item of pickerSelections[id]) {
        const chip = document.createElement("span");
        chip.className = "nerd-chip";

        const label = document.createElement("span");
        label.className = "nerd-chip-label";
        label.textContent = item.label;

        const remove = document.createElement("button");
        remove.type = "button";
        remove.className = "nerd-chip-remove";
        remove.setAttribute("aria-label", `Remove ${item.label}`);
        remove.innerHTML =
            '<i class="ms-Icon ms-Icon--Cancel"></i>';

        remove.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();

            pickerSelections[id] = pickerSelections[id].filter(
                (selected) => String(selected.id) !== String(item.id)
            );

            renderTaxonomyPicker(
                id,
                document.getElementById(`${id}-search`)?.value || ""
            );
        });

        chip.append(label, remove);
        container.appendChild(chip);
    }
}

function setPickerSelectionsByLabels(id, labels) {
    const source = getPickerSource(id);
    const wanted = labels.map(normalizeName);

    pickerSelections[id] = [];

    for (const label of wanted) {
        const item = source.find(
            (candidate) => normalizeName(candidate.label) === label
        );

        if (item) {
            pickerSelections[id].push({
                id: item.id,
                label: item.label,
                term: item.term || null,
                isChoice: item.isChoice || false,
                isNew: false,
            });
        }
    }

    renderTaxonomyPicker(id);
}

function openPicker(id) {
    closeAllPickers(id);
    closeAllPersonMenus();

    const menu = document.getElementById(`${id}-menu`);
    if (!menu) return;

    renderTaxonomyPicker(
        id,
        document.getElementById(`${id}-search`)?.value || ""
    );

    menu.classList.remove("nerd-hidden");
}

function closePicker(id) {
    document.getElementById(`${id}-menu`)
        ?.classList.add("nerd-hidden");
}

function togglePickerMenu(id) {
    const menu = document.getElementById(`${id}-menu`);
    if (!menu) return;

    if (menu.classList.contains("nerd-hidden")) {
        openPicker(id);
    } else {
        closePicker(id);
    }
}

function closeAllPickers(except = null) {
    ["system-platform", "collection", "tags", "audience"].forEach(
        (id) => {
            if (id !== except) closePicker(id);
        }
    );
}

/* PEOPLE */

function schedulePersonSearch(id, value) {
    clearTimeout(personSearchTimers[id]);

    const query = String(value || "").trim();

    if (query.length < 2) {
        closePersonMenu(id);
        return;
    }

    personSearchTimers[id] = setTimeout(async () => {
        try {
            const users = await searchDirectoryUsers(query);
            renderPersonMenu(id, users);
        } catch (error) {
            console.error("Person search failed:", error);
            renderPersonMessage(id, "Unable to search people");
        }
    }, 300);
}

async function searchDirectoryUsers(query) {
    const safe = String(query)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"');

    const search =
        `"displayName:${safe}" OR ` +
        `"userPrincipalName:${safe}" OR ` +
        `"mail:${safe}"`;

    const result = await graphGet(
        `${GRAPH_ROOT}/users?` +
        "$select=id,displayName,mail,userPrincipalName&" +
        `$search=${encodeURIComponent(search)}&$top=10`,
        { ConsistencyLevel: "eventual" }
    );

    return (result?.value || [])
        .filter(
            (user) =>
                user.id &&
                user.displayName &&
                user.userPrincipalName
        )
        .sort((a, b) =>
            a.displayName.localeCompare(b.displayName)
        );
}

function renderPersonMenu(id, users) {
    closeAllPersonMenus(id);
    closeAllPickers();

    const menu = document.getElementById(`${id}-menu`);
    if (!menu) return;

    menu.innerHTML = "";

    if (!users.length) {
        renderPersonMessage(id, "No people found");
        return;
    }

    for (const user of users) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "nerd-person-option";

        const name = document.createElement("span");
        name.className = "nerd-person-option-name";
        name.textContent = user.displayName;

        const detail = document.createElement("span");
        detail.className = "nerd-person-option-detail";
        detail.textContent = user.mail || user.userPrincipalName;

        button.append(name, detail);

        button.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            selectPerson(id, user);
        });

        menu.appendChild(button);
    }

    menu.classList.remove("nerd-hidden");
}

function renderPersonMessage(id, message) {
    const menu = document.getElementById(`${id}-menu`);
    if (!menu) return;

    menu.innerHTML = "";

    const empty = document.createElement("div");
    empty.className = "nerd-person-empty";
    empty.textContent = message;

    menu.appendChild(empty);
    menu.classList.remove("nerd-hidden");
}

function selectPerson(id, user) {
    personSelections[id] = {
        entraObjectId: user.id,
        displayName: user.displayName,
        mail: user.mail || "",
        userPrincipalName: user.userPrincipalName,
        sharePointLookupId: null,
        loginName: "",
        existing: false,
    };

    renderSelectedPerson(id);
    closePersonMenu(id);
}

function renderSelectedPerson(id) {
    const container = document.getElementById(`${id}-selected`);
    if (!container) return;

    container.innerHTML = "";

    const person = personSelections[id];
    if (!person) return;

    const chip = document.createElement("span");
    chip.className = "nerd-chip nerd-person-chip";

    const label = document.createElement("span");
    label.className = "nerd-chip-label";
    label.textContent = person.displayName;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "nerd-chip-remove";
    remove.setAttribute("aria-label", `Remove ${person.displayName}`);
    remove.innerHTML =
        '<i class="ms-Icon ms-Icon--Cancel"></i>';

    remove.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        clearPerson(id, true);
    });

    chip.append(label, remove);
    container.appendChild(chip);
}

function clearPerson(id, focus = false) {
    personSelections[id] = null;
    renderSelectedPerson(id);

    const input = document.getElementById(`${id}-search`);

    if (input) {
        input.value = "";
        if (focus) input.focus();
    }

    closePersonMenu(id);
}

function closePersonMenu(id) {
    const menu = document.getElementById(`${id}-menu`);

    if (menu) {
        menu.innerHTML = "";
        menu.classList.add("nerd-hidden");
    }
}

function closeAllPersonMenus(except = null) {
    ["owner", "responsible", "secondary"].forEach((id) => {
        if (id !== except) closePersonMenu(id);
    });
}
/* SAVE */

async function requestNumbering(action) {
    const token = await acquireToken(SHAREPOINT_SCOPES);
    const response = await fetch(`${API_BASE}/api/numbering`, {
        method: "POST",
        headers: { "X-NERD-SharePoint-Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({ action, itemId: currentListItem.id }),
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || "Numbering request failed.");
    if (result.enabled === false) return null;
    if (result.enabled !== true || !/^DORK-\d{4,}$/.test(result.documentId || "")) throw new Error("Numbering did not return a verified identifier.");
    return result.documentId;
}

async function saveMetadata() {
    if (isSaving) return;

    if (!currentDriveItem || !currentListItem) {
        setSaveStatus(
            "NERD has not located the current DORK document.",
            "error"
        );
        return;
    }

    isSaving = true;

    const saveButton = document.getElementById("save-metadata");
    if (saveButton) saveButton.disabled = true;

    try {
        setSaveStatus("Validating metadata...", "working");
        validateRequiredMetadata();
        buildDesiredFilename(); // Reject invalid filenames before reserving a number.

        setSaveStatus("Preparing taxonomy...", "working");
        await materializeNewTags();

        setSaveStatus("Resolving people...", "working");
        await resolvePeople();

        setSaveStatus("Resolving document type...", "working");
        const contentType = await resolveDesiredContentType();
        const formValues = buildSharePointFormValues(contentType);
        if (typeof intakeMode !== "undefined" && intakeMode) await assertIntakeFileClosed();
        const documentId = await requestNumbering("reserve");
        if (documentId) formValues.push({ FieldName: requireColumn("DORK ID").name, FieldValue: documentId });
        // Decide before verification replaces the loaded SharePoint fields.
        let desiredFilename = buildDesiredFilename();
        if (documentId) {
            const dot = currentDriveItem.name.lastIndexOf(".");
            const extension = dot > 0 ? currentDriveItem.name.substring(dot) : (typeof intakeMode !== "undefined" && intakeMode ? "" : ".docx");
            const title = document.getElementById("title").value.trim().replace(/^DORK-\d+\s+/i, "");
            desiredFilename = `${documentId} ${title}${extension}`;
        }

        const browserHost = typeof intakeMode !== "undefined" && intakeMode;
        const excelHost = !browserHost && Office.context.host === Office.HostType.Excel;
        let pageUpdated = false;
        if (excelHost) {
            // Refresh from the user's selections before hydration replaces them.
            // Include the newly reserved identifier on the first save.
            if (documentId) currentDocumentFields[requireColumn("DORK ID").name] = documentId;
            setSaveStatus("Saving workbook fields...", "working");
            pageUpdated = await refreshDocumentMetadata();
            if (!pageUpdated) throw new Error("Workbook fields could not update. Metadata save stopped.");
            const people = Object.fromEntries([["Owner", "owner"], ["Responsible", "responsible"], ["Secondary", "secondary"]]
                .map(([name, picker]) => [requireColumn(name).name, personSelections[picker]?.displayName || ""]));
            await syncWorkbookServerProperties(formValues, dorkColumns, people, contentType);
            await saveHostDocument();
        }
        setSaveStatus("Saving metadata...", "working");
        await validateUpdateListItem(formValues);

        setSaveStatus("Verifying SharePoint...", "working");
        await verifySavedMetadata(contentType, currentDriveItem.name);
        if (!excelHost && !browserHost) {
            pageUpdated = await refreshDocumentMetadata();
            if (!pageUpdated) throw new Error("Metadata saved, but document fields could not update. Filename unchanged; reopen in Editing mode and retry.");
            setSaveStatus("Saving document...", "working");
            await saveHostDocument();
        }

        // Finish all document edits before changing the path underneath its open session.
        // No content-control writes or document saves may follow this rename.
        if (desiredFilename !== currentDriveItem.name) {
            setSaveStatus("Renaming document...", "working");
            await validateUpdateListItem([{ FieldName: "FileLeafRef", FieldValue: desiredFilename }]);
        }
        await verifySavedMetadata(contentType, desiredFilename);
        rememberDocumentLocation(window.sessionStorage, dorkSite.id, dorkDocumentsLibrary.id, currentDocumentUrl, currentDriveItem);

        if (documentId && String(getFieldValue("DORK ID") || "") !== documentId) throw new Error("SharePoint did not verify: DORK ID.");
        let registryUpdated = true;
        if (documentId) {
            try { await requestNumbering("applied"); } catch { registryUpdated = false; }
        }
        if (!registryUpdated) {
            setSaveStatus("Metadata and number saved and verified. Registry completion is pending; save again to retry with the same number.", "error");
            return;
        }
        if (browserHost) {
            setSaveStatus("Metadata, filename, and DORK ID saved and verified.", "success");
            try { await refreshIntakeFiles(); } catch { /* Verified save remains successful if list refresh fails. */ }
            return;
        }
        setSaveStatus(pageUpdated
            ? "Metadata saved and verified. Document fields updated."
            : "Metadata saved and verified. Document fields could not refresh; reopen NERD to retry.",
            pageUpdated ? "success" : "error");
    } catch (error) {
        console.error("NERD save failed:", error);
        setSaveStatus(
            `Save failed: ${getErrorMessage(error)}`,
            "error"
        );
    } finally {
        isSaving = false;
        if (saveButton) saveButton.disabled = false;
        if (typeof intakeMode !== "undefined" && intakeMode) renderIntakeFiles();
    }
}

function validateRequiredMetadata() {
    const required = [
        ["title", "Title"],
        ["domain", "Domain"],
        ["function", "Function"],
        ["document-type", "Document Type"],
        ["classification", "Classification"],
        ["lifecycle", "Lifecycle"],
    ];

    for (const [id, label] of required) {
        const value = String(
            document.getElementById(id)?.value || ""
        ).trim();

        if (!value) {
            throw new Error(`${label} is required.`);
        }
    }

    if (!personSelections.owner) {
        throw new Error("Owner is required.");
    }
}

async function materializeNewTags() {
    for (const tag of pickerSelections.tags) {
        if (!tag.isNew) continue;

        const created = await createTagTerm(tag.label);
        tag.id = created.id;
        tag.term = created;
        tag.isNew = false;
        taxonomyData.tags.push(created);
    }

    taxonomyData.tags = sortTerms(taxonomyData.tags);
    renderTaxonomyPicker("tags");
}

async function resolvePeople() {
    for (const id of ["owner", "responsible", "secondary"]) {
        const person = personSelections[id];
        if (!person) continue;

        if (person.sharePointLookupId) continue;

        const resolved = await ensureSharePointUser(person);
        person.sharePointLookupId = resolved.id;
        person.loginName = resolved.loginName;
    }
}

async function resolveDesiredContentType() {
    const documentType = getSelectedTerm("document-type");
    if (!documentType) return null;

    const desiredName =
        (intakeMode && normalizeName(documentType.label) === "reference" && /\.xlsx$/i.test(currentDriveItem.name)
            ? "DORK Reference Sheet" : contentTypeMap[normalizeName(documentType.label)]);

    // Uploaded visuals and other non-template types belong to the general
    // DORK content type, never whichever template happened to be the default.
    if (!desiredName && !intakeMode) return null;

    const contentTypes = await graphGetAll(
        `${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}` +
        `/lists/${encodeURIComponent(dorkDocumentsLibrary.id)}` +
        "/contentTypes?$select=id,name"
    );

    const found = contentTypes.find(
        (contentType) =>
            normalizeName(contentType.name) ===
            normalizeName(desiredName || `DORK ${documentType.label}`)
    ) || (!desiredName && intakeMode ? contentTypes.find(type => normalizeName(type.name) === "dorkdocument") : null);

    if (!found) {
        throw new Error(
            `The SharePoint content type "${desiredName || "DORK Document"}" is not available in the DORK Documents library.`
        );
    }

    return found;
}

function buildSharePointFormValues(contentType) {
    const values = [];

    addFormValue(
        values,
        "Title",
        document.getElementById("title").value.trim()
    );
    addFormValue(
        values,
        "Domain",
        taxonomySingleValue(getSelectedTerm("domain"))
    );
    addFormValue(
        values,
        "Function",
        taxonomySingleValue(getSelectedTerm("function"))
    );
    addFormValue(
        values,
        "Document Type",
        taxonomySingleValue(getSelectedTerm("document-type"))
    );
    addFormValue(
        values,
        "System / Platform",
        taxonomyMultiValue(pickerSelections["system-platform"])
    );
    addFormValue(
        values,
        "Collection",
        taxonomyMultiValue(pickerSelections.collection)
    );
    addFormValue(
        values,
        "Tags",
        taxonomyMultiValue(pickerSelections.tags)
    );
    addFormValue(
        values,
        "Classification",
        document.getElementById("classification").value
    );
    addFormValue(
        values,
        "Lifecycle",
        document.getElementById("lifecycle").value
    );
    addFormValue(
        values,
        "Audience",
        pickerSelections.audience.map((item) => item.label).join(";#")
    );

    addPersonFormValue(values, "Owner", personSelections.owner);
    addPersonFormValue(
        values,
        "Responsible",
        personSelections.responsible
    );
    addPersonFormValue(values, "Secondary", personSelections.secondary);

    addFormValue(
        values,
        "Information Source",
        document.getElementById("information-source").value.trim()
    );
    addFormValue(
        values,
        "Last Reviewed",
        sharePointDateValue(document.getElementById("last-reviewed").value)
    );
    addFormValue(
        values,
        "Next Review",
        sharePointDateValue(document.getElementById("next-review").value)
    );

    if (contentType) {
        values.push({
            FieldName: "ContentTypeId",
            FieldValue: contentType.id,
        });
    }

    return values;
}

function addFormValue(values, displayName, value) {
    const column = requireColumn(displayName);

    values.push({
        FieldName: column.name,
        FieldValue: value == null ? "" : String(value),
    });
}

function addPersonFormValue(values, displayName, person) {
    const column = requireColumn(displayName);

    if (!person) {
        values.push({
            FieldName: column.name,
            FieldValue: "",
        });
        return;
    }

    // Preserve existing unresolved people without inventing a login.
    if (person.existing && !person.loginName) return;

    let login =
        person.loginName ||
        person.userPrincipalName ||
        person.mail;

    if (!login) {
        throw new Error(
            `${displayName} could not be resolved to a SharePoint account.`
        );
    }

    if (!login.startsWith("i:")) {
        login = `i:0#.f|membership|${login}`;
    }

    values.push({
        FieldName: column.name,
        FieldValue: JSON.stringify([{ Key: login }]),
    });
}

function taxonomySingleValue(term) {
    if (!term) return "";
    return `${term.label}|${term.id}`;
}

function taxonomyMultiValue(items) {
    if (!items.length) return "";

    // ValidateUpdateListItem expects label/GUID pairs, without WssId prefixes.
    return items.map((item) => `${item.label}|${item.id}`).join(";");
}

function sharePointDateValue(value) {
    if (!value) return "";
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
    if (!match) throw new Error("Review dates must use YYYY-MM-DD.");
    const [, year, month, day] = match;
    const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
    if (date.getUTCFullYear() !== Number(year) || date.getUTCMonth() + 1 !== Number(month) || date.getUTCDate() !== Number(day)) {
        throw new Error("Review date is not a valid calendar date.");
    }
    // DORK's form validator explicitly requires M/D/YYYY. Keep the calendar day.
    return `${Number(month)}/${Number(day)}/${year}`;
}

function buildDesiredFilename() {
    const title = document.getElementById("title").value.trim();

    if (!title) {
        throw new Error("Title is required.");
    }

    if (/["*:<>?\\/|]/.test(title)) {
        throw new Error(
            'Title cannot contain these filename characters: " * : < > ? \\ / |'
        );
    }

    if (/[. ]$/.test(title)) {
        throw new Error("Title cannot end with a period or space.");
    }

    const currentName = currentDriveItem.name;
    const loadedTitle = String(getFieldValue("Title") || "").trim();
    // Copies may share a Title while having deliberately different filenames.
    if (title === loadedTitle) return currentName;
    const dot = currentName.lastIndexOf(".");
    const extension = dot > 0 ? currentName.substring(dot) : (typeof intakeMode !== "undefined" && intakeMode ? "" : ".docx");

    return `${title}${extension}`;
}

async function verifySavedMetadata(contentType, desiredFilename) {
    currentDriveItem = await getDriveItemById(currentDriveItem.id);
    currentListItem = await getCurrentListItem(currentDriveItem);
    currentDocumentFields = currentListItem.fields || {};

    const failures = [];
    const expectedTitle =
        document.getElementById("title").value.trim();

    if (String(getFieldValue("Title") || "") !== expectedTitle) {
        failures.push("Title");
    }

    for (const [displayName, pickerId] of [
        ["Domain", "domain"], ["Function", "function"], ["Document Type", "document-type"],
    ]) {
        const term = getSelectedTerm(pickerId);
        verifyLabelSet(failures, displayName, taxonomyLabelsFromField(getFieldValue(displayName)), term ? [term.label] : []);
    }
    for (const [displayName, pickerId] of [
        ["System / Platform", "system-platform"], ["Collection", "collection"], ["Tags", "tags"],
    ]) {
        verifyLabelSet(failures, displayName, taxonomyLabelsFromField(getFieldValue(displayName)), pickerSelections[pickerId].map(item => item.label));
    }
    verifyLabelSet(failures, "Audience", choiceLabelsFromField(getFieldValue("Audience")), pickerSelections.audience.map(item => item.label));
    for (const [displayName, id] of [["Classification", "classification"], ["Lifecycle", "lifecycle"]]) {
        if (String(getFieldValue(displayName) || "") !== document.getElementById(id).value) failures.push(displayName);
    }
    for (const [displayName, id] of [["Information Source", "information-source"], ["Last Reviewed", "last-reviewed"], ["Next Review", "next-review"]]) {
        const actual = String(getFieldValue(displayName) || "");
        const expected = document.getElementById(id).value.trim();
        if ((id === "information-source" ? actual : actual.substring(0, 10)) !== expected) failures.push(displayName);
    }
    for (const [displayName, id] of [["Owner", "owner"], ["Responsible", "responsible"], ["Secondary", "secondary"]]) {
        const column = requireColumn(displayName);
        const raw = currentDocumentFields?.[`${column.name}LookupId`];
        const actualIds = (Array.isArray(raw) ? raw : raw == null || raw === "" ? [] : [raw]).map(String);
        const person = personSelections[id];
        const expectedIds = person ? [String(person.sharePointLookupId)] : [];
        if (!sameValueSet(actualIds, expectedIds)) failures.push(displayName);
    }

    if (currentDriveItem.name !== desiredFilename) {
        failures.push("Filename");
    }

    if (contentType) {
        const actualContentTypeId =
            String(currentListItem?.contentType?.id || currentDocumentFields?.ContentTypeId || "");

        if (
            !actualContentTypeId.toLowerCase().startsWith(
                String(contentType.id).toLowerCase()
            )
        ) {
            failures.push(`Content Type (expected ${contentType.name || contentType.id}; returned ${currentListItem?.contentType?.name || actualContentTypeId || "no content type ID"})`);
        }
    }

    if (failures.length) {
        throw new Error(
            `SharePoint did not verify: ${failures.join(", ")}.`
        );
    }

    await hydrateAllControls();
}

function sameValueSet(actual, expected) {
    const left = [...new Set(actual)].sort();
    const right = [...new Set(expected)].sort();
    return left.length === right.length && left.every((value, index) => value === right[index]);
}

function verifyLabelSet(failures, displayName, actual, expected) {
    const normalize = value => normalizeName(displayName === "Function"
        ? String(value || "").split(/[;:>\\/]/).pop().trim() : value);
    if (!sameValueSet(actual.map(normalize), expected.map(normalize))) failures.push(displayName);
}

/* FIELD PARSING */

function taxonomyLabelsFromField(value) {
    if (value == null || value === "") return [];

    if (Array.isArray(value)) {
        return value.flatMap(taxonomyLabelsFromField).filter(Boolean);
    }

    if (typeof value === "object") {
        const label =
            value.Label ||
            value.label ||
            value.displayName ||
            value.name;

        return label ? [String(label)] : [];
    }

    let text = String(value);
    text = text.replace(/-?\d+;#/g, "");

    return text
        .split(/;#/)
        .map((part) => part.trim())
        .filter(Boolean)
        .map((part) => {
            const pipe = part.indexOf("|");
            let label = pipe >= 0 ? part.substring(0, pipe) : part;

            const colon = label.lastIndexOf(":");
            if (colon >= 0) label = label.substring(colon + 1);

            return label.trim();
        })
        .filter(Boolean);
}

function choiceLabelsFromField(value) {
    if (value == null || value === "") return [];

    if (Array.isArray(value)) {
        return value
            .map(String)
            .map((item) => item.trim())
            .filter(Boolean);
    }

    return String(value)
        .split(/;#|;/)
        .map((item) => item.trim())
        .filter(Boolean);
}

/* UI */

function initializeInterface() {
    document.getElementById("domain")
        ?.addEventListener("change", async () => {
            try {
                await updateFunctions();
            } catch (error) {
                setSaveStatus(getErrorMessage(error), "error");
            }
        });

    document.getElementById("additional-toggle")
        ?.addEventListener("click", toggleAdditionalMetadata);

    document.getElementById("save-metadata")
        ?.addEventListener("click", saveMetadata);

    for (const id of [
        "system-platform",
        "collection",
        "tags",
        "audience",
    ]) {
        const search = document.getElementById(`${id}-search`);
        const toggle = document.getElementById(`${id}-toggle`);

        search?.addEventListener("focus", () => {
            openPicker(id);
        });

        search?.addEventListener("click", (event) => {
            event.stopPropagation();
            openPicker(id);
        });

        search?.addEventListener("input", (event) => {
            renderTaxonomyPicker(id, event.target.value);
            openPicker(id);
        });

        if (id === "tags") {
            search?.addEventListener("keydown", (event) => {
                if (event.key === "Enter") {
                    const value = event.target.value.trim();

                    if (value) {
                        event.preventDefault();

                        const existing = getPickerSource("tags").find(
                            (item) =>
                                normalizeName(item.label) ===
                                normalizeName(value)
                        );

                        if (existing) {
                            const alreadySelected =
                                pickerSelections.tags.some(
                                    (item) =>
                                        String(item.id) ===
                                        String(existing.id)
                                );

                            if (!alreadySelected) {
                                pickerSelections.tags.push({
                                    id: existing.id,
                                    label: existing.label,
                                    term: existing.term,
                                    isNew: false,
                                });
                            }

                            event.target.value = "";
                            renderTaxonomyPicker("tags");
                        } else {
                            addPendingTag(value);
                        }
                    }
                }
            });
        }

        toggle?.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            togglePickerMenu(id);
        });
    }

    for (const id of ["owner", "responsible", "secondary"]) {
        const input = document.getElementById(`${id}-search`);

        input?.addEventListener("click", (event) => {
            event.stopPropagation();
        });

        input?.addEventListener("input", (event) => {
            closeAllPickers();
            closeAllPersonMenus(id);
            schedulePersonSearch(id, event.target.value);
        });
    }

    document.addEventListener("click", (event) => {
        if (!event.target.closest(".nerd-picker")) {
            closeAllPickers();
        }

        if (!event.target.closest(".nerd-person-picker")) {
            closeAllPersonMenus();
        }
    });
}

function toggleAdditionalMetadata() {
    const content = document.getElementById("additional-content");
    const toggle = document.getElementById("additional-toggle");
    const chevron = document.getElementById("additional-chevron");

    if (!content || !toggle || !chevron) return;

    const opening = content.classList.contains("nerd-hidden");
    content.classList.toggle("nerd-hidden", !opening);
    toggle.setAttribute("aria-expanded", opening ? "true" : "false");

    chevron.className = opening
        ? "ms-Icon ms-Icon--ChevronUp"
        : "ms-Icon ms-Icon--ChevronDown";
}

function enableSave() {
    const button = document.getElementById("save-metadata");
    if (button) button.disabled = false;
}

function setConnectedStatus() {
    setConnectionStatus(
        `Connected as ${currentUser?.displayName || "Unknown User"}` +
        `\nSystem: ${
            dorkSite?.displayName ||
            "DORK - Documentation Organization & Reference Key"
        }`
    );
}

function setConnectionStatus(message, isError = false) {
    const element = document.getElementById("connection-status");
    if (!element) return;

    element.textContent = message || "";
    element.classList.toggle("is-error", isError);
    element.classList.toggle(
        "is-connected",
        !isError && String(message || "").startsWith("Connected")
    );
}

function setSaveStatus(message, state = "") {
    const element = document.getElementById("save-status");
    if (!element) return;

    element.textContent = message || "";
    element.classList.remove("is-working", "is-success", "is-error");

    if (state === "working") element.classList.add("is-working");
    if (state === "success") element.classList.add("is-success");
    if (state === "error") element.classList.add("is-error");
}

/* HELPERS */

function normalizeName(value) {
    return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
}

function extractMembershipLogin(loginName) {
    const value = String(loginName || "").trim();
    if (!value) return "";

    const marker = "|membership|";
    const index = value.toLowerCase().lastIndexOf(marker);

    return index >= 0
        ? value.substring(index + marker.length)
        : value;
}

function normalizeSearchText(value) {
    return String(value || "").trim().toLowerCase();
}

function getErrorMessage(error) {
    return error?.message || String(error || "Unknown error");
}

/* BROWSER INTAKE: list and selection are read-only until Save metadata. */
let intakeFiles = [];
let intakeUploading = false;
const intakeUploadedIds = new Set();
let intakeDefaults = null;
async function initializeIntake() {
    document.getElementById("intake-browser").hidden = false;
    document.getElementById("intake-refresh").onclick = () => refreshIntakeFiles().catch(error => setSaveStatus(getErrorMessage(error), "error"));
    document.getElementById("intake-search").oninput = renderIntakeFiles;
    document.getElementById("intake-unassigned").onchange = renderIntakeFiles;
    document.getElementById("intake-upload-input").onchange = event => uploadIntakeFiles(event.target.files).catch(error => setSaveStatus(getErrorMessage(error), "error"));
    const drop = document.getElementById("intake-drop");
    drop.ondragover = event => { event.preventDefault(); drop.classList.add("drag-over"); };
    drop.ondragleave = () => drop.classList.remove("drag-over");
    drop.ondrop = event => { event.preventDefault(); drop.classList.remove("drag-over"); uploadIntakeFiles(event.dataTransfer.files).catch(error => setSaveStatus(getErrorMessage(error), "error")); };
    document.getElementById("metadata-form").hidden = false;
    setChoiceValue("lifecycle", "Draft");
    document.getElementById("save-metadata").disabled = true;
    await refreshIntakeFiles();
    setConnectedStatus();
}
async function refreshIntakeFiles() {
    const items = await graphGetAll(`${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}/lists/${encodeURIComponent(dorkDocumentsLibrary.id)}/items?$expand=fields&$top=200`);
    intakeFiles = items.filter(item => item.fields?.FileLeafRef && String(item.fields.FSObjType || "0") === "0");
    renderIntakeFiles();
}
function renderIntakeFiles() {
    const query = document.getElementById("intake-search").value.trim().toLowerCase();
    const unassigned = document.getElementById("intake-unassigned").checked;
    const idName = requireColumn("DORK ID").name;
    const list = document.getElementById("intake-files");
    list.replaceChildren();
    const matches = intakeFiles.filter(item => (!unassigned || !item.fields[idName]) && String(item.fields.FileLeafRef).toLowerCase().includes(query));
    document.getElementById("intake-count").textContent = `${matches.length} files`;
    for (const item of matches) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "intake-file";
        button.textContent = `${item.fields.FileLeafRef} — ${item.fields[idName] || "Unnumbered"}`;
        button.disabled = isSaving;
        button.onclick = () => selectIntakeFile(item).catch(error => setSaveStatus(getErrorMessage(error), "error"));
        list.appendChild(button);
    }
    if (!matches.length) list.textContent = "No matching files. Upload files into DORK Documents, then refresh this list.";
}
async function selectIntakeFile(item) {
    if (isSaving || connecting || intakeUploading) return;
    const latestDraft = !currentDriveItem ? captureIntakeDefaults() : null;
    connecting = true;
    currentDriveItem = null;
    currentListItem = null;
    document.getElementById("metadata-form").hidden = true;
    setSaveStatus("Loading file metadata...", "working");
    try {
        const driveItem = await graphGet(`${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}/lists/${encodeURIComponent(dorkDocumentsLibrary.id)}/items/${encodeURIComponent(item.id)}/driveItem?$select=id,name,webUrl,parentReference,sharepointIds,file,folder`);
        if (!driveItem.file || driveItem.folder) throw new Error("Select a file, not a folder.");
        currentDriveItem = driveItem;
        currentDocumentUrl = driveItem.webUrl;
        currentListItem = await getCurrentListItem(driveItem);
        currentDocumentFields = currentListItem.fields || {};
        await hydrateAllControls();
        if (intakeUploadedIds.has(driveItem.id) && !getFieldValue("DORK ID")) {
            const defaults = latestDraft || intakeDefaults;
            if (defaults) await applyIntakeDefaults(defaults, Boolean(latestDraft));
        }
        document.getElementById("intake-selected").textContent = driveItem.name;
        document.getElementById("metadata-form").hidden = false;
        enableSave();
        setSaveStatus("File loaded. Close it in Word or Excel before saving here.", "");
    } finally { connecting = false; }
}

async function assertIntakeFileClosed() {
    const endpoint = `/_api/web/lists(guid'${dorkDocumentsLibrary.id}')/items(${currentListItem.id})/File?$select=CheckOutType,LockedByUser/Id&$expand=LockedByUser`;
    const response = await sharePointRequest(endpoint, "GET");
    const file = response?.d || response;
    if (!file || file.CheckOutType === undefined) throw new Error("Could not verify whether the file is open. No new number was assigned.");
    if (file.LockedByUser?.Id || Number(file.CheckOutType) !== 2) throw new Error("Close the file in Office and check it in before saving through GEEK. No new number was assigned.");
}

function captureIntakeDefaults() {
    const values = Object.fromEntries(["domain", "function", "document-type", "classification", "lifecycle", "information-source", "last-reviewed", "next-review"].map(id => [id, document.getElementById(id).value]));
    return { title: document.getElementById("title").value.trim(), values, pickers: Object.fromEntries(Object.entries(pickerSelections).map(([key, value]) => [key, value.map(item => ({...item}))])), people: Object.fromEntries(Object.entries(personSelections).map(([key, value]) => [key, value ? {...value} : null])) };
}
async function applyIntakeDefaults(defaults = intakeDefaults, preserveTitle = false) {
    if (preserveTitle && defaults.title) document.getElementById("title").value = defaults.title;
    document.getElementById("domain").value = defaults.values.domain;
    await updateFunctions();
    for (const [id, value] of Object.entries(defaults.values)) document.getElementById(id).value = value;
    for (const [id, value] of Object.entries(defaults.pickers)) { pickerSelections[id] = value.map(item => ({...item})); renderPickerChips(id); }
    for (const [id, value] of Object.entries(defaults.people)) { personSelections[id] = value ? {...value} : null; renderSelectedPerson(id); }
}
async function uploadIntakeFiles(files) {
    if (intakeUploading || isSaving || connecting) return;
    const queue = Array.from(files || []);
    if (!queue.length) return;
    intakeDefaults = captureIntakeDefaults();
    intakeUploading = true;
    document.getElementById("save-metadata").disabled = true;
    document.getElementById("intake-upload-input").disabled = true;
    let singleUploadedItem = null;
    const status = document.getElementById("intake-upload-status");
    status.replaceChildren();
    try {
        for (const file of queue) {
            const row = document.createElement("p");
            row.textContent = `${file.name}: uploading...`;
            status.appendChild(row);
            try {
                const uploaded = await uploadIntakeFile(file, {
                    createSession: (name, body) => graphRequest(`${GRAPH_ROOT}/sites/${encodeURIComponent(dorkSite.id)}/drive/root:/${encodeURIComponent(name)}:/createUploadSession`, {method:"POST", body:JSON.stringify(body)}),
                    // The upload URL authenticates the session; do not send bearer tokens to it.
                    putChunk: (url, body, headers) => fetch(url, {method:"PUT", body, headers}),
                });
                intakeUploadedIds.add(uploaded.id);
                if (queue.length === 1) singleUploadedItem = uploaded;
                row.textContent = `${file.name}: uploaded, unnumbered. Select it below to review and save metadata.`;
            } catch (error) { row.textContent = `${file.name}: ${getErrorMessage(error)}`; }
        }
        await refreshIntakeFiles();
    } catch (error) { setSaveStatus(`Upload list refresh failed: ${getErrorMessage(error)}. Refresh files before retrying.`, "error"); }
    finally {
        intakeUploading = false;
        document.getElementById("intake-upload-input").disabled = false;
        document.getElementById("save-metadata").disabled = !currentDriveItem;
        document.getElementById("intake-upload-input").value = "";
    }
    if (singleUploadedItem) {
        const driveItem = await getDriveItemById(singleUploadedItem.id);
        if (driveItem.sharepointIds?.listItemId) await selectIntakeFile({id: driveItem.sharepointIds.listItemId});
    }
}
