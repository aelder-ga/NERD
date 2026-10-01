/* global Office */
const tenant = "1169a3a9-3860-4a0b-80aa-410be007cde5";
const clientId = "f8371eb2-8758-48cf-8fa5-2f9b696d89c4";
const redirectUri = window.location.origin + "/auth.html";
const key = "nerd-dialog-pkce";
const resourceScopes = {
    graph: "https://graph.microsoft.com/User.Read https://graph.microsoft.com/User.ReadBasic.All https://graph.microsoft.com/Sites.ReadWrite.All https://graph.microsoft.com/TermStore.ReadWrite.All",
    sharepoint: "https://rocktwpnet.sharepoint.com/AllSites.Write",
};
const base64url = bytes => btoa(String.fromCharCode(...bytes)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");

async function authorize(saved, resource, includeBothResources) {
    const verifier = base64url(crypto.getRandomValues(new Uint8Array(32)));
    const challenge = base64url(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))));
    Object.assign(saved, {resource, verifier, state: crypto.randomUUID()});
    sessionStorage.setItem(key, JSON.stringify(saved));
    const scopes = includeBothResources
        ? `${resourceScopes[resource]} ${resourceScopes[resource === "graph" ? "sharepoint" : "graph"]} offline_access`
        : `${resourceScopes[resource]} offline_access`;
    const url = new URL(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize`);
    url.search = new URLSearchParams({
        client_id: clientId, response_type: "code", redirect_uri: redirectUri,
        response_mode: "query", scope: scopes, state: saved.state,
        code_challenge: challenge, code_challenge_method: "S256",
    }).toString();
    if (saved.loginHint) url.searchParams.set("login_hint", saved.loginHint);
    if (typeof saved.claims === "string") url.searchParams.set("claims", saved.claims);
    window.location.replace(url.href);
}

async function requestToken(parameters) {
    const response = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
        method: "POST", headers: {"Content-Type": "application/x-www-form-urlencoded"},
        body: new URLSearchParams({client_id: clientId, redirect_uri: redirectUri, ...parameters}),
    });
    const data = await response.json();
    if (!response.ok || !data.access_token) {
        const error = new Error(data.error_description || "Microsoft did not return an access token.");
        error.code = data.error;
        error.claims = data.claims;
        throw error;
    }
    return data;
}

function accessToken(data) {
    const lifetime = Number(data.expires_in);
    if (typeof data.access_token !== "string" || !data.access_token || !Number.isFinite(lifetime) || lifetime <= 0) {
        throw new Error("Microsoft returned an invalid token response.");
    }
    return {token: data.access_token, expires: Date.now() + lifetime * 1000};
}

Office.onReady(async () => {
    let saved;
    try {
        const params = new URLSearchParams(window.location.search);
        if (params.has("code") || params.has("error")) {
            saved = JSON.parse(sessionStorage.getItem(key) || "null");
            if (!saved || !saved.nonce || !Object.hasOwn(resourceScopes, saved.resource) ||
                saved.state !== params.get("state") || !Number.isFinite(saved.started) ||
                Date.now() - saved.started > 600000 || saved.started > Date.now()) {
                // Do not send a response using unvalidated callback state.
                saved = null;
                throw new Error("Sign-in state validation failed. Close this window and retry.");
            }
            if (typeof history.replaceState === "function") history.replaceState(null, "", "/auth.html");
            if (params.has("error")) throw new Error(params.get("error_description") || params.get("error"));
            const data = await requestToken({
                grant_type: "authorization_code", code: params.get("code"),
                code_verifier: saved.verifier, scope: `${resourceScopes[saved.resource]} offline_access`,
            });
            saved.tokens[saved.resource] = accessToken(data);
            const other = saved.resource === "graph" ? "sharepoint" : "graph";
            if (!saved.tokens[other]) {
                document.getElementById("status").textContent = "Connecting to DORK with your Microsoft session…";
                if (data.refresh_token) {
                    try {
                        const second = await requestToken({
                            grant_type: "refresh_token", refresh_token: data.refresh_token,
                            scope: resourceScopes[other],
                        });
                        saved.tokens[other] = accessToken(second);
                    } catch (error) {
                        if (!["interaction_required", "consent_required", "invalid_grant"].includes(error.code)) throw error;
                        // Additional Microsoft consent or MFA uses this same dialog.
                        saved.claims = error.claims;
                    }
                }
                if (!saved.tokens[other]) {
                    await authorize(saved, other, false);
                    return;
                }
            }
            sessionStorage.removeItem(key);
            document.getElementById("status").textContent = "Signed in. Returning to Office…";
            Office.context.ui.messageParent(JSON.stringify({nonce: saved.nonce, tokens: saved.tokens}), {targetOrigin: window.location.origin});
            return;
        }
        const resource = params.get("resource");
        const nonce = params.get("nonce");
        if (!nonce || !Object.hasOwn(resourceScopes, resource)) throw new Error("Open sign-in from the NERD pane in Word or Excel.");
        saved = {nonce, tokens: {}, started: Date.now(), loginHint: params.get("login_hint") || ""};
        await authorize(saved, resource, true);
    } catch (error) {
        sessionStorage.removeItem(key);
        document.getElementById("status").textContent = error.message;
        if (saved?.nonce) Office.context.ui.messageParent(JSON.stringify({nonce: saved.nonce, error: error.message}), {targetOrigin: window.location.origin});
    }
});
