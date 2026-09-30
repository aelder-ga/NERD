/* global Office */
const tenant = "1169a3a9-3860-4a0b-80aa-410be007cde5";
const clientId = "f8371eb2-8758-48cf-8fa5-2f9b696d89c4";
const redirectUri = window.location.origin + "/auth.html";
const key = "nerd-dialog-pkce";
const base64url = bytes => btoa(String.fromCharCode(...bytes)).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/, "");
Office.onReady(async () => {
    let saved;
    try {
        const params = new URLSearchParams(window.location.search);
        if (params.has("code") || params.has("error")) {
            saved = JSON.parse(sessionStorage.getItem(key) || "null");
            if (!saved || saved.state !== params.get("state") || Date.now() - saved.started > 600000) throw new Error("Sign-in state validation failed. Close this window and retry.");
            if (typeof history.replaceState === "function") history.replaceState(null, "", "/auth.html");
            if (params.has("error")) throw new Error(params.get("error_description") || params.get("error"));
            const response = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
                method:"POST", headers:{"Content-Type":"application/x-www-form-urlencoded"},
                body:new URLSearchParams({client_id:clientId,grant_type:"authorization_code",code:params.get("code"),redirect_uri:redirectUri,code_verifier:saved.verifier,scope:saved.scopes})
            });
            const data = await response.json();
            if (!response.ok || !data.access_token) throw new Error(data.error_description || "Microsoft did not return an access token.");
            sessionStorage.removeItem(key);
            document.getElementById("status").textContent = "Signed in. Returning to Word…";
            Office.context.ui.messageParent(JSON.stringify({nonce:saved.nonce,token:data.access_token,expires:Date.now()+Number(data.expires_in)*1000}), {targetOrigin:window.location.origin});
            return;
        }
        const resource = params.get("resource");
        const nonce = params.get("nonce");
        if (!nonce || !["graph","sharepoint"].includes(resource)) throw new Error("Open sign-in from the NERD pane in Word.");
        const scopes = resource === "graph" ? "https://graph.microsoft.com/User.Read https://graph.microsoft.com/User.ReadBasic.All https://graph.microsoft.com/Sites.ReadWrite.All https://graph.microsoft.com/TermStore.ReadWrite.All" : "https://rocktwpnet.sharepoint.com/AllSites.Write";
        const verifier = base64url(crypto.getRandomValues(new Uint8Array(32)));
        const challenge = base64url(new Uint8Array(await crypto.subtle.digest("SHA-256",new TextEncoder().encode(verifier))));
        saved = {nonce,scopes,verifier,state:crypto.randomUUID(),started:Date.now()};
        sessionStorage.setItem(key,JSON.stringify(saved));
        const url = new URL(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize`);
        url.search = new URLSearchParams({client_id:clientId,response_type:"code",redirect_uri:redirectUri,response_mode:"query",scope:scopes,state:saved.state,code_challenge:challenge,code_challenge_method:"S256"}).toString();
        window.location.replace(url.href);
    } catch (error) {
        sessionStorage.removeItem(key);
        document.getElementById("status").textContent = error.message;
        if (saved?.nonce) Office.context.ui.messageParent(JSON.stringify({nonce:saved.nonce,error:error.message}),{targetOrigin:window.location.origin});
    }
});
