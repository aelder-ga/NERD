// Validate destinations before obtaining or attaching a delegated bearer token.
function graphUrl(value) {
    let url;
    try { url = new URL(value); } catch { throw new Error('Invalid Microsoft Graph URL.'); }
    if (url.origin !== 'https://graph.microsoft.com' || url.username || url.password || url.hash ||
        !url.pathname.startsWith('/v1.0/') || /%2f|%5c|%25/i.test(url.pathname)) {
        throw new Error('NERD may only send Graph tokens to Microsoft Graph v1.0.');
    }
    return url.href;
}
function uploadUrl(value) {
    let url;
    try { url = new URL(value); } catch { throw new Error('Invalid Microsoft upload session.'); }
    const microsoftUpload = url.hostname.endsWith('.sharepoint.com') || url.hostname.endsWith('.up.1drv.com');
    if (url.protocol !== 'https:' || url.port || url.username || url.password || url.hash || !microsoftUpload) {
        throw new Error('Microsoft did not return a trusted upload session.');
    }
    return url.href;
}
module.exports = { graphUrl, uploadUrl };
