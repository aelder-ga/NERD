import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";
// This page only returns the MSAL v5 response; it never starts NERD or a dialog.
broadcastResponseToMainFrame().catch(() => {
    document.getElementById("status").textContent = "Return to Office to continue signing in.";
});
