const express = require("express");
const cors = require("cors");

const app = express();
const PORT = 3001;

const ALLOWED_ORIGINS = new Set([
  "https://localhost:3000",
]);

app.use(
  cors({
    origin(origin, callback) {
      // Allows requests with no Origin header, such as a local health check.
      if (!origin || ALLOWED_ORIGINS.has(origin)) {
        callback(null, true);
        return;
      }

      callback(new Error(`Origin not allowed by NERD API: ${origin}`));
    },
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

app.use(express.json({ limit: "1mb" }));

/*
 * --------------------------------------------------------------------------
 * Health Check
 * --------------------------------------------------------------------------
 */

app.get("/api/health", (request, response) => {
  response.json({
    status: "ok",
    service: "NERD API",
  });
});

/*
 * --------------------------------------------------------------------------
 * SharePoint REST Proxy
 * --------------------------------------------------------------------------
 *
 * NERD obtains the user's delegated SharePoint access token in the task pane.
 * The token is sent to this local bridge in the Authorization header.
 *
 * The bridge does NOT store credentials or tokens. It simply forwards the
 * authenticated request to the DORK SharePoint site.
 */

app.post("/api/sharepoint", async (request, response) => {
  try {
    const authorization = request.headers.authorization;

    if (!authorization || !authorization.startsWith("Bearer ")) {
      response.status(401).json({
        error: "Missing SharePoint bearer token.",
      });
      return;
    }

    const {
      url,
      method = "POST",
      headers = {},
      body = null,
    } = request.body || {};

    if (!url || typeof url !== "string") {
      response.status(400).json({
        error: "A SharePoint URL is required.",
      });
      return;
    }

    /*
     * SECURITY:
     * Do not allow this proxy to forward the user's SharePoint token
     * anywhere except the RTSD SharePoint tenant.
     */
    let targetUrl;

    try {
      targetUrl = new URL(url);
    } catch {
      response.status(400).json({
        error: "Invalid SharePoint URL.",
      });
      return;
    }

    if (
      targetUrl.protocol !== "https:" ||
      targetUrl.hostname !== "rocktwpnet.sharepoint.com"
    ) {
      response.status(403).json({
        error: "NERD may only proxy requests to rocktwpnet.sharepoint.com.",
      });
      return;
    }

    const forwardedHeaders = {
      Accept: "application/json;odata=nometadata",
      Authorization: authorization,
      ...headers,
    };

    /*
     * Never allow the client-supplied header object to replace the
     * Authorization token received by this API.
     */
    forwardedHeaders.Authorization = authorization;

    if (body !== null && body !== undefined) {
      forwardedHeaders["Content-Type"] =
        forwardedHeaders["Content-Type"] ||
        "application/json;odata=nometadata";
    }

    const sharePointResponse = await fetch(targetUrl.toString(), {
      method,
      headers: forwardedHeaders,
      body:
        body === null || body === undefined
          ? undefined
          : typeof body === "string"
            ? body
            : JSON.stringify(body),
    });

    const responseText = await sharePointResponse.text();

    let responseBody = null;

    if (responseText) {
      try {
        responseBody = JSON.parse(responseText);
      } catch {
        responseBody = responseText;
      }
    }

    if (!sharePointResponse.ok) {
      response.status(sharePointResponse.status).json({
        error: "SharePoint request failed.",
        status: sharePointResponse.status,
        statusText: sharePointResponse.statusText,
        details: responseBody,
      });
      return;
    }

    response.status(sharePointResponse.status).json({
      ok: true,
      status: sharePointResponse.status,
      data: responseBody,
    });
  } catch (error) {
    console.error("NERD SharePoint proxy error:", error);

    response.status(500).json({
      error: "NERD API encountered an unexpected error.",
      details: error instanceof Error ? error.message : String(error),
    });
  }
});

/*
 * --------------------------------------------------------------------------
 * Error Handler
 * --------------------------------------------------------------------------
 */

app.use((error, request, response, next) => {
  console.error("NERD API error:", error);

  response.status(500).json({
    error: error instanceof Error ? error.message : "Unknown NERD API error.",
  });
});

/*
 * --------------------------------------------------------------------------
 * Start NERD API
 * --------------------------------------------------------------------------
 */

app.listen(PORT, () => {
  console.log(`NERD API is running on http://localhost:${PORT}`);
});