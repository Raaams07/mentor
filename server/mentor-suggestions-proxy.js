/*
 * mentor-suggestions-proxy.js
 * -------------------------------
 * Small local HTTP server that holds the Anthropic API key server-side,
 * so the task pane (client-side-only JS, no backend of its own) never
 * embeds it in the shipped bundle. Exposes one endpoint the task pane
 * calls for the Ineligible ITC vendor-name suggestion feature (see
 * mentor-gst-reconciliation-ui.js) — nothing else runs through this.
 *
 * MOCK MODE: if ANTHROPIC_API_KEY isn't set (or MENTOR_MOCK_LLM=1 is set
 * explicitly), this runs in mock mode — returns a locally-generated,
 * CLEARLY FAKE response (mock-vendor-itc-screening.js) instead of calling
 * Anthropic at all. This exists so the whole feature (client -> proxy ->
 * structured response -> sheet write) can be built and tested end-to-end
 * with zero API cost and zero real network calls, before a real key is
 * ever supplied. The server logs loudly on startup and on every request
 * which mode it's in, and every mock response is tagged mock: true.
 *
 * Run: node server/mentor-suggestions-proxy.js
 * (in a separate terminal from the webpack dev server — see package.json's
 * "gst-suggestions-proxy" script)
 *
 * To go live later: set ANTHROPIC_API_KEY (env var, or a gitignored
 * server/.anthropic-api-key file containing just the key) and restart —
 * no code change needed.
 */

const http = require("http");
const fs = require("fs");
const path = require("path");
const { screenVendorsForIneligibleItc, callAnthropicApi } = require("./anthropic-vendor-itc-screening.js");
const { mockScreenVendors } = require("./mock-vendor-itc-screening.js");

const PORT = process.env.MENTOR_PROXY_PORT ? parseInt(process.env.MENTOR_PROXY_PORT, 10) : 3001;
const FORCE_MOCK = process.env.MENTOR_MOCK_LLM === "1";

function readLocalApiKeyFile() {
  const keyPath = path.join(__dirname, ".anthropic-api-key");
  try {
    return fs.readFileSync(keyPath, "utf8").trim();
  } catch (readError) {
    return "";
  }
}

const API_KEY = process.env.ANTHROPIC_API_KEY || readLocalApiKeyFile();
const MOCK_MODE = FORCE_MOCK || !API_KEY;

function isAllowedOrigin(origin) {
  return typeof origin === "string" && /^https?:\/\/localhost(:\d+)?$/.test(origin);
}

function sendJson(res, statusCode, body, origin) {
  const headers = { "content-type": "application/json" };
  if (isAllowedOrigin(origin)) headers["access-control-allow-origin"] = origin;
  res.writeHead(statusCode, headers);
  res.end(JSON.stringify(body));
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    let data = "";
    req.on("data", (chunk) => {
      data += chunk;
      if (data.length > 1_000_000) req.destroy(new Error("Request body too large"));
    });
    req.on("end", () => resolve(data));
    req.on("error", reject);
  });
}

function validateVendors(vendors) {
  if (!Array.isArray(vendors) || vendors.length === 0) return "`vendors` must be a non-empty array";
  for (const v of vendors) {
    if (!v || typeof v.gstin !== "string" || !v.gstin.trim() || typeof v.vendorName !== "string" || !v.vendorName.trim()) {
      return "every vendor needs a non-empty gstin and vendorName";
    }
  }
  return null;
}

const server = http.createServer(async (req, res) => {
  const origin = req.headers.origin;

  if (req.method === "OPTIONS") {
    const headers = { "access-control-allow-methods": "POST, OPTIONS", "access-control-allow-headers": "Content-Type" };
    if (isAllowedOrigin(origin)) headers["access-control-allow-origin"] = origin;
    res.writeHead(204, headers);
    res.end();
    return;
  }

  if (req.method === "GET" && req.url === "/health") {
    sendJson(res, 200, { status: "ok", mode: MOCK_MODE ? "mock" : "live" }, origin);
    return;
  }

  if (req.method !== "POST" || req.url !== "/api/ineligible-itc-vendor-suggestions") {
    sendJson(res, 404, { error: "not found" }, origin);
    return;
  }

  try {
    const rawBody = await readRequestBody(req);
    let parsedBody;
    try {
      parsedBody = JSON.parse(rawBody);
    } catch (parseError) {
      sendJson(res, 400, { error: "request body is not valid JSON" }, origin);
      return;
    }

    const validationError = validateVendors(parsedBody.vendors);
    if (validationError) {
      sendJson(res, 400, { error: validationError }, origin);
      return;
    }

    console.log("MENTOR suggestions proxy: screening " + parsedBody.vendors.length + " vendor(s) — mode: " + (MOCK_MODE ? "MOCK (no real LLM call)" : "LIVE (calling Anthropic)"));

    const result = MOCK_MODE ? mockScreenVendors(parsedBody.vendors) : await screenVendorsForIneligibleItc(parsedBody.vendors, (requestBody) => callAnthropicApi(requestBody, API_KEY));

    sendJson(res, 200, result, origin);
  } catch (error) {
    console.error("MENTOR suggestions proxy error:", error);
    sendJson(res, 500, { error: error.message }, origin);
  }
});

server.listen(PORT, () => {
  console.log("MENTOR suggestions proxy listening on http://localhost:" + PORT);
  console.log(MOCK_MODE ? "Running in MOCK MODE — no ANTHROPIC_API_KEY found (or MENTOR_MOCK_LLM=1 set). No real LLM calls, no cost." : "Running in LIVE MODE — calls to Anthropic's API WILL be made and billed.");
});

module.exports = { server, MOCK_MODE, PORT };
