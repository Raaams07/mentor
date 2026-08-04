/*
 * mentor-suggestions-proxy-test.js
 * -------------------------------------
 * Integration test for mentor-suggestions-proxy.js — starts the REAL
 * server (not a mocked function call) and hits it with REAL HTTP
 * requests, so this exercises the actual network path: CORS handling,
 * JSON body parsing, validation, error responses, and the mock-mode
 * screening pipeline end-to-end.
 *
 * Runs entirely in mock mode — no ANTHROPIC_API_KEY is read or required
 * (MENTOR_MOCK_LLM=1 is forced below regardless of the environment), so
 * this makes zero real network calls to Anthropic and costs nothing.
 *
 * Run with: node server/mentor-suggestions-proxy-test.js
 */

process.env.MENTOR_MOCK_LLM = "1"; // force mock mode regardless of any real key present in the environment
process.env.MENTOR_PROXY_PORT = "3091"; // distinct port so this never collides with a real proxy instance already running

const http = require("http");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

function request(options, body) {
  return new Promise((resolve, reject) => {
    const req = http.request(options, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        let parsed = null;
        try {
          parsed = data ? JSON.parse(data) : null;
        } catch (e) {
          parsed = null;
        }
        resolve({ statusCode: res.statusCode, headers: res.headers, body: parsed, raw: data });
      });
    });
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

async function run() {
  const { server, PORT } = require("./mentor-suggestions-proxy.js");

  await new Promise((resolve) => {
    if (server.listening) resolve();
    else server.once("listening", resolve);
  });

  try {
    console.log("-- Health check --\n");
    const health = await request({ hostname: "localhost", port: PORT, path: "/health", method: "GET" });
    assert(health.statusCode === 200 && health.body.mode === "mock", "GET /health reports mock mode");

    console.log("\n-- CORS preflight (OPTIONS) --\n");
    const preflight = await request({
      hostname: "localhost",
      port: PORT,
      path: "/api/ineligible-itc-vendor-suggestions",
      method: "OPTIONS",
      headers: { origin: "https://localhost:3000" },
    });
    assert(preflight.statusCode === 204, "OPTIONS preflight returns 204");
    assert(preflight.headers["access-control-allow-origin"] === "https://localhost:3000", "preflight echoes back the localhost task-pane origin");

    console.log("\n-- Real screening request (mock mode) --\n");
    const vendors = [
      { gstin: "27ZZAAA1111A1Z1", vendorName: "Speedway Motors Pvt Ltd" },
      { gstin: "07ZZBBB2222B1Z2", vendorName: "Bharat General Insurance Co Ltd" },
      { gstin: "29ZZCCC3333C1Z3", vendorName: "Sunrise Retail Traders" },
    ];
    const payload = JSON.stringify({ vendors });
    const screening = await request(
      {
        hostname: "localhost",
        port: PORT,
        path: "/api/ineligible-itc-vendor-suggestions",
        method: "POST",
        headers: { "content-type": "application/json", "content-length": Buffer.byteLength(payload), origin: "https://localhost:3000" },
      },
      payload
    );
    assert(screening.statusCode === 200, "screening request returns 200");
    assert(screening.headers["access-control-allow-origin"] === "https://localhost:3000", "screening response has the CORS header");
    assert(screening.body.mock === true, "response is clearly tagged mock: true");
    assert(screening.body.results.length === 3, "all 3 vendors came back");

    const motors = screening.body.results.find((r) => r.gstin === "27ZZAAA1111A1Z1");
    assert(motors.flagged === true && motors.subClause === "17(5)(a)/(ab)", "vehicle dealer flagged correctly by the mock rules, over real HTTP");

    const retailer = screening.body.results.find((r) => r.gstin === "29ZZCCC3333C1Z3");
    assert(retailer.flagged === false, "ambiguous retailer not flagged, over real HTTP");

    console.log("\n-- Malformed requests are rejected explicitly, not silently --\n");

    const badJson = await request(
      { hostname: "localhost", port: PORT, path: "/api/ineligible-itc-vendor-suggestions", method: "POST", headers: { "content-type": "application/json" } },
      "{not valid json"
    );
    assert(badJson.statusCode === 400, "invalid JSON body -> 400");

    const emptyVendors = JSON.stringify({ vendors: [] });
    const emptyResp = await request(
      { hostname: "localhost", port: PORT, path: "/api/ineligible-itc-vendor-suggestions", method: "POST", headers: { "content-type": "application/json", "content-length": Buffer.byteLength(emptyVendors) } },
      emptyVendors
    );
    assert(emptyResp.statusCode === 400, "empty vendors array -> 400, not a silent empty success");

    const missingName = JSON.stringify({ vendors: [{ gstin: "27ZZAAA1111A1Z1" }] });
    const missingNameResp = await request(
      { hostname: "localhost", port: PORT, path: "/api/ineligible-itc-vendor-suggestions", method: "POST", headers: { "content-type": "application/json", "content-length": Buffer.byteLength(missingName) } },
      missingName
    );
    assert(missingNameResp.statusCode === 400, "a vendor missing vendorName -> 400");

    const unknownRoute = await request({ hostname: "localhost", port: PORT, path: "/nope", method: "GET" });
    assert(unknownRoute.statusCode === 404, "unknown route -> 404");
  } finally {
    server.close();
  }

  console.log("");
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All proxy server integration checks passed (mock mode, real HTTP, zero API calls).");
  }
}

run();
