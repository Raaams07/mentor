/* global fetch */

/*
 * mentor-proxy-column-pattern-store.js
 * ----------------------------------------
 * Client-side access to Tier 1 (shared software-pattern) column-resolution
 * memory — talks to the local proxy server (server/mentor-suggestions-
 * proxy.js's new /api/column-pattern-memory/* routes, server/column-
 * pattern-store.js underneath), the same local process already used for
 * the Ineligible ITC vendor-name suggestion feature (see
 * mentorGstFetchIneligibleItcVendorSuggestions in
 * mentor-gst-reconciliation-ui.js — same http://localhost:3001 convention,
 * same MENTOR_PROXY_PORT caveat if that ever changes server-side).
 *
 * Tier 1 is a nice-to-have, not a dependency: every function here MUST
 * NEVER THROW. A network error, the proxy simply not running, a non-2xx
 * response, or a malformed body all resolve to the same "not available"
 * outcome — the caller (mentor-column-memory-ui.js) falls through to
 * asking the user exactly as if Tier 1 didn't exist at all. Nothing about
 * Tier 2 (per-client, workbook-embedded) resolution, prompting, or saving
 * a user's answer may ever be blocked or failed by a Tier-1 outage.
 */

const MENTOR_COLUMN_PATTERN_LOOKUP_URL = "http://localhost:3001/api/column-pattern-memory/lookup";
const MENTOR_COLUMN_PATTERN_REMEMBER_URL = "http://localhost:3001/api/column-pattern-memory/remember";

// -> { found: false } | { found: true, chosenHeader }
async function lookupColumnPattern(fieldName, candidateHeaders) {
  try {
    const response = await fetch(MENTOR_COLUMN_PATTERN_LOOKUP_URL, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ fieldName, candidateHeaders }),
    });
    if (!response.ok) return { found: false };
    const body = await response.json();
    return body && typeof body.found === "boolean" ? body : { found: false };
  } catch (error) {
    console.log("MENTOR Tier-1 column-pattern lookup unavailable (proxy not running?) — falling through to Tier 2/ask:", error.message);
    return { found: false };
  }
}

// Fire-and-forget / best-effort — callers should not (and don't need to)
// await this before continuing their own save flow.
async function rememberColumnPattern(fieldName, candidateHeaders, chosenHeader) {
  try {
    const response = await fetch(MENTOR_COLUMN_PATTERN_REMEMBER_URL, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ fieldName, candidateHeaders, chosenHeader }),
    });
    if (!response.ok) console.log("MENTOR Tier-1 column-pattern remember failed with status", response.status);
  } catch (error) {
    console.log("MENTOR Tier-1 column-pattern remember unavailable (proxy not running?) — Tier-2 save still succeeded:", error.message);
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { lookupColumnPattern, rememberColumnPattern };
}
