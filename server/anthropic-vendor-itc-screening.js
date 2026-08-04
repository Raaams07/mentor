/*
 * anthropic-vendor-itc-screening.js
 * -------------------------------------
 * Builds the batched vendor-name ITC-screening prompt, calls Anthropic's
 * Messages API with forced tool-use for structured output, and validates
 * the response.
 *
 * screenVendorsForIneligibleItc() takes the actual API call as an injected
 * function (callAnthropic) rather than calling https itself, specifically
 * so the prompt-building and response-parsing logic here is unit-testable
 * without a live network call or an API key — see
 * vendor-itc-screening-test.js, which passes in a fake callAnthropic.
 * callAnthropicApi() (the real network call) lives in this same file but
 * is never exercised by that test.
 *
 * This is a NEW, separate, optional capability — it never reads or writes
 * anything in src/gst-reconciliation/*.js, and its output is explicitly
 * labeled "unconfirmed" everywhere it's surfaced (see gst-report-writer.js's
 * appendIneligibleItcVendorSuggestions). Vendor identity alone never
 * determines ITC eligibility — the categories below are the same ones
 * ineligible-itc-config.js already validates against, applied here as a
 * screening aid for manual review, not a determination.
 */

const https = require("https");

const ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages";
const ANTHROPIC_MODEL = "claude-sonnet-5";
const ANTHROPIC_API_VERSION = "2023-06-01";

const SCREENING_TOOL = {
  name: "report_vendor_itc_screening",
  description: "Reports Section 17(5) CGST Act ITC-blocking screening results for a list of vendors, based on their likely primary business inferred from their name — not actual invoice content.",
  input_schema: {
    type: "object",
    properties: {
      results: {
        type: "array",
        items: {
          type: "object",
          properties: {
            gstin: { type: "string" },
            vendorName: { type: "string" },
            flagged: { type: "boolean", description: "true only if this vendor's likely business genuinely and directly matches a known Section 17(5) blocked category" },
            category: { type: ["string", "null"], description: "Short category label (e.g. 'Motor Vehicles', 'Food/Beverage/Catering'), null when not flagged" },
            subClause: { type: ["string", "null"], description: "e.g. '17(5)(a)/(ab)', null when not flagged" },
            reasoning: { type: "string", description: "1-2 sentences. If eligibility genuinely depends on what was purchased (e.g. an insurer selling both motor and property cover, a hotel providing both accommodation and catering), say so explicitly rather than flagging or clearing unconditionally." },
          },
          required: ["gstin", "vendorName", "flagged", "reasoning"],
        },
      },
    },
    required: ["results"],
  },
};

function buildScreeningPrompt(vendors) {
  const vendorLines = vendors.map((v) => "- GSTIN: " + v.gstin + ", Name: " + v.vendorName).join("\n");
  return (
    "You are screening a list of vendor names from an Indian company's GST purchase records, to identify which vendors' TYPICAL LINE OF BUSINESS commonly triggers a blocked/restricted Input Tax Credit (ITC) category under Section 17(5) of the CGST Act.\n\n" +
    "You have ONLY the vendor's GSTIN and name — no invoice amounts, no description of what was actually purchased, no other data. Decide based solely on what this vendor's name suggests about their likely primary business.\n\n" +
    "Categories to consider:\n" +
    "- 17(5)(a)/(ab): motor vehicles for transport of persons (≤13 seats) and related services — blocked unless used for further supply (resale), passenger transport as a business, driving training, or goods transport.\n" +
    "- 17(5)(b): food, beverages, outdoor catering; and separately, employee health/life insurance, rent-a-cab, leave travel benefits — blocked unless legally obligatory on the employer or the goods/services are further supplied in the same category.\n" +
    "- 17(5)(c)/(d): works contract services or goods for construction of an immovable property (excluding plant & machinery).\n" +
    "- 17(5)(fa): CSR (Corporate Social Responsibility) expenditure.\n" +
    "- 17(5)(g): goods/services for personal use or consumption.\n\n" +
    "Rules:\n" +
    "1. Be conservative. Only flag a vendor if their likely business GENUINELY and DIRECTLY matches one of these categories.\n" +
    "2. If a vendor's likely business doesn't clearly indicate a blocked-category purchase (e.g. a general retailer, a raw-material supplier, a generic B2B services vendor), do NOT flag it — explicitly decline rather than guess. Still return a `results` entry for it with flagged=false and a brief reason.\n" +
    "3. Some vendors sell a MIX of blocked and non-blocked goods/services (e.g. a general insurer selling both motor and property insurance; a hotel providing both room accommodation and banquet/catering). For these, hedge explicitly in your reasoning — state that eligibility depends on what was actually purchased or which policy/service was booked. Set flagged=true if there's a real, non-trivial likelihood of a blocked-category purchase from this vendor, so a human reviews it, but the hedge must be clear in the reasoning text, not an unconditional flag.\n" +
    "4. Every flagged vendor must cite exactly one specific sub-clause and give reasoning grounded in what that vendor plausibly sells — not a generic disclaimer.\n" +
    "5. Return one entry per vendor in the input list, covering every vendor — flagged and not flagged alike.\n\n" +
    "Vendors:\n" +
    vendorLines
  );
}

function buildAnthropicRequestBody(vendors) {
  return {
    model: ANTHROPIC_MODEL,
    max_tokens: 4096,
    tools: [SCREENING_TOOL],
    tool_choice: { type: "tool", name: SCREENING_TOOL.name },
    messages: [{ role: "user", content: buildScreeningPrompt(vendors) }],
  };
}

// apiResponse: the parsed JSON body Anthropic's Messages API returned.
// Throws a descriptive Error if the shape isn't what's expected, rather
// than silently returning something malformed to the caller.
function parseScreeningResponse(apiResponse, expectedVendors) {
  if (!apiResponse || !Array.isArray(apiResponse.content)) {
    throw new Error("Unexpected Anthropic API response shape (no content array)");
  }
  const toolUseBlock = apiResponse.content.find((block) => block.type === "tool_use" && block.name === SCREENING_TOOL.name);
  if (!toolUseBlock) {
    throw new Error("Anthropic response did not include the expected tool_use block — the model may have replied in prose instead of calling the tool");
  }
  const results = toolUseBlock.input && toolUseBlock.input.results;
  if (!Array.isArray(results)) {
    throw new Error("Tool call's `results` field is missing or not an array");
  }

  const validated = [];
  for (const r of results) {
    if (!r || typeof r.gstin !== "string" || typeof r.flagged !== "boolean" || typeof r.reasoning !== "string") {
      continue; // skip a malformed entry rather than failing the whole batch over one bad row
    }
    validated.push({
      gstin: r.gstin,
      vendorName: typeof r.vendorName === "string" ? r.vendorName : "",
      flagged: r.flagged,
      category: typeof r.category === "string" ? r.category : null,
      subClause: typeof r.subClause === "string" ? r.subClause : null,
      reasoning: r.reasoning,
    });
  }

  const missingGstins = expectedVendors.filter((v) => !validated.some((r) => r.gstin === v.gstin)).map((v) => v.gstin);

  return { results: validated, missingGstins };
}

// vendors: [{ gstin, vendorName }, ...], already deduplicated by the
// caller. callAnthropic: (requestBody) => Promise<parsedJsonResponse> —
// injected so this function runs the same way in tests (fake) and
// production (callAnthropicApi below).
async function screenVendorsForIneligibleItc(vendors, callAnthropic) {
  if (!Array.isArray(vendors) || vendors.length === 0) {
    return { results: [], missingGstins: [] };
  }
  const requestBody = buildAnthropicRequestBody(vendors);
  const apiResponse = await callAnthropic(requestBody);
  return parseScreeningResponse(apiResponse, vendors);
}

// The real network call — a thin wrapper Node's built-in https module
// (no new dependency needed). Not used by the unit test; used by
// mentor-suggestions-proxy.js in production.
function callAnthropicApi(requestBody, apiKey) {
  return new Promise((resolve, reject) => {
    const payload = JSON.stringify(requestBody);
    const req = https.request(
      ANTHROPIC_API_URL,
      {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": ANTHROPIC_API_VERSION,
          "content-length": Buffer.byteLength(payload),
        },
      },
      (res) => {
        let body = "";
        res.on("data", (chunk) => (body += chunk));
        res.on("end", () => {
          let parsed;
          try {
            parsed = JSON.parse(body);
          } catch (parseError) {
            reject(new Error("Anthropic API returned a non-JSON response (status " + res.statusCode + "): " + body.slice(0, 300)));
            return;
          }
          if (res.statusCode < 200 || res.statusCode >= 300) {
            const message = parsed && parsed.error && parsed.error.message ? parsed.error.message : JSON.stringify(parsed);
            reject(new Error("Anthropic API error (status " + res.statusCode + "): " + message));
            return;
          }
          resolve(parsed);
        });
      }
    );
    req.on("error", reject);
    req.write(payload);
    req.end();
  });
}

module.exports = {
  ANTHROPIC_API_URL,
  ANTHROPIC_MODEL,
  ANTHROPIC_API_VERSION,
  SCREENING_TOOL,
  buildScreeningPrompt,
  buildAnthropicRequestBody,
  parseScreeningResponse,
  screenVendorsForIneligibleItc,
  callAnthropicApi,
};
