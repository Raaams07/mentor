/*
 * vendor-itc-screening-test.js
 * ---------------------------------
 * Unit tests for anthropic-vendor-itc-screening.js's prompt-building and
 * response-parsing logic, using a FAKE injected callAnthropic — no real
 * network call, no API key needed or used.
 *
 * IMPORTANT LIMITATION: this validates that the plumbing (prompt shape,
 * request shape, response parsing/validation) is correct — it does NOT
 * validate that the real Claude model actually produces good screening
 * judgments for real vendor names. That requires a live call with a real
 * ANTHROPIC_API_KEY, which this test deliberately does not have access to.
 * See mentor-suggestions-proxy.js's own startup instructions for how to
 * run the real end-to-end check yourself.
 *
 * Run with: node server/vendor-itc-screening-test.js
 */

const { buildScreeningPrompt, buildAnthropicRequestBody, parseScreeningResponse, screenVendorsForIneligibleItc, SCREENING_TOOL, ANTHROPIC_MODEL } = require("./anthropic-vendor-itc-screening.js");

let failures = 0;
function assert(condition, description) {
  console.log(`${condition ? "PASS" : "FAIL"}  ${description}`);
  if (!condition) failures++;
}

const SYNTHETIC_VENDORS = [
  { gstin: "27ZZAAA1111A1Z1", vendorName: "Speedway Motors Pvt Ltd" }, // clear vehicle dealer
  { gstin: "07ZZBBB2222B1Z2", vendorName: "Bharat General Insurance Co Ltd" }, // clear insurer
  { gstin: "29ZZCCC3333C1Z3", vendorName: "Sunrise Retail Traders" }, // ambiguous
];

// A hand-built response shaped exactly like Anthropic's real Messages API
// tool_use response — this is what I'd EXPECT a well-behaved call to
// return for these three vendors, based on the prompt's own rules. It is
// NOT the output of an actual API call.
function fakeAnthropicResponse() {
  return {
    id: "msg_fake",
    type: "message",
    role: "assistant",
    content: [
      {
        type: "tool_use",
        id: "toolu_fake",
        name: "report_vendor_itc_screening",
        input: {
          results: [
            {
              gstin: "27ZZAAA1111A1Z1",
              vendorName: "Speedway Motors Pvt Ltd",
              flagged: true,
              category: "Motor Vehicles",
              subClause: "17(5)(a)/(ab)",
              reasoning: "Name indicates a motor vehicle dealer; a purchase here is likely a vehicle or vehicle-related service, blocked unless used for resale, passenger transport, driving training, or goods transport.",
            },
            {
              gstin: "07ZZBBB2222B1Z2",
              vendorName: "Bharat General Insurance Co Ltd",
              flagged: true,
              category: "Insurance (motor/employee-health portion)",
              subClause: "17(5)(b)",
              reasoning: "A general insurer sells both blocked policies (motor vehicle, employee health/life) and non-blocked ones (property, marine, liability) — eligibility depends on which policy was actually purchased; review the invoice.",
            },
            {
              gstin: "29ZZCCC3333C1Z3",
              vendorName: "Sunrise Retail Traders",
              flagged: false,
              category: null,
              subClause: null,
              reasoning: "Generic retail trading name with no indication of a Section 17(5) blocked category — declining to flag rather than guess.",
            },
          ],
        },
      },
    ],
  };
}

function runTests() {
  console.log("-- Prompt construction --\n");

  const prompt = buildScreeningPrompt(SYNTHETIC_VENDORS);
  assert(prompt.includes("Speedway Motors Pvt Ltd") && prompt.includes("Bharat General Insurance Co Ltd") && prompt.includes("Sunrise Retail Traders"), "prompt includes every vendor's name");
  assert(prompt.includes("27ZZAAA1111A1Z1"), "prompt includes GSTINs");
  // No separate "no amounts" assertion needed here -- the vendors objects
  // passed in only ever carry {gstin, vendorName} (see the extraction
  // logic in mentor-gst-reconciliation-ui.js), so there is structurally
  // nothing else buildScreeningPrompt() could include in the prompt.
  assert(prompt.includes("Be conservative"), "prompt includes the conservative/don't-guess instruction");
  assert(prompt.includes("hedge"), "prompt includes the explicit-hedge instruction for mixed-business vendors");

  const requestBody = buildAnthropicRequestBody(SYNTHETIC_VENDORS);
  assert(requestBody.model === ANTHROPIC_MODEL, "request uses the configured model");
  assert(requestBody.tool_choice.type === "tool" && requestBody.tool_choice.name === SCREENING_TOOL.name, "tool_choice forces the screening tool — no freeform prose response");
  assert(requestBody.messages.length === 1 && requestBody.messages[0].role === "user", "single batched user message, not one call per vendor");

  console.log("\n-- Response parsing: well-formed response --\n");

  const parsed = parseScreeningResponse(fakeAnthropicResponse(), SYNTHETIC_VENDORS);
  assert(parsed.results.length === 3, "all 3 vendors parsed");
  assert(parsed.missingGstins.length === 0, "no missing vendors");

  const motors = parsed.results.find((r) => r.gstin === "27ZZAAA1111A1Z1");
  assert(motors.flagged === true && motors.subClause === "17(5)(a)/(ab)", "clear vehicle dealer flagged under the motor vehicles sub-clause");

  const insurer = parsed.results.find((r) => r.gstin === "07ZZBBB2222B1Z2");
  assert(insurer.flagged === true && /depends on|which policy/i.test(insurer.reasoning), "insurer flagged WITH an explicit hedge in the reasoning, not an unconditional flag");

  const retailer = parsed.results.find((r) => r.gstin === "29ZZCCC3333C1Z3");
  assert(retailer.flagged === false && retailer.category === null, "ambiguous retailer correctly NOT flagged, no category guessed");

  console.log("\n-- Response parsing: malformed / partial responses handled explicitly, not silently --\n");

  const noToolUse = { content: [{ type: "text", text: "I'll just describe this in prose instead..." }] };
  let threw = false;
  try {
    parseScreeningResponse(noToolUse, SYNTHETIC_VENDORS);
  } catch (e) {
    threw = true;
    assert(/tool_use/.test(e.message), "error message explains the model didn't call the tool");
  }
  assert(threw, "a response with no tool_use block throws instead of returning something malformed");

  const missingOneVendor = fakeAnthropicResponse();
  missingOneVendor.content[0].input.results = missingOneVendor.content[0].input.results.slice(0, 2); // drop the retailer
  const partialParsed = parseScreeningResponse(missingOneVendor, SYNTHETIC_VENDORS);
  assert(partialParsed.results.length === 2 && partialParsed.missingGstins.length === 1 && partialParsed.missingGstins[0] === "29ZZCCC3333C1Z3", "a vendor the model silently dropped is reported as missing, not ignored");

  const oneMalformedEntry = fakeAnthropicResponse();
  oneMalformedEntry.content[0].input.results.push({ gstin: "BADROW", flagged: "not a boolean", reasoning: 12345 }); // wrong types
  const withBadEntry = parseScreeningResponse(oneMalformedEntry, SYNTHETIC_VENDORS);
  assert(withBadEntry.results.length === 3, "a structurally malformed entry is skipped rather than corrupting the whole batch");

  console.log("\n-- End-to-end with injected fake callAnthropic (no network) --\n");

  return screenVendorsForIneligibleItc(SYNTHETIC_VENDORS, async (requestBody) => {
    assert(requestBody.messages[0].content.includes("Speedway Motors"), "the injected callAnthropic received the built prompt");
    return fakeAnthropicResponse();
  }).then((result) => {
    assert(result.results.length === 3, "screenVendorsForIneligibleItc returns all 3 results end-to-end");

    console.log("\n-- Empty vendor list short-circuits without calling the API at all --\n");
    let called = false;
    return screenVendorsForIneligibleItc([], async () => {
      called = true;
      return fakeAnthropicResponse();
    }).then((emptyResult) => {
      assert(called === false, "callAnthropic is never invoked for an empty vendor list");
      assert(emptyResult.results.length === 0, "empty vendor list returns an empty result, not an error");
    });
  });
}

async function run() {
  await runTests();
  console.log("");
  if (failures > 0) {
    console.log(`${failures} check(s) FAILED.`);
    process.exitCode = 1;
  } else {
    console.log("All vendor ITC screening plumbing checks passed.");
    console.log("NOTE: this only validates prompt/request/response plumbing against a hand-built fake response.");
    console.log("It does NOT validate real model judgment quality -- that needs a live call with your own ANTHROPIC_API_KEY.");
  }
}

run();
