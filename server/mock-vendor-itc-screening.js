/*
 * mock-vendor-itc-screening.js
 * ---------------------------------
 * A deliberately simple, LOCAL, keyword-based stand-in for the real
 * Anthropic screening call — used ONLY when mentor-suggestions-proxy.js
 * is running in mock mode (no ANTHROPIC_API_KEY configured, or
 * MENTOR_MOCK_LLM=1 forced explicitly). This is NOT a substitute for the
 * real model's judgment — it exists purely so the surrounding pipeline
 * (client button -> HTTP call -> structured response -> sheet write) can
 * be built and verified end-to-end with zero API cost and zero real
 * network calls, before a real key is ever supplied.
 *
 * Every result this produces is tagged mock: true and its reasoning text
 * is prefixed "[MOCK]" so it can never be mistaken for a real screening
 * result if it somehow ended up on a real sheet.
 */

const MOCK_RULES = [
  {
    test: /\b(motor|vehicle|car|automobile)s?\b/i,
    category: "Motor Vehicles",
    subClause: "17(5)(a)/(ab)",
    reasoning: "[MOCK] Name suggests a motor vehicle dealer/related business — blocked unless used for resale, passenger transport, driving training, or goods transport. This is a mock rule, not a real screening.",
  },
  {
    test: /\binsuranc/i,
    category: "Insurance (motor/employee-health portion)",
    subClause: "17(5)(b)",
    reasoning: "[MOCK] General insurers sell both blocked (motor, employee health/life) and non-blocked (property, marine, liability) policies — depends on which policy was purchased. This is a mock rule, not a real screening.",
  },
  {
    test: /\b(hotel|resort)s?\b/i,
    category: "Food/Beverage/Catering (banquet portion)",
    subClause: "17(5)(b)(i)",
    reasoning: "[MOCK] Hotels commonly bill both accommodation (generally not blocked) and food/catering/banquet services (blocked) — depends on what was booked. This is a mock rule, not a real screening.",
  },
];

function mockScreenOneVendor(vendor) {
  for (const rule of MOCK_RULES) {
    if (rule.test.test(vendor.vendorName)) {
      return { gstin: vendor.gstin, vendorName: vendor.vendorName, flagged: true, category: rule.category, subClause: rule.subClause, reasoning: rule.reasoning, mock: true };
    }
  }
  return {
    gstin: vendor.gstin,
    vendorName: vendor.vendorName,
    flagged: false,
    category: null,
    subClause: null,
    reasoning: "[MOCK] No keyword match against the mock rule set (motor/vehicle/car, insurance, hotel/resort) — not a real screening decision.",
    mock: true,
  };
}

function mockScreenVendors(vendors) {
  return { results: vendors.map(mockScreenOneVendor), missingGstins: [], mock: true };
}

module.exports = { mockScreenVendors, MOCK_RULES };
