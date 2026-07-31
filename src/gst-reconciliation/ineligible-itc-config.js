/*
 * ineligible-itc-config.js
 * ---------------------------
 * Editable list of blocked-credit categories under Section 17(5) CGST Act
 * — Input Tax Credit is disallowed for these regardless of whether GST was
 * otherwise validly charged. Used by ineligible-itc-detector.js.
 *
 * Kept as one small, separate, easily-editable list (not embedded in
 * detection logic) for the same reason as rcm-config.js: blocked-credit
 * categories are added/amended by law over time (most recently the CSR
 * exclusion, inserted by the Finance Act 2023) — updating this list should
 * be a one-entry edit, not a code change.
 *
 * `keywords` match against free-text expense/particulars descriptions —
 * generic category vocabulary (common ledger/expense-head naming), not
 * tied to any specific company's chart of accounts. Detection by
 * description keyword is inherently approximate for this rule — several
 * of these categories genuinely require knowing the item/purpose, which a
 * bare taxable-value/GSTIN row often doesn't carry. See the detector's
 * docstring for how it surfaces that uncertainty rather than hiding it.
 */

const ITC_INELIGIBLE_CATEGORIES = [
  {
    id: "motor_vehicles",
    label: "Motor vehicles for transport of persons (≤13 seats incl. driver) and related services — except specified business uses (further supply, passenger transport, driving training, goods transport)",
    section: "Section 17(5)(a)/(ab), CGST Act",
    keywords: ["motor vehicle", "car purchase", "vehicle insurance", "car insurance", "vehicle repair", "car repair", "vehicle maintenance", "car maintenance"],
  },
  {
    id: "food_beverage",
    label: "Food and beverages, outdoor catering — except where used for making a further taxable outward supply of the same category",
    section: "Section 17(5)(b)(i), CGST Act",
    keywords: ["food", "beverage", "catering", "restaurant bill", "canteen"],
  },
  {
    id: "employee_insurance_travel",
    label: "Employee health/life insurance, rent-a-cab, leave travel benefits — except where the law mandates the employer provide them",
    section: "Section 17(5)(b)(iii)-(iv), CGST Act",
    keywords: ["health insurance", "life insurance", "rent a cab", "rent-a-cab", "leave travel", "ltc", "employee travel"],
  },
  {
    id: "works_contract_immovable",
    label: "Works contract services / goods for construction of an immovable property on own account — except plant and machinery",
    section: "Section 17(5)(c)/(d), CGST Act",
    keywords: ["works contract", "building construction", "civil construction", "construction of building", "capital work in progress", "capitalised construction"],
  },
  {
    id: "csr_expenditure",
    label: "Goods/services used for CSR (Corporate Social Responsibility) activities",
    section: "Section 17(5)(fa), CGST Act (inserted by the Finance Act 2023, effective 01-Oct-2023)",
    keywords: ["csr", "corporate social responsibility"],
  },
  {
    id: "personal_use",
    label: "Goods/services for personal consumption",
    section: "Section 17(5)(g), CGST Act",
    keywords: ["personal use", "personal consumption"],
  },
  {
    id: "composition_scheme_purchase",
    label: "Purchases from a supplier registered under the Composition Scheme (no tax separately charged/creditable)",
    section: "Section 17(5)(e), CGST Act; mechanics of Section 10 composition scheme",
    // Not description-keyword-detectable — this depends on the SUPPLIER's
    // registration type, not what was bought. Left with no keywords
    // deliberately; a real implementation would need a filing-status-type
    // signal (see ineligible-itc-detector.js) rather than text matching.
    keywords: [],
  },
];

if (typeof module !== "undefined" && module.exports) {
  module.exports = { ITC_INELIGIBLE_CATEGORIES };
}
