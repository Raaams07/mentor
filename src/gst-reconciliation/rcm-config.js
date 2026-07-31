/*
 * rcm-config.js
 * --------------
 * Editable list of currently-notified Reverse Charge Mechanism (RCM)
 * categories under Section 9(3)/9(4) CGST Act and Section 5(3)/5(4) IGST
 * Act — used by rcm-detector.js ONLY as a fallback, for rows where
 * GSTR-2A's own "Supply Attract Reverse Charge" flag isn't present.
 *
 * Kept as one small, separate, easily-editable list (not embedded in
 * detection logic) specifically because notified categories change over
 * time — adding a newly-notified category, or correcting a notification
 * reference, should be a one-entry edit here, not a code change.
 *
 * `notification` fields are best-effort references to the relevant CBIC
 * notification and should be re-verified against the current notification
 * before being relied on for actual compliance decisions — that's exactly
 * why they're isolated here rather than scattered through logic.
 *
 * `keywords` match against free-text expense/particulars descriptions
 * (see rcm-detector.js) — generic category vocabulary, not tied to any
 * specific company's ledger names.
 */

const RCM_CATEGORIES = [
  {
    id: "gta_freight",
    label: "Goods Transport Agency (GTA) services — freight",
    notification: "Notification 13/2017-CT(Rate), Entry 1",
    keywords: ["gta", "goods transport agency", "freight", "transportation charges", "lorry freight", "truck freight"],
  },
  {
    id: "legal_services",
    label: "Legal services from an advocate / firm of advocates",
    notification: "Notification 13/2017-CT(Rate), Entry 2",
    keywords: ["advocate", "legal fees", "legal services", "law firm", "attorney fees"],
  },
  {
    id: "security_services",
    label: "Security services (supply of security personnel) by a non-body-corporate supplier",
    notification: "Notification 13/2017-CT(Rate), Entry 14 (inserted by Notification 29/2018-CT(Rate))",
    keywords: ["security services", "security guard", "security personnel"],
  },
  {
    id: "director_fees",
    label: "Services supplied by a director to the company",
    notification: "Notification 13/2017-CT(Rate), Entry 6",
    keywords: ["director fees", "director sitting fees", "director remuneration", "sitting fees"],
  },
  {
    id: "sponsorship",
    label: "Sponsorship services",
    notification: "Notification 13/2017-CT(Rate), Entry 4",
    keywords: ["sponsorship"],
  },
  {
    id: "rent_unregistered_landlord",
    label: "Rent paid to an unregistered landlord",
    notification: "Notification 05/2022-CT(Rate) (residential renting to a registered person)",
    // Weaker signal deliberately: "rent" alone can't distinguish a
    // registered from an unregistered landlord — a keyword hit here should
    // be treated as "worth checking the landlord's registration status",
    // not a confirmed RCM case, unlike the other entries in this list.
    keywords: ["rent", "rental", "lease rent"],
    requiresManualConfirmation: true,
  },
  {
    id: "metal_scrap",
    label: "Purchase of metal scrap from an UNREGISTERED supplier only",
    // Confirmed: Notification No. 06/2024-CT(Rate), effective 10-Oct-2024.
    // Conditional, not a blanket rule: this only applies when the supplier
    // has NO GSTIN. If the supplier IS registered, ordinary forward-charge
    // GST applies instead, with a separate 2% TDS obligation under a
    // different mechanism (Section 51 CGST Act TDS provisions) — NOT RCM.
    // requiresUnregisteredSupplier tells the detector to only confirm this
    // match when the row's own GSTIN field is blank; a keyword hit with a
    // populated GSTIN is a registered supplier and must NOT be flagged
    // as RCM under this entry.
    notification: "Notification 06/2024-CT(Rate), effective 10-Oct-2024 — RCM only when supplier is unregistered; registered-supplier metal scrap purchases attract forward-charge GST + Section 51 TDS instead",
    keywords: ["metal scrap", "scrap metal", "iron scrap", "steel scrap", "ferrous scrap"],
    requiresUnregisteredSupplier: true,
  },
];

if (typeof module !== "undefined" && module.exports) {
  module.exports = { RCM_CATEGORIES };
}
