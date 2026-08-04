/*
 * fallback-sheet-memory-store.js
 * --------------------------------
 * Composes two stores that both satisfy the sheet-memory-store contract
 * (findExact/findSimilar/remember/list) so one is authoritative for reads
 * and the other is a fallback — used to make workbook-embedded storage the
 * source of truth for "has this sheet already been classified?" while
 * keeping localStorage as a secondary check, not the other way around.
 *
 * Pure composition, no platform dependency of its own — works with any two
 * stores implementing the same contract (Node's JsonFileSheetMemoryStore,
 * the browser's BrowserSheetMemoryStore, or a future Supabase-backed one).
 */

class FallbackSheetMemoryStore {
  constructor(primary, secondary) {
    this.primary = primary;
    this.secondary = secondary;
  }

  async findExact(clientId, structuralSignature) {
    const fromPrimary = await this.primary.findExact(clientId, structuralSignature);
    if (fromPrimary) return fromPrimary;
    return this.secondary.findExact(clientId, structuralSignature);
  }

  async findSimilar(clientId, structuralSignature, headerSignature) {
    const fromPrimary = await this.primary.findSimilar(clientId, structuralSignature, headerSignature);
    if (fromPrimary) return fromPrimary;
    return this.secondary.findSimilar(clientId, structuralSignature, headerSignature);
  }

  // Written to BOTH: the primary is authoritative for reads (that's the
  // whole point), but a secondary that's never actually written to isn't a
  // meaningful fallback either — this keeps it a real safety net if the
  // primary is briefly unavailable, not just an empty promise of one.
  // Each write is isolated: one side failing is logged loudly but doesn't
  // stop the other from being attempted, and doesn't abort the caller's
  // subsequent steps (e.g. the Audit Log write) unless BOTH fail.
  async remember(clientId, info) {
    let primaryResult = null;
    let primaryError = null;
    try {
      primaryResult = await this.primary.remember(clientId, info);
      console.log("MENTOR sheet-memory: primary (workbook) store write succeeded", primaryResult);
    } catch (err) {
      primaryError = err;
      console.error("MENTOR sheet-memory: primary (workbook) store write FAILED", err);
    }

    try {
      await this.secondary.remember(clientId, info);
      console.log("MENTOR sheet-memory: secondary (localStorage) store write succeeded");
    } catch (err) {
      console.error("MENTOR sheet-memory: secondary store write failed", err);
    }

    if (!primaryResult && primaryError) throw primaryError; // only fail the whole call if BOTH sides never got a result
    return primaryResult;
  }

  async list(clientId) {
    const [primaryList, secondaryList] = await Promise.all([this.primary.list(clientId), this.secondary.list(clientId)]);
    const bySignature = new Map();
    secondaryList.forEach((r) => bySignature.set(r.structural_signature, r));
    primaryList.forEach((r) => bySignature.set(r.structural_signature, r)); // primary wins on overlap
    return [...bySignature.values()];
  }
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = { FallbackSheetMemoryStore };
}
