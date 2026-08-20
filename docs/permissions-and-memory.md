# Permissions: where a future gate belongs, relative to memory

**Status: awareness only. Nothing here is built. This document exists so a
future permission layer doesn't require redesigning memory to make room for
it.**

## The problem this heads off

Today, `clientId` is a bare placeholder string — `"workbook:" + workbookName`
— constructed independently in three places:

- `mentorGetSheetMemoryClientId()` — `src/taskpane/mentor-sheet-memory-ui.js`
- `mentorGetColumnMemoryClientId()` — `src/taskpane/mentor-gst-reconciliation-ui.js`
- `mentorGetMemoryReviewClientId()` — `src/taskpane/mentor-memory-review-ui.js`

It's the sole scoping key for every Tier-2 (per-client) memory read and
write: `resolveSheetLabel`/`rememberSheetLabel` (`sheet-memory.js`),
`resolveColumnField`/`rememberColumnField` (`column-memory.js`), and the
Learned Answers review panel's direct store reads
(`mentor-memory-review-ui.js`). It identifies **which workbook file** is
asking, and nothing else — there's no notion of **who** (which person, which
role) is asking on that workbook's behalf.

That's fine as long as MENTOR has no permission model at all. It stops being
fine the moment one exists, for a structural reason worth flagging now: a
permission check answers "is *this actor* allowed to do *this* on *this
client's* data" — and `clientId` today can't express "this actor," only
"this workbook." Bolting permissions on later without having planned for
this would mean either overloading `clientId` with actor identity in a way
that breaks every existing memory record's key, or threading a *second*,
uncoordinated identity concept through the same call sites. Neither is a
good migration.

## Where the gate goes: in front of memory, not inside it

Memory (`sheet-memory.js`, `column-memory.js`, the `WorkbookSheetMemoryStore`
/ `WorkbookColumnMemoryStore` / `BrowserSheetMemoryStore` /
`FallbackSheetMemoryStore` stack) stays **actor-agnostic**. It trusts
whatever `clientId` it's handed and has no opinion about who's allowed to
read or write it. A permission check is a **gate that runs before** any of
these functions are called — not new logic woven into them.

Concretely, when this gets built, the shape is:

```js
// today:
const clientId = mentorGetSheetMemoryClientId(context.workbook.name);
const result = await resolveSheetLabel({ clientId, sheetName, sheetSignals, classification, store });

// with a permission layer:
const clientId = mentorGetSheetMemoryClientId(context.workbook.name, currentActor);
if (!permissionCheck(currentActor, clientId, "read")) { /* deny, don't call resolveSheetLabel at all */ }
const result = await resolveSheetLabel({ clientId, sheetName, sheetSignals, classification, store });
```

The three `mentorGet*ClientId()` functions above are exactly where the
actor identity needs to enter the picture, since they're the only places
`clientId` is constructed — every downstream call already just receives
whatever string they produce. Each one now carries a comment pointing here.

## `clientId` becomes a compound key: workbook × actor

`"workbook:" + workbookName` is a single-axis key. The moment two different
people (a preparer, a reviewing partner, a second staff member) can open the
*same* workbook file and need *different* memory/permission scopes, a
workbook-only key can't tell them apart — either they'd silently share one
memory record (wrong: one person's correction becomes invisible-but-active
for someone with no idea it happened), or the key needs a second axis. The
natural fix is a compound key: `workbook:<name>::actor:<id>` (or similar) —
still one string, still passed through every existing call site unchanged
in shape, just built from two inputs instead of one.

This is a `mentorGet*ClientId()`-level change only. Nothing downstream
(store classes, `resolveSheetLabel`/`resolveColumnField`, the review UI's
lookups) needs to know or care that the string now encodes two things —
they already treat `clientId` as an opaque scoping key, which is exactly
what makes this safe to defer.

## Tier 1 is a separate sub-problem, not covered by the above

The shared, cross-client pattern stores (`column-pattern-store.js`,
`sheet-label-pattern-store.js`) are deliberately **not** keyed by `clientId`
at all — that's the point of Tier 1, see their own docstrings. A permission
question there isn't "can this actor read/write *this client's* data," it's
a coarser "is this actor allowed to read/write the *shared* store at all" —
closer to a feature flag than a per-record ACL. Don't conflate this with the
compound-key change above when the time comes; it needs its own, simpler
design.

## What this document is NOT

Not a schema. Not an implementation plan. Not a promise about *when* this
gets built. It exists purely so that whenever a permission model becomes a
real, scoped piece of work, the person building it (human or MENTOR itself)
starts from "extend `mentorGet*ClientId()` and add a gate in front of the
existing calls" instead of rediscovering that memory itself needs to change.
Memory does not need to change.
