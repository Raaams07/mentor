# MENTOR architecture principles

## Propose, don't auto-execute

MENTOR computes, matches, and drafts. It never files, pays, or sends
anything on a human's behalf without that human first seeing exactly what
it intends to do and explicitly accepting it.

Concretely, in this codebase today: every reconciliation is computed
read-only first (`src/gst-reconciliation/*.js` — pure functions, no
workbook writes, no side effects) and rendered as a proposal card. Nothing
gets written to the user's workbook (`gst-report-writer.js`) until the user
clicks Accept in the task pane (`mentor-gst-reconciliation-ui.js`). This
applies uniformly across the pipeline — vendor matching, Wrong Head, RCM,
Ineligible ITC, Duplicate Invoices, the Summary netting — not just to the
final write step. See the docstrings of `gst-report-writer.js`,
`mentor-gst-reconciliation-ui.js`, `gst-summary-builder.js`,
`gstr2a-vs-books-reconciler.js`, and `ineligible-itc-detector.js` for where
each of those individually commits to this.

## Earned autonomy has a ceiling: action tiers

MENTOR is allowed to get quieter over time about things it's been reliably
right about — that's what the memory system (Tier 1/Tier 2 column- and
sheet-identity resolution) already does: ask once, then stop asking for the
same shape. That's a deliberate, working design, not an oversight to be
walked back.

But "MENTOR can earn the right to stop asking" cannot be a single, global
dial. Some actions are reversible and stay entirely inside the user's own
workbook; others leave the workbook, reach a third party, or carry legal
consequence. Trust earned in the first category must never spill over into
the second. To keep that boundary explicit rather than assumed, every
action MENTOR takes or could take falls into one of two tiers:

**🔴 — never auto-approve, regardless of accumulated trust.**
Filings, payments, and anything client-facing (sending a reconciliation,
a report, or a communication to someone outside this session). These
require an explicit human click every single time, forever — not "until
MENTOR has been right N times in a row." The reasoning: these actions are
either hard/impossible to reverse, or their cost of being silently wrong
is borne by someone who never reviewed the decision. No amount of prior
correctness on unrelated, reversible actions (a correctly-identified
column, a correctly-classified sheet, a hundred correct Wrong Head calls)
is evidence about correctness on *this* action — so it must never be used
to skip asking on this action.

**Everything else — may earn quieter behavior over time, scoped to itself.**
Column/sheet identity resolution, standing validation, proposal
computation, drafting output sheets pending Accept. All reversible, all
contained to the user's own workbook until they act on it, all specifically
scoped (a column-memory hit means "stop asking about *this field, on
sheets shaped like this*" — it says nothing about, and grants no leniency
toward, any 🔴-tier action).

**MENTOR does not currently perform any 🔴-tier action.** It has no filing,
payment, or send capability today — this section is a guardrail for
functionality that doesn't exist yet, not a restriction on anything
currently running. Any future feature that would file, pay, or send on a
user's behalf must be built as 🔴-tier from day one: proposal shown, human
clicks accept, every time, with no code path that lets repeated correct
proposals reduce that to zero clicks. If a feature seems to blur the line
between the two tiers, treat it as 🔴 until a human explicitly decides
otherwise — this list is closed by default, not open by default.

## Where this doc came from

Point 5 of a five-point architecture request (Aug 2026) modeled on how
Bridgewater's AIA Labs describes earned autonomy in their PAT tool: the
system may become less naggy about things it's reliably correct on, but a
short, explicit list of actions is permanently excluded from that dynamic.
The "propose, don't auto-execute" rule already existed informally,
scattered across five files' docstrings, before this doc consolidated it
and made the tiering explicit.
