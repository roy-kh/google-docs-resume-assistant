# Resume Assistant Review (Apps Script)

This captures the current weaknesses and the concrete remediation plan without touching auth or storage. References point to existing logic in `Code.gs` and `sidebar.html`.

## Core Problems
- **Apply-first architecture:** `processJobDescription()` calls OpenAI → parses JSON → immediately `applySuggestionsToDoc()` with no review/approval step.
- **Brittle matching & formatting loss:** `body.findText(original)` only swaps the first contiguous match; replacements operate on paragraphs/text runs that may span bullets/line breaks. Attribute copy uses a single character snapshot, missing paragraph styles and split runs, leading to italics/spacing drift.
- **Prompt allows expansion:** Current prompts invite elaboration; there are no length/line constraints, so outputs grow and overflow the page.
- **Generated metadata unused:** Suggestions include `reason` but the UI never surfaces it; there’s no validation on length or newlines.
- **Silent failure modes:** `parseSuggestions()` falls back to `[]` and still proceeds; single-pass replacement means partial rewrites feel random.

## Priority Order (do first)
1) **Prompt contract (biggest win)**
   - Add hard constraints: no new bullets, no new lines, no added paragraphs, keep within original character count (e.g., ≤110% of original), preserve bullet count.
   - Include explicit rules: “Do not add line breaks,” “Do not add new bullets,” “Return JSON only.”

2) **Suggest → review → apply flow**
   - Keep OpenAI call and auth untouched.
   - Split responsibilities: generation (returns suggestions) → sidebar review list → user approval → apply selected items.
   - Show `original`, `suggestion`, and `reason` in the sidebar; include an Apply checkbox per item and an Apply Selected action.

3) **Safer replacement granularity**
   - Avoid replacing full paragraphs. Match within text runs; do not insert paragraph breaks.
   - When a match spans multiple runs, concatenate text in-memory, map back to ranges, and replace only text content while keeping existing paragraph/bullet structure.

4) **Defensive validation before apply**
   - Reject suggestions that: introduce `\n`, exceed X% of original length, or are empty.
   - Skip-apply if validation fails; surface the reason in the UI instead of failing silently.

## Targeted Code Changes (no auth/storage changes)
- `processJobDescription()`: split into `generateSuggestions()` (returns parsed/validated suggestions) and a new sidebar-driven apply step; stop auto-calling `applySuggestionsToDoc()`.
- `parseSuggestions()`: fail loudly on bad JSON; surface parsing errors in the UI instead of returning `[]`.
- `applySuggestionsToDoc()`: operate on text runs, preserve formatting attributes per range, and never create new paragraphs; handle multiple matches or clarify only-first-match behavior in UI.
- `sidebar.html`: add a suggestion list view with reason text, approval checkboxes, and an Apply Selected button; include status for rejected suggestions (too long, contains newlines, not found).
- Prompts (both JD and selection flows): add explicit length/line/bullet constraints and “no added bullets/lines” rules.

## Quick Wins to Implement First
- Update prompts with strict length + no-newline/no-new-bullet constraints.
- Change the flow to: Generate → display suggestions with reasons → user chooses → apply.
- Add validation: reject if suggestion length >110% of original or contains `\n`.
- Improve error surfacing: parsing or apply failures show in the sidebar with counts.

## What to Keep Untouched (works fine)
- Auth flows, API key handling, and OpenAI request wiring are solid; do not modify.
