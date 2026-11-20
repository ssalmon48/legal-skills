---
name: toa-formatting
description: Formatting layer for a Word Table of Authorities. Forces long entries to break cleanly so page-number lists never split across lines, and, when a court or user convention calls for it, emphasizes (italics by default) the citation identifier of statutes, regulations, and rules like case names. Run after or alongside the table-of-authorities build.
---

## When to use
After building (or when fixing) a Word **Table of Authorities**, to apply two formatting conventions the base `table-of-authorities` skill doesn't cover:
1. Long entries must break so the auto-generated **page-number list never splits** across two lines.
2. **Optional:** statute / regulation / rule **identifiers** carry the same emphasis as case names. Bluebook court-document style (Bluepages) leaves statutory compilations roman, so do this only when the user asks for it or the brief's body already emphasizes those identifiers.

Triggers: "italicize the statutes/rules in the TOA," "the page cites are splitting onto two lines," "make the authorities table format cleanly." This is the formatting layer — the base skill handles extraction, TA marking, and TOA field insertion.

## Key mechanism — editing a TA field's `\l` in place
TOA entry text and formatting come from the `\l` switch of the TA fields; the compiled table regenerates on F9, so formatting must live in the `\l`, not in the rendered table. Never set `field.code` (strips `xml:space` and breaks the entry) and never `insertOoxml` onto a hidden TA field's empty `result` (it gets swallowed/nested and the field disappears). Rebuild the field in place, one field per two syncs:
```javascript
f.select();
await context.sync();
context.document.getSelection().insertOoxml(newPkg, "Replace");
await context.sync();
```
`f` may come from `body.fields` or from any `footnote.body.fields`. The dedup key is the exact `\l` text — rebuild **every** field sharing an authority's `\l` with identical new `\l`, or Word splits it into duplicate entries. Package shape = the flat-OPC single-paragraph package from the base skill; the `\l` value is split across runs (an emphasized `name` run + plain runs), `\s` stays the plain full cite, `\c` unchanged.

## Step 1 — Detect the emphasis convention
Read a case-name run already in the TOA (or a body cite): italic = `<w:i/>`, underlined = `<w:u w:val="single"/>`. Mirror whichever the table already uses. If nothing is emphasized yet, use italics (Bluepages B2 permits either, consistently applied).

## Step 2 — Emphasize statute/reg/rule identifiers (opt-in)
Skip this step unless the user asked for it or the brief's body emphasizes statute/rule identifiers. When it applies, rebuild each statute/reg/rule TA field so the leading identifier sits in the emphasized `name` run and the remainder is plain:
- **Statutes / regs:** emphasize the whole identifier incl. internal periods — `name = "N.J.S.A."`, rest `" 34:19-3"`; likewise `N.J.A.C.`, `U.S.C.`
- **Evidence / professional rules:** `name = "N.J.R.E."` rest `" 403"`; `name = "R.P.C."` rest `" 3.4(e)"`
- **Bare court rules (`R. 2:5-1`): emphasize ONLY the letter `R`** — `name = "R"`, rest `". 2:5-1"`. The period is NOT emphasized.
- Section/subsection numbers and reporters stay plain; "Other Authorities" stay fully plain.
Enumerate the target fields by parsing `\c` (2 = statute, 4 = rule, 6 = reg) and `\l` from each field's `code`.
**Done when:** in the recompiled table, the identifier reads `font.italic === true` (or `font.underline === "Single"`) and the numbers read plain; for `R.` cites the "R. " range reads mixed (R emphasized, period not).

## Step 3 — Stop page-number lists from splitting
Long entries wrap; the failure is the page list (e.g. `26, 39, 48, 72`) splitting across lines, or the reporter breaking mid-cite. Fix inside `\l`:
- **Non-breaking spaces in the reporter:** join reporter tokens with `\u00A0` (`140\u00A0N.J.\u00A0366\u00A0(1995)`) so it never splits internally; keep a normal breakable space *before* the volume.
- **Hard break before the reporter, only for entries that actually wrap:** insert a `<w:br/>` **run** between the name run and the reporter run so the reporter + tab + page list start together on a fresh (roomy) line. A raw U+000B is invalid XML and rejected — always use the `<w:br/>` element:
  ```
  ...emphasized name run... +
  `<w:r><w:instrText xml:space="preserve">,</w:instrText></w:r>` +
  `<w:r><w:br/></w:r>` +
  `<w:r><w:instrText xml:space="preserve">140\u00A0N.J.\u00A0366\u00A0(1995)" \s "<plain full>" \c 1 </w:instrText></w:r>`
  ```
  Verify: `field.result.text` shows a `\u000b` between name and reporter.
- **Identify which entries wrap** with `verify_doc_visual` (ask the reviewer to list case entries whose citation/page list wraps to a second line). Only break those — forcing a break on a short entry needlessly doubles its height.
- **Unpublished/docket cites:** break before the reporter *volume* (e.g. before `2005`), leaving `No. A-…,` on line one; don't bind the whole docket+reporter into one non-breaking run (it overflows the margin).
**Done when:** the visual review shows no split page-number list and no reporter broken mid-cite.

## Step 4 — Recompile and verify
`updateResult()` on every TOA field, read back category counts (unchanged — no new/duplicate entries), confirm emphasis and wrapping, and tell the user to press **Ctrl+A → F9** in Word to settle page numbers.
