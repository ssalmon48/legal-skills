---
name: table-of-authorities
description: Build an auto-updating Table of Authorities for a legal brief using real Word TA/TOA field codes — scans for every authority, marks each with a citation field, and generates category-split tables (cases, statutes, regulations, rules, constitutional provisions, other authorities) with live page numbers.
---

## When to use
When the user asks to build, generate, or fix a **Table of Authorities** for a legal brief — "create a table of authorities," "mark the citations," "make the TOA auto-update," "add statutes/rules to the authorities table." Produces a *real* Word TOA driven by TA field codes, so page numbers update on field refresh — never a hand-typed table.

## Inputs
- `target` — the open brief. No argument needed; operate on the whole document body.
- Scope is every authority cited anywhere in the body (all Points, Standard of Review, Preliminary Statement, footnotes).

## Goal
Every distinct authority is marked with a well-formed hidden `TA` field at each place it's cited, and one `TOA` field per category present renders a category-split, alphabetized, auto-updating table under the "TABLE OF AUTHORITIES" heading. Verified by reading the compiled result back in-session.

## Platform note
Desktop Word's `Word.FieldType` enum does **not** expose `tableOfAuthoritiesEntry`/`tableOfAuthorities`, so `range.insertField(...)` cannot build these. All TA/TOA fields are inserted via `range.insertOoxml(pkg, "After")` with a flat-OPC package. Do **not** use `field.code = "..."` to edit a TA field — assigning `code` strips `xml:space="preserve"` and the leading space from the instruction text, which leaves `field.type === "TA"` but makes Word's TOA compiler silently skip the entry (F9 builds nothing). If you must change a field, delete it and re-insert.

## Steps

### 1. Read the brief and extract authorities
Load the full body text. Find every citation. For each, determine:
- the **base citation** for the entry — the citation **without the pincite**. If the text says `203 N.J. 252, 289`, the entry citation is `203 N.J. 252 (2010)` (drop the `, 289`). The pincite is never written into the entry; the page a source appears on is captured automatically by *where* the TA field sits.
- every **search string** that locates an occurrence (full cite plus short forms like `Rosenblit, 166 N.J. at 407.` and `166 N.J. at 408.`). You mark *all* occurrences so every page is captured.
**Done when:** you have a list of distinct authorities, each with its base citation and the set of searchable occurrence strings.
**Rule:** Dedupe against any TA fields already in the document (`body.fields`, `f.type === "TA"`, parse `\l "..."`). Preserve existing native fields; don't double-mark an authority that already has a correct field.

### 2. Classify each authority into a category (silently)
Assign each to a Word TOA category by citation pattern. Use the built-in category numbers so the `TOA \h` switch prints the correct heading automatically:

| Category | `\c` number | Detect by |
|---|---|---|
| Cases | 1 | ` v. ` between parties + a reporter cite |
| Statutes | 2 | `U.S.C.`, `N.J.S.A.`, `§`, `Stat.`, state code names |
| Other Authorities | 3 | law-review articles, Restatements, anything uncategorized |
| Rules | 4 | `R. `, `Rule `, `Fed. R. Civ. P.`, `Fed. R. Evid.` |
| Treatises | 5 | named treatise / hornbook |
| Regulations | 6 | `C.F.R.`, `Fed. Reg.`, `N.J.A.C.` |
| Constitutional Provisions | 7 | `Const.`, `amend.`, `U.S. Const.`, `N.J. Const.` |

**Done when:** every authority has a `\c` number. Mark and report in one pass (no approval gate).

### 3. Set the emphasis style and what gets it
**Style: italics by default** (`<w:i/>`). Bluebook Bluepages B2 lets court documents use italics or underscoring as long as the choice is consistent, and the current edition lists italics as the first option. Before marking, check the brief's body citations (or any existing TA field's `\l` runs). If the body **consistently underlines** case names, use underlining (`<w:u w:val="single"/>`) instead, so the table matches the brief. If the body is mixed, use italics and tell the user the body is inconsistent.

**What gets emphasis** (Bluepages typeface conventions for court documents; Whitepages fill gaps):
- **Cases:** the full case name, including `v.` and procedural phrases that are part of the name (`In re`, `ex rel.`, `ex parte`). The comma after the name, the reporter, the court, and the year stay roman.
- **Prior/subsequent history phrases** that appear in an entry's base citation (`aff'd`, `rev'd`, `cert. denied`, `aff'd on other grounds`): emphasized. The case names and reporters around them follow the case rule.
- **Treatises, books, articles (categories 3 and 5):** emphasize the **title** only. Author names, volume numbers, journal/periodical names, sections, and edition/year parentheticals stay roman.
- **Statutes, regulations, rules, constitutional provisions (categories 2, 4, 6, 7):** fully roman. Bluepages court-document style leaves statutory compilations and code abbreviations unemphasized. (The companion `toa-formatting` skill can emphasize these identifiers if a court or user convention calls for it; that is opt-in, not the default.)
- **Punctuation:** emphasize punctuation only when it falls inside emphasized material (the period in `v.` or `Inc.`); a trailing comma or period after the emphasized text is roman.
- **Signals and `Id.`** never appear in TOA entries; strip them from the base citation.

**Done when:** you know the run property to apply and which segments of each entry get it.

### 4. Mark every occurrence with a TA field
For each occurrence, `body.search(needle)`, then for each hit insert the field `"After"` the hit's `getRange("End")`. Build the package per authority. The `\l` is passed as an ordered list of segments, each flagged emphasized or roman per Step 3, so an entry can carry more than one emphasized span (a case name plus `aff'd`, or a treatise title between a roman author and roman section):

```javascript
const esc = s => s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
// segs = [{text:"Rosenblit v. Zimmerman", em:true}, {text:", 166 N.J. 391 (2001)", em:false}]; cat = 1
// Statute: [{text:"42 U.S.C. 1983", em:false}]
// Article: [{text:"Jane Doe, ", em:false}, {text:"The Title of the Article", em:true},
//           {text:", 100 Harv. L. Rev. 1 (2020)", em:false}]
// style = "i" (italics, default) or "u" (underline, only if the brief's body underlines)
function taPkg(segs, full, cat, style = "i") {
  const emph = style === "u" ? `<w:rPr><w:u w:val="single"/></w:rPr>` : `<w:rPr><w:i/></w:rPr>`;
  const body = segs.map(s =>
    `<w:r>${s.em ? emph : ``}<w:instrText xml:space="preserve">${esc(s.text)}</w:instrText></w:r>`).join("");
  const runs =
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> TA \\l "</w:instrText></w:r>` +
    body +
    `<w:r><w:instrText xml:space="preserve">" \\s "${esc(full)}" \\c ${cat} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  return `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${runs}</w:p></w:body></w:document></pkg:xmlData></pkg:part>
</pkg:package>`;
}
```
`full` is always the plain concatenation of every segment's text (the `\s` short cite carries no formatting). For a case, the first segment is the emphasized name and the next is the roman `", <reporter> (<court year>)"`. For a statute, rule, regulation, or constitutional provision, pass a single roman segment.
**Done when:** each occurrence has a hidden TA field; read `body.fields` back and confirm the per-authority mark counts.
**Rule — skip the rendered table:** a bare citation (no pincite) can also appear verbatim inside an already-rendered TOA entry. For each search hit, load `hit.paragraphs.getFirst().style` and **skip** any hit whose style is `"Table of Authorities"` — never mark the table itself. (Use `paragraphs.getFirst()`, not `parentParagraph`, which can return undefined.)
**Rule — no pincites:** `\l` and `\s` use the base citation only (Step 1).
**Rule — both switches:** always include both `\l` and `\s` with the base citation (matches native fields; `\l` alone can be skipped by some compilers).
**Rule — XML-escape** `&`, `<`, `>` in all citation text.
**Rule — passim:** to get a page-heavy source to collapse to "passim," mark all 5+ occurrences and rely on the `\p` switch on the TOA field (Step 5). Don't hand-type "passim."
**Rule — non-atomic recovery:** `insertOoxml`/`delete` commit incrementally. If a batch throws mid-loop, re-read `body.fields` to see what actually landed and finish surgically — never blindly re-run the whole batch (it duplicates).

### 5. Insert one TOA field per category present
Under the "TABLE OF AUTHORITIES" heading, for each category that has entries, insert a field whose instruction is ` TOA \h \c "N" \p ` (via the same package shape, but a single `<w:instrText xml:space="preserve"> TOA \h \c "1" \p </w:instrText>` run between the begin/end fldChars — no emphasis run, no vanish). `\h` auto-prints the built-in category heading ("Cases", "Statutes", "Rules"…); `\p` enables passim.
**Done when:** one TOA field exists per populated category, in category order (Cases first).
**Rule — placement:** anchor each TOA field on a **plain body paragraph outside any other TOA field's result range** (e.g. an empty "Brief Body Text" paragraph after the prior table). Anchoring on another field's `result` End range causes the field to be **absorbed and deleted** when that field regenerates.

### 6. Compile and verify
Call `updateResult()` on every TOA field — it recompiles in-session once the TA fields are well-formed. Then read back: the compiled `field.result.text` should list all authorities under each heading; search the result range for a case name and confirm `font.italic === true` (or `font.underline === "Single"` if Step 3 chose underlining) while the reporter reads roman (`false` / `"None"`). Spot-check one treatise title and one statute the same way. Report the category counts and tell the user to press **Ctrl+A → F9** once in Word to settle page numbers after the pagination shift.
**Done when:** each TOA field's result lists its authorities with page numbers and the emphasis verifies in the rendered table.
