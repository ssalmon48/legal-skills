# Legal Skills for Claude

Two Claude skills for building a real, auto-updating Table of Authorities in Microsoft Word briefs.

| Skill | What it does |
|---|---|
| `table-of-authorities` | Finds every authority in the brief, marks each citation with a Word TA field, and inserts category-split TOA fields (Cases, Statutes, Rules, Regulations, Constitutional Provisions, Treatises, Other Authorities) with live page numbers and "passim." |
| `toa-formatting` | Keeps long entries from splitting their page-number lists across lines, and optionally emphasizes statute and rule identifiers when a court or house style calls for it. |

Install both. They're designed to work together.

## Formatting conventions

The skills follow Bluebook Bluepages conventions for court documents:

- Case names (including "v." and procedural phrases like "In re") are italicized. Reporters, courts, and years are roman.
- Book, treatise, and article titles are italicized. Authors and journal names are roman.
- Statutes, regulations, rules, and constitutional provisions are roman by default.
- If your brief consistently underlines case names instead, the table will match the brief.

## Requirements

- Claude with code execution and file creation turned on.
- Microsoft Word, with Claude working in the open document.

## Install in the Claude app

1. On this page, click **Code > Download ZIP** and unzip it.
2. Zip each skill folder on its own (`table-of-authorities` and `toa-formatting`). The `SKILL.md` file must sit directly inside the folder you zip.
3. In Claude, go to **Customize > Skills**, click **+**, then **Create skill > Upload a skill**, and upload each ZIP.
4. Make sure both skills are toggled on.

## Install in Claude Code

```bash
git clone https://github.com/ssalmon48/legal-skills.git
cp -r legal-skills/table-of-authorities legal-skills/toa-formatting ~/.claude/skills/
```

## Use

Open your brief and ask Claude to "build a table of authorities." Afterward, press **Ctrl+A, then F9** in Word to refresh page numbers.
