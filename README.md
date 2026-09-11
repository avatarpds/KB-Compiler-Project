# KB-Compiler

A documentation standard for knowledge bases (KB), packaged as a Cowork/Claude plugin, with a deterministic validation script and a test suite.

## What it does

Once installed, Claude knows a documentation standard for KB articles/runbooks:

- Section structure (Purpose, Prerequisites, Step by Step, Verification, Version History, etc.)
- Exact Word formatting (title, header, footer, styles, warning callouts)
- Naming and `KB-XXX` numbering conventions, with a global uniqueness constraint and no code recycling
- How to create and keep a master-index spreadsheet up to date and free of broken references
- How to turn a raw draft or a document in another format (PDF, PPTX, Outlook `.msg`, Markdown, an informally-shaped `.docx`) into a document that follows the standard — extracting content, order, and images instead of building from scratch through Q&A

The plugin isn't tied to any specific company, folder, or path. The first time it's used on a new base, it asks (or confirms from the conversation's context) what the base's root directory is, what organization/team label should appear in documents' footers, and — separately — what language the actual KB documents should be written in. It reuses those answers for every later document in the same base.

The skill's own instructions are written in English, so Claude follows them precisely and any team can install and read this plugin regardless of what language they work in day to day. **The documents it produces are a separate choice**: they're written in whatever target language the base's author confirms (the standard's canonical terms are English; a base authored in another language is confirmed once at setup and stays consistent, and the scripts recognize equivalent terms when reading it).

This helps anyone on the team create, review, or fix KB documents already following the same standard — without memorizing the formatting rules or manually checking every consistency detail.

## How to use it

Once installed, just ask normally, for example:

- "Create a new KB about [process X] in folder [directory]"
- "Review KB-0XX to check it follows the standard"
- "Bump this document's version and update the master-index spreadsheet"
- "Turn this PDF/PPTX/email/draft into a KB document"

Note: for ingestion from a raw draft or another format, the skill still drafts the standardized document and the proposed master-index row for you to confirm before anything is saved or registered — it doesn't assign a `KB-XXX` code or write to the spreadsheet unsupervised.

The skill triggers automatically whenever the conversation involves creating, updating, or reviewing a KB article/runbook for a knowledge base.

## Validating a base

`scripts/check_master_list.py` cross-checks a whole base against the standard. It is a **fixed** script rather than one generated on demand, so an audit runs the same way every time.

```bash
pip install -r skills/kb-compiler/scripts/requirements.txt
python3 skills/kb-compiler/scripts/check_master_list.py "<base path>"
```

What it reports:

| Check | Detects |
|---|---|
| Broken references | A spreadsheet row pointing at a file that doesn't exist |
| Orphaned files | A document with no row in the spreadsheet |
| Ambiguous file reference | A bare file name that exists in several folders, where the row's tab doesn't single one out |
| Duplicate codes | The same `KB-XXX` on two different rows |
| Name divergence | File name, title, header and spreadsheet disagreeing |
| Version divergence | Header, Version History and spreadsheet disagreeing |
| Legacy version table | A Version History still in the old column format |
| Header missing tab | Name and version glued together (`Namev1.2`) |
| Structure violations | A required Section 3 heading missing, out of order, or content after Version History |
| Folder/tab mismatch | A row on one category tab whose file physically lives in another folder |
| Formatting violations | Margins, page orientation, title size/weight/color, non-native heading styles, callout shading and text color, footer presence and page fields, hyphen vs. dash in the header, bullet/numbered list formatting per section, all-caps author names, and a footer label that differs across the base |

Two design decisions worth knowing:

**It reports, it never fixes.** For a name divergence it recommends which source to keep, based on which one changed most recently, and lists every variant found — but the choice and the edit stay with you. Its report also reminds you that resolving a name divergence is a content edit, and therefore requires a version bump and a Version History row.

**Formatting checks are conservative.** Word formatting cascades from run to paragraph style to base style, so the script resolves that chain before judging anything. When no level of the chain states a property, the effective value comes from Word's theme defaults, which can't be read reliably — those cases are listed separately as "could not be verified" and are never counted as problems. The tradeoff is deliberate: a report full of false positives is one nobody trusts.

Exit codes are distinct so a wrapper can tell the cases apart: `0` clean, `1` problems found, `2` couldn't run at all (no spreadsheet, bad path, aborted). That makes it usable as a gate for a batch edit or a CI-style check.

## Bootstrapping a base that has no spreadsheet yet

If a base doesn't have a master-index spreadsheet, the checker doesn't die with an error — it says so and, when run in a terminal, offers to create one by scanning the base or to be pointed at an existing file under a different name. Run non-interactively (CI, cron, a pipe) it prints the same guidance and exits `2` instead of blocking on input that will never come.

To generate the spreadsheet directly:

```bash
python3 skills/kb-compiler/scripts/bootstrap_master_list.py "<base path>" \
    [--lang en|pt|...] [--owner "Name"] [--output PATH.xlsx] [--force]
```

Run in a terminal, it asks where the index should live: the default name inside the base, a different name inside the base, or a full path of your choosing (it may sit outside the base entirely). `--output` makes the same choice non-interactively. Whatever you pick is recorded in a small `.kb-compiler.json` in the base, so **every later run uses that same file** — the index becomes a continuous insertion point rather than something re-derived from a file-name pattern each time. A configured path that has gone missing is an explicit error, never a silent fallback to some other spreadsheet.

It is deliberately conservative, because a base that needs bootstrapping is by definition one that doesn't follow the standard yet:

- fills in only what it can actually read (code and name from the file name, version from the header when the header is in the standard's shape);
- leaves blank whatever it couldn't read, and reports each one, instead of inventing a value;
- writes every row as `In review`, never `Active` — no row has been reviewed by a human yet;
- **never assigns a KB code** to a document that doesn't already have one, since codes are never recycled and a wrong guess is permanent;
- indexes documents sitting loose in the base root under an `Uncategorized` tab and flags each one, rather than skipping them and letting the checker report them as orphans;
- refuses to run at all if the base already has a master index, in any language or under any configured path — a base has exactly one.

Creation lives in this separate script on purpose: `check_master_list.py` stays read-only, so it can be run at any time without wondering what it might touch.

Note the overlap with the skill itself: Section 7 of the standard already has Claude create the master index conversationally, one document at a time. Use that when starting a base from scratch; use this script when you have a folder full of existing documents and want them indexed in one pass.

## Tests

```bash
pip install -r skills/kb-compiler/scripts/requirements.txt
python3 skills/kb-compiler/tests/run_tests.py
```

Three groups:

**Detections.** `build_fixture.py` builds a synthetic base with one deliberate instance of every problem the checker catches, plus cases that must stay clean: a fully conforming document; one whose formatting is inherited from its style rather than set on the run; attachment files that must never be flagged as orphans; an accented file name stored in NFD form. The suite compares every report section's count against the expected value.

**Bootstrap.** Verifies the generated spreadsheet is one the checker reads with zero broken references and zero orphans, that no row is written as reviewed, that no KB code is ever invented for a document whose file name doesn't already have one, and that a second run refuses to overwrite.

**Non-interactive safety.** Verifies that a base with no spreadsheet prints guidance and exits `2` instead of blocking on input, and never falls back to a raw traceback.

Run this after any change to either script. A detection that silently stops working is the worst failure mode for an audit tool: the report comes back clean and you believe it. This suite has already caught regressions that manual testing missed, and grew twice after external reviews found bugs it didn't cover — including a fix that had turned a crash into a silent false positive. Every one of those cases now has a test.

## Requirements

Python 3, plus the pinned versions in `scripts/requirements.txt` (`python-docx`, `openpyxl`). Versions are pinned because table and cell reading behavior changes between releases.

## Maintenance

If the documentation standard changes (a new section, a new naming convention, a new formatting rule), update `skills/kb-compiler/SKILL.md`, mirror any formatting constant in the constants block at the top of `check_master_list.py`, add a fixture case for it, bump the version in `.claude-plugin/plugin.json`, and distribute a new version of the plugin to the team.

## License

MIT — see `LICENSE`.

Maintained by Lucas Souza.
