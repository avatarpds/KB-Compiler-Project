# v0.9.1 — Fourteen fixes from remediating a real base, not just auditing one

0.9.0 audited a real knowledge base for the first time. 0.9.1 came from
**fixing** one. Working through 252 findings on a 60-document base surfaced a
different class of defect than reading a report does: half of what this release
fixes is the checker being *wrong*, not the base.

Every fix has a regression case, and **every new test group was confirmed to
fail against 0.9.0's checker before being kept** — a regression test that passes
before the fix is testing nothing. The suite went from 16 groups to 28.

## Measured, on an identical base state

A copy of the base was taken before any document was edited, so the two checker
versions could be run against exactly the same input:

| Run | Findings |
|---|---|
| 0.9.0's checker, base untouched | 252 |
| 0.9.1's checker, **same** untouched base | 238 |
| 0.9.1's checker, after remediating the base | 21 |

The first gap — **14 findings** — is this release's false positives: things that
were never defects. The second is real repair. Nothing that was a true positive
stopped being reported; every disappearance was checked category by category
against the previous report.

## A clean report that wasn't

These are the ones worth reading twice, because in each case the tool told you
everything was fine when it wasn't.

- **A skipped sheet reported a clean base.** A tab with no recognizable `Code`
  column printed a `WARNING`, contributed nothing to the count, and let the run
  finish with `SUMMARY: 0 problem(s) found` and exit `0` — the documented
  "clean" signal a CI gate keys on — while an entire sheet went unchecked.
  Whether you got a non-zero exit at all depended on orphan detection
  incidentally firing for the same documents.

- **A broken vocabulary exited `1`.** Adding a language is documented as a data
  edit to `languages.json`, but the module-level tables index some keys
  directly, and a missing one raised a bare `KeyError` at *import* time — before
  `main()` exists to catch it. Python exits `1` on an uncaught exception, which
  is this tool's own "problems found" code, so a half-finished language was
  indistinguishable from a completed audit to anything reading the exit status.
  The required keys are now validated where the file is loaded, with the
  graceful error and exit `2` that path already promised.

- **A row with a blank `Code` vanished.** It was skipped whole-cloth: nothing
  named the row, and the document it indexes resurfaced as an "orphaned file",
  sending the reader after the file instead of the row that was actually wrong.
  Section 2's `—` marker exists precisely so a codeless document stays visible.
  Now reported once, and no longer also an orphan.

- **A document named like the index was never checked.**
  `is_master_list_filename` matched on the name prefix with no extension check,
  so `Master List Guidelines.docx` was taken for the index itself and excluded
  from orphan detection — an unindexed document that appeared nowhere in the
  report.

- **A typed page number passed for a Word field.** `has_page_field` searched the
  footer XML for the substring `PAGE`, so a footer reading
  `ACME HOMEPAGE | Knowledge Base` was credited with an automatic page field it
  did not have. The field instructions are now inspected, on a word boundary —
  `NUMPAGES` contains `PAGE`, and a total-pages field is not a current-page
  number.

## Reports nobody should trust

False positives, all of them found by fixing a base rather than reading about
one. Nine of the first alone.

- **Bullets counted as a shared numbering sequence.** The check collected every
  direct `w:numId` without consulting `numFmt`, so two *bulleted* sections
  sharing one definition were reported as "the second continues the first's
  count" — meaningless for bullets, and something Word does routinely because it
  reuses one definition for identical bullet formatting. On the real base this
  was **9 of 19 findings from that check, every one of them false.**

- **An em dash inside a document's own name.** Section 4 constrains the
  separator *between the code and the name*; the check searched the whole
  header, so `KB-004 - Alterar Data de Expiração — Contingent Worker no
  SuccessFactors` was reported as using the wrong separator. That reads as an
  instruction to strip punctuation out of your own title.

- **Section 7's own arrangement, reported as a mismatch.** Section 7 gives
  legacy documents two homes: their own Legacy tab, or their original topical
  tab with Status `Legacy`. The folder/tab check only tolerated the first, and
  advised the reader to "move the file or move the row" — undoing an arrangement
  the standard had just offered them. The checker contradicted the spec it
  enforces.

- **Numbered section headings were invisible.** Matching is `startswith()`
  against the bare vocabulary term, so a heading reading `3. Pré-requisitos` was
  reported as a missing section — and, silently, the per-section list rule
  skipped it too, leaving a numbered document's Step by Step never verified to
  be a numbered list at all.

- **Two identical-looking names diverged.** Every *file-name* comparison already
  normalized to NFC, but the four-source name check normalized only the file
  side — so an accented name authored NFD and recorded NFC was reported as a
  divergence between two strings that are the same text.

- **Office lock files were orphans.** Auditing a base while one document was
  open in Word reported that document's `~$…` lock file as an orphaned file — a
  finding that disappears on its own. `.gitignore` had excluded them from the
  repository all along; the checker had not.

- **The Overview tab, renamed.** Matching its literal spelling meant a tab
  called `OVERVIEW` stopped being recognized, was read as a category sheet,
  found no `Code` column and was skipped — silently losing the tab-to-category
  map that exists to prevent false folder/tab mismatches.

## Correctness

- **Direct numbering now beats the style name.** `_list_format` read the style
  name first, so a paragraph styled `List Number` whose own `w:numPr` points at
  a bullet definition was called decimal — while Word renders a bullet. The
  three sources are consulted in Word's order of precedence: the paragraph's
  numbering, then its style's, then the style name as a fallback.

  This is how the suite's own shared-numbering fixture hid its unfaithfulness
  for two releases: it pinned two `List Number` paragraphs to the *bullet*
  definition created moments earlier by the document's first bulleted
  prerequisite, so those paragraphs rendered as bullets and modelled nothing.
  The test passed only because the checker ignored `numFmt`. It now resolves the
  decimal definition explicitly and fails loudly if none exists.

- **Ambiguous dates are settled by the base, and disclosed.**
  `parse_date_value` tried `%d/%m/%Y` unconditionally, so an English base's
  `03/04/2026` read as 3 April. The order now comes from the language recorded
  in `.kb-compiler.json`, defaulting to day-first so every existing base reads
  exactly as before, and `date_is_ambiguous()` marks the dates where the
  convention changes the answer. The name-divergence recommendation states the
  reading it used instead of presenting one as fact.

## Two rules added to the standard

Both are mistakes made while remediating the base, and neither was something
SKILL.md did anything to prevent:

- **A new Version History row must inherit the existing rows' formatting.** Font,
  size and colour are set on the runs rather than the table style, so a row
  added programmatically states nothing and falls back to the document defaults
  — rendering larger and in a different colour than every row above it. It is
  immediately obvious to a reader and the audit does not check it. Cell borders
  and shading live on the cell, so they carry over on their own.

- **Remediate each document exactly once.** Applying a batch in two sweeps
  (formatting first, missing sections afterwards) gives every document caught by
  both two version bumps and two near-identical history rows — the outcome the
  batching rule in Section 8.9 exists to avoid. It affected 27 documents before
  being caught and repaired.

## Known limitations

- `_list_format`'s numbering lookup re-walks the numbering part per call, which
  is O(paragraphs × definitions) per document. Not measured; not optimized.
- `HEADING_STYLE_MARKERS` matches by substring, so `heading 1` also matches
  `heading 10`–`heading 19`. Left deliberately: tightening it would stop
  recognizing the section headings of documents that use a mangled `Heading 10`
  style, and reporting "no recognizable structure" for them is a worse failure
  than the over-match.
- The CI matrix covers Python 3.9 and 3.13. This run was on 3.14, where the
  pinned dependencies install and pass; 3.9 is unverified locally.

## Upgrading

Nothing to migrate. `check_master_list.py` is still read-only, exit codes are
unchanged (`0` clean, `1` problems found, `2` couldn't run), and no report
section was removed — two were added: **Sheets skipped entirely** and **Row with
an empty Code cell**.

Expect your finding count to **drop** on a base that has bulleted sections
sharing a numbering definition, legacy documents on topical tabs, or names
containing an em dash. Expect it to **rise** if a sheet in your index has lost
its header row, or a row has a blank `Code` — both were previously invisible.
