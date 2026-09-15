# Changelog

All notable changes to this project are documented here.

## 0.9.1

Found by auditing a real 60-document base (63 indexed rows across six category
tabs) and then probing each suspicion on a synthetic one. Nothing here was
reproducible on the fixture as it stood, and the suite was green throughout —
every one of these is a branch the fixture never walked.

- **A skipped sheet is now a finding, not a footnote.** A tab with no
  recognizable `Code` column printed a `WARNING` and contributed nothing to the
  count, so a base whose sheet had lost its header row reported
  `SUMMARY: 0 problem(s) found` and exited `0` — the documented "clean" signal
  a CI gate keys on — while an entire sheet went unchecked. Whether the exit
  code happened to be non-zero at all depended on orphan detection incidentally
  firing for the same documents.

- **Bullets no longer count as a shared numbering sequence.** The check
  collected every direct `w:numId` without consulting `numFmt`, so two
  *bulleted* sections sharing one definition were reported as "the second
  continues the first's count" — meaningless for bullets, which have no
  counter, and something Word does routinely because it reuses one definition
  for identical bullet formatting. On the real base this was **9 of 19
  findings from that check, every one of them false**. The ordered case still
  fires; an unresolvable definition is left alone, as the rest of the
  formatting checks already do.

  The fixture's own case for this was unfaithful: it pinned two `List Number`
  paragraphs to `numbering.findall(w:num)[0]`, which is the *bullet*
  definition created moments earlier by the document's first bulleted
  prerequisite. Those paragraphs rendered as bullets, so the document never
  modeled the defect — the test passed only because the checker ignored
  `numFmt`. It now resolves the decimal definition explicitly, and a bulleted
  counter-case sits beside it.

- **A row with a blank `Code` is reported.** It used to be skipped whole-cloth:
  nothing named the row, and the document it indexes then resurfaced as an
  "orphaned file", sending the reader after the file instead of the row that is
  actually wrong. Section 2's `—` marker exists precisely so a document without
  a code stays visible in the index. Reported once — the row still references
  its file, so it is no longer also an orphan. A whitespace-only cell counts as
  blank, since that is how it reads to everyone.

- **Office lock files and OS junk are never orphans.** Auditing a base while a
  single document was open in Word reported that document's `~$…` lock file as
  an orphaned file — a finding that vanishes on its own, which is the fastest
  way to make a report look untrustworthy. `~$…`, `.~lock.…#`, `Thumbs.db`,
  `.DS_Store` and `desktop.ini` are now invisible to the audit. `.gitignore`
  had excluded them from the repository all along; the checker had not.

- **A `.docx` named like the index is still checked.** `is_master_list_filename`
  matched on the name prefix with no extension check, so a document called
  `Master List Guidelines.docx` was taken for the index itself and excluded
  from orphan detection — an unindexed document that never appeared in the
  report at all. Now the extension is required, and the prefix is compared
  case-insensitively.

- **The name comparison normalizes all four sources.** Every *file-name*
  comparison already went through `nfc()`, but the four-source name check
  normalized only the file side — so an accented name authored NFD (a document
  that passed through macOS) and recorded NFC in the index was reported as a
  name divergence between two strings that are the same text, with a
  recommendation to "fix" one of them.

- **The footer's page number must be a real field.** `has_page_field` tested
  the footer XML for the substring `PAGE`, so a footer reading
  `ACME HOMEPAGE | Knowledge Base` was credited with an automatic page field it
  did not have and its typed-in page number went unreported. Now the field
  instructions are inspected, on a word boundary — ` NUMPAGES ` contains
  `PAGE`, and a total-pages field is not a current-page number.

- **The Overview tab is recognized in any casing.** Matching its literal
  spelling meant a tab renamed `OVERVIEW` stopped being recognized, was read as
  a category sheet, found no `Code` column and was skipped — silently losing
  the tab-to-category map that exists to prevent false folder/tab mismatches.

- **A half-finished `languages.json` fails loudly.** Adding a language is
  documented as a data edit to that file, but the module-level tables index
  some keys directly, and a missing one raised a bare `KeyError` at *import*
  time — before `main()` exists to catch it — which exits `1`. That is this
  tool's own "problems found" code, so a broken vocabulary was
  indistinguishable from a completed audit to anything reading the exit status.
  The required keys are validated where the file is loaded, with the graceful
  error and exit `2` that path already promised.

Four more, found while remediating that base rather than only auditing it:

- **Direct numbering beats the style name.** `_list_format` read the style name
  first, so a paragraph styled `List Number` whose own `w:numPr` points at a
  bullet definition was called decimal — while Word renders it as a bullet. The
  three sources are now consulted in Word's order of precedence: the
  paragraph's numbering, then its style's, then the style name as a fallback
  for the built-in list styles. This is how the fixture's shared-numbering case
  hid its own unfaithfulness for two releases.

- **Ambiguous dates are settled by the base, and disclosed.** `parse_date_value`
  tried `%d/%m/%Y` unconditionally, so an English base's `03/04/2026` read as
  3 April. It now takes the order from the language recorded in
  `.kb-compiler.json`, defaulting to day-first so every existing base reads
  exactly as before, and `date_is_ambiguous()` marks the dates where the
  convention actually changes the answer. The name-divergence recommendation
  states the reading it used instead of presenting one as fact.

- **Numbered section headings are recognized.** Section matching is
  `startswith()` against the bare vocabulary term, so a heading reading
  `3. Pré-requisitos` was invisible: the structure check called the section
  missing, and — silently — the per-section list rule skipped it, leaving a
  numbered document's Step by Step never verified to be a numbered list.
  Numbering sections is not something the standard forbids.

- **An em dash inside a document's own name is allowed.** Section 4
  constrains the separator *between the code and the name*: "Use a regular
  hyphen `-`, never an en/em dash, between the code and the name." The check
  searched the entire header, so
  `KB-004 - Alterar Data de Expiração — Contingent Worker no SuccessFactors`
  was reported as using the wrong separator — which reads as an instruction to
  strip punctuation out of your own title. Only the code-to-name separator is
  examined now, on a header that actually carries a code. Two false positives
  on the real base.

- **A superseded document may stay on its topical tab.** Section 7 offers
  legacy documents two homes in the index: their own Legacy tab, or their
  original topical tab with Status `Legacy`. The folder/tab check only
  tolerated the first, and reported the second as a mismatch telling the reader
  to "move the file or move the row" — undoing an arrangement the standard had
  offered them. The checker contradicted the spec it enforces.

The suite grew from 16 groups to 28. Every new group was confirmed to fail
against 0.9.0's checker before being kept: two of them are the reason the
`Formatting violations` count alone could not have caught this release's
defects, because the bullet false positive (+1) and the missed typed page
number (-1) cancelled out exactly.

Two rules were added to the standard itself. Both are mistakes made while
remediating that base, and neither was something SKILL.md did anything to
prevent:

- **A new Version History row must inherit the formatting of the rows already
  there** (Section 3.8). Font, size and colour are set on the runs rather than
  on the table style, so a row added programmatically states nothing and falls
  back to the document's defaults — rendering larger and in a different colour
  than every row above it. It is immediately visible to a reader and the audit
  does not check it. Cell borders and shading live on the cell, so they carry
  over on their own.

- **Remediate each document exactly once** (Section 8.8). Applying a batch in
  more than one sweep — formatting first, missing sections afterwards — gives
  every document caught by both two version bumps and two near-identical
  Version History rows, which is the outcome the batching rule in 8.9 exists to
  avoid. It affected 27 documents before being caught and repaired.

Measured, on an identical base state. A copy of the base was taken before any
document was edited, so both checker versions could run against byte-identical
input:

| Run | Findings |
|---|---|
| 0.9.0's checker, base untouched | 252 |
| 0.9.1's checker, **same** untouched base | 238 |
| 0.9.1's checker, after remediating the base | 21 |

The first gap — **14 findings** — is this release's false positives: things that
were never defects. The second is real repair. Nothing that was a true positive
stopped being reported: every disappearance was checked against the previous
report, category by category. The evidence is in
[`test-runs/0.9.1/`](test-runs/0.9.1/).

## 0.9.0

Continuous integration, and two rules the standard has always stated that
nothing enforced.

- **CI on a matrix of operating systems.** `.github/workflows/tests.yml` runs
  the suite on Ubuntu and Windows, on Python 3.9 and 3.13. Both defects fixed
  in 0.8.4 were Windows-only and neither was reproducible on Linux, so a
  single-platform check would have missed them exactly as a human did. The
  workflow also audits the fixture twice — once to a console, once redirected
  to a file — and fails if anything reaches stderr or the report has no
  `SUMMARY` line, because a crash halfway through the report still exits with
  the code that means "problems found".

- **Status values are validated** (Section 7). Only `Active`, `In review` and
  `Legacy` are legal, in the spellings the tooling recognizes. A value it
  doesn't know breaks things silently instead of loudly: the Overview's COUNTIF
  formulas stop counting the row, and a misspelled legacy label lets a
  superseded document fall through to the structure checks it should be exempt
  from. `Em Revisao` without its cedilla is the case that motivated this.
  `LEGACY_STATUS_LABELS` is now derived from the same table rather than
  repeated beside it.

- **The Version History is checked as a sequence** (Section 3), not only by its
  last row. The first row must record the creation at 1.0, and versions must
  only move forward; a history that goes 1.0 → 1.2 → 1.1 is the signature of a
  botched edit, and one that never starts at 1.0 has lost its creation row.
  Versions are compared as integer tuples, so 1.9 → 1.10 is correctly read as
  moving forward — as text it looks like a regression, and the fixture carries
  that case as a negative control. Both checks are skipped for Legacy-status
  documents, on the same reasoning as the structure checks.

Found by running against a real 68-document base for the first time — none of
these reproduce on the synthetic fixture:

- **`paragraph.style` can be `None`**, and reading `.name` off it aborted the
  entire audit. python-docx returns None whenever a paragraph points at a style
  the document's styles part doesn't define, which is routine in documents
  converted from .doc or carrying pasted-in content: 15% of the real base hit
  it. A style-less paragraph is now treated as having no style name, which is
  the same outcome as one styled Normal.
- **A crash exited 1**, and this tool's contract reads 1 as "problems found" —
  so any wrapper, CI job or batch-edit gate keying on the exit code took a
  traceback for a completed audit. Unexpected failures now exit 2.
- **Blank rows at the bottom of a Version History** made the current version
  read as `""`. Since the comparison drops empty sources before comparing,
  `{header, "", index}` collapsed to a single value and the version-divergence
  check silently stopped running for that document. The last FILLED row is now
  what counts.

- **Prerequisites and Verification may be numbered, not only bulleted**, and a
  new check catches **two sections sharing one numbering sequence**. Both came
  out of a real base: it used numbered lists essentially everywhere, so the
  "bulleted" rule described a practice nobody followed, and 13 of its documents
  had sections silently continuing each other's count — a defect invisible on
  screen, since Word shows numbers either way, just the wrong ones. The heading
  of a section is always a heading; only its items are a list.

Interactivity is detected properly, instead of being guessed:

- **The non-interactive path no longer contradicts itself.** `isatty()` is not
  enough on Windows: stdin redirected from `NUL` is a character device, so it
  reports a TERMINAL. A scheduled task printed a numbered menu to nobody and
  only learned the truth when `input()` raised EOFError — with the menu already
  on screen, immediately followed by "running non-interactively". A real
  console answers `GetConsoleMode` and `NUL` does not, despite both being
  character devices, so that call is now the discriminator. Verified: under
  `NUL`, `isatty()` says True while the new check says False. The EOFError
  handlers stay as a second line of defence, and the suite now asserts the menu
  is never printed to a non-interactive run.

Language, and the standard's "any language" promise:

- **Every localized term moved out of the code into
  `scripts/languages.json`.** Eight separate tables hardcoded English and
  Portuguese, so Section 10's "a base may be authored in any language" really
  meant "either of two, and only by editing Python". The checker now unions the
  recognition terms of every language in that file, and the bootstrap writes
  using one of them. Adding a language is a data edit.
- **`--lang` is optional**: the language is detected by counting how many
  section headings match each language's vocabulary, and the scores are
  printed. A tie or an empty base is reported as undecided rather than guessed.
  An explicit `--lang` still wins.
- Status values are the one exception to accent tolerance: they must match the
  written spelling exactly, because the Overview's COUNTIF formulas do, and a
  row reading `Em revisao` stops being counted no matter what the checker says.

Also: the fixture's KB-017 gained a proper creation row so it exercises version
divergence and nothing else — with a single 2.1 row it would now trip the new
sequence rule too, putting two deliberate defects in one document. 16 test
groups, 15 detection sections, 38 fixture problems.

Documentation: the documented Python floor moved from "3.7 or newer" (inferred
from the language features used) to "3.9 or newer" (what CI actually
exercises), and HOW-TO-USE's own table of checks had fallen behind the
README's.

## 0.8.4

Windows portability, three report-correctness fixes, and one case of silent
data loss in the bootstrap. The previous release was tagged `v0.8.3` in git but
never bumped `plugin.json`/`marketplace.json`, so it shipped as 0.8.2; this
entry covers both.

Portability — the tooling could not actually be run or tested on Windows:

- The report is pinned to UTF-8. On Windows, stdout defaults to the locale
  encoding (cp1252) whenever it is not a console, so `check_master_list.py
  "<base>" > report.txt` died with a `UnicodeEncodeError` on the first section
  title containing `→` — after two lines of output, and exiting `1`, which
  reads as "problems found" rather than "crashed". That is the exact CI/gate
  usage the README advertises.
- An immediate EOF on stdin is now treated as "nobody there". `isatty()` alone
  was not enough: on Windows, stdin redirected from `NUL` is reported as a
  terminal, so non-interactive runs took the interactive path. The test suite
  drives the scripts with a real empty pipe for the same reason — 7 of its 12
  groups had been failing on Windows for this alone.

Report correctness:

- A header with a space where the tab belongs (`Name v1.0`) is reported as a
  missing tab separator instead of a NAME divergence, whose recommendation
  told you to propagate `Name v1.0` into the file name and the title. A
  trailing `v<n>` only counts as a version when the file name or the Version
  History agrees, so `Migracao Office v2` keeps its own name.
- A header with no version at all is now reported. It used to be invisible:
  the version comparison drops empty sources before comparing, so
  `{None, "1.0", "1.0"}` collapsed to one value and said nothing — precisely
  where the "header must match the Version History" rule stops being
  enforceable.
- A file that exists but cannot be parsed as a Word document is reported once,
  under its own heading. It used to produce three findings: a broken reference
  claiming a file that plainly exists doesn't, plus a structure and a
  formatting violation.
- The Overview's category→tab mapping is read by scanning the row's cells
  rather than trusting its last one. openpyxl pads rows to the sheet's widest
  column, so a note typed past the "Tab" column hid the reference and brought
  back the false folder/tab mismatches on truncated tab names that the mapping
  exists to prevent.

Bootstrap:

- Two folders that strip to the same category name (`01 - Rede`, `02 - Rede`)
  no longer lose one of them. The scan assigned instead of merging, so every
  document of whichever folder was scanned first vanished from the index and
  then surfaced as an orphan in the checker's report, with nothing in the
  bootstrap's own output saying they had been dropped. They are now merged into
  one tab and the merge is reported.
- A `--force` re-run no longer indexes the master index itself. A custom-named
  index (`--output "Internal Index.xlsx"`) matches none of the built-in name
  patterns, and `.xlsx` is indexable.
- The interactive language prompt offers only the languages that exist. It
  invited "another language — give its two-letter code", then crashed on
  `HEADERS[lang]` and surfaced it as a bare `ERROR: 'fr'`. Its confirmation
  message also named the Portuguese status label regardless of the language
  chosen.
- The Overview's count formulas cover rows 3–10000 instead of 3–1000, which
  truncated silently past ~998 rows.

Internals and tests:

- Each document is opened once per row and shared across every check, instead
  of being parsed four times.
- New regression groups: colliding category names, an extra Overview column,
  and `--force` never indexing the index. The ambiguous-reference detection
  now has a positive case — it had only ever been asserted as zero. 15 groups,
  13 detection sections, 33 fixture problems.
- The detections fixture is built once up front, so a broken builder fails
  loudly instead of showing up as a dozen misleading `FileNotFoundError`s.
- Documentation: `HOW-TO-USE.md` gained a Prerequisites section (what the
  plugin needs versus what the scripts need, why the package versions are
  pinned, and Windows-specific troubleshooting). `SKILL.md` said "when the
  target language isn't Portuguese" in two places while Section 10 declares
  English canonical; `--lang en|pt|...` advertised languages that do not exist;
  the README claimed three test groups.

## 0.8.2

- Added `.claude-plugin/marketplace.json`, so the repository is both the plugin
  and its own single-plugin catalog. Installing is now two commands rather than
  a manual copy of the skill folder:

  ```
  /plugin marketplace add avatarpds/KB-Compiler-Project
  /plugin install kb-compiler@kb-compiler
  ```

  Third-party marketplaces have auto-update disabled by default, so run
  `/plugin marketplace update kb-compiler` to pick up later versions. Note that
  a managed marketplace policy can block this path entirely; uploading the
  release zip in the desktop app works regardless.

- Added `HOW-TO-USE.md`: a walkthrough covering installation, base setup,
  creating and ingesting documents, reading an audit report, and
  troubleshooting.

## 0.8.1

- Duplicate file names across folders are resolved using the row's category tab
  instead of taking the first match, which previously meant comparing a
  document the row didn't refer to. When the tab can't single one out, it is
  reported as an ambiguous reference and nothing about that document is
  checked.
- Orphan detection now records the resolved path rather than the bare file
  name, so a same-named file in another folder no longer escapes detection by
  borrowing another row's reference.

## 0.8.0

- Section 4 coverage completed: page orientation, footer presence, footer tab
  separator and automatic page fields, hyphen vs. en/em dash in the header,
  bullet/numbered list formatting per section, and all-caps author names in the
  Version History.
- Footer labels are compared across the base: every document must carry the
  same organization label.
- List formatting is resolved through built-in style names and the style chain,
  not only direct paragraph numbering — documents using Word's native
  List Bullet / List Number styles were being reported as not lists at all.
- Code-prefix stripping accepts en and em dashes, so a header using the wrong
  dash is reported once as a dash violation instead of also surfacing as a
  bogus name divergence.

## 0.7.1

- Documents sitting loose in the base root are indexed under an `Uncategorized`
  tab and individually flagged, instead of being skipped by the bootstrap and
  then reported as orphans by the checker.

## 0.7.0

- The bootstrap asks where the master index should live (default name in the
  base, another name in the base, or any path including outside it), or takes
  `--output`. The choice, the base language and the default owner are recorded
  in a `.kb-compiler.json` at the base root.
- The checker resolves the index from that configuration before falling back to
  file-name patterns, so a custom name or an out-of-base location keeps working
  as a continuous insertion point. A configured path that has gone missing is an
  explicit error, never a silent fallback.
- The interactive path asks which language the base is authored in, instead of
  defaulting to English and producing an index whose headers don't match the
  documents it indexes.

## 0.6.1

- The bootstrap refuses to run when the base already has a master index in any
  language or under any configured path. Checking only its own target file name
  meant a run without `--lang` quietly created a second index beside the first.

## 0.6.0

- English is the standard's canonical language. Literal terms are defined once
  in English; a base may be authored in any language confirmed at setup, and the
  scripts recognize equivalent terms when reading an existing base.
- Folder/tab comparison reads the Category-to-Tab mapping recorded in the
  Overview sheet instead of recomputing the transformation, which had made the
  checker report that a folder didn't match itself whenever a category name had
  been sanitized or truncated to become a legal sheet title.
- Structure checks skip documents whose Status marks them as legacy.
- Overview count formulas use a wide range so an appended row is included.

## 0.5.0

- Section 3 structure validation: required sections present, in the mandated
  order, with Version History last. An extra custom section is allowed.
- Folder/tab mismatch detection.
- Fixes found by external review: accumulate spreadsheet candidates across all
  naming patterns before deciding; read the title from the first non-empty
  paragraph; only treat a header as glued when the version is attached directly
  to a word character; sanitize Excel sheet names and disambiguate collisions;
  recognize a numbered legacy folder; reject an unknown `--lang`; return exit
  code 2 for usage and dependency errors.

## 0.4.1

- MIT license.

## 0.4.0

- Test suite covering detections, bootstrap behaviour and non-interactive
  safety.

## 0.3.0

- `bootstrap_master_list.py`: generates a master index for a base that has none,
  filling in only what it can read, marking every row unreviewed, and never
  assigning a KB code to a document that doesn't already have one.
- A missing index no longer ends in a traceback: the checker explains the
  options and, on an interactive terminal, offers to create one. Run
  non-interactively it prints the same guidance and exits rather than blocking
  on input.

## 0.2.0

- `check_master_list.py`: read-only audit covering broken references, orphaned
  files, duplicate codes, name and version divergence across four sources per
  document, legacy Version History tables, headers missing the tab separator,
  and the Section 4 formatting constants. Conservative mode: a formatting
  property the run/style cascade doesn't state is listed as unverifiable rather
  than reported as a violation.
- Pinned dependency versions.
- Ingestion of raw drafts and other formats into the standard.

## 0.1.0

- Initial documentation standard: section structure, exact Word formatting,
  global KB-XXX numbering with uniqueness constraints, and the master-index
  spreadsheet convention.
