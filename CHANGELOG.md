# Changelog

All notable changes to this project are documented here.

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
