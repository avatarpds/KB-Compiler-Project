# v0.8.4 — Windows portability and report correctness

Paste the section below into the GitHub release description for `v0.8.4`.
It is kept here rather than at the repository root so it doesn't accumulate one
file per version; the permanent history lives in `CHANGELOG.md`.

---

## What changed

This release makes the tooling actually runnable on Windows, fixes three ways
the audit report could mislead you, and closes one case where the bootstrap
silently dropped documents from the index.

### If you are on Windows, update

Two defects made the scripts unusable outside an interactive console, which is
to say unusable for the CI/gate workflow the README describes.

- **The report crashed when redirected.** Windows defaults stdout to the locale
  encoding (`cp1252`) whenever it isn't a console, so
  `check_master_list.py "<base>" > report.txt` died with a
  `UnicodeEncodeError` on the first section title containing `→` — after two
  lines of output, and exiting `1`, which reads as *"problems found"* rather
  than *"crashed"*. Output is now pinned to UTF-8.
- **Non-interactive runs took the interactive path.** `isatty()` alone was not
  enough: on Windows, stdin redirected from `NUL` is reported as a *terminal*,
  so a scheduled task or pipeline hit a prompt nobody could answer. An
  immediate EOF is now treated the same as a non-interactive stdin. The test
  suite drives the scripts with a real empty pipe for the same reason — 7 of
  its 12 groups had been failing on Windows for this alone.

### Report correctness

- **A space where the tab belongs is now a tab problem.** `Name v1.0` used to
  be reported as a *name divergence*, and the recommendation told you to
  propagate `Name v1.0` into the file name and the title — actively the wrong
  fix. A trailing `v<n>` counts as a version only when the file name or the
  Version History agrees, so `Migracao Office v2` keeps its own name.
- **A header with no version at all is now reported.** It used to be invisible:
  the version comparison drops empty sources before comparing, so
  `{None, "1.0", "1.0"}` collapsed to a single value and said nothing —
  precisely where the "header must match the Version History" rule stops being
  enforceable.
- **An unreadable document counts once.** A file the index points at that
  exists but cannot be parsed as a Word document used to produce three
  findings: a broken reference claiming a file that plainly exists doesn't,
  plus a structure and a formatting violation. It now has its own heading.
- **An extra column in the Overview no longer breaks the tab mapping.**
  openpyxl pads every row out to the sheet's widest column, so a note typed
  past the "Tab" column hid the reference and brought back the false
  folder/tab mismatches on truncated tab names that the mapping exists to
  prevent.

### Bootstrap

- **Two folders sharing a category name no longer lose one of them.**
  `01 - Rede` and `02 - Rede` both strip to `Rede`; the scan assigned instead
  of merging, so every document of whichever folder was scanned first vanished
  from the index and then surfaced as an orphan in the checker's report — with
  nothing in the bootstrap's own output saying they had been dropped. They are
  now merged into one tab, and the merge is reported.
- **`--force` no longer indexes the index itself.** A custom-named index
  (`--output "Internal Index.xlsx"`) matches none of the built-in name patterns,
  and `.xlsx` is indexable.
- **The language prompt offers only the languages that exist.** It invited
  "another language — give its two-letter code", then crashed on
  `HEADERS[lang]` and surfaced it as a bare `ERROR: 'fr'`. Its confirmation
  message also named the Portuguese status label regardless of the language
  chosen.
- Overview count formulas now cover rows 3–10000 instead of 3–1000, which
  truncated silently past ~998 rows.

### Documentation

- `HOW-TO-USE.md` gained a **Prerequisites** section: what the plugin needs
  (nothing but Claude) versus what the scripts need (Python 3.7+, pip, the two
  pinned packages), why the versions are pinned, and the fact that auditing
  needs no write access. Plus Windows troubleshooting for the `python3`/`py`
  execution aliases and for Microsoft Store Python's `AppData\Roaming`
  redirection.
- `SKILL.md` said "when the target language isn't Portuguese" in two places
  while Section 10 declares English canonical. `--lang en|pt|...` advertised
  languages that do not exist. The README claimed three test groups.

### Internals

- Each document is opened once per row and shared across every check, instead
  of being parsed four times.
- New regression groups: colliding category names, an extra Overview column,
  `--force` never indexing the index. The ambiguous-reference detection gained
  a positive case — it had only ever been asserted as zero.
- The detections fixture is built once up front, so a broken builder fails
  loudly instead of showing up as a dozen misleading `FileNotFoundError`s.

## Verification

15 test groups, 13 detection sections, 33 fixture problems, 3 unverifiable —
plus 10 release scenarios that each reproduce one of the defects above. Both
suites exit `0` on Windows 11 / CPython 3.13.14 with **no encoding environment
variables set and a `cp1252` locale**, which is the default condition under
which the previous release crashed.

The captured output is in [`test-runs/0.8.4/`](../../test-runs/0.8.4/). To
re-run it yourself:

```
pip install -r skills/kb-compiler/scripts/requirements.txt
python3 skills/kb-compiler/tests/run_tests.py
python3 test-runs/0.8.4/probes/verify_fixes.py
```

## Note on versioning

The previous commit was tagged `v0.8.3` but never bumped
`plugin.json`/`marketplace.json`, so it shipped as `0.8.2` and installs of
0.8.2 saw no update available. 0.8.4 bumps both; its changelog entry covers
that release's documentation changes as well.

## Upgrading

Marketplace:

```
/plugin marketplace update kb-compiler
```

Or download the release zip and upload it in the desktop app under
Settings → Plugins → Add, which works regardless of managed marketplace policy.
