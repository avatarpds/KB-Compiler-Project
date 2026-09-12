# v0.9.0 — Tested against a real base for the first time

Paste the section below the line into the GitHub release description for
`v0.9.0`. The permanent history lives in `CHANGELOG.md`.

---

## What changed

Until now this tooling had only ever been run against its own synthetic
fixture. This release is what came back from pointing it at a real
68-document knowledge base: three ways it could fail silently, and two rules
the written standard got wrong about how documents are actually authored.

It also gains continuous integration, so the next release is not the first
time anyone finds out whether it works on another platform.

### Three silent failures the fixture could never produce

- **`paragraph.style` can be `None`.** python-docx returns None whenever a
  paragraph points at a style the document's styles part doesn't define —
  routine in files converted from `.doc` or carrying content pasted in from
  another document. 15% of the real base hit it, and reading `.name` off that
  None aborted the entire audit.
- **A crash exited `1`.** This tool's own contract reads `1` as "problems
  found", so any wrapper, CI job or batch-edit gate keying on the exit code
  took a traceback for a completed audit. Unexpected failures now exit `2`.
- **Blank rows at the bottom of a Version History** made the current version
  read as `""`. The comparison discards empty sources before comparing, so
  `{header, "", index}` collapsed to a single value and the version-divergence
  check quietly stopped running for that document — no output, no warning.

The shape these share is the one worth naming: each failed by going quiet
rather than loud. That is the failure mode an audit tool can least afford,
because the report still comes back and you still believe it.

### Two places the standard described a practice nobody followed

- **Prerequisites and Verification may be numbered, not only bulleted.** The
  real base contained bullets in 6 documents and none in 33. A rule that most
  of a mature base works around is not being enforced. What the rule actually
  cares about is that the items form a list rather than prose; the section
  heading remains a heading either way.
- **Prerequisites is required even when there are none** — write a single
  `N/A` item. An explicit `N/A` records that the author considered it; an
  absent section records nothing, and the two are indistinguishable to whoever
  reads the document a year later.

### New checks

| Check | Catches |
|---|---|
| Invalid Status | A Status cell empty or outside `Active` / `In review` / `Legacy`, which stops the Overview's COUNTIF formulas counting the row and can disable the legacy exemption |
| Version History sequence | A history that doesn't begin at the creation (1.0), or whose versions move backwards |
| Shared numbering sequence | Two sections on one numbering sequence, so the second continues the first's count instead of restarting at 1 |
| Unreadable documents | A file the index points at that exists but cannot be parsed as Word |
| Header missing its version | A header carrying no version at all — exactly where "the header must match the Version History" stops being enforceable |

The numbering check deserves a note: nothing about such a document looks wrong
on screen, because Word renders numbers either way. They are simply the wrong
ones, and it surfaces only when someone is told to follow step 9 of a
four-step procedure. The real base has 24 such pairs across 13 documents.

### Any language, for real this time

Section 10 has always said a base may be authored in any language, while eight
tables in the code hardcoded English and Portuguese — so "any language" meant
"either of two, and only by editing Python". Every localized term now lives in
`scripts/languages.json`. **Adding a language is an edit to that file, not to
either script.** The checker recognizes the terms of every language listed
there at once, so it reads a base in any of them without being told which.

`--lang` is now optional: the language is detected by counting how many
section headings match each language's vocabulary, and the scores are printed.
A tie or an empty base is reported as undecided rather than guessed. On the
real base: 40 documents read, Portuguese 122 to English 0.

### Continuous integration

`.github/workflows/tests.yml` runs the suite on Ubuntu and Windows, Python 3.9
and 3.13. Every defect fixed in 0.8.4 was Windows-only and none reproduced on
Linux, so a single-platform check would have missed them exactly as a human
did. The workflow also audits the fixture twice — to a console and redirected
to a file — and fails if anything reaches stderr or the report has no
`SUMMARY` line.

### Also

- Interactivity is detected with `GetConsoleMode` rather than `isatty()` alone.
  On Windows a `NUL` stdin reports as a terminal, so a scheduled run printed a
  menu to nobody and then contradicted itself with "running non-interactively".
- Each document is opened once per row instead of being parsed four times.
- The Overview's count formulas cover rows 3–10000 instead of 3–1000.

## Verification

16 test groups, 15 detection sections, 38 fixture problems, plus the previous
release's 10 scenarios re-run against this one. Captured output is in
[`test-runs/0.9.0/`](../../test-runs/0.9.0/).

```
pip install -r skills/kb-compiler/scripts/requirements.txt
python3 skills/kb-compiler/tests/run_tests.py
```

**Not verified:** CI has never actually run — it executes only after a push —
so **Python 3.9 is untested**; only 3.13 was exercised. If the first CI run
fails on 3.9, that is the workflow doing its job. The positive case of the
interactivity check ("a real console returns True") is likewise untested,
because every path available during development ran with a redirected stdin;
it rests on documented Windows behaviour, with the fallback defaulting to
interactive so that a prompt is never wrongly suppressed.

## Upgrading

```
/plugin marketplace update kb-compiler
```

Or download the release zip and upload it in the desktop app under
Settings → Plugins → Add, which works regardless of managed marketplace policy.
