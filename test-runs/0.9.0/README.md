# Test run — 0.9.0

Produced by running the tooling as it ships in 0.9.0, on the machine described
in [`05-environment.txt`](05-environment.txt). This is the evidence trail for
[`CHANGELOG.md`](../../CHANGELOG.md), not documentation — regenerate it rather
than editing the text.

What makes this release's run different from 0.8.4's: it includes an audit of a
**real 68-document knowledge base**, which is where every fix in this release
came from. The synthetic fixture had never produced any of them.

## What is here

| File | What it shows |
|---|---|
| [`01-test-suite.txt`](01-test-suite.txt) | 16 groups, 15 detection sections, 38 fixture problems, 3 unverifiable. Exit 0. |
| [`02-verify-fixes-0.8.4.txt`](02-verify-fixes-0.8.4.txt) | The previous release's 10 scenarios, re-run against this one — the older fixes still hold. Exit 0. |
| [`03-fixture-audit-report.txt`](03-fixture-audit-report.txt) | A complete audit report, redirected to a file, with an empty stderr. |
| [`04-real-base-summary.txt`](04-real-base-summary.txt) | Section counts from the real base. Findings only — no document content. |
| [`05-environment.txt`](05-environment.txt) | Platform, Python, pinned package versions, locale and encoding variables. |
| [`RELEASE-NOTES.md`](RELEASE-NOTES.md) | Ready to paste into the GitHub release. |
| [`probes/verify_fixes_0.8.4.py`](probes/verify_fixes_0.8.4.py) | The 0.8.4 scenario script, kept alongside its output. |

`04-real-base-summary.txt` deliberately holds counts and section headings only.
The base is a company's internal documentation; nothing from inside the
documents is reproduced here.

## Reproducing it

```
pip install -r skills/kb-compiler/scripts/requirements.txt
python3 skills/kb-compiler/tests/run_tests.py
python3 test-runs/0.8.4/probes/verify_fixes.py
```

Both must exit `0`.

## What is NOT covered

- **CI has never run.** It executes only after a push, so **Python 3.9 is
  unverified** — only 3.13 was exercised here. The workflow declares 3.9 and
  3.13 on Ubuntu and Windows; the first real run is what will confirm it.
- **The interactivity check's positive case is untested.** "A real console
  returns True" rests on documented Windows behaviour: every path available
  during development ran with a redirected stdin. The fallback defaults to
  interactive, so the failure mode is an extra prompt rather than a suppressed
  one.
- **One platform, one Python** for everything else: Windows 11, CPython 3.13.
- **The real base is one base, in one language.** Language detection scored
  Portuguese 122 to English 0 there; a base mixing languages, or written in a
  third one added to `languages.json`, has not been exercised.
- **`os.path.getctime`** still backs the "Creation Date" column, and means
  creation time on Windows but inode change time on Unix. Known, not fixed.
