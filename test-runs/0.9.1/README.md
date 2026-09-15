# Test run — 0.9.1

Produced by running the tooling as it ships in 0.9.1, on the machine described
in [`05-environment.txt`](05-environment.txt). This is the evidence trail for
[`CHANGELOG.md`](../../CHANGELOG.md), not documentation — regenerate it rather
than editing the text.

What makes this release's run different from 0.9.0's: 0.9.0 **audited** a real
base, 0.9.1 **remediated** one. Working through 252 findings on a 60-document
base surfaced a class of defect that reading a report does not — over half of
this release is the checker being wrong rather than the base.

The comparison in [`04-real-base-summary.txt`](04-real-base-summary.txt) is the
part worth reading: a copy of the base was taken **before any document was
edited**, so both checker versions could be run against byte-identical input.
That makes "14 of the 252 findings were false positives" a measurement rather
than a claim.

## What is here

| File | What it shows |
|---|---|
| [`01-test-suite.txt`](01-test-suite.txt) | 28 groups, 17 detection sections, 40 fixture problems, 3 unverifiable. Exit 0. |
| [`02-verify-fixes-0.8.4.txt`](02-verify-fixes-0.8.4.txt) | 0.8.4's 10 scenarios, re-run against this release — the older fixes still hold. Exit 0. |
| [`03-fixture-audit-report.txt`](03-fixture-audit-report.txt) | A complete audit report, redirected to a file, with an empty stderr. |
| [`04-real-base-summary.txt`](04-real-base-summary.txt) | Three runs on the real base: 0.9.0's checker, 0.9.1's checker on the same untouched copy, and 0.9.1's after remediation. Counts only. |
| [`05-environment.txt`](05-environment.txt) | Platform, Python, pinned package versions, locale and encoding variables. |
| [`RELEASE-NOTES.md`](RELEASE-NOTES.md) | Ready to paste into the GitHub release. |
| [`probes/verify_fixes_0.8.4.py`](probes/verify_fixes_0.8.4.py) | The 0.8.4 scenario script, kept alongside its output. |

`04-real-base-summary.txt` deliberately holds counts and section headings only.
The base is a company's internal documentation; nothing from inside the
documents is reproduced here.

## The suite grew for a specific reason

Every group added in this release was run against **0.9.0's** checker and
confirmed to fail there before being kept. A regression test that passes before
the fix is testing nothing.

Two of them assert on a named document rather than a section count, and they
have to: against 0.9.0's checker the bullet false positive (+1) and the missed
typed page number (−1) cancel out exactly, so `"Formatting violations": 17` held
for the wrong reasons and a count assertion alone could not have caught either
defect.

One case in the existing fixture turned out to be unfaithful — it claimed to
model two sections sharing a numbering sequence while pinning both paragraphs to
a *bullet* definition, so it modelled nothing and passed only because the
checker ignored `numFmt`. Fixing the checker broke the test, which is how it was
found.

## Reproducing it

```bash
pip install -r skills/kb-compiler/scripts/requirements.txt

# 01
python skills/kb-compiler/tests/run_tests.py

# 02
python test-runs/0.9.1/probes/verify_fixes_0.8.4.py

# 03  — exit 1 is correct here: the fixture has deliberate problems
python skills/kb-compiler/tests/build_fixture.py /tmp/fixture
python skills/kb-compiler/scripts/check_master_list.py /tmp/fixture > report.txt 2> report.err
# report.err must be empty, and report.txt must end with a SUMMARY line

# 04 needs a real base, and is not reproducible from this repository
```
