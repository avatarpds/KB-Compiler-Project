# Test run — 0.9.2

0.9.2 changes documentation only: two paragraphs in `SKILL.md`, the version in
both manifests, and a changelog entry. No Python was touched.

So this folder is deliberately thinner than
[`0.9.1/`](../0.9.1/). There is no fixture audit, no real-base summary and no
re-run of older scenario probes, because nothing that produces them changed.
What is here proves the suite is still green on the released tree:

| File | What it shows |
|---|---|
| [`01-test-suite.txt`](01-test-suite.txt) | 28 groups, 17 detection sections, 40 fixture problems, 3 unverifiable. Exit 0. |
| [`05-environment.txt`](05-environment.txt) | Platform, Python, pinned package versions, locale and encoding. |
| [`RELEASE-NOTES.md`](RELEASE-NOTES.md) | Ready to paste into the GitHub release. |

For the full evidence trail behind the code as it stands, see
[`test-runs/0.9.1/`](../0.9.1/) — it still describes this release's behaviour,
because the behaviour is unchanged.

## Reproducing it

```bash
pip install -r skills/kb-compiler/scripts/requirements.txt
python skills/kb-compiler/tests/run_tests.py
```
