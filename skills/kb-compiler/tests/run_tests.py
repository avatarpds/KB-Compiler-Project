#!/usr/bin/env python3
"""
run_tests.py — Test suite for the kb-compiler tooling.

Covers three areas:

  1. Detections — builds the fixture (one deliberate instance of every problem
     check_master_list.py should catch) and compares every report section's
     count with the expected value.
  2. Bootstrap — verifies bootstrap_master_list.py produces a spreadsheet the
     checker reads cleanly, never invents a KB code, marks every row as
     unreviewed, and refuses to overwrite an existing file.
  3. Non-interactive safety — verifies that a base with no spreadsheet does not
     block waiting for input when stdin isn't a terminal, and exits 2.

Run this after ANY change to the scripts. A detection that silently stops
working is the worst failure mode for an audit tool: the report comes back
clean and you believe it.

Usage:
    python3 tests/run_tests.py

Exit code: 0 if every expectation holds, 1 otherwise.
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, "..", "scripts")
CHECKER = os.path.join(SCRIPTS, "check_master_list.py")
BOOTSTRAP = os.path.join(SCRIPTS, "bootstrap_master_list.py")
BUILDER = os.path.join(HERE, "build_fixture.py")

TIMEOUT = 120  # seconds; also catches a script that blocks on input

# --- Expected counts for the detection fixture -----------------------------
EXPECTED = {
    "Broken references": 1,          # KB-050 points to a nonexistent file
    # KB-099 (no row at all) + the unindexed twin of KB-045, which only becomes
    # visible once duplicate basenames resolve to the right file.
    "Orphaned files": 2,
    "Duplicate/shared codes": 1,     # KB-009 reused across two sheets
    "Name divergence": 3,            # KB-010, KB-034 (blank first paragraph), KB-009 (duplicate row)
    "Version divergence": 1,         # KB-017 header/history mismatch
    "Version History table in the legacy format": 1,   # KB-020
    "Header missing the tab separator": 1,             # KB-011
    # KB-030 (4 margins + size + color), KB-031 (shading + color), KB-040
    # (landscape + no footer), KB-041 (en dash), KB-042 (2 non-list sections),
    # KB-043 (all-caps author), plus one base-level footer-label mismatch.
    "Formatting violations": 15,
    "Structure violations": 2,       # KB-036: missing Prerequisites + history not last
    "Folder/tab mismatch": 2,        # KB-035 (planted) + KB-009's duplicate row
    "Ambiguous file reference": 0,   # the tab resolves KB-045; nothing stays ambiguous
}
EXPECTED_UNVERIFIABLE = 3   # KB-032's second title run: bold, size, color
EXPECTED_TOTAL = sum(EXPECTED.values())


def run(args):
    return subprocess.run(
        [sys.executable] + args,
        capture_output=True, text=True, timeout=TIMEOUT,
        stdin=subprocess.DEVNULL,
    )


def parse_sections(output):
    found = {}
    for line in output.splitlines():
        m = re.match(r"^=== (.+?) \((\d+)\) ===$", line.strip())
        if m:
            found[m.group(1)] = int(m.group(2))
    return found


def match_section(found, prefix):
    for title, count in found.items():
        if title.startswith(prefix):
            return count
    return None


# ---------------------------------------------------------------------------
# 1. Detections
# ---------------------------------------------------------------------------
def test_detections(tmp):
    failures = []
    base = os.path.join(tmp, "detections")
    built = run([BUILDER, base])
    if built.returncode != 0:
        # Without this, a fixture that fails to build shows up disguised as
        # "detection missing" and sends you looking in the wrong file.
        return ["fixture builder failed (rc=%d): %s" % (built.returncode, built.stderr.strip()[:300])]

    proc = run([CHECKER, base])
    output = proc.stdout
    found = parse_sections(output)

    for prefix, expected_count in EXPECTED.items():
        actual = match_section(found, prefix)
        if actual is None:
            failures.append("section '%s' missing from the report" % prefix)
        elif actual != expected_count:
            failures.append("'%s': expected %d, got %d" % (prefix, expected_count, actual))

    m = re.search(r"could not be verified \((\d+)\)", output)
    unverifiable = int(m.group(1)) if m else None
    if unverifiable != EXPECTED_UNVERIFIABLE:
        failures.append("'could not be verified': expected %d, got %s"
                        % (EXPECTED_UNVERIFIABLE, unverifiable))

    m = re.search(r"SUMMARY: (\d+) problem", output)
    total = int(m.group(1)) if m else None
    if total != EXPECTED_TOTAL:
        failures.append("SUMMARY total: expected %d, got %s" % (EXPECTED_TOTAL, total))
    # The "could not verify" bucket must never inflate the problem count —
    # that separation is the whole point of conservative mode.
    if total is not None and total != sum(EXPECTED.values()):
        failures.append("unverifiable items leaked into the problem count")

    if proc.returncode != 1:
        failures.append("exit code with problems: expected 1, got %d" % proc.returncode)

    return failures


# ---------------------------------------------------------------------------
# 2. Bootstrap
# ---------------------------------------------------------------------------
def make_base_without_spreadsheet(base):
    """Reuses the fixture builder, then deletes the spreadsheet it wrote."""
    run([BUILDER, base])
    for name in os.listdir(base):
        if name.lower().endswith(".xlsx"):
            os.remove(os.path.join(base, name))


def test_bootstrap(tmp):
    failures = []
    base = os.path.join(tmp, "bootstrap")
    make_base_without_spreadsheet(base)

    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        return ["bootstrap exit code: expected 0, got %d (%s)"
                % (proc.returncode, proc.stdout.strip()[:200])]

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    if len(sheets) != 1:
        return ["expected exactly 1 spreadsheet created, found %d" % len(sheets)]

    wb = openpyxl.load_workbook(os.path.join(base, sheets[0]))
    rows = []
    for name in wb.sheetnames:
        if name in ("Visão Geral", "Overview"):
            continue
        ws = wb[name]
        for row in list(ws.iter_rows(values_only=True))[2:]:
            if row and row[0]:
                rows.append(row)

    if not rows:
        return ["bootstrap produced a spreadsheet with no rows"]

    # Every generated row is unreviewed: never "Ativo".
    bad = [r[0] for r in rows if r[3] not in ("Em revisão", "In review", "Legado", "Legacy")]
    if bad:
        failures.append("rows written with a reviewed/unknown status: %s" % bad)

    # No code may be invented: a document with no KB-XXX in its file name must
    # come out as the "—" marker, never as a freshly assigned code.
    for row in rows:
        code, filename = row[0], row[2]
        stem = os.path.splitext(str(filename))[0]
        has_code = bool(re.match(r"^(?:.*/)?KB-\d{3}\s*-\s*", stem))
        if not has_code and str(code).strip() != "—":
            failures.append("bootstrap invented a code for '%s': '%s'" % (filename, code))

    # The spreadsheet it wrote must be one the checker reads without
    # structural complaints (broken references / orphans).
    proc = run([CHECKER, base])
    found = parse_sections(proc.stdout)
    for prefix in ("Broken references", "Orphaned files"):
        count = match_section(found, prefix)
        if count != 0:
            failures.append(
                "after bootstrap, '%s' should be 0, got %s "
                "(the generated spreadsheet doesn't match the real files)" % (prefix, count)
            )

    # Running it again must refuse rather than overwrite.
    proc = run([BOOTSTRAP, base])
    if proc.returncode != 2:
        failures.append("bootstrap over an existing spreadsheet: expected exit 2, got %d"
                        % proc.returncode)
    if "already has a master-index spreadsheet" not in proc.stdout:
        failures.append("bootstrap didn't explain why it refused to overwrite")

    # And it must refuse even when the existing spreadsheet is in a DIFFERENT
    # language, i.e. when the new target file name wouldn't collide. Checking
    # only the target name used to create a second spreadsheet beside the first.
    proc = run([BOOTSTRAP, base, "--lang", "pt"])
    if proc.returncode != 2:
        failures.append("bootstrap with a different --lang on an indexed base: "
                        "expected exit 2, got %d" % proc.returncode)
    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    if len(sheets) != 1:
        failures.append("a second spreadsheet was created in another language: %s" % sheets)

    return failures


# ---------------------------------------------------------------------------
# 3. Non-interactive safety
# ---------------------------------------------------------------------------
def test_non_interactive(tmp):
    failures = []
    base = os.path.join(tmp, "nospreadsheet")
    make_base_without_spreadsheet(base)

    # stdin is DEVNULL in run(), so isatty() is False: the checker must print
    # guidance and exit instead of blocking on input that will never come.
    proc = run([CHECKER, base])

    if proc.returncode != 2:
        failures.append("missing spreadsheet: expected exit 2, got %d" % proc.returncode)
    if "bootstrap_master_list.py" not in proc.stdout:
        failures.append("the guidance doesn't point at bootstrap_master_list.py")
    if "Traceback" in proc.stderr:
        failures.append("a missing spreadsheet still produces a raw traceback")

    return failures


# ---------------------------------------------------------------------------
# 4. Regressions found by external review
# ---------------------------------------------------------------------------
def test_invalid_sheet_characters(tmp):
    """A folder name Excel forbids in a sheet title used to crash the bootstrap."""
    base = os.path.join(tmp, "sheetchars")
    os.makedirs(os.path.join(base, "01 - Rede [WAN]"))
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, "01 - Rede [WAN]"))

    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        return ["bootstrap failed on a folder with '[' in the name (rc=%d)" % proc.returncode]
    if "Traceback" in proc.stderr:
        return ["bootstrap still raises on an invalid sheet character"]

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    wb = openpyxl.load_workbook(os.path.join(base, sheets[0]))
    overview = wb[wb.sheetnames[0]]
    failures = []
    for row in list(overview.iter_rows(values_only=True))[2:]:
        if not row or not row[0]:
            continue
        referenced = str(row[4] or "")
        if not any("'%s'" % name in referenced for name in wb.sheetnames):
            failures.append("Overview tab reference %r points at no existing sheet" % referenced)
    return failures


def test_numbered_legacy_folder(tmp):
    """'09 - Legado' must be recognized as legacy, same as a bare 'Legado'."""
    base = os.path.join(tmp, "numberedlegacy")
    os.makedirs(os.path.join(base, "09 - Legado"))
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, "09 - Legado"))

    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        return ["bootstrap failed (rc=%d)" % proc.returncode]

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    wb = openpyxl.load_workbook(os.path.join(base, sheets[0]))
    failures = []
    for name in wb.sheetnames[1:]:
        for row in list(wb[name].iter_rows(values_only=True))[2:]:
            if not row or not row[0]:
                continue
            if row[3] not in ("Legado", "Legacy"):
                failures.append("numbered legacy folder produced status %r" % row[3])
            if "/" not in str(row[2]):
                failures.append("legacy row is missing its relative path: %r" % row[2])
    return failures


def test_invalid_lang_rejected(tmp):
    """--lang with an unknown value must fail, not fall back and write a second file."""
    base = os.path.join(tmp, "badlang")
    os.makedirs(os.path.join(base, "01 - Cat"))
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, "01 - Cat"))

    failures = []
    proc = run([BOOTSTRAP, base, "--lang", "xx"])
    if proc.returncode != 2:
        failures.append("--lang xx: expected exit 2, got %d" % proc.returncode)
    if [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]:
        failures.append("--lang xx wrote a spreadsheet instead of refusing")
    return failures


def test_ambiguous_spreadsheet(tmp):
    """
    Two spreadsheets matching different naming patterns must be an explicit
    error, never a silent pick. This was the worst bug found: a clean report
    on a base that had only been half audited.
    """
    base = os.path.join(tmp, "ambiguous")
    os.makedirs(os.path.join(base, "01 - Cat"))
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, "01 - Cat"))

    run([BOOTSTRAP, base, "--lang", "pt"])
    # The bootstrap now refuses to add a second spreadsheet, so put the base in
    # that state by hand — it can still happen (a copied file, two people, a
    # sync conflict), and the checker must not quietly pick one of them.
    existing = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")][0]
    shutil.copy(os.path.join(base, existing), os.path.join(base, "Master List - copy.xlsx"))
    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    if len(sheets) != 2:
        return ["setup: expected 2 spreadsheets, got %d" % len(sheets)]

    failures = []

    # With the base's config present, there is no ambiguity: the pointer says
    # which file is the index, and the checker must use it.
    proc = run([CHECKER, base])
    if proc.returncode == 2:
        failures.append("a configured base was treated as ambiguous")
    if "SUMMARY" not in proc.stdout:
        failures.append("a configured base produced no report")

    # Without the config (an older base, or one set up by hand), two candidate
    # spreadsheets must be an explicit error rather than a silent pick.
    os.remove(os.path.join(base, ".kb-compiler.json"))
    proc = run([CHECKER, base])
    if proc.returncode != 2:
        failures.append("ambiguous spreadsheets: expected exit 2, got %d" % proc.returncode)
    if "More than one" not in proc.stdout:
        failures.append("the checker didn't say the spreadsheet choice is ambiguous")
    if "SUMMARY" in proc.stdout:
        failures.append("the checker produced a report despite not knowing which spreadsheet to read")
    return failures


def test_sanitized_tab_no_false_mismatch(tmp):
    """
    A category whose name had to be altered to become a legal Excel sheet
    title (invalid character, 31-char cap) must NOT then be reported as living
    in the wrong folder. The checker reads the Overview's recorded mapping
    instead of recomputing the transformation.
    """
    base = os.path.join(tmp, "sanitizedtab")
    long_name = "01 - Identity Access and Corporate Governance"
    bracket_name = "02 - Network [WAN]"
    os.makedirs(os.path.join(base, long_name))
    os.makedirs(os.path.join(base, bracket_name))
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, long_name, "KB-001 - One.docx"))
    shutil.copy(fixture, os.path.join(base, bracket_name, "KB-002 - Two.docx"))

    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        return ["bootstrap failed (rc=%d)" % proc.returncode]

    proc = run([CHECKER, base])
    found = parse_sections(proc.stdout)
    count = match_section(found, "Folder/tab mismatch")
    if count:
        return ["sanitized/truncated tab names produced %d false mismatch(es)" % count]
    return []


def test_legacy_status_skips_structure(tmp):
    """
    A document marked Legacy predates the standard by definition; flagging it
    for missing sections would flood the report with findings nobody intends
    to fix.
    """
    base = os.path.join(tmp, "legacystatus")
    os.makedirs(os.path.join(base, "Legacy"))

    import docx
    d = docx.Document()
    d.sections[0].header.paragraphs[0].text = "Old Process\tv1.0"
    d.add_paragraph("Old Process")          # no sections at all
    d.save(os.path.join(base, "Legacy", "Old Process.docx"))

    wb = openpyxl.Workbook()
    ov = wb.active
    ov.title = "Overview"
    ov.cell(row=1, column=1, value="Knowledge Base — Fixture")
    ws = wb.create_sheet("Legacy")
    ws.cell(row=1, column=1, value="Knowledge Base — Legacy")
    for i, h in enumerate(["Code", "Document", "File", "Status", "Version",
                           "Creation Date", "Last Updated", "Owner"], 1):
        ws.cell(row=2, column=i, value=h)
    for i, v in enumerate(("—", "Old Process", "Legacy/Old Process.docx", "Legacy",
                           "1.0", "", "", "Lucas Souza"), 1):
        ws.cell(row=3, column=i, value=v)
    wb.save(os.path.join(base, "Master List - legacystatus.xlsx"))

    proc = run([CHECKER, base])
    found = parse_sections(proc.stdout)
    count = match_section(found, "Structure violations")
    if count:
        return ["a Legacy-status document produced %d structure violation(s)" % count]
    return []


def test_custom_destination_is_remembered(tmp):
    """
    A master index created under a custom name (or outside the base) must be
    found automatically on every later run — that's what makes it a continuous
    insertion point rather than a one-off file.
    """
    failures = []
    base = os.path.join(tmp, "customdest")
    os.makedirs(os.path.join(base, "01 - Cat"))
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, "01 - Cat"))

    proc = run([BOOTSTRAP, base, "--output", "Internal Index.xlsx"])
    if proc.returncode != 0:
        return ["bootstrap with --output failed (rc=%d)" % proc.returncode]
    if not os.path.isfile(os.path.join(base, "Internal Index.xlsx")):
        failures.append("--output didn't create the file at the requested name")
    if not os.path.isfile(os.path.join(base, ".kb-compiler.json")):
        failures.append("the destination wasn't recorded in the base's config")

    # No file name passed: the checker must resolve it from the config, even
    # though the name matches none of the built-in patterns.
    proc = run([CHECKER, base])
    if "Internal Index.xlsx" not in proc.stdout:
        failures.append("the checker didn't pick up the configured master index")
    if proc.returncode != 0:
        failures.append("checker on a clean base with a custom index: expected 0, got %d"
                        % proc.returncode)

    # A configured path that no longer exists must be an explicit error, never
    # a silent fallback to auto-detection on a different file.
    os.remove(os.path.join(base, "Internal Index.xlsx"))
    proc = run([CHECKER, base])
    if proc.returncode != 2:
        failures.append("missing configured index: expected exit 2, got %d" % proc.returncode)
    return failures


def test_out_of_base_destination(tmp):
    """The index may live outside the base folder; the pointer still resolves."""
    base = os.path.join(tmp, "outofbase")
    elsewhere = os.path.join(tmp, "elsewhere")
    os.makedirs(os.path.join(base, "01 - Cat"))
    os.makedirs(elsewhere)
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, "01 - Cat"))

    target = os.path.join(elsewhere, "index.xlsx")
    proc = run([BOOTSTRAP, base, "--output", target])
    if proc.returncode != 0:
        return ["bootstrap to an out-of-base path failed (rc=%d)" % proc.returncode]
    if not os.path.isfile(target):
        return ["the index wasn't created at the out-of-base path"]

    proc = run([CHECKER, base])
    failures = []
    if proc.returncode != 0:
        failures.append("checker with an out-of-base index: expected 0, got %d" % proc.returncode)
    if "index.xlsx" not in proc.stdout:
        failures.append("the checker didn't resolve the out-of-base index")
    return failures


def test_root_level_documents_are_indexed(tmp):
    """
    A document loose in the base root has no folder to infer a category from.
    It must still be indexed — otherwise the bootstrap skips it and the checker
    then reports it as an orphan, with the two tools contradicting each other
    about the same file.
    """
    base = os.path.join(tmp, "rootdocs")
    os.makedirs(os.path.join(base, "01 - Cat"))
    fixture = os.path.join(tmp, "detections", "01 - Identity and Access",
                           "KB-009 - Configurar MFA.docx")
    shutil.copy(fixture, os.path.join(base, "01 - Cat"))
    shutil.copy(fixture, os.path.join(base, "KB-050 - Loose Document.docx"))

    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        return ["bootstrap failed (rc=%d)" % proc.returncode]

    failures = []
    if "base root" not in proc.stdout:
        failures.append("the root-level document wasn't reported as needing a category")

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    wb = openpyxl.load_workbook(os.path.join(base, sheets[0]))
    indexed = []
    for name in wb.sheetnames[1:]:
        for row in list(wb[name].iter_rows(values_only=True))[2:]:
            if row and row[0]:
                indexed.append(str(row[2]))
    if "KB-050 - Loose Document.docx" not in indexed:
        failures.append("the root-level document was not indexed at all")

    # The index itself and the config file must never become rows.
    for unwanted in (sheets[0], ".kb-compiler.json"):
        if unwanted in indexed:
            failures.append("'%s' was indexed as if it were a document" % unwanted)

    proc = run([CHECKER, base])
    found = parse_sections(proc.stdout)
    for prefix in ("Orphaned files", "Folder/tab mismatch"):
        count = match_section(found, prefix)
        if count:
            failures.append("root-level document produced %d '%s'" % (count, prefix))
    return failures


TESTS = [
    ("detections", test_detections),
    ("bootstrap", test_bootstrap),
    ("non-interactive safety", test_non_interactive),
    ("invalid sheet characters", test_invalid_sheet_characters),
    ("numbered legacy folder", test_numbered_legacy_folder),
    ("invalid --lang rejected", test_invalid_lang_rejected),
    ("ambiguous spreadsheet", test_ambiguous_spreadsheet),
    ("sanitized tab / no false mismatch", test_sanitized_tab_no_false_mismatch),
    ("legacy status skips structure", test_legacy_status_skips_structure),
    ("custom destination remembered", test_custom_destination_is_remembered),
    ("out-of-base destination", test_out_of_base_destination),
    ("root-level documents indexed", test_root_level_documents_are_indexed),
]


def main():
    tmp = tempfile.mkdtemp(prefix="kb-tests-")
    all_failures = {}
    try:
        for name, fn in TESTS:
            try:
                failures = fn(tmp)
            except subprocess.TimeoutExpired:
                failures = ["timed out after %ds (a script may be blocking on input)" % TIMEOUT]
            except Exception as e:  # a crashing test is a failing test
                failures = ["test raised %s: %s" % (type(e).__name__, e)]
            if failures:
                all_failures[name] = failures
            print("[%s] %s" % ("FAIL" if failures else " OK ", name))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if all_failures:
        print("\nFAILED:")
        for name, failures in all_failures.items():
            for f in failures:
                print("  - [%s] %s" % (name, f))
        return 1

    print("\nPASSED: %d test groups, %d detection sections, %d problems, %d unverifiable."
          % (len(TESTS), len(EXPECTED), EXPECTED_TOTAL, EXPECTED_UNVERIFIABLE))
    return 0


if __name__ == "__main__":
    sys.exit(main())
