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

import json
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
    # KB-017 (header/history mismatch) + KB-058, whose divergence was hidden by
    # blank rows at the bottom of its Version History table.
    "Version divergence": 2,
    "Version History table in the legacy format": 1,   # KB-020
    # KB-011 (glued: "YubiKeyv1.0") + KB-046 (spaced: "Nome Colado v1.0")
    "Header missing the tab separator": 2,
    "Header missing its version": 1,   # KB-037: "... Office v2", no version at all
    "Unreadable documents": 1,         # KB-048: .docx extension, not a Word file
    "Invalid Status value": 1,         # KB-049: "Em Revisao", missing its cedilla
    # KB-051 (history starts at 2.0) + KB-052 (1.2 -> 1.1). KB-053 (1.9 -> 1.10)
    # must NOT appear: it moves forward, and only trips a text comparison.
    "Version History sequence": 2,
    # KB-030 (4 margins + size + color), KB-031 (shading + color), KB-040
    # (landscape + no footer), KB-041 (en dash), KB-042 (2 non-list sections),
    # KB-043 (all-caps author), KB-060 (two sections on one DECIMAL numbering
    # sequence), KB-063 (typed page number in a footer containing the word
    # PAGE), KB-064 (a "List Number" step whose direct numbering is a bullet
    # definition, so it renders as a bullet), plus one base-level footer-label
    # mismatch.
    # KB-061 must NOT appear here: it shares a BULLET definition, which has no
    # count for a second section to continue. KB-062 must not either: its NFD
    # title is the same text as the index's NFC name.
    "Formatting violations": 18,
    # Both of these are clean in the fixture; they have dedicated groups below.
    # Declared here so the report is asserted to still CARRY the sections — a
    # detection that silently stops being printed is the failure mode this
    # whole suite exists to catch.
    "Sheets skipped entirely": 0,
    "Row with an empty Code cell": 0,
    "Structure violations": 2,       # KB-036: missing Prerequisites + history not last
    "Folder/tab mismatch": 2,        # KB-035 (planted) + KB-009's duplicate row
    # KB-045 is resolved by its tab; KB-047's two copies are not.
    "Ambiguous file reference": 1,
}
EXPECTED_UNVERIFIABLE = 3   # KB-032's second title run: bold, size, color
EXPECTED_TOTAL = sum(EXPECTED.values())


def run(args):
    return subprocess.run(
        [sys.executable] + args,
        capture_output=True, text=True, timeout=TIMEOUT,
        # A real empty pipe, NOT subprocess.DEVNULL: on Windows stdin redirected
        # from NUL is reported as a terminal by isatty(), so DEVNULL made every
        # script under test take the INTERACTIVE path — 7 of these 12 groups
        # failed for that reason alone, on the only OS the author runs.
        input="",
        encoding="utf-8",
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
    # The per-section counts above already sum to EXPECTED_TOTAL, so comparing
    # the summary against sum(EXPECTED.values()) a second time could never fail
    # independently — it only looked like an extra assertion.

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
    # The interactive menu must not be printed at all. It used to appear first,
    # with "running non-interactively" immediately after it, because Windows
    # reports a NUL stdin as a terminal and the truth only surfaced when input()
    # hit EOF. The output contradicted itself.
    for line in ("[1] Create one now", "What would you like to do?"):
        if line in proc.stdout:
            failures.append("the interactive menu was printed to a non-interactive "
                            "run: %r" % line)

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


def _fixture_doc(tmp):
    """A known-good document from the detections fixture, to copy around."""
    return os.path.join(tmp, "detections", "01 - Identity and Access",
                        "KB-009 - Configurar MFA.docx")


def test_colliding_category_names(tmp):
    """
    Two folders that strip to the same category name must not lose one of them.
    Assigning instead of merging silently dropped every document of whichever
    folder was scanned first; the checker then reported them as orphans, with
    nothing in the bootstrap's output saying they had been lost.
    """
    base = os.path.join(tmp, "collide")
    for folder, name in (("01 - Rede", "KB-001 - Um.docx"),
                         ("02 - Rede", "KB-002 - Dois.docx")):
        os.makedirs(os.path.join(base, folder), exist_ok=True)
        shutil.copy(_fixture_doc(tmp), os.path.join(base, folder, name))

    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        return ["bootstrap failed (rc=%d)" % proc.returncode]

    failures = []
    if "strips to the category name" not in proc.stdout:
        failures.append("merging two folders into one tab wasn't reported to the user")

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    wb = openpyxl.load_workbook(os.path.join(base, sheets[0]))
    indexed = [str(row[2]) for name in wb.sheetnames[1:]
               for row in list(wb[name].iter_rows(values_only=True))[2:]
               if row and row[0]]
    for expected in ("KB-001 - Um.docx", "KB-002 - Dois.docx"):
        if expected not in indexed:
            failures.append("'%s' was dropped from the index entirely" % expected)

    count = match_section(parse_sections(run([CHECKER, base]).stdout), "Orphaned files")
    if count:
        failures.append("colliding category names left %s orphan(s)" % count)
    return failures


def test_overview_extra_column(tmp):
    """
    A note typed into a column past "Tab" must not hide the Overview's
    category -> tab mapping. Reading only the row's LAST cell lost it, which
    brought back the false folder/tab mismatches on truncated tab names that
    the mapping exists to prevent.
    """
    base = os.path.join(tmp, "overviewextra")
    folder = "01 - Identity Access and Corporate Governance Board"
    os.makedirs(os.path.join(base, folder))
    shutil.copy(_fixture_doc(tmp), os.path.join(base, folder, "KB-001 - One.docx"))

    if run([BOOTSTRAP, base]).returncode != 0:
        return ["bootstrap failed"]

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    path = os.path.join(base, sheets[0])
    wb = openpyxl.load_workbook(path)
    ov = wb[wb.sheetnames[0]]
    ov.cell(row=3, column=7, value="note typed by a human")
    wb.save(path)

    count = match_section(parse_sections(run([CHECKER, base]).stdout), "Folder/tab mismatch")
    if count:
        return ["an extra Overview column produced %s false mismatch(es)" % count]
    return []


def test_force_does_not_index_the_index(tmp):
    """
    A custom-named index is still an .xlsx sitting in the base, and .xlsx is
    indexable — so a --force re-run used to add the index to itself as if it
    were one of the base's documents.
    """
    base = os.path.join(tmp, "forceindex")
    os.makedirs(os.path.join(base, "01 - Cat"))
    shutil.copy(_fixture_doc(tmp), os.path.join(base, "01 - Cat", "KB-001 - One.docx"))

    if run([BOOTSTRAP, base, "--output", "Internal Index.xlsx"]).returncode != 0:
        return ["bootstrap with --output failed"]
    if run([BOOTSTRAP, base, "--output", "Internal Index.xlsx", "--force"]).returncode != 0:
        return ["bootstrap --force failed"]

    wb = openpyxl.load_workbook(os.path.join(base, "Internal Index.xlsx"))
    indexed = [str(row[2]) for name in wb.sheetnames[1:]
               for row in list(wb[name].iter_rows(values_only=True))[2:]
               if row and row[0]]
    if any("Internal Index" in f for f in indexed):
        return ["the master index indexed itself: %s" % indexed]
    return []


def test_language_is_detected(tmp):
    """
    A base whose documents are written in Portuguese must produce a Portuguese
    index without anyone passing --lang. Section 0 says to follow whichever
    language the base already uses; requiring the flag was the tooling
    contradicting its own standard, and getting it wrong writes an index whose
    column headers don't match the documents it indexes.
    """
    base = os.path.join(tmp, "langdetect")
    os.makedirs(os.path.join(base, "01 - Categoria"))
    # KB-039 in the fixture is the fully conforming Portuguese document.
    src = os.path.join(tmp, "detections", "03 - Infrastructure",
                       "KB-039 - Documento em Portugues.docx")
    for name in ("KB-001 - Um.docx", "KB-002 - Dois.docx"):
        shutil.copy(src, os.path.join(base, "01 - Categoria", name))

    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        return ["bootstrap failed (rc=%d): %s" % (proc.returncode, proc.stdout[-200:])]

    failures = []
    if "Language detected: pt" not in proc.stdout:
        failures.append("the detected language was not reported: %r" % proc.stdout[:220])

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    if not sheets:
        return ["no spreadsheet was created"]
    if not sheets[0].startswith("Lista Mestra"):
        failures.append("the index was not named in Portuguese: %r" % sheets[0])

    wb = openpyxl.load_workbook(os.path.join(base, sheets[0]))
    if wb.sheetnames[0] != "Visão Geral":
        failures.append("overview tab is %r, expected 'Visão Geral'" % wb.sheetnames[0])
    ws = wb[wb.sheetnames[1]]
    headers = [c for c in next(ws.iter_rows(min_row=2, max_row=2, values_only=True)) if c]
    if headers[:3] != ["Código", "Documento", "Arquivo"]:
        failures.append("index headers are not Portuguese: %r" % headers[:3])

    # An explicit --lang must still win over the detection.
    other = os.path.join(tmp, "langoverride")
    shutil.copytree(base, other)
    for f in os.listdir(other):
        if f.lower().endswith(".xlsx"):
            os.remove(os.path.join(other, f))
    os.remove(os.path.join(other, ".kb-compiler.json"))
    proc = run([BOOTSTRAP, other, "--lang", "en"])
    sheets = [f for f in os.listdir(other) if f.lower().endswith(".xlsx")]
    if not sheets or not sheets[0].startswith("Master List"):
        failures.append("--lang en did not override the detection: %r" % sheets)
    return failures


# ---------------------------------------------------------------------------
# Regressions fixed in 0.9.1
# ---------------------------------------------------------------------------
def _sole_index(base):
    """The one .xlsx in a base — whatever the bootstrap decided to call it."""
    names = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    if len(names) != 1:
        raise AssertionError("expected exactly one index in %s, found %s"
                             % (base, names))
    return os.path.join(base, names[0])


def _mini_base(tmp, name, doc_name="KB-001 - One.docx"):
    """A one-document base with a bootstrapped index, ready to be broken."""
    base = os.path.join(tmp, name)
    os.makedirs(os.path.join(base, "01 - Identity"))
    shutil.copy(_fixture_doc(tmp), os.path.join(base, "01 - Identity", doc_name))
    proc = run([BOOTSTRAP, base])
    if proc.returncode != 0:
        raise AssertionError("bootstrap failed (rc=%d)" % proc.returncode)
    return base


def test_skipped_sheet_is_counted(tmp):
    """
    A sheet the checker cannot read is a finding, not a footnote.

    It used to print only a WARNING and contribute nothing to the count, so a
    base whose sheet had lost its recognizable header row reported
    "SUMMARY: 0 problem(s) found" and exited 0 — the documented "clean" signal
    a CI gate keys on — while an entire sheet went unchecked. Whether the exit
    code happened to be non-zero depended on orphan detection incidentally
    firing for the same documents.
    """
    base = _mini_base(tmp, "skippedsheet")
    path = _sole_index(base)
    wb = openpyxl.load_workbook(path)
    ws = wb.create_sheet("SAP")
    ws.append(["Cod.", "Doc", "Arq."])       # none of these are column synonyms
    ws.append(["KB-900", "Planned", "KB-900 - Planned.docx"])
    wb.save(path)

    proc = run([CHECKER, base])
    failures = []
    count = match_section(parse_sections(proc.stdout), "Sheets skipped entirely")
    if count != 1:
        failures.append("expected 1 skipped-sheet finding, got %s" % count)
    if proc.returncode == 0:
        failures.append("a run that skipped a whole sheet must not exit 0")
    return failures


def test_empty_code_row_is_reported_once(tmp):
    """
    A row with real content but an empty Code used to be dropped in silence:
    nothing reported the row, and the document it indexes then resurfaced as an
    "orphan" — sending the reader after the file instead of the row that is
    actually wrong. Section 2's "—" marker exists precisely so a document
    without a code is still visible in the index.

    Asserts BOTH halves: the row is reported, and it is reported once.
    """
    base = _mini_base(tmp, "emptycode")
    path = _sole_index(base)
    wb = openpyxl.load_workbook(path)
    blanked = False
    for sheet in wb.sheetnames:
        for row in wb[sheet].iter_rows():
            for cell in row:
                if str(cell.value).strip() == "KB-001":
                    cell.value = None
                    blanked = True
    if not blanked:
        return ["could not find the KB-001 cell to blank"]
    wb.save(path)

    found = parse_sections(run([CHECKER, base]).stdout)
    failures = []
    if match_section(found, "Row with an empty Code cell") != 1:
        failures.append("expected 1 empty-Code finding, got %s"
                        % match_section(found, "Row with an empty Code cell"))
    if match_section(found, "Orphaned files") != 0:
        failures.append(
            "the row still references the file, so it must not ALSO be an "
            "orphan (got %s)" % match_section(found, "Orphaned files"))
    return failures


def test_office_lock_files_are_not_orphans(tmp):
    """
    Auditing a base while one document is open in Word reported that document's
    "~$" lock file as an orphan — a finding that disappears on its own, which
    is the fastest way to make a report look unreliable. Same for the junk
    Windows and macOS leave in a synced folder.
    """
    base = _mini_base(tmp, "lockfiles")
    folder = os.path.join(base, "01 - Identity")
    for name in ("~$B-001 - One.docx", "Thumbs.db", ".~lock.KB-001 - One.docx#"):
        with open(os.path.join(folder, name), "wb") as f:
            f.write(b"\x00")
    with open(os.path.join(base, ".DS_Store"), "wb") as f:
        f.write(b"\x00")

    count = match_section(parse_sections(run([CHECKER, base]).stdout),
                          "Orphaned files")
    if count:
        return ["lock files / OS junk produced %s false orphan(s)" % count]
    return []


def test_master_index_lookalike_docx_is_still_checked(tmp):
    """
    is_master_list_filename matched on the name prefix alone, with no extension
    check — so a DOCUMENT called "Master List Guidelines.docx" was treated as
    the index itself and silently excluded from orphan detection. An unindexed
    document that never appears in the report at all is the worst shape a miss
    can take.
    """
    base = _mini_base(tmp, "lookalike")
    # Copied AFTER the bootstrap, so nothing indexes it: a genuine orphan.
    shutil.copy(_fixture_doc(tmp),
                os.path.join(base, "01 - Identity", "Master List Guidelines.docx"))

    out = run([CHECKER, base]).stdout
    if "Master List Guidelines.docx" not in out:
        return ["an unindexed .docx named like the index was never reported"]
    return []


def test_bullets_sharing_a_definition_are_clean(tmp):
    """
    Word reuses one numbering definition for identical bullet formatting, so two
    bulleted sections sharing one is normal — and there is no number on screen
    for the second section to get wrong. Reporting it as a shared numbering
    sequence accounted for 9 of 19 findings from that check on a real base,
    every one of them false.

    The ordered case must keep firing: that is the defect the check is for.
    """
    out = run([CHECKER, os.path.join(tmp, "detections")]).stdout
    shared = [l for l in out.splitlines() if "share one numbering sequence" in l]
    failures = []
    if not any("KB-060" in l for l in shared):
        failures.append(
            "KB-060 shares a DECIMAL definition and must still be reported")
    if any("KB-061" in l for l in shared):
        failures.append(
            "KB-061 shares a BULLET definition — bullets have no count to "
            "continue, so this is a false positive")
    return failures


def test_languages_json_requires_its_keys(tmp):
    """
    Adding a language is documented as a data edit to languages.json, so a
    half-finished entry is an expected mistake. The module-level tables index
    some keys directly, and a missing one raised a bare KeyError at IMPORT
    time — before main() exists to catch it — which exits 1. That is this
    tool's own "problems found" code, so a broken vocabulary was
    indistinguishable from a completed audit to anything reading the exit
    status.
    """
    sandbox = os.path.join(tmp, "langsandbox")
    shutil.copytree(os.path.abspath(SCRIPTS), sandbox)
    langs_file = os.path.join(sandbox, "languages.json")
    with open(langs_file, encoding="utf-8") as f:
        data = json.load(f)
    data["languages"]["fr"] = {"spreadsheet_name": "Liste Maitresse"}
    with open(langs_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

    proc = run([os.path.join(sandbox, "check_master_list.py"),
                os.path.join(tmp, "detections")])
    out = proc.stdout + proc.stderr
    failures = []
    if proc.returncode != 2:
        failures.append("a half-finished language must exit 2 ('couldn't run'), "
                        "got %d" % proc.returncode)
    if "Traceback" in out:
        failures.append("raised a traceback instead of reporting the problem")
    if "overview_sheet" not in out:
        failures.append("the error message should name the missing key")
    return failures


def test_typed_page_number_is_reported(tmp):
    """
    The footer check credited any footer whose XML contained the substring
    "PAGE" with having a Word automatic page field. A footer reading
    "ACME HOMEPAGE | Knowledge Base" satisfied that while its page number was
    plain typed text, so the violation went unreported.

    Needs its own assertion: against the old checker this miss (-1) and the
    bullet false positive (+1) cancel out, leaving the Formatting violations
    COUNT correct for the wrong reasons.
    """
    out = run([CHECKER, os.path.join(tmp, "detections")]).stdout
    typed = [l for l in out.splitlines()
             if "page field" in l and "KB-063" in l]
    if not typed:
        return ["KB-063's typed page number was not reported; the word PAGE in "
                "the footer text must not pass for a page field"]
    return []


def test_direct_numbering_beats_the_style_name(tmp):
    """
    Word applies a paragraph's own w:numPr over whatever its style says, so a
    paragraph styled "List Number" whose direct numbering points at a bullet
    definition renders as a bullet. Consulting the style NAME first called it
    decimal and reported nothing — and that is how this suite's own
    shared-numbering fixture got away with pinning "numbered" steps to a bullet
    definition for two releases.

    Asserted by name, not only by count: a count can be right for the wrong
    reasons, as this release already demonstrated twice.
    """
    out = run([CHECKER, os.path.join(tmp, "detections")]).stdout
    hits = [l for l in out.splitlines()
            if "KB-064" in l and "not formatted as" in l]
    if not hits:
        return ["KB-064's Step by Step renders as bullets and must be reported "
                "as not a numbered list, whatever its style is named"]
    return []


def test_ambiguous_dates_are_settled_and_disclosed(tmp):
    """
    parse_date_value tried %d/%m/%Y unconditionally, so an English base's
    "03/04/2026" was read as 3 April. The order now follows the base's recorded
    language, and date_is_ambiguous marks the cases where the choice actually
    changes the date — those feed the name-divergence recommendation, which is
    the only thing that consumes this.
    """
    code = (
        "import sys; sys.path.insert(0, %r)\n"
        "import check_master_list as c\n"
        "print(c.parse_date_value('03/04/2026', day_first=True))\n"
        "print(c.parse_date_value('03/04/2026', day_first=False))\n"
        "print(c.date_is_ambiguous('03/04/2026'))\n"
        "print(c.date_is_ambiguous('13/04/2026'))\n"
        "print(c.date_is_ambiguous('05/05/2026'))\n"
        "import datetime\n"
        "print(c.date_is_ambiguous(datetime.date(2026, 4, 3)))\n"
    ) % os.path.abspath(SCRIPTS)

    proc = run(["-c", code])
    got = proc.stdout.split()
    want = ["2026-04-03",   # day-first: 3 April
            "2026-03-04",   # month-first: 4 March
            "True",         # both readings valid and different
            "False",        # 13 can only be a day
            "False",        # both readings agree
            "False"]        # a real date carries no ambiguity
    if got != want:
        return ["expected %s, got %s%s" % (want, got,
                ("  stderr: " + proc.stderr.strip()[:200]) if proc.stderr.strip() else "")]
    return []


def test_numbered_headings_are_recognized(tmp):
    """
    Section matching is startswith() against the bare vocabulary term, so a
    heading reading "2. Prerequisites" was invisible to it: the structure check
    reported the section missing, and the per-section list rule skipped it
    entirely, leaving a numbered document's Step by Step never verified to be a
    numbered list.

    KB-065 is conforming in every respect and must produce no finding at all.
    """
    out = run([CHECKER, os.path.join(tmp, "detections")]).stdout
    hits = [l for l in out.splitlines() if "KB-065" in l]
    if hits:
        return ["KB-065 numbers its headings and is otherwise conforming, but "
                "produced: " + "; ".join(h.strip()[:90] for h in hits[:3])]
    return []


def test_legacy_on_topical_tab_is_not_a_mismatch(tmp):
    """
    Section 7 gives legacy documents two homes in the index: their own Legacy
    tab, or their original topical tab with Status = Legacy. The folder/tab
    check only tolerated the first, so the second was reported as a mismatch
    telling the reader to "move the file or move the row" — undoing an
    arrangement the standard had offered them.

    KB-066 is that arrangement: Status Legacy, file in Legacy/, row on the
    Infrastructure tab.
    """
    out = run([CHECKER, os.path.join(tmp, "detections")]).stdout
    hits = [l for l in out.splitlines()
            if "KB-066" in l and "is filed under category" in l]
    if hits:
        return ["a Legacy-status row on its topical tab is permitted by "
                "Section 7, but was reported: " + hits[0].strip()[:110]]
    return []


def test_dash_inside_the_name_is_allowed(tmp):
    """
    Section 4 constrains the separator BETWEEN THE CODE AND THE NAME: "Use a
    regular hyphen -, never an en/em dash, between the code and the name." The
    check scanned the whole header, so a document whose own title contains an em
    dash was reported as using the wrong separator -- which reads as an
    instruction to strip punctuation out of the title.

    KB-067 has a correct hyphen after the code and an em dash inside its name.
    KB-041 remains the real violation, with the dash AS the separator.
    """
    out = run([CHECKER, os.path.join(tmp, "detections")]).stdout
    failures = []
    if [l for l in out.splitlines() if "KB-067" in l and "plain hyphen" in l]:
        failures.append("an em dash inside the document's own name was reported "
                        "as a wrong code/name separator")
    if not [l for l in out.splitlines() if "KB-041" in l and "plain hyphen" in l]:
        failures.append("KB-041 uses an en dash AS the separator and must still "
                        "be reported")
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
    ("colliding category names", test_colliding_category_names),
    ("Overview extra column", test_overview_extra_column),
    ("--force never indexes the index", test_force_does_not_index_the_index),
    ("language detected from the base", test_language_is_detected),
    ("skipped sheet is counted", test_skipped_sheet_is_counted),
    ("empty Code row reported once", test_empty_code_row_is_reported_once),
    ("Office lock files are not orphans", test_office_lock_files_are_not_orphans),
    ("index look-alike .docx still checked", test_master_index_lookalike_docx_is_still_checked),
    ("shared bullets are clean", test_bullets_sharing_a_definition_are_clean),
    ("typed page number reported", test_typed_page_number_is_reported),
    ("direct numbering beats style name", test_direct_numbering_beats_the_style_name),
    ("numbered headings recognized", test_numbered_headings_are_recognized),
    ("legacy on topical tab is fine", test_legacy_on_topical_tab_is_not_a_mismatch),
    ("dash inside the name is allowed", test_dash_inside_the_name_is_allowed),
    ("ambiguous dates settled and disclosed", test_ambiguous_dates_are_settled_and_disclosed),
    ("languages.json keys validated", test_languages_json_requires_its_keys),
]


def main():
    tmp = tempfile.mkdtemp(prefix="kb-tests-")
    all_failures = {}
    try:
        # Most groups below copy a document out of the detections fixture, so
        # build it ONCE up front and stop here if it can't be built. Letting
        # them run anyway made a broken builder show up as a dozen
        # FileNotFoundErrors pointing at the wrong file.
        built = run([BUILDER, os.path.join(tmp, "detections")])
        if built.returncode != 0:
            print("[FAIL] fixture build (rc=%d): %s"
                  % (built.returncode, built.stderr.strip()[:400]))
            return 1

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
