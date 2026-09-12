#!/usr/bin/env python3
"""
verify_fixes.py — Reproduces, one scenario per function, each defect fixed in
0.8.4, and asserts the corrected behavior.

This is deliberately SEPARATE from tests/run_tests.py. The suite guards the
tooling going forward; this file is the evidence trail for one release: each
scenario states what the old code did and what the new code must do, so the
claims in CHANGELOG.md can be re-checked by running one command rather than
taken on trust.

Usage:
    python3 verify_fixes.py [<path to skills/kb-compiler>]

Defaults to locating the skill folder three levels up from this file, which is
where it sits inside the repository.

Exit code: 0 if every scenario behaves as 0.8.4 requires, 1 otherwise.
"""

import os
import re
import subprocess
import sys
import tempfile

import docx
import openpyxl
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SKILL = os.path.abspath(os.path.join(HERE, "..", "..", "..",
                                             "skills", "kb-compiler"))
SKILL = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SKILL
CHECKER = os.path.join(SKILL, "scripts", "check_master_list.py")
BOOTSTRAP = os.path.join(SKILL, "scripts", "bootstrap_master_list.py")

HISTORY_COLUMNS = ["Author", "Version", "Date", "Change Description"]
INDEX_COLUMNS = ["Code", "Document", "File", "Status", "Version",
                 "Creation Date", "Last Updated", "Owner"]


def run(args):
    """Runs a script with a real empty pipe on stdin, never DEVNULL.

    On Windows, stdin redirected from NUL is reported as a TERMINAL by
    isatty(), so DEVNULL would send the script down the interactive path — the
    very defect scenario 2 covers.
    """
    return subprocess.run(
        [sys.executable] + args,
        capture_output=True, text=True, encoding="utf-8", input="", timeout=120,
    )


def sections(output):
    return {m.group(1): int(m.group(2))
            for m in re.finditer(r"^=== (.+?) \((\d+)\) ===$", output, re.M)}


def count(output, prefix):
    for title, n in sections(output).items():
        if title.startswith(prefix):
            return n
    return None


def make_doc(path, header, title):
    """A document conforming to the standard apart from what a caller changes."""
    d = docx.Document()
    section = d.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.75)
    section.left_margin = section.right_margin = Inches(0.75)
    section.header.paragraphs[0].text = header

    # A real PAGE/NUMPAGES field pair, not typed-in text: the standard requires
    # Word's automatic page fields, and the checker is right to say so.
    footer = section.footer.paragraphs[0]
    footer.text = ""
    footer.add_run("Fixture Team  |  Knowledge Base\tPage ")
    for instruction in (" PAGE ", " NUMPAGES "):
        field_run = footer.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = instruction
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        value = OxmlElement("w:t")
        value.text = "1"
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        for element in (begin, instr, separate, value, end):
            field_run._r.append(element)
        if instruction == " PAGE ":
            footer.add_run(" of ")

    paragraph = d.add_paragraph()
    run_ = paragraph.add_run(title)
    run_.bold = True
    run_.font.size = Pt(22)
    run_.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

    d.add_heading("Purpose", 1)
    d.add_heading("Prerequisites", 1)
    d.add_paragraph("A prerequisite", style="List Bullet")
    d.add_heading("Step by Step", 1)
    d.add_paragraph("Do the thing", style="List Number")
    d.add_heading("Version History", 1)
    table = d.add_table(rows=2, cols=4)
    for i, value in enumerate(HISTORY_COLUMNS):
        table.rows[0].cells[i].text = value
    for i, value in enumerate(["Lucas Souza", "1.0", "Sep/2026", "Created"]):
        table.rows[1].cells[i].text = value

    os.makedirs(os.path.dirname(path), exist_ok=True)
    d.save(path)


def make_index(base, rows, tab="Cat", overview_note=False, tab_name=None):
    """Writes a minimal master index with an Overview mapping one category."""
    wb = openpyxl.Workbook()
    ov = wb.active
    ov.title = "Overview"
    ov.cell(row=1, column=1, value="Knowledge Base — Probe")
    for i, h in enumerate(["Category", "Total", "Active", "In review", "Tab"], 1):
        ov.cell(row=2, column=i, value=h)
    sheet = tab_name or tab
    ov.cell(row=3, column=1, value=tab)
    ov.cell(row=3, column=5, value="→ Tab '%s'" % sheet)
    if overview_note:
        # A human typing a note into a column past "Tab". openpyxl pads every
        # row out to the sheet's widest column, so this used to become row[-1].
        ov.cell(row=3, column=7, value="note typed by a human")

    ws = wb.create_sheet(sheet)
    ws.cell(row=1, column=1, value="Knowledge Base — %s" % tab)
    for i, h in enumerate(INDEX_COLUMNS, 1):
        ws.cell(row=2, column=i, value=h)
    for r, row in enumerate(rows, 3):
        for i, value in enumerate(row, 1):
            ws.cell(row=r, column=i, value=value)
    path = os.path.join(base, "Master List - Probe.xlsx")
    wb.save(path)
    return path


def fresh(name):
    base = os.path.join(tempfile.mkdtemp(prefix="kbverify-"), name)
    os.makedirs(base)
    return base


# ---------------------------------------------------------------------------


def scenario_utf8_redirect():
    """
    BEFORE: redirecting the report on Windows raised UnicodeEncodeError on the
            first section title containing "->", after two lines of output,
            exiting 1 — which reads as "problems found", not "crashed".
    AFTER:  the full report is written, stderr is empty.
    """
    base = fresh("redirect")
    make_doc(os.path.join(base, "Cat", "KB-001 - Alpha.docx"),
             "KB-001 - Alpha\tv1.0", "KB-001 - Alpha")
    make_index(base, [("KB-001", "Alpha", "KB-001 - Alpha.docx", "Active",
                       "1.0", None, None, "Lucas Souza")])
    proc = run([CHECKER, base])
    problems = []
    if proc.stderr.strip():
        problems.append("stderr is not empty: %r" % proc.stderr.strip()[:200])
    if "SUMMARY" not in proc.stdout:
        problems.append("the report has no SUMMARY line — it stopped early")
    # A LOWER BOUND, not an exact count: what this scenario asserts is that the
    # report ran to completion, and the number of sections grows with every
    # release that adds a detection. Pinning it exactly made this scenario fail
    # on 0.9.0 for a reason that had nothing to do with the defect it covers.
    if len(sections(proc.stdout)) < 13:
        problems.append("expected at least 13 report sections, got %d — the "
                        "report was truncated" % len(sections(proc.stdout)))
    return problems


def scenario_non_interactive():
    """
    BEFORE: with no index and a non-terminal stdin, Windows' isatty() said
            "terminal", so the script printed an interactive menu and aborted
            instead of printing the guidance.
    AFTER:  the guidance is printed and the exit code is 2.
    """
    base = fresh("nonint")
    make_doc(os.path.join(base, "Cat", "KB-001 - Alpha.docx"),
             "KB-001 - Alpha\tv1.0", "KB-001 - Alpha")
    proc = run([CHECKER, base])
    problems = []
    if proc.returncode != 2:
        problems.append("expected exit 2, got %d" % proc.returncode)
    if "bootstrap_master_list.py" not in proc.stdout:
        problems.append("the guidance never mentions bootstrap_master_list.py")
    if "Traceback" in proc.stderr:
        problems.append("a missing index still produces a traceback")
    return problems


def scenario_colliding_categories():
    """
    BEFORE: "01 - Rede" and "02 - Rede" both strip to "Rede"; the scan assigned
            instead of merging, so the first folder's documents vanished from
            the index and the checker then reported them as orphans.
    AFTER:  both documents are indexed, the merge is reported, no orphans.
    """
    base = fresh("collide")
    make_doc(os.path.join(base, "01 - Rede", "KB-001 - Um.docx"),
             "KB-001 - Um\tv1.0", "KB-001 - Um")
    make_doc(os.path.join(base, "02 - Rede", "KB-002 - Dois.docx"),
             "KB-002 - Dois\tv1.0", "KB-002 - Dois")

    proc = run([BOOTSTRAP, base])
    problems = []
    if proc.returncode != 0:
        return ["bootstrap failed (rc=%d): %s" % (proc.returncode, proc.stdout[-200:])]
    if "strips to the category name" not in proc.stdout:
        problems.append("the merge of two folders was not reported to the user")

    sheets = [f for f in os.listdir(base) if f.lower().endswith(".xlsx")]
    wb = openpyxl.load_workbook(os.path.join(base, sheets[0]))
    indexed = [str(row[2]) for name in wb.sheetnames[1:]
               for row in list(wb[name].iter_rows(values_only=True))[2:]
               if row and row[0]]
    for expected in ("KB-001 - Um.docx", "KB-002 - Dois.docx"):
        if expected not in indexed:
            problems.append("'%s' was dropped from the index" % expected)

    orphans = count(run([CHECKER, base]).stdout, "Orphaned files")
    if orphans:
        problems.append("colliding categories left %s orphan(s)" % orphans)
    return problems


def scenario_spaced_header():
    """
    BEFORE: a header using a space where the tab belongs ("Alpha v1.0") was
            reported as a NAME divergence, and the recommendation told you to
            propagate "Alpha v1.0" into the file name and the title.
    AFTER:  reported as a missing tab separator, with no name divergence.
    """
    base = fresh("spaced")
    make_doc(os.path.join(base, "Cat", "KB-001 - Alpha.docx"),
             "KB-001 - Alpha v1.0", "KB-001 - Alpha")
    make_index(base, [("KB-001", "Alpha", "KB-001 - Alpha.docx", "Active",
                       "1.0", None, None, "Lucas Souza")])
    out = run([CHECKER, base]).stdout
    problems = []
    if count(out, "Header missing the tab separator") != 1:
        problems.append("missing-tab not detected (got %s)"
                        % count(out, "Header missing the tab separator"))
    if count(out, "Name divergence") != 0:
        problems.append("still misreported as a name divergence (%s)"
                        % count(out, "Name divergence"))
    return problems


def scenario_name_ending_in_version():
    """
    NEGATIVE control for the scenario above: a document whose name legitimately
    ends in something version-shaped must keep its own name. The header there
    genuinely carries no version, which 0.8.4 reports separately.
    """
    base = fresh("officev2")
    make_doc(os.path.join(base, "Cat", "KB-001 - Migracao Office v2.docx"),
             "KB-001 - Migracao Office v2", "KB-001 - Migracao Office v2")
    make_index(base, [("KB-001", "Migracao Office v2",
                       "KB-001 - Migracao Office v2.docx", "Active",
                       "1.0", None, None, "Lucas Souza")])
    out = run([CHECKER, base]).stdout
    problems = []
    if count(out, "Header missing the tab separator") != 0:
        problems.append("a legitimate name ending in 'v2' was called a glued header")
    if count(out, "Name divergence") != 0:
        problems.append("a legitimate name ending in 'v2' produced a name divergence")
    if count(out, "Header missing its version") != 1:
        problems.append("the genuinely absent header version was not reported (%s)"
                        % count(out, "Header missing its version"))
    return problems


def scenario_header_without_version():
    """
    BEFORE: a header carrying no version was invisible. The version comparison
            drops empty sources first, so {None, "1.0", "1.0"} collapsed to one
            value and reported nothing — exactly where the "header must match
            the Version History" rule stops being enforceable.
    AFTER:  reported under its own heading.
    """
    base = fresh("noversion")
    make_doc(os.path.join(base, "Cat", "KB-001 - Alpha.docx"),
             "KB-001 - Alpha", "KB-001 - Alpha")
    make_index(base, [("KB-001", "Alpha", "KB-001 - Alpha.docx", "Active",
                       "1.0", None, None, "Lucas Souza")])
    out = run([CHECKER, base]).stdout
    if count(out, "Header missing its version") != 1:
        return ["a header with no version was not reported (%s)"
                % count(out, "Header missing its version")]
    return []


def scenario_unreadable_document():
    """
    BEFORE: one .docx that cannot be parsed produced THREE findings — a broken
            reference claiming a file that plainly exists doesn't, plus a
            structure and a formatting violation.
    AFTER:  exactly one finding, under its own heading, and SUMMARY says 1.
    """
    base = fresh("corrupt")
    os.makedirs(os.path.join(base, "Cat"))
    with open(os.path.join(base, "Cat", "KB-001 - Corrompido.docx"), "wb") as f:
        f.write(b"\xd0\xcf\x11\xe0 not a docx at all")
    make_index(base, [("KB-001", "Corrompido", "KB-001 - Corrompido.docx",
                       "Active", "1.0", None, None, "Lucas Souza")])
    out = run([CHECKER, base]).stdout
    problems = []
    if count(out, "Unreadable documents") != 1:
        problems.append("not reported as unreadable (%s)"
                        % count(out, "Unreadable documents"))
    for other in ("Broken references", "Structure violations",
                  "Formatting violations"):
        if count(out, other):
            problems.append("also counted under '%s' (%s)" % (other, count(out, other)))
    total = re.search(r"SUMMARY: (\d+) problem", out)
    if not total or int(total.group(1)) != 1:
        problems.append("expected SUMMARY 1, got %s"
                        % (total.group(1) if total else "none"))
    return problems


def scenario_overview_extra_column():
    """
    BEFORE: the Overview's category->tab mapping was read from row[-1]. A note
            typed past the "Tab" column became the last cell, the mapping was
            lost, and truncated tab names produced false folder/tab mismatches.
    AFTER:  the mapping survives; no mismatch.
    """
    base = fresh("overview")
    folder = "01 - Identity Access and Corporate Governance Board"
    category = folder[5:]
    make_doc(os.path.join(base, folder, "KB-001 - Alpha.docx"),
             "KB-001 - Alpha\tv1.0", "KB-001 - Alpha")
    make_index(base,
               [("KB-001", "Alpha", "KB-001 - Alpha.docx", "Active",
                 "1.0", None, None, "Lucas Souza")],
               tab=category, tab_name=category[:31].strip(), overview_note=True)
    out = run([CHECKER, base]).stdout
    if count(out, "Folder/tab mismatch"):
        return ["an extra Overview column produced %s false mismatch(es)"
                % count(out, "Folder/tab mismatch")]
    return []


def scenario_force_self_index():
    """
    BEFORE: a custom-named index matches none of the built-in name patterns and
            .xlsx is indexable, so a --force re-run added the index to itself.
    AFTER:  the index is never one of its own rows.
    """
    base = fresh("forceindex")
    make_doc(os.path.join(base, "01 - Cat", "KB-001 - Alpha.docx"),
             "KB-001 - Alpha\tv1.0", "KB-001 - Alpha")
    if run([BOOTSTRAP, base, "--output", "Internal Index.xlsx"]).returncode != 0:
        return ["bootstrap with --output failed"]
    if run([BOOTSTRAP, base, "--output", "Internal Index.xlsx",
            "--force"]).returncode != 0:
        return ["bootstrap --force failed"]
    wb = openpyxl.load_workbook(os.path.join(base, "Internal Index.xlsx"))
    indexed = [str(row[2]) for name in wb.sheetnames[1:]
               for row in list(wb[name].iter_rows(values_only=True))[2:]
               if row and row[0]]
    if any("Internal Index" in name for name in indexed):
        return ["the index indexed itself: %s" % indexed]
    return []


def scenario_clean_base_is_clean():
    """
    The counterweight to every detection above: a conforming base must still
    come back with zero problems and exit 0. A checker that finds something
    everywhere is as useless as one that finds nothing.
    """
    base = fresh("clean")
    make_doc(os.path.join(base, "01 - Cat", "KB-001 - Alpha.docx"),
             "KB-001 - Alpha\tv1.0", "KB-001 - Alpha")
    if run([BOOTSTRAP, base]).returncode != 0:
        return ["bootstrap failed on a clean base"]
    proc = run([CHECKER, base])
    problems = []
    if proc.returncode != 0:
        problems.append("expected exit 0 on a clean base, got %d" % proc.returncode)
    total = re.search(r"SUMMARY: (\d+) problem", proc.stdout)
    if not total or int(total.group(1)) != 0:
        problems.append("expected 0 problems, got %s"
                        % (total.group(1) if total else "no summary"))
    return problems


SCENARIOS = [
    ("UTF-8 report survives a redirect", scenario_utf8_redirect),
    ("non-interactive stdin prints guidance", scenario_non_interactive),
    ("colliding category names lose nothing", scenario_colliding_categories),
    ("space-separated header is a tab problem", scenario_spaced_header),
    ("name ending in 'v2' keeps its name", scenario_name_ending_in_version),
    ("header with no version is reported", scenario_header_without_version),
    ("unreadable document counts once", scenario_unreadable_document),
    ("extra Overview column keeps the mapping", scenario_overview_extra_column),
    ("--force never indexes the index", scenario_force_self_index),
    ("a clean base stays clean", scenario_clean_base_is_clean),
]


def main():
    print("skill folder: %s\n" % SKILL)
    if not os.path.isfile(CHECKER):
        print("ERROR: %s not found. Pass the path to skills/kb-compiler." % CHECKER)
        return 2

    failed = {}
    for name, fn in SCENARIOS:
        try:
            problems = fn()
        except Exception as e:  # a crashing scenario is a failing scenario
            problems = ["raised %s: %s" % (type(e).__name__, e)]
        if problems:
            failed[name] = problems
        print("[%s] %s" % ("FAIL" if problems else " OK ", name))

    if failed:
        print("\nFAILED:")
        for name, problems in failed.items():
            for problem in problems:
                print("  - [%s] %s" % (name, problem))
        return 1

    print("\nPASSED: %d scenarios. Every 0.8.4 fix behaves as the changelog "
          "describes." % len(SCENARIOS))
    return 0


if __name__ == "__main__":
    sys.exit(main())
