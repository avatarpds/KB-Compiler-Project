#!/usr/bin/env python3
"""
bootstrap_master_list.py — Scans a knowledge base's folders and creates the
master-index spreadsheet ("Lista Mestra") from the files already there.

This is the WRITE counterpart to check_master_list.py, deliberately kept as a
separate script: the checker is read-only and says so in its report, and that
guarantee is worth more than the convenience of one combined command.

CONSERVATIVE BY DESIGN
----------------------
A base that needs bootstrapping is, by definition, a base that doesn't follow
the standard yet — so most of what could be "inferred" from its files would be
a guess. This script therefore:

- fills in only what it can actually read (code and name from the file name,
  version from the document header when the header is in the standard's shape);
- leaves blank whatever it couldn't read, instead of inventing a value;
- marks every generated row Status = "Em revisão"/"In review", never "Ativo" —
  no row here has been reviewed by a human yet;
- NEVER assigns a new KB code to a document that doesn't already have one in
  its file name. Codes are never recycled (Section 2 of the standard), so
  burning one on a wrong guess is not reversible;
- refuses to overwrite an existing spreadsheet unless --force is passed;
- skips attachment folders (see ATTACHMENT_DIR_SUFFIXES in check_master_list).

Usage:
    python3 bootstrap_master_list.py <base_path> [--lang en|pt] [--owner "Name"] [--output PATH.xlsx] [--force]
"""

import datetime
import glob
import os
import re
import sys

# Same reason as in check_master_list.py: this script's notes contain "—" and
# accented file names, and Windows would fail to encode them the moment output
# is redirected or piped.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

try:
    import openpyxl
except ImportError:
    print("ERROR: missing 'openpyxl'. Install with: pip install -r requirements.txt")
    sys.exit(2)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from check_master_list import (  # noqa: E402
    CONFIG_FILENAME,
    is_master_list_filename,
    MASTER_LIST_NAME_PATTERNS,
    NO_CODE_MARKER,
    is_inside_attachment_folder,
    nfc,
    read_config,
    write_config,
    read_docx_signals,
)

# Canonical column order, shared with the checker's COLUMN_SYNONYMS so the
# spreadsheet this script writes is the one the checker knows how to read.
HEADERS = {
    "pt": ["Código", "Documento", "Arquivo", "Status", "Versão",
           "Data de Criação", "Última Atualização", "Responsável"],
    "en": ["Code", "Document", "File", "Status", "Version",
           "Creation Date", "Last Updated", "Owner"],
}
OVERVIEW_SHEET = {"pt": "Visão Geral", "en": "Overview"}
OVERVIEW_HEADERS = {
    "pt": ["Categoria", "Total", "Ativos", "Em revisão", "Aba"],
    "en": ["Category", "Total", "Active", "In review", "Tab"],
}
STATUS_IN_REVIEW = {"pt": "Em revisão", "en": "In review"}
STATUS_LEGACY = {"pt": "Legado", "en": "Legacy"}
LEGACY_FOLDERS = ("legado", "legacy")
BASE_TITLE = {"pt": "Base de Conhecimento", "en": "Knowledge Base"}
# Documents sitting loose in the base root belong to no folder, so there is no
# category to infer. They are indexed under this tab rather than skipped: the
# checker would otherwise report every one of them as an orphan, leaving the
# two tools contradicting each other about the same files.
UNCATEGORIZED = {"pt": "Sem Categoria", "en": "Uncategorized"}
SPREADSHEET_NAME = {"pt": "Lista Mestra", "en": "Master List"}

INDEXABLE_EXTENSIONS = (".docx", ".pdf", ".pptx", ".xlsx")
CODE_RE = re.compile(r"^(KB-\d{3})\s*-\s*(.+)$", re.IGNORECASE)


# Characters Excel forbids in a sheet name. A folder like "01 - Rede [WAN]" is
# perfectly legal on disk and used to crash the whole bootstrap with a raw
# traceback from openpyxl.
INVALID_SHEET_CHARS = set(r"[]:*?/\\")
MAX_SHEET_NAME = 31


def safe_sheet_name(name, taken):
    """
    Turns a category name into a valid, unique Excel sheet name. Returns
    (sheet_name, changed) so the caller can tell the user what was renamed and
    keep the Overview's "Aba" column pointing at the name that actually exists.
    """
    cleaned = "".join("-" if ch in INVALID_SHEET_CHARS else ch for ch in name).strip()
    cleaned = cleaned or "Categoria"
    changed = cleaned != name

    candidate = cleaned[:MAX_SHEET_NAME].strip()
    if candidate != cleaned:
        changed = True
    # Excel silently renames duplicates, which would leave the Overview's tab
    # reference pointing at a sheet that doesn't exist. Disambiguate ourselves.
    suffix = 2
    while candidate.lower() in taken:
        tail = " (%d)" % suffix
        candidate = cleaned[: MAX_SHEET_NAME - len(tail)].strip() + tail
        suffix += 1
        changed = True
    taken.add(candidate.lower())
    return candidate, changed


def strip_numeric_prefix(folder_name):
    """'01 - Identity' -> 'Identity'. Leaves an unprefixed name untouched."""
    m = re.match(r"^\d+\s*-\s*(.+)$", folder_name.strip())
    return m.group(1).strip() if m else folder_name.strip()


def build_row(base_dir, abs_path, file_value, is_legacy, notes):
    """Builds one index row from a real file, reading only what is actually there."""
    rel_to_base = nfc(os.path.relpath(abs_path, base_dir).replace(os.sep, "/"))
    stem = os.path.splitext(nfc(os.path.basename(abs_path)))[0]
    m = CODE_RE.match(stem)
    if m:
        code, doc_name = m.group(1).upper(), m.group(2).strip()
    else:
        # No code in the file name. We do NOT invent one — codes are never
        # recycled, so a wrong guess is permanent.
        code = NO_CODE_MARKER
        doc_name = stem
        if not is_legacy:
            notes.append(
                f"{rel_to_base}: no KB code in the file name — left as "
                f"'{NO_CODE_MARKER}' for you to assign"
            )

    version = ""
    if abs_path.lower().endswith(".docx"):
        signals = read_docx_signals(abs_path)
        if signals["error"]:
            notes.append(f"{rel_to_base}: couldn't be opened ({signals['error']})")
        elif signals["header_version"]:
            version = signals["header_version"]
        else:
            notes.append(f"{rel_to_base}: no version readable from the header — left blank")
    else:
        notes.append(f"{rel_to_base}: not a .docx, so no version could be read — left blank")

    return {
        "code": code,
        "document": doc_name,
        "file": file_value,
        "legacy": is_legacy,
        "version": version,
        "created": datetime.datetime.fromtimestamp(os.path.getctime(abs_path)),
        "updated": datetime.datetime.fromtimestamp(os.path.getmtime(abs_path)),
    }


def scan(base_dir, lang="en", exclude=()):
    """
    Walks the base and returns {category: [row_dict, ...]} plus a list of notes
    about what could not be determined.

    `exclude` holds absolute paths to leave out — the master index itself, above
    all. A custom-named index (--output "Internal Index.xlsx") is not matched by
    MASTER_LIST_NAME_PATTERNS and .xlsx is indexable, so a re-run with --force
    used to index the index as if it were one of the base's documents.
    """
    categories = {}
    notes = []
    excluded = {os.path.abspath(p) for p in exclude}

    # Root-level documents first, so they head the index and are easy to file.
    root_rows = []
    for fname in sorted(os.listdir(base_dir)):
        abs_path = os.path.join(base_dir, fname)
        if not os.path.isfile(abs_path):
            continue
        if not fname.lower().endswith(INDEXABLE_EXTENSIONS):
            continue
        if fname.startswith("~$") or fname == CONFIG_FILENAME:
            continue
        if is_master_list_filename(nfc(fname)):
            continue
        if os.path.abspath(abs_path) in excluded:
            continue
        row = build_row(base_dir, abs_path, nfc(fname), is_legacy=False, notes=notes)
        notes.append(
            f"{nfc(fname)}: sits in the base root, so it has no category — indexed under "
            f"'{UNCATEGORIZED[lang]}'. Move it into a category folder and update the row."
        )
        root_rows.append(row)
    if root_rows:
        categories.setdefault(UNCATEGORIZED[lang], []).extend(root_rows)

    for entry in sorted(os.listdir(base_dir)):
        full = os.path.join(base_dir, entry)
        if not os.path.isdir(full):
            continue
        category = strip_numeric_prefix(entry)
        # Compare AFTER stripping the numeric prefix: "09 - Legado" is a
        # legacy folder just as much as "Legado".
        is_legacy = category.strip().lower() in LEGACY_FOLDERS
        rows = []

        for root, _dirs, files in os.walk(full):
            for fname in sorted(files):
                abs_path = os.path.join(root, fname)
                rel_to_base = nfc(os.path.relpath(abs_path, base_dir).replace(os.sep, "/"))
                if is_inside_attachment_folder(rel_to_base):
                    continue
                if not fname.lower().endswith(INDEXABLE_EXTENSIONS):
                    continue
                if fname.startswith("~$"):  # Office lock file
                    continue
                if os.path.abspath(abs_path) in excluded:
                    continue

                # The checker expects the full relative path only for Legacy.
                file_value = rel_to_base if is_legacy else nfc(fname)
                rows.append(build_row(base_dir, abs_path, file_value, is_legacy, notes))

        if rows:
            # setdefault/extend, never assignment: two folders can strip to the
            # same category name ("01 - Rede" and "02 - Rede"), and assigning
            # silently dropped every document of whichever folder was scanned
            # first — they then surfaced as orphans in the checker's report,
            # with nothing in this script's output saying they had been lost.
            if category in categories:
                notes.append(
                    "folder '%s' strips to the category name '%s', which another "
                    "folder already uses; their documents were merged into one tab. "
                    "Rename one of the folders if they are meant to be separate "
                    "categories." % (entry, category)
                )
            categories.setdefault(category, []).extend(rows)

    return categories, notes


def build_workbook(categories, lang, owner=""):
    headers = HEADERS[lang]
    wb = openpyxl.Workbook()

    # Resolve every sheet name first, so the Overview can reference names that
    # actually exist in the workbook.
    taken = set()
    sheet_names = {}
    renamed = []
    for category in categories:
        name, changed = safe_sheet_name(category, taken)
        sheet_names[category] = name
        if changed:
            renamed.append((category, name))

    ov = wb.active
    ov.title = OVERVIEW_SHEET[lang]
    ov.cell(row=1, column=1, value=f"{BASE_TITLE[lang]} — bootstrap")
    for i, h in enumerate(OVERVIEW_HEADERS[lang], 1):
        ov.cell(row=2, column=i, value=h)
    for r, (category, rows) in enumerate(categories.items(), start=3):
        sheet = sheet_names[category]
        status_col = "D"
        # Wide fixed range: a range ending at the current row count went
        # stale the moment anyone appended a row.
        first, last = 3, 10000
        quoted = f"'{sheet}'"
        active_label = "Ativo" if lang == "pt" else "Active"
        review_label = STATUS_IN_REVIEW[lang]
        ov.cell(row=r, column=1, value=category)
        # Live formulas rather than frozen numbers: static counts go stale the
        # first time anyone edits a category tab.
        ov.cell(row=r, column=2, value=f"=COUNTA({quoted}!B{first}:B{last})")
        ov.cell(row=r, column=3,
                value=f'=COUNTIF({quoted}!{status_col}{first}:{status_col}{last},"{active_label}")')
        ov.cell(row=r, column=4,
                value=f'=COUNTIF({quoted}!{status_col}{first}:{status_col}{last},"{review_label}")')
        tab_word = "Aba" if lang == "pt" else "Tab"
        ov.cell(row=r, column=5, value=f"→ {tab_word} '{sheet}'")

    for category, rows in categories.items():
        ws = wb.create_sheet(sheet_names[category])
        ws.cell(row=1, column=1, value=f"{BASE_TITLE[lang]} — {category}")
        for i, h in enumerate(headers, 1):
            ws.cell(row=2, column=i, value=h)
        for r, row in enumerate(rows, start=3):
            status = STATUS_LEGACY[lang] if row["legacy"] else STATUS_IN_REVIEW[lang]
            values = [
                row["code"], row["document"], row["file"], status,
                row["version"], row["created"], row["updated"], owner,
            ]
            for i, v in enumerate(values, 1):
                ws.cell(row=r, column=i, value=v)
        for col, width in zip("ABCDEFGH", (12, 45, 45, 14, 10, 18, 20, 20)):
            ws.column_dimensions[col].width = width

    return wb, renamed


def default_target(base_dir, lang):
    return os.path.join(
        base_dir,
        "%s - %s.xlsx" % (SPREADSHEET_NAME[lang], os.path.basename(os.path.abspath(base_dir))),
    )


def choose_target(base_dir, lang):
    """
    Asks where this base's master index should live. Only prompts on an
    interactive terminal; otherwise returns the default silently, so a
    scripted run never blocks. Whatever is chosen is recorded in the base's
    config file and becomes the single insertion point for every later run.
    """
    default = default_target(base_dir, lang)

    if not sys.stdin.isatty():
        return default

    print("\nWhere should this base's master index live?")
    print("  [1] %s   (default, inside the base)" % os.path.basename(default))
    print("  [2] A different file name, inside the base")
    print("  [3] A full path of my choosing (can be outside the base)")

    try:
        choice = input("> ").strip() or "1"
    except EOFError:
        # Nobody to answer, despite isatty() saying otherwise — which is what
        # Windows reports when stdin is redirected from NUL. Fall back to the
        # same default the non-interactive path above uses, instead of
        # aborting a run that asked for nothing unusual.
        print("(no input available — using the default)")
        return default
    except KeyboardInterrupt:
        print("\nAborted.")
        return None

    if choice == "2":
        try:
            name = input("File name (.xlsx): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nAborted.")
            return None
        if not name:
            return None
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        return os.path.join(base_dir, name)

    if choice == "3":
        try:
            path = input("Full path (.xlsx): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nAborted.")
            return None
        if not path:
            return None
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        parent = os.path.dirname(os.path.abspath(path))
        if not os.path.isdir(parent):
            print("ERROR: '%s' is not an existing directory." % parent)
            return None
        return path

    return default


def bootstrap(base_dir, lang="en", owner="", force=False, quiet=False, target=None):
    """Returns the path of the spreadsheet written, or raises on refusal."""
    if target is None:
        target = choose_target(base_dir, lang)
        if target is None:
            raise RuntimeError("no destination chosen; nothing was written.")

    # Refuse if the base ALREADY has any master-index spreadsheet, not just one
    # under this language's file name. Checking only the target name meant that
    # running without --lang against a base indexed in another language quietly
    # created a SECOND spreadsheet beside the first, which the checker then
    # (correctly) refuses to audit at all.
    existing = []
    for pattern in MASTER_LIST_NAME_PATTERNS:
        existing.extend(glob.glob(os.path.join(base_dir, pattern)))
    configured = read_config(base_dir).get("master_index")
    if configured:
        configured_path = (configured if os.path.isabs(configured)
                           else os.path.join(base_dir, configured))
        if os.path.isfile(configured_path):
            existing.append(configured_path)
    existing = sorted(set(existing))
    if existing and not force:
        names = ", ".join(f"'{os.path.basename(e)}'" for e in existing)
        raise FileExistsError(
            f"this base already has a master-index spreadsheet ({names}). Refusing to "
            "create another one beside it — a base must have exactly one. Pass --force "
            "only if you really mean to replace it, and delete the other file first."
        )

    # Never index the index. `existing` already holds any configured path.
    categories, notes = scan(base_dir, lang, exclude=[target] + existing)
    if not categories:
        raise RuntimeError(
            f"No indexable document ({', '.join(INDEXABLE_EXTENSIONS)}) found under "
            f"'{base_dir}'. Nothing to bootstrap."
        )

    wb, renamed = build_workbook(categories, lang, owner)
    wb.save(target)

    # Record the choice so every later run uses THIS spreadsheet, whatever its
    # name or location, instead of guessing from a file-name pattern.
    try:
        stored = os.path.relpath(target, base_dir)
        if stored.startswith(".."):
            stored = os.path.abspath(target)
    except ValueError:
        stored = os.path.abspath(target)
    write_config(base_dir, master_index=stored, language=lang, owner=owner or None)

    if not quiet:
        total = sum(len(r) for r in categories.values())
        print(f"\nCreated: {target}")
        print(f"Registered in {CONFIG_FILENAME} — later runs will use this file.")
        print(f"{total} document(s) across {len(categories)} category tab(s).")
        print("\nEvery row was written with Status = "
              f"'{STATUS_IN_REVIEW[lang]}' — none of them has been reviewed yet.")
        if renamed:
            print("\n--- Category folders renamed to be valid Excel sheet names ---")
            for original, sheet in renamed:
                print(f"- '{original}' -> tab '{sheet}'")
        if notes:
            print(f"\n--- Could not be determined ({len(notes)}) — fill these in manually ---")
            for n in notes:
                print("- " + n)
        print("\nNext step: review the spreadsheet, then run check_master_list.py "
              "against the base.")

    return target


def main():
    args = [a for a in sys.argv[1:]]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 2

    base_dir = args[0]
    force = "--force" in args
    lang = "en"
    if "--lang" in args:
        idx = args.index("--lang")
        value = args[idx + 1] if idx + 1 < len(args) else None
        if value not in HEADERS:
            # Falling back to the default silently used to write a SECOND
            # spreadsheet next to an existing one in another language, which
            # --force couldn't protect against because the new target didn't
            # exist yet.
            print("ERROR: --lang must be one of: %s (got: %r)"
                  % (", ".join(sorted(HEADERS)), value))
            return 2
        lang = value

    if not os.path.isdir(base_dir):
        print(f"ERROR: '{base_dir}' is not a directory.")
        return 2

    owner = ""
    if "--owner" in args:
        idx = args.index("--owner")
        if idx + 1 < len(args):
            owner = args[idx + 1]

    target = None
    if "--output" in args:
        idx = args.index("--output")
        if idx + 1 >= len(args):
            print("ERROR: --output needs a path.")
            return 2
        target = args[idx + 1]
        if not target.lower().endswith(".xlsx"):
            target += ".xlsx"
        if not os.path.isabs(target):
            target = os.path.join(base_dir, target)

    try:
        bootstrap(base_dir, lang=lang, owner=owner, force=force, target=target)
    except (FileExistsError, RuntimeError) as e:
        print(f"ERROR: {e}")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
