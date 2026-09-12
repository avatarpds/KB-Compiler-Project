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
    python3 bootstrap_master_list.py <base_path> [--lang CODE] [--owner "Name"] [--output PATH.xlsx] [--force]

--lang is optional: the language is detected from the documents themselves.
Pass it only to override. The available codes come from languages.json, where
adding a language is a data edit rather than a code change.
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
import docx  # noqa: E402

from check_master_list import (  # noqa: E402
    CONFIG_FILENAME,
    HEADING_STYLE_MARKERS,
    LANGUAGES,
    _style_name,
    is_master_list_filename,
    MASTER_LIST_NAME_PATTERNS,
    NO_CODE_MARKER,
    is_inside_attachment_folder,
    nfc,
    stdin_is_interactive,
    read_config,
    write_config,
    read_docx_signals,
)

# Every localized term comes from languages.json, via the checker — so the
# spreadsheet this script WRITES is by construction the one the checker knows
# how to READ, and adding a language never means editing either script.
HEADERS = {code: v["index_headers"] for code, v in LANGUAGES.items()}
OVERVIEW_SHEET = {code: v["overview_sheet"] for code, v in LANGUAGES.items()}
OVERVIEW_HEADERS = {code: v["overview_headers"] for code, v in LANGUAGES.items()}
STATUS_IN_REVIEW = {code: v["status_written"]["in_review"] for code, v in LANGUAGES.items()}
STATUS_LEGACY = {code: v["status_written"]["legacy"] for code, v in LANGUAGES.items()}
STATUS_ACTIVE = {code: v["status_written"]["active"] for code, v in LANGUAGES.items()}
TAB_WORD = {code: v["tab_word"] for code, v in LANGUAGES.items()}
# A folder is "legacy" if it matches that word in ANY language: a base can
# perfectly well be authored in Portuguese with a folder named "Legacy".
LEGACY_FOLDERS = tuple(sorted({f.lower() for v in LANGUAGES.values()
                               for f in v["legacy_folders"]}))
BASE_TITLE = {code: v["base_title"] for code, v in LANGUAGES.items()}


def detect_language(base_dir, sample=40):
    """
    Infers which language a base is authored in, by reading its documents.

    Section 0 of the standard says to default to whichever language the base
    already uses — but the tooling used to make you declare it with --lang, and
    getting it wrong writes an index whose column headers don't match the
    documents it indexes.

    Scoring is deliberately dumb and explainable: count how many Heading-1
    section titles match each language's section vocabulary. A conforming
    document contributes several hits in exactly one language, so even a
    handful of documents settles it. Ties and empty bases return None, and the
    caller decides rather than this function guessing.

    Returns (code_or_None, evidence_dict).
    """
    scores = {code: 0 for code in LANGUAGES}
    seen = 0
    for root, _dirs, files in os.walk(base_dir):
        if is_inside_attachment_folder(
                nfc(os.path.relpath(os.path.join(root, "x"), base_dir)).replace(os.sep, "/")):
            continue
        for fname in sorted(files):
            if not fname.lower().endswith(".docx") or fname.startswith("~$"):
                continue
            if seen >= sample:
                break
            path = os.path.join(root, fname)
            try:
                d = docx.Document(path)
            except Exception:
                continue
            seen += 1
            headings = [p.text.strip().lower() for p in d.paragraphs
                        if any(m in _style_name(p) for m in HEADING_STYLE_MARKERS)]
            for code, vocab in LANGUAGES.items():
                for terms in vocab["sections"].values():
                    if any(h.startswith(t) for h in headings for t in terms):
                        scores[code] += 1

    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    evidence = {"documents_read": seen, "scores": dict(ranked)}
    if seen == 0 or ranked[0][1] == 0:
        return None, evidence
    if len(ranked) > 1 and ranked[0][1] == ranked[1][1]:
        return None, evidence          # a tie is not a detection
    return ranked[0][0], evidence
# Documents sitting loose in the base root belong to no folder, so there is no
# category to infer. They are indexed under this tab rather than skipped: the
# checker would otherwise report every one of them as an orphan, leaving the
# two tools contradicting each other about the same files.
UNCATEGORIZED = {code: v["uncategorized"] for code, v in LANGUAGES.items()}
SPREADSHEET_NAME = {code: v["spreadsheet_name"] for code, v in LANGUAGES.items()}

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
    cleaned = cleaned or "Categoria"   # last-resort name for an unnameable folder
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
        active_label = STATUS_ACTIVE[lang]
        review_label = STATUS_IN_REVIEW[lang]
        ov.cell(row=r, column=1, value=category)
        # Live formulas rather than frozen numbers: static counts go stale the
        # first time anyone edits a category tab.
        ov.cell(row=r, column=2, value=f"=COUNTA({quoted}!B{first}:B{last})")
        ov.cell(row=r, column=3,
                value=f'=COUNTIF({quoted}!{status_col}{first}:{status_col}{last},"{active_label}")')
        ov.cell(row=r, column=4,
                value=f'=COUNTIF({quoted}!{status_col}{first}:{status_col}{last},"{review_label}")')
        tab_word = TAB_WORD[lang]
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

    if not stdin_is_interactive():
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


def bootstrap(base_dir, lang=None, owner="", force=False, quiet=False, target=None):
    """
    Returns the path of the spreadsheet written, or raises on refusal.

    lang=None means "read the base and work it out" — Section 0 says to follow
    whichever language the base already uses, and making the caller declare it
    was the tooling contradicting its own standard.
    """
    detection = None
    if lang is None:
        lang, detection = detect_language(base_dir)
        if lang is None:
            # Inconclusive. Say so plainly instead of quietly picking one and
            # writing an index whose headers don't match the documents.
            lang = "en"
            print("Could not tell which language this base is written in "
                  "(%d document(s) read, scores: %s)." % (detection["documents_read"],
                                                          detection["scores"]))
            print("Defaulting to '%s'. Pass --lang to say otherwise: %s"
                  % (lang, ", ".join(sorted(LANGUAGES))))
        elif not quiet:
            print("Language detected: %s (%s). %d document(s) read, scores: %s"
                  % (lang, LANGUAGES[lang]["display_name"],
                     detection["documents_read"], detection["scores"]))
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
    # None means "detect from the base". An explicit --lang always overrides.
    lang = None
    if "--lang" in args:
        idx = args.index("--lang")
        value = args[idx + 1] if idx + 1 < len(args) else None
        if value not in HEADERS:
            # Falling back to the default silently used to write a SECOND
            # spreadsheet next to an existing one in another language, which
            # --force couldn't protect against because the new target didn't
            # exist yet. The valid set comes from languages.json, so adding a
            # language makes it accepted here with no code change.
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
