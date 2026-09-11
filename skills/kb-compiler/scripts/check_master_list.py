#!/usr/bin/env python3
"""
check_master_list.py — Cross-check of a KB base's master-index spreadsheet
("Lista Mestra").

This script is FIXED (not rewritten on every run) so the consistency check
always runs the same way, with the same recommendation criteria — instead of
depending on a freshly generated script that could vary from one audit to
the next.

WHAT IT CHECKS
--------------
1. Broken references: every master-index row whose "File" column points to a
   file that doesn't actually exist in the folder.
2. "Orphaned" files: real files in the base that don't appear in any
   master-index row (files inside a folder whose name ends with "- Files",
   e.g. "KB-017 - Autopilot - Files/", are treated as attachments of that KB
   document and are never flagged this way — see ATTACHMENT_DIR_SUFFIXES).
3. Duplicate/shared codes: two different rows using the same code (except
   the "—" marker, which can repeat freely for documents without a code).
4. Name divergence: for every .docx document, compares the name derived from
   4 sources — file name, title (1st line of the body), header, and the
   master-index spreadsheet's "Document" column — and flags when they don't
   match.
5. Version divergence: compares the header's version, the last row of the
   Version History table, and the master-index spreadsheet's "Version"
   column.
6. Version History table in the legacy format (No. | Revision Date |
   Revision | Reviewer, with no semantic-version column) — reported as its
   own item, so it doesn't produce a false version divergence by comparing a
   date against a version number.
7. Header missing the tab separator between name and version (e.g. "Namev1.2"
   glued together instead of "Name<TAB>v1.2") — reported as its own item, so
   it doesn't show up disguised as a generic version divergence.

In addition, at the end of the report, an INFORMATIONAL section (never
counted as a problem) lists files whose name contains a non-ASCII character
(accents, etc.) — a portability map for a possible future migration, not a
base inconsistency.

Every file-name comparison (master-index spreadsheet vs. real files) is done
after normalizing the strings to Unicode NFC (see the nfc() function), so
that the same name saved as NFC and as NFD isn't flagged as a broken
reference just because the underlying bytes differ.

RECOMMENDATION ON NAME DIVERGENCE
----------------------------------
The script does NOT decide or apply the fix on its own. It points to a
recommendation based on which source was changed most recently:

- Compares the .docx file's modification date (file mtime, which reflects
  the last time the title/header were edited) with the date in the
  master-index spreadsheet's "Last Updated" column (parsed whether it's a
  real Excel date or plain text — see parse_date_value()).
- The most recently changed source is suggested as the "recommended" name —
  but the report always lists EVERY variant found, with its origin and date,
  so the person running it can decide which one to keep everywhere (file,
  title, header, and master-index spreadsheet).
- If the dates don't allow a decision (e.g. equal, or unavailable), the
  script marks the divergence as "no automatic recommendation — manual
  decision needed" instead of guessing. If the spreadsheet's "Last Updated"
  cell has a value but it couldn't be parsed as a date at all, the report
  says so explicitly instead of silently falling back to "unavailable".

COLUMN LOOKUP AND EXIT CODE
----------------------------
Master-index spreadsheet rows are read by matching each column's NAME (see
COLUMN_SYNONYMS), not by a fixed position — the header row itself is found
by scanning for a recognizable "Code"/"Código" column rather than assuming
it's always row 2. A sheet with no recognizable header row is skipped with a
WARNING rather than silently read as if its columns were in the usual order.

The process exit code reflects the result: 0 when no problems were found, 1
when at least one was — so this script can gate a batch edit or a CI-style
check on a clean run instead of always reporting success regardless of what
it found.

HOW TO USE
----------
    python3 check_master_list.py <base_path> [<spreadsheet_name.xlsx>]

- <base_path>: the knowledge base's root directory (contains the category
  subfolders plus the master-index spreadsheet).
- <spreadsheet_name.xlsx>: optional. If omitted, the script looks for the
  single matching master-index spreadsheet file at the base's root (see
  MASTER_LIST_NAME_PATTERNS below for the recognized naming conventions).

Requirements: python-docx, openpyxl (pip install python-docx openpyxl).

OUTPUT
------
A text report on the console, grouped by problem type, always ending with a
summary of how many items of each type were found. Nothing is changed in any
file — this script is read-only/diagnostic only.

NOTE ON LANGUAGE: this script's own docstring, comments, and printed report
are in English so the tool itself is easy to read, maintain, and share
across teams. The KB base it inspects can be in any target document
language (see the kb-compiler skill's Section 0/6) — the patterns used to
recognize structural elements (sheet names, column headers) below
deliberately match BOTH the Portuguese terms this plugin was originally
built around (Hasbro Brazil's live base) and their common English
equivalents, so the same fixed script keeps working for either without a
fork. If a base uses yet another language for these structural terms, widen
the matching sets below rather than writing a new script.
"""

import sys
import os
import glob
import json
import re
import datetime
import unicodedata

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:
    print("ERROR: missing 'python-docx'. Install with: pip install -r requirements.txt")
    sys.exit(2)

try:
    import openpyxl
except ImportError:
    print("ERROR: missing 'openpyxl'. Install with: pip install -r requirements.txt")
    sys.exit(2)


# Overview/summary tab names to skip when scanning category tabs — matches
# the Portuguese base this plugin was built for and the common English
# equivalent.
SKIP_SHEETS = {"Visão Geral", "Visao Geral", "Overview"}

# Master-index spreadsheet naming conventions to auto-detect, tried in order.
MASTER_LIST_NAME_PATTERNS = ("Lista Mestra*.xlsx", "Master List*.xlsx", "Master Index*.xlsx")

NO_CODE_MARKER = "—"

# Status values that mark a document as superseded. Structure checks are
# skipped for these.
LEGACY_STATUS_LABELS = {"legacy", "legado"}

# ---------------------------------------------------------------------------
# Formatting constants — these mirror Section 4 of SKILL.md ("Exact Word
# formatting"). They live here so the spec and the checker can be compared
# side by side; if Section 4 changes, change these too.
# ---------------------------------------------------------------------------
EXPECTED_MARGIN_INCHES = 0.75
EXPECTED_TITLE_SIZE_PT = 22.0
EXPECTED_TITLE_COLOR = "1F3864"
EXPECTED_CALLOUT_SHADING = "FFF3CD"
EXPECTED_CALLOUT_TEXT_COLOR = "7B5C00"
CALLOUT_PREFIX = "⚠"
HEADING_STYLE_MARKERS = ("heading 1", "título 1", "titulo 1")

# Sections whose body must be a list, and which kind. Checked only when the
# section actually has body content — an empty section can't be judged.
BULLET_SECTIONS = ("prerequisites", "pré-requisitos", "pre-requisitos",
                   "verification", "verificação", "verificacao")
NUMBERED_SECTIONS = ("step by step", "passo a passo")
DASHES = ("\u2013", "\u2014")  # en dash, em dash — the standard requires a plain hyphen

# A folder whose name ends with one of these (case-insensitive) is treated as
# an attachment/supporting-files folder for a KB document (e.g.
# "KB-017 - Autopilot - Files/"), not a set of independently indexable items.
# Files inside it are never flagged as "orphaned" even though they don't have
# their own master-index row — they're dependencies of the KB they sit under,
# not documents in their own right.
ATTACHMENT_DIR_SUFFIXES = ("- files",)

# Column names recognized in the master-index spreadsheet's header row, in
# Portuguese (this plugin's original base) and English, so rows are read by
# column NAME instead of a fixed position — a column inserted or reordered in
# the spreadsheet won't silently shift every other column's data.
COLUMN_SYNONYMS = {
    "code": ("código", "codigo", "code"),
    "document": ("documento", "document"),
    "file": ("arquivo", "file"),
    "status": ("status",),
    "version": ("versão", "versao", "version"),
    "creation_date": ("data de criação", "data de criacao", "creation date"),
    "last_updated": ("última atualização", "ultima atualizacao", "last updated"),
    "owner": ("responsável", "responsavel", "owner"),
}


def nfc(text):
    """
    Normalize to Unicode NFC (composed form). Visually identical file names
    can be saved as NFC or NFD (common when a file passes through
    macOS/OneDrive/different OSes) — e.g. an accented character as a single
    code point (NFC) vs. a base letter + combining accent (NFD) are
    different bytes and won't match in a naive string comparison, producing
    false "broken reference" positives. Every string used as a file-name
    comparison key goes through this first.
    """
    if not text:
        return text
    return unicodedata.normalize("NFC", text)


# Per-base configuration written by bootstrap_master_list.py. It records where
# this base's master index lives (and the setup answers that go with it), so
# every later run uses the SAME spreadsheet instead of re-deriving it from a
# file-name pattern. That is what makes a custom name or a location outside the
# base folder work as a continuous insertion point.
CONFIG_FILENAME = ".kb-compiler.json"


def read_config(base_dir):
    """Returns the base's config dict, or {} when there is none / it's unreadable."""
    path = os.path.join(base_dir, CONFIG_FILENAME)
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (ValueError, OSError):
        return {}


def write_config(base_dir, **values):
    """Merges values into the base's config file and returns the merged dict."""
    data = read_config(base_dir)
    data.update({k: v for k, v in values.items() if v is not None})
    path = os.path.join(base_dir, CONFIG_FILENAME)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")
    return data


def resolve_configured_master_list(base_dir):
    """
    Returns the configured master-index path, or None. A configured path that
    no longer exists is an explicit error rather than a silent fallback to
    pattern matching — falling back would audit a different file than the one
    the base was set up to use.
    """
    configured = read_config(base_dir).get("master_index")
    if not configured:
        return None
    path = configured if os.path.isabs(configured) else os.path.join(base_dir, configured)
    if not os.path.isfile(path):
        raise FileNotFoundError(
            "%s points at '%s', which no longer exists. Fix the path in that file, "
            "or delete it to fall back to auto-detection." % (CONFIG_FILENAME, configured)
        )
    return path


def find_master_list(base_dir, explicit_name=None):
    # An explicitly passed name always wins; then the base's own configuration;
    # only then pattern matching.
    if not explicit_name:
        configured = resolve_configured_master_list(base_dir)
        if configured:
            return configured

    if explicit_name:
        path = os.path.join(base_dir, explicit_name)
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Spreadsheet '{explicit_name}' not found in {base_dir}")
        return path

    # Accumulate matches across EVERY naming pattern before deciding. Breaking
    # on the first pattern that matched meant a base holding both a
    # "Lista Mestra" and a "Master List" silently audited only one of them —
    # a clean report you would have believed.
    candidates = []
    for pattern in MASTER_LIST_NAME_PATTERNS:
        candidates.extend(glob.glob(os.path.join(base_dir, pattern)))
    candidates = sorted(set(candidates))

    if not candidates:
        raise FileNotFoundError(
            f"No master-index spreadsheet found in {base_dir} (tried: "
            f"{', '.join(MASTER_LIST_NAME_PATTERNS)}). Pass the file name explicitly "
            "as the second argument."
        )
    if len(candidates) > 1:
        raise RuntimeError(
            f"More than one master-index spreadsheet found in {base_dir}: "
            f"{candidates}. Pass the one to use as the second argument."
        )
    return candidates[0]


def is_master_list_filename(basename):
    for pattern in MASTER_LIST_NAME_PATTERNS:
        prefix = pattern.split("*")[0]
        if basename.startswith(prefix):
            return True
    return False


def is_inside_attachment_folder(rel_path):
    """
    True if any directory component of rel_path (everything but the file's
    own name) ends with one of ATTACHMENT_DIR_SUFFIXES — e.g. a file under
    "KB-017 - Autopilot - Files/GetAutoPilot/tool.zip" is an attachment of
    KB-017, not a standalone item that belongs in the master-index
    spreadsheet on its own.
    """
    parts = rel_path.split("/")[:-1]
    for part in parts:
        if part.strip().lower().endswith(ATTACHMENT_DIR_SUFFIXES):
            return True
    return False


def find_header_row(rows):
    """
    Scans a sheet's rows (as returned by iter_rows(values_only=True)) for the
    one that is the column-header row, identified by having a cell that
    matches the "code" column's synonyms — rather than assuming it's always
    row 2. Returns (row_index, lowered_cell_values) or (None, None) if no
    such row is found.
    """
    for i, row in enumerate(rows):
        if not row:
            continue
        cells = [str(c).strip().lower() if c is not None else "" for c in row]
        if any(c in COLUMN_SYNONYMS["code"] for c in cells):
            return i, cells
    return None, None


def map_columns(header_cells):
    """Maps each COLUMN_SYNONYMS key to the column index found in header_cells."""
    mapping = {}
    for key, synonyms in COLUMN_SYNONYMS.items():
        for i, c in enumerate(header_cells):
            if c in synonyms:
                mapping[key] = i
                break
    return mapping


def parse_date_value(value):
    """
    Best-effort parse of a spreadsheet date-like cell into a date, whether
    it's a real datetime/date (the normal case for an Excel date cell) or
    plain text (e.g. a "Last Updated" column typed in as a string instead of
    a real date). Returns None if the value is empty or couldn't be parsed
    as a date in any of the tried formats.
    """
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def index_real_files(base_dir):
    """
    Returns (by_basename, by_relative_path) for every real file. Keys are
    normalized to NFC (see nfc()) so the comparison against the
    master-index spreadsheet's names doesn't depend on which Unicode form
    either side happens to use.
    """
    by_basename = {}
    by_path = {}
    for root, _dirs, files in os.walk(base_dir):
        for f in files:
            full = os.path.join(root, f)
            rel = os.path.relpath(full, base_dir).replace(os.sep, "/")
            f_norm = nfc(f)
            rel_norm = nfc(rel)
            by_basename.setdefault(f_norm, []).append(rel_norm)
            by_path[rel_norm] = full
    return by_basename, by_path


def strip_code_prefix(text):
    """Removes a 'KB-XXX - ' (or similar) prefix from the start of the text, if present."""
    if not text:
        return text
    text = text.strip()
    # Accept the dash variants too. The standard requires a plain hyphen and a
    # header using an en/em dash is reported separately — but failing to strip
    # the prefix here would ALSO surface it as a bogus name divergence, i.e.
    # the same defect counted twice, with a recommendation to propagate the
    # broken form to the other sources.
    for separator in (" - ", " \u2013 ", " \u2014 "):
        parts = text.split(separator, 1)
        if len(parts) == 2 and parts[0].upper().startswith("KB-"):
            return parts[1].strip()
    return text


def classify_table_header(header_cells):
    """
    Decides whether the Version History table's first row is a recognizable
    column header, and in which format:
      - "new": columns Author | Version | Date | Description (of the Change)
      - "legacy": columns No. | Revision Date | Revision | Reviewer
      - None: not a header row (it's already data), or an unrecognized format
    Returns (format, version_index_or_None, description_index_or_None).

    Matches both the Portuguese terms (Autor/Versão/Revisor/Revisão) this
    plugin was built around and their English equivalents
    (Author/Version/Reviewer/Revision), so the same script works for a base
    in either language.
    """
    if not header_cells or len(header_cells) < 4:
        return None, None, None
    lowered = [c.strip().lower() for c in header_cells]
    # new format: has a column clearly called "version" AND an "author" one
    if any("versão" in c or "versao" in c or "version" in c for c in lowered) and any(
        "autor" in c or "author" in c for c in lowered
    ):
        idx_version = next(
            i for i, c in enumerate(lowered) if "versão" in c or "versao" in c or "version" in c
        )
        idx_description = next((i for i, c in enumerate(lowered) if "descri" in c), None)
        return "new", idx_version, idx_description
    # legacy format: has "reviewer"/"revisor" and "revision"/"revisão" as column names (not data)
    if any("revisor" in c or "reviewer" in c for c in lowered) and any(
        c.strip() in ("revisão", "revisao", "revision") for c in lowered
    ):
        idx_description = next(
            (i for i, c in enumerate(lowered) if c.strip() in ("revisão", "revisao", "revision")),
            None,
        )
        return "legacy", None, idx_description
    return None, None, None


def read_docx_signals(full_path):
    """Extracts the title, header, header version, and Version History table signals."""
    signals = {
        "title": None,
        "header": None,
        "header_name": None,
        "header_version": None,
        "header_missing_tab": False,
        "history_version": None,
        "history_description": None,
        "table_format": None,  # "new", "legacy", None (unrecognized/empty)
        "error": None,
    }
    try:
        d = docx.Document(full_path)
        # First NON-EMPTY paragraph: a blank line at the top of the body used
        # to make the title read as "", which silently dropped it out of the
        # four-source name comparison (turning it into a three-source one).
        title_para = next((p for p in d.paragraphs if p.text.strip()), None)
        if title_para is not None:
            signals["title"] = title_para.text.strip()

        hdr_paragraphs = [p.text for p in d.sections[0].header.paragraphs if p.text.strip()]
        if hdr_paragraphs:
            hdr_text = hdr_paragraphs[0].strip()
            signals["header"] = hdr_text
            if "\t" in hdr_text:
                signals["header_name"] = hdr_text.split("\t")[0].strip()
                tail = hdr_text.split("\t")[-1].strip()
                if tail.lower().startswith("v") and tail[1:2].isdigit():
                    signals["header_version"] = tail[1:].strip()
            else:
                # no tab: only treated as "glued together" if it looks like
                # name+version were concatenated (ends in something like
                # "v1.2" with no separator before it)
                # Require the version to be glued directly to a word character:
                # "Namev1.2" is a broken header, but "Migracao Office v2" is a
                # legitimate document name that merely ends in something
                # version-shaped. Without the lookbehind, the latter produced a
                # false "missing tab", a truncated name, and a phantom version.
                m = re.search(r"(?<=\w)v\d+(\.\d+)?$", hdr_text.strip(), re.IGNORECASE)
                if m:
                    signals["header_missing_tab"] = True
                    signals["header_version"] = m.group(0)[1:]
                    # The regex already knows exactly where the version starts,
                    # so the name is everything before it. Without this, the
                    # name would keep the glued-on version ("Namev1.2") and the
                    # SAME defect would be reported twice: once correctly as a
                    # missing-tab header, and again as a bogus name divergence
                    # whose recommendation would tell the person to propagate
                    # the broken name to the other three sources.
                    signals["header_name"] = hdr_text[: m.start()].strip()
                else:
                    # No tab and no trailing version: nothing to split off, the
                    # whole header is the name.
                    signals["header_name"] = hdr_text

        if d.tables:
            last_table = d.tables[-1]
            if len(last_table.rows) >= 1:
                header_cells = [c.text for c in last_table.rows[0].cells]
                fmt, idx_version, idx_description = classify_table_header(header_cells)
                signals["table_format"] = fmt
                if fmt == "new" and len(last_table.rows) >= 2:
                    last_row = [c.text.strip() for c in last_table.rows[-1].cells]
                    if idx_version is not None and idx_version < len(last_row):
                        signals["history_version"] = last_row[idx_version]
                    if idx_description is not None and idx_description < len(last_row):
                        signals["history_description"] = last_row[idx_description]
                elif fmt == "legacy" and len(last_table.rows) >= 2:
                    last_row = [c.text.strip() for c in last_table.rows[-1].cells]
                    if idx_description is not None and idx_description < len(last_row):
                        signals["history_description"] = last_row[idx_description]
                # format None: we don't try to extract anything — safer than
                # risking comparing the wrong column (e.g. mistaking Date for Version).
    except Exception as e:
        signals["error"] = str(e)
    return signals


def read_tab_category_map(wb):
    """
    Reads the Overview sheet's Category -> Tab mapping and returns it inverted
    (tab name -> real category name).

    This mapping is DATA, not something to recompute. The bootstrap may have
    had to alter a category name to make it a legal Excel sheet title (invalid
    characters, the 31-character cap, a collision), so deriving the tab name
    from the folder a second time here would mean keeping two implementations
    of the same transformation in sync forever — and when they drifted, the
    checker reported that a folder didn't match itself. Reading what the
    workbook already recorded removes that whole class of bug.

    Returns {} when there's no usable Overview, in which case the caller falls
    back to a direct comparison.
    """
    mapping = {}
    for sheet_name in wb.sheetnames:
        if sheet_name not in SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            if not row or not row[0]:
                continue
            category = str(row[0]).strip()
            reference = str(row[-1] or "")
            m = re.search(r"'([^']+)'", reference)
            if m:
                mapping[m.group(1).strip()] = category
    return mapping


def strip_numeric_prefix_for_compare(folder_name):
    """'01 - Identity' -> 'Identity'. Used to compare a folder against its tab."""
    m = re.match(r"^\d+\s*-\s*(.+)$", (folder_name or "").strip())
    return m.group(1).strip() if m else (folder_name or "").strip()


def _style_chain(style, max_depth=10):
    """
    Yields a paragraph style and its base styles, outermost first. Word
    formatting cascades: a run may say nothing about bold/size/color while the
    paragraph's style (or the style that style is based on) does. Walking this
    chain is what keeps a visually correct document from being reported as a
    violation just because the property wasn't set on the run itself.
    """
    seen = 0
    while style is not None and seen < max_depth:
        yield style
        style = getattr(style, "base_style", None)
        seen += 1


def resolve_font_property(run, paragraph, prop):
    """
    Resolves an effective font property ("bold", "size", "color") for a run,
    following the cascade run -> paragraph style -> base styles.

    Returns (value, determined). `determined` is False when nothing in the
    cascade states the property explicitly — in that case the real value comes
    from the Word theme/defaults, which python-docx can't see reliably. In
    CONSERVATIVE mode we never call that a violation; it goes to the
    "could not verify" bucket instead.
    """
    def read(source_font, from_run):
        if prop == "bold":
            return source_font.bold if not from_run else source_font.bold
        if prop == "size":
            return source_font.size
        if prop == "color":
            color = source_font.color
            if color is None:
                return None
            try:
                return color.rgb
            except (AttributeError, ValueError):
                return None
        return None

    value = read(run.font, True) if prop != "bold" else run.bold
    if value is not None:
        return value, True

    for style in _style_chain(paragraph.style):
        font = getattr(style, "font", None)
        if font is None:
            continue
        value = read(font, False)
        if value is not None:
            return value, True

    return None, False


def _rgb_str(value):
    return str(value).upper() if value is not None else None


# Section 3 required headings, in the order the standard mandates. Optional
# sections (Roles & Responsibilities, Naming Convention, Verification) may sit
# between them, and an extra custom heading is allowed — the standard doesn't
# forbid one. What is checked: every required section exists, they appear in
# the mandated relative order, and Version History is last.
REQUIRED_SECTIONS = (
    ("purpose", ("objetivo", "purpose")),
    ("prerequisites", ("pré-requisitos", "pre-requisitos", "prerequisites")),
    ("steps", ("passo a passo", "step by step")),
    ("history", ("histórico de versões", "historico de versoes", "version history")),
)


def _list_format(paragraph, doc):
    """
    Returns "bullet", "decimal", None (not a list), or "unknown" when the
    paragraph is a list item whose numbering definition can't be resolved.
    Word stores bullet and numbered lists under the same "List Paragraph"
    style, so the style name alone can't tell them apart — the numbering part
    has to be consulted.
    """
    # A paragraph can get its list formatting three ways, in this order of
    # specificity. Only checking direct numbering flags documents that use
    # Word's built-in List Bullet / List Number styles — where the numbering
    # lives in the style definition, not on the paragraph — as if they were not
    # lists at all.
    style_name = (paragraph.style.name or "").strip().lower()
    if "list bullet" in style_name:
        return "bullet"
    if "list number" in style_name:
        return "decimal"

    pPr = paragraph._p.pPr
    numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
    if numPr is None:
        style_element = getattr(paragraph.style, "element", None)
        style_pPr = style_element.find(qn("w:pPr")) if style_element is not None else None
        numPr = style_pPr.find(qn("w:numPr")) if style_pPr is not None else None
    if numPr is None:
        return None
    num_id_el = numPr.find(qn("w:numId"))
    if num_id_el is None:
        return "unknown"
    num_id = num_id_el.get(qn("w:val"))
    try:
        numbering = doc.part.numbering_part.element
    except (AttributeError, KeyError, ValueError):
        return "unknown"
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) != num_id:
            continue
        abstract = num.find(qn("w:abstractNumId"))
        if abstract is None:
            return "unknown"
        abstract_id = abstract.get(qn("w:val"))
        for a in numbering.findall(qn("w:abstractNum")):
            if a.get(qn("w:abstractNumId")) != abstract_id:
                continue
            lvl = a.find(qn("w:lvl"))
            if lvl is None:
                return "unknown"
            fmt = lvl.find(qn("w:numFmt"))
            if fmt is None:
                return "unknown"
            value = (fmt.get(qn("w:val")) or "").lower()
            return "bullet" if value == "bullet" else "decimal"
    return "unknown"


def _footer_signals(doc):
    """Returns (text, has_tab, has_page_field) for the first section's footer."""
    try:
        footer = doc.sections[0].footer
    except (IndexError, AttributeError):
        return None, False, False
    text = "".join(p.text for p in footer.paragraphs).strip()
    xml = footer._element.xml if hasattr(footer, "_element") else ""
    has_page_field = "PAGE" in xml
    has_tab = any("\t" in p.text for p in footer.paragraphs)
    return text, has_tab, has_page_field


def check_document_structure(full_path, label):
    """
    Checks Section 3: required sections present, in order, with Version History
    last. Returns a list of issues.
    """
    try:
        d = docx.Document(full_path)
    except Exception as e:
        return ["%s: couldn't open the file for structure checks (%s)" % (label, e)]

    issues = []
    headings = []
    for p in d.paragraphs:
        style_name = (p.style.name or "").strip().lower()
        if any(marker in style_name for marker in HEADING_STYLE_MARKERS):
            headings.append(p.text.strip())

    if not headings:
        return ["%s: no Heading 1 sections found — the document has no recognizable structure" % label]

    lowered = [h.lower() for h in headings]
    positions = {}
    for key, synonyms in REQUIRED_SECTIONS:
        idx = next((i for i, h in enumerate(lowered)
                    if any(h.startswith(syn) for syn in synonyms)), None)
        if idx is None:
            issues.append(
                "%s: required section '%s' is missing (expected one of: %s)"
                % (label, key, ", ".join(synonyms))
            )
        else:
            positions[key] = idx

    ordered = [k for k, _ in REQUIRED_SECTIONS if k in positions]
    for earlier, later in zip(ordered, ordered[1:]):
        if positions[earlier] > positions[later]:
            issues.append(
                "%s: section '%s' appears after '%s'; the standard's order is "
                "Purpose -> Prerequisites -> Step by Step -> Version History"
                % (label, earlier, later)
            )

    if "history" in positions and positions["history"] != len(headings) - 1:
        trailing = headings[positions["history"] + 1:]
        issues.append(
            "%s: Version History must be the last section, but %s come after it"
            % (label, ", ".join("'%s'" % t for t in trailing))
        )

    return issues


def check_document_formatting(full_path, label):
    """
    Checks a .docx against the formatting constants above (Section 4 of the
    spec), in CONSERVATIVE mode: something is only reported as a violation when
    an explicit, divergent value was actually found. Anything the cascade can't
    settle is returned separately as "could not verify" and never counted as a
    problem — a checker that guesses produces false positives, and a formatting
    report nobody trusts is worse than no report at all.

    Returns (violations, unverifiable), both lists of strings.
    """
    violations = []
    unverifiable = []

    try:
        d = docx.Document(full_path)
    except Exception as e:
        return [f"{label}: couldn't open the file for formatting checks ({e})"], []

    # --- margins (always determinable: they're stored on the section) ---
    for section in d.sections:
        for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
            margin = getattr(section, side, None)
            if margin is None:
                unverifiable.append(f"{label}: '{side}' is not set explicitly")
                continue
            if abs(margin.inches - EXPECTED_MARGIN_INCHES) > 0.01:
                violations.append(
                    f"{label}: {side.replace('_', ' ')} is {margin.inches:.2f}in, "
                    f"expected {EXPECTED_MARGIN_INCHES}in"
                )
        break  # the standard is a single-section document; check the first

    # --- title: first non-empty paragraph of the body ---
    title_para = next((p for p in d.paragraphs if p.text.strip()), None)
    if title_para is None:
        unverifiable.append(f"{label}: no body paragraph found, title not checked")
    else:
        # Every run is checked, not just the first: a title split across runs
        # (autocorrect, pasted text) can have a correctly formatted first half
        # and an unformatted second half.
        for idx, run in enumerate(title_para.runs, 1):
            if not run.text.strip():
                continue
            where = f"{label}: title run {idx} ('{run.text.strip()[:30]}')"

            bold, determined = resolve_font_property(run, title_para, "bold")
            if not determined:
                unverifiable.append(f"{where}: bold not stated on the run or its style chain")
            elif bold is not True:
                violations.append(f"{where}: is not bold")

            size, determined = resolve_font_property(run, title_para, "size")
            if not determined:
                unverifiable.append(f"{where}: font size not stated on the run or its style chain")
            elif abs(size.pt - EXPECTED_TITLE_SIZE_PT) > 0.01:
                violations.append(
                    f"{where}: font size is {size.pt}pt, expected {EXPECTED_TITLE_SIZE_PT}pt"
                )

            color, determined = resolve_font_property(run, title_para, "color")
            if not determined or color is None:
                unverifiable.append(f"{where}: font color not stated on the run or its style chain")
            elif _rgb_str(color) != EXPECTED_TITLE_COLOR:
                violations.append(
                    f"{where}: font color is {_rgb_str(color)}, expected {EXPECTED_TITLE_COLOR}"
                )

    # --- section headings must use the NATIVE Heading 1 style ---
    for p in d.paragraphs:
        style_name = (p.style.name or "").strip().lower()
        if any(marker in style_name for marker in HEADING_STYLE_MARKERS):
            builtin = getattr(p.style, "builtin", None)
            if builtin is None:
                unverifiable.append(
                    f"{label}: heading '{p.text.strip()[:30]}' — couldn't tell whether the "
                    "style is native"
                )
            elif not builtin:
                violations.append(
                    f"{label}: heading '{p.text.strip()[:30]}' uses a custom style "
                    f"('{p.style.name}'), not Word's native Heading 1"
                )

    # --- warning callouts ---
    for p in d.paragraphs:
        if not p.text.strip().startswith(CALLOUT_PREFIX):
            continue
        snippet = p.text.strip()[:30]
        pPr = p._p.pPr
        shd = pPr.find(qn("w:shd")) if pPr is not None else None
        if shd is None:
            violations.append(
                f"{label}: callout '{snippet}' has no shading; expected "
                f"{EXPECTED_CALLOUT_SHADING}"
            )
        else:
            fill = (shd.get(qn("w:fill")) or "").upper()
            if fill != EXPECTED_CALLOUT_SHADING:
                violations.append(
                    f"{label}: callout '{snippet}' shading is {fill or 'empty'}, "
                    f"expected {EXPECTED_CALLOUT_SHADING}"
                )
        for idx, run in enumerate(p.runs, 1):
            if not run.text.strip():
                continue
            color, determined = resolve_font_property(run, p, "color")
            if not determined or color is None:
                unverifiable.append(
                    f"{label}: callout '{snippet}' run {idx} — text color not stated"
                )
            elif _rgb_str(color) != EXPECTED_CALLOUT_TEXT_COLOR:
                violations.append(
                    f"{label}: callout '{snippet}' run {idx} text color is "
                    f"{_rgb_str(color)}, expected {EXPECTED_CALLOUT_TEXT_COLOR}"
                )


    # --- page orientation ---
    try:
        from docx.enum.section import WD_ORIENT
        orientation = d.sections[0].orientation
        if orientation is not None and orientation != WD_ORIENT.PORTRAIT:
            violations.append("%s: page orientation is landscape, expected portrait" % label)
    except (ImportError, IndexError, AttributeError):
        unverifiable.append("%s: page orientation could not be read" % label)

    # --- footer: must exist, use a tab, and use Word's automatic page fields ---
    footer_text, footer_has_tab, footer_has_page = _footer_signals(d)
    if not footer_text:
        violations.append(
            "%s: the document has no footer; the standard requires "
            "'<Organization label>  |  Knowledge Base<TAB>Page <n> of <total>'" % label
        )
    else:
        if not footer_has_tab:
            violations.append(
                "%s: the footer has no tab separating the label from the page number" % label
            )
        if not footer_has_page:
            violations.append(
                "%s: the footer's page number is typed in rather than a Word automatic "
                "page field, so it won't update" % label
            )

    # --- header: plain hyphen between code and name, never an en/em dash ---
    try:
        header_text = "".join(p.text for p in d.sections[0].header.paragraphs)
    except (IndexError, AttributeError):
        header_text = ""
    for dash in DASHES:
        if dash in header_text:
            violations.append(
                "%s: the header uses '%s' between code and name; the standard "
                "requires a plain hyphen '-'" % (label, dash)
            )
            break

    # --- list formatting per section, only where the section has content ---
    headings = []
    for i, p in enumerate(d.paragraphs):
        style_name = (p.style.name or "").strip().lower()
        if any(marker in style_name for marker in HEADING_STYLE_MARKERS):
            headings.append((i, p.text.strip().lower()))

    for pos, (idx, title) in enumerate(headings):
        expected = None
        if any(title.startswith(x) for x in BULLET_SECTIONS):
            expected = "bullet"
        elif any(title.startswith(x) for x in NUMBERED_SECTIONS):
            expected = "decimal"
        if expected is None:
            continue
        end = headings[pos + 1][0] if pos + 1 < len(headings) else len(d.paragraphs)
        body_paras = [p for p in d.paragraphs[idx + 1:end] if p.text.strip()]
        if not body_paras:
            continue  # an empty section can't be judged
        formats = {_list_format(p, d) for p in body_paras}
        if "unknown" in formats and expected not in formats:
            unverifiable.append(
                "%s: section '%s' — list numbering definition could not be resolved"
                % (label, title)
            )
        elif expected not in formats:
            kind = "a bulleted list" if expected == "bullet" else "a numbered list"
            violations.append("%s: section '%s' is not formatted as %s" % (label, title, kind))

    # --- author names in the Version History must be Title Case, not all caps ---
    if d.tables:
        last_table = d.tables[-1]
        if len(last_table.rows) >= 2:
            header_cells = [c.text for c in last_table.rows[0].cells]
            fmt, _idx_v, _idx_d = classify_table_header(header_cells)
            if fmt is not None:
                lowered = [c.strip().lower() for c in header_cells]
                author_idx = next(
                    (i for i, c in enumerate(lowered)
                     if "autor" in c or "author" in c or "revisor" in c or "reviewer" in c),
                    None,
                )
                if author_idx is not None:
                    for row in last_table.rows[1:]:
                        if author_idx >= len(row.cells):
                            continue
                        value = row.cells[author_idx].text.strip()
                        letters = [ch for ch in value if ch.isalpha()]
                        if len(letters) > 1 and value == value.upper():
                            violations.append(
                                "%s: Version History author '%s' is in all caps; the "
                                "standard requires Title Case" % (label, value)
                            )

    return violations, unverifiable


def resolve_missing_master_list(base_dir):
    """
    Called when no master-index spreadsheet was found. Interactively offers to
    create one (delegating to bootstrap_master_list.py) or to be pointed at an
    existing file.

    IMPORTANT: this only prompts when stdin is an interactive terminal. Under
    CI, cron, or a pipe there is nobody to answer, and a script that blocks
    waiting for input would hang the pipeline — in that case it prints the same
    guidance and the caller exits with code 2. This is also why creation lives
    in a separate script: check_master_list.py itself never writes.

    Returns (path, created) — created is True only when a spreadsheet was
    actually generated in this run. Pointing at an existing file must not
    claim anything was written. (None, False) means abort.
    """
    print(f"No master-index spreadsheet found in '{base_dir}'.")
    print(f"(looked for: {', '.join(MASTER_LIST_NAME_PATTERNS)})\n")

    if not sys.stdin.isatty():
        print(
            "Running non-interactively, so I won't prompt. Either:\n\n"
            "  1) generate one from the files already in the base:\n"
            f"       python3 bootstrap_master_list.py \"{base_dir}\"\n\n"
            "  2) or point at an existing spreadsheet:\n"
            f"       python3 check_master_list.py \"{base_dir}\" \"<file name>.xlsx\"\n"
        )
        return None, False

    print("What would you like to do?")
    print("  [1] Create one now by scanning this base")
    print("  [2] Point me at an existing spreadsheet")
    print("  [3] Abort")

    try:
        choice = input("> ").strip()
    except (EOFError, KeyboardInterrupt):
        print("\nAborted.")
        return None, False

    if choice == "1":
        try:
            from bootstrap_master_list import bootstrap
        except ImportError as e:
            print(f"Couldn't load bootstrap_master_list.py ({e}).")
            return None, False
        print(
            "\nHeads-up: every row will be written with Status 'Em revisão', and any\n"
            "document whose file name has no KB code will be left with '—' instead of\n"
            "a new code — codes are never recycled, so that choice stays with you.\n"
        )
        # Ask which language the base is authored in. Defaulting silently to
        # English would write an index whose column headers and status values
        # don't match the documents it indexes.
        print("What language are this base's documents written in?")
        print("  [1] English (default)")
        print("  [2] Another language — give its two-letter code")
        try:
            lang_choice = input("> ").strip() or "1"
            lang = "en"
            if lang_choice == "2":
                lang = input("Language code: ").strip().lower()
            confirm = input("Proceed? [y/N] ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\nAborted.")
            return None, False
        if confirm not in ("y", "yes", "s", "sim"):
            print("Aborted. Nothing was written.")
            return None, False
        try:
            return bootstrap(base_dir, lang=lang), True
        except Exception as e:
            print(f"ERROR: {e}")
            return None, False

    if choice == "2":
        try:
            name = input("File name (inside the base folder): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nAborted.")
            return None, False
        candidate = os.path.join(base_dir, name)
        if not os.path.isfile(candidate):
            print(f"ERROR: '{name}' not found in '{base_dir}'.")
            return None, False
        return candidate, False

    print("Aborted. Nothing was written.")
    return None, False


def find_non_ascii_filenames(by_path):
    """
    Lists (informational, NEVER an error) the file paths that contain a
    non-ASCII character (accents, etc.).

    This is different from the NFC/NFD bug handled in nfc(): here the
    Unicode form doesn't matter, only whether the name has any character
    outside the basic ASCII range. It's a PORTABILITY note (the name could
    break if the base is ever exported to a system that doesn't handle
    Unicode well), not a base inconsistency — so it never enters the
    problem count, it's just a heads-up for a possible future migration.
    """
    found = []
    for path in sorted(by_path.keys()):
        if not path.isascii():
            found.append(path)
    return found


def format_date(dt):
    if dt is None:
        return "unknown"
    if isinstance(dt, datetime.datetime):
        return dt.strftime("%Y-%m-%d")
    return str(dt)


def check(base_dir, master_list_name=None):
    bootstrapped = False
    try:
        master_path = find_master_list(base_dir, master_list_name)
    except FileNotFoundError:
        # No spreadsheet: offer to create one or be pointed at it, instead of
        # dying with a traceback. Returns None to signal "couldn't run".
        master_path, bootstrapped = resolve_missing_master_list(base_dir)
        if master_path is None:
            return None
    print(f"Base: {base_dir}")
    print(f"Master-index spreadsheet: {os.path.basename(master_path)}\n")

    by_basename, by_path = index_real_files(base_dir)

    wb = openpyxl.load_workbook(master_path, data_only=True)
    tab_to_category = read_tab_category_map(wb)

    broken_reference_issues = []
    duplicate_code_issues = []
    name_divergence_issues = []
    version_divergence_issues = []
    legacy_table_issues = []
    header_missing_tab_issues = []
    formatting_issues = []
    formatting_unverifiable = []
    structure_issues = []
    folder_mismatch_issues = []
    ambiguous_reference_issues = []
    footer_labels = {}
    codes_seen = {}
    referenced_files = set()
    sheets_without_header = []

    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Find the column-header row by name instead of assuming it's always
        # row 2 — a sheet that doesn't have a recognizable "Code"/"Código"
        # column is skipped (with a warning) rather than silently reading
        # the wrong columns as data.
        header_idx, header_cells = find_header_row(rows)
        if header_idx is None:
            sheets_without_header.append(sheet_name)
            continue
        col = map_columns(header_cells)
        if "code" not in col or "document" not in col or "file" not in col:
            sheets_without_header.append(sheet_name)
            continue

        def cell(row, key):
            idx = col.get(key)
            return row[idx] if idx is not None and idx < len(row) else None

        for row in rows[header_idx + 1:]:
            if not row or not cell(row, "code"):
                continue
            codigo = str(cell(row, "code")).strip()
            documento = cell(row, "document")
            arquivo = cell(row, "file")
            status = cell(row, "status")
            versao_planilha = cell(row, "version")
            ultima_atualizacao = cell(row, "last_updated")

            # 1) duplicate code
            if codigo != NO_CODE_MARKER:
                if codigo in codes_seen:
                    duplicate_code_issues.append(
                        f"Code '{codigo}' appears in '{codes_seen[codigo]}' and in '{sheet_name}' "
                        f"(current document: '{documento}')"
                    )
                else:
                    codes_seen[codigo] = sheet_name

            if not arquivo:
                continue
            # normalize to NFC before any file-name comparison, so "same name,
            # different Unicode form" isn't mistaken for a real broken
            # reference (see nfc()).
            arquivo = nfc(str(arquivo))

            # 2) broken reference
            full_path = None
            resolved_rel = None
            if "/" in arquivo:
                full_path = by_path.get(arquivo)
                resolved_rel = arquivo if full_path else None
            else:
                matches = by_basename.get(arquivo) or []
                if len(matches) == 1:
                    resolved_rel = matches[0]
                    full_path = by_path.get(resolved_rel)
                elif len(matches) > 1:
                    # The same file name exists in more than one folder. Taking
                    # the first match would silently compare the WRONG document
                    # — its title, header and version would come from a file
                    # this row doesn't refer to. The row's tab tells us which
                    # category it belongs to, so use that to narrow it down.
                    expected_category = tab_to_category.get(sheet_name, sheet_name)
                    expected_label = strip_numeric_prefix_for_compare(expected_category).lower()
                    narrowed = [
                        m for m in matches
                        if strip_numeric_prefix_for_compare(m.split("/")[0]).lower()
                        == expected_label
                    ]
                    if len(narrowed) == 1:
                        resolved_rel = narrowed[0]
                        full_path = by_path.get(resolved_rel)
                    else:
                        ambiguous_reference_issues.append(
                            "[%s] %s ('%s'): the File column says '%s', but that name exists "
                            "in %d folders (%s) and the row's tab doesn't single one out. "
                            "Nothing about this document was checked. Use the full relative "
                            "path in the File column, or rename one of the files."
                            % (sheet_name, codigo, documento, arquivo, len(matches),
                               ", ".join("'%s'" % m for m in matches))
                        )
                        for m in matches:
                            referenced_files.add(m)
                        continue

            if resolved_rel:
                # Record the resolved path, not the bare name: otherwise a
                # same-named file in another folder that nothing indexes would
                # escape orphan detection by borrowing this row's reference.
                referenced_files.add(resolved_rel)

            if not full_path:
                broken_reference_issues.append(
                    f"[{sheet_name}] {codigo} ('{documento}'): master-index spreadsheet points to "
                    f"'{arquivo}', which doesn't exist in the base."
                )
                continue

            if not full_path.lower().endswith(".docx"):
                continue  # only .docx has a title/header/history to compare

            label = f"[{sheet_name}] {codigo}"
            fmt_violations, fmt_unverifiable = check_document_formatting(full_path, label)
            formatting_issues.extend(fmt_violations)
            formatting_unverifiable.extend(fmt_unverifiable)
            # Legacy documents are, by definition, the ones that predate the
            # standard — flagging every one of them for missing sections would
            # flood the report with findings nobody intends to fix, which is
            # exactly the failure the conservative mode exists to avoid.
            if str(status or "").strip().lower() not in LEGACY_STATUS_LABELS:
                structure_issues.extend(check_document_structure(full_path, label))

            # Collect the footer's organization label. It can't be validated in
            # isolation (the script doesn't know the base's label), but every
            # document in one base must carry the SAME one.
            try:
                footer_doc = docx.Document(full_path)
                footer_line = "".join(
                    p.text for p in footer_doc.sections[0].footer.paragraphs
                )
                prefix = footer_line.split("\t")[0].strip()
                if prefix:
                    footer_labels.setdefault(prefix, []).append(codigo)
            except Exception:
                pass

            # Folder/tab mismatch. The File column stores only the bare name
            # precisely because the tab is supposed to imply the folder; if the
            # file actually sits somewhere else, that assumption is broken and
            # the row is wrong, even though nothing else complains.
            rel = os.path.relpath(full_path, base_dir).replace(os.sep, "/")
            folder = rel.split("/")[0] if "/" in rel else ""
            folder_label = strip_numeric_prefix_for_compare(folder)
            # The tab may be a sanitized/truncated version of the real category
            # name; the Overview records which category each tab represents.
            expected = tab_to_category.get(sheet_name, sheet_name)
            expected_label = strip_numeric_prefix_for_compare(expected)
            if folder_label and folder_label.lower() != expected_label.lower():
                folder_mismatch_issues.append(
                    f"{label} ('{documento}'): the row is filed under category "
                    f"'{expected_label}', but the file is in folder '{folder}'. "
                    "Either move the file or move the row."
                )

            signals = read_docx_signals(full_path)
            if signals["error"]:
                broken_reference_issues.append(
                    f"[{sheet_name}] {codigo}: couldn't open '{arquivo}' ({signals['error']})"
                )
                continue

            # 3a) Version History table in the legacy format (no real semantic
            # version column) — reported as its own item, instead of producing
            # a false "version divergence" comparing something the table doesn't have.
            if signals["table_format"] == "legacy":
                legacy_table_issues.append(
                    f"[{sheet_name}] {codigo} ('{documento}'): the Version History table is "
                    "in the old format (No. | Revision Date | Revision | Reviewer), which has no "
                    "semantic-version column — just a revision counter. "
                    "Action needed: convert the table to the new format "
                    "(Author | Version | Date | Change Description)."
                )

            # 3b) header missing the tab separator between name and version
            # (e.g. 'Namev1.2' glued together instead of 'Name<TAB>v1.2').
            # Reported as its own item so it doesn't show up as a confusing
            # 'header: None' inside the generic version divergence.
            if signals["header_missing_tab"]:
                header_missing_tab_issues.append(
                    f"[{sheet_name}] {codigo} ('{documento}'): the header has no tab separator "
                    f"between name and version (found: '{signals['header']}'). "
                    "Action needed: fix the header to the format "
                    "'KB-XXX - <Document Name><TAB>v<Version>', with a real tab character "
                    "separating the name from the version."
                )

            file_name = strip_code_prefix(os.path.splitext(os.path.basename(arquivo))[0])
            title_name = strip_code_prefix(signals["title"])
            header_name = strip_code_prefix(signals["header_name"])
            spreadsheet_name = str(documento).strip() if documento else None

            candidates = {
                "file name": file_name,
                "title (document body)": title_name,
                "header": header_name,
                "master-index spreadsheet (Document column)": spreadsheet_name,
            }
            unique_names = set(v for v in candidates.values() if v)
            if len(unique_names) > 1:
                mtime = os.path.getmtime(full_path)
                mtime_dt = datetime.datetime.fromtimestamp(mtime)
                spreadsheet_date = parse_date_value(ultima_atualizacao)

                if ultima_atualizacao is not None and spreadsheet_date is None:
                    recommendation = (
                        "no automatic recommendation — the master-index spreadsheet's "
                        f"Last Updated value ('{ultima_atualizacao}') couldn't be read as a date"
                    )
                elif spreadsheet_date is None:
                    recommendation = "no automatic recommendation — manual decision needed"
                elif mtime_dt.date() > spreadsheet_date:
                    recommendation = (
                        f"use the DOCUMENT's name (title/header), since the file was "
                        f"last modified on {format_date(mtime_dt)}, after the last "
                        f"update recorded in the master-index spreadsheet ({format_date(spreadsheet_date)})"
                    )
                elif spreadsheet_date > mtime_dt.date():
                    recommendation = (
                        f"use the MASTER-INDEX SPREADSHEET's name, since it was updated on "
                        f"{format_date(spreadsheet_date)}, after the file's last "
                        f"modification ({format_date(mtime_dt)})"
                    )
                else:
                    recommendation = "no automatic recommendation — manual decision needed"

                detail = [f"[{sheet_name}] {codigo}: names diverge between sources:"]
                for origin, value in candidates.items():
                    detail.append(f"    - {origin}: '{value}'")
                detail.append(f"    Recommendation: {recommendation}")
                detail.append(
                    "    Action needed: pick which of these variants becomes the official name and "
                    "apply it in all 4 sources (file, title, header, master-index spreadsheet)."
                )
                detail.append(
                    "    Note: applying that choice to all 4 sources is a content edit, not just a "
                    "name change — it triggers the critical version rule (the header must always "
                    "match the last row of the Version History). In other words, resolving this name "
                    "divergence also requires bumping the document's version and adding a new row to "
                    "the Version History describing the name standardization — otherwise fixing one "
                    "inconsistency creates another."
                )
                name_divergence_issues.append("\n".join(detail))

            # 3) version divergence
            versions = {
                "header": signals["header_version"],
                "last row of Version History": signals["history_version"],
                "master-index spreadsheet (Version column)": str(versao_planilha).strip() if versao_planilha else None,
            }
            unique_versions = set(v for v in versions.values() if v)
            if len(unique_versions) > 1:
                detail = [f"[{sheet_name}] {codigo}: versions diverge between sources:"]
                for origin, value in versions.items():
                    detail.append(f"    - {origin}: '{value}'")
                detail.append(
                    "    Action needed: align all 3 sources to the most recent version "
                    "(usually the Version History's, since it's the audit record)."
                )
                version_divergence_issues.append("\n".join(detail))

    # 4) orphaned files (exist in the base, but aren't in any master-index row)
    orphan_file_issues = []
    for basename, paths in by_basename.items():
        # The spreadsheet in use is never an orphan — including when it was
        # passed by an explicit name that doesn't match the usual patterns.
        if (is_master_list_filename(basename)
                or basename == nfc(os.path.basename(master_path))
                or basename == CONFIG_FILENAME):
            continue
        for path in paths:
            if is_inside_attachment_folder(path):
                continue  # attachment of a KB document, not a standalone indexable item
            if path not in referenced_files and basename not in referenced_files:
                orphan_file_issues.append(f"'{path}' exists in the base, but isn't in any master-index row")

    if sheets_without_header:
        print(
            "WARNING: the following sheet(s) have no recognizable column-header row "
            "(a 'Code'/'Código' column) and were skipped entirely — nothing on them was "
            f"checked: {', '.join(sheets_without_header)}\n"
        )

    def section(title, items):
        print(f"=== {title} ({len(items)}) ===")
        if not items:
            print("No problems found.\n")
            return
        for item in items:
            print("- " + item)
        print()

    section("Broken references (master-index spreadsheet → file)", broken_reference_issues)
    section("Orphaned files (file → master-index spreadsheet)", orphan_file_issues)
    section("Duplicate/shared codes", duplicate_code_issues)
    section("Name divergence (file/title/header/master-index spreadsheet)", name_divergence_issues)
    section("Version divergence (header/history/master-index spreadsheet)", version_divergence_issues)
    section("Version History table in the legacy format", legacy_table_issues)
    section("Header missing the tab separator (name and version glued together)", header_missing_tab_issues)
    if len(footer_labels) > 1:
        detail = ["the base uses more than one footer label; every document must carry the same one:"]
        for prefix, codes in sorted(footer_labels.items()):
            shown = ", ".join(codes[:4]) + (", ..." if len(codes) > 4 else "")
            detail.append("    - '%s'  (%s)" % (prefix, shown))
        formatting_issues.append("\n".join(detail))

    section("Structure violations (Section 3 of the standard)", structure_issues)
    section("Ambiguous file reference (same name in several folders)", ambiguous_reference_issues)
    section("Folder/tab mismatch", folder_mismatch_issues)
    section("Formatting violations (Section 4 of the standard)", formatting_issues)

    # Informational — NOT a problem, and not part of the SUMMARY. Just a map
    # of files with a non-ASCII character in the name, in case of a future
    # migration to a system that doesn't handle Unicode well.
    # Informational — NOT counted as problems. Conservative mode: when the
    # run/style cascade doesn't state a formatting property explicitly, the
    # real value comes from Word's theme defaults, which can't be read
    # reliably here. Reporting those as violations would flood the report
    # with false positives, so they're listed apart instead.
    print(f"--- Informational: formatting properties that could not be verified "
          f"({len(formatting_unverifiable)}) ---")
    print(
        "These are NOT violations. The property simply isn't stated on the run or anywhere in\n"
        "its style chain, so its effective value comes from Word's defaults and can't be\n"
        "confirmed programmatically. Check visually if a specific document looks off."
    )
    for item in formatting_unverifiable:
        print("- " + item)
    print()

    non_ascii = find_non_ascii_filenames(by_path)
    print(f"--- Informational: files with a non-ASCII character in the name ({len(non_ascii)}) ---")
    print(
        "This is NOT a base consistency problem — it's only a portability note.\n"
        "Accented characters in file names work normally here; they'd only break if the base were\n"
        "ever exported to a system that doesn't handle Unicode well. Nothing to fix now."
    )
    if non_ascii:
        for path in non_ascii:
            print("- " + path)
    print()

    total = (
        len(broken_reference_issues)
        + len(orphan_file_issues)
        + len(duplicate_code_issues)
        + len(name_divergence_issues)
        + len(version_divergence_issues)
        + len(legacy_table_issues)
        + len(header_missing_tab_issues)
        + len(formatting_issues)
        + len(structure_issues)
        + len(folder_mismatch_issues)
        + len(ambiguous_reference_issues)
    )
    if bootstrapped:
        # The checker itself still wrote nothing; the spreadsheet was created by
        # bootstrap_master_list.py, which this run delegated to after asking.
        # Saying "no file was changed" flat out would be false here.
        print(
            f"SUMMARY: {total} problem(s) found. A master-index spreadsheet was created "
            "for this base (see above); the checks themselves changed no file."
        )
    else:
        print(f"SUMMARY: {total} problem(s) found. No file was changed by this script.")
    return total


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # Usage error is 'couldn't run' (2), not 'problems found' (1).
        print(__doc__)
        sys.exit(2)
    base_dir_arg = sys.argv[1]
    master_list_arg = sys.argv[2] if len(sys.argv) > 2 else None
    try:
        total_problems = check(base_dir_arg, master_list_arg)
    except (FileNotFoundError, RuntimeError) as e:
        print(f"ERROR: {e}")
        sys.exit(2)

    # Exit codes are distinct on purpose so a wrapper can tell the three cases
    # apart: 0 = clean, 1 = problems found, 2 = couldn't run at all.
    if total_problems is None:
        sys.exit(2)
    sys.exit(1 if total_problems else 0)
