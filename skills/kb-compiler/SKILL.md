---
name: kb-compiler
description: Use when creating, updating, or reviewing a knowledge-base (KB) article or runbook — including turning a raw draft or a document in another format (PDF, PPTX, Outlook .msg, Markdown) into one — to follow a fixed section structure, exact Word formatting, KB-XXX numbering, and a master-index spreadsheet convention. Ask for the base directory, the organization/team label, and the target document language if not already established in this conversation.
---

# Knowledge Base (KB) Documentation Standard

This skill defines a documentation standard for KB articles/runbooks: section structure, exact Word formatting, the `KB-XXX` numbering convention, and a master-index spreadsheet. It **does not assume any specific company, path, or folder** — discover or confirm that with the person before creating the first document, and never invent or reuse an example from another context. The final deliverable is always a `.docx` file (use python-docx or the docx skill to build/edit it), never plain loose text — unless the person explicitly asks for something else.

**The language of these instructions and the language of the documents this skill produces are independent** — see Sections 0 and 6.

## 0. Initial setup (once per base)

Before creating the first document in a new base, confirm with the person (or infer from context/folder already in use in the conversation — never ask again about something already clear):

- **Root directory** of the knowledge base (where the category folders will live).
- **Organization/team label**, used in the footer of every document (e.g., `"<Team Name> | Knowledge Base"`). Save this label and reuse it in every document of the same base — never invent a company/team name nor reuse an example from another base.
- **Default responsible person's name**, used as the author in Version History rows when no other name is given.
- **Target document language** — the language the actual KB documents (body text, section headings, the master-index spreadsheet's column headers, everything) should be written in. This is independent of the language of this SKILL.md: these instructions stay in English regardless, but the documents follow whatever language the base's author confirms here. Default to whichever language the existing base already uses; if starting a brand-new base, ask. Once confirmed for a base, reuse it for every document in that base without asking again — never mix languages within one base, and never guess a language from the person's own message language alone (they may be writing to Claude in one language while authoring documents in another).

The literal terms in this document (section headings, spreadsheet column headers, status values, folder and file names) are the standard's canonical English form. When the target document language confirmed in Section 0 is not English, translate every one of them into that language, keeping the same structure, order, and meaning — never mix languages within one base, and never write a heading as a slash pair.

If the base already exists, read the master-index spreadsheet and the existing folder structure instead of asking everything again.

## 1. Folder structure and categories

Inside the root directory, documents are organized into numbered category subfolders (e.g., `01 - <Category A>`, `02 - <Category B>`, ...), plus a `Legacy` folder for discontinued/superseded documents — translate that folder's own name too when the target language isn't English, but keep every base internally consistent. Use whatever categories already exist in the configured base; if building a base from scratch, ask the person which categories make sense for their team — never assume category names from another context.

## 2. Numbering and file naming

- **Global** sequential code `KB-XXX` (3 digits), unique across the ENTIRE base, numbered in a single sequence across all categories (it does not restart per folder). Before creating a new document, check the highest `KB-XXX` already used in the master-index spreadsheet and use the next number. Never reuse a code already used by another document, even if the old document was moved to Legacy — codes are never recycled.
- Documents in the Legacy folder **do not need** a KB code — the folder itself already signals the status. If a legacy document has no code, use `—` in the "Code" column of the master-index spreadsheet (never leave it blank, and never reuse/invent a code already used elsewhere).
- File name: `KB-XXX - <Document Name>.docx`. Legacy documents that do have a code follow the same pattern WITHOUT an extra suffix like "- Legacy" in the name (the folder already conveys that); documents without a code just use `<Document Name>.docx`.
- Prefixes like "Guide -" or "Manual -" can appear when the document is a longer tutorial, but they are not required.

### Mandatory name consistency (critical rule)

The **file name**, the Word **header**, the **title** (first line of the body), and the **"Document"** column in the master-index spreadsheet must all use exactly the same process name. Before delivering any document (new or revised), check all four and fix any divergence — this is a recurring failure mode in this kind of base, always check it.

## 3. Content structure (section order)

1. **Title** (first line of the body, "Normal" style): `KB-XXX - <Document Name>` (or just `<Name>` if there is no code).
2. **Purpose** (native Word Heading 1) — 1 to 3 sentences on what the document solves/teaches.
3. **[Roles & Responsibilities]** (Heading 1) — when different user profiles have different paths.
4. **[Naming Convention]** (Heading 1) — when there's a naming/code convention to explain before the step-by-step.
5. **Prerequisites** (Heading 1 — the heading itself is always a heading, never a list item) — its items are a **list, bulleted or numbered**; pick whichever the rest of the document already uses. This section is required even when the procedure has none: in that case write a single item reading `N/A`. An explicit `N/A` says "the author considered this and there are none"; an absent section says nothing at all, and the two are indistinguishable to whoever reads the document later.
6. **Step by Step** (Heading 1, can repeat as "Step by Step – <Context>" for alternate paths) — **numbered** (decimal) list, one imperative action per item.
7. **[Verification]** (Heading 1) — a list of success signals, bulleted or numbered, matching whatever the document already uses.
8. **Version History** (Heading 1, always last) — table `Author | Version | Date | Change Description`. The first row is always the creation, version `1.0`. Older documents may have a table in the legacy format (`No. | Revision Date | Revision | Reviewer`) — when editing one of these, convert it to the new standard while preserving all history. **Be careful rewriting this table**: confirm whether the first visible row is really a column header or already data — some old tables have no separate header row, and overwriting the wrong row erases history by mistake. **A new row must inherit the formatting of the rows already there** — font, size and colour are usually set on the runs, not on the table style, so a row created programmatically states nothing and falls back to the document's defaults. It renders larger and in a different colour than every row above it, which is immediately visible to a reader and is not something the audit checks. Copy the run formatting from an existing data row. Cell borders and shading live on the cell, not the run, so they carry over on their own.

Warnings/callouts inside the step-by-step use the callout style from Section 4 (formatting).

## 4. Exact Word formatting

- **Margins**: 0.75 inch on all sides, portrait.
- **Header**: ALWAYS in the format `KB-XXX - <Document Name>\tv<Current Version>` (with the `KB-XXX -` prefix included; documents without a code just use `<Name>\tv<Version>`). Use a **real tab character** (`\t`) between name and version — NEVER concatenate them with no separator (common bug: the text comes out glued together, e.g. "Namev2.4"). Use a regular hyphen `-`, never an en/em dash `–`/`—`, between the code and the name.
- **Footer**: `<Organization/Team label confirmed in Section 0>  |  Knowledge Base\tPage <n> of <total>` (Word's automatic page fields).
- **Title** (1st line of the body): "Normal" style, **bold**, **22pt**, color **RGB 1F3864** (dark blue).
- **Sections**: native Word "Heading 1" style — never create a custom style copied from another document.
- **Prerequisites/Verification lists**: a list — bulleted or numbered, "List Paragraph" style (or "List Bullet"/"List Number"). What matters is that the items are a list rather than prose; the choice between bullet and number follows whatever the document already uses.
- **Step-by-step lists**: decimal numbered, "List Paragraph" style (or "List Number").
- **One numbering sequence per section**: every section that contains a numbered list gets its OWN sequence, restarting at 1. A document covering several topics gives each topic its own list. Word counts all paragraphs sharing a numbering id as a single continuous list in document order, so when two sections share one, the second carries on from the first — "Step by Step – Alternate" opening at 9. Nothing looks wrong on screen, because Word does show numbers; they are simply the wrong ones, and the defect only surfaces when someone is told to follow step 9 of a four-step procedure.
- **Warning callout**: a paragraph with **FFF3CD** shading (background), **7B5C00** text color, starting with `⚠ `. Use sparingly, only for real risks.
- **Author name**: always Title Case (e.g., "First Last"), never all caps.

### Header version vs. Version History (critical rule)

The **header**'s version number must always be identical to the most recent version in the **Version History** table and in the "Version" column of the master-index spreadsheet. Any edit to the document — including purely formatting/consistency fixes (not just a content/process change) — must: bump the version (e.g. `1.0`→`1.1` for a small adjustment), update the header, add a row to the Version History describing the change, and update "Version" and "Last Updated" in the master-index spreadsheet.

## 5. Screenshots

- When creating a new document with no real screenshots available yet, insert a visible placeholder (`[Insert screenshot: <description>]`) after the relevant step and tell the person about it at the end.
- When inserting real screenshots (or editing an existing document that already has images): preserve the capture's **original aspect ratio** (don't distort it) and size it only for good human legibility — no other fixed size/ratio is required.

## 6. Language and tone

- The documents themselves follow the **target document language** confirmed in Section 0 for that base — this is independent from the language of this SKILL.md (which stays English). Once a base's target language is confirmed, every document in it — body text, section headings, spreadsheet columns, everything — stays in that one language; never ask again for the same base, and never mix languages within one document or one base.
- Third-party product/menu names stay exactly as they appear on the real screen, in whatever language that UI actually is (don't translate button/screen names).
- Use a direct, imperative tone, and short sentences per step, in the target document language.

## 7. The master-index spreadsheet

### Bundled tooling

This plugin ships two scripts alongside the standard, both under `scripts/`. Prefer them over improvising equivalent logic — they exist so the same operation runs the same way every time.

Both live in the `scripts/` directory next to this SKILL.md — resolve paths relative to that directory, not to the current working directory, which won't be the skill's folder when installed as a plugin. Before the first run: `pip install -r <skill dir>/scripts/requirements.txt`. If Python or the dependencies aren't available in the environment, say so and stop — do NOT improvise an equivalent check by hand, since the whole point of a fixed script is that the audit runs identically every time.

- `check_master_list.py <base path>` — read-only audit of a whole base against this standard: broken references, orphaned files, duplicate codes, name and version divergence across all four sources, legacy Version History tables, headers missing the tab separator (glued or space-separated) or missing their version entirely, documents that exist but cannot be parsed, Section 3 structure (required sections present, in order, Version History last), folder/tab mismatch, and the Section 4 formatting rules (margins, orientation, title, heading styles, callouts, footer, header dash, list formatting and author casing). It reports and recommends; it never edits. Exit codes: `0` clean, `1` problems found, `2` couldn't run. Formatting checks are conservative — when Word's run/style cascade doesn't state a property, it is listed as "could not be verified" rather than called a violation.
- `bootstrap_master_list.py <base path>` — the only script that writes. Generates a master-index spreadsheet by scanning a base that doesn't have one yet, filling in only what it can read, marking every row unreviewed, and never assigning a KB code to a document that doesn't already have one. Refuses to overwrite an existing spreadsheet without `--force`.

Where a base's index lives is a choice made once: the bootstrap asks (or takes `--output`) and records the answer, together with the base's language and default owner, in a `.kb-compiler.json` at the base root. Read that file to learn a base's setup instead of asking again, and treat the index it points at as the single place new rows are inserted.

Both depend on the pinned versions in `scripts/requirements.txt`. `tests/run_tests.py` verifies every detection against a synthetic fixture — run it after changing either script.

Note the overlap with the rest of this section: the workflow below has you create and update the spreadsheet conversationally, one document at a time, which is right when working document by document. `bootstrap_master_list.py` is for the other case — a folder of existing documents that needs indexing in one pass. And whenever you finish a batch edit, run `check_master_list.py` rather than eyeballing consistency.

The `Master List - <Base Name>.xlsx` spreadsheet (translate the file's own name and internal labels too, when the target document language isn't English, keeping the same structure) is the source of truth for every documented process.

### Updating an existing master-index spreadsheet

Whenever a document is created or revised:

- In the matching category tab, add/update the row: `Code | Document | File | Status | Version | Creation Date | Last Updated | Owner`.
  - `Code`/`Document` must match the title/file name (Section 2).
  - `File`: if the document is in the Legacy folder, use the full relative path (`Legacy/<file name>`) — for every other category, use just the file name (no folder), since the tab itself already indicates the category/folder.
  - `Version` must match the header/Version History (Section 4).
  - `Status`: `Active`, `In review`, or `Legacy`.
  - `Owner`: the name confirmed in Section 0, unless told otherwise.
- **Codes are never duplicated or shared** between two different rows — documents without their own code use `—`, never a shared generic code.
- Before finalizing any batch edit, run the bundled fixed script `scripts/check_master_list.py` (included with this plugin — don't write a new one each time) pointing at the base's root directory: `python3 scripts/check_master_list.py "<base path>"`. It deterministically checks, the same way every time: broken references, orphaned files, duplicate codes, name divergence (file/title/header/master-index spreadsheet), version divergence, Version History tables still in the old format, headers missing the tab separator or the version itself, and documents that exist but cannot be parsed — and at the end lists (only as a portability note, never as an error) file names containing non-ASCII characters. The script never decides or applies a fix on its own: for divergent names, it recommends which source to keep based on which one was changed most recently (file mtime vs. the spreadsheet's "Última Atualização"/"Last Updated"), but it's up to the person running it to choose the final version and apply it everywhere. Remember: resolving a name divergence edits content in up to 4 places and therefore also requires bumping the document's version and adding a Version History row — the script's own report already flags this.
- Ask whether the "Overview" tab also needs updating (it has cross-tab formulas — be careful not to break them).

### Creating a master-index spreadsheet from scratch (brand-new knowledge base)

Whenever this standard is applied to a brand-new knowledge base that doesn't have an index spreadsheet yet, **create the master-index spreadsheet following this structure** (don't invent a different format):

- **"Overview" tab** (always the first tab): a title row (`Knowledge Base — <Label>`), then a table with the columns `Category | Total | Active | In review | Tab`, one row per category/folder the base will have, with the "Tab" column referencing the matching tab (e.g., `→ Tab '<Category>'`).
- **One tab per folder category** (same name as the folder, without the numeric prefix), each with: a title row (`Knowledge Base — <Category>`) and the index table with columns `Code | Document | File | Status | Version | Creation Date | Last Updated | Owner`.
- Legacy documents in a new base follow the same convention: they can stay listed in their original topical category tab (with Status = "Legacy") or get their own "Legacy" tab — ask for the preference if it isn't clear, but never share a code between rows.
- Name the spreadsheet file following the same pattern: `Master List - <Base Name>.xlsx`.
- Don't skip this step even if the person doesn't explicitly ask for a spreadsheet — a brand-new knowledge base, by definition of this standard, always has a master-index spreadsheet from the very first document.

## 8. Workflow for creating or reviewing a document

1. If not already clear in this conversation: confirm the root directory, the organization/team label, the default owner, and the target document language (Section 0).
2. Ask/confirm: the process's title, category/folder, user profiles, any naming convention involved, prerequisites, whether a Verification section is warranted. If the person is instead handing over a raw draft or a document in another format to convert, see Section 9 first — it replaces this step's Q&A with extraction from that source, not the rest of the workflow.
3. If this is a brand-new base (no master-index spreadsheet yet), create the spreadsheet first (Section 7).
4. Check the master-index spreadsheet for the next available `KB-XXX` (new document) or the current version (existing document).
5. Generate/edit the `.docx` following Sections 2–5, ensuring name consistency (Section 2) and version consistency (Section 4), in the confirmed target document language (Section 6).
6. Update the master-index spreadsheet (Section 7) and confirm before saving.
7. Deliver the file and point out where screenshot placeholders need to be replaced.
8. When fixing inconsistencies in bulk (an audit), handle one at a time, show the result before moving to the next, and always ask when there's a naming/content decision that isn't obvious — never assume silently. **Run the remediation over a given document exactly once.** If a batch is applied in more than one sweep — say, formatting first and missing sections afterwards — every document caught by both gets two version bumps and two near-identical Version History rows, which is the very outcome the batching rule in 9 below exists to avoid. Before appending a row, check whether the document already carries one for this same change.
9. **When the audit reports a required section as missing, read the document before writing anything.** "Missing" as reported means "not found as a native Heading 1", which covers two very different situations:
   - **The section exists but its heading is off-standard** — styled Normal, or Heading 2, or a custom style. This is the common case by far. Promote that heading to native Heading 1 and change nothing else. Do NOT author a new section: the content is already there, and adding one duplicates it.
   - **The section genuinely does not exist.** Then read the document, synthesize the section from its actual content, and write it — for Purpose, one to three sentences on what the document solves; for Prerequisites, a single `N/A` bullet when the author defined none (see Section 3).

   Either way it is an edit, so the critical version rule in Section 4 applies: bump the version, update the header, add a Version History row, and update the master-index spreadsheet. Because each of those costs a version, batch every fix a document needs into ONE pass rather than making separate passes for structure, header and formatting.

## 9. Ingesting a raw draft or another source format

This skill also converts an existing source into a document that follows this standard, instead of building the content from scratch through Q&A: the person's own rough notes or an old undocumented write-up, a document in plain text or Markdown, `.docx` in some other/informal shape, `.pdf`, `.pptx`, or an Outlook `.msg` export. The rest of the workflow (Section 8) still applies in full — this section only replaces how the raw content is gathered, not whether the person reviews and confirms the result before it's saved.

- **Content and order**: extract the source's text content and treat its original order as a first draft of the sequence, then actively remap it onto this standard's section structure (Section 3) — Purpose, Prerequisites, Step by Step, and so on. Don't assume the source already separates these cleanly. A raw draft typically has prerequisites buried mid-paragraph, no explicit purpose statement, and context interleaved with instructions — the job is genuine restructuring into the standard's shape, not a reformat that keeps the source's own organization.
- **`.msg` specifically**: separate the live message body from quoted-reply/forwarded thread history before extracting steps. An email thread's visual top-to-bottom order is not the process's logical order once older quoted messages are involved — extract from the topmost/live content unless an earlier message in the thread is clearly part of the same procedure.
- **Images, in any source format**: extract embedded images preserving their original order of appearance in the source, and insert each one as a positioned placeholder immediately after the step it most plausibly belongs to, based on that order/proximity in the source — never collect extracted images at the end of the document. When a given image's step can't be confidently determined (the source was too loosely organized, or remapping the order broke the original proximity), still place it at your best guess but flag it explicitly to the person when delivering the document (e.g., "these N images could not be confidently anchored to a specific step — please confirm placement") rather than silently guessing and staying quiet about the uncertainty.
- **KB-XXX numbering and master-index registration stay exactly as described in Sections 2 and 7** — ingesting from a raw source does not change that. In particular, do not compute a number and register the document in the master-index spreadsheet automatically as part of an ingestion batch without the person confirming first: an extraction error (wrong category, mis-mapped title, or a raw document that turns out to already be covered by an existing `KB-XXX`) is much cheaper to fix before a code is burned and a row exists than after — codes are never recycled (Section 2), and once a document is in the master-index spreadsheet other people treat it as the source of truth. Draft the standardized document and the proposed master-index row, show them, and only save/register once confirmed — for a batch of several ingested documents, one consolidated confirmation covering all of them is enough, it doesn't need to be per document.

## 10. Producing documents in another language

English is this standard's canonical language. A base may be authored in any
language: confirm it once in Section 0, translate every literal term listed
throughout this document into that language, and keep it consistent for the
whole base.

The bundled scripts recognize equivalent terms in more than one language when
reading an existing base, so a base authored in another language stays fully
auditable. That recognition is a compatibility feature of the tooling, not a
second canonical form of the standard.
