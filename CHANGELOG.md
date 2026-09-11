# Changelog

All notable changes to this project are documented here.

## 0.8.2

- Added `.claude-plugin/marketplace.json`, so the repository is both the plugin
  and its own single-plugin catalog. Installing is now two commands rather than
  a manual copy of the skill folder:

  ```
  /plugin marketplace add avatarpds/KB-Compiler-Project
  /plugin install kb-compiler@kb-compiler
  ```

  Third-party marketplaces have auto-update disabled by default, so run
  `/plugin marketplace update kb-compiler` to pick up later versions.

## 0.8.1

- Duplicate file names across folders are resolved using the row's category tab
  instead of taking the first match, which previously meant comparing a
  document the row didn't refer to. When the tab can't single one out, it is
  reported as an ambiguous reference and nothing about that document is
  checked.
- Orphan detection now records the resolved path rather than the bare file
  name, so a same-named file in another folder no longer escapes detection by
  borrowing another row's reference.

## 0.8.0

- Section 4 coverage completed: page orientation, footer presence, footer tab
  separator and automatic page fields, hyphen vs. en/em dash in the header,
  bullet/numbered list formatting per section, and all-caps author names in the
  Version History.
- Footer labels are compared across the base: every document must carry the
  same organization label.
- List formatting is resolved through built-in style names and the style chain,
  not only direct paragraph numbering — documents using Word's native
  List Bullet / List Number styles were being reported as not lists at all.
- Code-prefix stripping accepts en and em dashes, so a header using the wrong
  dash is reported once as a dash violation instead of also surfacing as a
  bogus name divergence.

## 0.7.1

- Documents sitting loose in the base root are indexed under an `Uncategorized`
  tab and individually flagged, instead of being skipped by the bootstrap and
  then reported as orphans by the checker.

## 0.7.0

- The bootstrap asks where the master index should live (default name in the
  base, another name in the base, or any path including outside it), or takes
  `--output`. The choice, the base language and the default owner are recorded
  in a `.kb-compiler.json` at the base root.
- The checker resolves the index from that configuration before falling back to
  file-name patterns, so a custom name or an out-of-base location keeps working
  as a continuous insertion point. A configured path that has gone missing is an
  explicit error, never a silent fallback.
- The interactive path asks which language the base is authored in, instead of
  defaulting to English and producing an index whose headers don't match the
  documents it indexes.

## 0.6.1

- The bootstrap refuses to run when the base already has a master index in any
  language or under any configured path. Checking only its own target file name
  meant a run without `--lang` quietly created a second index beside the first.

## 0.6.0

- English is the standard's canonical language. Literal terms are defined once
  in English; a base may be authored in any language confirmed at setup, and the
  scripts recognize equivalent terms when reading an existing base.
- Folder/tab comparison reads the Category-to-Tab mapping recorded in the
  Overview sheet instead of recomputing the transformation, which had made the
  checker report that a folder didn't match itself whenever a category name had
  been sanitized or truncated to become a legal sheet title.
- Structure checks skip documents whose Status marks them as legacy.
- Overview count formulas use a wide range so an appended row is included.

## 0.5.0

- Section 3 structure validation: required sections present, in the mandated
  order, with Version History last. An extra custom section is allowed.
- Folder/tab mismatch detection.
- Fixes found by external review: accumulate spreadsheet candidates across all
  naming patterns before deciding; read the title from the first non-empty
  paragraph; only treat a header as glued when the version is attached directly
  to a word character; sanitize Excel sheet names and disambiguate collisions;
  recognize a numbered legacy folder; reject an unknown `--lang`; return exit
  code 2 for usage and dependency errors.

## 0.4.1

- MIT license.

## 0.4.0

- Test suite covering detections, bootstrap behaviour and non-interactive
  safety.

## 0.3.0

- `bootstrap_master_list.py`: generates a master index for a base that has none,
  filling in only what it can read, marking every row unreviewed, and never
  assigning a KB code to a document that doesn't already have one.
- A missing index no longer ends in a traceback: the checker explains the
  options and, on an interactive terminal, offers to create one. Run
  non-interactively it prints the same guidance and exits rather than blocking
  on input.

## 0.2.0

- `check_master_list.py`: read-only audit covering broken references, orphaned
  files, duplicate codes, name and version divergence across four sources per
  document, legacy Version History tables, headers missing the tab separator,
  and the Section 4 formatting constants. Conservative mode: a formatting
  property the run/style cascade doesn't state is listed as unverifiable rather
  than reported as a violation.
- Pinned dependency versions.
- Ingestion of raw drafts and other formats into the standard.

## 0.1.0

- Initial documentation standard: section structure, exact Word formatting,
  global KB-XXX numbering with uniqueness constraints, and the master-index
  spreadsheet convention.
