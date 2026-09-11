# How to use KB-Compiler

A short guide: installing, creating your first document, and auditing an
existing base. No prior knowledge of plugins is assumed.

If you only want to audit a folder of documents and don't care about the rest,
skip to [Auditing without installing anything](#auditing-without-installing-anything).

---

## 1. Install

### Claude desktop app (recommended)

1. Open the Claude desktop app.
2. Click the **account menu** in the bottom-left corner, then **Settings**.
3. Open **Plugins** and click **Add**. You get two options: add a marketplace,
   or upload one directly.

**Upload is the most reliable path.** Download the release zip from the
[Releases page](https://github.com/avatarpds/KB-Compiler-Project/releases) and
upload it. This works even in organizations whose policy restricts which plugin
marketplaces users are allowed to add — if you work in corporate IT, you have
probably run into that before.

**Marketplace** is the alternative, if your organization allows it. Choose that
option and enter:

```
https://github.com/avatarpds/KB-Compiler-Project
```

### Terminal

```
/plugin marketplace add avatarpds/KB-Compiler-Project
/plugin install kb-compiler@kb-compiler
```

Type these into the Claude prompt, not into your shell. The same managed-policy
restriction applies here.

### Choosing a scope

You'll be asked where the plugin should be available:

| Scope | Who gets it | When to choose it |
|---|---|---|
| **User** | You, in every project | The usual choice |
| **Project** | Everyone working in this repository | A team standardizing on it |
| **Local** | You, in this repository only | Trying it out |

### Updating

Third-party marketplaces don't auto-update. To pick up a later version:

```
/plugin marketplace update kb-compiler
```

If you installed by upload, download the newer release zip and upload it again.

---

## 2. Set up a base

A "base" is just a folder of documents. The first time you use the plugin on
one, it asks four things and then reuses the answers for every later document:

- **Root directory** — where the base lives
- **Organization label** — what appears in each document's footer
- **Default owner** — who is recorded as the author
- **Document language** — the language the documents are written in

English is the standard's canonical language, but the documents themselves can
be in any language. Say which one; the plugin does not guess it from the
language you are chatting in.

The answers are stored in a `.kb-compiler.json` at the base root, along with
where the master index lives.

---

## 3. Create your first document

Just ask, in plain language. There is no command to memorize — the skill
activates on its own when the conversation is about knowledge-base documents.

```
Create a KB about resetting a user password, in <folder>
```

What happens:

1. The content is structured into the standard's sections
2. The document is formatted to the specification
3. A KB code and an index row are **proposed** for you to confirm
4. Once you confirm, the file is saved and the index is updated

Step 3 is deliberate. Codes in this standard are never recycled, so assigning
one is irreversible — the plugin never does it unsupervised.

### Turning an existing draft into a document

Same idea, different input. A rough draft, a PDF, a slide deck, an Outlook
`.msg`, an old `.docx` nobody standardized:

```
Turn this PDF into a KB document in <folder>
```

The content and its order are extracted and remapped onto the standard's
structure. Images are carried over as positioned placeholders; any image that
couldn't be confidently anchored to a step is flagged rather than guessed.

### Reviewing or updating

```
Review KB-004 against the standard
Bump KB-012 to version 2.0 and update the index
```

---

## 4. Audit the whole base

Creating documents correctly is only half of it. The audit checks that the base
as a whole is still consistent.

```
pip install -r skills/kb-compiler/scripts/requirements.txt
python3 skills/kb-compiler/scripts/check_master_list.py "<base path>"
```

Paths are relative to the plugin's own folder, not to your current directory.

### What it checks

| Check | Detects |
|---|---|
| Broken references | An index row pointing at a file that doesn't exist |
| Orphaned files | A document with no row in the index |
| Ambiguous file reference | A bare file name that exists in several folders |
| Duplicate codes | The same `KB-XXX` on two different rows |
| Name divergence | File name, title, header and index disagreeing |
| Version divergence | Header, Version History and index disagreeing |
| Legacy version table | A Version History still in the old column format |
| Header missing tab | Name and version glued together (`Namev1.2`) |
| Structure violations | A required section missing, out of order, or content after Version History |
| Folder/tab mismatch | A row on one category tab whose file lives in another folder |
| Formatting violations | Margins, orientation, title, heading styles, callouts, footer, list formatting, author casing |

### Reading the report

Each section states how many findings it holds, then lists them. Findings name
the document as `[Tab] KB-XXX`.

Two things to know:

**It never edits anything.** For a name divergence it recommends which source to
keep, based on which one changed most recently, and lists every variant it
found. The choice and the edit stay with you.

**Items under "could not be verified" are not problems.** Word formatting
cascades from run to style to base style; when no level states a property, its
real value comes from Word's defaults and can't be read reliably. Those are
listed separately and never counted, because a report full of false alarms is
one nobody trusts.

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Base is clean |
| `1` | Findings exist |
| `2` | Couldn't run at all — no index, bad path, or aborted |

Distinct codes mean you can use the audit as a gate before a batch edit, or in a
CI-style check.

### Fixing findings

Work in this order: broken references and duplicate codes first, then name and
version divergence, then structure and formatting.

After correcting a document, bump its version and add a Version History row.
Standardizing a name is a content edit — fixing one inconsistency without
recording it creates another.

---

## 5. Auditing without installing anything

The scripts are plain Python and don't need the plugin, the desktop app, or
Claude at all. Clone the repository and run the audit against any folder:

```
git clone https://github.com/avatarpds/KB-Compiler-Project
cd KB-Compiler-Project
pip install -r skills/kb-compiler/scripts/requirements.txt
python3 skills/kb-compiler/scripts/check_master_list.py "<base path>"
```

If the base has no master index yet, generate one:

```
python3 skills/kb-compiler/scripts/bootstrap_master_list.py "<base path>"
```

It asks where the index should live, then scans the folder. It is deliberately
conservative: it fills in only what it can actually read, leaves the rest blank
and reports it, writes every row as unreviewed, and never assigns a KB code to a
document that doesn't already have one.

---

## Troubleshooting

**The plugin doesn't seem to activate.** It triggers on conversations about
knowledge-base documents. Be explicit the first time: "create a KB article
about…", "review this KB against the standard". Check it is listed under
Settings → Plugins.

**Adding the marketplace fails.** Your organization may restrict which
marketplaces can be added. Use the upload path with the release zip instead.

**`/plugin` isn't available.** Use the plugin browser in the desktop app, or
declare the plugin under `enabledPlugins` in `.claude/settings.json`.

**The audit reports a lot on an existing base.** Expected on the first run
against documents written before the standard existed. Treat it as an inventory
of the current state, not a list of urgent corrections. Documents whose status
marks them as legacy are skipped by the structure checks.

**`ModuleNotFoundError`.** Install the pinned dependencies:
`pip install -r skills/kb-compiler/scripts/requirements.txt`

**The audit refuses to run with exit code 2.** Either the base has no master
index (generate one with the bootstrap script), or it has more than one and the
script won't guess which is authoritative. Delete the extra, or pass the right
file name as the second argument.

---

## Verifying the tooling itself

If you change either script, run the test suite:

```
python3 skills/kb-compiler/tests/run_tests.py
```

It builds a synthetic base containing one deliberate instance of every
detection and checks that each one still fires. A detection that silently stops
working is the worst failure mode for an audit tool: the report comes back clean
and you believe it.

---

MIT licensed. See [CHANGELOG.md](CHANGELOG.md) for version history.
