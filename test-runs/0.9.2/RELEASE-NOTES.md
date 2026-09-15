# v0.9.2 — House style belongs to the document

A small, documentation-only follow-up to 0.9.1. No code changed, so the audit
behaves exactly as it did yesterday. Nothing to migrate.

## What prompted it

While fixing up a real knowledge base, rows added to a Version History table
came out looking wrong — bigger, and a different colour than every row above
them. Nothing was broken. The table just stopped matching itself, which is
glaringly obvious to anyone who opens the document and completely invisible to
the audit.

The reason turned out to be mundane. In these documents the font, size and
colour sit on the individual runs, not on the table style. Write a row without
them and Word fills the gap with its own defaults instead of the document's
look.

0.9.1 added a rule about that for Version History rows. This release makes it
general, because the problem was never really about that one table.

## What changed

Section 4 now says out loud something it had only implied:

- The properties Section 4 lists are the standard's. Those don't move.
- Everything it doesn't list — table fonts and borders, body text size,
  spacing, column widths — belongs to the document. Read it off the document
  and match it.

And when a document's own style genuinely clashes with a Section 4 rule, ask
instead of quietly rewriting it. Standardise it, or keep the house style and
accept that the audit will flag it — your call, not the tool's. Only when
there's a real clash, though; if the standard is silent there's nothing to
decide and nothing to ask about.

None of this is new thinking, really. Section 4 already worked this way for one
property: whether Prerequisites uses bullets or numbers "follows whatever the
document already uses". This just applies the same idea everywhere it belongs.

## Why there's no new check for it

The audit enforces six constants, and a history table's font was never one of
them — so matching the document can't produce a finding, and there's nothing to
reconcile.

A check was the obvious alternative and it's the worse one. Judging whether a
row's font "matches" means resolving a Word style cascade that most documents
leave unstated, which is exactly the false-positive territory the conservative
mode exists to stay out of. Not writing the problem beats detecting it after
the fact.
