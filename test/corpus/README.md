# Golden corpus

Drop real `.docx` files in this directory. The harness picks up every `.docx`
here automatically:

```bash
npm run test:roundtrip
```

**Use real documents.** Synthetic test files only exercise the paths someone
already thought of. Aim for twenty to thirty, deliberately varied: a resume, a
cover letter, a contract with numbered clauses, a report with tables and
figures, something exported from Google Docs, one with tracked changes, one
with a letterhead, one in landscape, one in an unusual font, and at least one
that is genuinely ugly.

`npm run corpus` writes a small **seed** corpus here - synthetic, but covering
structure this app does not itself generate (tables, headers and footers,
multi-level numbering, landscape and legal page sizes, tracked changes,
content controls). It is a floor to develop against, not the corpus.

The contents of this directory are gitignored apart from this file, so real
documents you add stay out of version control.
