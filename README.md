# thesis-docx-format

Turn a thesis draft `.docx` into the final, spec-compliant `.docx` — deterministically, with a
human gate on the one decision a machine should not make alone.

[中文说明 →](README.zh-CN.md)

Built for Chinese-university thesis formatting (where the spec dictates fonts, sizes, line
spacing, section-scoped page numbering and a strict front-matter order), but the engine is
school-agnostic: **one JSON profile per institution, a generic kernel underneath.**

## What it does

| Stage | What happens |
| --- | --- |
| `accept` | Accepts every tracked change, drops comments, merges paragraphs whose mark was deleted — then proves the result is character-identical to Word's "final" view |
| `outline` | Detects heading levels **even in drafts with no heading styles at all**, and emits a proposal you must confirm |
| `apply` | Applies heading/body formatting, rebuilds cover + title page + declaration from the profile, wires multi-section page numbering, runs hygiene fixes |
| `audit` | Three passes: structure/sections/page numbering, fonts/sizes/spacing, TOC & content consistency |
| `pdfcheck` | Verifies an exported PDF: page numbering sequence, embedded fonts, and whether every TOC page number matches where the heading actually landed |

Nothing in the draft's wording is ever edited. Content-level findings (half-width parentheses in
Chinese text, suspicious typos, ideographic-space padding) are **reported, never silently fixed**.

## Install

No dependencies, no build step. Python 3.9+ standard library only.

```bash
git clone https://github.com/<you>/thesis-docx-format.git
cd thesis-docx-format
bash tests/run_regression.sh          # end-to-end on synthetic drafts
```

As a Claude Code skill:

```bash
git clone https://github.com/<you>/thesis-docx-format.git ~/.claude/skills/thesis-docx-format
```

As a Codex (or any other agent) skill: copy the directory in. Every step is a plain CLI command;
nothing depends on a particular agent's tooling.

## Use

```bash
S=.   # or ~/.claude/skills/thesis-docx-format
python3 $S/scripts/tfmt.py init ./P --draft draft.docx --profile tfsu-mti
python3 $S/scripts/tfmt.py accept  --project ./P
python3 $S/scripts/tfmt.py outline --project ./P     # → P/02_outline.tsv
#   ← review the level column, change "# CONFIRMED: no" to yes
python3 $S/scripts/tfmt.py apply   --project ./P
python3 $S/scripts/tfmt.py audit   --project ./P --docx ./P/03_formatted.docx
```

`apply` refuses to run until the outline is confirmed. That is deliberate: a wrong heading level
silently corrupts the table of contents, the running headers, pagination and every page number.

## Heading detection without heading styles

The worst realistic input is a draft where every paragraph is `Normal` at the same size. So the
detector does not score paragraphs one by one. It works in two phases:

1. **Find the numbering system.** Collect leading numbers across the document
   (`第X章` / `1.1` / `1.1.1` / `一、` / `（一）` / `Chapter N`) and keep only schemes that are
   *sequentially consistent*: level 1 increments 1, 2, 3…; sub-levels restart under each parent;
   one document-wide restart is tolerated. This alone rejects flat pseudo-sequences such as
   `Case 1: … Case 27:` that look exactly like third-level headings.
2. **Score each paragraph.** Special names (acknowledgements / abstract / TOC / references /
   appendix) rank above scheme+valid-sequence, which ranks above an existing `outlineLvl`, which
   ranks above size/bold/centering/short-line heuristics. Negative signals: label brackets,
   `Case N:` prefixes, sentence-final punctuation, excessive length.

When a numbering system is already established, the shape-based fallback is switched **off** —
otherwise "University Name" on the cover gets promoted to a second-level heading.

Paragraphs that look like headings but were not classified as such are still listed as
**candidates**, because that is where genuinely missed headings are recovered.

## Profiles

`profiles/<name>.json` is data; `scripts/` is the engine. `tfsu-mti.json` is a complete, working
implementation. To support another institution, copy `profiles/_template.json` — it carries a
7-step filling guide — and edit. A profile declares page setup, body typography, three heading
levels, the front-matter layout (**including tab-stop coordinates for the cover's fill-in rules**),
the section/page-numbering plan, headers and footers, hygiene rules, report-only content checks
and hard limits (abstract length, minimum body length, title length).

## Notes worth reading before trusting any tool in this space

`references/` documents the failure modes this implementation is built around:

- **`01-outline.md`** — why sequence consistency beats per-paragraph scoring; how the TOC field
  corrupts the sequence; why `b <= a` is the wrong restart test.
- **`02-sections.md`** — `pgNumType`'s format does **not** inherit across sections; a section with
  no `headerReference` inherits the previous one's header; three page-break mechanisms stacking on
  one boundary produce blank pages; empty paragraphs before a section break push it onto the next page.
- **`03-hygiene.md`** — `"Times New Roman Regular"` is a PostScript style name, not a family name;
  Word silently falls back and reports nothing. `w14:textFill` lives in another namespace and
  survives naive cleanup while overriding `w:color`. Underlined *spaces* get trimmed at line end,
  so fill-in rules must be drawn with underlined **tabs**.
- **`04-pdf.md`** — a TOC's page numbers are cached by whichever engine laid the document out.
  Measured on a 169-page thesis: LibreOffice agreed with only 10 of 48 entries and drifted 6 pages
  by the appendix. Export the PDF from the same application that paginated it.

## Guarantees and non-goals

Every stage proves itself and refuses to emit a file when the proof fails: `accept` compares its
output against Word's final view character by character; `apply` compares the draft's body
paragraph by paragraph and aborts naming the first paragraph that changed.

Not in scope: editing content, updating TOC page numbers (that is a field refresh — F9 in
Word/WPS), exporting the PDF for you, and reformatting tables/images/footnotes (detected and
reported only).

## License

MIT — see [LICENSE](LICENSE). No institutional specification document is bundled; see
[`profiles/tfsu-mti/README.md`](profiles/tfsu-mti/README.md).
