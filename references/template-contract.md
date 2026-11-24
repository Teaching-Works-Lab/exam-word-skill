# Standard Template Contract

## Canonical Asset

`assets/reference.docx` is the sole formatting authority. Do not regenerate it from measurements and do not overwrite it. Before each build, record its SHA-256 hash in the run manifest.

## Geometry and Styles

- Paper: A4 portrait, 11906 × 16838 twips.
- Margins: left 1701, right 1134, top 1134, bottom 1134 twips.
- Header distance: 851 twips; footer distance: 992 twips.
- Normal: Chinese 宋体, Latin Times New Roman, 12 pt, exact 20 pt line spacing.
- Heading 1: 16 pt bold, centered, 30 pt line spacing.
- Score table: three rows and eight columns with exact template geometry.
- Footer: dynamic `PAGE` and `NUMPAGES` fields.

## Preserved OOXML Parts

The builder and verifier require these template parts to remain byte-identical when they exist:

- `word/styles.xml`
- `word/numbering.xml`
- `word/theme/theme1.xml`
- `word/footer1.xml`
- `word/fontTable.xml`
- `word/footnotes.xml`
- `word/endnotes.xml`

The final `sectPr` and page geometry must also match the template. A visual match without these structural checks is insufficient.

## Score Table Values

The three-row score table is a **grading sheet for the marker**, not a student-facing scorecard. For newly authored papers, leave the 分数 and 得分 rows empty (including the total cell); only the 题号 header row is filled. Section scores belong in the section headings. Do not copy the numeric `["分数","40","60","100"]` shape from `tests/` fixtures into a real paper — those fixtures only verify builder behaviour. See `references/exam-authoring-pitfalls.md` §3.

## Content Authority

- For PDF input, the PDF controls text, question order, figures, scores, and intended page boundaries.
- For Word input, the source Word controls content unless the user identifies another content authority.
- The template controls page geometry, styles, score table, headers/footers, numbering definitions, and visual system.

## Render Gate

Render the final DOCX in Microsoft Word when available and inspect every page PNG at 100%. LibreOffice rendering may be used as an additional check but does not replace Word compatibility when exact pagination is required.
