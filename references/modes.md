# Mode Contracts

## `compare`

Inputs: exactly two Word files. Identify left/standard and right/target explicitly. Run structural comparison first, then render both when visual pagination matters. Output `diff.json` and `diff.md`; do not modify either Word.

## `inspect`

Inputs: one target Word and the retained template unless the user supplies a different standard. Compare against `assets/reference.docx`. Output a read-only report grouped into content, page setup, styles, tables, numbering, headers/footers, and visual findings.

## `normalize`

Inputs: one target Word. Extract content into UTF-8 JSON, build from the canonical template, verify preserved OOXML parts, render, and inspect all pages. Output a new DOCX beside the requested delivery location; never overwrite the source.

## `pdf-to-word`

Inputs: one or more PDF exams. Render every PDF page before extraction. Use text extraction for reliable text layers and OCR only for scanned pages. Preserve source-page references and uncertainty in the JSON. Build directly from the template; do not make an intermediate generic Word.

## `compare-and-fix`

Inputs: two Word files. First complete `compare`, then use the identified standard/template and target content to complete `normalize`. Deliver the difference report and corrected DOCX.

## Structured Exam JSON

The builder accepts UTF-8 JSON with `metadata` and `pages`:

```json
{
  "metadata": {
    "school": "沧州交通学院",
    "academic_term": "2025－2026学年第二学期",
    "unit": "计算机与信息技术学院",
    "teacher": "命题教师",
    "course": "课程名称",
    "paper": "A 卷",
    "exam_mode": "闭卷",
    "summary": "本试卷共有五道大题（满分：100 分；考试时间：120 分钟）"
  },
  "score_table": [
    ["题号", "一", "二", "三", "四", "五", "", "总分"],
    ["分数", "", "", "", "", "", "", ""],
    ["得分", "", "", "", "", "", "", ""]
  ],
  "pages": [
    {
      "blocks": [
        {"type": "heading", "text": "一、单项选择题"},
        {"type": "question", "text": "1、题干"},
        {"type": "option-row", "items": ["A. 选项", "B. 选项", "C. 选项", "D. 选项"]}
      ]
    }
  ]
}
```

Supported block types are `heading`, `question`, `option`, `option-row`, `text`, `spacer`, and `table`. A new entry in `pages` inserts an explicit page break; keep these to a minimum so objective questions flow continuously instead of leaving half-empty pages.

### Score table

`score_table` is optional. For **newly authored** papers only the header row carries values: the 分数 and 得分 rows — including the total cell — must be empty strings so the marker can fill them in and sign. Put section scores in the section headings (`一、单项选择题（共 8 题，每题 2 分，共 16 分）`) instead. An omitted `score_table` also clears the cells, but passing explicit empty strings avoids stale template values. See `references/exam-authoring-pitfalls.md` §3.

### Per-block paragraph controls

All block types accept optional paragraph-format fields:

| Field | Effect | Typical use |
|---|---|---|
| `keep_next` | keep with next paragraph | bind a stem to its first option; never on option-less judgment items |
| `keep_together` | keep lines of one paragraph together | long stems that must not split mid-sentence |
| `left_twips` / `first_twips` | left indent / first-line indent (negative = hanging) | sub-question indentation |
| `before_pt` / `after_pt` / `line_pt` | space before / space after / line spacing | `spacer` uses `after_pt` for answer space |
| `align` | `left` / `center` / `right` | — |

Defaults: `question` = `left_twips 283`, `first_twips -283`; `option` = `left_twips 425`. A block's own values override these defaults.

### Sub-questions

Never put `（1）…（4）` or `①…⑤` in one paragraph. Split each item into its own `question` block with `{"left_twips": 425, "first_twips": -142, "keep_together": true}` so the number aligns with the stem text and wrapped lines hang further in. Only the last sub-question gets `keep_next` (to bind the answer space); intermediate items must not, or the chain creates large blank gaps. See `references/exam-authoring-pitfalls.md` §4.4.

### Render acceptance

`verify_docx.py` only proves template fidelity. Render the DOCX to PDF and check every page for orphan option lines, split stems, and short trailing fragments before delivery. See §5 of the pitfalls reference.
