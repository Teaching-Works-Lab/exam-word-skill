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
    "summary": "本试卷共有五道大题（满分：100 分）"
  },
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

Supported block types are `heading`, `question`, `option`, `option-row`, `text`, `spacer`, and `table`. A new entry in `pages` inserts an explicit page break.
