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

## `answer-keyed-paper` (default answer output)

Inputs: either (a) the student-paper spec (`exam-spec-*.json`), same-source objective-answer JSON and answer Markdown; or (b) a **finalized student-paper DOCX** plus an `answers.json` derived from the answer Markdown. **No attached file is required to route here** — any request mentioning 答案 / 参考答案 / 标准答案 / 评分标准 lands on this mode. See `references/exam-authoring-pitfalls.md` §11 and §12.

Run:

```text
# spec path
python scripts/build_answer_keyed_paper.py exam-spec-B.json objective.json answer.md answer-keyed.docx --spec-out keyed-spec.json
# finalized-DOCX path (no spec / no spacer blocks)
python scripts/derive_answer_keyed_from_docx.py student-final.docx answers.json answer-keyed.docx
python scripts/verify_answer_keyed_paper.py answer-keyed.docx
python scripts/verify_docx.py answer-keyed.docx assets/reference.docx
```

Output: one new DOCX beside the student paper; never overwrite the source. Gates: the content checks from `verify_answer_keyed_paper.py` (answer label, student header, blank score table, **every** objective bracket filled, **one** uniform rubric label per subjective question, mark format and per-question sums, label bound to first bullet, label not misplaced after a chapter heading, no answer space, consistent sub-question indent, no hard page break, no AI marks). Counts, section names, mark values and indent are **derived from the paper** (pitfalls §12.1–12.2); never assume 18 objective / 6 subjective. If a subjective rubric cannot be parsed into integer marks, pass `--marks-json` rather than hand-editing the DOCX. A script change is only trusted after the three-structure regression in pitfalls §12.5.

## `answer-booklet` (explicit opt-in)

Inputs: an answer spec (`answer-spec-*.json`). Trigger **only** when the request explicitly says 教师用 / 答案册 / 评分细则 / 答案速查 / 对照表. Build with `scripts/build_answer_docx.py`, then `verify_docx.py` + PDF review. See `references/exam-authoring-pitfalls.md` §10. This is not the default answer output.

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

### Code-heavy papers

Papers containing multi-line source code (程序阅读 / 查错与测试 / 代码填空) are built with `scripts/build_code_exam.py`, **not** `build_from_spec.py`. It accepts the same `metadata` / `score_table` / `pages` structure above plus two extra block types:

```json
{"type": "code", "lines": ["#include <stdio.h>", "int main(void) {", "    return 0;", "}"]}
{"type": "option-row", "items": ["A. ...", "B. ..."]}
```

`code` renders as a borderless, unnumbered block (Consolas + 宋体, 9.5pt / 12.5pt line spacing) indented 425 twips; `option-row` places two or four options on one line using widened tab stops. The builder wraps each question in a borderless table row with `w:cantSplit` so a stem, its code and its options can never be split across pages. `build_from_spec.py` does not recognise `code` and must not be used for these papers. See `references/exam-authoring-pitfalls.md` §9.

### Answer key / rubric (explicit opt-in only)

**Only for `answer-booklet`.** 参考答案与评分细则、客观题答案速查、A/B 卷答案对照用 `scripts/build_answer_docx.py` 构建，不走学生卷路线：

- **教师卷头**：学校 / 学期·命题单位·命题教师 / 课程·卷别 / 「参考答案与评分细则 · 教师用（请勿下发给学生）」，**无学号姓名栏、无卷首得分表**；卷头用全角空格串联而非 tab（长学院名会把姓名挤出页外）。
- **blocks**：`heading` / `subheading`（level 3/4）/ `text`（可 `bold`）/ `bullet`（可 `mark`）/ `quote` / `table`（`rows`, `header`）/ `code`（`lines`）/ `spacer`。
- **表格**：宽表（> 5 列，如双栏答案速查）自动拆成左右两张；首行 `tblHeader` 跨页重复；≤ 8 行的小表整表不拆。
- **绑定**：「评分要点：」等以冒号结尾的标签行，若下一块是 `table` / `code`，自动与它同页。

metadata 继承试卷 spec，只改 `paper` / `summary`，并设 `doc_type: "answer"`。生成后仍走 `verify_docx.py` + PDF 逐页复核。See `references/exam-authoring-pitfalls.md` §10.

### Answer-keyed paper (卷面标注答案卷, default)

Implements the `answer-keyed-paper` mode above; full rule set in `references/exam-authoring-pitfalls.md` §11. This is **not** a §10 teacher booklet.

- Derive from the **student-paper spec**, never author a second content source: all stems, options, question numbers and sub-question indentation stay byte-identical to the student paper.
- Keep the student header (insert 「答案」 after the paper letter), the 学号/姓名 column, and the score table with empty 分数/得分 rows. Keep the same page setup. Stem blocks stay at the student paper's 20 pt; answer blocks are uniformly 18 pt — both set explicitly by the generator, not pasted from an answer booklet (§11.3-6).
- Objective answers go **into the brackets already present in the stem** (`（B）` `(ABC)` `（×）` `(√)`); read them from the objective-answer JSON, never hand-copy.
- Replace every answer-space spacer with one uniform `参考答案与评分要点：` label bound to its first bullet, then one rubric paragraph per scoring point. Sub-questions stay as separate `425/-142` blocks; rubric paragraphs align with the stem at `283/-283`.
- Marks use `（N 分）` only, integers, and **the points of each subjective question must sum to that question's full mark** — checked by both the builder and `verify_answer_keyed_paper.py`.
- Merge all spec pages into one naturally flowing page; do not reserve hard page breaks. If the last page holds only one long rubric paragraph, tighten the answer-block line spacing (18 pt) rather than hand-editing.
- A file that was opened and saved by 腾讯文档 / Word / WPS is 人工修订态: skip `verify_docx.py` preserved-part checks and run content gates + PDF page review instead (§11.5).
