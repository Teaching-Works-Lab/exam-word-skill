---
name: exam-word
description: Use when exam or test-paper files from any department or subject need Word-to-Word comparison, format-difference checking, standard-template normalization, exam DOCX repair, PDF-to-editable-Word conversion, or authoring a new exam paper from a question bank / syllabus; triggers include 试卷改格式, 对比两个Word, 格式差异, 套用试卷模板, 修正试卷格式, PDF转试卷Word, 试卷排版, 页码, 得分表, 出题, 组卷, 生成试卷, 试卷明细, A/B卷, 得分表留空, 小问逐行, 选项跨页, 程序阅读题, 查错题, 代码块排版, 代码不跨页, 题号编排, 小分标注, 整数分值, 参考答案, 评分细则, 答案卷, 答案速查, A/B卷答案对照, and pagination drift. Do not use for ordinary non-exam Word editing.
---

# Exam Word

## Overview

Use the retained standard exam template as the formatting authority and the supplied Word/PDF as the content authority. Apply the workflow across departments and subjects, including engineering, computer science, mathematics, and general education. Keep comparison read-only unless the user explicitly requests correction.

Preserve subject-specific content exactly: department and course names, questions, source code, formulas, diagrams, scoring, and answer-space intent come from the source. The template controls page geometry, typography, tables, headers, footers, and pagination—not academic content.

## First Action: Route Deterministically

Before choosing tools, inspecting content, or loading a document sub-skill, run `scripts/route_request.py` with the exact user request and attached paths. The command's stdout is the mode contract. Do not substitute a generic document-diff route.

After routing, use `documents:documents` for DOCX creation, editing, rendering, and visual QA. Use `pdf:pdf` whenever a PDF is an input.

| Mode | Trigger | Default output |
|---|---|---|
| `compare` | Two Word files plus 对比/比较/差异 | JSON and Markdown difference report |
| `inspect` | One Word plus 检查/核对, or unclear intent | Read-only report |
| `normalize` | One Word plus 改格式/修正/套用/统一/标准化 | New corrected DOCX |
| `pdf-to-word` | PDF plus 转换/可编辑/套模板 | New editable DOCX |
| `compare-and-fix` | Two Word files plus comparison and correction intent | Report and new corrected DOCX |
| `answer-keyed-paper` | **Any request mentioning 答案 / 参考答案 / 标准答案 / 评分标准** (default answer output) | New answer-keyed DOCX derived from the student-paper spec |
| `answer-booklet` | Explicit 教师用 / 答案册 / 评分细则 / 答案速查 / 对照表 | New teacher answer-booklet DOCX |

**Default answer output**: from 2026-11-24, "答案" alone means the **卷面标注答案卷** (pitfalls §11) — the full student paper with answers filled into the objective brackets and rubric blocks replacing the answer space. The separate teacher booklet (§10) is now an explicit opt-in, not the default. `route_request.py` accepts an answer request with **no attached file**; do not ask the user to supply one.

When intent remains ambiguous, choose `inspect`. Never overwrite an input.

Explicit invocation is `$exam-word`. Also allow implicit invocation for clear natural-language exam requests such as “把这份计算机试卷改成标准格式”“对比这两份试卷 Word”“把试卷 PDF 转成可编辑 Word”. A department name alone is not a trigger; the request must concern an exam/test paper and one of the supported operations.

## Stable Workflow

1. Record input paths and hashes in a task-local run manifest.
2. For Word comparison, run `scripts/compare_docx.py`; keep content and formatting findings separate.
3. For Word normalization, extract exam content into structured JSON and run `scripts/build_from_spec.py` with `assets/reference.docx`.
4. For PDF conversion, inspect every source page, extract/OCR into the same structured JSON, then build directly from the template. When the user says “标准模板” without attaching another template, `assets/reference.docx` is the standard; do not ask the user to supply it again. Do not create an intermediate generic Word or use full-page images as the editable result.
5. Run `scripts/verify_docx.py`, render with the Documents workflow, and inspect every page. Treat requested page count as a render gate, not an OOXML estimate.

Read `references/modes.md` for mode-specific inputs and outputs. Read `references/template-contract.md` before any build or template change.

## Generating a New Exam (Question Bank / Syllabus)

When the user asks to **author** a paper (出题 / 组卷 / 生成试卷 / 出一份试卷明细) instead of converting or comparing attached files, the `compare`/`normalize` router modes do not apply. Use this sequence:

0. **Authoring is a three-layer pipeline (§15): L1 命题设计 → L2 题目生成 → L3 格式渲染.** Set up `assets/pipeline.template.json` as `pipeline.json` for the paper, declaring each layer's inputs / build / gates, then use `scripts/exam_pipeline.py` (`init` → `status` → `run`). When ANY layer changes, re-run the pipeline from that layer so all downstream gates re-execute automatically — never hand-pick build steps. This is mandatory: skipping it is exactly how "changed upstream, forgot a downstream gate" reworks happen. Cross-cutting checks (content safety, code completeness, A/B consistency) must still be run at every delivery.
1. Read `references/exam-authoring-pitfalls.md` in full before touching the builder. It is the authoritative checklist for content sourcing, score-table blanking, sub-question splitting, pagination guards, and render acceptance; every rule there came from a real rework. For **code-heavy papers** (程序阅读 / 查错与测试 / 含多行代码的试卷), §9 and §12.8 are mandatory: build with `scripts/build_code_exam.py` instead of `build_from_spec.py`, which has no `code` block and cannot keep a code block on the same page as its options.
2. Produce a **paper-detail alignment draft** first (依据材料 / 题库盘点 / 卷面结构 / 逐题明细 / 单元与课程目标覆盖 / 待确认事项). Do not build or deliver the final DOCX until the user confirms the exam type, section structure, score split, and paper count (A only or A/B).
3. Keep one structured JSON as the single content source. Generate the spec, the detail sheet, and the separate answer key from one generator script so they cannot drift; never hand-edit the DOCX.
4. Apply the gates in `references/exam-authoring-pitfalls.md` §5 (verify → content checks → PDF render → per-page split detection → PNG visual review) before delivery. `verify_docx.py` passing is necessary but **not sufficient**.
5. For the **answer key**, the default is the answer-keyed paper (§11) — "答案 / 参考答案 / 标准答案 / 评分标准" all produce it:
   - **Default: answer-keyed paper** → read pitfalls §11 and §12, then choose the derivation path by what you have:
     - Structured student-paper spec available → `scripts/build_answer_keyed_paper.py <paper-spec.json> <objective.json> <answer.md> <out.docx> [--spec-out keyed-spec.json]`.
     - Only a **finalized student-paper DOCX** plus answers (no spec/spacer) → `scripts/derive_answer_keyed_from_docx.py <student.docx> <answers.json> <out.docx>` (pitfalls §12.4). Parse the answer Markdown into `answers.json`; never hand-code per-course question data into the script. Both paths keep the 学号/姓名 column and blank score table, fill objective answers into the existing brackets, and replace answer space with one uniform `参考答案与评分要点：` label plus scored points.
   - Then run `scripts/verify_answer_keyed_paper.py <out.docx>` and `scripts/verify_docx.py <out.docx> assets/reference.docx`, render to PDF, and review every page for orphan pages/labels. The verifier is structure-agnostic (pitfalls §12.1–12.2) and must pass on **four different paper structures (incl. a table-layout student paper)** before a script change is trusted (§12.5).
   - **Explicit teacher booklet only** (教师用 / 答案册 / 评分细则 / 答案速查 / 对照表) → read pitfalls §10 and build with `scripts/build_answer_docx.py`: four-line teacher header, no 学号/姓名/score table, unbreakable code blocks and table rows.
   Never decide by the filename alone, and never use `build_answer_docx.py` for the default answer output.

If the syllabus says 考查 (no written exam) while the user requests a written final, flag the conflict explicitly instead of silently changing the assessment design.

## AI 生成痕迹门禁

Before delivering an exam paper, answer key, scoring rubric, or its PDF/DOCX output, check both visible text and document properties for generation-disclosure artifacts, including `AI生成`, `由 AI 生成`, `部分内容可能由 AI 生成`, model or product names, prompts, and `generated by` metadata.

- Do not treat a legitimate academic reference to artificial intelligence as a generation marker.
- When the user authorizes correction, remove these markers and generation metadata from unsigned working files before final rendering and delivery.
- Signed, stamped, or other evidentiary originals must not be edited. Report the issue and request a clean source or reissued version.
- A file that still contains a generation-disclosure marker must not be used as the final exam, answer, or archive sample. Re-scan edited DOCX/PDF files and verify that the cleanup has not removed answers, scoring points, or altered pagination.

## Retention Contract

Keep only the canonical template, scripts, references, tests, final deliverables, and a compact run manifest. Record source paths and hashes, but do not retain a duplicate source Word/PDF inside the task directory or Skill. Store OCR images, rendered PDFs/PNGs, unpacked OOXML, candidate DOCX files, and caches in one task directory.

On success, delete the task directory after final outputs and the manifest are copied out. On failure, keep the minimal diagnostic directory and report its path; never move failed candidates into the delivery folder.

Do not copy source exam PDFs, generated exams, student data, credentials, or QA renders into this Skill or its repository.

## Common Drift

| Temptation | Required choice |
|---|---|
| Use a generic Documents diff helper first | Run this Skill's router, then `scripts/compare_docx.py` |
| Ask for a template when none is attached | Use `assets/reference.docx` |
| Keep a duplicate source file for traceability | Keep only its original path and hash |
| Interpret “处理一下” as permission to edit | Route to read-only `inspect` |
| Print section scores in the header score table | Leave the 分数/得分 rows blank; put scores in section headings |
| Keep （1）（2）（3）（4） in one paragraph | Split into indented one-line blocks (pitfalls §4.4) |
| Add `keep_next` to every question | Skip judgment questions without options (pitfalls §4.2) |
| Deliver once `verify_docx.py` passes | Render PDF and run per-page split checks (pitfalls §5) |
| Generate A and B from the same pool and eyeball for overlap | Run the A/B dedup gate: stems, option sets, answer sequence, subjective-answer similarity (pitfalls §8) |
| Check only formatting and assume the questions are sound | Run `scripts/check_paper_content.py` for intra-paper duplication, front-to-back answer leakage, undefined slogan-style stems, and unit coverage; a paper with objective answers that reproduce later rubric points or asks students to guess自造代称 is not deliverable (pitfalls §13) |
| Trust leakage findings from an objective.json that came from another course/paper (same question count) | `check_source_alignment` fails fast (exit 2) when objective stems don't match the spec; only read `leakage` when `source_alignment.verdict==pass` (pitfalls §13.8) |
| Hand-edit the generated DOCX | Change the spec/generator and rebuild |
| Reuse the student-paper layout for a §10 teacher booklet | Teacher header, no 学号/姓名/score table (pitfalls §10.2); for an answer-keyed paper the student layout is required (pitfalls §11) |
| Build an 答案卷 with the §10 teacher-header builder | Default to the answer-keyed paper from the student spec (§11.1); §10 booklet is an explicit opt-in |
| Produce a separate teacher booklet when the user only says 答案 | Route to `answer-keyed-paper` and use `build_answer_keyed_paper.py` |
| Run `verify_docx.py` on a user-resaved DOCX | Treat it as 人工修订态: content gates only, no byte-identical check (pitfalls §11.5) |
| Hand-copy A/B answer differences into a comparison table | Parse both answer files and compute the counts (pitfalls §10.5) |
| Render a code-heavy paper with `build_from_spec.py` | Use `scripts/build_code_exam.py` (pitfalls §9.5) |
| Print internal unit IDs like `A9+A10` on the paper | Keep them in the teacher detail sheet only; number questions per section (pitfalls §9.1) |
| Use fractional marks or put the score before the sentence | Integers only, suffix `（N 分）` at sentence end, sync the answer key (pitfalls §9.2) |
| Paste incomplete code fragments, with borders or line numbers | Compilable full program; professional code block = clean NO shade / NO border (teacher's final call), monospace 10pt, 13.5pt line spacing, inner-indent, Consolas + 宋体 (pitfalls §9.3 v5, §12.8); verify with `scripts/check_code_layout.py` (default `plain`); compile every 'run the code' block and compare output with the answer key |
| Leave answers concentrated on one option (B×6/10) or all multiple-choice = ABC | Rebalance by permuting option positions without changing any question/answer content, then gate with `scripts/check_answer_distribution.py` — hard rule for EVERY paper (pitfalls §14) |
| Hand-pick build/gate steps after an upstream change (change answers → regenerate by memory) | Set up the three-layer `pipeline.json` and run `scripts/exam_pipeline.py` so downstream layers auto-rerun with fail-fast gates (pitfalls §15) |
| Bake Markdown backticks / `**` / `&lt;` into the answer surface (``输出：`sum=10` ``) — looks like "code format still wrong" but code blocks are actually fine | `clean_markdown()` in derive script before `add_run`, plus MANDATORY last gate `scripts/check_final_gate.py` on finished DOCX (code paragraphs exempt); add to render.gates (pitfalls §15.6) |
| Let an answer-key table autofit to its longest cell | Lock total width to the text column (pitfalls §10.4-5) |
| Leave 「参考答案」/「评分要点」 as a bare label line at page bottom | Bind it to the first bullet with keep_with_next (pitfalls §10.4-4) |
| Hard-code 18/6 question counts, `FULL_MARKS`, or 425/-142 indent in a verifier | Derive everything from the paper; regression-test on 3 structures (pitfalls §12.1, §12.5) |
| Insert an answer block right after the stem or after a chapter heading | Insert after the question's last element, stopping at `^[一二三四五六]、` (pitfalls §12.3) |
| Hand-code one course's answers into a "generic" derivation script | Externalize answers to JSON; let array lengths drive question counts (pitfalls §12.4) |
| Rebuild a finalized student-paper DOCX from spec just to add answers | Derive in DOCX body and restore the 7 template parts (pitfalls §12.4) |
| Bold a stem or `(1)(2)(3)(4)` sub-question stem | Never bold any stem; keep bold only on headers, section titles, answer labels, table headers. For finalized DOCX use `scripts/unbold_substems.py` (pitfalls §12.9) |

If the first action is not the router, the template is re-derived, or a source copy is retained, stop and return to this workflow.

## Stop Conditions

Pause instead of guessing when the template is not unique, source text is unreadable, formulas or image anchors are uncertain, total scores conflict, the paper-design alignment is unconfirmed, or the requested page count / split-item checks cannot pass Word rendering after three focused iterations.

## Quick Commands

Use the Python executable returned by the workspace dependency loader.

```text
python scripts/route_request.py "对比这两个 Word" standard.docx exam.docx
python scripts/compare_docx.py standard.docx exam.docx --json-output diff.json --markdown-output diff.md
python scripts/build_from_spec.py assets/reference.docx exam.json output.docx
python scripts/build_code_exam.py exam.json output.docx          # code-heavy papers (pitfalls §9, §12.8)
python scripts/build_answer_docx.py answer.json answer.docx     # explicit teacher booklet (pitfalls §10)
python scripts/build_answer_keyed_paper.py exam-spec-A.json objective.json answer.md answer-keyed.docx --spec-out keyed-spec.json
# 评分要点是自然语言（非「(1) N 分」结构）时，显式给定分值向量（pitfalls §12.11）
python scripts/build_answer_keyed_paper.py exam-spec-A.json objective.json answer.md answer-keyed.docx --marks-json marks.json
python scripts/derive_answer_keyed_from_docx.py student-final.docx answers.json answer-keyed.docx  # finalized DOCX path (§12.4)
python scripts/verify_answer_keyed_paper.py answer-keyed.docx
python scripts/route_request.py "出一下答案"                     # -> answer-keyed-paper (no attachment needed)
python scripts/check_paper_content.py exam-spec.json objective.json answer.md --blueprint blueprint.json --md-out content-report.md  # §13: check report.source_alignment.verdict==pass BEFORE trusting leakage; fail-fast exit 2 on cross-course mismatch (§13.8)
python scripts/check_paper_content.py spec.json objective.json answer.md --gate-mode alignment  # §15: pipeline L2 gate — source_alignment is hard, content errors downgraded to warn (§13.7)
python scripts/verify_docx.py output.docx assets/reference.docx
python scripts/check_code_completeness.py paper.md   # §12.8: run-the-code blocks must be full programs
python scripts/check_code_layout.py output.docx --body-indent 283  # §9.3: code-block visual gate (plain=NO shade/border + monospace/size/indent/no-grid/no-line-number/pPr order); --style shaded for shaded variant; exit 1 on FAIL
python scripts/check_answer_distribution.py answers-A.json         # §14: objective-answer distribution gate (single <=40%/no run>2, multi combo<=2 & distractor spread, judge true in 30-70%); exit 1 on FAIL
python scripts/check_pdf_pages.py paper.pdf --min-last-page-chars 60  # L3: PDF blank/lone-page gate (needs pymupdf)
python scripts/exam_pipeline.py init    --config pipeline.json     # §15: record per-layer input hashes baseline
python scripts/exam_pipeline.py status  --config pipeline.json     # §15: show dirty layers + where to re-run from
python scripts/exam_pipeline.py run     --config pipeline.json     # §15: re-run from earliest dirty layer, fail-fast gates, refresh baseline
python scripts/exam_pipeline.py run     --config pipeline.json --from content --dry-run  # §15: preview plan
python scripts/check_final_gate.py paper.docx answer.docx    # §15.6: MANDATORY last gate — scan finished DOCX surface text for Markdown/HTML/AI residue; exit 1 on FAIL; add to render.gates LAST
python scripts/unbold_substems.py paper.docx answer.docx  # §12.9: strip bold from (1)(2) sub-question stems, keep title/header bold
```

For authored papers, PDF rendering plus per-page split detection is mandatory; see `references/exam-authoring-pitfalls.md` §5. Code-heavy papers additionally follow §9 and build with `scripts/build_code_exam.py`. Every layout fix belongs in the spec/generator, never in the DOCX.
