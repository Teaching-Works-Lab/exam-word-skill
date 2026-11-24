---
name: exam-word
description: Use when exam or test-paper files from any department or subject need Word-to-Word comparison, format-difference checking, standard-template normalization, exam DOCX repair, PDF-to-editable-Word conversion, or authoring a new exam paper from a question bank / syllabus; triggers include 试卷改格式, 对比两个Word, 格式差异, 套用试卷模板, 修正试卷格式, PDF转试卷Word, 试卷排版, 页码, 得分表, 出题, 组卷, 生成试卷, 试卷明细, A/B卷, 得分表留空, 小问逐行, 选项跨页, 程序阅读题, 查错题, 代码块排版, 代码不跨页, 题号编排, 小分标注, 整数分值, and pagination drift. Do not use for ordinary non-exam Word editing.
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

1. Read `references/exam-authoring-pitfalls.md` in full before touching the builder. It is the authoritative checklist for content sourcing, score-table blanking, sub-question splitting, pagination guards, and render acceptance; every rule there came from a real rework. For **code-heavy papers** (程序阅读 / 查错与测试 / 含多行代码的试卷), §9 is mandatory: build with `scripts/build_code_exam.py` instead of `build_from_spec.py`, which has no `code` block and cannot keep a code block on the same page as its options.
2. Produce a **paper-detail alignment draft** first (依据材料 / 题库盘点 / 卷面结构 / 逐题明细 / 单元与课程目标覆盖 / 待确认事项). Do not build or deliver the final DOCX until the user confirms the exam type, section structure, score split, and paper count (A only or A/B).
3. Keep one structured JSON as the single content source. Generate the spec, the detail sheet, and the separate answer key from one generator script so they cannot drift; never hand-edit the DOCX.
4. Apply the gates in `references/exam-authoring-pitfalls.md` §5 (verify → content checks → PDF render → per-page split detection → PNG visual review) before delivery. `verify_docx.py` passing is necessary but **not sufficient**.

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
| Hand-edit the generated DOCX | Change the spec/generator and rebuild |
| Render a code-heavy paper with `build_from_spec.py` | Use `scripts/build_code_exam.py` (pitfalls §9.5) |
| Print internal unit IDs like `A9+A10` on the paper | Keep them in the teacher detail sheet only; number questions per section (pitfalls §9.1) |
| Use fractional marks or put the score before the sentence | Integers only, suffix `（N 分）` at sentence end, sync the answer key (pitfalls §9.2) |
| Paste incomplete code fragments, with borders or line numbers | Compilable full program, no border/line number, Consolas + 宋体 (pitfalls §9.3) |

If the first action is not the router, the template is re-derived, or a source copy is retained, stop and return to this workflow.

## Stop Conditions

Pause instead of guessing when the template is not unique, source text is unreadable, formulas or image anchors are uncertain, total scores conflict, the paper-design alignment is unconfirmed, or the requested page count / split-item checks cannot pass Word rendering after three focused iterations.

## Quick Commands

Use the Python executable returned by the workspace dependency loader.

```text
python scripts/route_request.py "对比这两个 Word" standard.docx exam.docx
python scripts/compare_docx.py standard.docx exam.docx --json-output diff.json --markdown-output diff.md
python scripts/build_from_spec.py assets/reference.docx exam.json output.docx
python scripts/build_code_exam.py exam.json output.docx          # code-heavy papers (pitfalls §9)
python scripts/verify_docx.py output.docx assets/reference.docx
```

For authored papers, PDF rendering plus per-page split detection is mandatory; see `references/exam-authoring-pitfalls.md` §5. Code-heavy papers additionally follow §9 and build with `scripts/build_code_exam.py`. Every layout fix belongs in the spec/generator, never in the DOCX.
