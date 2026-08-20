---
name: mechanical-exam-word
description: Use when mechanical college exam files need Word-to-Word comparison, format-difference checking, standard-template normalization, exam DOCX repair, or PDF-to-editable-Word conversion; triggers include 对比两个Word, 格式差异, 套用标准模板, 修正试卷格式, PDF转Word, 试卷排版, 页码, 得分表, and pagination drift.
---

# Mechanical Exam Word

## Overview

Use the retained mechanical college exam template as the formatting authority and the supplied Word/PDF as the content authority. Keep comparison read-only unless the user explicitly requests correction.

## First Action: Route Deterministically

Before choosing tools, inspecting content, or loading a document sub-skill, run `scripts/route_request.py` with the exact user request and attached paths. The command's stdout is the mode contract. Do not substitute a generic document-diff route.

After routing, use `documents:documents` for DOCX creation, editing, rendering, and visual QA. Use `pdf:pdf` whenever a PDF is an input.

| Mode | Trigger | Default output |
|---|---|---|
| `compare` | Two Word files plus 对比/比较/差异 | JSON and Markdown difference report |
| `inspect` | One Word plus 检查/核对, or unclear intent | Read-only report |
| `normalize` | One Word plus 修正/套用/统一/标准化 | New corrected DOCX |
| `pdf-to-word` | PDF plus 转换/可编辑/套模板 | New editable DOCX |
| `compare-and-fix` | Two Word files plus comparison and correction intent | Report and new corrected DOCX |

When intent remains ambiguous, choose `inspect`. Never overwrite an input.

## Stable Workflow

1. Record input paths and hashes in a task-local run manifest.
2. For Word comparison, run `scripts/compare_docx.py`; keep content and formatting findings separate.
3. For Word normalization, extract exam content into structured JSON and run `scripts/build_from_spec.py` with `assets/reference.docx`.
4. For PDF conversion, inspect every source page, extract/OCR into the same structured JSON, then build directly from the template. When the user says “标准模板” without attaching another template, `assets/reference.docx` is the standard; do not ask the user to supply it again. Do not create an intermediate generic Word or use full-page images as the editable result.
5. Run `scripts/verify_docx.py`, render with the Documents workflow, and inspect every page. Treat requested page count as a render gate, not an OOXML estimate.

Read `references/modes.md` for mode-specific inputs and outputs. Read `references/template-contract.md` before any build or template change.

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

If the first action is not the router, the template is re-derived, or a source copy is retained, stop and return to this workflow.

## Stop Conditions

Pause instead of guessing when the template is not unique, source text is unreadable, formulas or image anchors are uncertain, total scores conflict, or the requested page count cannot pass Word rendering after three focused iterations.

## Quick Commands

Use the Python executable returned by the workspace dependency loader.

```text
python scripts/route_request.py "对比这两个 Word" standard.docx exam.docx
python scripts/compare_docx.py standard.docx exam.docx --json-output diff.json --markdown-output diff.md
python scripts/build_from_spec.py assets/reference.docx exam.json output.docx
python scripts/verify_docx.py output.docx assets/reference.docx
```
