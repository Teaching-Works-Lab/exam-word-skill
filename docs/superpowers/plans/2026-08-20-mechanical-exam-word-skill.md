# Mechanical Exam Word Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build, validate, install, and privately publish a reusable multi-mode Word/PDF exam formatting skill.

**Architecture:** A concise routing skill delegates deterministic work to focused Python scripts and keeps the canonical DOCX template as a versioned asset. Structured JSON separates content extraction from document generation, while OOXML and Word-render checks protect template fidelity.

**Tech Stack:** Python 3.12, `python-docx`, OOXML ZIP processing, Microsoft Word COM, pytest, Git, GitHub CLI.

**Spec:** `docs/superpowers/specs/2026-08-20-mechanical-exam-word-skill-design.md`

## Global Constraints

- Repository and GitHub visibility must remain Private.
- Never overwrite user inputs.
- Never commit source exam PDFs, generated exam outputs, QA renders, caches, credentials, or personal data.
- Python source and text files use UTF-8; Chinese paths are passed through `pathlib.Path`, not inline shell scripts.
- Template-derived output preserves the canonical template's required OOXML parts.

---

### Task 1: Baseline Skill Behavior

**Files:**
- Create: `tests/baseline-observations.md`

**Interfaces:**
- Consumes: Three realistic Word/PDF exam requests without the new skill.
- Produces: Observed routing, fidelity, and cleanup failures that the skill must address.

- [ ] Run fresh-agent scenarios without the skill.
- [ ] Record concrete omissions and unstable choices.
- [ ] Convert observed failures into behavioral requirements.

### Task 2: Router and Discovery

**Files:**
- Create: `tests/test_route_request.py`
- Create: `scripts/route_request.py`
- Create: `SKILL.md`
- Create: `agents/openai.yaml`

**Interfaces:**
- Consumes: Attached file suffixes and user request text.
- Produces: One of `compare`, `inspect`, `normalize`, `pdf-to-word`, or `compare-and-fix` with a read-only default.

- [ ] Write routing tests and verify missing-module failure.
- [ ] Implement the minimal routing API and CLI.
- [ ] Run routing tests and verify all pass.
- [ ] Add discovery metadata matching the approved trigger phrases.

### Task 3: DOCX Comparison

**Files:**
- Create: `tests/test_compare_docx.py`
- Create: `scripts/compare_docx.py`

**Interfaces:**
- Consumes: Two `.docx` paths.
- Produces: JSON data separating content, page setup, styles, tables, headers/footers, and numbering differences.

- [ ] Write comparison tests using generated DOCX fixtures.
- [ ] Verify tests fail because comparison code is absent.
- [ ] Implement OOXML comparison and Markdown reporting.
- [ ] Run comparison tests and verify all pass.

### Task 4: Template Build and Verification

**Files:**
- Create: `tests/test_build_and_verify.py`
- Create: `scripts/build_from_spec.py`
- Create: `scripts/verify_docx.py`
- Create: `references/template-contract.md`
- Create: `references/modes.md`
- Copy: `assets/reference.docx`

**Interfaces:**
- Consumes: Canonical template plus structured exam JSON.
- Produces: Editable DOCX plus machine-readable verification report.

- [ ] Write output-package and template-part hash tests.
- [ ] Verify tests fail before implementation.
- [ ] Implement minimal template cloning and content insertion.
- [ ] Implement structural and optional Word-render verification.
- [ ] Run all tests and inspect representative render.

### Task 5: Packaging and Private Publication

**Files:**
- Create: `.gitignore`
- Create: `LICENSE` only if explicitly requested; otherwise omit.

**Interfaces:**
- Consumes: Verified skill directory.
- Produces: Installed personal skill and private GitHub repository `Teaching-Works-Lab/mechanical-exam-word-skill`.

- [ ] Run skill and Unicode validators.
- [ ] Inspect every staged path and secret-scan the repository.
- [ ] Initialize Git and commit only verified skill files.
- [ ] Create the GitHub repository with `--private`.
- [ ] Push the verified default branch and confirm private visibility.
