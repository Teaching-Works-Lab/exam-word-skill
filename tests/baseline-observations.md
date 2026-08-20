# Baseline Observations

Date: 2026-08-20

Three fresh agents received realistic requests without the proposed Skill. All produced plausible plans, but their defaults diverged enough to make repeated execution slow and inconsistent.

## Scenario 1: Word-to-Word Comparison

Observed behavior:

- Selected a broad OOXML plus rendering comparison.
- Proposed keeping generated PNG, PDF, text diff, structure summaries and optional unpacked OOXML until inspection ended.
- Did not define a stable machine-readable report schema or a canonical template asset.

Failure addressed by the Skill: use a fixed comparison report schema, separate content from formatting differences, and apply one cleanup contract.

## Scenario 2: PDF-to-Template Word

Observed behavior:

- Correctly treated the PDF as content authority and the Word template as formatting authority.
- Proposed a template-part hash allowlist and six-page render gate.
- Chose a low-level targeted OOXML replacement strategy that is safe but expensive to rediscover for every task.

Failure addressed by the Skill: retain the template, preserved-part list, exam intermediate representation and builder so these decisions are reused rather than re-derived.

## Scenario 3: Repeated Workflow Retention

Observed behavior:

- Proposed thirteen separate scripts, a schema, golden DOCX files, visual baselines, package locks and multiple specialized extractors.
- This is comprehensive but exceeds the currently proven scope and would increase maintenance cost.
- Cleanup and failure-retention advice differed from Scenario 1.

Failure addressed by the Skill: keep the smallest proven toolset—router, comparator, builder and verifier—while delegating OCR/rendering to the existing PDF and Documents skills.

## Baseline Conclusion

The missing capability is not basic reasoning; it is deterministic reuse. The Skill must freeze routing, template authority, preserved OOXML parts, output contracts, privacy boundaries and cleanup behavior without duplicating mature PDF/OCR/render helpers.
