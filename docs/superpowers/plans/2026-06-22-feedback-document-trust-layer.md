# Feedback Document Trust Layer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Strengthen factuality, traceability, and reader reassurance in generated feedback reports without changing the web UI, raw datasets, public generation API, or adding model calls.

**Architecture:** Add one focused deterministic trust module that builds a filter-scoped fact registry, anonymous snapshot fingerprint, contradiction summary, and conservative numeric-preservation checks. Feed this packet into the existing deliberation contract before narrative generation, enrich the existing appendix and prescriptive action table, and keep the final preflight as the blocking boundary.

**Tech Stack:** Python standard library, pandas, unittest, existing report pipeline and DOCX renderer.

---

### Task 1: Deterministic trust packet

**Files:**
- Create: `Sentiment analyzer/report_factuality.py`
- Create: `Sentiment analyzer/tests/test_report_factuality.py`

- [ ] **Step 1: Write failing tests**

Test that `ReportFactRegistry.build(...)` returns stable fact IDs, filter-scoped counts, confidence bases, contradiction data, and an anonymous fingerprint that changes when input facts change but does not expose record content.

- [ ] **Step 2: Verify RED**

Run: `python -m unittest tests.test_report_factuality -v`

Expected: import failure because `report_factuality` does not exist.

- [ ] **Step 3: Implement the minimal trust module**

Implement:

```python
class ReportFactRegistry:
    @classmethod
    def build(cls, dataframe, analysis_context, scope): ...

class NarrativeFactValidator:
    @classmethod
    def numeric_tokens(cls, text): ...

    @classmethod
    def preserves_numeric_facts(cls, original, candidate): ...
```

The registry must use only normalized, filter-scoped data; calculate an anonymous SHA-256 fingerprint; and expose no raw comments or record identifiers.

- [ ] **Step 4: Verify GREEN**

Run: `python -m unittest tests.test_report_factuality -v`

Expected: all factuality tests pass.

### Task 2: Pre-writing integration and conservative editing

**Files:**
- Modify: `Sentiment analyzer/report_pipeline.py`
- Modify: `Sentiment analyzer/report_evidence.py`
- Modify: `Sentiment analyzer/feedback_deliberation.py`
- Modify: `Sentiment analyzer/writing_quality.py`
- Test: `Sentiment analyzer/tests/test_enterprise_content_deliberation.py`
- Test: `Sentiment analyzer/tests/test_editorial_quality.py`

- [ ] **Step 1: Write failing integration tests**

Require `trust_packet`, `data_version`, scoped confidence bases, and contradiction guidance in the deliberation contract. Require the editor to reject a polished candidate that adds or changes numeric facts.

- [ ] **Step 2: Verify RED**

Run: `python -m unittest tests.test_enterprise_content_deliberation tests.test_editorial_quality -v`

Expected: assertions fail because the trust packet and numeric-preservation gate are absent.

- [ ] **Step 3: Integrate the trust packet before narrative generation**

Build it from the same filtered dataframe used by analytics, pass it through section context, include its fingerprint as `data_version`, and place contradiction/confidence guidance in the hidden deliberation payload.

- [ ] **Step 4: Add conservative editor fallback**

After an optional language-polish call, return the original deterministic narrative whenever numeric tokens differ.

- [ ] **Step 5: Verify GREEN**

Run: `python -m unittest tests.test_enterprise_content_deliberation tests.test_editorial_quality -v`

Expected: all selected tests pass.

### Task 3: Recommendation traceability and appendix reassurance

**Files:**
- Modify: `Sentiment analyzer/narratives/prescriptive.py`
- Modify: `Sentiment analyzer/feedback_deliberation.py`
- Modify: `Sentiment analyzer/report_quality.py`
- Test: `Sentiment analyzer/tests/test_report_contract.py`
- Test: `Sentiment analyzer/tests/test_enterprise_content_deliberation.py`

- [ ] **Step 1: Write failing contract tests**

Require every prioritized action row to carry an evidence ID, owner, review window, success indicator, and expected outcome. Require the appendix to show the anonymous snapshot fingerprint and confidence basis without raw dataset labels.

- [ ] **Step 2: Verify RED**

Run: `python -m unittest tests.test_report_contract tests.test_enterprise_content_deliberation -v`

Expected: traceability and fingerprint assertions fail.

- [ ] **Step 3: Implement localized narrative and validator changes**

Extend only the existing action matrix and appendix. Validate each action-table row rather than scanning the entire chapter for isolated keywords.

- [ ] **Step 4: Verify GREEN**

Run: `python -m unittest tests.test_report_contract tests.test_enterprise_content_deliberation -v`

Expected: all selected tests pass.

### Task 4: Regression, deployment, and live verification

**Files:**
- Modify only files already listed above.

- [ ] **Step 1: Run local regression suite**

Run all `Sentiment analyzer/tests/test_*.py` tests in a clean local virtual environment.

- [ ] **Step 2: Inspect diff and repository state**

Confirm no UI templates, connector policy, raw data, account databases, environment files, or unrelated files changed.

- [ ] **Step 3: Inspect production read-only and record preservation anchors**

Confirm `/opt/apps/inixindo-feedback/current`, `inixindo-feedback.service`, the production virtualenv, feedback row count, user count, and existing health/readiness endpoints.

- [ ] **Step 4: Create a neutral production backup and deploy focused files**

Back up changed production source files under a timestamped `pre-document-trust-layer-*` directory. Transfer only verified source and test files; do not replace production databases, connector configuration, profiles, job state, or generated reports.

- [ ] **Step 5: Run VPS targeted tests and restart**

Run the targeted test modules with the production virtualenv, restart `inixindo-feedback.service`, and poll loopback `/health` and `/ready` through any startup warm-up.

- [ ] **Step 6: Run a real report smoke and public verification**

Use the existing authenticated production flow to generate and download one report, inspect its extracted DOCX text for fingerprint, evidence IDs, traceable actions, factuality gate success, and source-identifier leakage. Verify `https://feedback.inworx.id` and re-check preservation counts.

- [ ] **Step 7: Clean temporary artifacts**

Remove uploaded staging files, temporary report probes, extracted DOCX directories, and transient health files. Preserve the neutral rollback backup.
