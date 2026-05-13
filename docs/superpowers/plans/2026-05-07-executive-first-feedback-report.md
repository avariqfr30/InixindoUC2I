# Executive-First Feedback Report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the UC2 generated DOCX report show executive findings and decisions first, with technical analytics and formulas later.

**Architecture:** Keep the existing analytics engine and DOCX renderer. Add an executive-first narrative builder that produces headline bullets, decision dashboard, and management actions before the existing technical chapters. Move detailed formula wording out of the opening snapshot.

**Tech Stack:** Python, Flask report generation, pandas analytics, python-docx rendering, unittest contract tests.

---

### Task 1: Update Report Contract Test

**Files:**
- Modify: `/Users/avariqfr30/Documents/InixindoUC2I/Sentiment analyzer/tests/test_report_contract.py`

- [ ] Add/adjust assertions so `build_executive_snapshot()` contains `## Executive Brief`, `### Headline untuk Manajemen`, `### Decision Dashboard`, and `### Prioritas Aksi Manajemen`.
- [ ] Assert the snapshot does not contain `Formula Experience Index`.
- [ ] Assert formula details remain in technical report sections.

### Task 2: Implement Executive-First Snapshot

**Files:**
- Modify: `/Users/avariqfr30/Documents/InixindoUC2I/Sentiment analyzer/report_narratives.py`

- [ ] Add helper methods for executive headline text, compact dashboard rows, and management actions from existing context metrics.
- [ ] Rewrite `build_executive_snapshot()` to start with headlines and decisions, not formulas.
- [ ] Keep `Formula Experience Index` in `_predictive_markdown()` for technical-later placement.

### Task 3: Verify Locally

**Files:**
- Verify UC2 only.

- [ ] Run `python3 -m unittest 'Sentiment analyzer.tests.test_report_contract'`.
- [ ] Run `python3 -m unittest 'Sentiment analyzer.tests.test_internal_api_settings'`.
- [ ] Run `python3 -m py_compile 'Sentiment analyzer/report_narratives.py' 'Sentiment analyzer/report_engine.py'`.
- [ ] Run `git diff --check -- 'Sentiment analyzer/report_narratives.py' 'Sentiment analyzer/tests/test_report_contract.py'`.

### Task 4: Deploy and Verify VPS

**Files:**
- Deploy: `/Users/avariqfr30/Documents/InixindoUC2I/Sentiment analyzer/report_narratives.py`
- Deploy: `/Users/avariqfr30/Documents/InixindoUC2I/Sentiment analyzer/tests/test_report_contract.py` only if tests are intentionally shipped; otherwise deploy runtime file only.

- [ ] Copy runtime file to VPS app path `/opt/apps/inixindo-feedback/current/report_narratives.py` with backup.
- [ ] Restart `inixindo-feedback`.
- [ ] Verify service is active.
- [ ] Verify `https://feedback.inworx.id/ready` returns ready.
