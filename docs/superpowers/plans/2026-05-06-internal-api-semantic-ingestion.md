# Internal API Semantic Ingestion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convert APIDog/Internal API data into normalized, evaluated, report-ready context across UC1, UC2, and UC3.

**Architecture:** Add app-specific semantic adapters at the ingestion boundary. Keep raw HTTP clients and document engines separate. Preserve demo fallback and reject misleading live-data output when required fields are absent.

**Tech Stack:** Python, Flask, pandas, python-docx, existing unittest suites.

---

### Task 1: UC2 ClassReport Semantic Normalizer

**Files:**
- Modify: `/Users/avariqfr30/Documents/InixindoUC2I/Sentiment analyzer/data_pipeline.py`
- Modify: `/Users/avariqfr30/Documents/InixindoUC2I/Sentiment analyzer/tests/test_internal_api_settings.py`

- [ ] Add helpers that convert all-caps APIDog question labels into title/sentence-case Indonesian labels.
- [ ] Classify numeric answers as rating rows and non-numeric answers as comment rows.
- [ ] Map known training evaluation labels into better `Layanan`, `Customer Journey Hint`, `Lokasi`, and `Tipe Instruktur` fields where possible.
- [ ] Keep `Rentang Waktu` explicit as `Semua Data APIDog (tanggal tidak tersedia)` until APIDog provides dates.
- [ ] Run `python3 -m unittest 'Sentiment analyzer.tests.test_internal_api_settings'`.

### Task 2: UC1 Proposal Context Naturalization Guardrails

**Files:**
- Modify: `/Users/avariqfr30/Documents/Inixindo-Use-Case-1/Proposal gen/main/runtime_components.py`
- Modify: `/Users/avariqfr30/Documents/Inixindo-Use-Case-1/Proposal gen/main/text_hygiene.py`
- Modify/Add tests under `/Users/avariqfr30/Documents/Inixindo-Use-Case-1/Proposal gen/tests/`

- [ ] Normalize account/project text fields before relationship/capability summaries.
- [ ] Remove dataset-code phrasing from proposal context.
- [ ] Keep `ProjectStandards` unavailable when APIDog returns zero rows rather than treating it as real standards.
- [ ] Run focused proposal tests.

### Task 3: UC3 Finance Contract Adapter

**Files:**
- Modify: `/Users/avariqfr30/Documents/InixindoUC3/Payment predictor/financial_analyzer_metrics.py`
- Modify: `/Users/avariqfr30/Documents/InixindoUC3/Payment predictor/data_contract.py`
- Modify/Add tests under `/Users/avariqfr30/Documents/InixindoUC3/Payment predictor/tests/`

- [ ] Normalize future invoice fields into period, partner, service, payment class, invoice value, and notes.
- [ ] Reject empty/missing finance production datasets with a clear readiness state.
- [ ] Keep demo mode intact while `FinanceInvoice` has zero records.
- [ ] Run focused payment predictor tests.

### Task 4: Verification and Deployment Decision

- [ ] Run compile/check tests for all touched apps.
- [ ] Review diffs for secret leakage, raw dataset leakage, and behavior regressions.
- [ ] Deploy to VPS only after local verification passes or when explicitly requested in the implementation turn.
