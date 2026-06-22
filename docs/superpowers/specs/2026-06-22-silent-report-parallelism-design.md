# Silent Report Parallelism Design

## Goal

Reduce report-generation latency and improve concurrent-job stability without changing hardware, service configuration, public APIs, visible UI/UX, generated wording, data handling, or account behavior.

## Constraints

- Keep the existing two report workers and current model/provider configuration.
- Keep DOCX assembly serial and preserve section order.
- Keep OSINT timeout/error fallback and writing-quality factuality fallback unchanged.
- Emit performance timings to application logs only; do not add fields to job or quality payloads.
- Use only bounded in-memory state. Do not create persistent caches or change schemas.
- Preserve exact narrative inputs and deterministic output content.

## Design

### 1. Stage timing and research/analysis overlap

`ReportPipeline` will record monotonic duration for research, analysis preparation, narrative, writing, preflight, rendering, and final quality. One summary log entry is emitted after a successful run.

Research is submitted to a two-slot orchestration executor while the local analytics engine and its prepared analysis are built on the request thread. Research resolution retains the current 45-second timeout and fallback text. Stage injection remains supported for tests and alternate runtimes.

### 2. Per-job prepared analysis

Each `FeedbackAnalyticsEngine` instance will own a single prepared-analysis packet keyed by timeframe, sentiment, segment, and score engine. It contains the scoped dataframe and deterministic derived values already reused by narrative, executive, trust, and final-quality paths.

The packet is request-local because a new analytics engine is created per report. No mutable dataframe or engine is shared across jobs. Existing methods remain callable without preparation and compute the same results on demand.

Small instance-local memoization will reuse governance, theme, and grouped-risk calculations for the same scoped dataframe. The cache is bounded to the lifetime of the report engine.

### 3. Bounded writing concurrency

Snapshot and section polishing will use one process-wide executor with two workers. Results are collected in original order before document-spine repair. The existing editor remains responsible for issue detection, protected facts, model fallback, and final validation.

The global two-task bound prevents two report jobs from multiplying model calls beyond the current server capacity. Failures continue to return the deterministic source text.

### 4. Bounded chart-byte cache

Chart renderers will cache immutable PNG bytes in a small process-wide LRU keyed by complete chart specification, color, and render version. Each caller receives a fresh `BytesIO` cursor.

A render lock protects Matplotlib cache misses because `pyplot` is process-global and not thread-safe. DOCX insertion remains serial within each report.

## Verification

- Unit tests prove research and analytics overlap, stage timings stay log-only, prepared analysis is computed once, polishing never exceeds two concurrent tasks and preserves order, and chart streams are independent.
- Existing report, factuality, quality, and architecture tests must pass.
- A deterministic local baseline and post-change document are compared by extracted paragraph/table content and quality result.
- One-job and two-job local runs are timed and checked for exceptions.
- Production deploy backs up only changed files, restarts `inixindo-feedback`, checks service state plus loopback/public health and readiness, verifies feedback/account counts, and removes temporary artifacts.

## Deferred

- Tensor/model parallelism, GPU changes, process workers, Waitress/systemd tuning, vector-index activation, and chart process pools.
