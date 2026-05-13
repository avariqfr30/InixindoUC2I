# Internal API Semantic Ingestion Design

**Goal:** Ensure APIDog/Internal API datasets are transformed into app-specific context before they reach document and visual generation.

**Approved Direction:** Use strict semantic adapters with graceful fallback. Required production fields must be validated. Incomplete live datasets should not silently produce misleading reports.

## Architecture

Each use case gets a small semantic adapter between raw APIDog records and the existing report/proposal/visual engines.

- UC1 Proposal Generator: adapter converts account and project history into natural client context, relationship evidence, and capability evidence without leaking dataset names or all-caps/raw labels.
- UC2 Feedback Intelligence Analyzer: adapter converts `ClassReport` and `ReferenceClassReport` into report-ready feedback rows, cleans question labels, separates numeric ratings from text answers, maps feedback to service/customer-journey dimensions, and preserves the explicit no-date limitation.
- UC3 Payment Predictor: adapter validates and normalizes future finance invoice records before analytics. If the finance endpoint is empty, production mode remains unavailable and demo fallback stays explicit.

## Non-Goals

- Do not invent missing dates, invoice values, or project standards.
- Do not expose APIDog credentials or raw internal endpoint values in repo docs.
- Do not change user-visible core flows except clearer, cleaner generated report context.

## Validation

- Unit tests cover semantic label cleaning, strict field mapping, and graceful fallback behavior.
- Each app keeps existing demo mode working.
- VPS deploy verification must include service status and `/ready` or `/health` checks if deployment is requested.
