# Lumen — transparent data analysis workspace

Lumen is a React + TypeScript dashboard with an Express API for profiling uploaded datasets, inspecting data quality, and applying explicitly selected cleaning operations.

## Features

- Batch upload up to 10 CSV, XLS/XLSX, and text-based PDF files at once.
- Independent dataset profiles and a dataset selector to switch between successful uploads.
- Schema profiling with missing-cell counts, unique values, numeric ranges, means, and sample previews.
- User-controlled cleaning options: trim text, normalize headers, remove fully empty rows, drop rows with missing cells, mean-impute numeric blanks, and remove exact duplicate rows.
- Transparent cleaning audit trail with applied/skipped steps and before/after row counts. Cleaning exports a new CSV and does not overwrite the in-memory source dataset.
- Deterministic Q&A and clearly labelled baseline estimates.

## Run locally

Requires a recent Node.js release compatible with Vite 8.

```bash
npm install
npm run dev:full
```

The Vite frontend runs on its usual development port and proxies `/api` requests to the Express API at `http://localhost:8787`. You can also run `npm run dev` and `npm run api` in separate terminals.

## Upload limits and PDF support

- Up to 10 files per batch.
- Maximum 25 MB per file.
- PDFs are text-extracted and profiled as line-number/text records. Scanned/image-only PDFs are not OCR'd and will return a clear error.
- The API keeps uploaded datasets in process memory; restarting the API clears them. This is intended for a local project/demo, not multi-user production hosting.

## Transparency notes

- Cleaning operations run only when their checkboxes are selected. The output is a separate CSV, and the audit response includes the chosen options and row counts.
- “Evaluate baseline” and “Estimate mean” are statistical baselines, not trained ML models. The current model endpoints should not be presented as production model training.
- “Ask Lumen” currently uses deterministic numeric summaries; an LLM provider has not been connected.
- CSV parsing and PDF extraction are not substitutes for OCR, semantic validation, or domain-specific data quality review.

## Checks

```bash
npm run build
npm run lint
```
