# Agent Operating Manual

## Scope

This repository centers on automating updates to the tire financial benchmark workbook.

## Files

- `Бенч финансовой отчетности_мэйджоры.xlsx` is the clean source template. In the public repo it should contain only the `Свод` sheet with structure/formatting/header metadata, not pre-filled benchmark values; company-detail sheets are not part of the default workflow. Do not overwrite it unless the user explicitly asks.
- `scripts/` contains repeatable automation proofs and ETL scripts.
- `scripts/launchers/advanced/` contains legacy/diagnostic launcher wrappers that should not clutter the project root.
- `outputs/` contains generated dry-run artifacts and copied workbooks.
- `.Codex/skills/taste-skill/` contains a project-local copy of `Leonxlnx/taste-skill` from `skills/taste-skill`.
- `.Codex/decisions.md` records non-obvious technical decisions.
- `.Codex/progress.md` records current state and next steps.

## Commands

Create a local Windows environment:

```bat
install_windows_requirements.cmd
```

Create a local macOS environment:

```bash
./install_macos_requirements.command
```

Run the current structured-source proof:

```bash
.venv/bin/python scripts/goodyear_sec_dry_run.py
```

Prepare AI Studio packages:

```bash
.venv/bin/python scripts/prepare_aistudio_sources.py --mode quarterly --download-mode full
.venv/bin/python scripts/prepare_aistudio_sources.py --mode annual --download-mode full
```

Apply returned AI Studio JSON:

```bash
.venv/bin/python scripts/apply_aistudio_json.py --mode quarterly
.venv/bin/python scripts/apply_aistudio_json.py --mode annual
```

Run the local dashboard:

```bash
.venv/bin/python scripts/dashboard_server.py --open
```

Root user-facing launchers after local setup:

```bash
./install_macos_requirements.command
./start_dashboard.command
./update_from_github.command
```

```bat
install_windows_requirements.cmd
update_from_github.cmd
start_dashboard.cmd
```

For user-facing demos, prefer `start_dashboard.command` on macOS and `start_dashboard.cmd` on Windows; both open the same local dashboard without requiring terminal commands. Keep root launchers limited to install/start/update. For diagnostics, use direct Python commands or the wrappers under `scripts/launchers/advanced/`. Package pointers live under `outputs/latest_aistudio_*_package_path.txt`; do not recreate root `OPEN_THIS_*` symlinks.

The public GitHub setup keeps generated outputs, local caches, internal Codex notes, and generated workbook files ignored. The clean source template `Бенч финансовой отчетности_мэйджоры.xlsx` is expected to be committed by default, while filled/generated benchmark workbooks are not. Do not stage `outputs/`, `.venv/`, `.playwright-*`, `.Codex/`, or other spreadsheet files unless the user explicitly clears that exact artifact for publication.

## Verification

For any workbook automation change:

1. Read the exact workbook structure before changing scripts.
2. Generate a copied workbook under `outputs/`; do not edit the original.
3. Include machine-readable provenance for every extracted value.
4. Compare extracted values against existing workbook values where a comparable period exists.
5. Treat values without source URL plus page/table evidence as review-only.

## Conventions

- Prefer StockAnalysis for standard income statement, balance sheet, cash-flow, and employee-count extraction because current accepted values mostly come from that source. Upload StockAnalysis pages to the LLM as the primary evidence. Do not add official/company sources for the same company and period just because they may be more authoritative; include them only when StockAnalysis is unavailable, lacks the target period, or lacks the requested metric.
- Annual StockAnalysis parsing must skip `TTM` columns and choose the latest full fiscal-year column; quarterly parsing must choose the latest standalone quarter and reject H1/H2/YTD/TTM when applying JSON.
- Use deterministic structured parsing for QA, diagnostics, and draft checks, but do not let it satisfy final dashboard JSON unless the user explicitly accepts parser output.
- Use LLM extraction only behind a strict JSON schema with period, unit, source, citation, and confidence fields.
- Keep source provenance explicit; do not collapse annual reports, IR decks, transcripts, and finance portals into vague "internet source" labels.
- Keep AI Studio batches within the current context budget. The preparer defaults to regional batches of up to 6 companies, estimates prompt plus upload-source input against a 64k-token budget, auto-splits oversized batches, converts `.xlsx`, `.pdf`, and HTML-like sources to `.txt`, uses compact SEC extracts instead of full Company Facts JSON, and mirrors only LLM evidence attachments into each batch's `FILES_FOR_AI_STUDIO` folder. Use `--max-companies-per-batch 1` only as the fallback if provider web errors return.
- Local StockAnalysis parser output is diagnostic only and belongs outside `aistudio_json/`; final dashboard builds should use JSON saved from the selected LLM provider unless the user explicitly accepts parser drafts.
- Keep the dashboard as a thin local control surface over the scripts. Source preparation, JSON validation, and Excel writing must remain in Python so the UI can be replaced without changing the data pipeline.
- Treat `.Codex/skills/taste-skill/` as a local reference, not a global Codex install. Read its `SKILL.md` only for frontend landing-page, portfolio, or redesign work; it explicitly does not target dashboards, data tables, or multi-step product UI.
