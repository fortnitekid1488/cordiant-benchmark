# Agent Operating Manual

## Scope

This repository centers on automating updates to the tire financial benchmark workbook.

## Files

- `Бенч финансовой отчетности_мэйджоры.xlsx` is the source workbook. Do not overwrite it unless the user explicitly asks.
- `scripts/` contains repeatable automation proofs and ETL scripts.
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

Portable equivalents after local setup:

```bash
.venv/bin/python scripts/dashboard_server.py --open
.venv/bin/python scripts/prepare_aistudio_sources.py --mode quarterly --download-mode full
.venv/bin/python scripts/prepare_aistudio_sources.py --mode annual --download-mode full
.venv/bin/python scripts/apply_aistudio_json.py --mode quarterly
.venv/bin/python scripts/apply_aistudio_json.py --mode annual
```

```bat
start_dashboard.cmd
prepare_quarterly_sources.cmd
prepare_annual_sources.cmd
update_quarterly_excel_from_aistudio.cmd
update_annual_excel_from_aistudio.cmd
```

For user-facing demos, prefer `start_dashboard.command` on macOS and `start_dashboard.cmd` on Windows; both open the same local dashboard without requiring terminal commands. Windows symlinks for `OPEN_THIS_*` are best-effort only, so rely on `outputs/latest_aistudio_*_package_path.txt` or the dashboard status when checking the current package.

The public GitHub setup keeps generated outputs, local caches, internal Codex notes, and generated workbook files ignored. The source workbook `Бенч финансовой отчетности_мэйджоры.xlsx` is project data and is expected to be committed by default. Do not stage `outputs/`, `.venv/`, `.playwright-*`, `.Codex/`, or other spreadsheet files unless the user explicitly clears that exact artifact for publication.

## Verification

For any workbook automation change:

1. Read the exact workbook structure before changing scripts.
2. Generate a copied workbook under `outputs/`; do not edit the original.
3. Include machine-readable provenance for every extracted value.
4. Compare extracted values against existing workbook values where a comparable period exists.
5. Treat values without source URL plus page/table evidence as review-only.

## Conventions

- Prefer official company filings, regulatory feeds, and investor-relations documents over scraper mirrors.
- Use structured sources without an LLM when available.
- Use LLM extraction only behind a strict JSON schema with period, unit, source, citation, and confidence fields.
- Keep source provenance explicit; do not collapse annual reports, IR decks, transcripts, and finance portals into vague "internet source" labels.
- Keep AI Studio batches within the current context budget. The preparer defaults to regional batches of up to 6 companies, estimates prompt plus upload-source input against a 64k-token budget, auto-splits oversized batches, converts `.xlsx` and HTML-like source pages to `.txt`, uses compact SEC extracts instead of full Company Facts JSON, and mirrors only upload-ready attachments into each batch's `FILES_FOR_AI_STUDIO` folder. Use `--max-companies-per-batch 1` only as the fallback if AI Studio web errors return.
- Keep the dashboard as a thin local control surface over the scripts. Source preparation, JSON validation, and Excel writing must remain in Python so the UI can be replaced without changing the data pipeline.
- Treat `.Codex/skills/taste-skill/` as a local reference, not a global Codex install. Read its `SKILL.md` only for frontend landing-page, portfolio, or redesign work; it explicitly does not target dashboards, data tables, or multi-step product UI.
