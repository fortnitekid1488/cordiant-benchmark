#!/usr/bin/env python3
"""Prepare LLM source bundles for latest-period extraction.

The output is meant for a non-technical workflow:
1. Run prepare_sources.command.
2. Upload everything from a batch folder's `FILES_FOR_AI_STUDIO` folder to Qwen Studio or the selected provider.
3. Paste that batch's prompt.
4. Save the returned JSON into the prepared `aistudio_json` folder.
5. Run update_excel_from_aistudio.command.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import shutil
import sys
import time
import urllib.request
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
REGISTRY_PATH = ROOT / "config" / "company_source_registry.json"
SCHEMA_PATH = ROOT / "templates" / "aistudio_latest_quarter_schema.json"
PROMPT_TEMPLATE_PATH = ROOT / "templates" / "aistudio_prompt_template.txt"
DEFAULT_MAX_COMPANIES_PER_BATCH = 6
DEFAULT_MAX_ESTIMATED_INPUT_TOKENS = 64000
APPROX_CHARS_PER_TOKEN = 4

METRICS = [
    "Total Revenues",
    "Cost Of Revenues",
    "Gross Profit",
    "Gross Profit Margin %",
    "Selling General & Admin Expenses",
    "Other Operating Expenses",
    "Operating Income",
    "EBIT Margin %",
    "EBITDA",
    "EBITDA Margin %",
    "Net Income",
    "Accounts Receivable, Total",
    "Inventory",
    "Total Current Liabilities",
    "Total Assets",
    "Total Equity",
    "Total Debt",
    "Capital Expenditure",
    "Levered Free Cash Flow",
    "Cash from Operations",
]

COMPACT_SCHEMA_TEMPLATE = """{
  "reporting_scope": "{{REPORTING_SCOPE}}",
  "currency_rule": "Use company reporting currency; do not convert currencies.",
  "unit_rule": "Use millions. Convert billions to millions and note it.",
  "companies": [{
    "company_key": "string",
    "company_name": "string",
    "period_label": "string",
    "period_start": "YYYY-MM-DD or null",
    "period_end": "YYYY-MM-DD or null",
    "currency": "string",
    "unit": "millions",
    "facts": [{
      "metric": "requested metric name",
      "value": "number or null",
      "source_file": "uploaded file name or null",
      "source_url": "source URL or null",
      "source_type": "official_api | official_pdf | official_excel | official_ir_page | official_regulatory_portal | readable_aggregator_fallback | unknown",
      "page": "page/table page or null",
      "table_label": "table/section label or null",
      "evidence_quote": "short row/quote proving the value or null",
      "confidence": "high | medium | low",
      "review_required": true,
      "review_reason": "reason if review_required is true"
    }],
    "company_review_notes": ["string"]
  }]
}"""

MODE_CONFIG = {
    "quarterly": {
        "folder_prefix": "aistudio_quarterly_package",
        "mode_label": "quarterly_update",
        "reporting_scope": "latest reported quarter financial benchmark for tire manufacturers",
        "period_hint": "latest reported standalone quarter",
        "period_rule_1": "Use the latest reported quarter available in the sources for each company.",
        "period_rule_2": "If a source reports cumulative H1/9M/YTD numbers instead of a standalone quarter, say so in `period_label` and `company_review_notes`, and mark affected values `review_required`.",
        "period_instructions": (
            "Extract the newest reported quarterly period available in the uploaded files or listed URLs. "
            "Do not default to 2025 or Q1 2026 just because those examples appear in old files or URLs. "
            "Prefer standalone Q1/Q2/Q3/Q4 values. If only cumulative H1/9M/YTD is available, return it but mark review_required."
        ),
    },
    "annual": {
        "folder_prefix": "aistudio_annual_package",
        "mode_label": "annual_update",
        "reporting_scope": "latest full-year annual financial benchmark for tire manufacturers",
        "period_hint": "latest reported full fiscal year",
        "period_rule_1": "Use the latest full fiscal year or annual period available in the sources for each company.",
        "period_rule_2": "Do not use quarterly, interim, H1, 9M, YTD, TTM, or Q1/Q2/Q3/Q4 values for annual mode; if only those are available, set the metric to null and mark `review_required`.",
        "period_instructions": (
            "Extract the newest full-year annual period available in the uploaded files or listed URLs. "
            "Do not default to 2025 if a newer annual report exists in the sources. "
            "Ignore quarterly/interim/YTD-only documents for values in annual mode."
        ),
    },
}


def slugify(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_") or "source"


def source_filename(company_key: str, index: int, url: str, content_type: str | None) -> str:
    parsed = urlparse(url)
    path_name = Path(parsed.path).name
    suffix = Path(path_name).suffix.lower()
    if not suffix:
        if content_type and "pdf" in content_type:
            suffix = ".pdf"
        elif content_type and ("spreadsheet" in content_type or "excel" in content_type):
            suffix = ".xlsx"
        elif content_type and "json" in content_type:
            suffix = ".json"
        else:
            suffix = ".html"
    base = slugify(Path(path_name).stem if path_name else parsed.netloc)
    return f"{company_key}_{index:02d}_{base}{suffix}"


FAST_DOWNLOAD_TYPES = {"official_api", "official_pdf", "official_excel", "official_press_release"}


def schema_text_for_mode(mode: str) -> str:
    return COMPACT_SCHEMA_TEMPLATE.replace("{{REPORTING_SCOPE}}", MODE_CONFIG[mode]["reporting_scope"])


def period_hint_for_mode(mode: str) -> str:
    return MODE_CONFIG[mode]["period_hint"]


def is_quarter_specific_source(source: dict) -> bool:
    raw_text = f"{source.get('url', '')} {source.get('note', '')}".lower()
    normalized = re.sub(r"[^a-z0-9]+", " ", raw_text)
    patterns = (
        r"\bq[1-4]\b",
        r"\b[1-4]q\b",
        r"\b10\s*q\b",
        r"\bfirst\s+quarter\b",
        r"\bsecond\s+quarter\b",
        r"\bthird\s+quarter\b",
        r"\bfourth\s+quarter\b",
        r"\bthree\s+months\b",
        r"\b3\s+months\b",
        r"\bquarterly\b",
        r"\b31\s+march\b",
        r"\b30\s+june\b",
        r"\b30\s+september\b",
    )
    return any(re.search(pattern, normalized) for pattern in patterns)


def source_for_mode(source: dict, mode: str) -> dict | None:
    source_entry = dict(source)
    source_type = source_entry.get("type")
    if mode == "quarterly":
        if source_type in {"official_pdf", "official_press_release", "official_sec_filing"} and is_quarter_specific_source(source_entry):
            return None
        if source_type == "readable_aggregator_fallback":
            source_entry["note"] = "Readable fallback with latest quarterly financial statement tables."
        elif is_quarter_specific_source(source_entry):
            source_entry["note"] = "Official source page. Use the latest quarterly materials available there; do not treat older period examples as the target."
        return source_entry

    if source_type == "readable_aggregator_fallback":
        source_entry["url"] = source_entry["url"].replace("?p=quarterly", "")
        source_entry["note"] = "Readable fallback with annual financial statement tables."
        return source_entry

    if source_type in {"official_pdf", "official_press_release", "official_sec_filing"} and is_quarter_specific_source(source_entry):
        return None

    if is_quarter_specific_source(source_entry):
        source_entry["note"] = "Official source page. Use the latest annual/full-year materials available there; do not use quarterly examples as the target."

    return source_entry


def fetch_url(url: str, timeout: int = 6) -> tuple[bytes, str, str]:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml,application/pdf,application/json,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read(), response.headers.get("content-type", ""), response.url


def html_to_text(html: str) -> str:
    text = re.sub(r"<(script|style)[\s\S]*?</\1>", " ", html, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def format_cell_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value).replace("\n", " ").replace("\t", " ").strip()


def xlsx_to_text(workbook_path: Path, max_chars: int = 200000) -> str:
    """Convert official Excel downloads to a tab-separated text extract for provider upload.

    Several provider web UIs accept PDF/TXT/JSON more reliably than XLSX uploads.
    The original workbook is still kept in the package for provenance.
    """
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    sections: list[str] = [
        f"# Extracted text from {workbook_path.name}",
        "# Values are tab-separated. Blank trailing cells are removed from each row.",
        "",
    ]
    for ws in wb.worksheets:
        sections.append(f"## Sheet: {ws.title}")
        for row in ws.iter_rows(values_only=True):
            values = [format_cell_value(value) for value in row]
            while values and values[-1] == "":
                values.pop()
            if values:
                sections.append("\t".join(values))
        sections.append("")
    return "\n".join(sections)[:max_chars]


SEC_RELEVANT_TAGS = {
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "CostOfRevenue",
    "CostOfGoodsAndServicesSold",
    "GrossProfit",
    "SellingGeneralAndAdministrativeExpense",
    "SellingAndMarketingExpense",
    "GeneralAndAdministrativeExpense",
    "OperatingIncomeLoss",
    "NetIncomeLoss",
    "AccountsReceivableNetCurrent",
    "AccountsReceivableNet",
    "InventoryNet",
    "InventoryFinishedGoodsNetOfReserves",
    "LiabilitiesCurrent",
    "Assets",
    "AssetsCurrent",
    "StockholdersEquity",
    "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    "LongTermDebtCurrent",
    "LongTermDebtNoncurrent",
    "ShortTermBorrowings",
    "LongTermDebtAndFinanceLeaseObligationsCurrent",
    "LongTermDebtAndFinanceLeaseObligationsNoncurrent",
    "OperatingLeaseLiability",
    "FinanceLeaseLiability",
    "PaymentsToAcquirePropertyPlantAndEquipment",
    "NetCashProvidedByUsedInOperatingActivities",
    "DepreciationDepletionAndAmortization",
    "DepreciationDepletionAndAmortizationExpense",
}


def slim_sec_company_facts(payload: bytes, max_facts_per_unit: int = 16) -> dict:
    raw = json.loads(payload)
    us_gaap = raw.get("facts", {}).get("us-gaap", {})
    tags = {}
    for tag in sorted(SEC_RELEVANT_TAGS):
        tag_payload = us_gaap.get(tag)
        if not tag_payload:
            continue
        units = {}
        for unit_name, facts in tag_payload.get("units", {}).items():
            relevant = [
                {
                    "fy": fact.get("fy"),
                    "fp": fact.get("fp"),
                    "form": fact.get("form"),
                    "start": fact.get("start"),
                    "end": fact.get("end"),
                    "filed": fact.get("filed"),
                    "accn": fact.get("accn"),
                    "val": fact.get("val"),
                }
                for fact in facts
                if fact.get("form") in {"10-K", "10-Q"}
            ]
            relevant.sort(
                key=lambda fact: (
                    str(fact.get("end") or ""),
                    str(fact.get("filed") or ""),
                    str(fact.get("start") or ""),
                ),
                reverse=True,
            )
            if relevant:
                units[unit_name] = relevant[:max_facts_per_unit]
        if units:
            tags[tag] = {
                "label": tag_payload.get("label"),
                "description": tag_payload.get("description"),
                "units": units,
            }
    return {
        "source": "SEC Company Facts compact extract",
        "entityName": raw.get("entityName"),
        "cik": raw.get("cik"),
        "note": "Compact extract for provider upload. Monetary values are in raw SEC units; convert to millions in the answer.",
        "tags": tags,
    }


def upload_files_for_batch(batch_manifest: dict) -> list[str]:
    files: list[str] = []
    for company in batch_manifest["companies"]:
        for source in company["sources"]:
            upload_file = source.get("ai_studio_upload_file")
            if not upload_file and source.get("downloaded_file", "").endswith(".xlsx"):
                upload_file = source.get("text_extract_file")
            if not upload_file:
                upload_file = source.get("downloaded_file")
            if upload_file:
                files.append(upload_file)
    return files


def compact_source_manifest(batch_manifest: dict) -> dict:
    compact = {
        "batch_id": batch_manifest.get("batch_id"),
        "title": batch_manifest.get("title"),
        "companies": [],
    }
    for company in batch_manifest["companies"]:
        compact_company = {
            "key": company.get("key"),
            "display_name": company.get("display_name"),
            "country": company.get("country"),
            "currency": company.get("currency"),
            "period_target": company.get("period_target"),
            "sources": [],
        }
        for source in company.get("sources", []):
            upload_file = source.get("ai_studio_upload_file") or source.get("downloaded_file")
            compact_company["sources"].append(
                {
                    "source_type": source.get("type"),
                    "upload_file": Path(upload_file).name if upload_file else None,
                    "source_url": source.get("final_url") or source.get("url"),
                    "download_status": source.get("download_status"),
                    "note": source.get("note") or source.get("ai_studio_upload_note") or source.get("note_for_user"),
                }
            )
        compact["companies"].append(compact_company)
    return compact


def write_upload_list(batch_dir: Path, batch_manifest: dict) -> None:
    files = upload_files_for_batch(batch_manifest)
    lines = [
        "# Upload exactly these source files to Qwen Studio or the selected provider for this batch.",
        "# Do not upload aistudio_latest_quarter_schema.json or source_manifest.json; the prompt already includes them.",
        "# Do not upload raw .xlsx files or full SEC Company Facts JSON when an extract is listed.",
        "",
        *files,
        "",
    ]
    (batch_dir / "FILES_TO_UPLOAD.txt").write_text("\n".join(lines), encoding="utf-8")
    write_upload_folder(batch_dir, files)


def write_upload_folder(batch_dir: Path, files: list[str]) -> None:
    upload_dir = batch_dir / "FILES_FOR_AI_STUDIO"
    if upload_dir.exists():
        shutil.rmtree(upload_dir)
    upload_dir.mkdir()
    used_names: set[str] = set()
    for rel_path in files:
        source_path = (batch_dir / rel_path).resolve()
        if not source_path.is_file() or batch_dir.resolve() not in source_path.parents:
            continue
        target_name = source_path.name
        if target_name in used_names:
            target_name = f"{slugify(source_path.parent.name)}_{target_name}"
        used_names.add(target_name)
        shutil.copy2(source_path, upload_dir / target_name)


def estimated_source_chars(source_path: Path) -> int:
    if source_path.suffix.lower() == ".pdf":
        try:
            from pypdf import PdfReader

            reader = PdfReader(source_path)
            text = "\n".join(page.extract_text() or "" for page in reader.pages)
            if text.strip():
                return len(text)
        except Exception:  # noqa: BLE001 - estimate only; package generation should continue.
            pass
        return max(1, source_path.stat().st_size // 8)
    return source_path.stat().st_size


def batch_input_estimate(batch_dir: Path) -> dict[str, int]:
    prompt_path = batch_dir / "prompt_for_aistudio.txt"
    prompt_chars = prompt_path.stat().st_size if prompt_path.exists() else 0
    upload_bytes = 0
    source_chars = 0
    upload_count = 0
    for rel_path in upload_files_from_written_list(batch_dir):
        source_path = batch_dir / rel_path
        if source_path.is_file():
            upload_bytes += source_path.stat().st_size
            source_chars += estimated_source_chars(source_path)
            upload_count += 1
    total_chars = prompt_chars + source_chars
    return {
        "prompt_chars": prompt_chars,
        "upload_bytes": upload_bytes,
        "estimated_source_chars": source_chars,
        "upload_count": upload_count,
        "approx_input_tokens": math.ceil(total_chars / APPROX_CHARS_PER_TOKEN),
    }


def upload_files_from_written_list(batch_dir: Path) -> list[Path]:
    upload_list = batch_dir / "FILES_TO_UPLOAD.txt"
    if not upload_list.exists():
        return []
    files = []
    for line in upload_list.read_text(encoding="utf-8").splitlines():
        rel_path = line.strip()
        if rel_path and not rel_path.startswith("#"):
            files.append(Path(rel_path))
    return files


def write_manifest(batch_dir: Path, batch_manifest: dict) -> None:
    (batch_dir / "source_manifest.json").write_text(
        json.dumps(batch_manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def prepare_batch(
    batch: dict,
    company_lookup: dict[str, dict],
    output_dir: Path,
    schema_text: str,
    prompt_template: str,
    mode: str,
    download_mode: str,
) -> dict:
    batch_dir = output_dir / batch["batch_id"]
    sources_dir = batch_dir / "sources"
    sources_dir.mkdir(parents=True, exist_ok=True)
    batch_manifest: dict = {
        "batch_id": batch["batch_id"],
        "title": batch["title"],
        "companies": [],
    }

    for company_key in batch["companies"]:
        company = company_lookup[company_key]
        company_entry = {
            "key": company["key"],
            "display_name": company["display_name"],
            "country": company["country"],
            "currency": company["currency"],
            "period_target": period_hint_for_mode(mode),
            "sources": [],
        }
        mode_sources = [
            mode_source
            for source in company.get("sources", [])
            if (mode_source := source_for_mode(source, mode)) is not None
        ]
        for i, source in enumerate(mode_sources, start=1):
            source_entry = dict(source)
            if download_mode == "fast" and source.get("type") not in FAST_DOWNLOAD_TYPES:
                source_entry.update(
                    {
                        "download_status": "skipped_fast_mode",
                        "note_for_user": "URL is included in the manifest/prompt, but the fast preparer did not download this web page.",
                    }
                )
                company_entry["sources"].append(source_entry)
                continue
            try:
                payload, content_type, final_url = fetch_url(source["url"])
                filename = source_filename(company["key"], i, final_url, content_type)
                file_path = sources_dir / filename
                file_path.write_bytes(payload)
                source_entry.update(
                    {
                        "download_status": "downloaded",
                        "downloaded_file": str(file_path.relative_to(batch_dir)),
                        "content_type": content_type,
                        "final_url": final_url,
                        "bytes": len(payload),
                    }
                )
                if "html" in content_type.lower() or filename.endswith(".html"):
                    text_path = file_path.with_suffix(".txt")
                    text_path.write_text(html_to_text(payload.decode("utf-8", "ignore"))[:200000], encoding="utf-8")
                    source_entry["text_extract_file"] = str(text_path.relative_to(batch_dir))
                    source_entry["ai_studio_upload_file"] = str(text_path.relative_to(batch_dir))
                if filename.endswith(".xlsx"):
                    text_path = file_path.with_suffix(".txt")
                    text_path.write_text(xlsx_to_text(file_path), encoding="utf-8")
                    source_entry["text_extract_file"] = str(text_path.relative_to(batch_dir))
                    source_entry["ai_studio_upload_file"] = str(text_path.relative_to(batch_dir))
                    source_entry["ai_studio_upload_note"] = (
                        "Some provider web UIs may reject XLSX files. Upload this text_extract_file instead."
                    )
                if source.get("type") == "official_api" and "data.sec.gov/api/xbrl/companyfacts" in final_url:
                    slim_path = file_path.with_name(f"{file_path.stem}_slim.json")
                    slim_path.write_text(
                        json.dumps(slim_sec_company_facts(payload), ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                    source_entry["slim_extract_file"] = str(slim_path.relative_to(batch_dir))
                    source_entry["ai_studio_upload_file"] = str(slim_path.relative_to(batch_dir))
                    source_entry["ai_studio_upload_note"] = (
                        "Upload this compact SEC extract instead of the full Company Facts JSON."
                    )
                time.sleep(0.5)
            except Exception as exc:  # noqa: BLE001 - record for user/AI Studio
                source_entry.update({"download_status": "failed", "error": str(exc)})
            company_entry["sources"].append(source_entry)
        batch_manifest["companies"].append(company_entry)

    write_manifest(batch_dir, batch_manifest)

    company_list = "\n".join(
        f"- {company_lookup[key]['display_name']} ({company_lookup[key]['currency']}, {period_hint_for_mode(mode)})"
        for key in batch["companies"]
    )
    metric_list = "\n".join(f"- {metric}" for metric in METRICS)
    prompt = (
        prompt_template.replace("{{COMPANY_LIST}}", company_list)
        .replace("{{METRIC_LIST}}", metric_list)
        .replace("{{PERIOD_MODE}}", MODE_CONFIG[mode]["mode_label"])
        .replace("{{PERIOD_RULE_1}}", MODE_CONFIG[mode]["period_rule_1"])
        .replace("{{PERIOD_RULE_2}}", MODE_CONFIG[mode]["period_rule_2"])
        .replace("{{PERIOD_INSTRUCTIONS}}", MODE_CONFIG[mode]["period_instructions"])
        .replace("{{SCHEMA_JSON}}", schema_text_for_mode(mode))
        .replace("{{SOURCE_MANIFEST}}", json.dumps(compact_source_manifest(batch_manifest), ensure_ascii=False, indent=2))
    )
    (batch_dir / "prompt_for_aistudio.txt").write_text(prompt, encoding="utf-8")
    shutil.copy2(SCHEMA_PATH, batch_dir / "aistudio_latest_quarter_schema.json")
    write_upload_list(batch_dir, batch_manifest)
    batch_manifest["input_estimate"] = batch_input_estimate(batch_dir)
    write_manifest(batch_dir, batch_manifest)
    return batch_manifest


def split_one_batch(batch: dict, max_companies_per_batch: int) -> list[dict]:
    if max_companies_per_batch <= 0 or len(batch["companies"]) <= max_companies_per_batch:
        return [batch]
    split: list[dict] = []
    companies = batch["companies"]
    for index in range(0, len(companies), max_companies_per_batch):
        part = index // max_companies_per_batch + 1
        split.append(
            {
                "batch_id": f"{batch['batch_id']}_part_{part}",
                "title": f"{batch['title']} part {part}",
                "companies": companies[index : index + max_companies_per_batch],
            }
        )
    return split


def split_batches(batches: list[dict], max_companies_per_batch: int) -> list[dict]:
    split: list[dict] = []
    for batch in batches:
        split.extend(split_one_batch(batch, max_companies_per_batch))
    return split


def split_batch_in_half(batch: dict) -> list[dict]:
    companies = batch["companies"]
    split_size = math.ceil(len(companies) / 2)
    return split_one_batch(batch, split_size)


def prepare_batch_with_budget(
    batch: dict,
    company_lookup: dict[str, dict],
    output_dir: Path,
    schema_text: str,
    prompt_template: str,
    mode: str,
    download_mode: str,
    max_estimated_input_tokens: int,
) -> list[dict]:
    batch_manifest = prepare_batch(
        batch,
        company_lookup,
        output_dir,
        schema_text,
        prompt_template,
        mode,
        download_mode,
    )
    estimate = batch_manifest.get("input_estimate", {})
    if (
        max_estimated_input_tokens > 0
        and estimate.get("approx_input_tokens", 0) > max_estimated_input_tokens
        and len(batch["companies"]) > 1
    ):
        shutil.rmtree(output_dir / batch["batch_id"])
        manifests: list[dict] = []
        for smaller_batch in split_batch_in_half(batch):
            manifests.extend(
                prepare_batch_with_budget(
                    smaller_batch,
                    company_lookup,
                    output_dir,
                    schema_text,
                    prompt_template,
                    mode,
                    download_mode,
                    max_estimated_input_tokens,
                )
            )
        return manifests
    return [batch_manifest]


def write_instructions(output_dir: Path, batch_manifests: list[dict], mode: str) -> None:
    lines = [
        f"# LLM {MODE_CONFIG[mode]['mode_label']} Package",
        "",
        "Use this package manually in Qwen Studio by default, or in the provider selected in the dashboard.",
        "",
        "1. Open one batch folder.",
        "2. In the provider web UI, turn off web search / grounding for this extraction run if that option is enabled.",
        "3. Open `FILES_FOR_AI_STUDIO` and upload every file in that folder.",
        "   That folder contains only source attachments for the selected model.",
        "   It already replaces raw `.xlsx` files and oversized SEC JSON files with upload-friendly extracts.",
        "   Do not upload `source_manifest.json`, the schema JSON, or `FILES_TO_UPLOAD.txt`; the prompt already includes the needed structure.",
        "   If some sources are marked `skipped_fast_mode` or `failed`, the prompt still contains their URLs. The model can use them as references, or you can manually download/upload those files later.",
        "4. Copy-paste `prompt_for_aistudio.txt` into the provider web UI.",
        "5. The model must return raw JSON only.",
        "6. Save each answer as `aistudio_json/<batch_folder_name>.json`, for example:",
    ]
    lines.extend(f"   - `aistudio_json/{manifest['batch_id']}.json`" for manifest in batch_manifests)
    lines.extend([
        "7. Run the matching update command from the project folder.",
        "",
        "Do not edit the original workbook. The update script creates a new workbook copy.",
        "",
        "## Batch Download Summary",
        "",
    ])
    for manifest in batch_manifests:
        downloaded = sum(1 for c in manifest["companies"] for s in c["sources"] if s.get("download_status") == "downloaded")
        failed = sum(1 for c in manifest["companies"] for s in c["sources"] if s.get("download_status") == "failed")
        estimate = manifest.get("input_estimate", {})
        approx_tokens = estimate.get("approx_input_tokens")
        estimate_text = f", approx input {approx_tokens:,} tokens" if approx_tokens else ""
        lines.append(f"- `{manifest['batch_id']}`: downloaded {downloaded}, failed {failed}{estimate_text}")
    lines.append("")
    (output_dir / "README_FOR_USER.md").write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=ROOT / "outputs")
    parser.add_argument("--mode", choices=sorted(MODE_CONFIG), default="quarterly")
    parser.add_argument("--workbook", type=Path, default=ROOT / "Бенч финансовой отчетности_мэйджоры.xlsx")
    parser.add_argument("--download-mode", choices=["fast", "full"], default="fast")
    parser.add_argument(
        "--max-companies-per-batch",
        type=int,
        default=DEFAULT_MAX_COMPANIES_PER_BATCH,
        help=(
            "Split LLM batches by company count. Default keeps the current registry at regional "
            "batches of up to 6 companies; use 1 for the older one-company fallback."
        ),
    )
    parser.add_argument(
        "--max-estimated-input-tokens",
        type=int,
        default=DEFAULT_MAX_ESTIMATED_INPUT_TOKENS,
        help="Auto-split a generated batch if prompt plus upload-source text is estimated above this input token budget. Use 0 to disable.",
    )
    args = parser.parse_args()

    registry = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
    schema_text = SCHEMA_PATH.read_text(encoding="utf-8")
    prompt_template = PROMPT_TEMPLATE_PATH.read_text(encoding="utf-8")
    company_lookup = {company["key"]: company for company in registry["companies"]}

    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    output_dir = (args.output_dir / f"{MODE_CONFIG[args.mode]['folder_prefix']}_{run_id}").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "aistudio_json").mkdir(exist_ok=True)

    batches = split_batches(registry["batches"], args.max_companies_per_batch)
    batch_manifests: list[dict] = []
    for batch in batches:
        batch_manifests.extend(
            prepare_batch_with_budget(
                batch,
                company_lookup,
                output_dir,
                schema_text,
                prompt_template,
                args.mode,
                args.download_mode,
                args.max_estimated_input_tokens,
            )
        )
    write_instructions(output_dir, batch_manifests, args.mode)
    (ROOT / "outputs" / f"latest_aistudio_{args.mode}_package_path.txt").write_text(str(output_dir), encoding="utf-8")
    print(output_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
