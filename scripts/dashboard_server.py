#!/usr/bin/env python3
"""Local web dashboard for the LLM workbook workflow."""

from __future__ import annotations

import argparse
import ast
import errno
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import urllib.parse
import webbrowser
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
WEB_DIR = ROOT / "web"
PROVIDER_PROFILES_PATH = ROOT / "config" / "ai_provider_profiles.json"

MODE_CONFIG = {
    "quarterly": {
        "label": "Квартал",
        "prepare_args": ["scripts/prepare_aistudio_sources.py", "--mode", "quarterly", "--download-mode", "full"],
        "apply_args": ["scripts/apply_aistudio_json.py", "--mode", "quarterly"],
        "symlink": ROOT / "OPEN_THIS_QUARTERLY_PACKAGE",
        "latest_file": ROOT / "outputs" / "latest_aistudio_quarterly_package_path.txt",
    },
    "annual": {
        "label": "Год",
        "prepare_args": ["scripts/prepare_aistudio_sources.py", "--mode", "annual", "--download-mode", "full"],
        "apply_args": ["scripts/apply_aistudio_json.py", "--mode", "annual"],
        "symlink": ROOT / "OPEN_THIS_ANNUAL_PACKAGE",
        "latest_file": ROOT / "outputs" / "latest_aistudio_annual_package_path.txt",
    },
}

PREPARE_LOCK = threading.RLock()
PREPARE_JOBS: dict[str, dict[str, Any]] = {}


def load_provider_profiles() -> dict[str, dict[str, Any]]:
    return json.loads(PROVIDER_PROFILES_PATH.read_text(encoding="utf-8"))


PROVIDER_PROFILES = load_provider_profiles()


def normalize_provider(provider: str | None) -> str:
    return provider if provider in PROVIDER_PROFILES else "qwen"


def latest_file_for_provider(mode: str, provider: str) -> Path:
    return ROOT / "outputs" / f"latest_aistudio_{mode}_{normalize_provider(provider)}_package_path.txt"


def read_path_pointer(path_file: Path) -> Path | None:
    if not path_file.exists():
        return None
    text = path_file.read_text(encoding="utf-8").strip()
    return Path(text).resolve() if text else None


def package_path(mode: str, provider: str | None = None) -> Path | None:
    config = MODE_CONFIG[mode]
    if provider:
        provider_path = read_path_pointer(latest_file_for_provider(mode, provider))
        if provider_path:
            return provider_path
    package_link = config["symlink"]
    if package_link.exists() and package_link.is_dir():
        return package_link.resolve()
    return read_path_pointer(config["latest_file"])


def update_package_pointer(mode: str, target: Path) -> None:
    """Create the convenient package pointer when the OS allows it.

    macOS accepts this as a normal symlink. Windows often requires Developer
    Mode or admin rights for directory symlinks, so failure is non-fatal; the
    canonical latest-package text file written by the preparer remains the
    cross-platform source of truth.
    """
    package_link = MODE_CONFIG[mode]["symlink"]
    pointer_file = package_link.with_suffix(".txt")
    if package_link.is_symlink() or package_link.is_file():
        package_link.unlink()
    if not package_link.exists():
        try:
            package_link.symlink_to(target, target_is_directory=True)
        except OSError:
            pass
    pointer_file.write_text(str(target), encoding="utf-8")


def copy_text_to_clipboard(text: str) -> None:
    if sys.platform == "darwin":
        subprocess.run(["pbcopy"], input=text, text=True, check=True, timeout=5)
        return
    if os.name == "nt":
        powershell = shutil.which("powershell") or shutil.which("pwsh")
        if powershell:
            command = (
                "[Console]::InputEncoding=[System.Text.UTF8Encoding]::new(); "
                "Set-Clipboard -Value ([Console]::In.ReadToEnd())"
            )
            subprocess.run([powershell, "-NoProfile", "-Command", command], input=text.encode("utf-8"), check=True, timeout=10)
            return
        clip = shutil.which("clip")
        if clip:
            subprocess.run([clip], input=text, text=True, check=True, timeout=5)
            return
        raise RuntimeError("Windows clipboard tool not found.")
    for command in (["wl-copy"], ["xclip", "-selection", "clipboard"]):
        executable = shutil.which(command[0])
        if executable:
            subprocess.run([executable, *command[1:]], input=text, text=True, check=True, timeout=5)
            return
    raise RuntimeError("Clipboard copy is not available on this platform.")


def open_path(path: Path) -> None:
    if sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    elif os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        opener = shutil.which("xdg-open")
        if opener:
            subprocess.run([opener, str(path)], check=False)


def reveal_path(path: Path) -> None:
    target = path if path.exists() else path.parent
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", str(target)], check=False)
    elif os.name == "nt":
        if target.is_dir():
            os.startfile(str(target))  # type: ignore[attr-defined]
        else:
            subprocess.run(["explorer", f"/select,{target}"], check=False)
    else:
        open_path(target if target.is_dir() else target.parent)


def run_project_command(args: list[str], timeout: int = 420) -> dict[str, Any]:
    proc = subprocess.run(
        [sys.executable, *args],
        cwd=ROOT,
        text=True,
        capture_output=True,
        timeout=timeout,
    )
    return {
        "ok": proc.returncode == 0,
        "returncode": proc.returncode,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
    }


def iso_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def prepare_job_key(mode: str, provider: str) -> str:
    return f"{mode}:{normalize_provider(provider)}"


def public_prepare_job(mode: str, provider: str) -> dict[str, Any] | None:
    with PREPARE_LOCK:
        job = PREPARE_JOBS.get(prepare_job_key(mode, provider))
        if not job:
            return None
        return {
            key: value
            for key, value in job.items()
            if key
            in {
                "mode",
                "state",
                "started_at",
                "finished_at",
                "message",
                "package_path",
                "returncode",
                "stderr_tail",
            }
        }


def prepare_worker(mode: str, provider: str) -> None:
    key = prepare_job_key(mode, provider)
    try:
        result = run_project_command([*MODE_CONFIG[mode]["prepare_args"], "--provider", normalize_provider(provider)])
        if not result["ok"]:
            with PREPARE_LOCK:
                PREPARE_JOBS[key].update(
                    {
                        "state": "error",
                        "finished_at": iso_now(),
                        "message": "Подготовка источников завершилась ошибкой.",
                        "returncode": result.get("returncode"),
                        "stderr_tail": (result.get("stderr") or "")[-4000:],
                    }
                )
            return

        lines = [line.strip() for line in result["stdout"].splitlines() if line.strip()]
        if not lines:
            raise ValueError("Prepare script did not print package path.")
        path = Path(lines[-1]).resolve()
        update_package_pointer(mode, path)
        with PREPARE_LOCK:
            PREPARE_JOBS[key].update(
                {
                    "state": "complete",
                    "finished_at": iso_now(),
                    "message": "Источники подготовлены.",
                    "package_path": str(path),
                    "returncode": result.get("returncode"),
                }
            )
    except Exception as exc:  # noqa: BLE001
        with PREPARE_LOCK:
            PREPARE_JOBS.setdefault(key, {"mode": mode, "provider": normalize_provider(provider)}).update(
                {
                    "state": "error",
                    "finished_at": iso_now(),
                    "message": str(exc),
                }
            )


def start_prepare_job(mode: str, provider: str) -> dict[str, Any]:
    provider = normalize_provider(provider)
    key = prepare_job_key(mode, provider)
    with PREPARE_LOCK:
        current = PREPARE_JOBS.get(key)
        if current and current.get("state") == "running":
            return public_prepare_job(mode, provider) or current
        PREPARE_JOBS[key] = {
            "mode": mode,
            "provider": provider,
            "state": "running",
            "started_at": iso_now(),
            "message": "Готовлю пакет источников. Full-download обычно занимает 1-3 минуты.",
        }
        thread = threading.Thread(target=prepare_worker, args=(mode, provider), daemon=True)
        thread.start()
        return public_prepare_job(mode, provider) or PREPARE_JOBS[key]


def safe_resolve(path_value: str | Path) -> Path:
    path = Path(path_value)
    if not path.is_absolute():
        path = ROOT / path
    resolved = path.resolve()
    root = ROOT.resolve()
    if resolved != root and root not in resolved.parents:
        raise ValueError("Path is outside the project folder.")
    return resolved


def read_json(path: Path) -> dict[str, Any] | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def summarize_payload(payload: dict[str, Any] | None) -> dict[str, Any] | None:
    if not payload:
        return None
    companies = payload.get("companies", [])
    if not isinstance(companies, list):
        return None
    non_null = 0
    review = 0
    facts = 0
    names = []
    for company in companies:
        names.append(company.get("company_name") or company.get("company_key") or "company")
        for fact in company.get("facts", []):
            facts += 1
            if fact.get("value") is not None:
                non_null += 1
            if fact.get("review_required") or fact.get("value") is None:
                review += 1
    return {
        "companies": names,
        "company_count": len(companies),
        "facts": facts,
        "non_null": non_null,
        "review": review,
    }


def package_metadata(path: Path | None) -> dict[str, Any] | None:
    if not path:
        return None
    return read_json(path / "package_metadata.json")


def provider_upload_limit(provider: str) -> int:
    profile = PROVIDER_PROFILES[normalize_provider(provider)]
    return int(profile.get("max_upload_files_per_batch") or 0)


def batch_summary(mode: str, batch_dir: Path, provider: str = "qwen") -> dict[str, Any]:
    manifest = read_json(batch_dir / "source_manifest.json") or {}
    upload_file = batch_dir / "FILES_TO_UPLOAD.txt"
    upload_files = []
    if upload_file.exists():
        upload_files = [
            line.strip()
            for line in upload_file.read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        ]
    prompt_path = batch_dir / "prompt_for_aistudio.txt"
    package = package_path(mode, provider)
    json_path = package / "aistudio_json" / f"{batch_dir.name}.json" if package else batch_dir / f"{batch_dir.name}.json"
    payload = read_json(json_path) if json_path.exists() else None
    companies = [
        item.get("display_name") or item.get("key")
        for item in manifest.get("companies", [])
    ]
    downloaded = sum(
        1
        for company in manifest.get("companies", [])
        for source in company.get("sources", [])
        if source.get("download_status") == "downloaded"
    )
    failed = sum(
        1
        for company in manifest.get("companies", [])
        for source in company.get("sources", [])
        if source.get("download_status") == "failed"
    )
    upload_count = len(upload_files)
    max_upload_files = provider_upload_limit(provider)
    return {
        "id": batch_dir.name,
        "title": manifest.get("title") or batch_dir.name,
        "companies": companies,
        "downloaded_sources": downloaded,
        "failed_sources": failed,
        "upload_count": upload_count,
        "provider_limit": max_upload_files,
        "provider_limit_exceeded": bool(max_upload_files and upload_count > max_upload_files),
        "json_saved": json_path.exists(),
        "json_path": str(json_path) if json_path.exists() else None,
        "json_summary": summarize_payload(payload),
        "prompt_chars": prompt_path.stat().st_size if prompt_path.exists() else 0,
    }


def upload_files_from_batch(batch_dir: Path) -> list[Path]:
    upload_file = batch_dir / "FILES_TO_UPLOAD.txt"
    if not upload_file.exists():
        return []
    files = []
    for line in upload_file.read_text(encoding="utf-8").splitlines():
        rel_path = line.strip()
        if not rel_path or rel_path.startswith("#"):
            continue
        source_path = (batch_dir / rel_path).resolve()
        if source_path.is_file() and batch_dir.resolve() in source_path.parents:
            files.append(source_path)
    return files


def slugify_file_part(text: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", text).strip("_") or "source"


def ensure_upload_folder(batch_dir: Path) -> Path:
    upload_dir = batch_dir / "FILES_FOR_AI_STUDIO"
    if upload_dir.exists():
        shutil.rmtree(upload_dir)
    upload_dir.mkdir()
    used_names: set[str] = set()
    for source_path in upload_files_from_batch(batch_dir):
        target_name = source_path.name
        if target_name in used_names:
            target_name = f"{slugify_file_part(source_path.parent.name)}_{target_name}"
        used_names.add(target_name)
        shutil.copy2(source_path, upload_dir / target_name)
    return upload_dir


def status_payload(provider: str = "qwen") -> dict[str, Any]:
    provider = normalize_provider(provider)
    modes = {}
    for mode in MODE_CONFIG:
        path = package_path(mode, provider)
        metadata = package_metadata(path)
        package_provider = (metadata or {}).get("provider") or {}
        batches = []
        json_count = 0
        if path and path.exists():
            batches = [
                batch_summary(mode, batch_dir, provider)
                for batch_dir in sorted(path.glob("batch_*"))
                if batch_dir.is_dir()
            ]
            json_dir = path / "aistudio_json"
            json_count = len(list(json_dir.glob("*.json"))) if json_dir.exists() else 0
        modes[mode] = {
            "label": MODE_CONFIG[mode]["label"],
            "package_path": str(path) if path else None,
            "package_exists": bool(path and path.exists()),
            "batch_count": len(batches),
            "json_count": json_count,
            "batches": batches,
            "prepare_job": public_prepare_job(mode, provider),
            "package_provider": package_provider,
            "provider_mismatch": bool(package_provider and package_provider.get("id") != provider),
        }
    return {"root": str(ROOT), "provider": provider, "providers": PROVIDER_PROFILES, "modes": modes}


def clean_aistudio_text(text: str) -> str:
    return (
        str(text or "")
        .replace("\ufeff", "")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .strip()
    )


def parse_json_candidate(candidate: str) -> Any:
    parse_errors: list[Exception] = []
    for variant in relaxed_json_candidates(candidate):
        try:
            parsed = json.loads(variant)
            break
        except json.JSONDecodeError as exc:
            parse_errors.append(exc)
    else:
        for variant in relaxed_json_candidates(candidate):
            try:
                parsed = ast.literal_eval(variant)
                break
            except (SyntaxError, ValueError) as exc:
                parse_errors.append(exc)
        else:
            raise parse_errors[-1] if parse_errors else ValueError("JSON candidate is empty.")
    if isinstance(parsed, str):
        nested = clean_aistudio_text(parsed)
        if nested and nested != candidate and any(marker in nested for marker in ("{", "[")):
            return extract_json_payload(nested)
    return parsed


def relaxed_json_candidates(candidate: str) -> list[str]:
    cleaned = clean_aistudio_text(candidate)
    variants = [cleaned]
    relaxed = (
        cleaned.replace("“", '"')
        .replace("”", '"')
        .replace("„", '"')
        .replace("‟", '"')
        .replace("‘", "'")
        .replace("’", "'")
    )
    relaxed = re.sub(r",(\s*[}\]])", r"\1", relaxed)
    relaxed = re.sub(r"([{,]\s*)([A-Za-z_][A-Za-z0-9_]*)(\s*:)", r'\1"\2"\3', relaxed)
    relaxed = re.sub(r"\bNone\b", "null", relaxed)
    relaxed = re.sub(r"\bTrue\b", "true", relaxed)
    relaxed = re.sub(r"\bFalse\b", "false", relaxed)
    if relaxed != cleaned:
        variants.append(relaxed)
    return variants


def iter_fenced_json(text: str) -> list[str]:
    return [
        match.group(1).strip()
        for match in re.finditer(r"```(?:json|JSON)?\s*([\s\S]*?)```", text, flags=re.IGNORECASE)
    ]


def iter_balanced_json(text: str) -> list[str]:
    candidates = []
    pairs = {"{": "}", "[": "]"}
    openings = set(pairs)
    closings = set(pairs.values())
    for start, opening in enumerate(text):
        if opening not in openings:
            continue
        stack = [pairs[opening]]
        in_string = False
        escape = False
        for index in range(start + 1, len(text)):
            char = text[index]
            if in_string:
                if escape:
                    escape = False
                elif char == "\\":
                    escape = True
                elif char == '"':
                    in_string = False
                continue
            if char == '"':
                in_string = True
            elif char in openings:
                stack.append(pairs[char])
            elif char in closings:
                if not stack or char != stack[-1]:
                    break
                stack.pop()
                if not stack:
                    candidates.append(text[start : index + 1])
                    break
    return candidates


def normalize_aistudio_payload(payload: Any) -> dict[str, Any]:
    if isinstance(payload, dict) and isinstance(payload.get("companies"), list):
        return payload
    if isinstance(payload, list):
        return {"companies": payload}
    if isinstance(payload, dict) and payload.get("company_key") and isinstance(payload.get("facts"), list):
        return {"companies": [payload]}

    if isinstance(payload, dict):
        for key in ("json", "output", "response", "text", "content"):
            nested = payload.get(key)
            if isinstance(nested, str) and any(marker in nested for marker in ("{", "[")):
                return normalize_aistudio_payload(extract_json_payload(nested))

    raise ValueError(
        "JSON найден, но в нём нет `companies`. Вставь ответ модели целиком: объект с `companies`, массив компаний или один объект компании."
    )


def normalize_identity(value: Any) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").strip().lower().split())


def company_identity_values(company: dict[str, Any]) -> set[str]:
    values = {
        normalize_identity(company.get("key")),
        normalize_identity(company.get("display_name")),
        normalize_identity(company.get("company_key")),
        normalize_identity(company.get("company_name")),
    }
    return {value for value in values if value}


def validate_payload_matches_batch(batch_dir: Path, payload: dict[str, Any]) -> None:
    manifest = read_json(batch_dir / "source_manifest.json") or {}
    expected = set()
    for company in manifest.get("companies", []):
        expected.update(company_identity_values(company))

    actual = set()
    for company in payload.get("companies", []):
        if isinstance(company, dict):
            actual.update(company_identity_values(company))

    if expected and actual and expected.isdisjoint(actual):
        expected_text = ", ".join(sorted(expected))
        actual_text = ", ".join(sorted(actual))
        raise ValueError(
            f"Этот ответ похож на другой батч. Выбранный батч ждёт: {expected_text}; в ответе найдено: {actual_text}."
        )


def extract_json_payload(text: str) -> Any:
    stripped = clean_aistudio_text(text)
    if not stripped:
        raise ValueError("Вставь ответ модели.")

    candidates = [stripped, *iter_fenced_json(stripped), *iter_balanced_json(stripped)]
    seen: set[str] = set()
    last_error: Exception | None = None
    for candidate in candidates:
        candidate = clean_aistudio_text(candidate)
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            return parse_json_candidate(candidate)
        except (json.JSONDecodeError, ValueError) as exc:
            last_error = exc

    if not any(marker in stripped for marker in ("{", "[")):
        raise ValueError("JSON не найден. Скопируй полный ответ модели, не только пояснение.")
    if last_error:
        raise ValueError(f"JSON найден, но не читается: {last_error}") from last_error
    raise ValueError("JSON обрывается до конца. Скопируй ответ модели целиком.")


def extract_normalized_aistudio_payload(text: str) -> dict[str, Any]:
    stripped = clean_aistudio_text(text)
    if not stripped:
        raise ValueError("Вставь ответ модели.")

    candidates = [stripped, *iter_fenced_json(stripped), *iter_balanced_json(stripped)]
    seen: set[str] = set()
    last_parse_error: Exception | None = None
    last_shape_error: Exception | None = None
    for candidate in candidates:
        candidate = clean_aistudio_text(candidate)
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            parsed = parse_json_candidate(candidate)
        except (json.JSONDecodeError, ValueError, SyntaxError) as exc:
            last_parse_error = exc
            continue
        try:
            return normalize_aistudio_payload(parsed)
        except ValueError as exc:
            last_shape_error = exc

    if last_shape_error:
        raise last_shape_error
    if not any(marker in stripped for marker in ("{", "[")):
        raise ValueError("JSON не найден. Скопируй полный ответ модели, не только пояснение.")
    if last_parse_error:
        raise ValueError(f"JSON найден, но не читается: {last_parse_error}") from last_parse_error
    raise ValueError("JSON обрывается до конца. Скопируй ответ модели целиком.")


def get_batch_dir(mode: str, batch_id: str, provider: str = "qwen") -> Path:
    package = package_path(mode, provider)
    if not package:
        raise ValueError("Package is not prepared yet.")
    batch_dir = (package / batch_id).resolve()
    safe_resolve(batch_dir)
    if not batch_dir.is_dir() or not batch_dir.name.startswith("batch_"):
        raise ValueError("Batch folder not found.")
    return batch_dir


def assert_complete_json_set(mode: str, provider: str = "qwen") -> Path:
    package = package_path(mode, provider)
    if not package or not package.exists():
        raise ValueError("Сначала подготовь пакет источников.")
    batches = [batch_dir for batch_dir in package.glob("batch_*") if batch_dir.is_dir()]
    json_dir = package / "aistudio_json"
    json_files = list(json_dir.glob("*.json")) if json_dir.exists() else []
    if not batches:
        raise ValueError("В пакете нет батчей для обработки.")
    if len(json_files) < len(batches):
        raise ValueError(
            f"Финальный Excel можно собрать только после всех JSON: сохранено {len(json_files)} из {len(batches)}."
        )
    return package


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "CordiantDashboard/1.0"

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A003
        return

    def send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path, content_type: str) -> None:
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        return json.loads(raw or "{}")

    def do_GET(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in {"/", "/index.html"}:
            return self.send_file(WEB_DIR / "dashboard.html", "text/html; charset=utf-8")
        if parsed.path == "/dashboard.css":
            return self.send_file(WEB_DIR / "dashboard.css", "text/css; charset=utf-8")
        if parsed.path == "/dashboard.js":
            return self.send_file(WEB_DIR / "dashboard.js", "application/javascript; charset=utf-8")
        if parsed.path == "/api/status":
            query = urllib.parse.parse_qs(parsed.query)
            provider = normalize_provider(query.get("provider", ["qwen"])[0])
            return self.send_json({"ok": True, **status_payload(provider)})
        if parsed.path == "/api/batch":
            query = urllib.parse.parse_qs(parsed.query)
            try:
                mode = query.get("mode", ["quarterly"])[0]
                provider = normalize_provider(query.get("provider", ["qwen"])[0])
                batch_id = query["batch"][0]
                if mode not in MODE_CONFIG:
                    raise ValueError("Unknown mode.")
                batch_dir = get_batch_dir(mode, batch_id, provider)
                prompt = (batch_dir / "prompt_for_aistudio.txt").read_text(encoding="utf-8")
                upload_folder = ensure_upload_folder(batch_dir)
                return self.send_json(
                    {
                        "ok": True,
                        "batch": batch_summary(mode, batch_dir, provider),
                        "prompt": prompt,
                        "batch_path": str(batch_dir),
                        "upload_folder_path": str(upload_folder),
                    }
                )
            except Exception as exc:  # noqa: BLE001
                return self.send_json({"ok": False, "error": str(exc)}, 400)
        self.send_error(404)

    def do_POST(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        try:
            payload = self.read_body()
            if parsed.path == "/api/status":
                provider = normalize_provider(payload.get("provider", "qwen"))
                return self.send_json({"ok": True, **status_payload(provider)})

            if parsed.path == "/api/batch":
                mode = payload.get("mode", "quarterly")
                provider = normalize_provider(payload.get("provider", "qwen"))
                batch_id = payload.get("batch")
                if mode not in MODE_CONFIG or not batch_id:
                    raise ValueError("Mode and batch are required.")
                batch_dir = get_batch_dir(mode, batch_id, provider)
                prompt = (batch_dir / "prompt_for_aistudio.txt").read_text(encoding="utf-8")
                upload_folder = ensure_upload_folder(batch_dir)
                return self.send_json(
                    {
                        "ok": True,
                        "batch": batch_summary(mode, batch_dir, provider),
                        "prompt": prompt,
                        "batch_path": str(batch_dir),
                        "upload_folder_path": str(upload_folder),
                    }
                )

            if parsed.path == "/api/prepare":
                mode = payload.get("mode", "quarterly")
                provider = normalize_provider(payload.get("provider", "qwen"))
                if mode not in MODE_CONFIG:
                    raise ValueError("Unknown mode.")
                return self.send_json(
                    {
                        "ok": True,
                        "prepare_job": start_prepare_job(mode, provider),
                        **status_payload(provider),
                    }
                )

            if parsed.path == "/api/save-json":
                mode = payload.get("mode", "quarterly")
                provider = normalize_provider(payload.get("provider", "qwen"))
                batch_id = payload.get("batch")
                text = payload.get("text", "")
                if mode not in MODE_CONFIG or not batch_id:
                    raise ValueError("Mode and batch are required.")
                batch_dir = get_batch_dir(mode, batch_id, provider)
                parsed_json = extract_normalized_aistudio_payload(text)
                validate_payload_matches_batch(batch_dir, parsed_json)
                package = package_path(mode, provider)
                assert package is not None
                json_dir = package / "aistudio_json"
                json_dir.mkdir(exist_ok=True)
                json_path = json_dir / f"{batch_dir.name}.json"
                json_path.write_text(json.dumps(parsed_json, ensure_ascii=False, indent=2), encoding="utf-8")
                return self.send_json(
                    {
                        "ok": True,
                        "json_path": str(json_path),
                        "summary": summarize_payload(parsed_json),
                        **status_payload(provider),
                    }
                )

            if parsed.path == "/api/start-batch":
                mode = payload.get("mode", "quarterly")
                provider = normalize_provider(payload.get("provider", "qwen"))
                batch_id = payload.get("batch")
                if mode not in MODE_CONFIG or not batch_id:
                    raise ValueError("Mode and batch are required.")
                batch_dir = get_batch_dir(mode, batch_id, provider)
                prompt = (batch_dir / "prompt_for_aistudio.txt").read_text(encoding="utf-8")
                upload_folder = ensure_upload_folder(batch_dir)
                copy_text_to_clipboard(prompt)
                open_path(upload_folder)
                return self.send_json(
                    {
                        "ok": True,
                        "batch": batch_summary(mode, batch_dir, provider),
                        "prompt_chars": len(prompt),
                        "upload_folder_path": str(upload_folder),
                    }
                )

            if parsed.path == "/api/copy-prompt":
                mode = payload.get("mode", "quarterly")
                provider = normalize_provider(payload.get("provider", "qwen"))
                batch_id = payload.get("batch")
                if mode not in MODE_CONFIG or not batch_id:
                    raise ValueError("Mode and batch are required.")
                batch_dir = get_batch_dir(mode, batch_id, provider)
                prompt = (batch_dir / "prompt_for_aistudio.txt").read_text(encoding="utf-8")
                copy_text_to_clipboard(prompt)
                return self.send_json(
                    {
                        "ok": True,
                        "batch": batch_summary(mode, batch_dir, provider),
                        "prompt_chars": len(prompt),
                    }
                )

            if parsed.path == "/api/apply":
                mode = payload.get("mode", "quarterly")
                provider = normalize_provider(payload.get("provider", "qwen"))
                if mode not in MODE_CONFIG:
                    raise ValueError("Unknown mode.")
                package = assert_complete_json_set(mode, provider)
                result = run_project_command(
                    [
                        *MODE_CONFIG[mode]["apply_args"],
                        "--json-dir",
                        str(package / "aistudio_json"),
                    ]
                )
                if not result["ok"]:
                    return self.send_json({"ok": False, **result}, 500)
                output_path = None
                for line in result["stdout"].splitlines():
                    if line.strip().endswith(".xlsx"):
                        candidate = Path(line.strip())
                        output_path = str(candidate.resolve() if candidate.is_absolute() else (ROOT / candidate).resolve())
                return self.send_json({"ok": True, "output_path": output_path, **result})

            if parsed.path == "/api/reveal":
                path = safe_resolve(payload.get("path", ROOT))
                reveal_path(path)
                return self.send_json({"ok": True})

            self.send_error(404)
        except Exception as exc:  # noqa: BLE001
            self.send_json({"ok": False, "error": str(exc)}, 400)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8787)
    parser.add_argument("--open", action="store_true")
    args = parser.parse_args()

    server = None
    selected_port = args.port
    for port in range(args.port, args.port + 20):
        try:
            server = ThreadingHTTPServer((args.host, port), DashboardHandler)
            selected_port = port
            break
        except OSError as exc:
            if getattr(exc, "errno", None) not in {errno.EADDRINUSE, 48, 10048}:
                raise
    if server is None:
        raise OSError(f"No free dashboard port found from {args.port} to {args.port + 19}.")

    url = f"http://{args.host}:{selected_port}/"
    print(url, flush=True)
    if args.open:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
