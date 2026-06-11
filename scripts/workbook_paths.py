#!/usr/bin/env python3
"""Resolve the project source workbook across Unicode filename variants."""

from __future__ import annotations

import unicodedata
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CANONICAL_WORKBOOK_NAME = "Бенч финансовой отчетности_мэйджоры.xlsx"
DEFAULT_WORKBOOK = ROOT / CANONICAL_WORKBOOK_NAME


def filename_key(name: str) -> str:
    return unicodedata.normalize("NFC", name).casefold()


def path_name_variants(path: Path) -> list[Path]:
    names = [
        path.name,
        unicodedata.normalize("NFC", path.name),
        unicodedata.normalize("NFD", path.name),
    ]
    variants: list[Path] = []
    seen: set[str] = set()
    for name in names:
        if name in seen:
            continue
        seen.add(name)
        variants.append(path.with_name(name))
    return variants


def candidate_base_paths(path: Path) -> list[Path]:
    expanded = path.expanduser()
    bases = [expanded] if expanded.is_absolute() else [Path.cwd() / expanded, ROOT / expanded]
    unique: list[Path] = []
    seen: set[Path] = set()
    for base in bases:
        candidate = base.expanduser()
        if candidate in seen:
            continue
        seen.add(candidate)
        unique.append(candidate)
    return unique


def find_name_normalized_match(path: Path) -> Path | None:
    parent = path.parent
    if not parent.is_dir():
        return None
    target_key = filename_key(path.name)
    for candidate in sorted(parent.glob("*.xlsx")):
        if candidate.is_file() and filename_key(candidate.name) == target_key:
            return candidate
    return None


def available_root_workbooks() -> list[Path]:
    return sorted(path for path in ROOT.glob("*.xlsx") if path.is_file())


def resolve_workbook_path(path: Path = DEFAULT_WORKBOOK) -> Path:
    searched: list[Path] = []
    for base in candidate_base_paths(path):
        for candidate in path_name_variants(base):
            searched.append(candidate)
            if candidate.is_file():
                return candidate

        normalized_match = find_name_normalized_match(base)
        if normalized_match:
            return normalized_match

    available = available_root_workbooks()
    expected = ", ".join(str(item) for item in searched)
    available_text = "\n".join(f"- {item.name}" for item in available) or "- нет .xlsx файлов в корне проекта"
    raise FileNotFoundError(
        "Не найден исходный Excel-шаблон для сборки.\n"
        f"Искали: {expected}\n"
        f"Ожидаемое имя: {CANONICAL_WORKBOOK_NAME}\n"
        f"Доступные .xlsx в корне проекта:\n{available_text}\n"
        "Положите исходный workbook в корень проекта или запустите обновление из свежей версии."
    )


def workbook_status(path: Path = DEFAULT_WORKBOOK) -> dict[str, str | bool | None]:
    try:
        resolved = resolve_workbook_path(path)
    except FileNotFoundError as exc:
        return {
            "exists": False,
            "path": None,
            "expected_name": CANONICAL_WORKBOOK_NAME,
            "message": str(exc),
        }

    return {
        "exists": True,
        "path": str(resolved),
        "expected_name": CANONICAL_WORKBOOK_NAME,
        "message": None,
    }
