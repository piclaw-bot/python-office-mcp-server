"""Persistent JSON metadata cache for Office template analysis."""

from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

CACHE_SCHEMA_VERSION = 1
ENV_CACHE_DIR = "OFFICE_MCP_METADATA_CACHE_DIR"
DEFAULT_CACHE_DIRNAME = ".office-metadata-cache"


def metadata_cache_dir() -> Path:
    configured = os.environ.get(ENV_CACHE_DIR)
    if configured:
        return Path(configured)
    return Path.cwd() / DEFAULT_CACHE_DIRNAME


def _file_fingerprint(path: str | Path) -> dict[str, Any]:
    resolved = Path(path).resolve()
    stat = resolved.stat()
    return {
        "path": str(resolved),
        "mtime_ns": stat.st_mtime_ns,
        "size": stat.st_size,
    }


def metadata_cache_key(path: str | Path, document_type: str, analysis_type: str) -> str:
    fingerprint = _file_fingerprint(path)
    raw = f"{document_type}|{analysis_type}|{fingerprint['path']}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def metadata_cache_path(path: str | Path, document_type: str, analysis_type: str) -> Path:
    cache_dir = metadata_cache_dir() / document_type / analysis_type
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir / f"{metadata_cache_key(path, document_type, analysis_type)}.json"


def load_cached_metadata(
    path: str | Path,
    document_type: str,
    analysis_type: str,
    *,
    schema_version: int = CACHE_SCHEMA_VERSION,
) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    cache_path = metadata_cache_path(path, document_type, analysis_type)
    fingerprint = _file_fingerprint(path)

    if not cache_path.exists():
        return None, {"hit": False, "reason": "miss", "cache_file": str(cache_path)}

    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001
        return None, {
            "hit": False,
            "reason": "corrupt",
            "cache_file": str(cache_path),
            "error": str(exc),
        }

    if payload.get("schemaVersion") != schema_version:
        return None, {
            "hit": False,
            "reason": "schema_mismatch",
            "cache_file": str(cache_path),
            "cached_schema_version": payload.get("schemaVersion"),
            "expected_schema_version": schema_version,
        }

    cached_source = payload.get("source", {})
    if (
        cached_source.get("path") != fingerprint["path"]
        or cached_source.get("mtime_ns") != fingerprint["mtime_ns"]
        or cached_source.get("size") != fingerprint["size"]
    ):
        return None, {
            "hit": False,
            "reason": "stale",
            "cache_file": str(cache_path),
        }

    return payload.get("metadata"), {
        "hit": True,
        "reason": "hit",
        "cache_file": str(cache_path),
        "generated_at": payload.get("generatedAt"),
    }


def store_cached_metadata(
    path: str | Path,
    document_type: str,
    analysis_type: str,
    metadata: dict[str, Any],
    *,
    schema_version: int = CACHE_SCHEMA_VERSION,
) -> dict[str, Any]:
    cache_path = metadata_cache_path(path, document_type, analysis_type)
    payload = {
        "schemaVersion": schema_version,
        "documentType": document_type,
        "analysisType": analysis_type,
        "source": _file_fingerprint(path),
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "metadata": metadata,
    }
    temp_path = cache_path.with_suffix(cache_path.suffix + ".tmp")
    temp_path.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    temp_path.replace(cache_path)
    return {
        "hit": False,
        "reason": "stored",
        "cache_file": str(cache_path),
        "generated_at": payload["generatedAt"],
    }
