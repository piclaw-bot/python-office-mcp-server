"""Shared diagnostics helpers for Office mutation tools."""

from __future__ import annotations

from typing import Any


DEFAULT_SUCCESS_NEXT_TOOLS = ["office_read", "office_inspect", "office_audit"]
DEFAULT_RECOVERY_NEXT_TOOLS = ["office_help", "office_inspect", "office_audit"]


def summarize_status(matched_targets: list[dict[str, Any]], unmatched_targets: list[dict[str, Any]], skipped_targets: list[dict[str, Any]]) -> str:
    if matched_targets and not unmatched_targets and not skipped_targets:
        return "success"
    if matched_targets and (unmatched_targets or skipped_targets):
        return "partial_success"
    if skipped_targets and not matched_targets and not unmatched_targets:
        return "skipped"
    return "failed"


def build_mutation_diagnostics(
    *,
    matched_targets: list[dict[str, Any]] | None = None,
    unmatched_targets: list[dict[str, Any]] | None = None,
    skipped_targets: list[dict[str, Any]] | None = None,
    warnings: list[str] | None = None,
    diagnostics: dict[str, Any] | None = None,
    next_tools: list[str] | None = None,
) -> dict[str, Any]:
    matched_targets = matched_targets or []
    unmatched_targets = unmatched_targets or []
    skipped_targets = skipped_targets or []
    warnings = warnings or []
    diagnostics = diagnostics or {}
    status = summarize_status(matched_targets, unmatched_targets, skipped_targets)
    success = bool(matched_targets) and status in {"success", "partial_success"}

    if next_tools is None:
        next_tools = DEFAULT_SUCCESS_NEXT_TOOLS if success else DEFAULT_RECOVERY_NEXT_TOOLS

    return {
        "success": success,
        "status": status,
        "warnings": warnings,
        "matched_targets": matched_targets,
        "unmatched_targets": unmatched_targets,
        "skipped_targets": skipped_targets,
        "diagnostics": diagnostics,
        "next_tools": next_tools,
    }
