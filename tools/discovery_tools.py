"""Structured discovery and recommendation tools for Office MCP workflows."""

from __future__ import annotations

from typing import Any, Literal


CORE_TOOLS = [
    "office_help",
    "office_read",
    "office_inspect",
    "office_patch",
    "office_table",
    "office_template",
    "office_audit",
]


WORKFLOW_GUIDANCE: dict[str, dict[str, Any]] = {
    "fill_sow_from_markdown": {
        "document_type": "word",
        "summary": "Fill a consulting SOW template from markdown, then audit gaps before cleanup.",
        "recommended": [
            {
                "tool": "word_create_sow_from_markdown",
                "why": "Best first-pass workflow for consulting SOW templates driven by markdown.",
            },
            {
                "tool": "word_insert_at_anchor",
                "why": "Use for additive narrative insertion when a section needs more prose without replacing the full body.",
            },
            {
                "tool": "word_audit_completion",
                "why": "Confirms whether required sections, tables, and placeholders were actually completed.",
            },
        ],
        "fallbacks": ["word_generate_sow", "office_patch", "office_audit"],
        "watch_for": [
            "unmapped_sections",
            "section_diagnostics.matched=false",
            "table_diagnostics.matched=false",
        ],
        "next_step": "Run word_create_sow_from_markdown first, review diagnostics, then fill remaining gaps with word_insert_at_anchor or office_patch(section:...).",
        "related_tools": [
            "word_parse_sow_template",
            "word_get_section_guidance",
            "office_template",
        ],
        "notes": [
            "Prefer template-preserving edits over rebuilding a customer-facing SOW from scratch.",
            "Audit completion before removing template guidance or handing the draft to stakeholders.",
        ],
    },
    "insert_architecture_narrative": {
        "document_type": "word",
        "summary": "Add architecture prose into an existing Word deliverable without replacing more content than necessary.",
        "recommended": [
            {
                "tool": "word_insert_at_anchor",
                "why": "Safest path when you know the insertion point and want additive narrative edits.",
            },
            {
                "tool": "office_patch",
                "why": "Use section: targets when you want to replace or seed the body of a named section.",
            },
            {
                "tool": "word_audit_completion",
                "why": "Verifies the document still has complete narrative coverage after patching.",
            },
        ],
        "fallbacks": ["word_get_section_guidance", "office_inspect"],
        "watch_for": [
            "anchor not found",
            "zero modifications",
            "unmapped_sections",
        ],
        "next_step": "Inspect headings or anchors first, then choose word_insert_at_anchor for additive changes or office_patch(section:...) for targeted replacement.",
        "related_tools": ["word_list_anchors", "word_document_map", "word_get_section_guidance", "office_inspect"],
        "notes": [
            "Use additive insertion for customer-ready sections where preserving surrounding content matters.",
        ],
    },
    "find_insertion_point": {
        "document_type": "word",
        "summary": "Find a reliable insertion point before adding narrative content to a Word document.",
        "recommended": [
            {
                "tool": "word_list_anchors",
                "why": "Lists likely heading and paragraph anchors for deterministic insertion workflows.",
            },
            {
                "tool": "office_inspect",
                "why": "Use sections/tables inspection to understand nearby structure before editing.",
            },
            {
                "tool": "word_insert_at_anchor",
                "why": "Apply the insertion once you have chosen a stable anchor.",
            },
        ],
        "fallbacks": ["word_get_section_guidance", "office_read"],
        "watch_for": ["anchor ambiguity", "empty sections", "heading normalization mismatches"],
        "next_step": "Enumerate anchors, choose one near the target section, then insert after the heading or anchor paragraph.",
        "related_tools": ["word_insert_at_anchor", "office_patch"],
        "notes": [
            "Dedicated anchor discovery is preferred over guessing paragraph indexes in complex consulting templates.",
        ],
    },
    "append_table_row": {
        "document_type": "word",
        "summary": "Add a structured row to an existing delivery, staffing, or milestone table.",
        "recommended": [
            {
                "tool": "office_table",
                "why": "Primary unified path for additive row operations across document types.",
            },
            {
                "tool": "office_inspect",
                "why": "Inspect table ids and headers before inserting to avoid schema mismatches.",
            },
            {
                "tool": "office_audit",
                "why": "Use audit or read-back to verify the row landed in the expected table.",
            },
        ],
        "fallbacks": ["word_insert_table_row", "excel_update_table_row"],
        "watch_for": ["header mismatch", "table id mismatch", "merged cell layouts"],
        "next_step": "Inspect the target table first, then add a row using exact header names.",
        "related_tools": ["office_read", "office_template"],
        "notes": [
            "For consulting deliverables, exact column-name alignment matters more than row order.",
        ],
    },
    "patch_estimate_workbook": {
        "document_type": "excel",
        "summary": "Safely update estimate, staffing, or planning values in an existing workbook.",
        "recommended": [
            {
                "tool": "office_patch",
                "why": "Primary unified path for workbook cell/range updates while preserving the original package.",
            },
            {
                "tool": "office_inspect",
                "why": "Inspect sheets, named ranges, or tables before mutating a complex workbook.",
            },
            {
                "tool": "office_audit",
                "why": "Use read-back/audit after patching to verify expected values and preserved structure.",
            },
        ],
        "fallbacks": ["office_read", "excel_list_sheets"],
        "watch_for": [
            "unsupported workbook features",
            "range targeting mistakes",
            "package-preservation warnings",
        ],
        "next_step": "Inspect workbook structure first for non-trivial files, then patch only the intended cells or ranges and verify the saved workbook reopens cleanly.",
        "related_tools": ["office_template", "list_supported_formats"],
        "notes": [
            "Prefer narrow cell/range patches over broad rewrites in customer estimate workbooks.",
            "Current default mutation mode is best_effort; use diagnostics to catch partial success.",
        ],
    },
    "inspect_template_structure": {
        "document_type": "word",
        "summary": "Inspect a template before filling it so you can preserve structure and choose the right mutation path.",
        "recommended": [
            {
                "tool": "office_template",
                "why": "Primary unified entry point for template analysis and copy workflows.",
            },
            {
                "tool": "office_inspect",
                "why": "Lists sections, tables, slides, or sheets needed for targeted edits.",
            },
            {
                "tool": "office_help",
                "why": "Use again with a specific goal after inspection to choose the next mutation workflow.",
            },
        ],
        "fallbacks": ["word_parse_sow_template", "office_read"],
        "watch_for": ["split placeholders", "template instructions", "complex tables"],
        "next_step": "Copy the template, inspect its sections/tables, then choose a generation or patch workflow based on what must be preserved.",
        "related_tools": ["office_patch", "office_table", "word_create_sow_from_markdown"],
        "notes": [
            "This is a good first step for new customer templates or when diagnostics show skipped matches.",
        ],
    },
    "create_review_deck": {
        "document_type": "powerpoint",
        "summary": "Build or update a stakeholder review deck while preserving slide structure and notes.",
        "recommended": [
            {
                "tool": "pptx_from_markdown",
                "why": "Fastest way to create a draft deck from markdown content for architecture reviews.",
            },
            {
                "tool": "office_patch",
                "why": "Use for targeted shape text updates when you must preserve an existing template deck.",
            },
            {
                "tool": "office_inspect",
                "why": "Inspect slides and shapes before patching customer-facing presentations.",
            },
        ],
        "fallbacks": ["pptx_add_slide", "pptx_set_notes", "office_audit"],
        "watch_for": ["shape targeting mistakes", "layout mismatch", "speaker note drift"],
        "next_step": "Choose between generating a draft deck from markdown or patching an existing review deck, then inspect slides before editing speaker notes or tables.",
        "related_tools": ["pptx_recommend_layout", "pptx_list_slides"],
        "notes": [
            "Use existing customer templates when branding/layout must be preserved.",
        ],
    },
    "audit_deliverable_completeness": {
        "document_type": "word",
        "summary": "Check whether a consulting deliverable is complete before review or handoff.",
        "recommended": [
            {
                "tool": "office_audit",
                "why": "Primary cross-format audit path for placeholders, completion, and tracking status.",
            },
            {
                "tool": "office_read",
                "why": "Read back the generated content to confirm critical sections and tables are present.",
            },
            {
                "tool": "office_help",
                "why": "Use again with a recovery goal if the audit reports gaps or skipped mappings.",
            },
        ],
        "fallbacks": ["word_audit_completion", "word_audit_sow"],
        "watch_for": ["remaining placeholders", "low completion score", "skipped table diagnostics"],
        "next_step": "Run an audit after generation or patching, then target unresolved sections, placeholders, or tables explicitly.",
        "related_tools": ["office_patch", "word_insert_at_anchor", "office_table"],
        "notes": [
            "A successful mutation is not the same as a complete deliverable; always audit customer-facing outputs.",
        ],
    },
}


TASK_KEYWORDS: list[tuple[str, str]] = [
    ("fill_sow_from_markdown", "sow markdown template statement of work"),
    ("insert_architecture_narrative", "insert narrative architecture prose section heading"),
    ("find_insertion_point", "anchor insertion point heading paragraph before after"),
    ("append_table_row", "append add row staffing milestone table"),
    ("patch_estimate_workbook", "excel workbook estimate budget cost staffing sheet range cell"),
    ("inspect_template_structure", "inspect analyze template structure preserve sections tables placeholders"),
    ("create_review_deck", "powerpoint ppt slide deck review presentation notes"),
    ("audit_deliverable_completeness", "audit completion placeholders verify review handoff"),
]


GOAL_ALIASES = {
    "fill_template_from_markdown": "fill_sow_from_markdown",
    "generate_statement_of_work": "fill_sow_from_markdown",
    "replace_section_content": "insert_architecture_narrative",
    "safe_patch_complex_workbook": "patch_estimate_workbook",
    "prepare_review_deck": "create_review_deck",
    "verify_delivery_document": "audit_deliverable_completeness",
}


class DiscoveryTools:
    """Structured discovery tools for office workflows."""

    def _canonical_goal(self, goal: str | None, task: str | None) -> str | None:
        if goal:
            normalized = goal.strip().lower().replace("-", "_").replace(" ", "_")
            return GOAL_ALIASES.get(normalized, normalized)

        if not task:
            return None

        normalized_task = task.lower()
        best_goal = None
        best_score = 0
        for workflow_goal, keywords in TASK_KEYWORDS:
            score = sum(1 for keyword in keywords.split() if keyword in normalized_task)
            if score > best_score:
                best_goal = workflow_goal
                best_score = score
        return best_goal

    def _infer_document_type(self, document_type: str | None, goal: str | None, task: str | None) -> str | None:
        if document_type:
            normalized = document_type.strip().lower()
            aliases = {
                "docx": "word",
                "word": "word",
                "excel": "excel",
                "xlsx": "excel",
                "xlsm": "excel",
                "ppt": "powerpoint",
                "pptx": "powerpoint",
                "powerpoint": "powerpoint",
            }
            return aliases.get(normalized, normalized)

        if goal and goal in WORKFLOW_GUIDANCE:
            return WORKFLOW_GUIDANCE[goal].get("document_type")

        task_text = (task or "").lower()
        if any(token in task_text for token in ["ppt", "powerpoint", "slide", "deck"]):
            return "powerpoint"
        if any(token in task_text for token in ["excel", "workbook", "sheet", "cell", "range"]):
            return "excel"
        if any(token in task_text for token in ["word", "docx", "section", "sow", "heading"]):
            return "word"
        return None

    def tool_office_help(
        self,
        goal: str | None = None,
        document_type: Literal["word", "excel", "powerpoint"] | None = None,
        constraints: list[str] | None = None,
        task: str | None = None,
        format: Literal["summary", "detailed"] = "summary",
    ) -> dict[str, Any]:
        """Get structured workflow help and recommendations for office document work.

        Use this as the preferred discovery entry point for systems architecture
        and consulting workflows. Prefer `goal` plus optional `document_type`
        and `constraints`. `task` is supported only as a thin convenience layer
        for mapping common natural-language requests onto the structured workflow
        catalog.
        """
        normalized_constraints = [str(item).strip() for item in (constraints or []) if str(item).strip()]
        canonical_goal = self._canonical_goal(goal, task)
        resolved_document_type = self._infer_document_type(document_type, canonical_goal, task)

        if canonical_goal is None:
            return {
                "success": True,
                "mode": format,
                "scope": "systems_architecture_and_consulting",
                "core_tools": CORE_TOOLS,
                "common_goals": sorted(WORKFLOW_GUIDANCE.keys()),
                "message": "Start with a structured goal such as fill_sow_from_markdown, patch_estimate_workbook, create_review_deck, or audit_deliverable_completeness.",
                "next_step": "Call office_help again with a goal and optional document_type/constraints for deterministic recommendations.",
            }

        guidance = WORKFLOW_GUIDANCE.get(canonical_goal)
        if not guidance:
            return {
                "success": False,
                "error": f"Unknown goal: {canonical_goal}",
                "supported_goals": sorted(WORKFLOW_GUIDANCE.keys()),
                "message": "Use one of the supported consulting workflow goals or provide a clearer task description.",
            }

        recommendations = list(guidance["recommended"])
        if "preserve_template_structure" in normalized_constraints:
            recommendations = sorted(
                recommendations,
                key=lambda item: 0 if item["tool"] in {"office_template", "office_inspect", "word_create_sow_from_markdown", "office_patch"} else 1,
            )
        if "additive_narrative_edits" in normalized_constraints:
            recommendations = sorted(
                recommendations,
                key=lambda item: 0 if item["tool"] == "word_insert_at_anchor" else 1,
            )

        payload: dict[str, Any] = {
            "success": True,
            "scope": "systems_architecture_and_consulting",
            "goal": canonical_goal,
            "document_type": resolved_document_type,
            "constraints": normalized_constraints,
            "summary": guidance["summary"],
            "recommended": recommendations,
            "fallbacks": guidance["fallbacks"],
            "watch_for": guidance["watch_for"],
            "next_step": guidance["next_step"],
            "related_tools": guidance["related_tools"],
            "core_tools": CORE_TOOLS,
        }

        if task:
            payload["task_interpreted_as"] = canonical_goal

        if format == "detailed":
            payload["notes"] = guidance["notes"]
            payload["workflow_steps"] = [
                f"1. Start with {recommendations[0]['tool']}",
                f"2. If needed, continue with {recommendations[1]['tool'] if len(recommendations) > 1 else guidance['fallbacks'][0]}",
                f"3. Validate results with {recommendations[-1]['tool'] if recommendations else guidance['fallbacks'][0]}",
            ]
        else:
            payload["notes"] = guidance["notes"][:1]

        return payload
