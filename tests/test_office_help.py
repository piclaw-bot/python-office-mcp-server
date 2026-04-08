"""Tests for structured office_help discovery and recommendation guidance."""

import asyncio
import json

from office_server import OfficeServer
from tools.discovery_tools import DiscoveryTools


class TestOfficeHelp:
    def test_summary_returns_compact_word_sow_guidance(self):
        tool = DiscoveryTools()

        result = tool.tool_office_help(
            goal="fill_sow_from_markdown",
            document_type="word",
            format="summary",
        )

        assert result["success"] is True
        assert result["goal"] == "fill_sow_from_markdown"
        assert result["document_type"] == "word"
        assert result["recommended"][0]["tool"] == "word_create_sow_from_markdown"
        assert "word_generate_sow" in result["fallbacks"]
        assert "unmapped_sections" in result["watch_for"]
        assert result["tool_model"]["primary"] == "core_first"
        assert "office_help" in result["tool_model"]["core_tools"]
        assert result["notes"]
        assert len(result["notes"]) == 1
        assert "workflow_steps" not in result

    def test_detailed_returns_fallbacks_and_workflow_steps(self):
        tool = DiscoveryTools()

        result = tool.tool_office_help(
            goal="patch_estimate_workbook",
            document_type="excel",
            format="detailed",
        )

        assert result["success"] is True
        assert result["goal"] == "patch_estimate_workbook"
        assert result["document_type"] == "excel"
        assert result["recommended"][0]["tool"] == "office_patch"
        assert "package-preservation warnings" in result["watch_for"]
        assert "workflow_steps" in result
        assert len(result["notes"]) >= 2

    def test_task_mapping_resolves_common_consulting_phrase(self):
        tool = DiscoveryTools()

        result = tool.tool_office_help(
            task="Fill a Word SOW template from markdown and preserve template structure",
            format="summary",
        )

        assert result["success"] is True
        assert result["goal"] == "fill_sow_from_markdown"
        assert result["task_interpreted_as"] == "fill_sow_from_markdown"
        assert result["document_type"] == "word"

    def test_additive_narrative_constraint_prioritizes_anchor_insertion(self):
        tool = DiscoveryTools()

        result = tool.tool_office_help(
            goal="insert_architecture_narrative",
            constraints=["additive_narrative_edits"],
            format="summary",
        )

        assert result["success"] is True
        assert result["recommended"][0]["tool"] == "word_insert_at_anchor"

    def test_powerpoint_goal_supported(self):
        tool = DiscoveryTools()

        result = tool.tool_office_help(
            goal="create_review_deck",
            document_type="powerpoint",
            format="summary",
        )

        assert result["success"] is True
        assert result["document_type"] == "powerpoint"
        assert result["recommended"][0]["tool"] == "pptx_from_markdown"
        assert "pptx_add_slide" in result["fallbacks"]

    def test_no_goal_returns_core_overview(self):
        tool = DiscoveryTools()

        result = tool.tool_office_help()

        assert result["success"] is True
        assert "office_help" in result["core_tools"]
        assert "fill_sow_from_markdown" in result["common_goals"]
        assert "fallback" in result["advanced_tool_classes"]
        assert "goal" not in result

    def test_mcp_server_exposes_office_help_and_returns_recommendations(self):
        server = OfficeServer()

        response = asyncio.run(
            server.handle_tools_call_async(
                request_id=1,
                params={
                    "name": "office_help",
                    "arguments": {
                        "goal": "fill_sow_from_markdown",
                        "document_type": "word",
                        "format": "summary",
                    },
                },
            )
        )

        payload = response.get("result", {}).get("content", [])[0]["text"]
        result = json.loads(payload)
        assert result.get("success") is True
        assert result.get("goal") == "fill_sow_from_markdown"
        assert result.get("recommended", [])[0]["tool"] == "word_create_sow_from_markdown"
