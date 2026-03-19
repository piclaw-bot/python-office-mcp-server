"""Tests for MCP tool schema validation and strict parameter handling."""

import asyncio

from office_server import OfficeServer


def _get_tool_schema(tools, name):
    for tool in tools:
        if tool.get("name") == name:
            return tool.get("inputSchema", {})
    return {}


def test_tool_schema_disallows_unknown_properties():
    server = OfficeServer()
    tools = server.discover_tools().get("tools", [])
    schema = _get_tool_schema(tools, "office_comment")
    assert schema.get("additionalProperties") is False


def test_rejects_unknown_parameters():
    server = OfficeServer()
    response = asyncio.run(
        server.handle_tools_call_async(
            request_id=1,
            params={
                "name": "office_comment",
                "arguments": {
                    "file_path": "dummy.pptx",
                    "comment": "invalid",
                },
            },
        )
    )

    error = response.get("error", {})
    assert "Unrecognized parameter" in error.get("message", "")


def test_markdown_tools_expose_markdown_file_and_oneof():
    server = OfficeServer()
    tools = server.discover_tools().get("tools", [])

    for tool_name in [
        "word_from_markdown",
        "excel_from_markdown",
        "pptx_from_markdown",
        "word_create_sow_from_markdown",
    ]:
        schema = _get_tool_schema(tools, tool_name)
        props = schema.get("properties", {})

        assert "markdown" in props
        assert "markdown_file" in props

        # Agent-oriented contract: at least one markdown source must be provided
        assert "oneOf" in schema
        assert {"required": ["markdown"]} in schema["oneOf"]
        assert {"required": ["markdown_file"]} in schema["oneOf"]


def test_word_create_sow_schema_marks_template_required():
    server = OfficeServer()
    tools = server.discover_tools().get("tools", [])
    schema = _get_tool_schema(tools, "word_create_sow_from_markdown")

    required = schema.get("required", [])
    assert "output_path" in required
    assert "template_path" in required
