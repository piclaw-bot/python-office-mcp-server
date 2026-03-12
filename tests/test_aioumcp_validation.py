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
