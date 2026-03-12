#!/usr/bin/env python3
"""
office_server.py - Modular MCP server for Office document processing

Dynamically loads tool modules from the tools/ directory.
Each tool module provides a mixin class with tool_ methods.
"""

import os
import sys

# Force unbuffered stdio — critical for MCP stdio transport, especially
# inside PyInstaller one-file binaries on Windows where buffering causes
# the host to hang waiting for the first response.
os.environ.setdefault("PYTHONUNBUFFERED", "1")
from typing import Any

# Import the async MCP base class
from aioumcp import AsyncMCPServer

# Dynamically import tool classes from the tools package
from tools import TOOL_CLASSES

# Author name for track changes (from environment or default)
TRACK_CHANGES_AUTHOR = os.environ.get("MCP_AUTHOR", "Solution Architect Agent")

# Tools that have been consolidated into unified tools and should not be exposed
# These implementations still exist for internal use by office_unified_tools
# NOTE: Some tools remain exposed for parity with C# MCP server (excel_list_sheets, pptx_list_slides, pptx_get_notes)
DEPRECATED_TOOLS = {
    # Excel tools → office_read, office_inspect, office_patch, office_table, office_template, office_audit, office_comment
    # NOTE: excel_list_sheets is exposed as basic tool for C# parity
    "excel_extract", "excel_to_markdown",
    "excel_list_named_ranges", "excel_list_tables", "excel_list_merged_cells",
    "excel_get_range", "excel_get_table", "excel_get_comments", "excel_get_change_log",
    "excel_patch_cell", "excel_patch_range", "excel_replace_placeholders",
    "excel_append_table_row", "excel_update_table_row",
    "excel_copy_template", "excel_audit_placeholders",
    "excel_add_comment",
    # Word tools → office_read, office_inspect, office_patch, office_table, office_template, office_audit, office_comment
    "word_extract", "word_to_markdown",
    "word_list_sections", "word_list_tables", "word_get_section", "word_get_table", "word_check_tracking",
    "word_patch_section", "word_patch_placeholder", "word_fix_split_placeholders", "word_replace_global_variables",
    "word_insert_table_row", "word_patch_table_row", "word_create_new_table", "word_duplicate_table_structure",
    "word_copy_template", "word_analyze_template_formatting",
    "word_audit_sow", "word_audit_completion",
    "word_add_comment",
    # PowerPoint tools → office_read, office_inspect, office_patch, office_table, office_template, office_audit, office_comment
    # NOTE: pptx_list_slides, pptx_get_notes exposed as basic tools for C# parity
    "pptx_extract", "pptx_to_markdown",
    "pptx_list_masters", "pptx_list_shapes", "pptx_get_slide",
    "pptx_get_table", "pptx_get_hidden_slides", "pptx_get_comments",
    "pptx_patch_shape", "pptx_replace_text", "pptx_replace_placeholders",
    "pptx_set_text_autofit", "pptx_clear_bullets", "pptx_patch_table_cell",
    "pptx_add_bullet", "pptx_insert_table_row",
    "pptx_copy_template", "pptx_analyze_layouts",
    "pptx_audit_placeholders",
    "pptx_add_comment",
}


def create_server_class(base_class: type, tool_classes: list[type]) -> type:
    """Dynamically create a server class that inherits from all tool mixins."""

    class DynamicOfficeServer(base_class, *tool_classes):
        """MCP server for Office document processing with dynamically loaded tools."""

        def __init__(self):
            super().__init__()
            # Runtime defaults used by comment tools when author is omitted.
            self._comment_author = TRACK_CHANGES_AUTHOR
            self._comment_identity = os.environ.get("MCP_AUTHOR_IDENTITY")
            initials = os.environ.get("MCP_AUTHOR_INITIALS", "").strip().upper()
            if not initials:
                tokens = [part for part in str(self._comment_author).split() if part]
                initials = "".join(part[0].upper() for part in tokens[:2]) if tokens else "SA"
            self._comment_initials = initials

        def get_instructions(self) -> str:
            # Build instructions from loaded tools, excluding deprecated ones
            tool_names = [
                name[5:] for name in dir(self)
                if name.startswith("tool_")
                and callable(getattr(self, name))
                and name[5:] not in DEPRECATED_TOOLS
            ]
            return (
                "Document processing server for Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files. "
                f"Available tools: {', '.join(sorted(tool_names))}. "
                "Can extract text/content, parse structure, and generate new documents."
            )

        def discover_tools(self):
            """Override to filter out deprecated tools from MCP exposure."""
            result = super().discover_tools()
            # Filter out deprecated tools
            result["tools"] = [
                tool for tool in result["tools"]
                if tool["name"] not in DEPRECATED_TOOLS
            ]
            return result

        def tool_list_supported_formats(self) -> dict[str, Any]:
            """List supported document formats and their availability.

            Returns:
                Dictionary showing which formats are available
            """
            has_unified = hasattr(self, "tool_office_read")

            return {
                "unified_tools": {
                    "available": has_unified,
                    "tools": [
                        "office_read", "office_inspect", "office_patch",
                        "office_comment", "office_set_comment_identity",
                        "office_table", "office_template", "office_audit"
                    ] if has_unified else []
                },
                "word": {
                    "extensions": [".docx"],
                    "available": has_unified,
                    "specialized": [
                        "word_from_markdown", "word_generate_sow", "word_cleanup_sow",
                        "word_create_sow_from_markdown", "word_enable_track_changes",
                        "word_extract_sow_structure", "word_get_comments",
                        "word_get_section_guidance", "word_parse_sow_template",
                        "word_patch_with_track_changes",
                    ]
                },
                "excel": {
                    "extensions": [".xlsx", ".xlsm"],
                    "available": has_unified,
                    "specialized": [
                        "excel_from_markdown", "excel_add_sheet", "excel_list_sheets",
                    ]
                },
                "powerpoint": {
                    "extensions": [".pptx"],
                    "available": has_unified,
                    "specialized": [
                        "pptx_from_markdown", "pptx_add_slide", "pptx_delete_slide",
                        "pptx_duplicate_slide", "pptx_get_notes", "pptx_hide_slide",
                        "pptx_add_table",
                        "pptx_list_slides", "pptx_log_changes", "pptx_recommend_layout",
                        "pptx_reorder_slides", "pptx_set_notes",
                    ]
                }
            }

        def tool_restart_server(self) -> dict[str, Any]:
            """Restart the MCP server to reload code changes.

            Use this tool after modifying tool modules in .github/mcp/tools/
            to pick up the changes without manually restarting.

            The server will exit and VS Code will automatically restart it,
            picking up any code changes in the tools/ directory.

            Example:
                restart_server()

            Returns:
                Status message (the server exits immediately after responding)
            """
            import logging
            import platform
            import threading

            logger = logging.getLogger(__name__)
            logger.info("Server restart requested")

            def do_restart():
                """Execute restart after a brief delay to allow response to be sent."""
                import time
                time.sleep(0.5)  # Allow response to be sent
                logger.info("Executing server restart")
                # Re-execute the same script with the same arguments
                python = sys.executable
                script = os.path.abspath(__file__)
                args = [python, script] + sys.argv[1:]
                try:
                    # On Windows, os.execv doesn't work the same way
                    # Use sys.exit and let VS Code restart the server
                    if platform.system() == "Windows":
                        logger.info("Windows detected, using sys.exit for restart")
                        sys.exit(0)
                    else:
                        # On Unix-like systems, execv replaces the process
                        os.execv(python, args)
                except Exception as e:
                    logger.error(f"Restart failed: {e}, falling back to sys.exit")
                    sys.exit(0)

            # Start restart in background thread so response can be sent first
            restart_thread = threading.Thread(target=do_restart, daemon=False)
            restart_thread.start()

            return {
                "success": True,
                "message": "Server restarting to reload code changes...",
                "action": "The server will restart in ~0.5 seconds. Tools will be unavailable briefly."
            }

    return DynamicOfficeServer


# Create the server class with all discovered tools
OfficeServer = create_server_class(AsyncMCPServer, TOOL_CLASSES)


def main():
    """CLI entry point for the MCP server."""
    # Required for PyInstaller one-file builds on Windows: prevents the
    # bootloader from spawning duplicate child processes.
    import multiprocessing
    multiprocessing.freeze_support()

    server = OfficeServer()
    server.run()


if __name__ == "__main__":
    main()
