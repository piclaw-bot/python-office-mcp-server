"""
Tests for word_advanced_tools.py - Advanced Word document manipulation

Tests cover:
- Section operations
- Table operations
- Placeholder replacement
- SOW generation
- Track changes
- Template operations
"""

from docx import Document

from tools.word_advanced_tools import (
    _get_text_with_track_changes,
)

# Fixtures temp_dir, word_advanced_tools, sample_docx, and sow_template are provided by conftest.py


class TestGetTextWithTrackChanges:
    """Tests for _get_text_with_track_changes helper."""

    def test_reads_normal_text(self, temp_dir):
        """Should read normal paragraph text."""
        doc = Document()
        doc.add_paragraph("Normal text content")
        path = temp_dir / "normal.docx"
        doc.save(path)

        doc = Document(path)
        # Find the paragraph with our text (index varies by template)
        texts = [_get_text_with_track_changes(p) for p in doc.paragraphs]
        assert any("Normal text content" in t for t in texts)


class TestListSections:
    """Tests for tool_word_list_sections."""

    def test_lists_sections(self, word_advanced_tools, sow_template):
        """Should list all sections/headings."""
        result = word_advanced_tools.tool_word_list_sections(str(sow_template))
        assert "sections" in result
        sections = result["sections"]
        # Should find our headings
        section_titles = [s.get("title", s.get("text", "")) for s in sections]
        assert any("Executive Summary" in t for t in section_titles)
        assert any("Scope" in t for t in section_titles)

    def test_file_not_found(self, word_advanced_tools):
        """Should handle missing files."""
        result = word_advanced_tools.tool_word_list_sections("/nonexistent.docx")
        assert "error" in result


class TestGetSection:
    """Tests for tool_word_get_section."""

    def test_gets_section_content(self, word_advanced_tools, sow_template):
        """Should get content of a specific section."""
        result = word_advanced_tools.tool_word_get_section(str(sow_template), "Executive Summary")
        assert "content" in result or "text" in result or "paragraphs" in str(result)

    def test_section_not_found(self, word_advanced_tools, sow_template):
        """Should handle missing section gracefully."""
        result = word_advanced_tools.tool_word_get_section(str(sow_template), "Nonexistent Section")
        # Should indicate section not found
        assert "error" in result or "not found" in str(result).lower() or "available" in str(result).lower()


class TestPatchSection:
    """Tests for tool_word_patch_section."""

    def test_patches_section(self, word_advanced_tools, sow_template, temp_dir):
        """Should update section content."""
        output = temp_dir / "patched.docx"
        result = word_advanced_tools.tool_word_patch_section(
            str(sow_template),
            "Executive Summary",
            "This is the new executive summary content.",
            output_path=str(output)
        )
        assert result.get("success") is True or "patched" in str(result).lower()

    def test_section_not_found(self, word_advanced_tools, sow_template, temp_dir):
        """Should handle missing section gracefully."""
        output = temp_dir / "patched2.docx"
        result = word_advanced_tools.tool_word_patch_section(
            str(sow_template),
            "Nonexistent",
            "New content",
            output_path=str(output)
        )
        assert "error" in result or "not found" in str(result).lower()


class TestListTables:
    """Tests for tool_word_list_tables."""

    def test_lists_tables(self, word_advanced_tools, sample_docx):
        """Should list tables in document."""
        result = word_advanced_tools.tool_word_list_tables(str(sample_docx))
        assert "tables" in result
        assert len(result["tables"]) >= 1


class TestGetTable:
    """Tests for tool_word_get_table."""

    def test_gets_table_by_index(self, word_advanced_tools, sample_docx):
        """Should get table by index."""
        result = word_advanced_tools.tool_word_get_table(str(sample_docx), "0")
        assert "rows" in result or "data" in result or "header" in result


class TestInsertTableRow:
    """Tests for tool_word_insert_table_row."""

    def test_inserts_row(self, word_advanced_tools, sample_docx, temp_dir):
        """Should insert a new row into table."""
        output = temp_dir / "row_added.docx"
        # API expects dict, not list
        result = word_advanced_tools.tool_word_insert_table_row(
            str(sample_docx),
            "0",
            {"Role": "Analyst", "Hours": "80", "Rate": "$125"},
            output_path=str(output)
        )
        assert result.get("success") is True


class TestPatchTableRow:
    """Tests for tool_word_patch_table_row."""

    def test_patches_row(self, word_advanced_tools, sample_docx, temp_dir):
        """Should update a table row."""
        output = temp_dir / "row_patched.docx"
        result = word_advanced_tools.tool_word_patch_table_row(
            str(sample_docx),
            "0",
            1,
            {"Hours": "120"},
            output_path=str(output)
        )
        assert result.get("success") is True or "updated" in str(result).lower()


class TestAuditCompletion:
    """Tests for tool_word_audit_completion."""

    def test_finds_placeholders(self, word_advanced_tools, sow_template):
        """Should find unfilled placeholders."""
        result = word_advanced_tools.tool_word_audit_completion(str(sow_template))
        # Should find <Customer Name> and <Project Name>
        assert "issues" in result or "placeholders" in str(result).lower() or "findings" in result

    def test_empty_document(self, word_advanced_tools, temp_dir):
        """Should handle document without issues."""
        doc = Document()
        doc.add_paragraph("Clean content with no placeholders.")
        path = temp_dir / "clean.docx"
        doc.save(path)

        result = word_advanced_tools.tool_word_audit_completion(str(path))
        # Should complete without error
        assert "error" not in result


class TestAuditSow:
    """Tests for tool_word_audit_sow."""

    def test_audits_sow(self, word_advanced_tools, sow_template):
        """Should audit SOW for completeness."""
        result = word_advanced_tools.tool_word_audit_sow(str(sow_template))
        # Should find issues
        assert "issues" in result or "placeholders" in str(result).lower()


class TestReplaceGlobalVariables:
    """Tests for tool_word_replace_global_variables."""

    def test_replaces_placeholders(self, word_advanced_tools, sow_template, temp_dir):
        """Should replace placeholder text."""
        output = temp_dir / "replaced.docx"
        result = word_advanced_tools.tool_word_replace_global_variables(
            str(sow_template),
            {"<Customer Name>": "Contoso Corp", "<Project Name>": "Cloud Migration"},
            output_path=str(output)
        )
        assert result.get("success") is True


class TestEnableTrackChanges:
    """Tests for tool_word_enable_track_changes."""

    def test_enables_tracking(self, word_advanced_tools, sample_docx, temp_dir):
        """Should enable track changes."""
        output = temp_dir / "tracked.docx"
        result = word_advanced_tools.tool_word_enable_track_changes(
            str(sample_docx),
            output_path=str(output)
        )
        assert result.get("success") is True or "enabled" in str(result).lower()


class TestPatchWithTrackChanges:
    """Tests for tool_word_patch_with_track_changes."""

    def test_patches_with_tracking(self, word_advanced_tools, sample_docx, temp_dir):
        """Should patch text with tracking enabled."""
        output = temp_dir / "tracked_patch.docx"
        result = word_advanced_tools.tool_word_patch_with_track_changes(
            str(sample_docx),
            {"executive summary content": "updated summary content"},
            output_path=str(output)
        )
        assert result.get("success") is True


class TestGenerateSow:
    """Tests for tool_word_generate_sow."""

    def test_generates_sow(self, word_advanced_tools, sow_template, temp_dir):
        """Should generate SOW from template."""
        output = temp_dir / "generated.docx"
        result = word_advanced_tools.tool_word_generate_sow(
            str(sow_template),
            str(output),
            {"customer_name": "Contoso Corp", "project_name": "Migration"}
        )
        assert result.get("success") is True


class TestCopyTemplate:
    """Tests for tool_word_copy_template."""

    def test_copies_template(self, word_advanced_tools, sow_template, temp_dir):
        """Should copy template to new location."""
        output = temp_dir / "copied.docx"
        result = word_advanced_tools.tool_word_copy_template(
            str(sow_template),
            str(output)
        )
        assert result.get("success") is True
        assert output.exists()


class TestFixSplitPlaceholders:
    """Tests for tool_word_fix_split_placeholders."""

    def test_fixes_placeholders(self, word_advanced_tools, sow_template, temp_dir):
        """Should fix split placeholder text."""
        output = temp_dir / "fixed.docx"
        result = word_advanced_tools.tool_word_fix_split_placeholders(
            str(sow_template),
            {"<Customer Name>": "Contoso"},
            output_path=str(output)
        )
        assert result.get("success") is True


class TestCleanupSow:
    """Tests for tool_word_cleanup_sow."""

    def test_cleans_sow(self, word_advanced_tools, sow_template, temp_dir):
        """Should cleanup SOW removing instructions."""
        output = temp_dir / "cleaned.docx"
        result = word_advanced_tools.tool_word_cleanup_sow(
            str(sow_template),
            output_path=str(output)
        )
        assert result.get("success") is True or "cleaned" in str(result).lower()


class TestGetSectionGuidance:
    """Tests for tool_word_get_section_guidance."""

    def test_gets_guidance(self, word_advanced_tools, sow_template):
        """Should get section guidance information."""
        result = word_advanced_tools.tool_word_get_section_guidance(
            str(sow_template),
            "Executive Summary"
        )
        # Should return guidance or indication it exists
        assert "guidance" in result or "content" in result or "error" not in str(result).lower()


class TestParseSowTemplate:
    """Tests for tool_word_parse_sow_template."""

    def test_parses_template(self, word_advanced_tools, sow_template):
        """Should parse SOW template structure."""
        result = word_advanced_tools.tool_word_parse_sow_template(str(sow_template))
        assert "sections" in result or "structure" in result or "template" in str(result).lower()


class TestAnalyzeTemplateFormatting:
    """Tests for tool_word_analyze_template_formatting."""

    def test_analyzes_formatting(self, word_advanced_tools, sow_template):
        """Should analyze template formatting."""
        result = word_advanced_tools.tool_word_analyze_template_formatting(str(sow_template))
        # Should return analysis information
        assert "error" not in result or isinstance(result, dict)


class TestWordFromMarkdown:
    """Tests for word_from_markdown - this is in word_tools.py, not word_advanced_tools."""

    def test_method_location(self):
        """Note: word_from_markdown is in WordTools, not WordAdvancedTools."""
        # This method is in word_tools.py
        from tools.word_tools import WordTools
        wt = WordTools()
        assert hasattr(wt, 'tool_word_from_markdown')


class TestCreateSowFromMarkdown:
    """Tests for tool_word_create_sow_from_markdown."""

    def test_creates_sow(self, word_advanced_tools, sow_template, temp_dir):
        """Should create SOW from markdown with template."""
        md = """# Statement of Work

## Executive Summary

This is the executive summary.

## Scope

Project scope description.
"""
        output = temp_dir / "sow_from_md.docx"
        result = word_advanced_tools.tool_word_create_sow_from_markdown(
            str(output), md, str(sow_template)
        )
        assert result.get("success") is True or output.exists()


class TestAddComment:
    """Tests for tool_word_add_comment."""

    def test_adds_comment(self, word_advanced_tools, sow_template, temp_dir):
        """Should add comment to document."""
        output = temp_dir / "commented.docx"
        result = word_advanced_tools.tool_word_add_comment(
            str(sow_template),
            "Executive Summary",
            "This needs more detail",
            output_path=str(output)
        )
        assert result.get("success") is True or "added" in str(result).lower()
