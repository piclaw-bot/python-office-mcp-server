"""Tests for web_tools.py - Web content extraction tools."""

from importlib.util import find_spec

import pytest

# Check for required libraries
HAS_REQUESTS = find_spec("requests") is not None
HAS_READABILITY = find_spec("readability") is not None
HAS_MARKDOWNIFY = find_spec("markdownify") is not None
HAS_BS4 = find_spec("bs4") is not None


# Fixture web_tools is provided by conftest.py


class TestWebFetch:
    """Tests for web_fetch tool."""

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    @pytest.mark.parametrize("invalid_url", [
        "not-a-url",
        "ftp://example.com",
        "://missing-scheme",
    ])
    def test_fetch_invalid_urls(self, web_tools, invalid_url):
        """Test fetching invalid URLs returns error."""
        result = web_tools.tool_web_fetch(invalid_url)
        assert "error" in result

    @pytest.mark.skipif(
        not all([HAS_REQUESTS, HAS_READABILITY, HAS_MARKDOWNIFY]),
        reason="Web dependencies not installed"
    )
    def test_fetch_valid_url(self, web_tools):
        """Test fetching a valid URL."""
        # Use a simple, stable page
        result = web_tools.tool_web_fetch("https://example.com")
        assert result.get("success") is True
        assert "title" in result
        assert "content" in result
        assert result["word_count"] > 0


class TestWebSearch:
    """Tests for web_search tool."""

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    def test_search_empty_query(self, web_tools):
        """Test searching with empty query."""
        result = web_tools.tool_web_search("")
        assert "error" in result

    @pytest.mark.skipif(
        not all([HAS_REQUESTS, HAS_BS4]),
        reason="Web dependencies not installed"
    )
    def test_search_valid_query(self, web_tools):
        """Test searching with valid query."""
        result = web_tools.tool_web_search("python programming", max_results=2)
        assert result.get("success") is True
        assert "results" in result
        # DuckDuckGo may not always return results, so just check structure
        if result["result_count"] > 0:
            assert "title" in result["results"][0]
            assert "url" in result["results"][0]


class TestWebExtractLinks:
    """Tests for web_extract_links tool."""

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    def test_extract_links_invalid_url(self, web_tools):
        """Test extracting links from invalid URL."""
        result = web_tools.tool_web_extract_links("not-a-url")
        assert "error" in result

    @pytest.mark.skipif(
        not all([HAS_REQUESTS, HAS_BS4]),
        reason="Web dependencies not installed"
    )
    def test_extract_links_valid_url(self, web_tools):
        """Test extracting links from valid URL."""
        result = web_tools.tool_web_extract_links("https://example.com")
        assert result.get("success") is True
        assert "links" in result
        assert "link_count" in result


class TestWebExtractTables:
    """Tests for web_extract_tables tool."""

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    def test_extract_tables_invalid_url(self, web_tools):
        """Test extracting tables from invalid URL."""
        result = web_tools.tool_web_extract_tables("not-a-url")
        assert "error" in result

    @pytest.mark.skipif(
        not all([HAS_REQUESTS, HAS_BS4]),
        reason="Web dependencies not installed"
    )
    def test_extract_tables_valid_url(self, web_tools):
        """Test extracting tables from valid URL."""
        # example.com has no tables, but should return empty list
        result = web_tools.tool_web_extract_tables("https://example.com")
        assert result.get("success") is True
        assert "tables" in result
        assert result["table_count"] == 0


class TestWebCheckUrl:
    """Tests for web_check_url tool."""

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    @pytest.mark.parametrize("invalid_url", [
        "not-a-url",
        "ftp://example.com",
        "://missing-scheme",
    ])
    def test_check_invalid_urls(self, web_tools, invalid_url):
        """Test checking invalid URLs returns error."""
        result = web_tools.tool_web_check_url(invalid_url)
        assert result["exists"] is False
        assert "error" in result

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    def test_check_valid_url(self, web_tools):
        """Test checking a valid, accessible URL."""
        result = web_tools.tool_web_check_url("https://example.com")
        assert result["exists"] is True
        assert result["status_code"] == 200
        assert "content_type" in result

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    def test_check_nonexistent_url(self, web_tools):
        """Test checking a URL that returns 404."""
        result = web_tools.tool_web_check_url("https://example.com/this-page-does-not-exist-12345")
        assert result["exists"] is False
        assert result["status_code"] == 404

    @pytest.mark.skipif(not HAS_REQUESTS, reason="requests not installed")
    def test_check_url_with_redirect(self, web_tools):
        """Test checking a URL that redirects."""
        # http://example.com redirects to https://example.com
        result = web_tools.tool_web_check_url("http://example.com", follow_redirects=True)
        assert result["exists"] is True
        # Check that it detected the redirect
        if result.get("redirected"):
            assert result["final_url"] != result["url"]
