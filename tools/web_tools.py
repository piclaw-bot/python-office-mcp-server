#!/usr/bin/env python3
"""
web_tools.py - MCP tools for web content extraction and search.

Provides tools to fetch, extract, and convert web content to Markdown format
using readability for content extraction and markdownify for HTML conversion.
"""

import re
import warnings
from typing import Any
from urllib.parse import urljoin, urlparse

try:
    from urllib3.exceptions import InsecureRequestWarning
except ImportError:
    InsecureRequestWarning = None

# Check for required libraries
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from readability import Document
    HAS_READABILITY = True
except ImportError:
    HAS_READABILITY = False

try:
    from markdownify import markdownify as md
    HAS_MARKDOWNIFY = True
except ImportError:
    HAS_MARKDOWNIFY = False

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False


# Default headers to mimic a browser
DEFAULT_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
}


def _request_with_ssl_fallback(method: str, url: str, **kwargs):
    """Perform an HTTP request and retry with verify=False on SSL cert failures."""
    if not HAS_REQUESTS:
        raise RuntimeError("requests is not available")

    kwargs = dict(kwargs)
    kwargs.setdefault("verify", True)

    try:
        return requests.request(method, url, **kwargs)
    except requests.exceptions.SSLError:
        kwargs["verify"] = False
        with warnings.catch_warnings():
            if InsecureRequestWarning is not None:
                warnings.simplefilter("ignore", InsecureRequestWarning)
            return requests.request(method, url, **kwargs)


class WebTools:
    """MCP tool mixin for web content extraction.

    Provides tools to:
    - Fetch and extract readable content from web pages
    - Convert HTML to clean Markdown
    - Search the web (via DuckDuckGo)
    - Extract specific elements from pages

    Required packages:
    - requests: HTTP requests
    - readability-lxml: Content extraction
    - markdownify: HTML to Markdown conversion
    - beautifulsoup4: HTML parsing
    """

    def tool_web_fetch(
        self,
        url: str,
        extract_content: bool = True,
        include_links: bool = True,
        include_images: bool = False,
        timeout: int = 30,
    ) -> dict[str, Any]:
        """Fetch a web page and extract its content as Markdown.

        Uses readability to extract the main content from the page,
        removing navigation, ads, and other clutter. Then converts
        the clean HTML to Markdown format.

        Example:
            web_fetch(url="https://example.com/article")

            web_fetch(
                url="https://docs.microsoft.com/en-us/azure/...",
                include_links=True,
                include_images=True
            )

        Args:
            url: The URL to fetch
            extract_content: Use readability to extract main content (default: True)
            include_links: Include hyperlinks in output (default: True)
            include_images: Include image references (default: False)
            timeout: Request timeout in seconds (default: 30)

        Returns:
            Dictionary with title, content (markdown), url, and metadata
        """
        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        if extract_content and not HAS_READABILITY:
            return {"error": "readability-lxml not installed. Run: pip install readability-lxml"}

        if not HAS_MARKDOWNIFY:
            return {"error": "markdownify not installed. Run: pip install markdownify"}

        # Validate URL
        parsed = urlparse(url)
        if not parsed.scheme or not parsed.netloc:
            return {"error": f"Invalid URL: {url}"}

        if parsed.scheme not in ('http', 'https'):
            return {"error": f"Unsupported URL scheme: {parsed.scheme}"}

        try:
            response = _request_with_ssl_fallback("GET", url, headers=DEFAULT_HEADERS, timeout=timeout)
            response.raise_for_status()
        except requests.exceptions.Timeout:
            return {"error": f"Request timed out after {timeout} seconds"}
        except requests.exceptions.RequestException as e:
            return {"error": f"Failed to fetch URL: {str(e)}"}

        # Get content type
        content_type = response.headers.get('content-type', '')
        if 'text/html' not in content_type and 'application/xhtml' not in content_type:
            return {
                "error": f"Not an HTML page (content-type: {content_type})",
                "url": url,
            }

        html = response.text

        # Extract main content using readability
        if extract_content:
            try:
                doc = Document(html)
                title = doc.title()
                content_html = doc.summary()
            except Exception as e:
                return {"error": f"Failed to extract content: {str(e)}"}
        else:
            title = _extract_title(html)
            content_html = html

        # Convert to Markdown
        try:
            markdown_content = _html_to_markdown(
                content_html,
                base_url=url,
                include_links=include_links,
                include_images=include_images,
            )
        except Exception as e:
            return {"error": f"Failed to convert to Markdown: {str(e)}"}

        # Clean up the markdown
        markdown_content = _clean_markdown(markdown_content)

        return {
            "success": True,
            "url": url,
            "title": title,
            "content": markdown_content,
            "content_length": len(markdown_content),
            "word_count": len(markdown_content.split()),
        }

    def tool_web_search(
        self,
        query: str,
        max_results: int = 5,
        region: str = "wt-wt",
    ) -> dict[str, Any]:
        """Search the web using DuckDuckGo and return results.

        Performs a web search and returns titles, URLs, and snippets
        for the top results. Does not fetch the full page content -
        use web_fetch for that.

        Example:
            web_search(query="Microsoft Fabric lakehouse architecture")

            web_search(
                query="python-pptx table formatting",
                max_results=10
            )

        Args:
            query: Search query string
            max_results: Maximum number of results to return (default: 5)
            region: DuckDuckGo region code (default: "wt-wt" for worldwide)

        Returns:
            Dictionary with search results (title, url, snippet)
        """
        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        if not query or not query.strip():
            return {"error": "Search query cannot be empty"}

        try:
            # Use DuckDuckGo HTML search (no API key required)
            results = _duckduckgo_search(query, max_results, region)
            return {
                "success": True,
                "query": query,
                "result_count": len(results),
                "results": results,
            }
        except Exception as e:
            return {"error": f"Search failed: {str(e)}"}

    def tool_web_extract_links(
        self,
        url: str,
        filter_pattern: str = None,
        same_domain_only: bool = False,
        timeout: int = 30,
    ) -> dict[str, Any]:
        """Extract all links from a web page.

        Fetches a page and extracts all hyperlinks, optionally filtering
        by pattern or domain. Useful for discovering related pages or
        building navigation maps.

        Example:
            web_extract_links(url="https://docs.microsoft.com/...")

            web_extract_links(
                url="https://example.com",
                filter_pattern=r"/docs/",
                same_domain_only=True
            )

        Args:
            url: The URL to extract links from
            filter_pattern: Regex pattern to filter links (optional)
            same_domain_only: Only return links to the same domain (default: False)
            timeout: Request timeout in seconds (default: 30)

        Returns:
            Dictionary with extracted links and their text
        """
        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        if not HAS_BS4:
            return {"error": "beautifulsoup4 not installed. Run: pip install beautifulsoup4"}

        parsed_url = urlparse(url)
        if not parsed_url.scheme or not parsed_url.netloc:
            return {"error": f"Invalid URL: {url}"}

        try:
            response = _request_with_ssl_fallback("GET", url, headers=DEFAULT_HEADERS, timeout=timeout)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            return {"error": f"Failed to fetch URL: {str(e)}"}

        soup = BeautifulSoup(response.text, 'html.parser')
        links = []

        for a_tag in soup.find_all('a', href=True):
            href = a_tag['href']
            text = a_tag.get_text(strip=True)

            # Make absolute URL
            absolute_url = urljoin(url, href)
            parsed_link = urlparse(absolute_url)

            # Filter by domain if requested
            if same_domain_only and parsed_link.netloc != parsed_url.netloc:
                continue

            # Filter by pattern if provided
            if filter_pattern and not re.search(filter_pattern, absolute_url):
                continue

            # Skip non-http(s) links
            if parsed_link.scheme not in ('http', 'https'):
                continue

            links.append({
                "url": absolute_url,
                "text": text[:200] if text else "(no text)",
            })

        # Deduplicate by URL
        seen = set()
        unique_links = []
        for link in links:
            if link["url"] not in seen:
                seen.add(link["url"])
                unique_links.append(link)

        return {
            "success": True,
            "source_url": url,
            "link_count": len(unique_links),
            "links": unique_links,
        }

    def tool_web_extract_tables(
        self,
        url: str,
        table_index: int = None,
        timeout: int = 30,
    ) -> dict[str, Any]:
        """Extract tables from a web page as structured data.

        Fetches a page and extracts HTML tables, converting them to
        a structured format with headers and rows.

        Example:
            web_extract_tables(url="https://example.com/data")

            web_extract_tables(
                url="https://example.com/report",
                table_index=0  # Get only the first table
            )

        Args:
            url: The URL to extract tables from
            table_index: Specific table index to extract (optional, 0-based)
            timeout: Request timeout in seconds (default: 30)

        Returns:
            Dictionary with extracted tables
        """
        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        if not HAS_BS4:
            return {"error": "beautifulsoup4 not installed. Run: pip install beautifulsoup4"}

        try:
            response = _request_with_ssl_fallback("GET", url, headers=DEFAULT_HEADERS, timeout=timeout)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            return {"error": f"Failed to fetch URL: {str(e)}"}

        soup = BeautifulSoup(response.text, 'html.parser')
        html_tables = soup.find_all('table')

        if not html_tables:
            return {
                "success": True,
                "url": url,
                "table_count": 0,
                "tables": [],
                "note": "No tables found on page",
            }

        tables = []
        for idx, html_table in enumerate(html_tables):
            if table_index is not None and idx != table_index:
                continue

            # Extract headers
            headers = []
            header_row = html_table.find('thead')
            if header_row:
                headers = [th.get_text(strip=True) for th in header_row.find_all(['th', 'td'])]
            else:
                # Try first row as header
                first_row = html_table.find('tr')
                if first_row:
                    headers = [cell.get_text(strip=True) for cell in first_row.find_all(['th', 'td'])]

            # Extract rows
            rows = []
            tbody = html_table.find('tbody') or html_table
            for tr in tbody.find_all('tr'):
                cells = [cell.get_text(strip=True) for cell in tr.find_all(['td', 'th'])]
                if cells:
                    # Skip if this is the header row we already captured
                    if cells == headers:
                        continue
                    rows.append(cells)

            tables.append({
                "index": idx,
                "headers": headers,
                "row_count": len(rows),
                "rows": rows,
            })

        return {
            "success": True,
            "url": url,
            "table_count": len(tables),
            "tables": tables,
        }

    def tool_web_check_url(
        self,
        url: str,
        timeout: int = 10,
        follow_redirects: bool = True,
    ) -> dict[str, Any]:
        """Check if a URL exists and is accessible.

        Performs a HEAD request (or GET if HEAD fails) to verify the URL
        returns a valid response. Useful for validating links before
        including them in documents.

        Example:
            web_check_url(url="https://example.com/page")

            web_check_url(
                url="https://docs.microsoft.com/...",
                timeout=5,
                follow_redirects=True
            )

        Args:
            url: The URL to check
            timeout: Request timeout in seconds (default: 10)
            follow_redirects: Whether to follow redirects (default: True)

        Returns:
            Dictionary with exists, status_code, final_url, and content_type
        """
        if not HAS_REQUESTS:
            return {"error": "requests not installed. Run: pip install requests"}

        # Validate URL format
        parsed = urlparse(url)
        if not parsed.scheme or not parsed.netloc:
            return {
                "exists": False,
                "url": url,
                "error": "Invalid URL format",
            }

        if parsed.scheme not in ('http', 'https'):
            return {
                "exists": False,
                "url": url,
                "error": f"Unsupported URL scheme: {parsed.scheme}",
            }

        try:
            # Try HEAD first (faster, less bandwidth)
            response = _request_with_ssl_fallback(
                "HEAD",
                url,
                headers=DEFAULT_HEADERS,
                timeout=timeout,
                allow_redirects=follow_redirects,
            )

            # Some servers don't support HEAD, fall back to GET
            if response.status_code == 405:  # Method Not Allowed
                response = _request_with_ssl_fallback(
                    "GET",
                    url,
                    headers=DEFAULT_HEADERS,
                    timeout=timeout,
                    allow_redirects=follow_redirects,
                    stream=True,  # Don't download body
                )
                response.close()

            # Consider 2xx and 3xx as "exists"
            exists = response.status_code < 400

            return {
                "exists": exists,
                "url": url,
                "final_url": response.url if follow_redirects else url,
                "status_code": response.status_code,
                "content_type": response.headers.get('content-type', ''),
                "redirected": response.url != url if follow_redirects else False,
            }

        except requests.exceptions.Timeout:
            return {
                "exists": False,
                "url": url,
                "error": f"Request timed out after {timeout} seconds",
            }
        except requests.exceptions.ConnectionError:
            return {
                "exists": False,
                "url": url,
                "error": "Connection failed - host may be unreachable",
            }
        except requests.exceptions.RequestException as e:
            return {
                "exists": False,
                "url": url,
                "error": str(e),
            }


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def _extract_title(html: str) -> str:
    """Extract title from HTML."""
    if HAS_BS4:
        soup = BeautifulSoup(html, 'html.parser')
        title_tag = soup.find('title')
        if title_tag:
            return title_tag.get_text(strip=True)

    # Fallback: regex
    match = re.search(r'<title[^>]*>([^<]+)</title>', html, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    return "(no title)"


def _html_to_markdown(
    html: str,
    base_url: str = None,
    include_links: bool = True,
    include_images: bool = False,
) -> str:
    """Convert HTML to Markdown using markdownify."""
    if not HAS_MARKDOWNIFY:
        # Basic fallback - strip HTML tags
        text = re.sub(r'<[^>]+>', '', html)
        return text.strip()

    # Configure markdownify options
    options = {
        'heading_style': 'atx',
        'bullets': '-',
        'strip': ['script', 'style', 'nav', 'footer', 'header', 'aside'],
    }

    if not include_links:
        options['strip'].append('a')

    if not include_images:
        options['strip'].append('img')

    markdown = md(html, **options)

    # Make relative URLs absolute if base_url provided
    if base_url and include_links:
        # Fix relative links
        def make_absolute(match):
            link_text = match.group(1)
            link_url = match.group(2)
            if link_url and not urlparse(link_url).scheme:
                link_url = urljoin(base_url, link_url)
            return f'[{link_text}]({link_url})'

        markdown = re.sub(r'\[([^\]]*)\]\(([^)]+)\)', make_absolute, markdown)

    return markdown


def _clean_markdown(text: str) -> str:
    """Clean up markdown output."""
    # Remove excessive blank lines
    text = re.sub(r'\n{3,}', '\n\n', text)

    # Remove leading/trailing whitespace from lines
    lines = [line.rstrip() for line in text.split('\n')]
    text = '\n'.join(lines)

    # Remove excessive spaces
    text = re.sub(r'[ \t]{2,}', ' ', text)

    return text.strip()


def _duckduckgo_search(query: str, max_results: int = 5, region: str = "wt-wt") -> list[dict]:
    """Perform a DuckDuckGo search using the HTML interface."""
    if not HAS_BS4:
        raise ImportError("beautifulsoup4 required for search")

    # DuckDuckGo HTML search URL
    search_url = "https://html.duckduckgo.com/html/"
    params = {
        'q': query,
        'kl': region,
    }

    response = requests.post(search_url, data=params, headers=DEFAULT_HEADERS, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, 'html.parser')
    results = []

    # Find search results
    for result in soup.find_all('div', class_='result'):
        if len(results) >= max_results:
            break

        # Extract title and URL
        title_elem = result.find('a', class_='result__a')
        if not title_elem:
            continue

        title = title_elem.get_text(strip=True)
        url = title_elem.get('href', '')

        # DuckDuckGo wraps URLs - extract actual URL
        if url.startswith('//duckduckgo.com/l/'):
            # Parse the redirect URL
            url_match = re.search(r'uddg=([^&]+)', url)
            if url_match:
                from urllib.parse import unquote
                url = unquote(url_match.group(1))

        # Extract snippet
        snippet_elem = result.find('a', class_='result__snippet')
        snippet = snippet_elem.get_text(strip=True) if snippet_elem else ""

        if url and title:
            results.append({
                "title": title,
                "url": url,
                "snippet": snippet[:300] if snippet else "",
            })

    return results
