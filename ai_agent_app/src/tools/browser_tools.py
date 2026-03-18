"""
Browser automation and web scraping tools.
Uses subprocess for browser launch, urllib/html.parser for static scraping,
and Playwright when available for rendered JavaScript pages.
"""
import os
import shutil
import subprocess
import tempfile
from html.parser import HTMLParser
from urllib.error import HTTPError, URLError
from urllib.parse import quote_plus
from urllib.request import Request, urlopen


_LAST_BROWSER_PROCESS = None
_LAST_BROWSER_NAME = ""
_INTERACTIVE_DEBUG_PORT = None
_INTERACTIVE_USER_DATA_DIR = ""


def _get_sync_playwright():
    try:
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except Exception:
        return None


def _normalize_url(url: str = "", search_query: str = "") -> str:
    cleaned_url = (url or "").strip()
    cleaned_query = (search_query or "").strip()
    if cleaned_query and not cleaned_url:
        return f"https://www.google.com/search?q={quote_plus(cleaned_query)}"
    if not cleaned_url:
        return "https://www.google.com"
    if "://" not in cleaned_url:
        return f"https://{cleaned_url}"
    return cleaned_url


def _candidate_browser_paths(browser: str) -> list[str]:
    name = (browser or "edge").strip().lower()
    candidates = []
    executable_name = {
        "edge": "msedge.exe",
        "chrome": "chrome.exe",
        "firefox": "firefox.exe",
    }.get(name, "msedge.exe")

    resolved = shutil.which(executable_name)
    if resolved:
        candidates.append(resolved)

    program_files = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"), os.environ.get("LocalAppData")]
    known_suffixes = {
        "edge": [r"Microsoft\Edge\Application\msedge.exe"],
        "chrome": [r"Google\Chrome\Application\chrome.exe"],
        "firefox": [r"Mozilla Firefox\firefox.exe"],
    }
    for base in program_files:
        if not base:
            continue
        for suffix in known_suffixes.get(name, known_suffixes["edge"]):
            candidate = os.path.join(base, suffix)
            if os.path.exists(candidate):
                candidates.append(candidate)

    # Preserve order while removing duplicates.
    unique = []
    for candidate in candidates:
        if candidate not in unique:
            unique.append(candidate)
    return unique


def _resolve_browser_command(browser: str) -> tuple[str | None, str]:
    name = (browser or "edge").strip().lower()
    candidates = _candidate_browser_paths(name)
    if candidates:
        return candidates[0], name
    return None, name


def _playwright_channel(browser: str) -> tuple[str, str | None]:
    name = (browser or "edge").strip().lower()
    if name == "chrome":
        return "chromium", "chrome"
    if name == "firefox":
        return "firefox", None
    return "chromium", "msedge"


def _start_interactive_browser(url: str, browser: str) -> str:
    global _LAST_BROWSER_PROCESS, _LAST_BROWSER_NAME, _INTERACTIVE_DEBUG_PORT, _INTERACTIVE_USER_DATA_DIR

    sync_playwright = _get_sync_playwright()
    if sync_playwright is None:
        return "❌ Playwright is not installed. Install it with `pip install playwright` and `playwright install`."

    browser_cmd, browser_name = _resolve_browser_command(browser)
    if not browser_cmd:
        return f"❌ Browser not found for '{browser}'. Install Edge or Chrome for interactive scraping."
    if browser_name not in {"edge", "chrome"}:
        return "❌ Interactive scraping currently supports Edge and Chrome."

    try:
        if _LAST_BROWSER_PROCESS is not None and _LAST_BROWSER_PROCESS.poll() is None:
            try:
                _LAST_BROWSER_PROCESS.terminate()
                _LAST_BROWSER_PROCESS.wait(timeout=5)
            except Exception:
                pass

        _INTERACTIVE_DEBUG_PORT = 9222 if browser_name == "edge" else 9223
        _INTERACTIVE_USER_DATA_DIR = tempfile.mkdtemp(prefix=f"{browser_name}_interactive_")
        process = subprocess.Popen([
            browser_cmd,
            f"--remote-debugging-port={_INTERACTIVE_DEBUG_PORT}",
            f"--user-data-dir={_INTERACTIVE_USER_DATA_DIR}",
            url,
        ])
        _LAST_BROWSER_PROCESS = process
        _LAST_BROWSER_NAME = browser_name
        return f"✅ Opened interactive {browser_name} automation browser at {url}"
    except Exception as exc:
        return f"❌ Error opening interactive browser: {exc}"


class _SimpleHTMLExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_script = False
        self.in_style = False
        self.text_chunks = []
        self.links = []
        self.title = ""
        self._in_title = False

    def handle_starttag(self, tag, attrs):
        tag_name = (tag or "").lower()
        if tag_name == "script":
            self.in_script = True
        elif tag_name == "style":
            self.in_style = True
        elif tag_name == "a":
            href = dict(attrs).get("href", "").strip()
            if href:
                self.links.append(href)
        elif tag_name == "title":
            self._in_title = True

    def handle_endtag(self, tag):
        tag_name = (tag or "").lower()
        if tag_name == "script":
            self.in_script = False
        elif tag_name == "style":
            self.in_style = False
        elif tag_name == "title":
            self._in_title = False

    def handle_data(self, data):
        if self.in_script or self.in_style:
            return
        cleaned = " ".join((data or "").split())
        if not cleaned:
            return
        if self._in_title:
            self.title = cleaned
        else:
            self.text_chunks.append(cleaned)


def _extract_content_from_html(html: str, max_chars: int, include_links: bool, url: str) -> str:
    parser = _SimpleHTMLExtractor()
    parser.feed(html)
    page_text = " ".join(parser.text_chunks)
    page_text = page_text[:max(max_chars, 200)].strip()
    if not page_text:
        page_text = "No readable text found."

    parts = [f"✅ Scraped {url}"]
    if parser.title:
        parts.append(f"Title: {parser.title}")
    parts.append(f"Text: {page_text}")
    if include_links and parser.links:
        preview_links = ", ".join(parser.links[:10])
        parts.append(f"Links: {preview_links}")
    return "\n".join(parts)


def _scrape_page_with_playwright(url: str, browser: str, max_chars: int, include_links: bool, wait_until: str) -> str:
    sync_playwright = _get_sync_playwright()
    if sync_playwright is None:
        return (
            "❌ Playwright is not installed. Install it with `pip install playwright` "
            "to enable rendered JavaScript scraping."
        )

    engine_name, channel = _playwright_channel(browser)
    try:
        with sync_playwright() as playwright:
            engine = getattr(playwright, engine_name)
            launch_kwargs = {"headless": True}
            if channel:
                launch_kwargs["channel"] = channel
            browser_instance = engine.launch(**launch_kwargs)
            page = browser_instance.new_page()
            strategies = []
            preferred = (wait_until or "").strip() or "domcontentloaded"
            for candidate in [preferred, "domcontentloaded", "load", "networkidle"]:
                if candidate not in strategies:
                    strategies.append(candidate)

            last_error = None
            for strategy in strategies:
                try:
                    page.goto(url, wait_until=strategy, timeout=15000)
                    break
                except Exception as exc:
                    last_error = exc
            else:
                raise last_error or RuntimeError("Could not load page in rendered mode.")

            page.wait_for_timeout(1500)
            html = page.content()
            browser_instance.close()
        rendered = _extract_content_from_html(html, max_chars=max_chars, include_links=include_links, url=url)
        return f"{rendered}\nMode: Rendered browser page"
    except Exception as exc:
        return f"❌ Rendered scraping failed: {exc}"


def _connect_to_interactive_page():
    sync_playwright = _get_sync_playwright()
    if sync_playwright is None:
        raise RuntimeError("Playwright is not installed.")
    if _INTERACTIVE_DEBUG_PORT is None:
        raise RuntimeError("No interactive browser session is active.")

    playwright = sync_playwright().start()
    try:
        browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{_INTERACTIVE_DEBUG_PORT}")
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.pages[0] if context.pages else context.new_page()
        return playwright, browser, page
    except Exception:
        playwright.stop()
        raise


def browser_open(url: str = "", browser: str = "edge", search_query: str = "",
                 interactive: bool = False) -> str:
    """Open a browser with a URL or search query."""
    global _LAST_BROWSER_PROCESS, _LAST_BROWSER_NAME

    target_url = _normalize_url(url=url, search_query=search_query)
    if interactive:
        return _start_interactive_browser(target_url, browser)
    browser_cmd, browser_name = _resolve_browser_command(browser)
    if not browser_cmd:
        return f"❌ Browser not found for '{browser_name}'. Install Edge, Chrome, or Firefox."

    try:
        process = subprocess.Popen([browser_cmd, target_url])
        _LAST_BROWSER_PROCESS = process
        _LAST_BROWSER_NAME = browser_name
        return f"✅ Opened {browser_name} with {target_url}"
    except Exception as exc:
        return f"❌ Error opening browser: {exc}"


def browser_close(browser: str = "", close_all: bool = False) -> str:
    """Close the last opened browser window or all windows for a browser."""
    global _LAST_BROWSER_PROCESS, _LAST_BROWSER_NAME
    global _INTERACTIVE_DEBUG_PORT, _INTERACTIVE_USER_DATA_DIR

    browser_name = (browser or _LAST_BROWSER_NAME or "edge").strip().lower()
    process_name = {
        "edge": "msedge.exe",
        "chrome": "chrome.exe",
        "firefox": "firefox.exe",
    }.get(browser_name)

    if close_all:
        if not process_name:
            return f"❌ Unsupported browser '{browser_name}'."
        try:
            _INTERACTIVE_DEBUG_PORT = None
            completed = subprocess.run(
                ["taskkill", "/IM", process_name, "/F"],
                capture_output=True,
                text=True,
                check=False,
            )
            if completed.returncode == 0:
                _LAST_BROWSER_PROCESS = None
                if _INTERACTIVE_USER_DATA_DIR:
                    shutil.rmtree(_INTERACTIVE_USER_DATA_DIR, ignore_errors=True)
                    _INTERACTIVE_USER_DATA_DIR = ""
                return f"✅ Closed all {browser_name} windows."
            stderr = (completed.stderr or completed.stdout or "").strip()
            return f"❌ Could not close {browser_name}: {stderr or 'process not running'}"
        except Exception as exc:
            return f"❌ Error closing browser: {exc}"

    if _INTERACTIVE_DEBUG_PORT is not None and _LAST_BROWSER_PROCESS is not None:
        try:
            if _LAST_BROWSER_PROCESS.poll() is None:
                _LAST_BROWSER_PROCESS.terminate()
                _LAST_BROWSER_PROCESS.wait(timeout=5)
            _INTERACTIVE_DEBUG_PORT = None
            if _INTERACTIVE_USER_DATA_DIR:
                shutil.rmtree(_INTERACTIVE_USER_DATA_DIR, ignore_errors=True)
                _INTERACTIVE_USER_DATA_DIR = ""
            _LAST_BROWSER_PROCESS = None
            return f"✅ Closed the interactive {browser_name} browser."
        except Exception as exc:
            return f"❌ Error closing browser: {exc}"

    if _LAST_BROWSER_PROCESS is None:
        return "❌ No browser process is currently tracked. Use close_all=true to close all browser windows."

    try:
        if _LAST_BROWSER_PROCESS.poll() is None:
            _LAST_BROWSER_PROCESS.terminate()
            _LAST_BROWSER_PROCESS.wait(timeout=5)
        _LAST_BROWSER_PROCESS = None
        return f"✅ Closed the last opened {browser_name} window."
    except Exception as exc:
        return f"❌ Error closing browser: {exc}"


def browser_scrape_current_page(max_chars: int = 2000, include_links: bool = True) -> str:
    """Scrape the currently open interactive automation page."""
    if _INTERACTIVE_DEBUG_PORT is None:
        return "❌ No interactive browser page is open. Open a browser with interactive=true first."

    playwright = None
    browser = None
    try:
        playwright, browser, page = _connect_to_interactive_page()
        page.wait_for_timeout(1000)
        html = page.content()
        current_url = page.url
        result = _extract_content_from_html(
            html,
            max_chars=max_chars,
            include_links=include_links,
            url=current_url,
        )
        return f"{result}\nMode: Interactive current page"
    except Exception as exc:
        return f"❌ Error scraping current browser page: {exc}"
    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass
        if playwright is not None:
            try:
                playwright.stop()
            except Exception:
                pass


def web_scrape_page(url: str, max_chars: int = 2000, include_links: bool = True,
                    render_js: bool = False, browser: str = "edge",
                    wait_until: str = "domcontentloaded") -> str:
    """Fetch a webpage and return readable text plus optional links."""
    try:
        sync_playwright = _get_sync_playwright()
        target_url = _normalize_url(url=url)
        if render_js:
            rendered_result = _scrape_page_with_playwright(
                url=target_url,
                browser=browser,
                max_chars=max_chars,
                include_links=include_links,
                wait_until=wait_until,
            )
            if not rendered_result.startswith("❌"):
                return rendered_result
            if sync_playwright is not None:
                return rendered_result
        request = Request(
            target_url,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
        )
        with urlopen(request, timeout=20) as response:
            content_type = response.headers.get("Content-Type", "")
            if "html" not in content_type.lower():
                return f"❌ The page at {target_url} is not HTML content."
            html = response.read().decode("utf-8", errors="replace")

        static_result = _extract_content_from_html(
            html,
            max_chars=max_chars,
            include_links=include_links,
            url=target_url,
        )
        if render_js and sync_playwright is None:
            return f"{static_result}\nMode: Static fallback (Playwright not installed)"
        return f"{static_result}\nMode: Static HTML"
    except HTTPError as exc:
        return f"❌ HTTP error while scraping {url}: {exc.code}"
    except URLError as exc:
        return f"❌ Network error while scraping {url}: {exc.reason}"
    except Exception as exc:
        return f"❌ Error scraping webpage: {exc}"
