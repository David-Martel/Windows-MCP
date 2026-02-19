"""Tests for ScraperService covering edge cases not exercised in test_security.py.

Missing coverage targets:
- Line 50:  validate_url() raises when getaddrinfo returns empty list
- Line 89:  scrape() breaks out of redirect loop when Location header is absent
- Lines 91-92: scrape() follows a valid redirect and re-fetches the redirect target
- Lines 96-101: scrape() converts requests.HTTPError, requests.Timeout,
                requests.ConnectionError into the appropriate Python exception types
- Happy path: successful scrape returns markdownified HTML content
- Happy path: validate_url() accepts a well-formed public URL without raising
"""

import socket
from unittest.mock import MagicMock, call, patch

import pytest
import requests

from windows_mcp.scraper.service import ScraperService

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_PUBLIC_DNS_RESPONSE = [(socket.AF_INET, socket.SOCK_STREAM, 0, "", ("93.184.216.34", 0))]

_PUBLIC_URL = "https://example.com"


def _make_response(
    *,
    is_redirect: bool = False,
    location: str | None = None,
    status_code: int = 200,
    text: str = "<html><body><h1>Hello</h1></body></html>",
    raise_for_status_exc: Exception | None = None,
) -> MagicMock:
    """Build a minimal mock requests.Response."""
    resp = MagicMock()
    resp.is_redirect = is_redirect
    resp.status_code = status_code
    resp.text = text
    headers: dict[str, str] = {}
    if location is not None:
        headers["Location"] = location
    resp.headers = headers
    if raise_for_status_exc is not None:
        resp.raise_for_status.side_effect = raise_for_status_exc
    else:
        resp.raise_for_status = MagicMock()
    return resp


# ---------------------------------------------------------------------------
# TestValidateUrlEdgeCases
# ---------------------------------------------------------------------------


class TestValidateUrlEdgeCases:
    """Edge cases in ScraperService.validate_url not covered by test_security.py."""

    def test_empty_getaddrinfo_result_raises(self):
        """Line 50: getaddrinfo returning [] must raise ValueError (hostname unresolvable)."""
        with patch("socket.getaddrinfo", return_value=[]):
            with pytest.raises(ValueError, match="Could not resolve hostname"):
                ScraperService.validate_url(_PUBLIC_URL)

    def test_valid_public_url_does_not_raise(self):
        """Happy path: a well-formed public HTTPS URL passes validation without error."""
        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            ScraperService.validate_url(_PUBLIC_URL)  # must not raise


# ---------------------------------------------------------------------------
# TestScrapeRedirectHandling
# ---------------------------------------------------------------------------


class TestScrapeRedirectHandling:
    """Tests for the manual redirect-following loop in scrape()."""

    def test_redirect_with_missing_location_header_breaks_loop(self):
        """Line 89: when a redirect response carries no Location header the loop
        exits immediately and the redirect response itself is used."""
        scraper = ScraperService()

        redirect_resp = _make_response(is_redirect=True)  # headers has no Location key
        # raise_for_status succeeds, text is plain HTML
        redirect_resp.text = "<p>Fallback</p>"

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=redirect_resp) as mock_get:
                result = scraper.scrape(_PUBLIC_URL)

        # requests.get should only be called once (the initial fetch; redirect loop broke)
        mock_get.assert_called_once()
        assert "Fallback" in result

    def test_redirect_with_empty_location_header_breaks_loop(self):
        """Line 89: an empty Location string is falsy -- the loop must break."""
        scraper = ScraperService()

        redirect_resp = _make_response(is_redirect=True, location="")
        redirect_resp.text = "<p>Empty location</p>"

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=redirect_resp) as mock_get:
                result = scraper.scrape(_PUBLIC_URL)

        mock_get.assert_called_once()
        assert "Empty location" in result

    def test_follows_redirect_to_valid_public_url(self):
        """Lines 91-92: scrape() follows a redirect that passes validate_url and
        issues a second requests.get for the redirect target."""
        scraper = ScraperService()

        redirect_url = "https://www.example.com/page"
        first_resp = _make_response(is_redirect=True, location=redirect_url)
        second_resp = _make_response(text="<h1>Redirected page</h1>")

        # DNS: both the initial URL and the redirect URL need public IP resolution
        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", side_effect=[first_resp, second_resp]) as mock_get:
                result = scraper.scrape(_PUBLIC_URL)

        assert mock_get.call_count == 2
        assert mock_get.call_args_list[1] == call(redirect_url, timeout=10, allow_redirects=False)
        assert "Redirected page" in result

    def test_redirect_to_private_ip_is_blocked(self):
        """Redirect pointing at a private IP must be blocked by validate_url."""
        scraper = ScraperService()

        first_resp = _make_response(is_redirect=True, location="http://192.168.0.1/secret")

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=first_resp):
                with pytest.raises(ValueError, match="private/reserved"):
                    scraper.scrape(_PUBLIC_URL)


# ---------------------------------------------------------------------------
# TestScrapeExceptionConversion
# ---------------------------------------------------------------------------


class TestScrapeExceptionConversion:
    """Lines 96-101: verify that requests exceptions are re-raised as stdlib types."""

    def test_http_error_becomes_value_error(self):
        """Line 96-97: requests.HTTPError -> ValueError with URL context in message."""
        scraper = ScraperService()

        http_err = requests.exceptions.HTTPError("404 Not Found")
        response = _make_response(raise_for_status_exc=http_err)

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=response):
                with pytest.raises(ValueError, match="HTTP error"):
                    scraper.scrape(_PUBLIC_URL)

    def test_http_error_message_contains_url(self):
        """ValueError raised for HTTPError must include the requested URL."""
        scraper = ScraperService()

        response = _make_response(
            raise_for_status_exc=requests.exceptions.HTTPError("500 Server Error")
        )

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=response):
                with pytest.raises(ValueError) as exc_info:
                    scraper.scrape(_PUBLIC_URL)

        assert _PUBLIC_URL in str(exc_info.value)

    def test_requests_timeout_becomes_timeout_error(self):
        """Line 100-101: requests.Timeout -> TimeoutError."""
        scraper = ScraperService()

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch(
                "requests.get",
                side_effect=requests.exceptions.Timeout("timed out"),
            ):
                with pytest.raises(TimeoutError, match="timed out"):
                    scraper.scrape(_PUBLIC_URL)

    def test_requests_connection_error_becomes_connection_error(self):
        """Line 98-99: requests.ConnectionError -> ConnectionError."""
        scraper = ScraperService()

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch(
                "requests.get",
                side_effect=requests.exceptions.ConnectionError("refused"),
            ):
                with pytest.raises(ConnectionError, match="Failed to connect"):
                    scraper.scrape(_PUBLIC_URL)

    def test_connection_error_message_contains_url(self):
        """ConnectionError message must include the target URL."""
        scraper = ScraperService()

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch(
                "requests.get",
                side_effect=requests.exceptions.ConnectionError("Connection refused"),
            ):
                with pytest.raises(ConnectionError) as exc_info:
                    scraper.scrape(_PUBLIC_URL)

        assert _PUBLIC_URL in str(exc_info.value)

    def test_timeout_error_message_contains_url(self):
        """TimeoutError message must include the target URL."""
        scraper = ScraperService()

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch(
                "requests.get",
                side_effect=requests.exceptions.Timeout("deadline exceeded"),
            ):
                with pytest.raises(TimeoutError) as exc_info:
                    scraper.scrape(_PUBLIC_URL)

        assert _PUBLIC_URL in str(exc_info.value)

    def test_value_error_from_validate_url_propagates_unchanged(self):
        """Line 94-95: ValueError (e.g. from validate_url inside redirect loop)
        is re-raised as-is, not wrapped again."""
        scraper = ScraperService()

        # Trigger: initial URL is public, redirect URL is a private IP
        first_resp = _make_response(is_redirect=True, location="http://10.0.0.1/secret")

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=first_resp):
                with pytest.raises(ValueError, match="private/reserved"):
                    scraper.scrape(_PUBLIC_URL)


# ---------------------------------------------------------------------------
# TestScrapeSuccessPath
# ---------------------------------------------------------------------------


class TestScrapeSuccessPath:
    """Happy-path tests for scrape() ensuring correct content conversion."""

    def test_returns_markdown_content(self):
        """scrape() converts the response HTML to markdown via markdownify."""
        scraper = ScraperService()
        html = "<html><body><h1>Main Heading</h1><p>Some text</p></body></html>"
        response = _make_response(text=html)

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=response):
                result = scraper.scrape(_PUBLIC_URL)

        # markdownify converts <h1> to '# Main Heading' and preserves paragraph text
        assert "Main Heading" in result
        assert "Some text" in result

    def test_no_network_call_for_blocked_scheme(self):
        """validate_url raises before requests.get is invoked for unsafe schemes."""
        scraper = ScraperService()

        with patch("requests.get") as mock_get:
            with pytest.raises(ValueError, match="scheme"):
                scraper.scrape("ftp://example.com/file.txt")

        mock_get.assert_not_called()

    def test_requests_get_called_with_correct_args(self):
        """scrape() passes timeout=10 and allow_redirects=False to requests.get."""
        scraper = ScraperService()
        response = _make_response()

        with patch("socket.getaddrinfo", return_value=_PUBLIC_DNS_RESPONSE):
            with patch("requests.get", return_value=response) as mock_get:
                scraper.scrape(_PUBLIC_URL)

        mock_get.assert_called_once_with(_PUBLIC_URL, timeout=10, allow_redirects=False)
