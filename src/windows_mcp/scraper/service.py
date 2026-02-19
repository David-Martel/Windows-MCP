"""Web scraping with SSRF protection.

Validates URLs against private/reserved IP ranges and non-HTTP schemes
before fetching. Follows redirects manually to validate each hop.
"""

import ipaddress
import socket
from urllib.parse import urlparse

import requests
from markdownify import markdownify


class ScraperService:
    """Fetch and convert web pages to markdown with SSRF guards."""

    @staticmethod
    def validate_url(url: str) -> None:
        """Validate a URL for SSRF safety.

        Blocks:
        - Non-HTTP(S) schemes (file://, ftp://, data:, etc.)
        - Private/reserved IP ranges (RFC 1918, link-local, loopback)
        - Cloud metadata endpoints (169.254.169.254, fd00::, etc.)

        Raises ValueError if the URL is unsafe.
        """
        parsed = urlparse(url)

        # Scheme check
        if parsed.scheme not in ("http", "https"):
            raise ValueError(
                f"Unsupported URL scheme '{parsed.scheme}'. Only http and https are allowed."
            )

        # Extract hostname
        hostname = parsed.hostname
        if not hostname:
            raise ValueError("URL has no hostname.")

        # Resolve to IP and check for private/reserved ranges
        try:
            addr = ipaddress.ip_address(hostname)
        except ValueError:
            # It's a domain name -- resolve it
            try:
                resolved = socket.getaddrinfo(hostname, None, socket.AF_UNSPEC, socket.SOCK_STREAM)
                if not resolved:
                    raise ValueError(f"Could not resolve hostname: {hostname}")
                # Check ALL resolved addresses (prevent DNS rebinding with multiple A records)
                for family, _, _, _, sockaddr in resolved:
                    ip_str = sockaddr[0]
                    addr = ipaddress.ip_address(ip_str)
                    if (
                        addr.is_private
                        or addr.is_reserved
                        or addr.is_loopback
                        or addr.is_link_local
                    ):
                        raise ValueError(
                            f"URL resolves to private/reserved IP {addr}. "
                            "Scraping internal network addresses is not allowed."
                        )
            except socket.gaierror as e:
                raise ValueError(f"Could not resolve hostname '{hostname}': {e}") from e
        else:
            # Direct IP address in URL
            if addr.is_private or addr.is_reserved or addr.is_loopback or addr.is_link_local:
                raise ValueError(
                    f"URL points to private/reserved IP {addr}. "
                    "Scraping internal network addresses is not allowed."
                )

    def scrape(self, url: str) -> str:
        """Fetch a URL and return its content as markdown.

        Validates URL for SSRF safety, follows redirects manually
        (validating each hop), and converts HTML to markdown.
        """
        self.validate_url(url)
        try:
            response = requests.get(url, timeout=10, allow_redirects=False)
            # Follow redirects manually to validate each hop
            redirects = 0
            while response.is_redirect and redirects < 5:
                redirect_url = response.headers.get("Location", "")
                if not redirect_url:
                    break
                self.validate_url(redirect_url)
                response = requests.get(redirect_url, timeout=10, allow_redirects=False)
                redirects += 1
            response.raise_for_status()
        except (ValueError, ConnectionError, TimeoutError):
            raise  # Re-raise our own validation errors
        except requests.exceptions.HTTPError as e:
            raise ValueError(f"HTTP error for {url}: {e}") from e
        except requests.exceptions.ConnectionError as e:
            raise ConnectionError(f"Failed to connect to {url}: {e}") from e
        except requests.exceptions.Timeout as e:
            raise TimeoutError(f"Request timed out for {url}: {e}") from e
        html = response.text
        content = markdownify(html=html)
        return content
