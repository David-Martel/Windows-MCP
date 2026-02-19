"""Tests for the VisionService module.

Tests cover configuration, API interaction (mocked), and JSON response
parsing without requiring a real vision API endpoint.
"""

import json
from unittest.mock import MagicMock, patch

import pytest

from windows_mcp.vision.service import VisionService

# ---------------------------------------------------------------------------
# Configuration tests
# ---------------------------------------------------------------------------


class TestVisionServiceConfig:
    """Test VisionService initialization and configuration."""

    def test_default_not_configured(self):
        svc = VisionService()
        assert not svc.is_configured

    def test_configured_via_constructor(self):
        svc = VisionService(api_url="http://localhost:8080/v1", api_key="test-key")
        assert svc.is_configured
        assert svc.api_url == "http://localhost:8080/v1"
        assert svc.api_key == "test-key"

    def test_configured_via_env(self):
        with patch.dict(
            "os.environ",
            {"VISION_API_URL": "http://vision:8080/v1", "VISION_API_KEY": "env-key"},
        ):
            svc = VisionService()
            assert svc.is_configured
            assert svc.api_key == "env-key"

    def test_model_default(self):
        svc = VisionService()
        assert svc.model == "gpt-4o"

    def test_model_override(self):
        svc = VisionService(model="claude-3.5-sonnet")
        assert svc.model == "claude-3.5-sonnet"

    def test_model_from_env(self):
        with patch.dict("os.environ", {"VISION_MODEL": "llava:latest"}):
            svc = VisionService()
            assert svc.model == "llava:latest"

    def test_url_trailing_slash_stripped(self):
        svc = VisionService(api_url="http://localhost:8080/v1/")
        assert svc.api_url == "http://localhost:8080/v1"

    def test_analyze_raises_when_not_configured(self):
        svc = VisionService()
        with pytest.raises(RuntimeError, match="not configured"):
            svc.analyze(b"fake-png-bytes")


# ---------------------------------------------------------------------------
# API interaction tests (mocked)
# ---------------------------------------------------------------------------


class TestVisionServiceAnalyze:
    """Test VisionService.analyze with mocked HTTP responses."""

    @pytest.fixture
    def svc(self):
        return VisionService(api_url="http://test:8080/v1", api_key="test-key")

    def _mock_response(self, content: str) -> MagicMock:
        mock = MagicMock()
        mock.raise_for_status = MagicMock()
        mock.json.return_value = {"choices": [{"message": {"content": content}}]}
        return mock

    def test_analyze_sends_correct_request(self, svc):
        with patch("windows_mcp.vision.service.requests.post") as mock_post:
            mock_post.return_value = self._mock_response("A desktop screenshot.")

            result = svc.analyze(b"\x89PNG\r\n", prompt="Describe this")

            mock_post.assert_called_once()
            call_args = mock_post.call_args
            assert call_args.kwargs["json"]["model"] == "gpt-4o"
            assert call_args.kwargs["json"]["messages"][0]["content"][0]["text"] == "Describe this"
            assert "Bearer test-key" in call_args.kwargs["headers"]["Authorization"]
            assert result == "A desktop screenshot."

    def test_analyze_returns_content(self, svc):
        with patch("windows_mcp.vision.service.requests.post") as mock_post:
            mock_post.return_value = self._mock_response("Two buttons and a text field.")
            result = svc.analyze(b"\x89PNG", prompt="Describe")
            assert result == "Two buttons and a text field."

    def test_analyze_no_auth_header_when_no_key(self):
        svc = VisionService(api_url="http://test:8080/v1")
        with patch("windows_mcp.vision.service.requests.post") as mock_post:
            mock_post.return_value = self._mock_response("ok")
            svc.analyze(b"\x89PNG")
            headers = mock_post.call_args.kwargs["headers"]
            assert "Authorization" not in headers


# ---------------------------------------------------------------------------
# JSON parsing tests
# ---------------------------------------------------------------------------


class TestVisionServiceIdentifyElements:
    """Test identify_elements JSON parsing from LLM responses."""

    @pytest.fixture
    def svc(self):
        return VisionService(api_url="http://test:8080/v1", api_key="k")

    def _mock_analyze(self, svc, response_text):
        """Patch analyze to return a canned string."""
        return patch.object(svc, "analyze", return_value=response_text)

    def test_parse_clean_json_array(self, svc):
        elements = [
            {"type": "button", "text": "OK", "x": 100, "y": 200, "width": 80, "height": 30},
            {"type": "text_field", "text": "", "x": 300, "y": 200, "width": 200, "height": 30},
        ]
        with self._mock_analyze(svc, json.dumps(elements)):
            result = svc.identify_elements(b"\x89PNG")
            assert len(result) == 2
            assert result[0]["type"] == "button"
            assert result[1]["type"] == "text_field"

    def test_parse_json_in_markdown_code_block(self, svc):
        elements = [{"type": "button", "text": "Save", "x": 50, "y": 50, "width": 60, "height": 25}]
        raw = f"Here are the elements:\n```json\n{json.dumps(elements)}\n```"
        with self._mock_analyze(svc, raw):
            result = svc.identify_elements(b"\x89PNG")
            assert len(result) == 1
            assert result[0]["text"] == "Save"

    def test_parse_json_in_bare_code_block(self, svc):
        elements = [{"type": "label", "text": "Name:", "x": 10, "y": 10, "width": 50, "height": 20}]
        raw = f"```\n{json.dumps(elements)}\n```"
        with self._mock_analyze(svc, raw):
            result = svc.identify_elements(b"\x89PNG")
            assert len(result) == 1

    def test_returns_empty_on_non_json(self, svc):
        with self._mock_analyze(svc, "I can see a desktop with some windows."):
            result = svc.identify_elements(b"\x89PNG")
            assert result == []

    def test_returns_empty_on_non_array(self, svc):
        with self._mock_analyze(svc, '{"type": "button"}'):
            result = svc.identify_elements(b"\x89PNG")
            assert result == []

    def test_target_parameter_added_to_prompt(self, svc):
        with patch.object(svc, "analyze", return_value="[]") as mock:
            svc.identify_elements(b"\x89PNG", target="Save button")
            prompt = mock.call_args.args[1]
            assert "Save button" in prompt


class TestVisionServiceDescribeScreen:
    """Test describe_screen method."""

    @pytest.fixture
    def svc(self):
        return VisionService(api_url="http://test:8080/v1", api_key="k")

    def test_returns_description(self, svc):
        with patch.object(svc, "analyze", return_value="Notepad is open with empty document."):
            result = svc.describe_screen(b"\x89PNG")
            assert "Notepad" in result

    def test_context_included_in_prompt(self, svc):
        with patch.object(svc, "analyze", return_value="ok") as mock:
            svc.describe_screen(b"\x89PNG", context="User is editing a spreadsheet")
            prompt = mock.call_args.args[1]
            assert "editing a spreadsheet" in prompt
