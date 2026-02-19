"""LLM vision analysis for screenshots and UI element identification.

Sends screenshots to vision-capable LLMs (Claude, GPT-4o, Gemini, or local
models via OpenAI-compatible APIs) and returns structured descriptions of
visible UI elements.  This complements the UIAutomation accessibility tree
for apps with poor a11y support (games, custom-rendered UIs, remote desktops).

Configuration is via environment variables:

    VISION_API_URL      Base URL for the OpenAI-compatible API (default: none)
    VISION_API_KEY      Bearer token for the API (default: none)
    VISION_MODEL        Model identifier (default: "gpt-4o")

When ``VISION_API_URL`` is unset, :meth:`VisionService.analyze` raises
``RuntimeError`` to signal that vision is not configured.
"""

import base64
import json
import logging
import os
from typing import Any

import requests

logger = logging.getLogger(__name__)


class VisionService:
    """Analyse screenshots using a vision-capable LLM.

    Works with any OpenAI-compatible ``/v1/chat/completions`` endpoint:

    - **OpenAI** GPT-4o / GPT-4 Vision
    - **Anthropic** Claude 3.5+ (via OpenAI-compatible proxy)
    - **Google** Gemini (via OpenAI-compatible proxy)
    - **Local** llama.cpp server, Ollama, vLLM, or PC-AI pcai-inference
    """

    def __init__(
        self,
        api_url: str | None = None,
        api_key: str | None = None,
        model: str | None = None,
    ) -> None:
        self.api_url = (api_url or os.environ.get("VISION_API_URL", "")).rstrip("/")
        self.api_key = api_key or os.environ.get("VISION_API_KEY", "")
        self.model = model or os.environ.get("VISION_MODEL", "gpt-4o")

    @property
    def is_configured(self) -> bool:
        """Return ``True`` if a vision API endpoint is configured."""
        return bool(self.api_url)

    def analyze(
        self,
        image_bytes: bytes,
        prompt: str = "Describe all UI elements visible in this screenshot.",
        max_tokens: int = 1024,
    ) -> str:
        """Send an image to a vision-capable LLM and return the text response.

        Parameters
        ----------
        image_bytes:
            PNG or JPEG image as raw bytes.
        prompt:
            Instruction for the vision model.
        max_tokens:
            Maximum tokens in the response.

        Raises
        ------
        RuntimeError
            If ``VISION_API_URL`` is not configured.
        requests.HTTPError
            If the API returns a non-2xx status.
        """
        if not self.is_configured:
            raise RuntimeError(
                "Vision API not configured.  Set VISION_API_URL and "
                "VISION_API_KEY environment variables."
            )

        b64_image = base64.b64encode(image_bytes).decode("ascii")

        headers: dict[str, str] = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{b64_image}",
                            },
                        },
                    ],
                }
            ],
            "max_tokens": max_tokens,
        }

        response = requests.post(
            f"{self.api_url}/chat/completions",
            headers=headers,
            json=payload,
            timeout=60,
        )
        response.raise_for_status()

        result = response.json()
        return result["choices"][0]["message"]["content"]

    def identify_elements(
        self,
        image_bytes: bytes,
        target: str = "",
        max_tokens: int = 2048,
    ) -> list[dict[str, Any]]:
        """Use a vision LLM to identify UI elements from a screenshot.

        Returns a list of dicts, each with keys:

        - ``type``: element type (button, text_field, menu, label, etc.)
        - ``text``: visible text on the element
        - ``x``: centre x coordinate (pixels)
        - ``y``: centre y coordinate (pixels)
        - ``width``: approximate width (pixels)
        - ``height``: approximate height (pixels)

        Parameters
        ----------
        image_bytes:
            PNG screenshot bytes.
        target:
            Optional focus hint (e.g. ``"Save button"``).
        max_tokens:
            Maximum tokens for the vision response.

        Returns
        -------
        list[dict]
            Parsed element descriptions.  Returns ``[]`` if the model
            response cannot be parsed as JSON.
        """
        prompt = (
            "Analyze this Windows desktop screenshot. "
            "List all visible UI elements with their approximate coordinates. "
            "For each element, provide a JSON object with these keys:\n"
            '- "type": element type (button, text_field, menu, label, checkbox, '
            "dropdown, tab, icon, link, toolbar, statusbar)\n"
            '- "text": visible text on the element (empty string if none)\n'
            '- "x": center x coordinate in pixels from left edge\n'
            '- "y": center y coordinate in pixels from top edge\n'
            '- "width": approximate width in pixels\n'
            '- "height": approximate height in pixels\n'
            "Return ONLY a JSON array of objects, no other text."
        )
        if target:
            prompt += f"\nFocus especially on elements related to: {target}"

        raw = self.analyze(image_bytes, prompt, max_tokens=max_tokens)

        # Extract JSON from the response (handles markdown code blocks).
        try:
            cleaned = raw.strip()
            if "```json" in cleaned:
                cleaned = cleaned.split("```json", 1)[1].split("```", 1)[0]
            elif "```" in cleaned:
                cleaned = cleaned.split("```", 1)[1].split("```", 1)[0]
            elements = json.loads(cleaned.strip())
            if isinstance(elements, list):
                return elements
            logger.warning("Vision response was not a JSON array: %s", type(elements))
            return []
        except (json.JSONDecodeError, IndexError):
            logger.warning("Could not parse vision response as JSON: %.200s", raw)
            return []

    def describe_screen(
        self,
        image_bytes: bytes,
        context: str = "",
    ) -> str:
        """Return a natural-language description of a screenshot.

        Useful for providing LLM agents with a high-level understanding of
        what is currently visible on screen, especially when the accessibility
        tree is sparse or unavailable.

        Parameters
        ----------
        image_bytes:
            PNG screenshot bytes.
        context:
            Optional context about the current task (e.g. ``"user is trying
            to save a document in Word"``).
        """
        prompt = (
            "Describe this Windows desktop screenshot concisely. "
            "Include: which application(s) are visible, the current state "
            "(dialogs, menus, selections), and any notable content. "
            "Be specific about button labels, menu items, and text fields."
        )
        if context:
            prompt += f"\nContext: {context}"

        return self.analyze(image_bytes, prompt, max_tokens=512)
