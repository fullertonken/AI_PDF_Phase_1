"""
ollama_client.py — Lightweight wrapper around the Ollama REST API.

Supports both text-only and vision models (e.g. qwen2.5-vl:7b).
All requests go to localhost. No data leaves the machine.
"""

import base64
import json
import urllib.request
import urllib.error
from pathlib import Path
from typing import Optional


DEFAULT_BASE_URL = "http://192.168.1.101:11434"
DEFAULT_MODEL = "qwen3-vl:8b"


class OllamaError(Exception):
    """Raised when Ollama returns an error or is unreachable."""
    pass


class OllamaClient:
    """Minimal Ollama REST client with vision support — no external dependencies."""

    def __init__(self, base_url: str = DEFAULT_BASE_URL, model: str = DEFAULT_MODEL):
        self.base_url = base_url.rstrip("/")
        self.model = model

    def is_available(self) -> bool:
        """Check if Ollama is running and the model is loaded."""
        try:
            req = urllib.request.Request(f"{self.base_url}/api/tags")
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read())
                model_names = [m.get("name", "") for m in data.get("models", [])]
                for name in model_names:
                    if name == self.model or name.startswith(self.model.split(":")[0]):
                        return True
                return False
        except Exception:
            return False

    def list_models(self) -> list[str]:
        """Return list of available model names."""
        try:
            req = urllib.request.Request(f"{self.base_url}/api/tags")
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read())
                return [m.get("name", "") for m in data.get("models", [])]
        except Exception:
            return []

    # ------------------------------------------------------------------
    # Image helpers
    # ------------------------------------------------------------------
    @staticmethod
    def image_to_base64(image_path: str) -> str:
        """Read an image file and return its base64-encoded content."""
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

    @staticmethod
    def resize_image_if_needed(image_path: str, max_pixels: int = 1344 * 1344) -> str:
        """
        If the image exceeds max_pixels, resize it proportionally and save
        to a temp file. Returns the (possibly new) path. This keeps VRAM
        usage under control for large scanned pages.
        """
        try:
            from PIL import Image
            img = Image.open(image_path)
            w, h = img.size
            if w * h <= max_pixels:
                return image_path
            scale = (max_pixels / (w * h)) ** 0.5
            new_w, new_h = int(w * scale), int(h * scale)
            img = img.resize((new_w, new_h), Image.LANCZOS)
            tmp_path = str(Path(image_path).with_suffix(".resized.png"))
            img.save(tmp_path, "PNG")
            return tmp_path
        except ImportError:
            return image_path

    # ------------------------------------------------------------------
    # Core generate — supports text-only and vision (image) calls
    # ------------------------------------------------------------------
    def generate(
        self,
        prompt: str,
        system: str = "",
        images: list[str] | None = None,
        temperature: float = 0.1,
        timeout: int = 180,
    ) -> str:
        """
        Send a prompt to Ollama and return the full response text.

        Args:
            prompt: The text prompt.
            system: Optional system prompt.
            images: Optional list of base64-encoded images. For vision models
                    like qwen2.5-vl, this is how you pass page images.
            temperature: Sampling temperature (low = deterministic).
            timeout: Request timeout in seconds.
        """
        payload = {
            "model": self.model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": temperature,
                "num_predict": 16384,
                "num_ctx": 16384,
            },
        }
        if system:
            payload["system"] = system
        if images:
            payload["images"] = images

        body = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            f"{self.base_url}/api/generate",
            data=body,
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read())
                return data.get("response", "")
        except urllib.error.HTTPError as e:
            raise OllamaError(f"Ollama returned HTTP {e.code}: {e.read().decode()}")
        except urllib.error.URLError as e:
            raise OllamaError(f"Cannot reach Ollama at {self.base_url}: {e}")
        except json.JSONDecodeError as e:
            raise OllamaError(f"Invalid JSON from Ollama: {e}")
        except Exception as e:
            raise OllamaError(f"Unexpected error calling Ollama: {e}")

    # ------------------------------------------------------------------
    # Vision convenience: send image file(s) + prompt
    # ------------------------------------------------------------------
    def generate_with_image(
        self,
        prompt: str,
        image_paths: list[str],
        system: str = "",
        temperature: float = 0.1,
        timeout: int = 180,
    ) -> str:
        """
        Send one or more images along with a text prompt.

        Args:
            prompt: Text prompt describing what to do with the images.
            image_paths: List of file paths to images (PNG, JPG).
            system: Optional system prompt.
        """
        b64_images = []
        temp_files = []
        for path in image_paths:
            resized = self.resize_image_if_needed(path)
            b64_images.append(self.image_to_base64(resized))
            if resized != path:
                temp_files.append(resized)

        try:
            return self.generate(
                prompt=prompt,
                system=system,
                images=b64_images,
                temperature=temperature,
                timeout=timeout,
            )
        finally:
            for tf in temp_files:
                try:
                    Path(tf).unlink()
                except OSError:
                    pass

    # ------------------------------------------------------------------
    # JSON extraction — text-only or with images
    # ------------------------------------------------------------------
    def generate_json(
        self,
        prompt: str,
        system: str = "",
        images: list[str] | None = None,
        temperature: float = 0.1,
        timeout: int = 180,
        retries: int = 2,
    ) -> dict:
        """
        Send a prompt (optionally with base64 images) expecting JSON back.
        Attempts to parse the response, retrying on malformed JSON.
        """
        last_error = None
        raw = ""
        for attempt in range(retries + 1):
            raw = self.generate(
                prompt, system=system, images=images,
                temperature=temperature, timeout=timeout,
            )
            try:
                return self._extract_json(raw)
            except (json.JSONDecodeError, ValueError) as e:
                last_error = e
                if attempt < retries:
                    if not raw.strip():
                        # Empty response — model likely hit context limit.
                        # Don't lengthen the prompt; just retry with a short nudge.
                        prompt = "Respond with a JSON object only. No explanation."
                    else:
                        prompt = (
                            f"{prompt}\n\n"
                            f"IMPORTANT: Your previous response was not valid JSON. "
                            f"Respond with ONLY a JSON object, no markdown, no explanation."
                        )
        raise OllamaError(
            f"Failed to get valid JSON after {retries + 1} attempts. "
            f"Last error: {last_error}\nLast response: {raw[:500]}"
        )

    def generate_json_with_image(
        self,
        prompt: str,
        image_paths: list[str],
        system: str = "",
        temperature: float = 0.1,
        timeout: int = 180,
        retries: int = 2,
    ) -> dict:
        """
        Convenience: send image files + prompt, expect JSON back.
        Handles base64 encoding and resizing internally.
        """
        b64_images = []
        temp_files = []
        for path in image_paths:
            resized = self.resize_image_if_needed(path)
            b64_images.append(self.image_to_base64(resized))
            if resized != path:
                temp_files.append(resized)

        try:
            return self.generate_json(
                prompt=prompt, system=system, images=b64_images,
                temperature=temperature, timeout=timeout, retries=retries,
            )
        finally:
            for tf in temp_files:
                try:
                    Path(tf).unlink()
                except OSError:
                    pass

    # ------------------------------------------------------------------
    # JSON parser
    # ------------------------------------------------------------------
    @staticmethod
    def _extract_json(text: str) -> dict:
        """Extract a JSON object from a response that might have extra text."""
        text = text.strip()

        # Try direct parse
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # Strip markdown code fences
        if "```" in text:
            parts = text.split("```")
            for part in parts:
                cleaned = part.strip()
                if cleaned.startswith("json"):
                    cleaned = cleaned[4:].strip()
                try:
                    return json.loads(cleaned)
                except json.JSONDecodeError:
                    continue

        # Find JSON object in text
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            try:
                return json.loads(text[start:end + 1])
            except json.JSONDecodeError:
                pass

        raise ValueError(f"No valid JSON found in response")
