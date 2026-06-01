"""Tests for the OpenAI-compatible client adapter (host detection + message translation)."""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from llm_client import is_openai_compat_host, OpenAICompatClient


class HostDetectionTests(unittest.TestCase):
    def test_native_ollama_hosts_are_not_openai_compat(self):
        self.assertFalse(is_openai_compat_host("http://localhost:11434"))
        self.assertFalse(is_openai_compat_host("https://ollama.com"))

    def test_openai_compat_hosts_detected(self):
        self.assertTrue(is_openai_compat_host("http://localhost:3001"))
        self.assertTrue(is_openai_compat_host("https://my-freellmapi.example.com/v1"))
        self.assertTrue(is_openai_compat_host("http://gateway:8080/v1"))


class MessageTranslationTests(unittest.TestCase):
    def test_text_message_passthrough(self):
        m = OpenAICompatClient._translate_message({"role": "user", "content": "hello"})
        self.assertEqual(m, {"role": "user", "content": "hello"})

    def test_image_message_becomes_content_blocks(self):
        m = OpenAICompatClient._translate_message(
            {"role": "user", "content": "look", "images": ["BASE64DATA"]}
        )
        self.assertEqual(m["role"], "user")
        self.assertIsInstance(m["content"], list)
        self.assertEqual(m["content"][0], {"type": "text", "text": "look"})
        self.assertEqual(m["content"][1]["type"], "image_url")
        self.assertTrue(m["content"][1]["image_url"]["url"].startswith("data:image/jpeg;base64,"))

    def test_http_image_url_not_double_wrapped(self):
        m = OpenAICompatClient._translate_message(
            {"role": "user", "content": "x", "images": ["http://example.com/a.png"]}
        )
        self.assertEqual(m["content"][1]["image_url"]["url"], "http://example.com/a.png")


if __name__ == "__main__":
    unittest.main()
