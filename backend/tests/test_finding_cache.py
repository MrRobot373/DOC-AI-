"""The content-addressed finding cache returns identical results without re-calling the LLM."""

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


class _FakeClient:
    def __init__(self):
        self.calls = 0

    def chat(self, **kwargs):
        self.calls += 1
        payload = [{
            "category": "GRAMMAR_SPELLING", "severity": "MINOR", "page": "-",
            "section": "X", "comment": "c", "fix": "f", "fix_type": "MANUAL", "evidence": "e",
        }]
        return {"message": {"content": json.dumps(payload)}}


class FindingCacheTests(unittest.TestCase):
    def test_second_call_hits_cache(self):
        os.environ["DOCAI_CACHE_DIR"] = tempfile.mkdtemp()
        try:
            import review_engine as R
            client = _FakeClient()
            a = R._review_chunk_multipass(client, "m", "some chunk text", "summary", 1, ["GRAMMAR_SPELLING"], "")
            b = R._review_chunk_multipass(client, "m", "some chunk text", "summary", 1, ["GRAMMAR_SPELLING"], "")
            self.assertEqual(client.calls, 1, "Second identical call must be served from cache")
            self.assertEqual(a, b)
        finally:
            os.environ.pop("DOCAI_CACHE_DIR", None)

    def test_no_cache_when_disabled(self):
        os.environ.pop("DOCAI_CACHE_DIR", None)
        import review_engine as R
        client = _FakeClient()
        R._review_chunk_multipass(client, "m", "text2", "summary", 1, ["GRAMMAR_SPELLING"], "")
        R._review_chunk_multipass(client, "m", "text2", "summary", 1, ["GRAMMAR_SPELLING"], "")
        self.assertEqual(client.calls, 2, "Without DOCAI_CACHE_DIR every call hits the LLM")


if __name__ == "__main__":
    unittest.main()
