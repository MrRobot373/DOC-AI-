"""
LLM client adapters.

DOC-AI's review engine is written against the Ollama client's `.chat()` / `.list()`
shape. This module adds an adapter for any OpenAI-compatible endpoint
(FreeLLMAPI, LM Studio, vLLM, llama.cpp server, a remote gateway) so the same
engine can talk to a much wider set of free/open models behind one endpoint.

The adapter translates an Ollama-style call into an OpenAI `/v1/chat/completions`
call and returns the response in the Ollama shape the engine already parses, so
no engine code needs to change.
"""

from __future__ import annotations


class OpenAICompatClient:
    """Ollama-shaped wrapper over the OpenAI SDK pointed at any /v1 endpoint."""

    def __init__(self, api_key, host):
        from openai import OpenAI
        base = host.rstrip("/")
        if not base.endswith("/v1"):
            base = base + "/v1"
        # OpenAI SDK requires a non-empty key string even when the server ignores it.
        self._client = OpenAI(api_key=api_key or "none", base_url=base)

    def chat(self, model, messages, format=None, options=None, **kwargs):
        """
        Translate an Ollama-style chat call to OpenAI-compatible, and return the
        response in Ollama's {"message": {"content": ...}} shape.

        - `format` (a JSON schema dict in our engine) → response_format json_object
        - `options` (Ollama: temperature/seed/num_predict) → temperature/seed/max_tokens
        - image messages: Ollama uses message["images"] = [b64]; OpenAI uses
          content blocks with image_url data URLs — translated here.
        """
        options = options or {}
        oai_messages = [self._translate_message(m) for m in messages]

        req = {
            "model": model,
            "messages": oai_messages,
            "temperature": options.get("temperature", 0),
        }
        if "num_predict" in options:
            req["max_tokens"] = options["num_predict"]
        seed = options.get("seed")
        if seed is not None:
            req["seed"] = seed
        if format is not None:
            # Most OpenAI-compatible servers accept json_object; the prompt already
            # instructs the model to return a JSON array, so this is enough.
            req["response_format"] = {"type": "json_object"}

        resp = self._client.chat.completions.create(**req)
        content = resp.choices[0].message.content or ""
        return {"message": {"content": content}}

    @staticmethod
    def _translate_message(m):
        """Convert an Ollama message (possibly with `images`) to OpenAI content blocks."""
        images = m.get("images")
        if not images:
            return {"role": m["role"], "content": m.get("content", "")}
        blocks = [{"type": "text", "text": m.get("content", "")}]
        for b64 in images:
            url = b64 if str(b64).startswith("http") else f"data:image/jpeg;base64,{b64}"
            blocks.append({"type": "image_url", "image_url": {"url": url}})
        return {"role": m["role"], "content": blocks}

    def list(self):
        """Return available models in the Ollama-ish shape the engine expects."""
        models = self._client.models.list()

        class _Resp:
            pass

        r = _Resp()
        r.models = [type("M", (), {"model": m.id})() for m in models.data]
        return r


def is_openai_compat_host(host):
    """
    True when `host` is an OpenAI-compatible endpoint rather than native Ollama.

    Heuristics: explicit /v1 path, the FreeLLMAPI default port (3001), or a host
    that mentions 'freellmapi'. Native Ollama (11434, ollama.com) returns False.
    """
    h = (host or "").lower()
    if "ollama.com" in h or ":11434" in h:
        return False
    return "/v1" in h or ":3001" in h or "freellmapi" in h
