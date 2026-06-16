# ai_supplier_matcher.py
# Cloud AI supplier detection for Maafushivaru Hub
# Supports: Anthropic Claude, OpenAI GPT, Google Gemini, custom OpenAI-compatible endpoints

import re
import json
import logging
import threading
import urllib.request
import urllib.error
from typing import Optional, Tuple, List, Dict


# ---------------------------------------------------------------------------
# PROVIDER DEFINITIONS
# ---------------------------------------------------------------------------
PROVIDERS = {
    "anthropic": {
        "label":    "Anthropic Claude",
        "base_url": "https://api.anthropic.com/v1/messages",
        "models":   ["claude-haiku-4-5-20251001", "claude-sonnet-4-6", "claude-opus-4-6"],
        "default_model": "claude-haiku-4-5-20251001",
        "doc_url":  "https://console.anthropic.com/settings/keys",
    },
    "openai": {
        "label":    "OpenAI GPT",
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models":   ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"],
        "default_model": "gpt-4o-mini",
        "doc_url":  "https://platform.openai.com/api-keys",
    },
    "gemini": {
        "label":    "Google Gemini",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        "models":   ["gemini-1.5-flash", "gemini-1.5-pro"],
        "default_model": "gemini-1.5-flash",
        "doc_url":  "https://aistudio.google.com/app/apikey",
    },
    "custom": {
        "label":    "Custom / Local (OpenAI-compatible)",
        "base_url": "http://localhost:11434/v1/chat/completions",
        "models":   ["llama3", "mistral", "phi3"],
        "default_model": "llama3",
        "doc_url":  "",
    },
}

# Status codes
STATUS_CONNECTED    = "connected"     # green dot
STATUS_DISCONNECTED = "disconnected"  # red dot
STATUS_LOW_CREDIT   = "low_credit"    # yellow dot
STATUS_OFFLINE      = "offline"       # grey dot (using local OCR)


# ---------------------------------------------------------------------------
# MAIN AI MATCHER CLASS
# ---------------------------------------------------------------------------
class AISupplierMatcher:
    """
    Calls a cloud or local AI to identify the supplier from extracted PDF text.
    Thread-safe; caches the last connectivity status.
    """

    def __init__(self, config: dict):
        self.config = config
        self._status = STATUS_OFFLINE
        self._status_lock = threading.Lock()
        self._last_error: str = ""

    def _get_ai_cfg(self) -> dict:
        return self.config.get("ai_settings", {})

    @property
    def status(self) -> str:
        with self._status_lock:
            return self._status

    @status.setter
    def status(self, val: str):
        with self._status_lock:
            self._status = val

    def is_enabled(self) -> bool:
        cfg = self._get_ai_cfg()
        return bool(cfg.get("enabled", False) and cfg.get("api_key", "").strip())

    # ------------------------------------------------------------------
    # CONNECTIVITY TEST
    # ------------------------------------------------------------------
    def test_connection(self) -> Tuple[str, str]:
        """
        Returns (status_code, message).
        Runs a minimal ping to the selected provider.
        """
        if not self.is_enabled():
            self.status = STATUS_OFFLINE
            return STATUS_OFFLINE, "AI mode disabled or no API key set."

        try:
            name, conf = self.match_supplier(
                "INVOICE FROM ACME CORP TOTAL USD 100.00",
                candidates=["ACME CORP"],
                aliases={},
            )
            if name:
                self.status = STATUS_CONNECTED
                return STATUS_CONNECTED, f"Connected — test match: {name}"
            else:
                self.status = STATUS_CONNECTED
                return STATUS_CONNECTED, "Connected (no match on test text)"
        except _CreditError:
            self.status = STATUS_LOW_CREDIT
            return STATUS_LOW_CREDIT, "API key valid but credits exhausted or quota exceeded."
        except _AuthError:
            self.status = STATUS_DISCONNECTED
            return STATUS_DISCONNECTED, "Authentication failed — check your API key."
        except Exception as e:
            self.status = STATUS_DISCONNECTED
            return STATUS_DISCONNECTED, f"Connection failed: {e}"

    # ------------------------------------------------------------------
    # MATCH SUPPLIER
    # ------------------------------------------------------------------
    def match_supplier(
        self,
        text: str,
        candidates: List[str],
        aliases: Dict[str, List[str]],
    ) -> Tuple[Optional[str], float]:
        """
        Ask the AI to identify the supplier from the extracted text.
        Returns (supplier_name, confidence_0_to_100) or (None, 0.0).
        """
        if not self.is_enabled():
            return None, 0.0

        cfg = self._get_ai_cfg()
        provider = cfg.get("provider", "anthropic")
        cand_list = "\n".join(f"- {c}" for c in candidates[:80])  # cap to avoid huge prompts

        prompt = (
            "You are a document processing assistant for a hotel in the Maldives.\n"
            "Given the following extracted text from a supplier invoice or receiving report, "
            "identify which supplier from the list below issued this document.\n\n"
            f"SUPPLIER LIST:\n{cand_list}\n\n"
            f"EXTRACTED TEXT (first 1500 chars):\n{text[:1500]}\n\n"
            "Reply ONLY with a JSON object like:\n"
            '{"supplier": "EXACT NAME FROM LIST", "confidence": 92}\n'
            "If you cannot determine the supplier, reply:\n"
            '{"supplier": null, "confidence": 0}\n'
            "Do not add any other text."
        )

        try:
            raw_response = self._call_provider(provider, cfg, prompt)
            return self._parse_response(raw_response, candidates, aliases)
        except _CreditError:
            self.status = STATUS_LOW_CREDIT
            logging.warning("[AI] Credit/quota error")
            return None, 0.0
        except _AuthError:
            self.status = STATUS_DISCONNECTED
            logging.error("[AI] Authentication error")
            return None, 0.0
        except Exception as e:
            logging.error(f"[AI] match_supplier failed: {e}")
            self._last_error = str(e)
            self.status = STATUS_DISCONNECTED
            return None, 0.0

    # ------------------------------------------------------------------
    # PROVIDER ROUTING
    # ------------------------------------------------------------------
    def _call_provider(self, provider: str, cfg: dict, prompt: str) -> str:
        if provider == "anthropic":
            return self._call_anthropic(cfg, prompt)
        elif provider == "gemini":
            return self._call_gemini(cfg, prompt)
        else:
            # OpenAI-compatible (openai, custom/local)
            return self._call_openai_compatible(cfg, prompt)

    def _call_anthropic(self, cfg: dict, prompt: str) -> str:
        api_key = cfg.get("api_key", "")
        model   = cfg.get("model", PROVIDERS["anthropic"]["default_model"])
        url     = PROVIDERS["anthropic"]["base_url"]
        timeout = int(cfg.get("timeout_seconds", 20))

        payload = json.dumps({
            "model": model,
            "max_tokens": 256,
            "messages": [{"role": "user", "content": prompt}],
        }).encode("utf-8")

        req = urllib.request.Request(
            url,
            data=payload,
            headers={
                "Content-Type":      "application/json",
                "x-api-key":         api_key,
                "anthropic-version": "2023-06-01",
            },
            method="POST",
        )
        return self._do_request(req, timeout)

    def _call_openai_compatible(self, cfg: dict, prompt: str) -> str:
        provider = cfg.get("provider", "openai")
        api_key  = cfg.get("api_key", "")
        model    = cfg.get("model", PROVIDERS.get(provider, PROVIDERS["openai"])["default_model"])
        base_url = cfg.get("custom_base_url", "") or PROVIDERS.get(provider, {}).get("base_url", "")
        timeout  = int(cfg.get("timeout_seconds", 20))

        payload = json.dumps({
            "model": model,
            "max_tokens": 256,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
        }).encode("utf-8")

        headers = {"Content-Type": "application/json"}
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"

        req = urllib.request.Request(url=base_url, data=payload, headers=headers, method="POST")
        return self._do_request(req, timeout)

    def _call_gemini(self, cfg: dict, prompt: str) -> str:
        api_key = cfg.get("api_key", "")
        model   = cfg.get("model", PROVIDERS["gemini"]["default_model"])
        timeout = int(cfg.get("timeout_seconds", 20))
        url     = PROVIDERS["gemini"]["base_url"].format(model=model)
        full_url = f"{url}?key={api_key}"

        payload = json.dumps({
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"maxOutputTokens": 256, "temperature": 0.1},
        }).encode("utf-8")

        req = urllib.request.Request(
            full_url, data=payload,
            headers={"Content-Type": "application/json"}, method="POST",
        )
        return self._do_request(req, timeout)

    def _do_request(self, req: urllib.request.Request, timeout: int) -> str:
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8")
                self.status = STATUS_CONNECTED
                return raw
        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode("utf-8")
            except Exception:
                pass
            if e.code in (401, 403):
                raise _AuthError(f"HTTP {e.code}: {body}")
            if e.code == 429 or "credit" in body.lower() or "quota" in body.lower():
                raise _CreditError(f"HTTP {e.code}: {body}")
            raise RuntimeError(f"HTTP {e.code}: {body}")
        except urllib.error.URLError as e:
            raise RuntimeError(f"Network error: {e.reason}")

    # ------------------------------------------------------------------
    # RESPONSE PARSER
    # ------------------------------------------------------------------
    def _parse_response(
        self,
        raw: str,
        candidates: List[str],
        aliases: Dict[str, List[str]],
    ) -> Tuple[Optional[str], float]:
        """Extract JSON from AI response and validate against candidate list."""
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            pass
        else:
            # Direct JSON (Gemini / some custom models)
            return self._extract_from_dict(data, candidates, aliases)

        # Try to find JSON embedded in a larger response (Anthropic / OpenAI)
        try:
            # OpenAI-style
            data = json.loads(raw)
            content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
            return self._extract_from_json_string(content, candidates, aliases)
        except Exception:
            pass

        try:
            # Anthropic-style
            data = json.loads(raw)
            content = data.get("content", [{}])[0].get("text", "")
            return self._extract_from_json_string(content, candidates, aliases)
        except Exception:
            pass

        try:
            # Gemini-style
            data = json.loads(raw)
            content = (
                data.get("candidates", [{}])[0]
                    .get("content", {})
                    .get("parts", [{}])[0]
                    .get("text", "")
            )
            return self._extract_from_json_string(content, candidates, aliases)
        except Exception:
            pass

        logging.warning(f"[AI] Could not parse response: {raw[:300]}")
        return None, 0.0

    def _extract_from_json_string(
        self, content: str, candidates: List[str], aliases: Dict[str, List[str]]
    ) -> Tuple[Optional[str], float]:
        m = re.search(r"\{.*?\}", content, re.DOTALL)
        if not m:
            return None, 0.0
        try:
            obj = json.loads(m.group(0))
            return self._extract_from_dict(obj, candidates, aliases)
        except Exception:
            return None, 0.0

    def _extract_from_dict(
        self, obj: dict, candidates: List[str], aliases: Dict[str, List[str]]
    ) -> Tuple[Optional[str], float]:
        name = obj.get("supplier")
        conf = float(obj.get("confidence", 0))
        if not name:
            return None, 0.0

        name_upper = str(name).upper().strip()

        # Exact match against candidates
        for c in candidates:
            if c.upper() == name_upper:
                return c, conf

        # Alias resolution
        for main, alias_list in aliases.items():
            for alias in alias_list:
                if alias.upper() == name_upper:
                    return main, conf
            if main.upper() == name_upper:
                return main, conf

        # Fuzzy fallback — AI may have slightly paraphrased
        from rapidfuzz import process, fuzz
        all_cands = list(candidates)
        for sub in aliases.values():
            all_cands.extend(sub)
        m2 = process.extractOne(name_upper, all_cands, scorer=fuzz.token_sort_ratio)
        if m2 and m2[1] >= 80:
            matched = m2[0]
            for main, alias_list in aliases.items():
                if matched in alias_list:
                    return main, conf * 0.9
            return matched, conf * 0.9

        return name, conf * 0.7  # Return raw AI name with reduced confidence


# ---------------------------------------------------------------------------
# PRIVATE EXCEPTIONS
# ---------------------------------------------------------------------------
class _AuthError(Exception):
    pass

class _CreditError(Exception):
    pass
