# ai_supplier_matcher.py
# OCR.space OCR + optional AI supplier matcher
# Supported AI providers: OpenAI, custom OpenAI-compatible endpoint
# Default OCR backend: OCR.space

import os
import re
import json
import time
import base64
import fitz
import tempfile
import threading
import urllib.request
import urllib.error
import urllib.parse
import logging
from typing import Optional, Tuple, List, Dict, Callable


# ---------------------------------------------------------------------------
# PROVIDER DEFINITIONS
# ---------------------------------------------------------------------------
PROVIDERS = {
    "openai": {
        "label": "OpenAI GPT",
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models": "gpt-4o-mini",
        "default_model": "gpt-4o-mini",
        "doc_url": "https://platform.openai.com/api-keys",
    },
    "custom": {
        "label": "Custom / Local (OpenAI-compatible)",
        "base_url": "http://localhost:11434/v1/chat/completions",
        "models": ["llama3", "mistral", "phi3"],
        "default_model": "llama3",
        "doc_url": "",
    },
}

# ---------------------------------------------------------------------------
# OCR.SPACE CONFIGURATION
# ---------------------------------------------------------------------------
OCR_SPACE_API_URL = "https://api.ocr.space/parse/image"
OCR_SPACE_API_KEY = "K88109865088957"

OCR_SPACE_DEFAULTS = {
    "api_key": OCR_SPACE_API_KEY,
    "language": "eng",
    "isOverlayRequired": False,
    "detectOrientation": True,
    "scale": True,
    "OCREngine": 2,
    "isTable": False,
    "filetype": "PDF",
    "timeout_seconds": 30,
    "max_upload_mb": 1.0,
}

STATUS_CONNECTED = "connected"
STATUS_DISCONNECTED = "disconnected"
STATUS_LOW_CREDIT = "low_credit"
STATUS_OFFLINE = "offline"


class _AuthError(Exception):
    pass


class _CreditError(Exception):
    pass


class OCRSpaceExtractor:
    def __init__(self, config: Optional[dict] = None, logger_func: Optional[Callable[[str], None]] = None):
        cfg = {**OCR_SPACE_DEFAULTS, **(config or {})}
        self.api_key = cfg.get("api_key", OCR_SPACE_API_KEY)
        self.language = cfg.get("language", "eng")
        self.is_overlay = cfg.get("isOverlayRequired", False)
        self.detect_orient = cfg.get("detectOrientation", True)
        self.scale = cfg.get("scale", True)
        self.ocr_engine = int(cfg.get("OCREngine", 2))
        self.is_table = cfg.get("isTable", False)
        self.filetype = cfg.get("filetype", "PDF")
        self.timeout = int(cfg.get("timeout_seconds", 30))
        self.max_upload_mb = float(cfg.get("max_upload_mb", 1.0))
        self._logger_func = logger_func

    def _log(self, msg: str):
        if self._logger_func:
            try:
                self._logger_func(msg)
            except Exception:
                pass
        logging.info(f"[OCR.SPACE] {msg}")

    def extract_from_file(self, file_path: str) -> Tuple[str, str]:
        try:
            if not os.path.exists(file_path):
                return "", f"File not found: {file_path}"

            upload_path = file_path
            orig_size = os.path.getsize(file_path)
            self._log(f"Preparing OCR upload: {os.path.basename(file_path)} ({orig_size / 1024:.1f} KB)")

            if file_path.lower().endswith(".pdf"):
                upload_path = self._ensure_pdf_below_limit(file_path, self.max_upload_mb)
                upload_size = os.path.getsize(upload_path)
                self._log(f"Upload PDF ready: {os.path.basename(upload_path)} ({upload_size / 1024:.1f} KB)")
                if upload_size > int(self.max_upload_mb * 1024 * 1024):
                    return "", (
                        f"Compressed PDF is still above {self.max_upload_mb:.0f} MB "
                        f"({upload_size / 1024 / 1024:.2f} MB)."
                    )

            with open(upload_path, "rb") as fh:
                file_bytes = fh.read()

            b64 = base64.b64encode(file_bytes).decode("utf-8")
            ext = upload_path.rsplit(".", 1)[-1].upper() if "." in upload_path else self.filetype
            filetype = ext if ext in ("PDF", "PNG", "JPG", "JPEG", "GIF", "BMP", "TIFF") else self.filetype

            payload = self._build_payload(base64_image=b64, filetype=filetype)
            text, err = self._post(payload)

            if upload_path != file_path:
                try:
                    os.remove(upload_path)
                except Exception:
                    pass

            return text, err

        except Exception as e:
            self._log(f"extract_from_file error: {e}")
            return "", str(e)

    def extract_from_url(self, url: str) -> Tuple[str, str]:
        payload = self._build_payload(url=url)
        return self._post(payload)

    def test_connection(self) -> Tuple[bool, str]:
        tiny_png = (
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8"
            "z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg=="
        )
        payload = self._build_payload(base64_image=tiny_png, filetype="PNG")
        text, err = self._post(payload)
        if err:
            return False, f"OCR.space connection failed: {err}"
        return True, "OCR.space connected successfully."

    def _ensure_pdf_below_limit(self, pdf_path: str, max_mb: float) -> str:
        max_bytes = int(max_mb * 1024 * 1024)
        original_size = os.path.getsize(pdf_path)

        if original_size <= max_bytes:
            return pdf_path

        self._log(
            f"PDF exceeds limit ({original_size / 1024 / 1024:.2f} MB). "
            f"Starting compression target < {max_mb:.0f} MB."
        )

        temp_dir = tempfile.mkdtemp(prefix="ocrspace_pdf_")
        best_path = os.path.join(temp_dir, "compressed.pdf")

        try:
            doc = fitz.open(pdf_path)
            doc.save(best_path, garbage=4, deflate=True, clean=True)
            doc.close()
            if os.path.getsize(best_path) <= max_bytes:
                return best_path
        except Exception as e:
            self._log(f"Compression pass 1 failed: {e}")

        attempts = [
            (160, 80),
            (140, 72),
            (120, 65),
            (110, 60),
            (96, 55),
        ]

        for dpi, jpg_quality in attempts:
            try:
                out_path = os.path.join(temp_dir, f"compressed_{dpi}_{jpg_quality}.pdf")
                self._raster_rebuild_pdf(pdf_path, out_path, dpi=dpi, jpg_quality=jpg_quality)
                if os.path.getsize(out_path) <= max_bytes:
                    return out_path
            except Exception as e:
                self._log(f"Compression raster pass failed dpi={dpi}, q={jpg_quality}: {e}")

        raise RuntimeError("Unable to compress PDF below 1 MB without excessive degradation.")

    def _raster_rebuild_pdf(self, src_pdf: str, out_pdf: str, dpi: int = 140, jpg_quality: int = 72):
        src = fitz.open(src_pdf)
        out = fitz.open()
        zoom = dpi / 72.0

        for page in src:
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img_bytes = pix.tobytes("jpg", jpg_quality)
            rect = page.rect
            new_page = out.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=img_bytes)

        out.save(out_pdf, garbage=4, deflate=True, clean=True)
        out.close()
        src.close()

    def _build_payload(self, base64_image: str = "", url: str = "", filetype: str = "") -> bytes:
        fields: Dict[str, str] = {
            "apikey": self.api_key,
            "language": self.language,
            "isOverlayRequired": str(self.is_overlay).lower(),
            "detectOrientation": str(self.detect_orient).lower(),
            "scale": str(self.scale).lower(),
            "OCREngine": str(self.ocr_engine),
            "isTable": str(self.is_table).lower(),
        }

        if url:
            fields["url"] = url

        if base64_image:
            mime_ext = (filetype or self.filetype).lower()
            mime_prefix = "application/pdf" if mime_ext == "pdf" else f"image/{mime_ext}"
            fields["base64Image"] = f"data:{mime_prefix};base64,{base64_image}"

        if filetype or self.filetype:
            fields["filetype"] = (filetype or self.filetype).upper()

        return urllib.parse.urlencode(fields).encode("utf-8")

    def _post(self, payload: bytes) -> Tuple[str, str]:
        req = urllib.request.Request(
            OCR_SPACE_API_URL,
            data=payload,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read().decode("utf-8")
        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode("utf-8")
            except Exception:
                pass

            if e.code in (401, 403):
                return "", f"Auth error (HTTP {e.code}): invalid API key"
            if e.code == 429:
                return "", f"Rate-limit / quota exceeded (HTTP {e.code})"
            return "", f"HTTP {e.code}: {body[:500]}"
        except urllib.error.URLError as e:
            return "", f"Network error: {e.reason}"
        except Exception as e:
            return "", f"Unexpected request error: {e}"

        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            return "", f"Unexpected response (not JSON): {raw[:300]}"

        if data.get("IsErroredOnProcessing", False):
            msgs = []
            for p in data.get("ParsedResults", []):
                em = p.get("ErrorMessage", "Unknown")
                if isinstance(em, list):
                    msgs.extend(str(x) for x in em)
                else:
                    msgs.append(str(em))
            if not msgs:
                msgs = ["Unknown OCR.space processing error"]
            return "", "OCR.space processing error: " + "; ".join(msgs)

        texts: List[str] = []
        for page in data.get("ParsedResults", []):
            t = page.get("ParsedText", "").strip()
            if t:
                texts.append(t)

        # Join pages with a form-feed (\f) instead of a newline so the
        # downstream parser can tell the pages apart. A single PDF may contain
        # several Receiving Reports (one per page); keeping the page boundaries
        # lets the app (a) recover the supplier from a second report when the
        # first is unreadable, and (b) SUM the invoice totals of every report
        # instead of collapsing equal amounts together.
        return "\f".join(texts), ""


class AISupplierMatcher:
    """
    OCR.space extractor + optional AI supplier matcher.
    AI provider is OpenAI or any OpenAI-compatible endpoint.
    """

    def __init__(self, config: dict, logger_func: Optional[Callable[[str], None]] = None):
        self.config = config
        self._status = STATUS_OFFLINE
        self._status_lock = threading.Lock()
        self._last_error: str = ""
        self._logger_func = logger_func
        self.ocr_extractor = OCRSpaceExtractor(config=config.get("ocr_space", {}), logger_func=self._log)

    def _log(self, msg: str):
        if self._logger_func:
            try:
                self._logger_func(msg)
            except Exception:
                pass
        logging.info(f"[AI-MATCHER] {msg}")

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
        provider = cfg.get("provider", "openai")
        if provider == "custom":
            return bool(cfg.get("enabled", False) and cfg.get("custom_base_url", "").strip())
        return bool(cfg.get("enabled", False) and cfg.get("api_key", "").strip())

    def test_connection(self) -> Tuple[str, str]:
        messages: List[str] = []

        ocr_ok, ocr_msg = self.ocr_extractor.test_connection()
        messages.append(f"OCR.space: {ocr_msg}")

        if not self.is_enabled():
            self.status = STATUS_OFFLINE if not ocr_ok else STATUS_CONNECTED
            return self.status, " | ".join(messages)

        try:
            name, conf = self.match_supplier(
                "INVOICE FROM ACME CORP TOTAL USD 100.00",
                candidates=["ACME CORP"],
                aliases={},
            )
            ai_msg = f"AI connected - test match: {name}" if name else "AI connected (no match on test text)"
            messages.append(ai_msg)
            self.status = STATUS_CONNECTED
        except _CreditError:
            self.status = STATUS_LOW_CREDIT
            messages.append("AI: credits exhausted or quota exceeded.")
        except _AuthError:
            self.status = STATUS_DISCONNECTED
            messages.append("AI: authentication failed - check API key.")
        except Exception as e:
            self.status = STATUS_DISCONNECTED
            messages.append(f"AI connection failed: {e}")

        return self.status, " | ".join(messages)

    def extract_text(self, file_path: str = "", url: str = "") -> Tuple[str, str]:
        if file_path:
            return self.ocr_extractor.extract_from_file(file_path)
        if url:
            return self.ocr_extractor.extract_from_url(url)
        return "", "Provide either file_path or url."

    def match_supplier(self, text: str, candidates: List[str], aliases: Dict[str, List[str]]) -> Tuple[Optional[str], float]:
        if not self.is_enabled():
            self._log("AI matching disabled or API key/base URL missing.")
            return None, 0.0

        cfg = self._get_ai_cfg()
        provider = cfg.get("provider", "openai")
        cand_list = "\n".join(f"- {c}" for c in candidates[:80])

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
            raw_response = self._call_openai_compatible(cfg, prompt)
            supplier, confidence = self._parse_response(raw_response, candidates, aliases)
            self._log(f"AI match result: supplier={supplier}, confidence={confidence}")
            return supplier, confidence
        except _CreditError:
            self.status = STATUS_LOW_CREDIT
            return None, 0.0
        except _AuthError:
            self.status = STATUS_DISCONNECTED
            return None, 0.0
        except Exception as e:
            self._last_error = str(e)
            self.status = STATUS_DISCONNECTED
            self._log(f"AI match_supplier failed: {e}")
            return None, 0.0

    def _call_openai_compatible(self, cfg: dict, prompt: str) -> str:
        provider = cfg.get("provider", "openai")
        api_key = cfg.get("api_key", "")
        model = cfg.get("model", PROVIDERS.get(provider, PROVIDERS["openai"])["default_model"])
        base_url = cfg.get("custom_base_url", "").strip() or PROVIDERS.get(provider, {}).get("base_url", "")
        timeout = int(cfg.get("timeout_seconds", 20))

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

    def _parse_response(self, raw: str, candidates: List[str], aliases: Dict[str, List[str]]) -> Tuple[Optional[str], float]:
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            self._log(f"Could not parse AI response: {raw[:300]}")
            return None, 0.0

        direct = self._extract_from_dict(data, candidates, aliases)
        if direct != (None, 0.0):
            return direct

        try:
            content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
            return self._extract_from_json_string(content, candidates, aliases)
        except Exception:
            return None, 0.0

    def _extract_from_json_string(self, content: str, candidates: List[str], aliases: Dict[str, List[str]]) -> Tuple[Optional[str], float]:
        m = re.search(r"\{.*?\}", content, re.DOTALL)
        if not m:
            return None, 0.0
        try:
            obj = json.loads(m.group(0))
            return self._extract_from_dict(obj, candidates, aliases)
        except Exception:
            return None, 0.0

    def _extract_from_dict(self, obj: dict, candidates: List[str], aliases: Dict[str, List[str]]) -> Tuple[Optional[str], float]:
        name = obj.get("supplier")
        conf = float(obj.get("confidence", 0))
        if not name:
            return None, 0.0

        name_upper = str(name).upper().strip()

        for c in candidates:
            if c.upper() == name_upper:
                return c, conf

        for main, alias_list in aliases.items():
            for alias in alias_list:
                if alias.upper() == name_upper:
                    return main, conf
            if main.upper() == name_upper:
                return main, conf

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

        return name, conf * 0.7
