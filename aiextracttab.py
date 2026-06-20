# aiextracttab.py
# AI Extract Tab for Maafushivaru Hub
# - Trims PDFs to only Receiving Report pages (RC-MAM + PO number)
# - Saves trimmed copies into "TEMP API PDFS" with the SAME filename
# - OCRs each trimmed PDF via OCR.space API
# - Extracts supplier / GRN / invoice / totals using main app helpers
# - "Send to Tabs" pushes results into Rename + Dispatch tabs and renames originals

import os
import re
import json
import queue
import shutil
import logging
import threading
import io
from datetime import datetime
from typing import List, Dict

import tkinter as tk
from tkinter import ttk, messagebox

import fitz  # PyMuPDF

from maafushivaru_hub import (
    PANEL,
    PANEL2,
    TEXT,
    MUTED,
    ACCENT,
    SUCCESS,
    WARNING,
    ERROR,
)

try:
    from ai_supplier_matcher import OCRSpaceExtractor
    OCR_SPACE_AVAILABLE = True
except ImportError:
    OCR_SPACE_AVAILABLE = False
    OCRSpaceExtractor = None


# ---------------------------------------------------------------------------
# PAGE DETECTION: is this page a Receiving Report?
# ---------------------------------------------------------------------------
_RC_MAM_PAT = re.compile(r"RC[-\s]*MAM[-\s]*\d{3,}", re.IGNORECASE)
_PO_NUM_PAT = re.compile(
    r"(?:PURCHASE\s*ORDER|P\.?\s*O\.?)\s*(?:NO\.?|NUMBER|#)?\s*[:\-]?\s*(?:MAM[-\s]?)?\d{3,}",
    re.IGNORECASE,
)


def _page_is_receiving_report(page_text: str) -> bool:
    if not page_text:
        return False
    u = page_text.upper()
    has_rc = bool(_RC_MAM_PAT.search(u))
    has_po = bool(_PO_NUM_PAT.search(u))
    has_record_label = "RECEIVING RECORD" in u or "RECEIVING REPORT" in u
    return (has_rc and has_po) or (has_record_label and has_po)

# ---------------------------------------------------------------------------
# OCR.SPACE USAGE / CREDIT TRACKING  (1 request = 1 page, 2500 free / month)
# ---------------------------------------------------------------------------
_OCR_FREE_MONTHLY_QUOTA = 25000

# After this many results accumulate in the AI Extract result tree, the app
# offers to generate (export) the Excel sheet. Re-offered at every further
# multiple (60, 90, ...). Reset when the user presses "Clear".
_AIX_SHEET_PROMPT_EVERY = 30


def _aix_usage_file_path(app) -> str:
    return os.path.join(app.dirs.get("base", "."), "ocr_usage.json")


def _aix_load_usage(app) -> dict:
    path = _aix_usage_file_path(app)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _aix_save_usage(app, data: dict):
    path = _aix_usage_file_path(app)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        logging.warning(f"[AI EXTRACT] Could not save usage file: {e}")


def _aix_current_month_key() -> str:
    return datetime.now().strftime("%Y-%m")


def _aix_add_usage_pages(app, n_pages: int):
    if n_pages <= 0:
        return
    data = _aix_load_usage(app)
    key = _aix_current_month_key()
    data[key] = int(data.get(key, 0)) + int(n_pages)
    _aix_save_usage(app, data)
    _aix_refresh_usage_label(app)


def _aix_get_usage_this_month(app) -> int:
    data = _aix_load_usage(app)
    return int(data.get(_aix_current_month_key(), 0))


def _aix_refresh_usage_label(app):
    used = _aix_get_usage_this_month(app)
    remaining = max(_OCR_FREE_MONTHLY_QUOTA - used, 0)
    text = f"OCR.space credits — used: {used} / {_OCR_FREE_MONTHLY_QUOTA}   remaining: {remaining}"
    color = SUCCESS
    if remaining <= 0:
        color = ERROR
    elif remaining <= 250:
        color = WARNING
    if hasattr(app, "_aix_usage_var"):
        app._aix_usage_var.set(text)
    if hasattr(app, "_aix_usage_lbl"):
        try:
            app._aix_usage_lbl.configure(fg=color)
        except Exception:
            pass


def _aix_set_api_status(app, text: str, color=None):
    """Live 'what is the API doing right now' indicator."""
    if color is None:
        color = ACCENT
    if hasattr(app, "_aix_api_status_var"):
        app.after(0, lambda: app._aix_api_status_var.set(text))
    if hasattr(app, "_aix_api_status_lbl"):
        app.after(0, lambda: app._aix_api_status_lbl.configure(fg=color))
    _aix_log(app, "API", text)


# ---------------------------------------------------------------------------
# UNIFIED LOGS  (Compress / API / Process / Send / System in one window)
# ---------------------------------------------------------------------------
_AIX_LOG_CATEGORIES = ["COMPRESS", "API", "PROCESS", "SEND", "SYSTEM"]
# Special pseudo-category in the Logs window that reveals a per-PDF dropdown.
_AIX_PDF_LOG_LABEL = "PDF Log"
_AIX_LOG_MAX = 5000

def _aix_log(app, category: str, message: str):
    """Append one entry to the unified AI-Extract log store.

    Thread-safe enough for our use (CPython list.append is atomic). The Logs
    window renders from this store on a timer, so worker threads never touch Tk.
    """
    if not message:
        return
    cat = (category or "SYSTEM").upper()
    entry = {
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "cat": cat,
        "msg": str(message),
    }
    store = getattr(app, "_aix_logs", None)
    if store is None:
        store = app._aix_logs = []
    store.append(entry)
    if len(store) > _AIX_LOG_MAX:
        del store[: len(store) - _AIX_LOG_MAX]
    try:
        logging.info(f"[AIX:{cat}] {entry['msg'].splitlines()[0]}")
    except Exception:
        pass

def _aix_clog(app, message: str):
    """Compress/Trim log: keep the legacy buffer AND feed the unified Logs."""
    store = getattr(app, "_aix_compress_log", None)
    if store is None:
        store = app._aix_compress_log = []
    store.append(message)
    _aix_log(app, "COMPRESS", message)

def _aix_log_color(cat: str):
    return {
        "COMPRESS": WARNING,
        "API": ACCENT,
        "PROCESS": TEXT,
        "SEND": SUCCESS,
        "SYSTEM": MUTED,
    }.get((cat or "").upper(), TEXT)
# ---------------------------------------------------------------------------
# PDF SIZE COMPRESSION (target < 1 MB so OCR.space free tier accepts it)
# ---------------------------------------------------------------------------
def _aix_compress_pdf_to_size(src_path: str, dest_path: str, target_mb: float = 1.0) -> bool:
    """
    Copies/compresses src_path -> dest_path so the result is <= target_mb.
    Returns True if compression was actually needed/applied.
    Returns False if the file was already small enough (plain copy).
    """
    from PIL import Image

    target_bytes = target_mb * 1024 * 1024
    src_size = os.path.getsize(src_path)

    if src_size <= target_bytes:
        shutil.copy2(src_path, dest_path)
        return False

    # --- Pass 1: lossless cleanup (strip junk, recompress streams) ---
    try:
        doc = fitz.open(src_path)
        doc.save(dest_path, garbage=4, deflate=True, deflate_images=True, clean=True)
        doc.close()
    except Exception as e:
        logging.warning(f"[AI EXTRACT] lossless compress failed [{src_path}]: {e}")
        shutil.copy2(src_path, dest_path)

    if os.path.getsize(dest_path) <= target_bytes:
        return True

    # --- Pass 2: rasterize pages at decreasing quality until under target ---
    for scale, quality in [(1.6, 75), (1.3, 65), (1.0, 55), (0.8, 45), (0.6, 35)]:
        try:
            doc = fitz.open(src_path)
            out = fitz.open()
            for page in doc:
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                buf = io.BytesIO()
                img.save(buf, format="JPEG", quality=quality)
                jpeg_bytes = buf.getvalue()

                rect = page.rect
                new_page = out.new_page(width=rect.width, height=rect.height)
                new_page.insert_image(rect, stream=jpeg_bytes)

            out.save(dest_path, garbage=4, deflate=True)
            out.close()
            doc.close()
        except Exception as e:
            logging.error(f"[AI EXTRACT] rasterize compress failed [{src_path}] @ scale {scale}: {e}")
            continue

        if os.path.getsize(dest_path) <= target_bytes:
            break

    return True

# ---------------------------------------------------------------------------
# TAB BUILDER
# ---------------------------------------------------------------------------
def add_ai_extract_tab(app):
    app._aix_results = []
    app._aix_row_map = {}
    app._aix_all_rows = []
    app._aix_queue = queue.Queue()
    app._aix_running = False
    app._aix_temp_folder = os.path.join(app.dirs.get("base", "."), "TEMP API PDFS")
    app._aix_compress_log = []       # legacy trim/compress buffer (still used)
    app._aix_logs = []               # unified log store (compress/api/process/send)
    app._aix_logs_window = None
    app._aix_logs_text = None
    # Per-PDF debug store (raw OCR text + parsed fields), shown in the "PDF Log"
    # view of the Logs window. Temporary - wiped on Clear.
    if not hasattr(app, "_aix_pdf_logs"):
        app._aix_pdf_logs = {}
    # Next result count at which the "generate sheet now?" dialog should fire.
    app._aix_next_sheet_prompt = _AIX_SHEET_PROMPT_EVERY

    frame = app._make_tab("AI Extract")
    frame.configure(style="TFrame")

    ctrl_bar = tk.Frame(frame, bg=PANEL2, height=52)
    ctrl_bar.pack(fill=tk.X)
    ctrl_bar.pack_propagate(False)

    app._aix_btn_scan = ttk.Button(
        ctrl_bar,
        text="✂  Scan & Trim PDFs",
        style="Accent.TButton",
        command=lambda: _aix_start_scan_trim(app),
    )
    app._aix_btn_scan.pack(side=tk.LEFT, padx=(16, 6), pady=8)

    app._aix_btn_process = ttk.Button(
        ctrl_bar,
        text="🤖  Process (OCR.space API)",
        style="Success.TButton",
        command=lambda: _aix_start_process(app),
    )
    app._aix_btn_process.pack(side=tk.LEFT, padx=4, pady=8)

    app._aix_btn_send = ttk.Button(
        ctrl_bar,
        text="📤  Send to Tabs",
        command=lambda: _aix_send_to_tabs(app),
    )
    app._aix_btn_send.pack(side=tk.LEFT, padx=4, pady=8)

    ttk.Button(
        ctrl_bar,
        text="🗑  Clear",
        command=lambda: _aix_clear(app),
    ).pack(side=tk.LEFT, padx=4, pady=8)

    ttk.Button(
        ctrl_bar,
        text="📊  Export Excel",
        command=lambda: _aix_export_excel(app),
    ).pack(side=tk.LEFT, padx=4, pady=8)

    ttk.Button(
        ctrl_bar,
        text="📂  Open TEMP Folder",
        command=lambda: app._open_folder(app._aix_temp_folder),
    ).pack(side=tk.LEFT, padx=4, pady=8)

    # --- Compress Log button ---
    app._aix_btn_compress_log = ttk.Button(
        ctrl_bar,
        text="📋  Logs",
        command=lambda: _aix_show_logs(app),
    )
    app._aix_btn_compress_log.pack(side=tk.LEFT, padx=4, pady=8)

    # --- Debug button ---
    ttk.Button(
        ctrl_bar,
        text="🔧  Debug OCR",
        command=lambda: _aix_open_debug_window(app),
    ).pack(side=tk.LEFT, padx=4, pady=8)
    
    # --- OCR.space engine selector ---
    tk.Label(
        ctrl_bar,
        text="OCR Engine:",
        bg=PANEL2,
        fg=TEXT,
        font=("Segoe UI", 9),
    ).pack(side=tk.LEFT, padx=(12, 2), pady=8)

    _ocr_engine_options = [
        "Engine 1 (Default)",
        "Engine 2 (Enhanced)",
        "Engine 3 (Extra Accurate)",
    ]
    app._aix_ocr_engine_var = tk.StringVar(value=_ocr_engine_options[1])  # default Engine 2

    # Pre-select based on what's in config
    _cfg_engine = app.cfg.get("ocr_space", {}).get("OCREngine", 2)
    if _cfg_engine == 1:
        app._aix_ocr_engine_var.set(_ocr_engine_options[0])
    elif _cfg_engine == 3:
        app._aix_ocr_engine_var.set(_ocr_engine_options[2])
    else:
        app._aix_ocr_engine_var.set(_ocr_engine_options[1])

    app._aix_ocr_engine_combo = ttk.Combobox(
        ctrl_bar,
        textvariable=app._aix_ocr_engine_var,
        values=_ocr_engine_options,
        state="readonly",
        width=22,
    )
    app._aix_ocr_engine_combo.pack(side=tk.LEFT, padx=(0, 8), pady=8)

    # --- Status badge ---
    badge_text  = "OCR.space: ready"   if OCR_SPACE_AVAILABLE else "OCR.space: MISSING"
    badge_color = SUCCESS              if OCR_SPACE_AVAILABLE else ERROR
    tk.Label(
        ctrl_bar,
        text=badge_text,
        bg=PANEL2,
        fg=badge_color,
        font=("Segoe UI", 9, "bold"),
    ).pack(side=tk.RIGHT, padx=16, pady=8)

    prog_bar = tk.Frame(frame, bg=PANEL, height=6)
    prog_bar.pack(fill=tk.X)
    app._aix_progress = ttk.Progressbar(prog_bar, orient="horizontal", mode="determinate")
    app._aix_progress.pack(fill=tk.X)
    # --- Live API status + usage/credit bar ---
    status_bar2 = tk.Frame(frame, bg=PANEL, height=28)
    status_bar2.pack(fill=tk.X)
    status_bar2.pack_propagate(False)

    app._aix_api_status_var = tk.StringVar(value="API status: idle")
    app._aix_api_status_lbl = tk.Label(
        status_bar2,
        textvariable=app._aix_api_status_var,
        bg=PANEL,
        fg=MUTED,
        font=("Segoe UI", 9, "bold"),
    )
    app._aix_api_status_lbl.pack(side=tk.LEFT, padx=16, pady=4)

    app._aix_usage_var = tk.StringVar(value="OCR.space credits — loading...")
    app._aix_usage_lbl = tk.Label(
        status_bar2,
        textvariable=app._aix_usage_var,
        bg=PANEL,
        fg=MUTED,
        font=("Segoe UI", 9, "bold"),
    )
    app._aix_usage_lbl.pack(side=tk.RIGHT, padx=16, pady=4)

    hdr_bar = tk.Frame(frame, bg=PANEL)
    hdr_bar.pack(fill=tk.X)
    tk.Label(
        hdr_bar,
        text="AI Extract Results",
        bg=PANEL,
        fg=TEXT,
        font=("Segoe UI", 11, "bold"),
    ).pack(side=tk.LEFT, padx=20, pady=(12, 4))
    tk.Label(
        hdr_bar,
        text="1) Scan & Trim  ->  2) Process (OCR.space API)  ->  3) Send to Tabs  ·  Double-click a cell to edit",
        bg=PANEL,
        fg=MUTED,
        font=("Segoe UI", 8),
    ).pack(side=tk.LEFT, padx=8, pady=(12, 4))

    # Live count of how many PDFs are currently in the result tree.
    app._aix_count_var = tk.StringVar(value="0 PDFs in results")
    tk.Label(
        hdr_bar,
        textvariable=app._aix_count_var,
        bg=PANEL,
        fg=ACCENT,
        font=("Segoe UI", 9, "bold"),
    ).pack(side=tk.RIGHT, padx=20, pady=(12, 4))

    tf = tk.Frame(frame, bg=PANEL)
    tf.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))

    cols = ("File", "Date", "Supplier", "Confidence", "PO #", "Invoice #",
            "USD", "MVR", "EUR", "GBP", "SGD", "GRN")
    cw = {
        "File": 190, "Date": 90, "Supplier": 200, "Confidence": 80,
        "PO #": 115, "Invoice #": 130, "USD": 85, "MVR": 85,
        "EUR": 85, "GBP": 85, "SGD": 85, "GRN": 220,
    }

    app._aix_tree = ttk.Treeview(tf, columns=cols, show="headings", height=20)
    for c in cols:
        app._aix_tree.heading(c, text=c, anchor="center")
        app._aix_tree.column(c, width=cw.get(c, 100), anchor="center", stretch=False)

    ys = ttk.Scrollbar(tf, orient="vertical", command=app._aix_tree.yview)
    xs = ttk.Scrollbar(tf, orient="horizontal", command=app._aix_tree.xview)
    app._aix_tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)

    app._aix_tree.grid(row=0, column=0, sticky="nsew")
    ys.grid(row=0, column=1, sticky="ns")
    xs.grid(row=1, column=0, sticky="ew")
    tf.rowconfigure(0, weight=1)
    tf.columnconfigure(0, weight=1)

    app._bind_tree_mousewheel(app._aix_tree)

    app._make_tree_editable(
        app._aix_tree,
        on_edit_callback=lambda r, ci, ov, nv: _aix_on_tree_edit(app, r, ci, ov, nv),
        editable_cols=set(range(1, 12)),
    )
    app._aix_tree._edit_on_preview_cb = lambda row_id: _aix_preview(app, row_id)
    _aix_refresh_usage_label(app)

# ---------------------------------------------------------------------------
# PREVIEW
# ---------------------------------------------------------------------------
def _aix_preview(app, row_id: str):
    result = app._aix_row_map.get(row_id, {})
    path = result.get("temp_path", "")
    if not path or not os.path.exists(path):
        path = result.get("raw_path", "")
    app._show_pdf_preview(path)


# ---------------------------------------------------------------------------
# STEP 1: SCAN & TRIM
# ---------------------------------------------------------------------------
def _aix_start_scan_trim(app):
    if app._aix_running:
        return

    app._aix_running = True
    app._set_buttons_state("disabled")
    app._set_status("AI Extract: scanning & trimming PDFs...", ACCENT)

    def worker():
        kept = 0
        skipped = 0
        app._aix_compress_log.clear()
        _aix_clog(app, 
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Compress/Trim session started."
        )
        try:
            scanned = app.dirs.get("scanned")
            temp = app._aix_temp_folder
            os.makedirs(temp, exist_ok=True)

            files = app._get_pdf_files_strict_order(scanned)
            total = len(files)

            if total == 0:
                app._set_status("SCANNED folder is empty - nothing to trim.", WARNING)
                return

            app.after(0, lambda: app._aix_progress.configure(maximum=max(total, 1), value=0))

            for idx, src in enumerate(files, 1):
                fname = os.path.basename(src)
                app._set_status(f"[AI EXTRACT] Trimming {fname} ({idx}/{total})", ACCENT)

                keep_indices = _aix_find_receiving_pages(app, src)

                if not keep_indices:
                    skipped += 1
                    skip_msg = (
                        f"[{datetime.now().strftime('%H:%M:%S')}]  SKIPPED: {fname} "
                        f"— no receiving-report pages found."
                    )
                    _aix_clog(app, skip_msg)
                    logging.info(f"[AI EXTRACT] No receiving pages found in {fname} - skipped.")
                    app._aix_queue.put(("progress", idx))
                    continue

                dest = os.path.join(temp, fname)
                try:
                    src_size_kb  = os.path.getsize(src)  / 1024
                    _aix_write_trimmed_pdf(src, keep_indices, dest)
                    dest_size_kb = os.path.getsize(dest) / 1024
                    ratio        = (1 - dest_size_kb / src_size_kb) * 100 if src_size_kb > 0 else 0
                    kept += 1

                    log_msg = (
                        f"[{datetime.now().strftime('%H:%M:%S')}]  {fname}\n"
                        f"   Pages kept  : {len(keep_indices)} of {keep_indices}\n"
                        f"   Original    : {src_size_kb:,.1f} KB\n"
                        f"   Trimmed     : {dest_size_kb:,.1f} KB\n"
                        f"   Reduction   : {ratio:.1f}%\n"
                        f"   Saved to    : {dest}\n"
                    )
                    _aix_clog(app, log_msg)
                except Exception as e:
                    skipped += 1
                    err_msg = (
                        f"[{datetime.now().strftime('%H:%M:%S')}]  ERROR trimming {fname}: {e}"
                    )
                    _aix_clog(app, err_msg)
                    logging.error(f"[AI EXTRACT] Trim failed [{fname}]: {e}", exc_info=True)

                app._aix_queue.put(("progress", idx))

            summary = (
                f"[{datetime.now().strftime('%H:%M:%S')}]  Session complete. "
                f"Kept: {kept}   Skipped: {skipped}"
            )
            _aix_clog(app, summary)
            app._set_status(
                f"Trim complete - {kept} file(s) saved to TEMP API PDFS, {skipped} skipped.",
                SUCCESS if skipped == 0 else WARNING,
            )

        except Exception as e:
            logging.error(f"[AI EXTRACT] Scan/trim error: {e}", exc_info=True)
            app._set_status(f"AI Extract scan failed: {e}", ERROR)
        finally:
            app._aix_running = False
            app._aix_queue.put(None)

    threading.Thread(target=worker, daemon=True).start()
    app.after(60, lambda: _aix_poll_queue(app))


def _aix_find_receiving_pages(app, pdf_path: str) -> List[int]:
    keep = []
    try:
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            native = page.get_text("text") or ""
            text = native

            if len(native.strip()) < 50:
                try:
                    rect = page.rect
                    clip = fitz.Rect(0, 0, rect.width, rect.height * 0.40)
                    scale = app.cfg["app_settings"].get("image_scale_factor", 2)
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip)
                    arr = app._preprocess_image(pix)
                    arr = app._correct_rotation(arr)
                    text = app._run_ocr(arr)
                except Exception as e:
                    logging.debug(f"[AI EXTRACT] page OCR failed p{i}: {e}")

            if _page_is_receiving_report(text):
                keep.append(i)
        doc.close()
    except Exception as e:
        logging.error(f"[AI EXTRACT] page scan failed [{pdf_path}]: {e}", exc_info=True)
    return keep


def _aix_write_trimmed_pdf(src_path: str, keep_indices: List[int], dest_path: str):
    src = fitz.open(src_path)
    out = fitz.open()
    for i in keep_indices:
        out.insert_pdf(src, from_page=i, to_page=i)
    out.save(dest_path)
    out.close()
    src.close()


# ---------------------------------------------------------------------------
# STEP 2: PROCESS VIA OCR.SPACE
# ---------------------------------------------------------------------------
def _aix_start_process(app, auto=False):
    """Run OCR.space extraction over every PDF in SCANNED.

    auto=True  -> unattended mode used by the auto-ingest watcher:
                  no blocking message boxes, and on completion the results
                  are pushed straight to the OCR Renamer / GRN Dispatch tabs
                  via the `app._aix_on_complete` one-shot callback.
    """
    if app._aix_running:
        return

    if not OCR_SPACE_AVAILABLE:
        if auto:
            app._set_status(
                "[AUTO-INGEST API] OCR.space extractor unavailable - skipped.", ERROR
            )
            logging.error("[AUTO-INGEST API] OCRSpaceExtractor import failed; cannot process.")
            return
        messagebox.showerror(
            "OCR.space Unavailable",
            "OCRSpaceExtractor could not be imported from ai_supplier_matcher.py.\n"
            "Make sure that file exists and loads without errors.",
        )
        return

    scanned = app.dirs.get("scanned", "")
    temp = app._aix_temp_folder
    os.makedirs(temp, exist_ok=True)

    files = app._get_pdf_files_strict_order(scanned)
    if not files:
        if auto:
            app._set_status("[AUTO-INGEST API] SCANNED empty - nothing to process.", WARNING)
            return
        messagebox.showwarning("AI Extract", "SCANNED folder is empty. Nothing to process.")
        return

    # Results are PERSISTENT and ACCUMULATE across runs (manual or auto-ingest);
    # they are only wiped when the user presses "Clear". Because the originals
    # stay in SCANNED until "Send to Tabs" is clicked, skip any file we have
    # already extracted so re-runs (and the folder watcher) never duplicate rows.
    _already = {
        os.path.basename(r.get("raw_path", "") or r.get("file", ""))
        for r in app._aix_results
    }
    files = [f for f in files if os.path.basename(f) not in _already]
    if not files:
        msg = "AI Extract: no new PDFs (all already in the result tree)."
        _aix_set_api_status(app, msg, MUTED)
        app._set_status(msg, MUTED)
        if auto:
            # Still fire the completion hook (e.g. the sheet-generation prompt).
            _cb = getattr(app, "_aix_on_complete", None)
            app._aix_on_complete = None
            if callable(_cb):
                try:
                    _cb()
                except Exception as e:
                    logging.error(f"[AI EXTRACT] on-complete hook failed: {e}", exc_info=True)
        return

    used_this_month = _aix_get_usage_this_month(app)
    remaining = _OCR_FREE_MONTHLY_QUOTA - used_this_month
    if remaining <= 0:
        if auto:
            logging.warning(
                "[AUTO-INGEST API] Free OCR.space quota reached "
                f"({used_this_month}/{_OCR_FREE_MONTHLY_QUOTA}); proceeding automatically."
            )
        elif not messagebox.askyesno(
            "OCR.space Free Quota Reached",
            f"You have already used {used_this_month} of your {_OCR_FREE_MONTHLY_QUOTA} "
            f"free OCR.space requests this month.\n\nContinue anyway?",
        ):
            return

    # Compression threshold is driven by the configurable OCR.space upload limit
    # (Settings -> OCR.space "Max upload MB"), defaulting to 1.0 MB. Files at or
    # below this size are copied as-is; larger files are compressed first.
    try:
        target_mb = float(app.cfg.get("ocr_space", {}).get("max_upload_mb", 1.0))
    except (TypeError, ValueError):
        target_mb = 1.0
    if target_mb <= 0:
        target_mb = 1.0

    _aix_log(app, "PROCESS",
             f"Run started ({'auto' if auto else 'manual'}): {len(files)} file(s), "
             f"compress threshold {target_mb:.2f} MB.")

    app._aix_running = True
    app._set_buttons_state("disabled")
    app._set_status("AI Extract: preparing & OCR.space processing started...", ACCENT)
    _aix_set_api_status(app, "Preparing files...", ACCENT)

    # NOTE: we intentionally do NOT clear app._aix_results / the result tree
    # here. Results persist and accumulate until the user presses "Clear".
    # Duplicate processing is prevented by the skip-set built above.

    ocr_cfg = dict(app.cfg.get("ocr_space", {}))
    _engine_label = getattr(app, "_aix_ocr_engine_var", None)
    if _engine_label is not None:
        _sel = _engine_label.get()
        if "Engine 1" in _sel:
            ocr_cfg["OCREngine"] = 1
        elif "Engine 3" in _sel:
            ocr_cfg["OCREngine"] = 3
        else:
            ocr_cfg["OCREngine"] = 2

    extractor = OCRSpaceExtractor(config=ocr_cfg)

    def worker():
        n_ok = 0
        n_fail = 0
        total_pages_used = 0
        try:
            total = len(files)
            app.after(0, lambda: app._aix_progress.configure(maximum=max(total, 1), value=0))

            for idx, src_path in enumerate(files, 1):
                fname = os.path.basename(src_path)
                temp_path = os.path.join(temp, fname)

                # --- Prepare the temp copy that will actually be OCR'd ---
                temp_is_fresh = (
                    os.path.exists(temp_path)
                    and os.path.getmtime(temp_path) >= os.path.getmtime(src_path)
                )

                if temp_is_fresh:
                    # Reuse it — likely a manual Scan & Trim output, not stale
                    was_modified = True
                else:
                    app._set_status(f"[AI EXTRACT] Preparing {fname} ({idx}/{total})", ACCENT)
                    try:
                        was_modified = _aix_compress_pdf_to_size(src_path, temp_path, target_mb=target_mb)
                    except Exception as e:
                        logging.error(f"[AI EXTRACT] Prep failed [{fname}]: {e}", exc_info=True)
                        shutil.copy2(src_path, temp_path)
                        was_modified = False

                try:
                    _src_mb = os.path.getsize(src_path) / (1024 * 1024)
                    _tmp_mb = os.path.getsize(temp_path) / (1024 * 1024)
                    if temp_is_fresh:
                        _aix_log(app, "COMPRESS", f"{fname}: reused existing trimmed copy ({_tmp_mb:.2f} MB)")
                    elif was_modified:
                        _aix_log(app, "COMPRESS", f"{fname}: {_src_mb:.2f} MB > {target_mb:.2f} MB limit -> compressed to {_tmp_mb:.2f} MB")
                    else:
                        _aix_log(app, "COMPRESS", f"{fname}: {_src_mb:.2f} MB <= {target_mb:.2f} MB limit -> sent as-is")
                except Exception:
                    pass

                # --- Count pages for usage/credit tracking ---
                try:
                    pdoc = fitz.open(temp_path)
                    n_pages = pdoc.page_count
                    pdoc.close()
                except Exception:
                    n_pages = 1

                # --- OCR via OCR.space ---
                _aix_set_api_status(app, f"Uploading {fname} to OCR.space...", ACCENT)
                app._set_status(f"[AI EXTRACT] OCR.space {fname} ({idx}/{total})", ACCENT)

                try:
                    text, err = extractor.extract_from_file(temp_path)
                except Exception as e:
                    text, err = "", str(e)

                total_pages_used += n_pages
                _aix_add_usage_pages(app, n_pages)

                if err:
                    _aix_set_api_status(app, f"Error on {fname}: {err}", ERROR)
                else:
                    _aix_set_api_status(
                        app, f"OK — {fname} ({n_pages} page(s), {len(text or '')} chars)", SUCCESS
                    )

                text = (text or "").upper()
                fields = _aix_extract_fields_from_text(app, temp_path, text)

                result = {
                    "doc_id": f"{fname}|aix|{idx}",
                    "file": fname,
                    "date": fields["date"],
                    "supplier": fields["supplier"],
                    "po": fields["po"] or "MAM-0000",
                    "invoice": fields["invoice"],
                    "usd": fields["usd"],
                    "mvr": fields["mvr"],
                    "eur": fields["eur"],
                    "gbp": fields["gbp"],
                    "sgd": fields["sgd"],
                    "grn": fields["grn"] or "RC-MAM-0000",
                    "confidence": fields["confidence"],
                    "is_valid": (err == ""),
                    "errors": err,
                    "temp_path": temp_path,
                    "raw_path": src_path,
                    "scan_index": idx,
                    "was_modified": was_modified,
                    "pages_used": n_pages,
                    # Stashed for the "PDF Log" view (moved to app._aix_pdf_logs
                    # on the main thread by the queue poller).
                    "raw_ocr_text": text,
                    "parsed_fields": dict(fields),
                }

                if err == "":
                    n_ok += 1
                else:
                    n_fail += 1
                    logging.warning(f"[AI EXTRACT] OCR.space error [{fname}]: {err}")

                _aix_log(app, "PROCESS",
                         f"{fname}: supplier={result['supplier']}, grn={result['grn']}, "
                         f"po={result['po']}, invoice={result['invoice'] or '-'}, "
                         f"conf={result['confidence']:.0f}%" + (f", ERROR: {err}" if err else ""))
                app._aix_queue.put(("row", result))
                app._aix_queue.put(("progress", idx))

            _aix_set_api_status(
                app,
                f"Done — {n_ok} ok, {n_fail} error(s), {total_pages_used} page(s) used this run.",
                SUCCESS if n_fail == 0 else WARNING,
            )
            app._set_status(
                f"OCR complete - {n_ok} ok, {n_fail} with errors. Click 'Send to Tabs' to apply.",
                SUCCESS if n_fail == 0 else WARNING,
            )
            _aix_log(app, "PROCESS",
                     f"Run complete - {n_ok} ok, {n_fail} error(s), {total_pages_used} page(s) used.")
            app._notify("Maafushivaru - AI Extract Complete", f"{n_ok} processed, {n_fail} errors.")

        except Exception as e:
            logging.error(f"[AI EXTRACT] process error: {e}", exc_info=True)
            app._set_status(f"AI Extract process failed: {e}", ERROR)
            _aix_set_api_status(app, f"Fatal error: {e}", ERROR)
        finally:
            app._aix_running = False
            app._aix_queue.put(None)

    threading.Thread(target=worker, daemon=True).start()
    app.after(60, lambda: _aix_poll_queue(app))

def _aix_extract_fields_from_text(app, pdf_path: str, text: str) -> Dict:
    supplier = _aix_extract_supplier_raw(app, text)
    confidence = 100.0 if supplier else 0.0

    # FAST PATH: OCR.space already returned the text, so parse every field
    # directly from it. The old code called _extract_receiving_report_fields()
    # / _extract_grn_full(), which re-opened the PDF and re-ran LOCAL OCR (twice
    # per page, at the configured scale) - the cause of the slowdown after
    # characters were extracted. No re-OCR happens here anymore.
    rr = app._extract_receiving_report_fields_from_text(text)

    grn = rr.get("grn", "")
    if not grn:
        # Text-based GRN fallback (no PDF re-open, no re-OCR).
        try:
            nums = app._extract_grn_candidates_from_receiving_text((text or "").upper())
            grn = app._build_grn_chain(nums) if nums else ""
        except Exception:
            grn = ""
    grn = grn or "RC-MAM-0000"

    po = rr.get("po", "") or app._extract_po_from_receiving_text(text) or "MAM-0000"
    date_v = rr.get("date", "") or app._extract_date_from_receiving_text(text) or ""
    totals = rr.get("totals", {"USD": "", "MVR": "", "EUR": "", "GBP": "", "SGD": ""})

    invoice = app._extract_invoice(text, supplier_hint=supplier) or ""

    return {
        "supplier": supplier or "UNKNOWN SUPPLIER",
        "confidence": confidence,
        "grn": grn,
        "po": po,
        "date": date_v,
        "invoice": invoice,
        "usd": totals.get("USD", ""),
        "mvr": totals.get("MVR", ""),
        "eur": totals.get("EUR", ""),
        "gbp": totals.get("GBP", ""),
        "sgd": totals.get("SGD", ""),
    }

def _aix_extract_supplier_raw(app, text: str) -> str:
    """
    Extract supplier name from OCR text.

    A single PDF may contain MORE THAN ONE Receiving Report (one per page).
    OCR.space returns those pages separated by a form-feed ("\f"). The supplier
    name is usually printed on every report, so if it is unreadable on one
    report we can still recover it from another ("if one is not visible try to
    grab from the other one").

    Strategy:
      1. Try to match a KNOWN config supplier in the WHOLE text, then in each
         individual page/report block.
      2. If still nothing, fall back to a simple inline pattern on the whole
         text and then on each block.
    """
    if not text:
        return ""

    # ------------------------------------------------------------------
    # Build canonical supplier name list from config (no aliases)
    # ------------------------------------------------------------------
    suppliers_cfg = app.cfg.get("suppliers", [])
    supplier_names = []
    if isinstance(suppliers_cfg, list):
        supplier_names = [str(n).upper() for n in suppliers_cfg if n]
    elif isinstance(suppliers_cfg, dict):
        supplier_names = [str(n).upper() for n in suppliers_cfg.keys() if n]

    # ------------------------------------------------------------------
    # Normalizer: uppercase + collapse spaces + unify company suffixes
    # ------------------------------------------------------------------
    def _norm(v: str) -> str:
        v = v.upper()
        v = re.sub(r"\bPRIVATE\s+LIMITED\b",  "PVT LTD", v)
        v = re.sub(r"\bPRIVATE\s+LTD\b",      "PVT LTD", v)
        v = re.sub(r"\bPVT\.?\s*LTD\.?\b",     "PVT LTD", v)
        v = re.sub(r"\bPTE\.?\s*LTD\.?\b",     "PTE LTD", v)
        v = re.sub(r"\bLIMITED\b",             "LTD",     v)
        v = re.sub(r"[^A-Z0-9 &.()\-]+",       " ",       v)
        v = re.sub(r"\s+",                      " ",       v)
        return v.strip()

    norm_supplier_names = [_norm(s) for s in supplier_names]

    # ------------------------------------------------------------------
    # Match a known config supplier inside ONE block of text.
    # ------------------------------------------------------------------
    def _match_known(block: str) -> str:
        if not block:
            return ""
        lines = [ln.strip() for ln in block.splitlines() if ln.strip()]

        for idx, line in enumerate(lines):
            if not re.search(r"\bSUPPLIER\b", line, re.IGNORECASE):
                continue

            line_before = lines[idx - 1] if idx > 0 else ""
            line_after  = lines[idx + 1] if idx + 1 < len(lines) else ""

            m_inline = re.search(
                r"\b(?:SUPPLIER(?:\s+NAME)?|VENDOR)\s*[:\-]\s*(.+)",
                line,
                re.IGNORECASE,
            )
            same_val = m_inline.group(1).strip() if m_inline else ""

            candidates = []
            if line_before and same_val:
                candidates.append(f"{line_before} {same_val}")
            if same_val:
                candidates.append(same_val)
            if same_val and line_after:
                candidates.append(f"{same_val} {line_after}")
            if line_before:
                candidates.append(line_before)
            if line_after:
                candidates.append(line_after)
            if line_before and same_val and line_after:
                candidates.append(f"{line_before} {same_val} {line_after}")

            # Exact / substring match (normalised)
            for cand in candidates:
                nc = _norm(cand)
                for i, ns in enumerate(norm_supplier_names):
                    if ns and (ns in nc or nc in ns):
                        return supplier_names[i]

            # Fuzzy match via rapidfuzz (if available)
            try:
                from rapidfuzz import process as rf_proc, fuzz as rf_fuzz
                for cand in candidates:
                    nc = _norm(cand)
                    result = rf_proc.extractOne(
                        nc,
                        norm_supplier_names,
                        scorer=rf_fuzz.token_sort_ratio,
                        score_cutoff=75,
                    )
                    if result:
                        try:
                            return supplier_names[norm_supplier_names.index(result[0])]
                        except ValueError:
                            pass
            except Exception:
                pass
        return ""

    # ------------------------------------------------------------------
    # Simple inline pattern fallback inside ONE block of text.
    # ------------------------------------------------------------------
    def _inline_fallback(block: str) -> str:
        if not block:
            return ""
        patterns = [
            r"SUPPLIER\s*[:\-]\s*(.+)",
            r"VENDOR\s*[:\-]\s*(.+)",
            r"SUPPLIER\s+NAME\s*[:\-]\s*(.+)",
        ]
        for line in [ln.strip() for ln in block.splitlines() if ln.strip()]:
            for pat in patterns:
                m = re.search(pat, line, re.IGNORECASE)
                if m:
                    value = m.group(1).strip()
                    value = re.split(
                        r"\s{2,}|(?:INVOICE|DATE|GRN|RC[-\s]*MAM|PO|PURCHASE\s+ORDER)\s*[:\-]?",
                        value,
                        maxsplit=1,
                        flags=re.IGNORECASE,
                    )[0].strip(" :-")
                    if value:
                        return value.upper()
        return ""

    # ------------------------------------------------------------------
    # Try the whole text first, then each report/page block. This is what
    # makes a 2-receiving-report PDF behave like a normal one: the supplier
    # is recovered from whichever report it is readable on.
    # ------------------------------------------------------------------
    blocks = [b for b in text.split("\f") if b.strip()]

    # --- Single Receiving Record: match on the whole text. ---
    if len(blocks) <= 1:
        return _match_known(text) or _inline_fallback(text) or ""

    # --- Multiple Receiving Records in ONE pdf ---
    # Every GRN document starts with "RECEIVING RECORD". A pdf can hold several,
    # each (ideally) naming the supplier. Sometimes the supplier on the FIRST
    # record is unreadable while a LATER record on the same pdf shows it. Two
    # records that share the SAME invoice number belong to the same supplier,
    # so we can safely borrow the name from the readable record.
    per = []  # (supplier, invoice) per record block, in document order
    for b in blocks:
        sup = _match_known(b) or _inline_fallback(b)
        try:
            inv = (app._extract_invoice(b) or "").strip().upper()
        except Exception:
            inv = ""
        per.append((sup, inv))

    # 1) First record already names the supplier -> use it.
    if per[0][0]:
        return per[0][0]

    # 2) First record has no supplier: borrow it from another record on the
    #    same pdf that shares the same invoice number (same supplier).
    first_inv = per[0][1]
    if first_inv:
        for sup, inv in per[1:]:
            if sup and inv and inv == first_inv:
                return sup

    # 3) Otherwise take the supplier from the first record that names one
    #    (document order - earliest GRN wins).
    for sup, _inv in per:
        if sup:
            return sup

    return ""

# ---------------------------------------------------------------------------
# STEP 3: SEND TO TABS
# ---------------------------------------------------------------------------
def _aix_send_to_tabs(app, auto=False):
    if app._aix_running:
        if auto:
            return
        messagebox.showinfo("AI Extract", "Please wait for the current task to finish.")
        return

    if not app._aix_results:
        if auto:
            app._set_status("[AUTO-INGEST API] No results to send to tabs.", WARNING)
            return
        messagebox.showwarning("AI Extract", "No results to send. Run 'Process' first.")
        return

    scanned = app.dirs.get("scanned", "")
    processed = app.dirs.get("processed", "")
    os.makedirs(processed, exist_ok=True)

    n_renamed = 0
    n_missing = 0

    # Iterate in the EXACT order shown in the result tree so the OCR Renamer
    # and GRN Dispatch tabs end up sorted identically to the AI Extract tree.
    ordered = []
    for _rid in app._aix_tree.get_children():
        _r = app._aix_row_map.get(_rid)
        if _r is not None:
            ordered.append(_r)
    if not ordered:
        ordered = list(app._aix_results)

    for result in ordered:
        # Skip rows already pushed to the tabs so re-clicking "Send to Tabs"
        # (after more files were auto-ingested) never double-adds them.
        if result.get("sent"):
            continue
        fname = result.get("file", "")
        supplier = (result.get("supplier") or "UNKNOWN SUPPLIER").strip().upper()
        grn = (result.get("grn") or "RC-MAM-0000").strip().upper()
        inv_raw = (result.get("invoice") or "").strip()

        inv_part = f"IN {inv_raw}" if inv_raw and inv_raw.upper() != "NO-INVOICE" else "NO-INVOICE"
        new_name = app._safe_filename(f"{supplier} GRN {grn} {inv_part}.pdf")

        src = os.path.join(scanned, fname)
        if not os.path.exists(src):
            alt = os.path.join(processed, fname)
            src = alt if os.path.exists(alt) else src

        if not os.path.exists(src):
            n_missing += 1
            logging.warning(f"[AI EXTRACT] Original not found for {fname}")
            continue

        dest = os.path.join(processed, new_name)
        base, ext = os.path.splitext(dest)
        cnt = 1
        while os.path.exists(dest) and os.path.abspath(dest) != os.path.abspath(src):
            dest = f"{base}_{cnt}{ext}"
            cnt += 1

        try:
            shutil.move(src, dest)
            n_renamed += 1
            _aix_log(app, "SEND", f"{fname} -> {os.path.basename(dest)}")
        except Exception as e:
            logging.error(f"[AI EXTRACT] Rename failed [{fname}]: {e}", exc_info=True)
            n_missing += 1
            continue

        rename_result = {
            "doc_id": result["doc_id"],
            "file": os.path.basename(dest),
            "new_name": os.path.basename(dest),
            "supplier": supplier,
            "grn": grn,
            "invoice": inv_raw if inv_raw.upper() != "NO-INVOICE" else "",
            "invoice_dispatch": inv_raw if inv_raw.upper() != "NO-INVOICE" else "",
            "date": result.get("date", ""),
            "po": result.get("po", "") or "MAM-0000",
            "usd": result.get("usd", ""),
            "mvr": result.get("mvr", ""),
            "eur": result.get("eur", ""),
            "gbp": result.get("gbp", ""),
            "sgd": result.get("sgd", ""),
            "status": "success",
            "dest_path": dest,
            "duplicate_warning": "",
            "confidence": result.get("confidence", 0.0),
        }
        app._rename_results.append(rename_result)
        app._add_rename_tree_row(rename_result)

        dispatch_result = app._build_dispatch_from_rename_result(rename_result, result.get("scan_index", 0))
        dispatch_result["raw_path"] = dest
        app._dispatch_results.append(dispatch_result)
        app._add_dispatch_tree_row(dispatch_result)

        # Mark this AI Extract result as already dispatched.
        result["sent"] = True

    app._refresh_dashboard_stats()
    _aix_log(app, "SEND", f"Sent to tabs - {n_renamed} renamed, {n_missing} missing.")
    app._set_status(
        f"Sent to tabs - {n_renamed} renamed, {n_missing} missing.",
        SUCCESS if n_missing == 0 else WARNING,
    )

    # Processing is complete and originals are renamed -> clear the temporary
    # OCR upload copies from TEMP API PDFS.
    if n_renamed > 0:
        _aix_cleanup_temp(app)

    if auto:
        app._notify(
            "Maafushivaru - Auto-Ingest (API) Complete",
            f"{n_renamed} renamed and dispatched, {n_missing} missing.",
        )
        return
    messagebox.showinfo(
        "Send to Tabs",
        f"Done.\n\nRenamed: {n_renamed}\nMissing originals: {n_missing}\n\nCheck the OCR Renamer and GRN Dispatch tabs.",
    )


def _aix_cleanup_temp(app):
    """Delete every PDF in the TEMP API PDFS folder once processing is complete
    and the originals have been renamed/dispatched. Originals in SCANNED and
    PROCESSED are never touched. Controlled by the 'delete_temp_after_send'
    setting (default on)."""
    if not app.cfg.get("app_settings", {}).get("delete_temp_after_send", True):
        return
    temp = getattr(app, "_aix_temp_folder", "")
    if not temp or not os.path.isdir(temp):
        return
    removed = 0
    for fn in os.listdir(temp):
        if fn.lower().endswith(".pdf"):
            try:
                os.remove(os.path.join(temp, fn))
                removed += 1
            except Exception as e:
                logging.warning(f"[AI EXTRACT] Could not delete temp PDF {fn}: {e}")
    if removed:
        _aix_log(app, "SYSTEM", f"TEMP API PDFS cleaned - deleted {removed} temp PDF(s) after send.")
        app._set_status(f"TEMP API PDFS cleaned ({removed} file(s) removed).", SUCCESS)


# ---------------------------------------------------------------------------
# TREE
# ---------------------------------------------------------------------------
def _aix_add_row(app, result: Dict):
    iv = result.get("is_valid", False)
    conf = result.get("confidence", 0.0)
    conf_thr = app.cfg.get("app_settings", {}).get("confidence_warn_threshold", 80)
    sup = result.get("supplier", "")
    was_modified = result.get("was_modified", False)

    if not iv:
        tag = "invalid"
    elif was_modified:
        tag = "modified"          # was compressed or pre-trimmed — flagged orange
    elif sup == "UNKNOWN SUPPLIER" or conf <= 0:
        tag = "unknown"
    elif conf < conf_thr:
        tag = "low_conf"
    else:
        tag = "valid"

    for t, fg in [
        ("valid", SUCCESS),
        ("invalid", ERROR),
        ("unknown", WARNING),
        ("low_conf", WARNING),
        ("modified", "#F97316"),   # orange
    ]:
        app._aix_tree.tag_configure(t, foreground=fg)

    row_id = app._aix_tree.insert(
        "",
        "end",
        tags=(tag,),
        values=(
            result.get("file", ""),
            result.get("date", ""),
            sup,
            app._conf_display(conf),
            result.get("po", "") or "MAM-0000",
            result.get("invoice", ""),
            result.get("usd", ""),
            result.get("mvr", ""),
            result.get("eur", ""),
            result.get("gbp", ""),
            result.get("sgd", ""),
            result.get("grn", ""),
        ),
    )
    app._aix_row_map[row_id] = result
    app._aix_all_rows.append(row_id)
    _aix_update_count(app)


def _aix_update_count(app):
    """Refresh the 'N PDFs in results' label in the AI Extract header."""
    var = getattr(app, "_aix_count_var", None)
    if var is None:
        return
    try:
        n = len(app._aix_tree.get_children())
    except Exception:
        n = len(getattr(app, "_aix_results", []) or [])
    var.set(f"{n} PDF{'' if n == 1 else 's'} in results")


def _aix_on_tree_edit(app, row_id, col_index, old_val, new_val):
    vals = list(app._aix_tree.item(row_id, "values"))

    if col_index == 11:
        nv = str(new_val).strip().upper()
        if nv and not nv.startswith("RC-MAM-"):
            nv = "RC-MAM-" + nv
        new_val = nv

    if col_index == 5:
        raw = str(new_val).strip()
        if raw.upper().startswith("IN "):
            raw = raw[3:].strip()
        new_val = raw

    if col_index == 2:
        new_val = str(new_val).strip().upper()

    vals[col_index] = new_val
    app._aix_tree.item(row_id, values=vals)

    result = app._aix_row_map.get(row_id)
    if result:
        keys = ["file", "date", "supplier", "_conf", "po", "invoice",
                "usd", "mvr", "eur", "gbp", "sgd", "grn"]
        if col_index < len(keys) and keys[col_index] != "_conf":
            result[keys[col_index]] = new_val

        # If this document was already sent to the OCR Renamer / GRN Dispatch
        # tabs, propagate the edit there and rename the file on disk to match.
        propagated = False
        try:
            propagated = app._apply_ai_extract_edit(result)
        except Exception as e:
            logging.error(f"[AI EXTRACT] propagate edit failed: {e}", exc_info=True)

        if propagated:
            # Reflect the (possibly) renamed file back in the AI Extract row too.
            try:
                rid, rr = app._find_rename_row_by_doc_id(result.get("doc_id", ""))
                if rr and rr.get("file"):
                    result["file"] = rr["file"]
                    rvals = list(app._aix_tree.item(row_id, "values"))
                    rvals[0] = rr["file"]
                    app._aix_tree.item(row_id, values=rvals)
            except Exception:
                pass
            _aix_log(app, "SEND",
                     f"Edit synced to tabs: {result.get('file','')} "
                     f"(supplier={result.get('supplier','')}, grn={result.get('grn','')}, "
                     f"invoice={result.get('invoice','') or '-'})")
            app._set_status("AI Extract edit applied + synced to OCR Renamer / GRN Dispatch.", SUCCESS)
        else:
            app._set_status("AI Extract cell updated.")
    else:
        app._set_status("AI Extract cell updated.")


def _aix_clear(app):
    app._aix_results = []
    app._aix_row_map.clear()
    app._aix_all_rows.clear()
    for i in app._aix_tree.get_children():
        app._aix_tree.delete(i)
    # Reset the "generate sheet now?" milestone.
    app._aix_next_sheet_prompt = _AIX_SHEET_PROMPT_EVERY
    # Clearing AI Extract also clears the OCR Renamer and GRN Dispatch tabs.
    try:
        app._clear_rename_results()
    except Exception as e:
        logging.error(f"[AI EXTRACT] clear rename failed: {e}", exc_info=True)
    try:
        app._clear_dispatch_results()
    except Exception as e:
        logging.error(f"[AI EXTRACT] clear dispatch failed: {e}", exc_info=True)
    # Also wipe ALL logs (unified log store, compress buffer and per-PDF logs).
    app._aix_logs = []
    if getattr(app, "_aix_compress_log", None) is not None:
        app._aix_compress_log.clear()
    if getattr(app, "_aix_pdf_logs", None) is not None:
        app._aix_pdf_logs.clear()
    # Force the Logs window (if open) to redraw and refresh the PDF dropdown.
    app._aix_logs_last_sig = None
    try:
        _aix_refresh_pdf_log_combo(app)
    except Exception:
        pass
    _aix_update_count(app)
    app._set_status("Cleared AI Extract, OCR Renamer, GRN Dispatch results and all logs.")


def _aix_export_excel(app):
    """Export the CURRENT AI Extract results to a brand-new Excel sheet.
    Each call writes a new timestamped GRN_OUTPUT_*.xlsx (never overwrites)."""
    if not app._aix_results:
        messagebox.showwarning("AI Extract", "No results to export. Run 'Process' first.")
        return
    # Use the live tree order so the sheet matches what the user sees.
    rows = []
    for rid in app._aix_tree.get_children():
        r = app._aix_row_map.get(rid)
        if r:
            rows.append(r)
    if not rows:
        rows = list(app._aix_results)
    try:
        out = app._write_grn_excel(rows)
    except Exception as e:
        logging.error(f"[AI EXTRACT] Excel export failed: {e}", exc_info=True)
        messagebox.showerror("Export Failed", f"Could not export Excel:\n\n{e}")
        return
    _aix_log(app, "SYSTEM", f"Excel exported ({len(rows)} rows) -> {os.path.basename(out)}")
    app._set_status(f"Excel exported: {out}", SUCCESS)
    app._notify("Maafushivaru — Export Complete", f"{os.path.basename(out)} saved ({len(rows)} rows).")
    try:
        if messagebox.askyesno("Export Successful", f"Saved:\n{out}\n\nOpen now?"):
            os.startfile(out)
    except Exception:
        pass


def _aix_maybe_prompt_sheet(app):
    """When the result tree reaches a 30-result milestone, offer to generate
    the Excel sheet now. Re-offered at 60, 90, ... and reset on Clear."""
    n = len(app._aix_results)
    nxt = getattr(app, "_aix_next_sheet_prompt", _AIX_SHEET_PROMPT_EVERY)
    if n < nxt:
        return
    # Advance to the next milestone so we don't re-prompt for the same batch.
    app._aix_next_sheet_prompt = ((n // _AIX_SHEET_PROMPT_EVERY) + 1) * _AIX_SHEET_PROMPT_EVERY
    # Informational only - no export button here. The user generates the sheet
    # whenever they like with the "Export Excel" button on the AI Extract tab.
    messagebox.showinfo(
        "AI Extract",
        f"{n} PDFs have been processed and are ready in the results.\n\n"
        f"You can generate the Excel sheet now — just click "
        f"\"📊 Export Excel\" on the AI Extract tab whenever you're ready.",
    )
    
# ---------------------------------------------------------------------------
# COMPRESS LOG VIEWER
# ---------------------------------------------------------------------------
def _aix_render_logs(app):
    """(Re)draw the Logs window from the unified log store, honoring the current
    category filter. Runs on a timer while the window is open so entries added by
    worker threads appear automatically (threads never touch Tk directly)."""
    win = getattr(app, "_aix_logs_window", None)
    txt = getattr(app, "_aix_logs_text", None)
    if not win or not txt:
        return
    try:
        if not win.winfo_exists():
            app._aix_logs_window = None
            return
    except Exception:
        app._aix_logs_window = None
        return

    flt = getattr(app, "_aix_logs_filter_var", None)
    selected = flt.get() if flt else "ALL"

    # ---- PDF Log view: show one selected PDF's extracted data + fields ----
    if selected == _AIX_PDF_LOG_LABEL:
        pdf_logs = getattr(app, "_aix_pdf_logs", {}) or {}
        pdf_name = ""
        pv = getattr(app, "_aix_pdf_log_var", None)
        if pv is not None:
            pdf_name = pv.get()
        sig = ("PDFLOG", pdf_name, len(pdf_logs))
        if getattr(app, "_aix_logs_last_sig", None) == sig:
            app._aix_logs_after_id = win.after(800, lambda: _aix_render_logs(app))
            return
        app._aix_logs_last_sig = sig

        txt.configure(state="normal")
        txt.delete("1.0", tk.END)
        if not pdf_logs:
            txt.insert("1.0", "No PDF logs yet. Run 'Process' first.")
        elif not pdf_name or pdf_name not in pdf_logs:
            txt.insert("1.0", "Select a PDF from the dropdown above to view its log.")
        else:
            txt.insert(tk.END, _aix_format_pdf_log(pdf_name, pdf_logs[pdf_name]))
        txt.configure(state="disabled")
        txt.see("1.0")
        app._aix_logs_after_id = win.after(800, lambda: _aix_render_logs(app))
        return

    logs = getattr(app, "_aix_logs", []) or []
    if selected and selected != "ALL":
        logs = [e for e in logs if e.get("cat") == selected]

    sig = (selected, len(logs))
    if getattr(app, "_aix_logs_last_sig", None) == sig:
        app._aix_logs_after_id = win.after(800, lambda: _aix_render_logs(app))
        return
    app._aix_logs_last_sig = sig

    txt.configure(state="normal")
    txt.delete("1.0", tk.END)
    if not logs:
        txt.insert("1.0", "No log entries yet.")
    else:
        for e in logs:
            txt.insert(tk.END, f"[{e['ts']}] [{e['cat']}] {e['msg']}\n", e["cat"])
    txt.configure(state="disabled")
    txt.see(tk.END)
    app._aix_logs_after_id = win.after(800, lambda: _aix_render_logs(app))


def _aix_format_pdf_log(pdf_name: str, info: dict) -> str:
    """Build the debug-style text block for a single PDF's stored log."""
    fields = info.get("fields", {}) or {}
    raw = info.get("raw", "") or ""
    err = info.get("err", "") or ""
    lines = [
        f"File           : {pdf_name}",
        f"Pages          : {info.get('pages', '')}",
        f"OCR Error      : {err or 'None'}",
        "",
        "--- Extracted Fields ---",
        f"Supplier       : {fields.get('supplier', '')}",
        f"Confidence     : {fields.get('confidence', 0.0):.1f}%",
        f"GRN            : {fields.get('grn', '')}",
        f"PO #           : {fields.get('po', '')}",
        f"Invoice #      : {fields.get('invoice', '')}",
        f"Date           : {fields.get('date', '')}",
        "",
        "--- Totals ---",
        f"USD            : {fields.get('usd', '')}",
        f"MVR            : {fields.get('mvr', '')}",
        f"EUR            : {fields.get('eur', '')}",
        f"GBP            : {fields.get('gbp', '')}",
        f"SGD            : {fields.get('sgd', '')}",
        "",
        "--- Raw OCR Stats ---",
        f"Char count     : {len(raw)}",
        f"Line count     : {len(raw.splitlines())}",
        "",
        "--- Raw OCR Text ---",
        raw or "(no text returned)",
    ]
    return "\n".join(lines)


def _aix_refresh_pdf_log_combo(app):
    """Repopulate the PDF dropdown in the Logs window from the per-PDF store."""
    combo = getattr(app, "_aix_pdf_log_combo", None)
    if combo is None:
        return
    try:
        if not combo.winfo_exists():
            return
    except Exception:
        return
    names = list((getattr(app, "_aix_pdf_logs", {}) or {}).keys())
    combo["values"] = names
    cur = app._aix_pdf_log_var.get() if getattr(app, "_aix_pdf_log_var", None) else ""
    if names and cur not in names:
        app._aix_pdf_log_var.set(names[0])
    elif not names:
        app._aix_pdf_log_var.set("")


def _set_sig_dirty(app):
    app._aix_logs_last_sig = None

def _aix_show_logs(app):
    """Unified Logs window: Compress, API, Process, Send & System in one place,
    with a category filter and live auto-refresh."""
    existing = getattr(app, "_aix_logs_window", None)
    if existing is not None:
        try:
            if existing.winfo_exists():
                existing.deiconify(); existing.lift(); existing.focus_force()
                return
        except Exception:
            pass

    win = tk.Toplevel(app)
    win.title("Logs")
    win.geometry("860x560")
    win.configure(bg=PANEL)
    win.transient(app)
    app._aix_logs_window = win
    app._aix_logs_last_sig = None

    tk.Label(win, text="📜  Logs", bg=PANEL, fg=TEXT,
             font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
    tk.Label(win,
             text="All AI Extract activity in one place - Compress, API, Process & Send. Updates live.",
             bg=PANEL, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(0, 8))

    filt_bar = tk.Frame(win, bg=PANEL)
    filt_bar.pack(fill=tk.X, padx=16, pady=(0, 6))
    tk.Label(filt_bar, text="Show:", bg=PANEL, fg=TEXT,
             font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 6))
    app._aix_logs_filter_var = tk.StringVar(value="ALL")
    combo = ttk.Combobox(filt_bar, textvariable=app._aix_logs_filter_var,
                         values=["ALL"] + _AIX_LOG_CATEGORIES + [_AIX_PDF_LOG_LABEL],
                         state="readonly", width=14, font=("Segoe UI", 9))
    combo.pack(side=tk.LEFT)

    # Second dropdown - only visible when "PDF Log" is chosen. Lets the user
    # pick an individual PDF and see its extracted data + parsed fields.
    app._aix_pdf_log_label = tk.Label(filt_bar, text="PDF:", bg=PANEL, fg=TEXT,
                                      font=("Segoe UI", 9))
    app._aix_pdf_log_var = tk.StringVar(value="")
    app._aix_pdf_log_combo = ttk.Combobox(
        filt_bar, textvariable=app._aix_pdf_log_var,
        values=[], state="readonly", width=46, font=("Segoe UI", 9),
    )
    app._aix_pdf_log_combo.bind(
        "<<ComboboxSelected>>", lambda _e: (_set_sig_dirty(app), _aix_render_logs(app))
    )

    def _on_filter_changed(_e=None):
        if app._aix_logs_filter_var.get() == _AIX_PDF_LOG_LABEL:
            _aix_refresh_pdf_log_combo(app)
            app._aix_pdf_log_label.pack(side=tk.LEFT, padx=(14, 6))
            app._aix_pdf_log_combo.pack(side=tk.LEFT)
        else:
            try:
                app._aix_pdf_log_combo.pack_forget()
                app._aix_pdf_log_label.pack_forget()
            except Exception:
                pass
        _set_sig_dirty(app)
        _aix_render_logs(app)

    combo.bind("<<ComboboxSelected>>", _on_filter_changed)
    # Start hidden (default filter is ALL).
    app._aix_pdf_log_combo.pack_forget()
    app._aix_pdf_log_label.pack_forget()

    txt_frame = tk.Frame(win, bg=PANEL)
    txt_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))
    txt = tk.Text(txt_frame, bg=PANEL2, fg=TEXT, font=("Consolas", 9), wrap=tk.WORD,
                  relief="flat", borderwidth=0, selectbackground=ACCENT)
    sb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    app._aix_logs_text = txt
    for cat in _AIX_LOG_CATEGORIES:
        txt.tag_configure(cat, foreground=_aix_log_color(cat))

    btn_bar = tk.Frame(win, bg=PANEL)
    btn_bar.pack(fill=tk.X, padx=16, pady=(0, 12))

    def _copy_log():
        win.clipboard_clear()
        win.clipboard_append(txt.get("1.0", tk.END))
        win.update()

    def _clear_log():
        app._aix_logs = []
        if getattr(app, "_aix_compress_log", None) is not None:
            app._aix_compress_log.clear()
        _set_sig_dirty(app)
        _aix_render_logs(app)

    def _on_close():
        try:
            aid = getattr(app, "_aix_logs_after_id", None)
            if aid:
                win.after_cancel(aid)
        except Exception:
            pass
        app._aix_logs_window = None
        app._aix_logs_text = None
        win.destroy()

    ttk.Button(btn_bar, text="📋  Copy", command=_copy_log).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_bar, text="🗑  Clear", command=_clear_log).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_bar, text="Close", command=_on_close).pack(side=tk.RIGHT)
    win.protocol("WM_DELETE_WINDOW", _on_close)

    _aix_render_logs(app)

# Backward-compatible alias (older code may still call this name).
def _aix_show_compress_log(app):
    _aix_show_logs(app)
    try:
        app._aix_logs_filter_var.set("COMPRESS")
        _set_sig_dirty(app)
        _aix_render_logs(app)
    except Exception:
        pass

# ---------------------------------------------------------------------------
# DEBUG WINDOW - Test OCR engines on individual files
# ---------------------------------------------------------------------------
def _aix_open_debug_window(app):
    """
    Debug window: pick a file from TEMP API PDFS, pick an OCR engine,
    run OCR.space, and inspect raw text + parsed fields side by side.
    """
    if not OCR_SPACE_AVAILABLE:
        messagebox.showerror(
            "OCR.space Unavailable",
            "OCRSpaceExtractor is not available. Cannot run debug.",
        )
        return

    temp_folder = app._aix_temp_folder

    # Collect available PDF files
    if os.path.isdir(temp_folder):
        pdf_files = sorted(
            [f for f in os.listdir(temp_folder) if f.lower().endswith(".pdf")]
        )
    else:
        pdf_files = []

    # -----------------------------------------------------------------------
    # Build window
    # -----------------------------------------------------------------------
    win = tk.Toplevel(app)
    win.title("AI Extract - OCR Engine Debugger")
    win.geometry("1020x680")
    win.configure(bg=PANEL)
    win.transient(app)

    # --- Title bar ---
    tk.Label(
        win,
        text="🔧  OCR Engine Debugger",
        bg=PANEL,
        fg=TEXT,
        font=("Segoe UI", 12, "bold"),
    ).pack(anchor="w", padx=16, pady=(14, 2))

    tk.Label(
        win,
        text="Select a trimmed PDF and an engine, then click Run OCR Test.",
        bg=PANEL,
        fg=MUTED,
        font=("Segoe UI", 8),
    ).pack(anchor="w", padx=16, pady=(0, 10))

    # --- Controls row ---
    ctrl = tk.Frame(win, bg=PANEL2, height=52)
    ctrl.pack(fill=tk.X, padx=0)
    ctrl.pack_propagate(False)

    # File dropdown
    tk.Label(
        ctrl,
        text="File:",
        bg=PANEL2,
        fg=TEXT,
        font=("Segoe UI", 9),
    ).pack(side=tk.LEFT, padx=(14, 4), pady=10)

    _debug_file_var = tk.StringVar(
        value=pdf_files[0] if pdf_files else "(no files in TEMP API PDFS)"
    )
    _debug_file_combo = ttk.Combobox(
        ctrl,
        textvariable=_debug_file_var,
        values=pdf_files if pdf_files else ["(no files)"],
        state="readonly",
        width=36,
    )
    _debug_file_combo.pack(side=tk.LEFT, padx=(0, 14), pady=10)

    # Refresh file list button
    def _refresh_file_list():
        if os.path.isdir(temp_folder):
            updated = sorted(
                [f for f in os.listdir(temp_folder) if f.lower().endswith(".pdf")]
            )
        else:
            updated = []
        _debug_file_combo["values"] = updated if updated else ["(no files)"]
        if updated:
            _debug_file_var.set(updated[0])

    ttk.Button(
        ctrl,
        text="↺ Refresh",
        command=_refresh_file_list,
    ).pack(side=tk.LEFT, padx=(0, 14), pady=10)

    # Engine selector
    tk.Label(
        ctrl,
        text="Engine:",
        bg=PANEL2,
        fg=TEXT,
        font=("Segoe UI", 9),
    ).pack(side=tk.LEFT, padx=(0, 4), pady=10)

    _debug_engine_var = tk.StringVar(value="Engine 2 (Enhanced)")
    _debug_engine_combo = ttk.Combobox(
        ctrl,
        textvariable=_debug_engine_var,
        values=[
            "Engine 1 (Default)",
            "Engine 2 (Enhanced)",
            "Engine 3 (Extra Accurate)",
        ],
        state="readonly",
        width=22,
    )
    _debug_engine_combo.pack(side=tk.LEFT, padx=(0, 14), pady=10)

    # Status label (right side of ctrl bar)
    _debug_status_var = tk.StringVar(value="Ready.")
    _debug_status_lbl = tk.Label(
        ctrl,
        textvariable=_debug_status_var,
        bg=PANEL2,
        fg=MUTED,
        font=("Segoe UI", 9),
    )
    _debug_status_lbl.pack(side=tk.RIGHT, padx=14, pady=10)

    # Run button
    _debug_btn = ttk.Button(ctrl, text="▶  Run OCR Test", style="Accent.TButton")
    _debug_btn.pack(side=tk.LEFT, padx=(0, 8), pady=10)

    # -----------------------------------------------------------------------
    # Output area — left: raw OCR text / right: parsed fields
    # -----------------------------------------------------------------------
    pane = tk.Frame(win, bg=PANEL)
    pane.pack(fill=tk.BOTH, expand=True, padx=16, pady=(10, 0))

    # Left panel - raw OCR text
    left = tk.Frame(pane, bg=PANEL)
    left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

    tk.Label(
        left,
        text="Raw OCR Text",
        bg=PANEL,
        fg=MUTED,
        font=("Segoe UI", 9, "bold"),
    ).pack(anchor="w", pady=(0, 4))

    raw_frame = tk.Frame(left, bg=PANEL2)
    raw_frame.pack(fill=tk.BOTH, expand=True)

    raw_text = tk.Text(
        raw_frame,
        bg=PANEL2,
        fg=TEXT,
        font=("Consolas", 9),
        wrap=tk.WORD,
        relief="flat",
        borderwidth=0,
        selectbackground=ACCENT,
    )
    raw_sb = ttk.Scrollbar(raw_frame, orient="vertical", command=raw_text.yview)
    raw_text.configure(yscrollcommand=raw_sb.set)
    raw_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=6, pady=6)
    raw_sb.pack(side=tk.RIGHT, fill=tk.Y)

    # Right panel - parsed fields
    right = tk.Frame(pane, bg=PANEL)
    right.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(6, 0))
    right.configure(width=320)

    tk.Label(
        right,
        text="Parsed Fields",
        bg=PANEL,
        fg=MUTED,
        font=("Segoe UI", 9, "bold"),
    ).pack(anchor="w", pady=(0, 4))

    fields_frame = tk.Frame(right, bg=PANEL2)
    fields_frame.pack(fill=tk.BOTH, expand=True)

    fields_text = tk.Text(
        fields_frame,
        bg=PANEL2,
        fg=TEXT,
        font=("Consolas", 9),
        wrap=tk.WORD,
        relief="flat",
        borderwidth=0,
        selectbackground=ACCENT,
        width=38,
    )
    fields_text.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

    # -----------------------------------------------------------------------
    # Bottom bar
    # -----------------------------------------------------------------------
    bot = tk.Frame(win, bg=PANEL)
    bot.pack(fill=tk.X, padx=16, pady=10)

    def _copy_raw():
        win.clipboard_clear()
        win.clipboard_append(raw_text.get("1.0", tk.END))
        win.update()

    ttk.Button(bot, text="📋  Copy Raw Text", command=_copy_raw).pack(
        side=tk.LEFT, padx=(0, 8)
    )
    ttk.Button(bot, text="Close", command=win.destroy).pack(side=tk.RIGHT)

    # -----------------------------------------------------------------------
    # Run OCR logic
    # -----------------------------------------------------------------------
    def _run_debug():
        fname = _debug_file_var.get()
        if not fname or fname.startswith("("):
            messagebox.showwarning("Debug OCR", "No file selected.", parent=win)
            return

        fpath = os.path.join(temp_folder, fname)
        if not os.path.exists(fpath):
            messagebox.showerror(
                "Debug OCR",
                f"File not found:\n{fpath}",
                parent=win,
            )
            return

        # Determine engine number from selection
        sel = _debug_engine_var.get()
        if "Engine 1" in sel:
            engine_num = 1
        elif "Engine 3" in sel:
            engine_num = 3
        else:
            engine_num = 2

        _debug_btn.configure(state="disabled")
        _debug_status_var.set(f"Running OCR.space Engine {engine_num} on {fname} ...")
        win.update_idletasks()

        # Clear previous output
        raw_text.configure(state="normal")
        raw_text.delete("1.0", tk.END)
        fields_text.configure(state="normal")
        fields_text.delete("1.0", tk.END)

        def _worker():
            try:
                ocr_cfg = dict(app.cfg.get("ocr_space", {}))
                ocr_cfg["OCREngine"] = engine_num

                extractor = OCRSpaceExtractor(config=ocr_cfg)
                text, err = extractor.extract_from_file(fpath)
                text_upper = (text or "").upper()

                # Extract fields
                fields = _aix_extract_fields_from_text(app, fpath, text_upper)

                def _update_ui():
                    # Raw text panel
                    raw_text.configure(state="normal")
                    raw_text.delete("1.0", tk.END)
                    if err:
                        raw_text.insert(tk.END, f"[OCR ERROR]\n{err}\n\n")
                    raw_text.insert(tk.END, text or "(no text returned)")
                    raw_text.configure(state="disabled")

                    # Parsed fields panel
                    fields_text.configure(state="normal")
                    fields_text.delete("1.0", tk.END)
                    lines_out = [
                        f"Engine         : {engine_num}",
                        f"File           : {fname}",
                        f"OCR Error      : {err or 'None'}",
                        "",
                        "--- Extracted Fields ---",
                        f"Supplier       : {fields.get('supplier', '')}",
                        f"Confidence     : {fields.get('confidence', 0.0):.1f}%",
                        f"GRN            : {fields.get('grn', '')}",
                        f"PO #           : {fields.get('po', '')}",
                        f"Invoice #      : {fields.get('invoice', '')}",
                        f"Date           : {fields.get('date', '')}",
                        "",
                        "--- Totals ---",
                        f"USD            : {fields.get('usd', '')}",
                        f"MVR            : {fields.get('mvr', '')}",
                        f"EUR            : {fields.get('eur', '')}",
                        f"GBP            : {fields.get('gbp', '')}",
                        f"SGD            : {fields.get('sgd', '')}",
                        "",
                        "--- Raw OCR Stats ---",
                        f"Char count     : {len(text or '')}",
                        f"Line count     : {len((text or '').splitlines())}",
                    ]
                    fields_text.insert(tk.END, "\n".join(lines_out))
                    fields_text.configure(state="disabled")

                    status_color = MUTED if not err else WARNING
                    _debug_status_lbl.configure(fg=status_color)
                    _debug_status_var.set(
                        f"Done. Engine {engine_num}  |  "
                        f"Supplier: {fields.get('supplier', 'UNKNOWN')}  |  "
                        f"Chars: {len(text or '')}"
                    )
                    _debug_btn.configure(state="normal")

                win.after(0, _update_ui)

            except Exception as ex:
                def _show_err():
                    _debug_status_var.set(f"Error: {ex}")
                    _debug_status_lbl.configure(fg=ERROR)
                    _debug_btn.configure(state="normal")
                    raw_text.configure(state="normal")
                    raw_text.insert(tk.END, f"[EXCEPTION]\n{ex}")
                    raw_text.configure(state="disabled")
                win.after(0, _show_err)

        threading.Thread(target=_worker, daemon=True).start()

    _debug_btn.configure(command=_run_debug)

# ---------------------------------------------------------------------------
# QUEUE POLLING
# ---------------------------------------------------------------------------
def _aix_poll_queue(app):
    try:
        while True:
            item = app._aix_queue.get_nowait()
            if item is None:
                app._set_buttons_state("normal")
                app._refresh_dashboard_stats()
                # One-shot completion hook used by the auto-ingest API watcher
                # to automatically push results to the tabs once OCR finishes.
                cb = getattr(app, "_aix_on_complete", None)
                app._aix_on_complete = None
                if callable(cb):
                    try:
                        cb()
                    except Exception as e:
                        logging.error(f"[AI EXTRACT] on-complete hook failed: {e}", exc_info=True)
                else:
                    # Manual run finished -> offer the sheet at the 30 milestone.
                    try:
                        _aix_maybe_prompt_sheet(app)
                    except Exception as e:
                        logging.error(f"[AI EXTRACT] sheet prompt failed: {e}", exc_info=True)
                return

            kind, data = item

            if kind == "row":
                app._aix_results.append(data)
                # Record per-PDF debug info (raw OCR text + parsed fields) so it
                # can be inspected later from the Logs window "PDF Log" view.
                try:
                    if not hasattr(app, "_aix_pdf_logs"):
                        app._aix_pdf_logs = {}
                    app._aix_pdf_logs[data.get("file", "")] = {
                        "raw": data.get("raw_ocr_text", "") or "",
                        "fields": data.get("parsed_fields", {}) or {},
                        "err": data.get("errors", "") or "",
                        "pages": data.get("pages_used", ""),
                    }
                    _aix_refresh_pdf_log_combo(app)
                except Exception as e:
                    logging.error(f"[AI EXTRACT] pdf-log store failed: {e}", exc_info=True)
                _aix_add_row(app, data)

            elif kind == "progress":
                app._aix_progress.configure(value=data)

    except queue.Empty:
        pass

    if app._aix_running:
        app.after(60, lambda: _aix_poll_queue(app))
