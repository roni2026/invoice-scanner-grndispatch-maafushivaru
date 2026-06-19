# scan_tab.py
# Adds a "Scan" tab to an existing Tkinter app (MaafushivaruHub) using Windows WIA via pywin32.
# Supports: save-as, item type, page sides, page size, color mode, auto-orient, send-to folder, preview.
# Requires: Windows + WIA driver, and: pip install pywin32 pillow pypdf2

import os
import io
import sys
import time
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import warnings
warnings.simplefilter("ignore", Image.DecompressionBombWarning)
Image.MAX_IMAGE_PIXELS = None

from PIL import Image
from PyPDF2 import PdfWriter
try:
    import pythoncom
    import win32com.client  # pywin32
except ImportError:
    pythoncom = None
    win32com = None


# --- WIA constants (subset) ---
WIA_DeviceType_ScannerDeviceType = 1

# Common WIA property IDs (useful subset)
WIA_IPS_PAGE_SIZE           = 3097   # 0=Auto, 1=Letter, 2=Legal, 3=A4...
WIA_IPS_XRES                = 6147   # DPI X
WIA_IPS_YRES                = 6148   # DPI Y
WIA_IPS_CUR_INTENT          = 6146   # 1=Color, 2=Grayscale, 4=Text/BW
WIA_IPS_DOCUMENT_HANDLING_SELECT = 3088  # 1=Feeder, 2=Flatbed, 4=Duplex (combos allowed)
WIA_IPS_DOCUMENT_HANDLING_STATUS = 3087  # status (ADF ready etc.)
WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES = 3086
WIA_IPS_ORIENTATION         = 6151   # 0=Portrait, 1=Landscape, 2=Rot180, 3=Rot270, 4=Auto if driver supports

# Document handling flags
FEEDER  = 0x0001
FLATBED = 0x0002
DUPLEX  = 0x0004

# Page size values (WIA)
WIA_PAGE_AUTO   = 0
WIA_PAGE_LETTER = 1
WIA_PAGE_LEGAL  = 2
WIA_PAGE_A4     = 3

# Intent flags
INTENT_COLOR     = 0x0001
INTENT_GRAYSCALE = 0x0002
INTENT_TEXT      = 0x0004  # often B/W


def _ensure_pywin32():
    if win32com is None or pythoncom is None:
        raise RuntimeError(
            "pywin32 not installed correctly. Run:\n"
            "  pip install pywin32\n"
            "  python -m pywin32_postinstall -install"
        )


def list_wia_scanners():
    _ensure_pywin32()
    pythoncom.CoInitialize()
    try:
        dev_manager = win32com.client.Dispatch("WIA.DeviceManager")
        scanners = []
        for info in dev_manager.DeviceInfos:
            try:
                if info.Type == WIA_DeviceType_ScannerDeviceType:
                    scanners.append((info.DeviceID, info.Properties("Name").Value))
            except Exception:
                pass
        return scanners
    finally:
        pythoncom.CoUninitialize()



def _set_prop(props, pid, val):
    try:
        props.Item(pid).Value = val
    except Exception:
        # Not all devices support all props; ignore silently
        pass


def _acquire_pages(device_id, settings, progress_cb=None, status_cb=None):
    """
    Acquire one or more pages from the selected device using WIA.
    Returns a list of PIL.Image objects.
    """
    _ensure_pywin32()

    pythoncom.CoInitialize()
    try:
        dev_manager = win32com.client.Dispatch("WIA.DeviceManager")
        device = None
        for info in dev_manager.DeviceInfos:
            if info.DeviceID == device_id:
                device = info.Connect()
                break
        if device is None:
            raise RuntimeError("Selected scanner not found/connected.")

        item = device.Items[1]  # WIA uses 1-based index

        dpi = settings.get("dpi", 300)
        page_size = settings.get("page_size", "A4")
        color_mode = settings.get("color_mode", "Color")
        sides = settings.get("sides", "Simplex")
        source = settings.get("source", "ADF")
        auto_orient = settings.get("auto_orient", True)

        ps_map = {"A4": WIA_PAGE_A4, "Letter": WIA_PAGE_LETTER, "Auto": WIA_PAGE_AUTO}
        wia_ps = ps_map.get(page_size, WIA_PAGE_A4)

        if color_mode == "Color":
            intent = INTENT_COLOR
        elif color_mode == "Grayscale":
            intent = INTENT_GRAYSCALE
        else:
            intent = INTENT_TEXT

        doc_flags = 0
        if source == "ADF":
            doc_flags |= FEEDER
        else:
            doc_flags |= FLATBED
        if sides == "Duplex":
            doc_flags |= DUPLEX

        props = item.Properties
        _set_prop(props, WIA_IPS_XRES, dpi)
        _set_prop(props, WIA_IPS_YRES, dpi)
        _set_prop(props, WIA_IPS_PAGE_SIZE, wia_ps)
        _set_prop(props, WIA_IPS_CUR_INTENT, intent)
        if auto_orient:
            _set_prop(props, WIA_IPS_ORIENTATION, 4)

        try:
            _set_prop(device.Properties, WIA_IPS_DOCUMENT_HANDLING_SELECT, doc_flags)
        except Exception:
            _set_prop(props, WIA_IPS_DOCUMENT_HANDLING_SELECT, doc_flags)

        images = []
        page_idx = 0

        while True:
            page_idx += 1
            if status_cb:
                status_cb(f"Scanning page {page_idx}...")

            try:
                imgfile = item.Transfer("{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")
            except Exception as e:
                if page_idx == 1:
                    raise RuntimeError(f"Scan failed: {e}")
                break

            buf = imgfile.FileData.BinaryData
            bio = io.BytesIO(buf)
            try:
                pil = Image.open(bio)
                pil.load()
                images.append(pil)
            except Exception:
                pass

            if source == "Flatbed":
                break

            if progress_cb:
                progress_cb(page_idx)

            time.sleep(0.1)

        if not images:
            raise RuntimeError("No pages acquired (feeder empty?).")

        return images

    finally:
        pythoncom.CoUninitialize()


def _save_images(images, out_path, fmt):
    """
    Save list of PIL images to either a single PDF or multiple PNGs.
    fmt: "PDF" or "PNG"
    """
    if fmt == "PDF":
        # Convert all to RGB (PDF doesn’t support mode "1" directly)
        rgb_pages = []
        for im in images:
            if im.mode in ("RGBA", "P"):
                rgb_pages.append(im.convert("RGB"))
            elif im.mode == "1":
                rgb_pages.append(im.convert("L"))
            else:
                rgb_pages.append(im)
        # Fast path: PIL can save multipage PDF directly
        first, rest = rgb_pages[0], rgb_pages[1:]
        first.save(out_path, "PDF", save_all=True, append_images=rest)
        return [out_path]
    else:
        # Save as separate PNG files with counter
        base, ext = os.path.splitext(out_path)
        if ext.lower() not in (".png",):
            out_path = base + ".png"
        saved = []
        for i, im in enumerate(images, 1):
            p = f"{base}_p{i:02d}.png"
            im.save(p, "PNG", optimize=True)
            saved.append(p)
        return saved


def add_scan_tab(app):
    """
    Mounts a 'Scan' tab into the existing ttk.Notebook on the given app (MaafushivaruHub).
    Expects: app.notebook, app.dirs, app._set_status, app._show_pdf_preview (optional).
    """
    # Styles/colors already defined by the host app; reuse their constants if present
    BG       = getattr(app, "BG", "#0A0F1E") if hasattr(app, "BG") else "#0A0F1E"
    PANEL    = getattr(app, "PANEL", "#111827") if hasattr(app, "PANEL") else "#111827"
    PANEL2   = getattr(app, "PANEL2", "#1F2937") if hasattr(app, "PANEL2") else "#1F2937"
    TEXT     = getattr(app, "TEXT", "#F9FAFB") if hasattr(app, "TEXT") else "#F9FAFB"
    MUTED    = getattr(app, "MUTED", "#9CA3AF") if hasattr(app, "MUTED") else "#9CA3AF"
    ACCENT   = getattr(app, "ACCENT", "#3B82F6") if hasattr(app, "ACCENT") else "#3B82F6"
    SUCCESS  = getattr(app, "SUCCESS", "#10B981") if hasattr(app, "SUCCESS") else "#10B981"
    WARNING  = getattr(app, "WARNING", "#F59E0B") if hasattr(app, "WARNING") else "#F59E0B"
    ERROR    = getattr(app, "ERROR", "#EF4444") if hasattr(app, "ERROR") else "#EF4444"

    # Create tab
    tab = ttk.Frame(app.notebook)
    app.notebook.add(tab, text="  Scan  ")

    # Outer frames
    top = tk.Frame(tab, bg=PANEL2, height=60)
    top.pack(fill=tk.X)
    top.pack_propagate(False)
    body = tk.Frame(tab, bg=PANEL)
    body.pack(fill=tk.BOTH, expand=True)

    tk.Label(top, text="Scan Documents (WIA)", bg=PANEL2, fg=TEXT,
             font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=16, pady=14)

    # Controls frame
    ctrl = tk.Frame(body, bg=PANEL)
    ctrl.pack(fill=tk.X, padx=16, pady=12)

    # Left column (device + options)
    left = tk.Frame(ctrl, bg=PANEL)
    left.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

    # Right column (destination)
    right = tk.Frame(ctrl, bg=PANEL)
    right.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    # Scanner selection
    dev_lbl = tk.Label(left, text="Scanner:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold"))
    dev_lbl.grid(row=0, column=0, sticky="w", pady=(2, 6))
    scanner_var = tk.StringVar(value="")
    scanner_cb = ttk.Combobox(left, textvariable=scanner_var, state="readonly", width=40)
    scanner_cb.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=(2, 6))

    def refresh_scanners():
        try:
            sc = list_wia_scanners()
            names = [f"{name}  |  {did}" for did, name in sc]
            scanner_cb["values"] = names
            if names:
                scanner_cb.current(0)
        except Exception as e:
            messagebox.showerror("WIA Error", f"Could not enumerate scanners:\n\n{e}")

    ttk.Button(left, text="Refresh", command=refresh_scanners).grid(row=0, column=2, padx=8, pady=(2, 6))

    # Item type (document/photo)
    tk.Label(left, text="Item type:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=6)
    item_type_var = tk.StringVar(value="Document")
    ttk.Combobox(left, textvariable=item_type_var, values=["Document", "Photo"], state="readonly", width=20)\
        .grid(row=1, column=1, sticky="w", padx=(8, 0), pady=6)

    # Page sides
    tk.Label(left, text="Page sides:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=6)
    sides_var = tk.StringVar(value="Simplex")
    ttk.Combobox(left, textvariable=sides_var, values=["Simplex", "Duplex"], state="readonly", width=20)\
        .grid(row=2, column=1, sticky="w", padx=(8, 0), pady=6)

    # Page size
    tk.Label(left, text="Page size:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky="w", pady=6)
    page_size_var = tk.StringVar(value="A4")
    ttk.Combobox(left, textvariable=page_size_var, values=["A4", "Letter", "Auto"], state="readonly", width=20)\
        .grid(row=3, column=1, sticky="w", padx=(8, 0), pady=6)

    # DPI
    tk.Label(left, text="DPI:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold")).grid(row=4, column=0, sticky="w", pady=6)
    dpi_var = tk.IntVar(value=300)
    ttk.Spinbox(left, from_=100, to=600, textvariable=dpi_var, width=8)\
        .grid(row=4, column=1, sticky="w", padx=(8, 0), pady=6)

    # Color mode
    tk.Label(left, text="Color mode:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold")).grid(row=5, column=0, sticky="w", pady=6)
    color_var = tk.StringVar(value="Color")
    ttk.Combobox(left, textvariable=color_var, values=["Color", "Grayscale", "Black & White"], state="readonly", width=20)\
        .grid(row=5, column=1, sticky="w", padx=(8, 0), pady=6)

    # Source (ADF/Flatbed)
    tk.Label(left, text="Source:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold")).grid(row=6, column=0, sticky="w", pady=6)
    source_var = tk.StringVar(value="ADF")
    ttk.Combobox(left, textvariable=source_var, values=["ADF", "Flatbed"], state="readonly", width=20)\
        .grid(row=6, column=1, sticky="w", padx=(8, 0), pady=6)

    # Auto orient
    auto_orient_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(left, text="Auto orient (if supported)", variable=auto_orient_var)\
        .grid(row=7, column=1, sticky="w", padx=(6, 0), pady=6)

    # Destination controls
    tk.Label(right, text="Send to folder:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold"))\
        .grid(row=0, column=0, sticky="w", pady=(2, 6))
    out_dir_var = tk.StringVar(value=app.dirs.get("scanned", app.dirs.get("base", ".")))

    out_entry = tk.Entry(right, textvariable=out_dir_var, bg=PANEL2, fg=TEXT,
                         insertbackground=TEXT, relief="flat", bd=0,
                         highlightthickness=1, highlightbackground="#374151",
                         highlightcolor=ACCENT, width=44)
    out_entry.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=(2, 6))

    def choose_folder():
        d = filedialog.askdirectory(initialdir=out_dir_var.get() or app.dirs.get("base", "."))
        if d:
            out_dir_var.set(d)

    ttk.Button(right, text="Browse", command=choose_folder).grid(row=0, column=2, padx=8, pady=(2, 6))

    # Save as
    tk.Label(right, text="Save as:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold"))\
        .grid(row=1, column=0, sticky="w", pady=6)
    save_as_var = tk.StringVar(value="PDF")
    ttk.Combobox(right, textvariable=save_as_var, values=["PDF", "PNG"], state="readonly", width=12)\
        .grid(row=1, column=1, sticky="w", padx=(8, 0), pady=6)

    # File name base
    tk.Label(right, text="File name:", bg=PANEL, fg=MUTED, font=("Segoe UI", 9, "bold"))\
        .grid(row=2, column=0, sticky="w", pady=6)
    fname_var = tk.StringVar(value="SCAN")
    ttk.Entry(right, textvariable=fname_var, width=24).grid(row=2, column=1, sticky="w", padx=(8, 0), pady=6)

    # Show preview after scan
    preview_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(right, text="Show viewer after scan", variable=preview_var)\
        .grid(row=3, column=1, sticky="w", padx=(6, 0), pady=6)

    # Progress + buttons
    foot = tk.Frame(body, bg=PANEL)
    foot.pack(fill=tk.X, padx=16, pady=8)
    pbar = ttk.Progressbar(foot, orient="horizontal", mode="determinate")
    pbar.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 12))
    scan_btn = ttk.Button(foot, text="▶  Scan Now")
    scan_btn.pack(side=tk.LEFT)
    cancel_btn = ttk.Button(foot, text="⏹  Cancel")
    cancel_btn.pack(side=tk.LEFT, padx=(8, 0))

    # Status
    stat = tk.Frame(body, bg=PANEL, height=26)
    stat.pack(fill=tk.X, padx=16, pady=(0, 12))
    status_var = tk.StringVar(value="Ready.")
    tk.Label(stat, textvariable=status_var, bg=PANEL, fg=TEXT, font=("Segoe UI", 9)).pack(side=tk.LEFT)

    # Internal flag
    cancel_flag = {"stop": False}

    def set_status(txt, color=None):
        status_var.set(txt)
        if hasattr(app, "_set_status"):
            try:
                app._set_status(txt, None)
            except Exception:
                pass

    def on_cancel():
        cancel_flag["stop"] = True
        set_status("Cancel requested...", WARNING)

    cancel_btn.configure(command=on_cancel)

    def progress_cb(n):
        pbar.configure(value=n)

    def status_cb(txt):
        set_status(txt)

    def do_scan():
        try:
            cancel_flag["stop"] = False
            pbar.configure(value=0, maximum=10)
            # Device ID from combobox
            sel = scanner_var.get().strip()
            if not sel:
                raise RuntimeError("No scanner selected.")
            # Extract DeviceID from "Name | DeviceID"
            if "  |  " in sel:
                parts = sel.split("  |  ")
                device_id = parts[-1].strip()
            else:
                # fallback: first scanner
                sc = list_wia_scanners()
                if not sc:
                    raise RuntimeError("No WIA scanners found.")
                device_id = sc[0][0]

            # Settings map
            settings = {
                "dpi": int(dpi_var.get() or 300),
                "page_size": page_size_var.get(),
                "color_mode": color_var.get(),
                "sides": sides_var.get(),
                "source": source_var.get(),
                "auto_orient": bool(auto_orient_var.get()),
            }

            # Acquire images
            imgs = _acquire_pages(device_id, settings, progress_cb=progress_cb, status_cb=status_cb)

            if cancel_flag["stop"]:
                set_status("Scan cancelled.", WARNING)
                return

            # Output path
            out_dir = out_dir_var.get().strip() or app.dirs.get("scanned", ".")
            os.makedirs(out_dir, exist_ok=True)
            base = fname_var.get().strip() or "SCAN"

            # Unique target path
            ts = time.strftime("%Y%m%d_%H%M%S")
            if save_as_var.get() == "PDF":
                out_path = os.path.join(out_dir, f"{base}_{ts}.pdf")
                saved = _save_images(imgs, out_path, "PDF")
            else:
                out_path = os.path.join(out_dir, f"{base}_{ts}.png")
                saved = _save_images(imgs, out_path, "PNG")

            pbar.configure(value=pbar["maximum"])
            set_status(f"Scan complete — saved: {', '.join(os.path.basename(x) for x in saved)}", SUCCESS)

            # Preview first output (PDF preferred)
            if preview_var.get() and hasattr(app, "_show_pdf_preview"):
                try:
                    if save_as_var.get() == "PDF":
                        app._show_pdf_preview(saved[0])
                    else:
                        # When PNG, quickly wrap into a temp one-page PDF for preview using existing viewer
                        tmp_pdf = os.path.join(out_dir, f"{base}_{ts}_preview.pdf")
                        im0 = imgs[0]
                        if im0.mode in ("RGBA", "P"):
                            im0 = im0.convert("RGB")
                        im0.save(tmp_pdf, "PDF", save_all=True)
                        app._show_pdf_preview(tmp_pdf)
                except Exception:
                    pass

        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Scan Error", str(e))
            set_status(f"Scan failed: {e}", ERROR)

    def on_scan_click():
        scan_btn.configure(state="disabled")
        cancel_btn.configure(state="normal")
        t = threading.Thread(target=lambda: (do_scan(), scan_btn.configure(state="normal")), daemon=True)
        t.start()

    scan_btn.configure(command=on_scan_click)

    # Populate scanners on open
    refresh_scanners()

    return tab
