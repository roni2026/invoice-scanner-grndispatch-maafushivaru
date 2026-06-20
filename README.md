# 🧾 Maafushivaru Document Processing Hub

**An end-to-end desktop application that turns stacks of scanned supplier invoices & Goods Received Notes (GRNs) into a clean, ready-to-file Excel register — automatically.**

Built for the stores / receiving department of **Maafushivaru (Maldives)**, this tool replaces hours of manual typing, sorting and cross-checking with a guided, mostly one-click workflow: **scan → read → identify supplier → extract figures → rename → dispatch to Excel.**

---

## 📌 What this project is

Receiving teams at a resort handle dozens of paper invoices and GRNs every day. Each document has to be:

1. Scanned,
2. Read and understood (supplier name, GRN no., PO no., invoice no., amount & currency),
3. Renamed to a consistent file name,
4. Recorded into a master spreadsheet (the "GRN Dispatch" register).

Doing this by hand is slow, error-prone and tedious. **Maafushivaru Document Processing Hub** automates the whole chain with OCR (Optical Character Recognition), fuzzy supplier matching and optional AI, wrapped in a friendly tabbed Windows GUI.

> In short: **drop in scanned PDFs, and get back a tidy Excel file with every invoice's supplier, document numbers, currency and amount — plus consistently renamed source files.**

---

## ⚙️ How it works

The app reads scanned PDFs, extracts the text, figures out *which supplier* the document belongs to, pulls out the key numbers, and writes everything to Excel. Here is the pipeline:

```
 ┌──────────┐   ┌─────────────┐   ┌──────────────┐   ┌────────────────┐   ┌──────────────┐   ┌─────────────┐
 │  SCAN    │ → │   OCR /     │ → │  SUPPLIER    │ → │   FIELD        │ → │   RENAME     │ → │  DISPATCH   │
 │  PDF in  │   │ TEXT EXTRACT│   │  MATCHING    │   │  EXTRACTION    │   │   FILES      │   │  TO EXCEL   │
 └──────────┘   └─────────────┘   └──────────────┘   └────────────────┘   └──────────────┘   └─────────────┘
   scanner/      Tesseract /        fuzzy + alias       GRN, PO, invoice    Scan_001.pdf →    GRN_OUTPUT_*.xlsx
   folder        OCR.space API      + AI matching       no., amount, ccy    proper names      master register
```

### Step by step

1. **Ingest** — Pull documents straight from a scanner (Windows WIA) or from a watched `SCANNED/` folder. PDFs can be auto-renamed to a clean `Scan_001.pdf … Scan_NNN.pdf` sequence.
2. **Read the page (OCR)** — Text is extracted using one of several engines:
   - **Tesseract** (local, default) with image enhancement, scaling and configurable PSM/OEM modes;
   - **OCR.space API** (cloud) for the "AI Extract" workflow;
   - optional **PaddleOCR / EasyOCR** engines (installed on demand).
   The app can read **only the relevant zone** (fast header scan) or the **full page**, and only the pages that are actual *Receiving Reports* (detected by `RC-MAM-…` and PO-number patterns) are kept — trimming away noise.
3. **Identify the supplier** — Raw OCR text is rarely clean, so the app uses a layered matching system:
   - a curated list of **177 known suppliers** plus an **alias dictionary** (e.g. `EURO MARKETING` → `EURO MARKETING PVT LTD`);
   - **fuzzy matching** (RapidFuzz + Levenshtein) to survive OCR typos like `EASRN AND AIIED` → `EASTERN AND ALLIED`;
   - **progressive prefix matching**, **OCR word correction**, and an optional **smart bidirectional cross-matcher** (supplier ⇄ invoice format);
   - an optional **AI matcher** (OpenAI `gpt-4o-mini` or any OpenAI-compatible / local model) for the hard cases.
4. **Extract the fields** — GRN number, Purchase Order number, invoice number, currency and amount are parsed out using configurable regex patterns (e.g. GRN prefix `RC-MAM-`, PO prefix `MAM-`).
5. **Rename** — Source PDFs are renamed consistently and (optionally) archived or moved to `PROCESSED/` / `FAILED/` folders.
6. **Dispatch to Excel** — Everything is written into a formatted `GRN_OUTPUT_<timestamp>.xlsx` with columns:

   | INVOICE DATE | SUPPLIER NAME | PURCHASE ORDER # | INVOICE # | USD | MVR | EUR | GBP | SGD | GRN NO. |
   |---|---|---|---|---|---|---|---|---|---|

   Multi-currency amounts are placed in the correct column automatically.

Throughout, a **confidence score** flags low-certainty reads so a human can double-check them, and a full activity **log** is written to `LOGS/`.

---

## 🧩 The application (tabs & modules)

The GUI is a single Tkinter app (`MaafushivaruHub`) organized into tabs:

| Tab | Purpose |
|---|---|
| **Dashboard** | Live status, counts, and at-a-glance processing overview. |
| **Scan** | Acquire pages directly from a scanner via Windows WIA (page size, DPI, color mode, duplex, auto-orient, preview). |
| **Renamer** | OCR + extract from scanned PDFs and rename them consistently (with a **Dry Run** preview before committing). |
| **AI Extract** | Trim PDFs to receiving-report pages, OCR them via OCR.space, extract fields, then *"Send to Tabs"*. |
| **Dispatch** | Review extracted rows and **export the GRN Excel register**. |
| **Settings** | Choose OCR engine & mode, matching strategy, thresholds, AI provider/key, folders, notifications, auto-ingest, etc. |
| **About** | App / version info. |

### Key source files

| File | Role |
|---|---|
| `maafushivaru_hub.py` | Main application & GUI — the document-processing hub (v5.0). |
| `aiextracttab.py` | "AI Extract" tab: page trimming + OCR.space + field extraction. |
| `scan_tab.py` | Scanner acquisition tab (Windows WIA via pywin32). |
| `ai_supplier_matcher.py` | OCR.space client + optional AI (OpenAI / local) supplier matcher. |
| `supplier_matcher.py` | Multiple selectable fuzzy supplier-matching strategies. |
| `smart_cross_matcher.py` | Bidirectional supplier ⇄ invoice-format cross-matching. |
| `ocr_word_corrector.py` | Repairs broken OCR words against known supplier names/patterns. |
| `engine_installer.py` | Detects & installs OCR engines (Tesseract / PaddleOCR / EasyOCR) on demand. |
| `pdfrename.py` | Standalone helper to batch-rename PDFs to `Scan_NNN.pdf`. |
| `config.json` | All settings: suppliers, aliases, patterns, OCR/AI options, folders. |

---

## ✨ Features

- 🖨️ **Direct scanner capture** (Windows WIA) — feeder/flatbed/duplex, DPI, color mode, auto-orient.
- 🔍 **Multi-engine OCR** — Tesseract (local), OCR.space (cloud), optional PaddleOCR/EasyOCR.
- ⚡ **Zone vs. full-page OCR** — fast header-only scanning when you don't need the whole page.
- 📄 **Smart page detection** — keeps only true Receiving Report pages, ignores the rest.
- 🏷️ **177-supplier knowledge base + alias dictionary** for accurate identification.
- 🤝 **Layered matching** — fuzzy + prefix + word-correction + cross-matching + optional AI.
- 🔢 **Automatic field extraction** — GRN, PO, invoice numbers, currency & amount via configurable patterns.
- 💱 **Multi-currency support** — USD, MVR, EUR, GBP, SGD columns.
- 📊 **Formatted Excel export** — timestamped `GRN_OUTPUT_*.xlsx` master register.
- 🗂️ **Automatic file renaming & foldering** — `PROCESSED/`, `FAILED/`, `ARCHIVE/`.
- 👀 **Dry Run mode** — preview every rename/extract before anything is moved.
- 📈 **Confidence scoring** — low-certainty results are flagged for review.
- 🔔 **Desktop notifications** & **auto-ingest watcher** (watchdog) for hands-free processing.
- 🧵 **Multi-threaded** processing for speed.
- 🧠 **Self-learning aliases** — new supplier spellings can be remembered.
- 📝 **Full logging** to `LOGS/` for traceability.

---

## 📈 How it boosts productivity & efficiency

| Before (manual) | After (with this Hub) |
|---|---|
| Read each invoice and type the supplier, GRN, PO, invoice no. & amount by hand | OCR + matching extracts them automatically |
| Spelling/typo errors creep into the register | Fuzzy matching & word correction fix garbled OCR text |
| Inconsistent file names, hard to find documents later | Standardized `Scan_NNN.pdf` naming + organized folders |
| Sorting documents by supplier is slow and manual | 177-supplier knowledge base classifies them instantly |
| Each multi-page scan has to be split/checked manually | Auto-detects and trims to the actual receiving-report pages |
| Re-typing into Excel, one row at a time | One-click export to a formatted, multi-currency GRN register |
| No record of what was processed | Confidence scores + full logs for audit & review |

**Net effect:** what used to take **hours of repetitive data entry per batch** becomes a **few minutes of supervised, mostly automated processing** — with **fewer errors**, **consistent records**, and a **clear audit trail**. The team spends its time *verifying flagged items* instead of *typing every line*.

---

## 🚀 Getting started

> **Platform:** Windows (scanner capture uses Windows WIA via `pywin32`).

### 1. Install dependencies

```bash
pip install pytesseract PyMuPDF opencv-python numpy pillow rapidfuzz openpyxl
# Optional features:
pip install watchdog plyer pywin32 PyPDF2   # auto-ingest, notifications, scanning
pip install paddlepaddle paddleocr          # PaddleOCR engine (optional)
pip install easyocr                         # EasyOCR engine (optional)
```

Install the **Tesseract OCR** binary (Windows): <https://github.com/UB-Mannheim/tesseract/wiki>
and set its path in `config.json` → `tesseract_cmd`.

> The built-in **OCR Engine Installer** (`engine_installer.py`) can install the optional pip-based engines for you on demand.

### 2. Configure

Open `config.json` and review:
- `tesseract_cmd` — path to your Tesseract executable;
- `suppliers` / `aliases` — your supplier list and known spellings;
- `patterns` — GRN/PO prefixes & digit rules (`RC-MAM-`, `MAM-`, …);
- `app_settings` — OCR engine/mode, matching strategy, thresholds, folders;
- `ai_settings` — enable AI matching and add your OpenAI (or local) API key;
- `ocr_space` — OCR.space API key for the AI Extract workflow.

### 3. Run

```bash
python maafushivaru_hub.py
```

Then:
1. **Scan** or drop PDFs into the `SCANNED/` folder.
2. Use **Renamer** / **AI Extract** to OCR and extract (try **Dry Run** first).
3. Review rows in **Dispatch** and **Export to Excel**.

> 💡 **Tip:** Use `pdfrename.py` to quickly standardize a folder of PDFs:
> ```bash
> python pdfrename.py "C:\path\to\folder"
> ```

---

## 📂 Repository layout

```
.
├── maafushivaru_hub.py        # Main GUI application (the Hub)
├── aiextracttab.py            # AI Extract tab (OCR.space + extraction)
├── scan_tab.py                # Scanner capture tab (Windows WIA)
├── ai_supplier_matcher.py     # OCR.space + AI supplier matcher
├── supplier_matcher.py        # Fuzzy matching strategies
├── smart_cross_matcher.py     # Supplier ⇄ invoice cross-matching
├── ocr_word_corrector.py      # OCR word repair
├── engine_installer.py        # On-demand OCR engine installer
├── pdfrename.py               # Batch PDF renamer (CLI)
├── config.json                # Suppliers, aliases, patterns, settings
├── SCANNED/                   # Input PDFs
├── TEMP API PDFS/             # Trimmed pages sent to OCR.space
├── LOGS/                      # Activity logs
├── GRN_OUTPUT_*.xlsx          # Exported GRN dispatch registers
├── logo.png / logo.ico        # App branding
└── LICENSE
```

---

## 🔐 Notes & security

- Keep API keys (OpenAI, OCR.space) private — prefer environment variables / a local untracked config over committing real keys.
- OCR accuracy depends on scan quality; **always review confidence-flagged rows** before filing.
- Scanner capture requires **Windows + a WIA-compatible driver**.

---

## 📄 License

See [`LICENSE`](LICENSE).

---

*Built to make receiving-room paperwork fast, accurate and painless. 🌴*
