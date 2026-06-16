# smart_cross_matcher.py
# Bidirectional supplier ↔ invoice cross-matching.
#
# Feature A: Know the supplier → apply its invoice format rules
# Feature B: Know the invoice format → infer the supplier
#
# This module is toggled on/off from Settings.
# When OFF the system works exactly as before.
#
# ── PATTERN DESIGN NOTES ────────────────────────────────────────────────────
# Each XREF entry carries TWO pattern lists:
#
#   "invoice_patterns"    Used ONLY for Feature B (invoice text → supplier).
#                         Must be prefix-anchored or otherwise specific enough
#                         to avoid false positives across suppliers.
#                         Bare numeric-only patterns are excluded here because
#                         the same number could belong to dozens of suppliers.
#
#   "validation_patterns" Used ONLY for Feature A (validate/clean a known
#                         supplier's invoice). May include broad numeric regexes.
#
# Dict order matters for Feature B: entries that could shadow each other are
# placed so more-specific patterns come first (e.g. SH1-prefixed SALESCO
# before the generic YY/#### pattern used by ALAAN COMPANY).
# ─────────────────────────────────────────────────────────────────────────────

import re
from typing import Optional, List, Dict, Tuple


# ---------------------------------------------------------------------------
# CROSS-REFERENCE TABLE
# ---------------------------------------------------------------------------
SUPPLIER_INVOICE_XREF: Dict[str, Dict] = {

    # ── ORIGINAL ENTRIES ─────────────────────────────────────────────────

    "EASTERN AND ALLIED": {
        # EA-##### / EAL-#####
        "invoice_patterns":    [r"\bEA[-/]\d{4,8}\b", r"\bEAL[-/]\d{4,8}\b"],
        "validation_patterns": [r"\bEA[-/]\d{4,8}\b", r"\bEAL[-/]\d{4,8}\b"],
        "invoice_prefixes": ["EA", "EAL"],
        "canonical_prefix": "EA",
        "separator": "-",
    },

    "BEST BUY": {
        # BB-#####
        "invoice_patterns":    [r"\bBB[-/]\d{4,8}\b"],
        "validation_patterns": [r"\bBB[-/]\d{4,8}\b"],
        "invoice_prefixes": ["BB"],
        "canonical_prefix": "BB",
        "separator": "-",
    },

    "STANDARD": {
        # SI/##### or IS/#####
        "invoice_patterns":    [r"\bSI[-/]\d{4,8}\b", r"\bIS[-/]\d{4,8}\b"],
        "validation_patterns": [r"\bSI[-/]\d{4,8}\b", r"\bIS[-/]\d{4,8}\b"],
        "invoice_prefixes": ["SI", "IS"],
        "canonical_prefix": "SI",
        "separator": "/",
    },

    "SAWHNEY": {
        # SB/SBL prefix variants; 7-digit numeric starting 110xxxx
        "invoice_patterns": [
            r"\bSB[-/]\d{4,8}\b",
            r"\bSBL[-/]\d{4,8}\b",
            r"\b110\d{4}\b",
        ],
        "validation_patterns": [
            r"\bSB[-/]\d{4,8}\b",
            r"\bSBL[-/]\d{4,8}\b",
            r"\b[A-Z]{2,4}/\d+/\d{4}\b",
            r"\b110\d{4}\b",
        ],
        "invoice_prefixes": ["SB", "SBL"],
        "canonical_prefix": "SB",
        "separator": "/",
    },

    "LYCORN": {
        # LY-##### prefix only for inference; bare numeric too ambiguous
        "invoice_patterns":    [r"\bLY[-/]\d{4,8}\b"],
        "validation_patterns": [r"\bLY[-/]\d{4,8}\b", r"\b\d{6,10}\b"],
        "invoice_prefixes": ["LY"],
        "canonical_prefix": "LY",
        "separator": "-",
    },

    "COSMO": {
        # CI/CO prefix; also X/####/YYYY compound form
        "invoice_patterns": [
            r"\bCI[-/]\d{4,8}\b",
            r"\b[A-Z]{2,4}[-/]\d{4,8}/\d{4}\b",
        ],
        "validation_patterns": [
            r"\bCI[-/]\d{4,8}\b",
            r"\b[A-Z]{2,4}[-/]\d{4,8}/\d{4}\b",
        ],
        "invoice_prefixes": ["CI", "CO"],
        "canonical_prefix": "CI",
        "separator": "/",
    },

    # ── EXPANDED ENTRIES — GRN dispatch Mar–May 2026 ─────────────────────
    # Ordered most-specific → least-specific to prevent shadowing.

    # ── Highly distinctive prefix + structure ────────────────────────────

    "SHENZHEN SANHE": {
        # AF-YYYYMMDD-##### — extremely distinctive compound form
        "invoice_patterns":    [r"\bAF[-/]\d{8}[-/]\d{5}\b"],
        "validation_patterns": [r"\bAF[-/]\d{8}[-/]\d{5}\b"],
        "invoice_prefixes": ["AF"],
        "canonical_prefix": "AF",
        "separator": "-",
    },

    "SAFCO INTERNATIONAL": {
        # SF/EX/######
        "invoice_patterns":    [r"\bSF[/-]EX[/-]\d{6}\b"],
        "validation_patterns": [r"\bSF[/-]EX[/-]\d{6}\b"],
        "invoice_prefixes": ["SF"],
        "canonical_prefix": "SF",
        "separator": "/",
    },

    "SALESCO": {
        # SH1/YY/#### or SH1 YY #### — must come BEFORE generic YY/#### patterns
        "invoice_patterns":    [r"\bSH1[-/ ]\d{2}[-/ ]\d{4}\b"],
        "validation_patterns": [r"\bSH1[-/ ]\d{2}[-/ ]\d{4}\b", r"\b\d{4}\b"],
        "invoice_prefixes": ["SH1"],
        "canonical_prefix": "SH1",
        "separator": "/",
    },

    "MAM PETTY CASH": {
        # GTS-INV-######
        "invoice_patterns":    [r"\bGTS[-/]INV[-/]\d{6}\b"],
        "validation_patterns": [r"\bGTS[-/]INV[-/]\d{6}\b", r"\b\d{4,5}\b"],
        "invoice_prefixes": ["GTS"],
        "canonical_prefix": "GTS",
        "separator": "-",
    },

    "BIZ SOLUTIONS": {
        # BIS/INV/YYYY/### — specific compound prefix
        "invoice_patterns":    [r"\bBIS[/-]INV[/-]\d{4}[/-]\d{3}\b"],
        "validation_patterns": [r"\bBIS[/-]INV[/-]\d{4}[/-]\d{3}\b", r"\b20\d{2}[-/]\d{3}\b"],
        "invoice_prefixes": ["BIS"],
        "canonical_prefix": "BIS",
        "separator": "/",
    },

    "BIGWIG": {
        # OM-INV-YYYY-### — distinctive; bare YYYY-### excluded from inference
        "invoice_patterns":    [r"\bOM[-/]INV[-/]\d{4}[-/]\d{3}\b"],
        "validation_patterns": [r"\bOM[-/]INV[-/]\d{4}[-/]\d{3}\b", r"\b20\d{2}-\d{3}\b"],
        "invoice_prefixes": ["OM"],
        "canonical_prefix": "OM",
        "separator": "-",
    },

    "CHEMLAB": {
        # CL-I-####
        "invoice_patterns":    [r"\bCL[-/]I[-/]\d{4}\b"],
        "validation_patterns": [r"\bCL[-/]I[-/]\d{4}\b", r"\b\d{4}\b"],
        "invoice_prefixes": ["CL"],
        "canonical_prefix": "CL",
        "separator": "-",
    },

    "DHIAAGU": {
        # BA-#####-YY (new supplier)
        "invoice_patterns":    [r"\bBA[-/]\d{5}[-/]\d{2}\b"],
        "validation_patterns": [r"\bBA[-/]\d{5}[-/]\d{2}\b"],
        "invoice_prefixes": ["BA"],
        "canonical_prefix": "BA",
        "separator": "-",
    },

    "ADK PHARMACEUTICAL SUPPLY": {
        # 14/ or 14- + 7 digits
        "invoice_patterns":    [r"\b14[/-]\d{7}\b"],
        "validation_patterns": [r"\b14[/-]\d{7}\b"],
        "invoice_prefixes": ["14"],
        "canonical_prefix": "14",
        "separator": "/",
    },

    "ALIHAVA CONSTRUCTION": {
        # AH/INV/YY/### — distinctive compound prefix
        "invoice_patterns":    [r"\bAH[-/]INV[-/]\d{2}[-/]\d{2,4}\b"],
        "validation_patterns": [r"\bAH[-/]INV[-/]\d{2}[-/]\d{2,4}\b", r"\b\d{2}[/-]\d{3,4}\b"],
        "invoice_prefixes": ["AH"],
        "canonical_prefix": "AH",
        "separator": "/",
    },

    "D BLUE": {
        # IN + 7 digits (no separator)
        "invoice_patterns":    [r"\bIN\d{7}\b"],
        "validation_patterns": [r"\bIN\d{7}\b"],
        "invoice_prefixes": ["IN"],
        "canonical_prefix": "IN",
        "separator": "",
    },

    "COSMOPOLITAN": {
        # INV + space + 7 digits (the space makes it distinctive)
        "invoice_patterns":    [r"\bINV\s+\d{7}\b"],
        "validation_patterns": [r"\bINV\s+\d{7}\b", r"\b1[78]\d{4}\b"],
        "invoice_prefixes": ["INV"],
        "canonical_prefix": "",
        "separator": "",
    },

    "INTERNATIONAL FOOD SOLUTION": {
        # MSI-######
        "invoice_patterns":    [r"\bMSI[-/]\d{6}\b"],
        "validation_patterns": [r"\bMSI[-/]\d{6}\b", r"\b\d{6}\b"],
        "invoice_prefixes": ["MSI"],
        "canonical_prefix": "MSI",
        "separator": "-",
    },

    "S&J SALES": {
        # TRV-YY-#
        "invoice_patterns":    [r"\bTRV[-/]\d{2}[-/]\d{1,3}\b"],
        "validation_patterns": [r"\bTRV[-/]\d{2}[-/]\d{1,3}\b", r"\b\d{2}[/-]\d{3}\b"],
        "invoice_prefixes": ["TRV"],
        "canonical_prefix": "TRV",
        "separator": "-",
    },

    "EMPARAL": {
        # EYY/#### (year-embedded prefix)
        "invoice_patterns":    [r"\bE\d{2}[/-]\d{4}\b"],
        "validation_patterns": [r"\bE\d{2}[/-]\d{4}\b", r"\b\d{4}\b"],
        "invoice_prefixes": ["E"],
        "canonical_prefix": "E",
        "separator": "/",
    },

    "LILY INTERNATIONAL": {
        # INV-####### (7 digits) or LLS/####/YY
        "invoice_patterns": [
            r"\bINV[-/]\d{7}\b",
            r"\bLLS[/-]\d{4}[/-]\d{2}\b",
        ],
        "validation_patterns": [
            r"\bINV[-/]\d{7}\b",
            r"\bLLS[/-]\d{4}[/-]\d{2}\b",
        ],
        "invoice_prefixes": ["INV", "LLS"],
        "canonical_prefix": "INV",
        "separator": "-",
    },

    "KAILAAN": {
        # INV/YYYY/######
        "invoice_patterns":    [r"\bINV[/-]\d{4}[/-]\d{6}\b"],
        "validation_patterns": [r"\bINV[/-]\d{4}[/-]\d{6}\b", r"\b\d{3}\b"],
        "invoice_prefixes": ["INV"],
        "canonical_prefix": "",
        "separator": "",
    },

    "FRONTIER MARKETS": {
        # INV-YYYY-##### must come BEFORE POISE so the INV prefix is matched first
        "invoice_patterns": [
            r"\bINV[-/]\d{4}[-/]\d{5}\b",
        ],
        "validation_patterns": [
            r"\bINV[-/]\d{4}[-/]\d{5}\b",
            r"\b20\d{2}[/-]\d{5}\b",
            r"\b\d{3,5}\b",
        ],
        "invoice_prefixes": ["INV"],
        "canonical_prefix": "",
        "separator": "",
    },

    "POISE DISTRIBUTORS": {
        # YYYY-##### (5-digit suffix after full 4-digit year)
        "invoice_patterns":    [r"\b20\d{2}[-/]\d{5}\b"],
        "validation_patterns": [r"\b20\d{2}[-/]\d{5}\b", r"\b\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "-",
    },

    "HASSAN MARINE": {
        # 8-digit starting 1002xxxx
        "invoice_patterns":    [r"\b1002\d{4}\b"],
        "validation_patterns": [r"\b1002\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },

    "RESTOFAIR MALDIVES": {
        # 8-digit starting 1800xxxx (new supplier)
        "invoice_patterns":    [r"\b1800\d{4}\b"],
        "validation_patterns": [r"\b1800\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },

    "EURO MARKETING": {
        # 9-digit starting 10018/10019
        "invoice_patterns":    [r"\b1001[89]\d{4}\b"],
        "validation_patterns": [r"\b1001[89]\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },

    "COUNTLINE": {
        # YY-###### or YY/###### (6-digit suffix distinguishes from YY-####)
        "invoice_patterns":    [r"\b\d{2}[-/]\d{6}\b"],
        "validation_patterns": [r"\b\d{2}[-/]\d{6}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "-",
    },

    "SUN FRONT LIGHTING": {
        # #####-YY or ######-YY (number THEN year — reversed order)
        "invoice_patterns":    [r"\b\d{5,6}[-/]\d{2}\b"],
        "validation_patterns": [r"\b\d{5,6}[-/]\d{2}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "-",
    },

    "NASREENA ABDULLA": {
        # YYYY-0### — zero-padded 4-digit suffix (observed: 0042, 0046, 0050)
        # This distinguishes from PRINTLAB which uses non-zero-leading suffixes
        "invoice_patterns":    [r"\b20\d{2}[-/]0\d{3}\b"],
        "validation_patterns": [r"\b20\d{2}[-/]\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "-",
    },

    "PRINTLAB": {
        # YYYY-#### where suffix starts with [1-9] (observed: 2725, 4182)
        # Zero-padded suffixes belong to NASREENA
        "invoice_patterns":    [r"\b20\d{2}[-/][1-9]\d{3}\b"],
        "validation_patterns": [r"\b20\d{2}[-/]\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "-",
    },

    "PREMIUM SUPPLIES": {
        # YY-#### with DASH separator only (observed: 26-2481, 26-1542)
        # Slash form (26/XXXX) belongs to ALAAN COMPANY
        "invoice_patterns":    [r"(?<![A-Z0-9/\d])\b\d{2}-\d{4}\b"],
        "validation_patterns": [r"\b\d{2}[-/]\d{4}\b", r"\b\d{6}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "-",
    },

    "ALAAN COMPANY": {
        # YY/#### with SLASH separator only (observed: 26/0164)
        # Dash form belongs to PREMIUM SUPPLIES
        # Negative lookbehind prevents matching mid-string (e.g. SH1/26/1039)
        "invoice_patterns":    [r"(?<![A-Z0-9])\b\d{2}/\d{4}\b"],
        "validation_patterns": [r"\b\d{2}[/-]\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "/",
    },

    # ── Numeric-only (no inference) ─────────────────────────────────────
    # These suppliers use bare numbers that are indistinguishable at inference
    # time. Patterns here are for Feature A validation only.

    "ADK GENERAL TRADING": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{6}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },
    "ALBA INTERNATIONAL": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{6}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },
    "ARAABY": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{5}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },
    "HARDWARE LAB": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },
    "MIYAMI": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{6}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },
    "NOVELTY BOOKSHOP": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{3}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },
    "SEALANDS": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{5}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },
    "SUPER POWER": {
        "invoice_patterns":    [],
        "validation_patterns": [r"\b\d{1,4}\b"],
        "invoice_prefixes": [],
        "canonical_prefix": "",
        "separator": "",
    },

    # ── HOW TO ADD MORE ──────────────────────────────────────────────────
    # "YOUR SUPPLIER NAME": {
    #     "invoice_patterns":    [r"\bXX[-/]\d{4,8}\b"],   # prefix-anchored only
    #     "validation_patterns": [r"\bXX[-/]\d{4,8}\b"],   # may include numeric
    #     "invoice_prefixes": ["XX"],
    #     "canonical_prefix": "XX",
    #     "separator": "-",
    # },
}


# ---------------------------------------------------------------------------
# DISPATCH ALIASES — sourced from GRN dispatch Mar–May 2026
# Merged at runtime with config.json aliases via merge_dispatch_aliases().
# ---------------------------------------------------------------------------
DISPATCH_ALIASES: Dict[str, List[str]] = {
    "ADK GENERAL TRADING": [
        "ADK GENERAL TRADERS",
    ],
    "ADK PHARMACEUTICAL SUPPLY": [
        "ADK PHARMACEUTICALS",
    ],
    "ALBA INTERNATIONAL": [
        "ALBAINTERNATIONAL",
    ],
    "ALIHAVA CONSTRUCTION & TRADING CO PVT LTD": [
        "ALIHAVA",
        "ALIHAVA CONSTRUCTION & TRADING",
        "ALIHAVA CONSTRUCTION & TRADING CO",
        "ALIHAVA CONSTRUCTION AND TRADING",
    ],
    "ARAABY PVT LTD": [
        "ARAABY",
    ],
    "BIGWIG": [
        "BIG WIG",
    ],
    "BIZ SOLUTIONS": [
        "BIZ SOLUTIONS P",
    ],
    "CHEMLAB PRIVATE LTD": [
        "CHEMLAB PRIVATE",
    ],
    "COSMOPOLITAN": [
        "COSMOPOLTION",
    ],
    "COUNTLINE": [
        "COUNTLINF",
    ],
    "D BLUE PRIVATE LTD": [
        "D BLUE PRIVATE",
    ],
    "EMPARAL PVT LTD": [
        "EMPARAL PVT",
    ],
    "EURO MARKETING PVT LTD": [
        "EURO MARKET",
        "EURO MARKETING  PVT LTD",
    ],
    "FRONTIER MARKETS ENTERPRISES PVT LTD": [
        "FRONTIER MARKETS ENTERPRISE",
        "FRONTIER MARKETS ENTERPRISES",
    ],
    "HARDWARE LAB": [
        "HARDWARE",
    ],
    "HASSAN MARINE EQUIPMENT SHOP": [
        "HASSAN MARINE EQUIPMENT",
    ],
    "INTERNATIONAL FOOD SOLUTION PVT LTD": [
        "INTERNATIONAL FOOD SOLUATION",
    ],
    "KAILAAN TRENDS": [
        "KAILAAN",
    ],
    "LILY INTERNATIONAL PVT LTD": [
        "LILY INTERNATIONAL PVT LTD.",
    ],
    "MAM PETTY CASH": [
        "MAM PETTY CASH - MVR",
    ],
    "MIYAMI TRADERS": [
        "MIYAMI TRADING",
    ],
    "NASREENA ABDULLA (FISH SUPPLIER)": [
        "NASREEN ABDULLAH (FISH)",
    ],
    "NOVELTY BOOKSHOP": [
        "NOVELTY BOOK SHOP",
    ],
    "POISE DISTRIBUTORS": [
        "POISE DISTRIBI",
        "POISE DISTRIBUTOR",
    ],
    "PREMIUM SUPPLIES": [
        "PREMIUM SUPP",
    ],
    "PRINTLAB": [
        "PRINTALAB",
    ],
    "S&J SALES CORPORATION": [
        "S&J SALES PO 17833,17787,17783",
    ],
    "SAFCO INTERNATIONAL GEN TRD LLC": [
        "SAFCO INTERNATIONAL GEN TRD",
    ],
    "SALESCO PVT LTD": [
        "SALESCO PVT",
    ],
    "SAWHNEY FOOD STAFF TRADING CO.": [
        "SAWHNEY FOOD STAFF TRADING",
        "SAWHNEY FOOD STAFF TRADING CO",
    ],
    "SEALANDS PVT LTD": [
        "SEALANDS",
    ],
    "SHENZHEN SANHE E-COMMERCE CO. LTD": [
        "SHENZHEN SANHE E CONNERCE CO.LTD",
        "SHENZHEN SANHE E-COMMERCE",
    ],
    "SUN FRONT LIGHTING": [
        "SUN FRONT LIGHTIN",
    ],
    "SUPER POWER CO. LTD": [
        "SUPER POWER CO",
    ],
    # ── Truly new suppliers (not previously in config.json) ──────────────
    "DHIAAGU": [],
    "ALAAN COMPANY": [],
    "RESTOFAIR MALDIVES": [],
}


# ---------------------------------------------------------------------------
# FEATURE A  —  Given supplier → validate / clean an invoice number
# ---------------------------------------------------------------------------
def infer_invoice_from_supplier(
    raw_invoice: str,
    supplier_name: str,
) -> str:
    """
    Given an already-identified supplier, clean up a garbled invoice number
    using that supplier's known format.

    Example:
        supplier    = "EASTERN AND ALLIED SRI LANKA"
        raw_invoice = "EH-123456"   ← OCR misread EA as EH
        result      → "EA-123456"

    Returns the cleaned invoice string, or the original if no rule applies.
    """
    sup_upper = supplier_name.upper()
    entry = _find_entry(sup_upper)
    if not entry:
        return raw_invoice

    canonical_prefix = entry.get("canonical_prefix", "")
    separator        = entry.get("separator", "-")
    known_prefixes   = entry.get("invoice_prefixes", [])

    if not canonical_prefix or not raw_invoice:
        return raw_invoice

    m = re.match(r"([A-Z]{1,4})([-/])(\d{4,10})", raw_invoice.upper().strip())
    if not m:
        return raw_invoice

    raw_prefix, _sep, number = m.group(1), m.group(2), m.group(3)

    best_dist = 999
    for kp in known_prefixes + [canonical_prefix]:
        d = _edit_distance(raw_prefix, kp)
        if d < best_dist:
            best_dist = d

    if best_dist <= 2:
        return f"{canonical_prefix}{separator}{number}"

    return raw_invoice


# ---------------------------------------------------------------------------
# FEATURE B  —  Given invoice text → infer supplier
# ---------------------------------------------------------------------------
def infer_supplier_from_invoice(
    invoice_text: str,
    suppliers: List[str],
    aliases: Dict[str, List[str]],
) -> Tuple[Optional[str], float]:
    """
    Scans invoice_text for prefix-anchored patterns that uniquely identify
    a supplier.  Returns (canonical_supplier_name, confidence) or (None, 0.0).

    Only "invoice_patterns" (specific/prefix-anchored) are used here.
    Confidence is 0–100; a direct pattern match returns 88.
    """
    upper = invoice_text.upper()

    for sup_keyword, entry in SUPPLIER_INVOICE_XREF.items():
        for pattern in entry.get("invoice_patterns", []):
            if re.search(pattern, upper):
                canonical = _resolve_to_canonical(sup_keyword, suppliers, aliases)
                if canonical:
                    return canonical, 88.0

    return None, 0.0


# ---------------------------------------------------------------------------
# FEATURE C  —  Merge dispatch aliases into the live aliases dict
# ---------------------------------------------------------------------------
def merge_dispatch_aliases(aliases: Dict[str, List[str]]) -> Dict[str, List[str]]:
    """
    Merge DISPATCH_ALIASES into an existing aliases dict loaded from config.json.
    Existing entries are preserved; new aliases are appended without duplicates.
    New suppliers absent from config are created.
    Returns the merged dict.
    """
    merged = {k: list(v) for k, v in aliases.items()}
    for supplier, new_aliases in DISPATCH_ALIASES.items():
        if supplier not in merged:
            merged[supplier] = []
        for alias in new_aliases:
            if alias not in merged[supplier]:
                merged[supplier].append(alias)
    return merged


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def _find_entry(sup_upper: str) -> Optional[Dict]:
    """Return the first XREF entry whose keyword is contained in, or
    contains, the supplier name (case-insensitive substring match)."""
    for key, entry in SUPPLIER_INVOICE_XREF.items():
        if key.upper() in sup_upper or sup_upper in key.upper():
            return entry
    return None


def _resolve_to_canonical(
    keyword: str,
    suppliers: List[str],
    aliases: Dict[str, List[str]],
) -> Optional[str]:
    """Map an XREF keyword back to the canonical supplier name from config."""
    kw = keyword.upper()
    for s in suppliers:
        if kw in s.upper() or s.upper() in kw:
            return s
    for main, alias_list in aliases.items():
        if kw in main.upper() or main.upper() in kw:
            return main
        for alias in alias_list:
            if kw in alias.upper():
                return main
    return keyword  # last resort — return keyword itself


def _edit_distance(a: str, b: str) -> int:
    """Levenshtein distance (uses rapidfuzz when available)."""
    try:
        from rapidfuzz.distance import Levenshtein
        return Levenshtein.distance(a, b)
    except ImportError:
        m, n = len(a), len(b)
        dp = list(range(n + 1))
        for i in range(1, m + 1):
            prev = dp[:]
            dp[0] = i
            for j in range(1, n + 1):
                cost = 0 if a[i - 1] == b[j - 1] else 1
                dp[j] = min(dp[j] + 1, dp[j - 1] + 1, prev[j - 1] + cost)
        return dp[n]
