# smart_cross_matcher.py
# Bidirectional supplier ↔ invoice cross-matching.
#
# Feature A: Know the supplier → apply its invoice format rules
# Feature B: Know the invoice format → infer the supplier
#
# This module is toggled on/off from Settings.
# When OFF the system works exactly as before.

import re
from typing import Optional, List, Dict, Tuple


# ---------------------------------------------------------------------------
# CROSS-REFERENCE TABLE
# ---------------------------------------------------------------------------
# Each entry maps a supplier keyword to:
#   "invoice_patterns": list of regex strings that match valid invoice numbers
#   "invoice_prefixes": canonical prefix letters for that supplier's invoices
#
# When we see an invoice matching one of these patterns AND no supplier has
# been found yet, we assign the mapped supplier.
#
# When we HAVE found the supplier, we use the patterns to validate / clean
# the extracted invoice number.
#
SUPPLIER_INVOICE_XREF: Dict[str, Dict] = {
    "EASTERN AND ALLIED": {
        "invoice_patterns": [
            r"\bEA[-/]\d{4,8}\b",
            r"\bEAL[-/]\d{4,8}\b",
        ],
        "invoice_prefixes": ["EA", "EAL"],
        "canonical_prefix": "EA",
        "separator": "-",
    },
    "BEST BUY": {
        "invoice_patterns": [
            r"\bBB[-/]\d{4,8}\b",
        ],
        "invoice_prefixes": ["BB"],
        "canonical_prefix": "BB",
        "separator": "-",
    },
    "STANDARD": {
        "invoice_patterns": [
            r"\bSI[-/]\d{4,8}\b",
            r"\bIS[-/]\d{4,8}\b",
        ],
        "invoice_prefixes": ["SI", "IS"],
        "canonical_prefix": "SI",
        "separator": "/",
    },
    "SAWHNEY": {
        "invoice_patterns": [
            r"\bSB[-/]\d{4,8}\b",
            r"\bSBL[-/]\d{4,8}\b",
            r"\b[A-Z]{2,4}/\d+/\d{4}\b",
        ],
        "invoice_prefixes": ["SB", "SBL"],
        "canonical_prefix": "SB",
        "separator": "/",
    },
    "LYCORN": {
        "invoice_patterns": [
            r"\bLY[-/]\d{4,8}\b",
            r"\b\d{6,10}\b",
        ],
        "invoice_prefixes": ["LY"],
        "canonical_prefix": "LY",
        "separator": "-",
    },
    "COSMO": {
        "invoice_patterns": [
            r"\b[A-Z]{2,4}[-/]\d{4,8}/\d{4}\b",
            r"\bCI[-/]\d{4,8}\b",
        ],
        "invoice_prefixes": ["CI", "CO"],
        "canonical_prefix": "CI",
        "separator": "/",
    },
    # ── HOW TO ADD MORE ────────────────────────────────────────────────────
    # "YOUR SUPPLIER NAME": {
    #     "invoice_patterns": [r"\bXX[-/]\d{4,8}\b"],  # regex(es) for invoice numbers
    #     "invoice_prefixes": ["XX"],                    # prefix letters OCR might see
    #     "canonical_prefix": "XX",                     # the correct prefix to use
    #     "separator": "-",                              # - or /
    # },
}


# ---------------------------------------------------------------------------
# FEATURE A: Given supplier → infer correct invoice format
# ---------------------------------------------------------------------------
def infer_invoice_from_supplier(
    raw_invoice: str,
    supplier_name: str,
) -> str:
    """
    Given an already-identified supplier, clean up a garbled invoice number
    using that supplier's known format.

    Example:
        supplier = "EASTERN AND ALLIED"
        raw_invoice = "EH-123456"   ← OCR misread E A as E H
        result: "EA-123456"

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

    # Try to extract prefix + number from raw_invoice
    m = re.match(r"([A-Z]{1,4})([-/])(\d{4,10})", raw_invoice.upper().strip())
    if not m:
        return raw_invoice

    raw_prefix, raw_sep, number = m.group(1), m.group(2), m.group(3)

    # Check if raw_prefix is close to any known prefix (≤ 2 errors)
    best_prefix = raw_prefix
    best_dist   = 999
    for kp in known_prefixes + [canonical_prefix]:
        dist = _edit_distance(raw_prefix, kp)
        if dist < best_dist:
            best_dist   = dist
            best_prefix = kp

    if best_dist <= 2:
        return f"{canonical_prefix}{separator}{number}"

    return raw_invoice


# ---------------------------------------------------------------------------
# FEATURE B: Given invoice number → infer supplier
# ---------------------------------------------------------------------------
def infer_supplier_from_invoice(
    invoice_text: str,
    suppliers: List[str],
    aliases: Dict[str, List[str]],
) -> Tuple[Optional[str], float]:
    """
    Scans invoice_text for patterns that uniquely identify a supplier.
    Returns (canonical_supplier_name, confidence) or (None, 0.0).

    Confidence is 0–100. A direct pattern match gives 88.
    """
    upper = invoice_text.upper()

    for sup_keyword, entry in SUPPLIER_INVOICE_XREF.items():
        for pattern in entry.get("invoice_patterns", []):
            if re.search(pattern, upper):
                # Resolve to canonical supplier name
                canonical = _resolve_to_canonical(sup_keyword, suppliers, aliases)
                if canonical:
                    return canonical, 88.0

    return None, 0.0


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def _find_entry(sup_upper: str) -> Optional[Dict]:
    """Find the XREF entry whose key appears in the supplier name."""
    for key, entry in SUPPLIER_INVOICE_XREF.items():
        if key.upper() in sup_upper or sup_upper in key.upper():
            return entry
    return None


def _resolve_to_canonical(keyword: str, suppliers: List[str], aliases: Dict[str, List[str]]) -> Optional[str]:
    """Map a XREF keyword back to the canonical supplier name from config."""
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
    return keyword   # return the keyword itself as a last resort


def _edit_distance(a: str, b: str) -> int:
    """Simple Levenshtein distance."""
    try:
        from rapidfuzz.distance import Levenshtein
        return Levenshtein.distance(a, b)
    except ImportError:
        # Pure Python fallback
        m, n = len(a), len(b)
        dp = list(range(n + 1))
        for i in range(1, m + 1):
            prev = dp[:]
            dp[0] = i
            for j in range(1, n + 1):
                cost = 0 if a[i-1] == b[j-1] else 1
                dp[j] = min(dp[j] + 1, dp[j-1] + 1, prev[j-1] + cost)
        return dp[n]
