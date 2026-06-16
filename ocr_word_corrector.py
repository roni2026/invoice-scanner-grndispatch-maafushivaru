# ocr_word_corrector.py
# Corrects broken OCR words by fuzzy-matching each word against known
# supplier words and invoice number patterns.
#
# Example:
#   OCR output:  "EASRN AND AIIED SRILANKA"
#   Known name:  "EASTERN AND ALLIED SRILANKA"
#   Result:      corrects to "EASTERN AND ALLIED SRILANKA"
#
# Last updated: 2026-06 — aliases from GRN Dispatch Mar–May 2026;
#               3 new suppliers (DHIAAGU, ALAAN COMPANY, RESTOFAIR MALDIVES);
#               invoice patterns added for all active suppliers.

import re
from typing import Optional, List, Dict, Tuple
from rapidfuzz import fuzz


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
def _max_errors_for_word(word: str) -> int:
    n = len(word)
    if n <= 3:
        return 0
    if n <= 5:
        return 1
    return 2


_MIN_WORD_COVERAGE = 0.60


# ---------------------------------------------------------------------------
# CORE WORD FUZZY MATCH
# ---------------------------------------------------------------------------
def _word_matches(ocr_word: str, known_word: str) -> bool:
    max_err = _max_errors_for_word(known_word)
    if max_err == 0:
        return ocr_word == known_word
    from rapidfuzz.distance import Levenshtein
    dist = Levenshtein.distance(ocr_word, known_word)
    return dist <= max_err


# ---------------------------------------------------------------------------
# SUPPLIER NAME CORRECTOR
# ---------------------------------------------------------------------------
def correct_supplier_in_text(
    text: str,
    suppliers: List[str],
    aliases: Dict[str, List[str]],
) -> Tuple[str, Optional[str]]:
    """
    Scans every line of `text` for sequences of words that approximately
    match any known supplier name (or alias).

    Returns:
        (corrected_text, matched_canonical_name)
        If no match found, returns (text, None) — original text unchanged.
    """
    upper_text = text.upper()

    candidates: List[Tuple[str, str, List[str]]] = []

    for sup in suppliers:
        words = [w for w in re.sub(r"[^A-Z0-9]", " ", sup.upper()).split() if w]
        if words:
            candidates.append((sup.upper(), sup, words))

    for main, alias_list in aliases.items():
        for alias in alias_list:
            words = [w for w in re.sub(r"[^A-Z0-9]", " ", alias.upper()).split() if w]
            if words:
                candidates.append((alias.upper(), main, words))

    best_match: Optional[str] = None
    best_score: float = 0.0
    best_original_span: str = ""
    best_replacement: str = ""

    for line in upper_text.splitlines():
        line_words = [w for w in re.sub(r"[^A-Z0-9]", " ", line).split() if w]
        if not line_words:
            continue

        for display_name, canonical, cand_words in candidates:
            if len(cand_words) == 0:
                continue

            for start in range(len(line_words)):
                window = line_words[start: start + len(cand_words)]
                if not window:
                    break

                match_count = 0
                for ocr_w, known_w in zip(window, cand_words):
                    if _word_matches(ocr_w, known_w):
                        match_count += 1

                coverage = match_count / len(cand_words)
                if coverage >= _MIN_WORD_COVERAGE and coverage > best_score:
                    best_score = coverage
                    best_match = canonical
                    best_original_span = " ".join(window)
                    best_replacement = display_name

    if best_match and best_original_span:
        corrected = upper_text.replace(best_original_span, best_replacement, 1)
        return corrected, best_match

    return text, None


# ---------------------------------------------------------------------------
# INVOICE NUMBER CORRECTOR
# ---------------------------------------------------------------------------
# Each entry: (supplier_key, regex_pattern, description)
#
# supplier_key: substring of canonical supplier name (uppercase), or "*" for any.
# Regex uses named groups: prefix, sep, number (plus year, seq, date, etc.).
#
INVOICE_PATTERNS: List[Tuple[str, str, str]] = [

    # ── ADK GENERAL TRADING ─────────────────────────────────────────────────
    # 6-digit numeric   e.g. 747497
    ("ADK GENERAL TRADING", r"(?P<number>\d{6})", "ADK General: 6-digit"),

    # ── ADK PHARMACEUTICAL SUPPLY ───────────────────────────────────────────
    # 14/####### or 14-#######   e.g. 14/2786520, 14-2823014
    ("ADK PHARMACEUTICAL",
     r"(?P<prefix>14)(?P<sep>[-/])(?P<number>\d{7,8})",
     "ADK Pharma: 14/NNNNNNN"),

    # ── ALBA INTERNATIONAL ──────────────────────────────────────────────────
    # 6-digit numeric   e.g. 120685
    ("ALBA INTERNATIONAL", r"(?P<number>\d{6})", "Alba: 6-digit"),

    # ── ALAAN COMPANY ───────────────────────────────────────────────────────
    # YY/####   e.g. 26/0164
    ("ALAAN",
     r"(?P<year>\d{2})(?P<sep>/)(?P<number>\d{4})",
     "Alaan: YY/NNNN"),

    # ── ALIHAVA CONSTRUCTION & TRADING ──────────────────────────────────────
    # AH/INV/YY/### or AH-INV-YY-### or YY/### or YY-###   e.g. AH/INV/26/318
    ("ALIHAVA",
     r"(?:(?P<prefix>AH)(?P<sep>[-/])(?:INV(?P<sep2>[-/]))?)?(?P<year>\d{2})(?P<sep3>[-/])(?P<number>\d{2,4})",
     "Alihava: AH/INV/YY/NNN or YY/NNN"),

    # ── AMAGI EXPORTS ───────────────────────────────────────────────────────
    # ##AFE #####   e.g. 02AFE 02987
    ("AMAGI",
     r"(?P<prefix>\d{2}AFE)\s*(?P<number>\d{5})",
     "Amagi: NNAFe NNNNN"),

    # ── ANAKEE ──────────────────────────────────────────────────────────────
    # numeric or YYYY-C#-###   e.g. 3482, 2026-C2-960
    ("ANAKEE",
     r"(?:(?P<year>\d{4})-(?P<type>[A-Z]\d)-)?(?P<number>\d{3,4})",
     "Anakee: NNNN or YYYY-CN-NNN"),

    # ── ARAABY PVT LTD ──────────────────────────────────────────────────────
    # 5-digit numeric   e.g. 36851
    ("ARAABY", r"(?P<number>\d{5})", "Araaby: 5-digit"),

    # ── ATOLL MARKET ────────────────────────────────────────────────────────
    # YYYY-####   e.g. 2026-0337
    ("ATOLL MARKET",
     r"(?P<year>\d{4})(?P<sep>-)(?P<number>\d{4})",
     "Atoll Market: YYYY-NNNN"),

    # ── AYAAN EXPORT ────────────────────────────────────────────────────────
    # AY-YY-YY-INV/##   e.g. AY-25-26-INV/64
    ("AYAAN",
     r"(?P<prefix>AY)(?P<sep>-)(?P<y1>\d{2})-(?P<y2>\d{2})-(?P<type>INV)/(?P<number>\d+)",
     "Ayaan: AY-YY-YY-INV/NN"),

    # ── BARQUE INVESTMENT ───────────────────────────────────────────────────
    # YY/##### or YY-##### or 3-4 digit   e.g. 26/00899, 26-01923, 2050
    ("BARQUE",
     r"(?:(?P<year>\d{2})(?P<sep>[-/])(?P<number>\d{4,5}))|(?P<num2>\d{3,4})",
     "Barque: YY/NNNNN or NNNN"),

    # ── BESTBUY MALDIVES ────────────────────────────────────────────────────
    # 7-digit numeric   e.g. 1218265
    ("BESTBUY", r"(?P<number>\d{7})", "Bestbuy: 7-digit"),

    # ── BIGWIG ──────────────────────────────────────────────────────────────
    # YYYY-### or OM-INV-YYYY-###   e.g. 2026-048, OM-INV-2026-052
    ("BIGWIG",
     r"(?:(?P<prefix>OM)-(?P<type>INV)-)?(?P<year>\d{4})-(?P<number>\d{3})",
     "Bigwig: YYYY-NNN or OM-INV-YYYY-NNN"),

    # ── BIZ SOLUTIONS ───────────────────────────────────────────────────────
    # BIS/INV/YYYY/### or YYYY-### or YYYY/###   e.g. BIS/INV/2026/143
    ("BIZ SOLUTIONS",
     r"(?:(?P<prefix>BIS)/(?P<type>INV)/)?(?P<year>\d{4})(?P<sep>[-/])(?P<number>\d{3})",
     "Biz Solutions: BIS/INV/YYYY/NNN or YYYY-NNN"),

    # ── BLENX ───────────────────────────────────────────────────────────────
    # BL-####### or BL/####### or 7-digit   e.g. BL-1459973, 1464109
    ("BLENX",
     r"(?:(?P<prefix>BL)(?P<sep>[-/]))?(?P<number>\d{7})",
     "Blenx: BL-NNNNNNN or NNNNNNN"),

    # ── BLUE BIRD COMPANY ───────────────────────────────────────────────────
    # B##-####   e.g. B26-0379
    ("BLUE BIRD",
     r"(?P<prefix>B)(?P<year>\d{2})-(?P<number>\d{4})",
     "Blue Bird: BYY-NNNN"),

    # ── BROTHERHOOD SHOP ────────────────────────────────────────────────────
    # 9-digit numeric   e.g. 103067853
    ("BROTHERHOOD", r"(?P<number>\d{9})", "Brotherhood: 9-digit"),

    # ── CGT PVT LTD ─────────────────────────────────────────────────────────
    # numeric or INV NNNNNN   e.g. 62695, INV 062698
    ("CGT",
     r"(?:(?P<prefix>INV)\s*)?(?P<number>\d{5,6})",
     "CGT: NNNNN or INV NNNNNN"),

    # ── CHEF 2 CHEF ─────────────────────────────────────────────────────────
    # 3-digit numeric   e.g. 205
    ("CHEF 2 CHEF", r"(?P<number>\d{3})", "Chef 2 Chef: 3-digit"),

    # ── CHEMLAB PRIVATE LTD ─────────────────────────────────────────────────
    # CL-I-#### or 4-digit numeric   e.g. CL-I-5019, 5222
    ("CHEMLAB",
     r"(?:(?P<prefix>CL)-(?P<type>I)-)?(?P<number>\d{4})",
     "Chemlab: CL-I-NNNN or NNNN"),

    # ── COLORLAND ───────────────────────────────────────────────────────────
    # DC/N/YY/##### or DCL/N/YY/#####   e.g. DC/5/26/04567, DCL/2/26/38566
    ("COLORLAND",
     r"(?:(?P<prefix>DCL?)(?P<sep>/)(?P<sub>\d)/(?P<year>\d{2})/)?(?P<number>\d{5})",
     "Colorland: DC/N/YY/NNNNN or NNNNN"),

    # ── COSMOPOLITAN ────────────────────────────────────────────────────────
    # 6-digit or INV NNNNNNN   e.g. 181020, INV 0177504
    ("COSMOPOLITAN",
     r"(?:(?P<prefix>INV)\s*)?(?P<number>\d{6,7})",
     "Cosmopolitan: NNNNNN or INV NNNNNNN"),

    # ── COUNTLINE ───────────────────────────────────────────────────────────
    # YY-###### or YY/######   e.g. 26-005307, 26/005125
    ("COUNTLINE",
     r"(?P<year>\d{2})(?P<sep>[-/])(?P<number>\d{6})",
     "Countline: YY-NNNNNN"),

    # ── CSC ENGINEERING ─────────────────────────────────────────────────────
    # ####-YYYY or IN-####-YYYY   e.g. 1441-2026, IN-1449-2026
    ("CSC ENGINEERING",
     r"(?:(?P<prefix>IN)-)?(?P<number>\d{4})-(?P<year>\d{4})",
     "CSC Engineering: NNNN-YYYY or IN-NNNN-YYYY"),

    # ── D BLUE PRIVATE LTD ──────────────────────────────────────────────────
    # IN####### (IN + 7 digits)   e.g. IN0483445
    ("D BLUE",
     r"(?P<prefix>IN)(?P<number>\d{7})",
     "D Blue: INNNNNNNN"),

    # ── DHIAAGU ─────────────────────────────────────────────────────────────
    # BA-#####-YY   e.g. BA-19482-26
    ("DHIAAGU",
     r"(?P<prefix>BA)-(?P<number>\d{5})-(?P<year>\d{2})",
     "Dhiaagu: BA-NNNNN-YY"),

    # ── DIANA TRADING ───────────────────────────────────────────────────────
    # 5-digit numeric   e.g. 12021
    ("DIANA TRADING", r"(?P<number>\d{5})", "Diana Trading: 5-digit"),

    # ── DITECH CONSTRUCTION ─────────────────────────────────────────────────
    # DIN-####   e.g. DIN-0144
    ("DITECH",
     r"(?P<prefix>DIN)-(?P<number>\d{4})",
     "Ditech: DIN-NNNN"),

    # ── EASTERN AND ALLIED ──────────────────────────────────────────────────
    # FV#####/FV##### or FV#####-FV#####   e.g. FV20103/FV20104
    ("EASTERN",
     r"(?P<prefix>FV)(?P<number1>\d{5})(?P<sep>[/-])(?:FV)?(?P<number2>\d{5})",
     "Eastern & Allied: FV#####/FV#####"),

    # ── ECOCOAST ────────────────────────────────────────────────────────────
    # EC#####   e.g. EC25886
    ("ECOCOAST",
     r"(?P<prefix>EC)(?P<number>\d{5})",
     "Ecocoast: ECNNNNN"),

    # ── EMPARAL PVT LTD ─────────────────────────────────────────────────────
    # 4-digit or EYY/####   e.g. 2437, E26/2251
    ("EMPARAL",
     r"(?:(?P<prefix>E)(?P<year>\d{2})/)?(?P<number>\d{4})",
     "Emparal: NNNN or EYY/NNNN"),

    # ── EURO MARKETING PVT LTD ──────────────────────────────────────────────
    # 9-digit starting 1   e.g. 100188337
    ("EURO MARKETING",
     r"(?P<number>1\d{8})",
     "Euro Marketing: 9-digit (starts 1XXXXXXXX)"),

    # ── F&B EVENING STORE ───────────────────────────────────────────────────
    # YY/ES/ICS/#### or YY-ES-ICS-####   e.g. 26/ES/ICS/5639
    ("F&B EVENING",
     r"(?P<year>\d{2})(?P<sep>[-/])(?P<dept>ES)(?P<sep2>[-/])(?P<type>ICS)(?P<sep3>[-/])(?P<number>\d{4})",
     "F&B Evening: YY/ES/ICS/NNNN"),

    # ── FANTASY PVT LTD ─────────────────────────────────────────────────────
    # #####-YYYY   e.g. 02972-2026
    ("FANTASY",
     r"(?P<number>\d{5})-(?P<year>\d{4})",
     "Fantasy: NNNNN-YYYY"),

    # ── FEATHER INVESTMENT ──────────────────────────────────────────────────
    # 4-digit numeric   e.g. 1032
    ("FEATHER", r"(?P<number>\d{4})", "Feather Investment: 4-digit"),

    # ── FOOD SPECIALTY MALDIVES ─────────────────────────────────────────────
    # 5-digit numeric   e.g. 92542
    ("FOOD SPECIALTY", r"(?P<number>\d{5})", "Food Specialty: 5-digit"),

    # ── FRELLA INTERNATIONAL ────────────────────────────────────────────────
    # 3-digit or INV/YYYY/#####   e.g. 450, INV/2026/00378
    ("FRELLA",
     r"(?:(?P<prefix>INV)/(?P<year>\d{4})/)?(?P<number>\d{3,5})",
     "Frella: NNN or INV/YYYY/NNNNN"),

    # ── FRONTIER MARKETS ENTERPRISES ────────────────────────────────────────
    # 4-5 digit or YYYY/##### or INV-YYYY-#####
    ("FRONTIER MARKETS",
     r"(?:(?P<prefix>INV)-)?(?P<year>\d{4})(?P<sep>[-/])?(?P<number>\d{4,5})",
     "Frontier Markets: NNNNN or YYYY/NNNNN or INV-YYYY-NNNNN"),

    # ── GGT ─────────────────────────────────────────────────────────────────
    # INV-YYYY-######   e.g. INV-2026-089283
    ("GGT",
     r"(?P<prefix>INV)-(?P<year>\d{4})-(?P<number>\d{6})",
     "GGT: INV-YYYY-NNNNNN"),

    # ── GLOBAL RIWA ─────────────────────────────────────────────────────────
    # YY/#### or YY-#### or 4-digit   e.g. 26/0494, 26-1067, 1039
    ("GLOBAL RIWA",
     r"(?:(?P<year>\d{2})(?P<sep>[-/]))?(?P<number>\d{4})",
     "Global Riwa: YY/NNNN or NNNN"),

    # ── GRAPE EXPECTATIONS ──────────────────────────────────────────────────
    # 5-digit numeric   e.g. 64264
    ("GRAPE", r"(?P<number>\d{5})", "Grape Expectations: 5-digit"),

    # ── GREEN PATH ──────────────────────────────────────────────────────────
    # 4-digit numeric   e.g. 1396
    ("GREEN PATH", r"(?P<number>\d{4})", "Green Path: 4-digit"),

    # ── HAPPY MARKET ────────────────────────────────────────────────────────
    # 6-digit numeric   e.g. 147228
    ("HAPPY MARKET", r"(?P<number>\d{6})", "Happy Market: 6-digit"),

    # ── HARDWARE LAB ────────────────────────────────────────────────────────
    # 3-4 digit numeric   e.g. 583, 1064
    ("HARDWARE LAB", r"(?P<number>\d{3,4})", "Hardware Lab: 3-4 digit"),

    # ── HASSAN MARINE EQUIPMENT SHOP ────────────────────────────────────────
    # 8-digit starting 1002   e.g. 10024616
    ("HASSAN MARINE", r"(?P<number>1002\d{4})", "Hassan Marine: 1002NNNN"),

    # ── HENMAN TRADING ──────────────────────────────────────────────────────
    # SD#####   e.g. SD14341
    ("HENMAN",
     r"(?P<prefix>SD)(?P<number>\d{5})",
     "Henman: SDNNNNN"),

    # ── HMHI COMPANY ────────────────────────────────────────────────────────
    # 3-digit numeric   e.g. 396
    ("HMHI", r"(?P<number>\d{3})", "HMHI: 3-digit"),

    # ── HORIZON FISHERIES ───────────────────────────────────────────────────
    # 6-digit or FDS-#########   e.g. 112275, FDS-000006514
    ("HORIZON",
     r"(?:(?P<prefix>FDS)-)?(?P<number>\d{6,9})",
     "Horizon Fisheries: NNNNNN or FDS-NNNNNNNNN"),

    # ── HYDROGLOBAL ─────────────────────────────────────────────────────────
    # 3-4 digit or HGU/####   e.g. 880, HGU/7460
    ("HYDROGLOBAL",
     r"(?:(?P<prefix>HGU)/)?(?P<number>\d{3,4})",
     "Hydroglobal: NNN or HGU/NNNN"),

    # ── IFS GLOBAL ──────────────────────────────────────────────────────────
    # 5-digit or GS1-##### or INV #####   e.g. 11698, GS1-11960, INV 11813
    ("IFS GLOBAL",
     r"(?:(?P<prefix>GS[I1])-|(?P<pfx2>INV)\s*)?(?P<number>\d{5})",
     "IFS Global: NNNNN or GS1-NNNNN or INV NNNNN"),

    # ── ILAA MALDIVES ───────────────────────────────────────────────────────
    # INV-##-########   e.g. INV-20-00010321
    ("ILAA MALDIVES",
     r"(?P<prefix>INV)-(?P<sub>\d{2})-(?P<number>\d{8})",
     "Ilaa Maldives: INV-NN-NNNNNNNN"),

    # ── INNOVO MALDIVES ─────────────────────────────────────────────────────
    # 4-digit or ##-####   e.g. 5565, 18-5881
    ("INNOVO",
     r"(?:(?P<prefix>\d{2})-)?(?P<number>\d{4})",
     "Innovo: NNNN or NN-NNNN"),

    # ── INTERNATIONAL FOOD SOLUTION ─────────────────────────────────────────
    # 6-digit or MSI-######   e.g. 270568, MSI-577020
    ("INTERNATIONAL FOOD",
     r"(?:(?P<prefix>MSI)-)?(?P<number>\d{6})",
     "Int'l Food Solution: NNNNNN or MSI-NNNNNN"),

    # ── JIM TRADERS ─────────────────────────────────────────────────────────
    # JIM###### or JM###### or HN######   e.g. JIM000488, JM000396, HN000518
    ("JIM TRADERS",
     r"(?P<prefix>JIM|JM|HN)(?P<number>\d{6})",
     "Jim Traders: JIM/JM/HN + NNNNNN"),

    # ── KAILAAN TRENDS ──────────────────────────────────────────────────────
    # 3-digit or INV/YYYY/######   e.g. 845, INV/2026/000755
    ("KAILAAN",
     r"(?:(?P<prefix>INV)/(?P<year>\d{4})/)?(?P<number>\d{3,6})",
     "Kailaan: NNN or INV/YYYY/NNNNNN"),

    # ── LILY ENTERPRISES ────────────────────────────────────────────────────
    # ####/YY or ####-YY   e.g. 0551/26, 1073-26
    ("LILY ENTERPRISES",
     r"(?P<number>\d{4})(?P<sep>[-/])(?P<year>\d{2})",
     "Lily Enterprises: NNNN/YY"),

    # ── LILY INTERNATIONAL ──────────────────────────────────────────────────
    # 6-digit or ####-YY or INV-#######   e.g. 501259, 1140-26, INV-0509456
    ("LILY INTERNATIONAL",
     r"(?:(?P<prefix>INV)-)?(?P<number>\d{6,7})|(?P<num2>\d{4})-(?P<year>\d{2})",
     "Lily International: NNNNNN or NNNN-YY or INV-NNNNNNN"),

    # ── LOLLO WHOLESALE ─────────────────────────────────────────────────────
    # 5-digit numeric   e.g. 10165
    ("LOLLO", r"(?P<number>\d{5})", "Lollo Wholesale: 5-digit"),

    # ── LOTUS FIHAARA ───────────────────────────────────────────────────────
    # 7-digit or LFF-#######   e.g. 1057441, LFF-1052549
    ("LOTUS",
     r"(?:(?P<prefix>LFF)-)?(?P<number>\d{7})",
     "Lotus Fihaara: NNNNNNN or LFF-NNNNNNN"),

    # ── LYCORN MALDIVES ─────────────────────────────────────────────────────
    # 4-digit numeric   e.g. 1064
    ("LYCORN", r"(?P<number>\d{4})", "Lycorn: 4-digit"),

    # ── MAA HARDWARE ────────────────────────────────────────────────────────
    # 2-3 digit or YYYY/MMM-##   e.g. 122, 2026/FEB-87
    ("MAA HARDWARE",
     r"(?:(?P<year>\d{4})/(?P<month>[A-Z]{3})-)?(?P<number>\d{2,3})",
     "Maa Hardware: NN or YYYY/MMM-NN"),

    # ── MAM PETTY CASH ──────────────────────────────────────────────────────
    # 4-5 digit or GTS-INV-######   e.g. 11638, GTS-INV-000284
    ("MAM PETTY CASH",
     r"(?:(?P<prefix>GTS)-(?P<type>INV)-)?(?P<number>\d{4,6})",
     "MAM Petty Cash: NNNNN or GTS-INV-NNNNNN"),

    # ── MARITIME & MERCANTILE ────────────────────────────────────────────────
    # 8-digit or ARINV-########   e.g. 10030348, ARINV-10030993
    ("MARITIME",
     r"(?:(?P<prefix>ARINV)-)?(?P<number>\d{8})",
     "Maritime & Mercantile: NNNNNNNN or ARINV-NNNNNNNN"),

    # ── MIAAF PVT LTD ───────────────────────────────────────────────────────
    # 4-digit numeric   e.g. 4272
    ("MIAAF", r"(?P<number>\d{4})", "Miaaf: 4-digit"),

    # ── MISRAAB ─────────────────────────────────────────────────────────────
    # 7-digit starting 300   e.g. 3003749
    ("MISRAAB", r"(?P<number>300\d{4})", "Misraab: 300NNNN"),

    # ── MIYAMI TRADERS ──────────────────────────────────────────────────────
    # 6-digit numeric   e.g. 408528
    ("MIYAMI", r"(?P<number>\d{6})", "Miyami Traders: 6-digit"),

    # ── MMX TRADERS ─────────────────────────────────────────────────────────
    # 3-4 digit numeric   e.g. 161, 0260
    ("MMX TRADERS", r"(?P<number>\d{3,4})", "MMX Traders: NNN or NNNN"),

    # ── NASREENA ABDULLA ────────────────────────────────────────────────────
    # YYYY-####   e.g. 2026-0042
    ("NASREENA",
     r"(?P<year>\d{4})-(?P<number>\d{4})",
     "Nasreena Abdulla: YYYY-NNNN"),

    # ── NOVELTY BOOKSHOP ────────────────────────────────────────────────────
    # 3-digit numeric   e.g. 716
    ("NOVELTY", r"(?P<number>\d{3})", "Novelty Bookshop: 3-digit"),

    # ── OSMOSIS ASIA ────────────────────────────────────────────────────────
    # 3-digit or OSM/INV/YY/####   e.g. 370, OSM/INV/26/0285
    ("OSMOSIS",
     r"(?:(?P<prefix>OSM)/(?P<type>INV)/(?P<year>\d{2})/)?(?P<number>\d{3,4})",
     "Osmosis: NNN or OSM/INV/YY/NNNN"),

    # ── POISE DISTRIBUTORS ──────────────────────────────────────────────────
    # 4-digit or YYYY-#####   e.g. 1514, 2026-04907
    ("POISE",
     r"(?:(?P<year>\d{4})-)?(?P<number>\d{4,5})",
     "Poise: NNNN or YYYY-NNNNN"),

    # ── PREMIUM SUPPLIES ────────────────────────────────────────────────────
    # YY-#### or 6-digit   e.g. 26-1542, 262815
    ("PREMIUM",
     r"(?:(?P<year>\d{2})-)?(?P<number>\d{4,6})",
     "Premium Supplies: YY-NNNN or NNNNNN"),

    # ── PRINTLAB ────────────────────────────────────────────────────────────
    # YYYY-####   e.g. 2026-2725
    ("PRINTLAB",
     r"(?P<year>\d{4})-(?P<number>\d{4})",
     "Printlab: YYYY-NNNN"),

    # ── PROCURE PLUS ────────────────────────────────────────────────────────
    # 8-digit starting 170   e.g. 17012424
    ("PROCURE PLUS", r"(?P<number>170\d{5})", "Procure Plus: 170NNNNN"),

    # ── RESTOFAIR MALDIVES ──────────────────────────────────────────────────
    # 8-digit starting 180   e.g. 18000736
    ("RESTOFAIR", r"(?P<number>180\d{5})", "Restofair: 180NNNNN"),

    # ── RESUINSA EXPERIENCES ────────────────────────────────────────────────
    # EX:######   e.g. EX:007631
    ("RESUINSA",
     r"(?P<prefix>EX):(?P<number>\d{6})",
     "Resuinsa: EX:NNNNNN"),

    # ── S&J SALES CORPORATION ───────────────────────────────────────────────
    # TRV-YY-# or YY-## or YY/##   e.g. TRV-26-7, 26-46, 25/512
    ("S&J SALES",
     r"(?:(?P<prefix>TRV)-(?P<year>\d{2})-(?P<number>\d+))|(?:(?P<yr2>\d{2})(?P<sep>[-/])(?P<num2>\d+))",
     "S&J Sales: TRV-YY-N or YY-NN or YY/NNN"),

    # ── SAFCO INTERNATIONAL ─────────────────────────────────────────────────
    # SF/EX/######   e.g. SF/EX/132050
    ("SAFCO",
     r"(?P<prefix>SF)/(?P<type>EX)/(?P<number>\d{6})",
     "Safco: SF/EX/NNNNNN"),

    # ── SALESCO PVT LTD ─────────────────────────────────────────────────────
    # SH1/YY/#### or SH1 YY #### or 4-digit   e.g. SH1/26/1039, 4659
    ("SALESCO",
     r"(?:(?P<prefix>SH1)[\s/](?P<year>\d{2})[\s/])?(?P<number>\d{4})",
     "Salesco: SH1/YY/NNNN or NNNN"),

    # ── SAMAN TRADING ───────────────────────────────────────────────────────
    # 4-digit or INV/YYYY/#####   e.g. 2541, INV/2026/04019
    ("SAMAN TRADING",
     r"(?:(?P<prefix>INV)/(?P<year>\d{4})/)?(?P<number>\d{4,5})",
     "Saman Trading: NNNN or INV/YYYY/NNNNN"),

    # ── SAWHNEY FOOD STAFF TRADING ──────────────────────────────────────────
    # 7-digit starting 110   e.g. 1102989, 1105624
    ("SAWHNEY", r"(?P<number>110\d{4})", "Sawhney: 110NNNN"),

    # ── SEAFOOD ENTERPRISES ─────────────────────────────────────────────────
    # 4-digit or YY-#### or YY/####   e.g. 2316, 26-3899, 26/2001
    ("SEAFOOD ENTERPRISES",
     r"(?:(?P<year>\d{2})(?P<sep>[-/]))?(?P<number>\d{4})",
     "Seafood Enterprises: NNNN or YY-NNNN"),

    # ── SEAGEAR ─────────────────────────────────────────────────────────────
    # 6-digit or INV-######/SG/YYYY   e.g. 106234, INV-105265/SG/2026
    ("SEAGEAR",
     r"(?:(?P<prefix>INV)-)?(?P<number>\d{6})(?:/(?P<suffix>SG)/(?P<year>\d{4}))?",
     "Seagear: NNNNNN or INV-NNNNNN/SG/YYYY"),

    # ── SEALANDS PVT LTD ────────────────────────────────────────────────────
    # 5-digit numeric   e.g. 22520
    ("SEALANDS", r"(?P<number>\d{5})", "Sealands: 5-digit"),

    # ── SHENZHEN SANHE E-COMMERCE ───────────────────────────────────────────
    # AF-YYYYMMDD-#####   e.g. AF-20260303-10306
    ("SHENZHEN SANHE",
     r"(?P<prefix>AF)-(?P<date>\d{8})-(?P<number>\d{5})",
     "Shenzhen Sanhe: AF-YYYYMMDD-NNNNN"),

    # ── SIMDI CONSUMER PRODUCTS ─────────────────────────────────────────────
    # 4-digit numeric   e.g. 1574, 4396
    ("SIMDI CONSUMER", r"(?P<number>\d{4})", "Simdi Consumer: 4-digit"),

    # ── SONEE HARDWARE ──────────────────────────────────────────────────────
    # 7-digit or CINV-####### or INV #######   e.g. 1759865, CINV-1744936
    ("SONEE",
     r"(?:(?P<prefix>CINV|INV)[-\s])?(?P<number>\d{7})",
     "Sonee: NNNNNNN or CINV-NNNNNNN or INV NNNNNNN"),

    # ── STANDARD & ORIGIN MARKETING ─────────────────────────────────────────
    # 4-digit numeric   e.g. 2169
    ("STANDARD", r"(?P<number>\d{4})", "Standard & Origin: 4-digit"),

    # ── SUN FRONT LIGHTING ──────────────────────────────────────────────────
    # #####-YY or ######-YY   e.g. 47873-26, 200651-26
    ("SUN FRONT",
     r"(?P<number>\d{5,6})-(?P<year>\d{2})",
     "Sun Front Lighting: NNNNN-YY or NNNNNN-YY"),

    # ── SUPER POWER CO. LTD ─────────────────────────────────────────────────
    # 1-3 digit numeric   e.g. 20
    ("SUPER POWER", r"(?P<number>\d{1,3})", "Super Power: 1-3 digit"),

    # ── THANDIYA ────────────────────────────────────────────────────────────
    # ###/YYYY   e.g. 097/2026
    ("THANDIYA",
     r"(?P<number>\d{3})/(?P<year>\d{4})",
     "Thandiya: NNN/YYYY"),

    # ── TOTAL BEVERAGES ─────────────────────────────────────────────────────
    # 3-digit or INV-YYYY-#####   e.g. 391, INV-2026-00426
    ("TOTAL BEVERAGES",
     r"(?:(?P<prefix>INV)-(?P<year>\d{4})-)?(?P<number>\d{3,5})",
     "Total Beverages: NNN or INV-YYYY-NNNNN"),

    # ── TRADE MALDIVES ──────────────────────────────────────────────────────
    # 3-digit or TM-YY-######   e.g. 566, TM-26-000622
    ("TRADE MALDIVES",
     r"(?:(?P<prefix>TM)-(?P<year>\d{2})-)?(?P<number>\d{3,6})",
     "Trade Maldives: NNN or TM-YY-NNNNNN"),

    # ── UNID HARDWARE ───────────────────────────────────────────────────────
    # ####/YY or ####-YY   e.g. 4210/26, 4259-26
    ("UNID HARDWARE",
     r"(?P<number>\d{4})(?P<sep>[-/])(?P<year>\d{2})",
     "Unid Hardware: NNNN/YY"),

    # ── VB BROTHERS ─────────────────────────────────────────────────────────
    # 5-digit numeric   e.g. 19144
    ("VB BROTHERS", r"(?P<number>\d{5})", "VB Brothers: 5-digit"),

    # ── VELIGAA HARDWARE ────────────────────────────────────────────────────
    # YY/CORP/ICS/#### or YY CORP ICS ####   e.g. 26/CORP/ICS/1959
    ("VELIGAA",
     r"(?P<year>\d{2})(?P<sep>[/ ])CORP(?P<sep2>[/ ])ICS(?P<sep3>[/ ])(?P<number>\d{4})",
     "Veligaa: YY/CORP/ICS/NNNN"),

    # ── VILLA HAKATHA ────────────────────────────────────────────────────────
    # 6-digit numeric   e.g. 406594
    ("VILLA HAKATHA", r"(?P<number>\d{6})", "Villa Hakatha: 6-digit"),

    # ── VIRTUS GROUP ────────────────────────────────────────────────────────
    # YYYY-##### or RE-YYYY-#####   e.g. 2026-11031, RE-2026-10240
    ("VIRTUS",
     r"(?:(?P<prefix>RE)-)?(?P<year>\d{4})-(?P<number>\d{5})",
     "Virtus Group: YYYY-NNNNN or RE-YYYY-NNNNN"),

    # ── WANKUN (HANGZHOU) ────────────────────────────────────────────────────
    # 8-digit or WK########/########   e.g. 25092401, WK26013101/26022401
    ("WANKUN",
     r"(?:(?P<prefix>WK)(?P<n1>\d{8})/(?P<n2>\d{8}))|(?P<num>\d{8})",
     "Wankun: NNNNNNNN or WKNNNNNNNN/NNNNNNNN"),

    # ── ZIP INVESTMENT ───────────────────────────────────────────────────────
    # 5-digit numeric   e.g. 15370
    ("ZIP INVESTMENT", r"(?P<number>\d{5})", "Zip Investment: 5-digit"),

    # ── GENERIC FALLBACKS ────────────────────────────────────────────────────
    # Alphanumeric with slash: XX/YYYY/NNN
    ("*",
     r"(?P<prefix>[A-Z]{1,4})/(?P<year>\d{4})/(?P<seq>\d{1,6})",
     "Alphanumeric slash: XX/YYYY/NNN"),

    # Generic 5-10 digit
    ("*",
     r"(?P<number>\d{5,10})",
     "Generic 5-10 digit invoice"),
]


# Known correct prefixes per supplier — used to fix OCR-garbled prefix letters.
SUPPLIER_INVOICE_PREFIXES: Dict[str, List[str]] = {
    "ADK PHARMACEUTICAL": [],           # prefix is "14" (numeric, not garbled)
    "ALAAN":            [],
    "ALIHAVA":          ["AH"],
    "AMAGI":            [],
    "AYAAN":            ["AY"],
    "BLENX":            ["BL"],
    "BLUE BIRD":        ["B"],
    "BIZ SOLUTIONS":    ["BIS"],
    "CGT":              ["INV"],
    "CHEMLAB":          ["CL"],
    "COLORLAND":        ["DC", "DCL"],
    "CSC ENGINEERING":  ["IN"],
    "D BLUE":           ["IN"],
    "DHIAAGU":          ["BA"],
    "DITECH":           ["DIN"],
    "ECOCOAST":         ["EC"],
    "EASTERN":          ["FV"],
    "F&B EVENING":      ["ES"],
    "FRELLA":           ["INV"],
    "GGT":              ["INV"],
    "HASSAN MARINE":    [],
    "HENMAN":           ["SD"],
    "HORIZON":          ["FDS"],
    "HYDROGLOBAL":      ["HGU"],
    "IFS GLOBAL":       ["GS1", "GSI", "INV"],
    "ILAA MALDIVES":    ["INV"],
    "JIM TRADERS":      ["JIM", "JM", "HN"],
    "KAILAAN":          ["INV"],
    "LOTUS":            ["LFF"],
    "MAM PETTY CASH":   ["GTS"],
    "MARITIME":         ["ARINV"],
    "OSMOSIS":          ["OSM"],
    "RESUINSA":         ["EX"],
    "S&J SALES":        ["TRV"],
    "SAFCO":            ["SF"],
    "SALESCO":          ["SH1"],
    "SAMAN TRADING":    ["INV"],
    "SEAGEAR":          ["INV"],
    "SHENZHEN SANHE":   ["AF"],
    "SONEE":            ["CINV", "INV"],
    "TOTAL BEVERAGES":  ["INV"],
    "TRADE MALDIVES":   ["TM"],
    "VIRTUS":           ["RE"],
    "WANKUN":           ["WK"],
}


def correct_invoice_number(
    raw_invoice: str,
    supplier_name: str,
) -> str:
    """
    Given a raw OCR-extracted invoice string and the identified supplier name,
    attempts to fix garbled prefix letters using known supplier invoice formats.

    Returns the corrected invoice string, or the original if no fix is possible.
    """
    if not raw_invoice:
        return raw_invoice

    sup_upper = supplier_name.upper()

    correct_prefixes: List[str] = []
    for key, prefixes in SUPPLIER_INVOICE_PREFIXES.items():
        if key in sup_upper:
            correct_prefixes = prefixes
            break

    if not correct_prefixes:
        return raw_invoice

    for sup_key, pattern, _desc in INVOICE_PATTERNS:
        if sup_key != "*" and sup_key.upper() not in sup_upper:
            continue

        m = re.match(pattern, raw_invoice.upper().strip(), re.IGNORECASE)
        if not m:
            continue

        groups = m.groupdict()
        raw_prefix = groups.get("prefix", "")
        sep = groups.get("sep", "-")
        number = groups.get("number", "")

        if not raw_prefix or not number:
            return raw_invoice

        best_prefix = raw_prefix
        best_dist = 999
        for cp in correct_prefixes:
            from rapidfuzz.distance import Levenshtein
            d = Levenshtein.distance(raw_prefix.upper(), cp.upper())
            if d < best_dist:
                best_dist = d
                best_prefix = cp

        if best_dist <= 2:
            corrected = f"{best_prefix}{sep}{number}"
            return corrected

    return raw_invoice


# ---------------------------------------------------------------------------
# HOW TO ADD YOUR OWN INVOICE FORMATS
# ---------------------------------------------------------------------------
# 1. Add an entry to INVOICE_PATTERNS:
#
#    ("YOUR SUPPLIER KEYWORD",
#     r"(?P<prefix>[A-Z]{1,4})(?P<sep>[-/])(?P<number>\d{4,8})",
#     "Human description: PREFIX-NNNNN"),
#
# 2. Add an entry to SUPPLIER_INVOICE_PREFIXES:
#
#    "YOUR SUPPLIER KEYWORD": ["PFX"],
#
# Example — "HAPPY MARKET" with invoices like "HM-00123":
#
#   ("HAPPY MARKET",
#    r"(?P<prefix>[A-Z]{2})(?P<sep>-)(?P<number>\d{4,6})",
#    "Happy Market: HM-NNNNN"),
#
#   "HAPPY MARKET": ["HM"],
