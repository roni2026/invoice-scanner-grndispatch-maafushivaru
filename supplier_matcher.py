# supplier_matcher.py
# Advanced supplier name matching strategies for Maafushivaru Hub
# Provides multiple algorithms selectable from Settings

import re
import logging
from typing import Optional, Tuple, List, Dict
from rapidfuzz import process, fuzz
from difflib import SequenceMatcher


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", text.upper())).strip()


def _tokenize(name: str) -> List[str]:
    return [w for w in _normalize(name).split() if len(w) >= 2]


# ---------------------------------------------------------------------------
# STRATEGY 1: Progressive Prefix Word Matching (NEW — as requested)
# ---------------------------------------------------------------------------
def match_progressive_prefix(
    text: str,
    candidates: List[str],
    aliases: Dict[str, List[str]],
    threshold: int = 70,
) -> Tuple[Optional[str], float]:
    """
    For each candidate, build progressive prefix groups:
      EURO MARKETING PVT LTD → try "EURO", then "EURO MARKETING",
      then "EURO MARKETING PVT", then "EURO MARKETING PVT LTD"

    The LONGEST matching prefix wins.
    For 3-4 word names, the last word is treated as optional.
    If only the first word matches and nothing longer does, that first-word
    match is returned (low confidence).
    If NO word matches at all, returns None.
    """
    upper_text = _normalize(text)

    # Build a flat list: (candidate_string, canonical_name)
    all_cands: List[Tuple[str, str]] = []
    for sup in candidates:
        all_cands.append((sup, sup))
    for main, alias_list in aliases.items():
        for alias in alias_list:
            all_cands.append((alias, main))

    best_canonical: Optional[str] = None
    best_score: float = 0.0
    best_match_length: int = 0   # how many words matched in the winning prefix

    first_word_fallback: Optional[str] = None
    first_word_fallback_score: float = 0.0

    for raw_cand, canonical in all_cands:
        words = _tokenize(raw_cand)
        if not words:
            continue

        n = len(words)
        # For names with 3+ words, last word is optional — try without it too
        max_optional_drop = 1 if n >= 3 else 0

        found_first = False

        # Try from longest prefix down to just the first word
        for prefix_len in range(n, 0, -1):
            prefix = " ".join(words[:prefix_len])
            if not prefix:
                continue

            # Check if this prefix appears as a whole-word sequence in text
            pattern = r"\b" + r"\s+".join(re.escape(w) for w in words[:prefix_len]) + r"\b"
            match = re.search(pattern, upper_text)

            if match:
                if prefix_len == 1:
                    # Only first word matched
                    found_first = True
                    score = 40.0 + (10.0 if n == 1 else 0.0)
                    if score > first_word_fallback_score:
                        first_word_fallback_score = score
                        first_word_fallback = canonical
                else:
                    # Multi-word match
                    # Score based on how many words of the full name matched
                    # If we dropped the last word(s), slight penalty
                    dropped = n - prefix_len
                    if dropped > max_optional_drop:
                        continue  # too many words missing — not a reliable match

                    coverage = prefix_len / n
                    score = 60.0 + (coverage * 40.0)

                    if prefix_len > best_match_length or score > best_score:
                        best_match_length = prefix_len
                        best_score = score
                        best_canonical = canonical
                    break  # found longest matching prefix for this candidate

    # Decision logic
    if best_canonical and best_score >= threshold:
        return best_canonical, best_score

    # First-word fallback only if nothing better found
    if first_word_fallback and best_canonical is None:
        return first_word_fallback, first_word_fallback_score

    return best_canonical, best_score


# ---------------------------------------------------------------------------
# STRATEGY 2: Unique Word Index (existing logic, extracted here cleanly)
# ---------------------------------------------------------------------------
def match_unique_word_index(
    text: str,
    candidates: List[str],
    aliases: Dict[str, List[str]],
    threshold: int = 85,
) -> Tuple[Optional[str], float]:
    """
    Builds an index of words that appear in only ONE candidate.
    If those unique words appear in the text, it's a strong signal.
    """
    from collections import Counter

    all_cands = list(candidates)
    for sub in aliases.values():
        all_cands.extend(sub)

    word_freq: Counter = Counter()
    cand_words: Dict[str, set] = {}
    for c in all_cands:
        words = set(re.sub(r"[^A-Z0-9 ]", " ", c.upper()).split())
        words = {w for w in words if len(w) >= 3}
        cand_words[c] = words
        word_freq.update(words)

    unique_index: Dict[str, set] = {}
    for c, words in cand_words.items():
        unique_index[c] = {w for w in words if word_freq[w] == 1}

    upper = text.upper()
    best_name: Optional[str] = None
    best_score: float = 0.0

    for c in all_cands:
        unique_words = unique_index.get(c, set())
        if not unique_words:
            continue
        hits = sum(
            1 for w in unique_words
            if re.search(r"\b" + re.escape(w) + r"\b", upper)
        )
        if hits > 0:
            score = (hits / len(unique_words)) * 100.0
            if score > best_score:
                best_score = score
                best_name = c

    if best_name:
        # Resolve alias
        for main, alias_list in aliases.items():
            if best_name in alias_list:
                best_name = main
                break
        return best_name, best_score

    return None, 0.0


# ---------------------------------------------------------------------------
# STRATEGY 3: RapidFuzz Token Sort (existing primary fuzzy method)
# ---------------------------------------------------------------------------
def match_rapidfuzz_token_sort(
    text: str,
    candidates: List[str],
    aliases: Dict[str, List[str]],
    threshold: int = 85,
) -> Tuple[Optional[str], float]:
    """
    Classic rapidfuzz token_sort_ratio on each line vs all candidates.
    """
    all_cands = list(candidates)
    alias_to_main: Dict[str, str] = {}
    for main, alias_list in aliases.items():
        for alias in alias_list:
            all_cands.append(alias)
            alias_to_main[alias] = main
        alias_to_main[main] = main

    upper = text.upper()
    best_name: Optional[str] = None
    best_score: float = 0.0

    for line in upper.splitlines():
        line = line.strip()
        if len(line) < 3:
            continue
        m = process.extractOne(line, all_cands, scorer=fuzz.token_sort_ratio)
        if m and m[1] >= threshold:
            name = alias_to_main.get(m[0], m[0])
            if m[1] > best_score:
                best_score = float(m[1])
                best_name = name

    return best_name, best_score


# ---------------------------------------------------------------------------
# STRATEGY 4: RapidFuzz Partial Ratio
# ---------------------------------------------------------------------------
def match_rapidfuzz_partial(
    text: str,
    candidates: List[str],
    aliases: Dict[str, List[str]],
    threshold: int = 85,
) -> Tuple[Optional[str], float]:
    """
    Uses partial_ratio — good when the supplier name is embedded in a longer line.
    """
    all_cands = list(candidates)
    alias_to_main: Dict[str, str] = {}
    for main, alias_list in aliases.items():
        for alias in alias_list:
            all_cands.append(alias)
            alias_to_main[alias] = main
        alias_to_main[main] = main

    upper = text.upper()
    best_name: Optional[str] = None
    best_score: float = 0.0

    for line in upper.splitlines():
        line = line.strip()
        if len(line) < 3:
            continue
        m = process.extractOne(line, all_cands, scorer=fuzz.partial_ratio)
        if m and m[1] >= threshold:
            name = alias_to_main.get(m[0], m[0])
            if m[1] > best_score:
                best_score = float(m[1])
                best_name = name

    return best_name, best_score


# ---------------------------------------------------------------------------
# STRATEGY 5: Sequence Matcher (difflib — no external dependency)
# ---------------------------------------------------------------------------
def match_sequence_matcher(
    text: str,
    candidates: List[str],
    aliases: Dict[str, List[str]],
    threshold: int = 70,
) -> Tuple[Optional[str], float]:
    """
    Uses Python's built-in difflib.SequenceMatcher.
    Slower but no extra dependencies.
    """
    alias_to_main: Dict[str, str] = {}
    all_cands = list(candidates)
    for main, alias_list in aliases.items():
        for alias in alias_list:
            all_cands.append(alias)
            alias_to_main[alias] = main
        alias_to_main[main] = main

    norm_text = _normalize(text)
    best_name: Optional[str] = None
    best_score: float = 0.0

    for c in all_cands:
        cn = _normalize(c)
        if not cn:
            continue
        ratio = SequenceMatcher(None, norm_text[:2000], cn).ratio() * 100.0
        if ratio >= threshold and ratio > best_score:
            best_score = ratio
            best_name = alias_to_main.get(c, c)

    return best_name, best_score


# ---------------------------------------------------------------------------
# STRATEGY 6: Combined (runs multiple strategies, takes highest-confidence result)
# ---------------------------------------------------------------------------
def match_combined(
    text: str,
    candidates: List[str],
    aliases: Dict[str, List[str]],
    threshold: int = 75,
) -> Tuple[Optional[str], float]:
    """
    Runs all strategies and picks the result with the highest confidence.
    Strategies are weighted: unique-word and progressive-prefix get priority.
    """
    results: List[Tuple[Optional[str], float]] = []

    # Weight multipliers to bias toward more reliable strategies
    weights = {
        "unique":      1.20,
        "progressive": 1.15,
        "token_sort":  1.00,
        "partial":     0.95,
        "sequence":    0.85,
    }

    r_unique = match_unique_word_index(text, candidates, aliases, threshold)
    results.append(("unique", r_unique[0], r_unique[1] * weights["unique"]))

    r_prog = match_progressive_prefix(text, candidates, aliases, threshold)
    results.append(("progressive", r_prog[0], r_prog[1] * weights["progressive"]))

    r_ts = match_rapidfuzz_token_sort(text, candidates, aliases, threshold)
    results.append(("token_sort", r_ts[0], r_ts[1] * weights["token_sort"]))

    r_part = match_rapidfuzz_partial(text, candidates, aliases, threshold)
    results.append(("partial", r_part[0], r_part[1] * weights["partial"]))

    best_strategy = ""
    best_name: Optional[str] = None
    best_score: float = 0.0

    for strategy, name, score in results:
        if name and score > best_score:
            best_score = score
            best_name = name
            best_strategy = strategy

    logging.debug(f"Combined match winner: {best_strategy} → {best_name} ({best_score:.1f}%)")
    return best_name, min(best_score, 100.0)


# ---------------------------------------------------------------------------
# DISPATCHER — called from the main hub
# ---------------------------------------------------------------------------
STRATEGY_FUNCTIONS = {
    "progressive_prefix": match_progressive_prefix,
    "unique_word_index":  match_unique_word_index,
    "rapidfuzz_token_sort": match_rapidfuzz_token_sort,
    "rapidfuzz_partial":  match_rapidfuzz_partial,
    "sequence_matcher":   match_sequence_matcher,
    "combined":           match_combined,
}

STRATEGY_LABELS = {
    "progressive_prefix":   "Progressive Prefix (Word-by-word buildup)",
    "unique_word_index":    "Unique Word Index (Fingerprint matching)",
    "rapidfuzz_token_sort": "RapidFuzz Token Sort (Original)",
    "rapidfuzz_partial":    "RapidFuzz Partial Ratio",
    "sequence_matcher":     "Sequence Matcher (difflib, no dependencies)",
    "combined":             "Combined (All strategies, highest confidence wins)",
}


def dispatch_match(
    strategy: str,
    text: str,
    candidates: List[str],
    aliases: Dict[str, List[str]],
    threshold: int = 75,
) -> Tuple[Optional[str], float]:
    fn = STRATEGY_FUNCTIONS.get(strategy, match_combined)
    return fn(text, candidates, aliases, threshold)