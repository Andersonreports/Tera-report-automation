"""
PGT-A Classification Engine
Auto-detects embryo classification from raw result strings and derives
all patient-facing display fields.
"""

import re

# ── Classification constants ──────────────────────────────────────────────────
EUPLOID        = "EUPLOID"
ANEUPLOID      = "ANEUPLOID"
SEGMENTAL      = "SEGMENTAL"
LOW_MOSAIC     = "LOW_MOSAIC"
HIGH_MOSAIC    = "HIGH_MOSAIC"
COMPLEX_MOSAIC = "COMPLEX_MOSAIC"
FAILED         = "FAILED"

# ── Summary table Result column text ─────────────────────────────────────────
SUMMARY_TEXT = {
    EUPLOID:        "Normal chromosome complement",
    ANEUPLOID:      "Multiple chromosomal abnormalities",
    SEGMENTAL:      "Multiple chromosomal abnormalities",
    LOW_MOSAIC:     "Low level mosaic",
    HIGH_MOSAIC:    "High level mosaic",
    COMPLEX_MOSAIC: "Complex mosaic",
    FAILED:         "No result obtained",
}

# ── Individual embryo page Result field text ──────────────────────────────────
RESULT_TEXT = {
    EUPLOID:        "The embryo contains normal chromosome complement",
    ANEUPLOID:      "The embryo contains abnormal chromosome complement",
    SEGMENTAL:      "The embryo contains abnormal chromosome complement",
    LOW_MOSAIC:     "The embryo contains low level mosaic chromosome complement",
    HIGH_MOSAIC:    "The embryo contains high level mosaic chromosome complement",
    COMPLEX_MOSAIC: "The embryo contains complex mosaic chromosome complement",
    FAILED:         "No result obtained",
}

# Keyword map: pre-classified strings → constant
_KEYWORD_MAP = {
    "EUPLOID":                     EUPLOID,
    "NORMAL":                      EUPLOID,
    "NORMAL CHROMOSOME COMPLEMENT":EUPLOID,
    "ANEUPLOID":                   ANEUPLOID,
    "MULTIPLE CHROMOSOMAL ABNORMALITIES": ANEUPLOID,
    "SEGMENTAL":                   SEGMENTAL,
    "LOW LEVEL MOSAIC":            LOW_MOSAIC,
    "LOW MOSAIC":                  LOW_MOSAIC,
    "HIGH LEVEL MOSAIC":           HIGH_MOSAIC,
    "HIGH MOSAIC":                 HIGH_MOSAIC,
    "COMPLEX MOSAIC":              COMPLEX_MOSAIC,
    "FAILED":                      FAILED,
    "NO RESULT":                   FAILED,
    "NO RESULT OBTAINED":          FAILED,
    "INCONCLUSIVE":                FAILED,
}


def _extract_pct(text):
    """Return numeric percentage from text like '~35%', '40%'. None if absent."""
    m = re.search(r'~?\s*(\d+(?:\.\d+)?)\s*%', text)
    return float(m.group(1)) if m else None


def classify_embryo(raw_result):
    """
    Auto-classify an embryo from a raw result string.

    Classification rules:
      - No abnormality              → EUPLOID
      - Any Loss/Gain               → ANEUPLOID
      - Partial arm change (del/dup)→ SEGMENTAL
      - 30–50 % mosaic              → LOW_MOSAIC
      - 51–80 % mosaic              → HIGH_MOSAIC
      - 3+ chromosomes mosaic       → COMPLEX_MOSAIC
      - >80 % mosaic                → ANEUPLOID
      - <30 % mosaic                → EUPLOID

    Returns dict:
      classification, summary_text, result_text, is_abnormal, is_mosaic
    """
    if not raw_result:
        return _make(EUPLOID)

    s = str(raw_result).strip()
    su = s.upper()

    # 1. Direct keyword match (handles pre-classified values from form/Excel)
    if su in _KEYWORD_MAP:
        return _make(_KEYWORD_MAP[su])

    # 2. Failed / inconclusive
    if any(k in su for k in ("FAILED", "NO RESULT", "INCONCLUSIVE")):
        return _make(FAILED)

    # 3. Mosaic detection — look for percentage OR "mosaic" keyword
    pcts = [_extract_pct(t) for t in re.findall(r'[~\d.%]+', s) if _extract_pct(t) is not None]
    has_mosaic_kw = bool(re.search(r'\bmosaic\b', s, re.IGNORECASE))

    if has_mosaic_kw or pcts:
        pct = max(pcts) if pcts else None

        # Count distinct chromosomes labelled as mosaic
        mosaic_chr_nums = set(re.findall(
            r'(?:mosaic\s+)?(?:[+-]\s*)(1[0-9]|2[0-2]|[1-9])\b',
            s, re.IGNORECASE
        ))
        mosaic_chr_nums |= set(re.findall(
            r'\bmosaic\s+(?:del|dup)\s*\(\s*(1[0-9]|2[0-2]|[1-9])',
            s, re.IGNORECASE
        ))
        if not mosaic_chr_nums:
            # No explicit chr# — count comma-separated tokens loosely
            mosaic_chr_nums = {str(i) for i in range(len(re.split(r'[,;]', s)))}

        n_mosaic_chrs = max(1, len(mosaic_chr_nums))

        # ≥3 mosaic chromosomes → COMPLEX (regardless of %)
        if n_mosaic_chrs >= 3:
            return _make(COMPLEX_MOSAIC)

        # Percentage-based rules
        if pct is not None:
            if pct < 30:
                return _make(EUPLOID)
            if pct > 80:
                return _make(ANEUPLOID)
            if 30 <= pct <= 50:
                return _make(LOW_MOSAIC)
            if 51 <= pct <= 80:
                return _make(HIGH_MOSAIC)

        # Mosaic keyword but no percentage — default LOW
        return _make(LOW_MOSAIC)

    # 4. Segmental change (del/dup notation)
    if re.search(r'\b(del|dup)\s*\(|segmental\s+(loss|gain)\b', s, re.IGNORECASE):
        return _make(SEGMENTAL)

    # 5. Aneuploid (loss/gain mentions or +/- chromosome notation)
    if re.search(
        r'([+-]\s*(?:1[0-9]|2[0-2]|[1-9])\b'
        r'|monosomy|trisomy|nullisomy|tetrasomy'
        r'|\baneuploid\b|\babnormal\b'
        r'|\bloss\b|\bgain\b)',
        s, re.IGNORECASE
    ):
        return _make(ANEUPLOID)

    # 6. Default → EUPLOID
    return _make(EUPLOID)


def _make(cls):
    return {
        "classification": cls,
        "summary_text":   SUMMARY_TEXT[cls],
        "result_text":    RESULT_TEXT[cls],
        "is_abnormal":    cls in (ANEUPLOID, SEGMENTAL, LOW_MOSAIC, HIGH_MOSAIC, COMPLEX_MOSAIC),
        "is_mosaic":      cls in (LOW_MOSAIC, HIGH_MOSAIC, COMPLEX_MOSAIC),
    }


# ─────────────────────────────────────────────────────────────────────────────
# CNV STATUS DERIVATION
# ─────────────────────────────────────────────────────────────────────────────

def derive_chromosome_statuses(raw_result):
    """
    Parse raw result string → {str(1-22): status_code}.
    Default all 22 chromosomes = 'N'.

    Status codes:
      N   – normal (black)
      L   – Loss (red)
      G   – Gain (red)
      SL  – Segmental Loss (red)
      SG  – Segmental Gain (red)
      ML  – Mosaic Loss (orange)
      MG  – Mosaic Gain (orange)
      SML – Segmental Mosaic Loss (orange)
      SMG – Segmental Mosaic Gain (orange)
      NR  – No result / Failed (grey)
    """
    statuses = {str(i): 'N' for i in range(1, 23)}

    if not raw_result:
        return statuses

    s = str(raw_result).strip()
    su = s.upper()

    # FAILED → all NR
    if any(k in su for k in ("FAILED", "NO RESULT")):
        return {str(i): 'NR' for i in range(1, 23)}

    # Euploid / Normal → all N
    if su in ("EUPLOID", "NORMAL", "NORMAL CHROMOSOME COMPLEMENT"):
        return statuses

    # Split into comma/semicolon tokens
    tokens = re.split(r'[,;]|\band\b', s, flags=re.IGNORECASE)

    for token in tokens:
        tok = token.strip()
        is_mos = bool(re.search(r'\bmosaic\b', tok, re.IGNORECASE))
        is_seg = bool(re.search(r'\bsegmental\b', tok, re.IGNORECASE))

        # +/- N  (e.g., "+3", "-18", "+X")
        for m in re.finditer(r'([+-])\s*(1[0-9]|2[0-2]|[1-9])\b', tok):
            sign, num = m.group(1), m.group(2)
            if is_mos and is_seg:
                statuses[num] = 'SMG' if sign == '+' else 'SML'
            elif is_mos:
                statuses[num] = 'MG' if sign == '+' else 'ML'
            elif is_seg:
                statuses[num] = 'SG' if sign == '+' else 'SL'
            else:
                statuses[num] = 'G' if sign == '+' else 'L'

        # del(Nq/p) / dup(Np/q)
        for m in re.finditer(r'\b(del|dup)\s*\(\s*(1[0-9]|2[0-2]|[1-9])\s*[pq]?', tok, re.IGNORECASE):
            op, num = m.group(1).lower(), m.group(2)
            statuses[num] = ('SMG' if op == 'dup' else 'SML') if is_mos else ('SG' if op == 'dup' else 'SL')

        # Monosomy / Trisomy
        for m in re.finditer(r'\b(monosomy|nullisomy)\s+(1[0-9]|2[0-2]|[1-9])\b', tok, re.IGNORECASE):
            statuses[m.group(2)] = 'ML' if is_mos else 'L'
        for m in re.finditer(r'\b(trisomy|tetrasomy)\s+(1[0-9]|2[0-2]|[1-9])\b', tok, re.IGNORECASE):
            statuses[m.group(2)] = 'MG' if is_mos else 'G'

        # "Loss ChrN" / "Gain ChrN" / "Segmental Loss ChrN"
        for m in re.finditer(r'\bloss\s+(?:chr)?\s*(1[0-9]|2[0-2]|[1-9])\b', tok, re.IGNORECASE):
            num = m.group(1)
            statuses[num] = ('SML' if is_seg else 'ML') if is_mos else ('SL' if is_seg else 'L')
        for m in re.finditer(r'\bgain\s+(?:chr)?\s*(1[0-9]|2[0-2]|[1-9])\b', tok, re.IGNORECASE):
            num = m.group(1)
            statuses[num] = ('SMG' if is_seg else 'MG') if is_mos else ('SG' if is_seg else 'G')

    return statuses


def validate_statuses(statuses, raw_result):
    """
    Ensure every chromosome explicitly mentioned in raw_result has a non-N status.
    Fills gaps with ANEUPLOID-safe defaults if needed.
    """
    if not raw_result:
        return statuses
    # Re-derive and fill any missing non-N entries
    derived = derive_chromosome_statuses(raw_result)
    for num, st in derived.items():
        if st != 'N' and statuses.get(num, 'N') == 'N':
            statuses[num] = st
    return statuses


# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY FIELD DERIVATION
# ─────────────────────────────────────────────────────────────────────────────

def derive_autosomes(raw_result, chromosome_statuses, existing_autosomes=""):
    """
    Return the Autosomes display string.
    - EUPLOID → "Normal"
    - FAILED  → "No result"
    - Otherwise: list affected chromosomes, e.g. "-12", "+3, -18", "del(7q)", "Mosaic -15"
    - If existing_autosomes is non-empty and is already a clean label, prefer it.
    """
    cls = classify_embryo(raw_result)["classification"]

    if cls == EUPLOID:
        return "Normal"
    if cls == FAILED:
        return "No result"

    # If existing value looks clean (not a raw code), use it
    existing = (existing_autosomes or "").strip()
    raw_codes_pattern = re.compile(
        r'\b(euploid|aneuploid|mosaic|normal|multiple chromosomal|low level|high level|complex|no result)\b',
        re.IGNORECASE
    )
    if existing and not raw_codes_pattern.search(existing):
        # Sanitise: remove any accidental XX/XY mentions
        existing = re.sub(r'\b(XX|XY)\b', '', existing, flags=re.IGNORECASE).strip(', ')
        if existing:
            return existing

    # Auto-derive from chromosome_statuses
    parts = []
    for i in range(1, 23):
        st = chromosome_statuses.get(str(i), 'N')
        if st == 'N':
            continue
        pct = _find_pct(raw_result, i)
        arm = ''
        if st == 'L':
            parts.append(f"-{i}")
        elif st == 'G':
            parts.append(f"+{i}")
        elif st == 'SL':
            arm = _find_arm(raw_result, i, 'del')
            parts.append(f"del({i}{arm})" if arm else f"del({i})")
        elif st == 'SG':
            arm = _find_arm(raw_result, i, 'dup')
            parts.append(f"dup({i}{arm})" if arm else f"dup({i})")
        elif st == 'ML':
            parts.append(f"Mosaic -{i}" + (f"(~{pct}%)" if pct else ""))
        elif st == 'MG':
            parts.append(f"Mosaic +{i}" + (f"(~{pct}%)" if pct else ""))
        elif st == 'SML':
            arm = _find_arm(raw_result, i, 'del')
            parts.append(f"Mosaic del({i}{arm})" + (f"(~{pct}%)" if pct else ""))
        elif st == 'SMG':
            arm = _find_arm(raw_result, i, 'dup')
            parts.append(f"Mosaic dup({i}{arm})" + (f"(~{pct}%)" if pct else ""))
        elif st == 'NR':
            parts.append(f"Chr{i}: No result")

    return ", ".join(parts) if parts else "Normal"


def sanitize_sex_chromosomes(sex_text, raw_result="", classification=None):
    """
    Return safe sex-chromosome display string.
    NEVER reveals XX or XY per PNDT Act 1994.
    """
    s = str(sex_text or "").strip()
    su = s.upper()

    # Remove any accidental XX/XY literal
    s_clean = re.sub(r'\b(XX|XY)\b', '', s, flags=re.IGNORECASE).strip(', ')

    # If result is already clean and specific, return it
    if s_clean and s_clean.upper() not in ('NORMAL', ''):
        return s_clean

    # Derive from raw result or classification
    cls = (classification or classify_embryo(raw_result or "")["classification"])

    if cls in (EUPLOID, FAILED):
        return s_clean if s_clean else ("Normal" if cls == EUPLOID else "No result")

    r = str(raw_result or "")
    has_mosaic = bool(re.search(r'\bmosaic\b', r, re.IGNORECASE))

    if re.search(r'[-]\s*X\b|\bmonosomy\s+x\b', r, re.IGNORECASE):
        return "Mosaic -X" if has_mosaic else "-X"
    if re.search(r'[+]\s*X\b|\btrisomy\s+x\b', r, re.IGNORECASE):
        return "Mosaic +X" if has_mosaic else "+X"
    if re.search(r'[-]\s*Y\b|\bmonosomy\s+y\b', r, re.IGNORECASE):
        return "Mosaic -Y" if has_mosaic else "-Y"
    if re.search(r'[+]\s*Y\b|\btrisomy\s+y\b', r, re.IGNORECASE):
        return "Mosaic +Y" if has_mosaic else "+Y"

    return s_clean if s_clean else "Normal"


def _find_arm(raw, chr_num, op):
    """Extract the chromosomal arm (p/q) for a segmental change."""
    m = re.search(rf'\b{op}\s*\(\s*{chr_num}\s*([pq])', raw, re.IGNORECASE)
    return m.group(1).lower() if m else ""


def _find_pct(raw, chr_num):
    """Extract percentage associated with a specific chromosome number."""
    m = re.search(
        rf'(?:[+-]\s*{chr_num}|monosomy\s+{chr_num}|trisomy\s+{chr_num}|chr\s*{chr_num})'
        rf'\s*\(?\s*~?\s*(\d+(?:\.\d+)?)\s*%',
        raw, re.IGNORECASE
    )
    if m:
        return int(float(m.group(1)))
    # Fallback: any percentage in the full string
    m = re.search(r'~?\s*(\d+(?:\.\d+)?)\s*%', raw)
    return int(float(m.group(1))) if m else None


def any_mosaic(embryos_data):
    """Return True if any embryo in the list is classified as mosaic."""
    for emb in (embryos_data or []):
        raw = (emb.get('result_summary') or emb.get('result_description') or '')
        if classify_embryo(raw)["is_mosaic"]:
            return True
    return False
