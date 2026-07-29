"""
Matches a parsed customer-inquiry row (part_number, description) against the
parts inventory: exact part-number match first, falling back to fuzzy
description matching when there's no part number or no exact hit.
"""
import re

import pandas as pd

PARTS_COLS = ("row_id, sr, part_name, std_part_name, brand, model, group_code, condition, qty, unit, "
             "location, rack, part_number, photo, stars, available")


def build_corpus(con):
    """Row ids + searchable text for every part, for fuzzy description matching."""
    rows = con.execute(f"SELECT row_id, part_name||' '||model||' '||brand FROM parts").fetchall()
    return [r[0] for r in rows], [r[1] for r in rows]


def _by_ids(con, ids, score_by_id, match_type):
    if not ids:
        return pd.DataFrame()
    df = pd.read_sql_query(
        f"SELECT {PARTS_COLS} FROM parts WHERE row_id IN ({','.join('?' * len(ids))})",
        con, params=ids)
    df["match_score"] = df["row_id"].map(score_by_id)
    df["match_type"] = match_type
    return df.sort_values("match_score", ascending=False)


def _brand_ok(row_brand, brand):
    return _exact(row_brand, brand) or (len(_norm(brand)) >= 3 and _bidir_substr(row_brand, brand))


def _model_ok(row, model, cust_group):
    if cust_group and row["group_code"] == cust_group:
        return True
    return len(_norm(model)) >= 2 and _bidir_substr(row["model"], model)


def _rank(df, part_number, brand, model, description, cust_group, limit):
    """Add confidence, sort strongest-first, keep the top `limit`. Rows that earn
    no tier carry NaN confidence and sort last (na_position defaults to 'last'),
    so they stay listed as near-misses without claiming a score."""
    if df.empty:
        return df
    df = df.copy()
    if "match_score" not in df.columns:
        df["match_score"] = 0
    df["confidence"] = df.apply(
        lambda r: confidence_level(r, part_number, brand, model, description, cust_group), axis=1)
    return df.sort_values(["confidence", "match_score"], ascending=False).head(limit)


def _scoped_candidates(con, brand, model, cust_group, description):
    """All inventory rows for the extracted brand / model group, scored by how
    well their part name matches the customer's description. This is the hard
    brand/model filter — we search *within* the brand/model, not the whole DB."""
    conds, params = [], []
    if _norm(brand):
        conds.append("UPPER(brand) = ?"); params.append(_norm(brand))
    if cust_group:
        conds.append("group_code = ?"); params.append(cust_group)
    if _norm(model):
        conds.append("UPPER(model) LIKE ?"); params.append(f"%{_norm(model)}%")
    if not conds:
        return pd.DataFrame()
    df = pd.read_sql_query(
        f"SELECT {PARTS_COLS} FROM parts WHERE ({' OR '.join(conds)}) LIMIT 3000",
        con, params=params)
    if df.empty:
        return df
    from rapidfuzz import fuzz
    desc = (description or "").strip().lower()
    df["match_score"] = (df["part_name"].map(lambda n: fuzz.WRatio(desc, (n or "").lower()))
                         if desc else 0)
    df["match_type"] = "brand/model"
    return df


def match_row(con, corpus, part_number: str, description: str,
              brand: str = "", model: str = "", cust_group: str = "", limit: int = 10) -> pd.DataFrame:
    part_number = (part_number or "").strip()
    description = (description or "").strip()

    if part_number:
        # One query for exact + substring + stripped-match siblings: an exact hit
        # no longer hides the interchangeable ones ('90904-0029-169' and
        # '90904-29-0169' are the same gasket), it just outranks them.
        pats = [f"%{part_number}%"]
        segs = _pn_segments(part_number)
        if segs:
            # LIKE on the raw text can't see zero-padding differences, so also
            # pull same-first-and-last-segment rows and let scoring judge them.
            pats.append(f"%{segs[0]}%{segs[-1]}")
        hits = pd.read_sql_query(
            f"SELECT {PARTS_COLS} FROM parts WHERE "
            f"({' OR '.join('part_number LIKE ? COLLATE NOCASE' for _ in pats)}) "
            # exact first, so the LIMIT can never truncate it away
            "ORDER BY (part_number = ? COLLATE NOCASE) DESC LIMIT ?",
            con, params=(*pats, part_number, max(limit, 80)))
        if not hits.empty:
            is_exact = hits["part_number"].map(lambda p: _exact(p, part_number))
            hits["match_score"] = is_exact.map({True: 100, False: 90})
            hits["match_type"] = is_exact.map({True: "exact part no.", False: "partial part no."})
            return _rank(hits, part_number, brand, model, description, cust_group, limit)

    # no part-number hit: if we have a brand/model, search hard-scoped to it
    scoped = _scoped_candidates(con, brand, model, cust_group, description)
    if not scoped.empty:
        return _rank(scoped, part_number, brand, model, description, cust_group, limit)

    # no brand/model (or brand/model not stocked): fall back to global fuzzy description
    if not description:
        return pd.DataFrame()
    from rapidfuzz import process, fuzz
    ids, choices = corpus
    hits = process.extract(description, choices, scorer=fuzz.WRatio,
                           limit=max(limit, 40), score_cutoff=55)
    if not hits:
        return pd.DataFrame()
    score_by_id = {ids[i]: score for _, score, i in hits}
    df = _by_ids(con, list(score_by_id), score_by_id, "fuzzy description")
    return _rank(df, part_number, brand, model, description, cust_group, limit)


# ----------------------------------------------------------------- brand/model extraction
def known_brands_models(con):
    """Distinct brand/model names in inventory, longest-first so e.g. 'MAN B&W'
    is tried before a shorter overlapping name. Letterless values (bare numbers
    / decimals like '2.6', '11') are dropped — they're noise in the model column
    and would substring-hit dimensions in part descriptions."""
    def has_letter(s):
        return any(ch.isalpha() for ch in s)
    brands = [r[0] for r in con.execute(
        "SELECT DISTINCT brand FROM parts WHERE brand NOT IN ('', 'Unknown')").fetchall()
        if has_letter(r[0])]
    models = [r[0] for r in con.execute(
        "SELECT DISTINCT model FROM parts WHERE model != ''").fetchall()
        if has_letter(r[0])]
    return sorted(brands, key=len, reverse=True), sorted(models, key=len, reverse=True)


_TRAILING_BRACKET_RE = re.compile(r"\(([^()]+)\)\s*$")


def _best_inventory_match(text: str, candidates, fuzzy: bool = False, cutoff: int = 82) -> str:
    """Return the inventory brand/model that best matches `text`. An exact match
    (any length) wins first; otherwise the inventory value must appear in the
    customer text and be specific enough (>=3 chars) — so junk 1-2 char
    inventory entries like 'X'/'L' don't substring-hit everything, and a common
    word like 'BELT' doesn't map to a brand named 'Link Belt'. Candidates are
    longest-first, so the most specific match is returned. `fuzzy=True` adds a
    typo-tolerant fallback ('WIECHAI' -> 'WEICHAI') — only pass it for a short
    Brand/Model cell value, never for free-text description (too many false hits)."""
    text_u = (text or "").strip().upper()
    if not text_u:
        return ""
    for c in candidates:
        if c.upper() == text_u:
            return c
    for c in candidates:
        cu = c.upper()
        if len(cu) >= 3 and cu in text_u:
            return c
    if not fuzzy:
        return ""
    specific = [c for c in candidates if len(c) >= 3]
    from rapidfuzz import process, fuzz
    hit = process.extractOne(text_u, [c.upper() for c in specific],
                             scorer=fuzz.WRatio, score_cutoff=cutoff)
    return specific[hit[2]] if hit else ""


def extract_brand_model(description: str, brands, models, brand_hint="", model_hint=""):
    """Map whatever brand/model the customer mentioned to the inventory's own
    brand/model names. Looks in the customer's Brand/Model cell first (fuzzy,
    tolerates typos), then a trailing '(...)' on the description line, then the
    whole description (substring only)."""
    desc = description or ""
    bracket = _TRAILING_BRACKET_RE.search(desc)
    bracket_txt = bracket.group(1) if bracket else ""
    # When the customer gave a Brand/Maker or Model cell, use only that (normalize
    # to inventory, else the caller keeps the raw text). Scanning the description
    # is a fallback for when no such cell exists — doing it anyway lets a part
    # word like 'COMPRESSOR' in the description wrongly override the real maker.
    if (brand_hint or "").strip():
        brand = _best_inventory_match(brand_hint, brands, fuzzy=True)
    else:
        brand = _best_inventory_match(bracket_txt, brands) or _best_inventory_match(desc, brands)
    # model: no fuzzy — inventory models are messy compounds ('... C-12 C12 C18'),
    # so a short customer model like 'C18' stays as its own value (caller keeps
    # the raw hint when this returns "") rather than snap onto a compound string.
    if (model_hint or "").strip():
        model = _best_inventory_match(model_hint, models)
    else:
        model = _best_inventory_match(bracket_txt, models) or _best_inventory_match(desc, models)
    return brand, model


def model_group_code(con, model: str) -> str:
    if not model:
        return ""
    row = con.execute(
        "SELECT group_code FROM model_groups WHERE model = ? COLLATE NOCASE LIMIT 1",
        (model,)).fetchone()
    return row[0] if row else ""


# ----------------------------------------------------------------- confidence level
# MAN B&W part numbers are usually "<drawing#>-<item#>*<dimension>"
# (e.g. "773627-7*95"): the drawing number + dimension identify the physical
# part, the middle item number doesn't — so two part numbers with matching
# ends but a different middle are still the same interchangeable part.
_MIDBIT_RE = re.compile(r"(\d{4,})\s*-\s*[\d.]+\s*\*\s*(\w+)$")


def _middle_bit_match(pn_a: str, pn_b: str) -> bool:
    ma, mb = _MIDBIT_RE.search((pn_a or "").strip()), _MIDBIT_RE.search((pn_b or "").strip())
    return bool(ma and mb and ma.groups() == mb.groups())


# The same MAN B&W part number gets typed with the zero-padding in different
# places: '90904-0029-169' and '90904-29-0169' are one physical gasket, while
# '90904-0046-169' (different middle) is a different part. So: split on the
# separators, drop leading zeros per segment, require every segment to agree.
# Dots stay *inside* a segment: MAK writes '7.1130-16' and '8.1130-016', where
# the leading 7./8./1. is a real part-family prefix, not decoration. Matching on
# the trailing digits alone would call those two the same part — they are not.
_DASHED_PN_RE = re.compile(r"([\d.]+(?:[-/][\d.]+)+)$")


def _pn_segments(pn: str):
    m = _DASHED_PN_RE.search(_norm(pn))
    if not m:
        return ()
    return tuple(s.lstrip("0") or "0" if s.isdigit() else s
                for s in re.split(r"[-/]", m.group(1)))


def _zero_pad_match(pn_a: str, pn_b: str) -> bool:
    sa, sb = _pn_segments(pn_a), _pn_segments(pn_b)
    return bool(sa) and sa == sb


def _stripped_match(pn_a: str, pn_b: str) -> bool:
    """Either MAN B&W stripping rule: '*'-form middle-bit, or dashed zero-padding."""
    return _middle_bit_match(pn_a, pn_b) or _zero_pad_match(pn_a, pn_b)


def _norm(s) -> str:
    return (s or "").strip().upper()


def _exact(a, b) -> bool:
    a, b = _norm(a), _norm(b)
    return bool(a) and bool(b) and a == b


def _substr(a, b) -> bool:
    """One-directional: does inventory value `a` contain the customer's target `b`?"""
    a, b = _norm(a), _norm(b)
    return bool(a) and bool(b) and b in a


def _bidir_substr(a, b) -> bool:
    a, b = _norm(a), _norm(b)
    return bool(a) and bool(b) and (b in a or a in b)


def _pn_exact_or_stripped(cand_pn, cust_pn) -> bool:
    return _exact(cand_pn, cust_pn) or _stripped_match(cand_pn, cust_pn)


def _pn_substr_or_stripped(cand_pn, cust_pn) -> bool:
    return _substr(cand_pn, cust_pn) or _stripped_match(cand_pn, cust_pn)


def _part_name_ok(cand, cust_description) -> bool:
    """Inventory column K (PART NAME) *or* column L (Standard_Part_Name, the
    synonym column) contains-or-equals the customer's part name."""
    return (_bidir_substr(cand["part_name"], cust_description)
            or _bidir_substr(cand.get("std_part_name", ""), cust_description))


def _model_or_group_ok(cand, cust_model, cust_group) -> bool:
    """Either query value (target model name / target model group code) against
    either inventory value (column F Model Group Code / column G Model)."""
    return any(_bidir_substr(inv, q)
               for inv in (cand["group_code"], cand["model"])
               for q in (cust_model, cust_group))


def _group_ok(cand, cust_group) -> bool:
    """Inventory model group code (col F) matches the customer's model group."""
    return _bidir_substr(cand["group_code"], cust_group)


# Tiers as specified in confidance_level/Copy of CL sorter -1.xlsx
# (sheet inv-column-CLmap) — see confidance_level/CONFIDENCE_LEVEL_RULES.md
# for the human-readable write-up. Checked strongest-first.
# ponytail: CL4/CL4.1 merged per request — CL4.1's condition (part number match
# alone) is a strict subset of CL4's, so ORing them collapses to the broader check.
# CL2, CL1 and CL0 were removed per request: a row that earns none of the tiers
# below scores NaN, which sorts to the bottom and renders as a blank cell — still
# visible for eyeballing near-misses, but never presented as a confidence score.
def confidence_level(cand, cust_part_number: str, cust_brand: str, cust_model: str,
                     cust_description: str, cust_group: str) -> float:
    pn_c, pn_q = cand["part_number"], cust_part_number
    pn_strong = _pn_exact_or_stripped(pn_c, pn_q)   # exact or MAN B&W-stripped
    pn_any = _pn_substr_or_stripped(pn_c, pn_q)      # + substring
    grp_match = _group_ok(cand, cust_group)          # inventory col F == customer group
    grp_given = bool(_norm(cust_group))

    if _exact(cand["brand"], cust_brand) and _exact(cand["model"], cust_model) and _exact(pn_c, pn_q):
        return 5.0
    # CL4: a part-number match confirmed by the right engine group — or a strong
    # (exact/stripped) part number when the customer named no group to check against.
    if pn_any and grp_match:
        return 4.0
    if pn_strong and not grp_given:
        return 4.0
    # CL3 (per request): an exact/stripped part number that lands in a DIFFERENT
    # engine group than the customer's — right part number, wrong engine, so 3★.
    if pn_strong and grp_given and not grp_match:
        return 3.0
    model_grp_bidir = _model_or_group_ok(cand, cust_model, cust_group)
    pname_bidir = _part_name_ok(cand, cust_description)
    # CL3 (classic): model/group AND part number AND part name must all line up.
    if model_grp_bidir and pn_any and pname_bidir:
        return 3.0
    # CL2.5: only when the customer gave no part number at all.
    if not pn_q and model_grp_bidir and pname_bidir:
        return 2.5
    return float("nan")


if __name__ == "__main__":
    from math import isnan

    assert _middle_bit_match("773627-7*95", "MAN B&W 773627-7*95")
    assert not _middle_bit_match("773627-7*95", "773627-7*96")
    assert not _middle_bit_match("abc", "773627-7*95")

    # dashed zero-padding: same segments once leading zeros are dropped
    assert _zero_pad_match("90904-0029-169", "90904-29-0169")
    assert _zero_pad_match("90904-29/157", "90904-0029-157")     # '/' is the same separator
    assert not _zero_pad_match("90904-0029-169", "90904-0046-169")   # different middle = different part
    assert not _zero_pad_match("90904-29-169", "90904-29-062")
    assert not _zero_pad_match("1W-1699", "")
    # MAK part-family prefix is significant: 7.1130-16 is not 8.1130-016
    assert _zero_pad_match("7.1130-16", "7.1130-016")
    assert not _zero_pad_match("7.1130-16", "8.1130-016")
    assert not _zero_pad_match("8.2104-48", "1.2104-048")
    assert not _zero_pad_match("7.1630-2", "1630-002")

    cand = {"part_number": "773627-7*95", "part_name": "Cylinder Liner",
           "std_part_name": "", "brand": "MAN B&W", "model": "L20", "group_code": "L20+"}

    # CL5: brand + model + part number all exact
    assert confidence_level(cand, "773627-7*95", "MAN B&W", "L20", "Cylinder Liner", "L20+") == 5.0
    # CL4 (merged 4/4.1): part number exact-or-stripped-match, regardless of brand/model
    assert confidence_level(cand, "773627-14*95", "", "L20", "", "L20+") == 4.0
    assert confidence_level(cand, "773627-14*95", "", "", "", "") == 4.0

    # ---- CL4 needs the right engine group; a strong PN with no group also ok ----
    # group matches (L20+) + PN substring -> 4★ even though the part name is wrong
    assert confidence_level(cand, "627-7*95", "", "", "Snorkel", "L20+") == 4.0
    assert confidence_level(cand, "627-7*95", "", "", "Cylinder", "L20+") == 4.0
    # exact PN, RIGHT group -> 4★
    assert confidence_level(cand, "773627-7*95", "", "", "", "L20+") == 4.0
    # exact / stripped PN but the WRONG engine group -> 3★, not 4★ (per request)
    off = {**cand, "model": "9L70", "group_code": "L70+"}
    assert confidence_level(off, "773627-7*95", "", "", "", "L20+") == 3.0
    assert confidence_level(off, "773627-14*95", "", "", "", "L20+") == 3.0
    # substring PN in the wrong group with no part-name match -> no tier
    assert isnan(confidence_level({**cand, "model": "ZZ9", "group_code": "X99+"},
                                 "627-7*95", "", "", "Snorkel", "L20+"))

    # ---- CL3: model/group AND part number AND part name, no group-code match ----
    assert confidence_level(cand, "627-7*95", "", "L20", "Cylinder", "") == 3.0
    # part number alone is no longer enough — wrong model drops it out of CL3
    assert isnan(confidence_level({**cand, "model": "9L70", "group_code": "L70+"},
                                 "627-7*95", "", "", "", ""))
    # right model + part number but wrong part name and no group code -> no tier
    assert isnan(confidence_level(cand, "627-7*95", "", "L20", "Snorkel", ""))
    # column L (Standard_Part_Name) satisfies the part-name leg on its own
    gasket = {"part_number": "90904-29-0169", "part_name": "No Name",
             "std_part_name": "Gasket", "brand": "MAN B&W",
             "model": "S50MC", "group_code": "50MC+"}
    assert confidence_level(gasket, "90904-0029-0169", "", "S50MC", "Gasket", "50MC+") == 4.0
    # group 50MC+ matches + PN substring -> 4★ via the extra path (was 3★)
    assert confidence_level(gasket, "904-29-0169", "", "S50MC", "Gasket", "50MC+") == 4.0
    # PN substring but no group given -> stays CL3 (needs part name, which holds)
    assert confidence_level(gasket, "904-29-0169", "", "S50MC", "Gasket", "") == 3.0

    # CL2.5: no customer part number, model/group + part name bidirectional substring
    assert confidence_level(cand, "", "", "L20", "Liner", "") == 2.5
    # ---- removed tiers: CL2 and CL0 now score NaN, not a number ----
    # was CL2 (model variant): part number supplied but didn't match this candidate
    assert isnan(confidence_level(cand, "999999", "", "L20", "Liner", ""))
    # was CL2 (brand variant): brand + part name matched but nothing else
    assert isnan(confidence_level({**cand, "model": "9L70", "group_code": "L70+"},
                                 "999999", "MAN B&W", "", "Liner", ""))
    # was CL0: only same model group, nothing else lines up
    assert isnan(confidence_level({**cand, "part_name": "Gasket"}, "999999", "", "L20", "Snorkel", ""))
    # CL2.5 needs the part name too — model/group alone is no longer a tier
    assert isnan(confidence_level({**cand, "part_name": "Gasket"}, "", "", "L20", "Snorkel", ""))

    brands, models = ["MAN B&W", "WEICHAI", "Yanmar"], ["6L20", "X6170ZC", "L20"]
    # brand/model mentioned inside the description text
    assert extract_brand_model("Need MAN B&W L20 plunger", brands, models) == ("MAN B&W", "L20")
    assert extract_brand_model("no brand mentioned here", brands, models) == ("", "")
    # trailing "(...)" on the line is checked before the rest of the description
    assert extract_brand_model("Cylinder Liner (MAN B&W L20)", brands, models) == ("MAN B&W", "L20")
    # customer's own Brand/Model cell, wordier than inventory -> normalized to inventory value
    assert extract_brand_model("GASKET", brands, models,
                               brand_hint="WEICHAI 170Z SERIES",
                               model_hint="WEICHAI X6170ZC") == ("WEICHAI", "X6170ZC")
    # fuzzy fallback absorbs a typo in the customer's Brand cell (not free text)
    assert extract_brand_model("GASKET", brands, models, brand_hint="WIECHAI")[0] == "WEICHAI"
    # an unrelated brand cell must NOT fuzzy-hit
    assert extract_brand_model("GASKET", brands, models, brand_hint="ACME PUMPS")[0] == ""

    # brand/model hard-filter helpers
    row_cat = {"brand": "Caterpillar", "model": "C18", "group_code": "CAT+"}
    row_man = {"brand": "MAN B&W", "model": "L20", "group_code": "L20+"}
    assert _brand_ok(row_cat["brand"], "Caterpillar")           # exact brand kept
    assert not _brand_ok(row_man["brand"], "Caterpillar")       # off-brand dropped
    assert _model_ok(row_man, "L20", "")                        # model bidirectional-substring kept
    assert _model_ok(row_cat, "", "CAT+")                       # same model group kept
    assert not _model_ok(row_man, "C18", "CAT+")                # wrong model + wrong group dropped
    print("inquiry_match self-check OK")
