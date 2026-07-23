"""
build_db.py  -  AbhiMarine Parts Finder
Reads the inventory + model-group lookup Excel files, cleans the data,
computes a 1-5 star "info-richness" score per part, resolves model-group
interchangeability, and writes everything into a single SQLite database
(parts.db) with a full-text-search index.

Run:   python build_db.py
Re-run any time the source spreadsheets change.
"""

import os
import re
import sqlite3
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
DB_PATH = os.path.join(HERE, "parts.db")

INVENTORY_FILE = os.path.join(DATA, "INVENTORY_FILE_30_06_2026.xlsx")
LOOKUP_FILE = os.path.join(DATA, "Model_Group_code_look-up_sheet.xlsx")

# ---------------------------------------------------------------- helpers
EMPTY = {"", "nan", "none", "?", "na", "n/a", "-"}


def clean(series: pd.Series) -> pd.Series:
    """NaN-safe string cleaner -> stripped string, blanks normalised to ''."""
    s = series.fillna("").astype(str).str.strip()
    return s.mask(s.str.lower().isin(EMPTY), "")


def filled(series: pd.Series) -> pd.Series:
    return clean(series).ne("")


def to_qty(series: pd.Series) -> pd.Series:
    """Pull the first number out of messy qty cells ('1 PC', '6.0', '10')."""
    nums = series.fillna("").astype(str).str.extract(r"([\d]+\.?[\d]*)")[0]
    return pd.to_numeric(nums, errors="coerce").fillna(0.0)


CONDITION_MAP = [
    ("NEW/UNUSED", "New / Unused"),
    ("NEW/NOT GOOD", "New / Not Good"),
    ("NEW/GOOD", "New / Good"),
    ("NEW RUSTED", "New (Rusted)"),
    ("UNUSED", "Unused"),
    ("RECON", "Reconditioned"),
    ("USED/GOOD", "Used / Good"),
    ("USED GOOD", "Used / Good"),
    ("USED", "Used"),
    ("NEW", "New"),
]


def norm_condition(val: str) -> str:
    u = val.upper()
    for key, label in CONDITION_MAP:
        if key in u:
            return label
    return val.title() if val else "Unspecified"


def norm_brand(val: str) -> str:
    if not val or val.lower() in EMPTY:
        return "Unknown"
    return " ".join(w if w.isupper() and len(w) <= 4 else w.capitalize()
                    for w in val.split())


def norm_avail(val: str, qty: float) -> int:
    """1 = available, 0 = not. Anything starting NO -> 0, else qty>0 -> 1."""
    u = val.strip().upper()
    if u.startswith("NO"):
        return 0
    if u.startswith("YES") and qty > 0:
        return 1
    return 1 if qty > 0 else 0


# ---------------------------------------------------------------- scoring
def info_stars(df: pd.DataFrame) -> pd.Series:
    """1-5 stars based on how complete / sales-ready each listing is."""
    photo = df["Photos Availabiltity"].astype(str).str.upper().str.contains("YES", na=False)
    pts = (
        filled(df["PART NUMBER"]).astype(float) * 2.0
        + filled(df["Details/Dimension"]).astype(float) * 2.0
        + filled(df["Marking"]).astype(float) * 1.5
        + photo.astype(float) * 1.5
        + filled(df["Genuinity"]).astype(float) * 1.0
        + filled(df["key  specifications"]).astype(float) * 1.0
        + filled(df["Sub-Category"]).astype(float) * 0.5
        + filled(df["Fittings"]).astype(float) * 0.5
        + filled(df["Weight"]).astype(float) * 0.5
    )

    def to_star(p):
        if p >= 6.0:
            return 5
        if p >= 4.5:
            return 4
        if p >= 3.0:
            return 3
        if p >= 1.5:
            return 2
        return 1

    return pts.apply(to_star).astype(int)


# ---------------------------------------------------------------- build
def build():
    print("Reading inventory ...")
    inv = pd.read_excel(INVENTORY_FILE, sheet_name="Ind+UAE_inv")
    print(f"  {len(inv):,} rows")

    lk = pd.read_excel(LOOKUP_FILE)

    out = pd.DataFrame()
    out["sr"] = clean(inv["SR"])
    out["brand"] = clean(inv["Std. Brand"]).map(norm_brand)
    out["old_brand"] = clean(inv["old Brand"])
    out["group_code"] = clean(inv["Model Group Code"])
    out["model"] = clean(inv["Model"])
    out["std_model"] = clean(inv["Standardized model"])
    out["part_number"] = clean(inv["PART NUMBER"])
    out["category"] = clean(inv["Category"]).str.title()
    out["sub_category"] = clean(inv["Sub-Category"]).str.title()
    out["part_name"] = clean(inv["PART NAME"]).str.title()
    out["std_part_name"] = clean(inv["Standard_Part_Name"])
    out["fittings"] = clean(inv["Fittings"])
    out["condition"] = clean(inv["Condition"]).map(norm_condition)
    out["details"] = clean(inv["Details/Dimension"])
    out["marking"] = clean(inv["Marking"])
    out["key_specs"] = clean(inv["key  specifications"])
    out["genuinity"] = clean(inv["Genuinity"])
    out["qty"] = to_qty(inv["QTY"])
    out["unit"] = clean(inv["UNIT"])
    out["rack"] = clean(inv["RACK"])
    out["location"] = clean(inv["Location"])
    out["weight"] = clean(inv["Weight"])
    out["photo"] = inv["Photos Availabiltity"].astype(str).str.upper().str.contains("YES", na=False).astype(int)
    out["stars"] = info_stars(inv)
    out["available"] = [norm_avail(a, q) for a, q in zip(clean(inv["Availability"]), out["qty"])]
    out["last_update"] = clean(inv["DATE of last update"])

    # fill missing group_code from lookup via standardized model
    lk_map = dict(zip(clean(lk["Standardized model"]), clean(lk["Model Group Code"])))
    need = out["group_code"].eq("") & out["std_model"].ne("")
    out.loc[need, "group_code"] = out.loc[need, "std_model"].map(lk_map).fillna("")

    out["row_id"] = range(1, len(out) + 1)

    # ---- model_groups table (union of lookup + inventory observed pairs) ----
    pairs = []
    for m, g in zip(clean(lk["Standardized model"]), clean(lk["Model Group Code"])):
        if m and g:
            pairs.append((m, g))
    for m, g in zip(out["model"], out["group_code"]):
        if m and g:
            pairs.append((m, g))
    for m, g in zip(out["std_model"], out["group_code"]):
        if m and g:
            pairs.append((m, g))
    mg = pd.DataFrame(pairs, columns=["model", "group_code"]).drop_duplicates()
    mg = mg[mg["group_code"] != "UNKNOWN"]

    # ---- write SQLite ----
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    con = sqlite3.connect(DB_PATH)
    out.to_sql("parts", con, index=False)
    mg.to_sql("model_groups", con, index=False)

    cur = con.cursor()
    for col in ["group_code", "model", "brand", "stars", "available", "category"]:
        cur.execute(f"CREATE INDEX idx_parts_{col} ON parts({col})")
    cur.execute("CREATE INDEX idx_mg_group ON model_groups(group_code)")
    cur.execute("CREATE INDEX idx_mg_model ON model_groups(model)")

    # full text search index
    cur.execute("""CREATE VIRTUAL TABLE parts_fts USING fts5(
        part_name, brand, model, group_code, part_number, marking,
        details, category, sub_category, content='parts', content_rowid='row_id')""")
    cur.execute("""INSERT INTO parts_fts(rowid, part_name, brand, model, group_code,
        part_number, marking, details, category, sub_category)
        SELECT row_id, part_name, brand, model, group_code, part_number,
        marking, details, category, sub_category FROM parts""")
    con.commit()

    print("\nStar distribution:")
    print(out["stars"].value_counts().sort_index().to_string())
    print(f"\nGroups: {mg['group_code'].nunique():,}  |  In-stock parts: {int((out['available']==1).sum()):,}")
    print(f"Wrote {DB_PATH}")
    con.close()


if __name__ == "__main__":
    build()
