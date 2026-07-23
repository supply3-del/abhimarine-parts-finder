# AbhiMarine Parts Finder — Build Spec (for vibe-coding / AI coding tools)

Paste this whole file into Cursor, Claude Code, Windsurf, v0, or any AI coding
assistant to (re)generate or extend the app. It is written to be self-contained.

---

## 1. What we're building

An internal web app for a marine & industrial spare-parts trading company
(AbhiMarine, Mumbai). The sales/ops team uses it to answer customer enquiries:
"do we have spares for a Wärtsilä L20 plunger?" → search → see matching parts,
which models they're interchangeable with, how complete each listing is, and
export a model's full spares list as Excel or PDF.

**Stack:** Python · SQLite · Streamlit. Single machine / small LAN deployment.
**Look & feel:** modern, clean, **teal palette** (primary `#0F9D8F`, dark
`#0B6B61`, ink text `#0B2E2A`, mint surfaces `#F1FAF8`). Rounded cards, soft
shadows, pill-shaped tags, gold stars (`#F2A900`).

---

## 2. Data sources

Two Excel files live in `./data/`:

### `INVENTORY_FILE_30_06_2026.xlsx` — sheet `Ind+UAE_inv`, ~41,225 rows
Relevant columns (headers exactly as in the file, **note the messy spacing**):
`SR, Location, Std. Brand, old Brand, Model Group Code, Model,
PART NUMBER, Category, Sub-Category, PART NAME, Standard_Part_Name, Fittings,
Condition, Details/Dimension, Marking, "key  specifications" (two spaces),
Genuinity, QTY, UNIT, RACK, Weight, Photos Availabiltity (sic), Availability,
DATE of last update, Standardized model`.

Data is dirty: NaNs, case variants (`Yanmar`/`YANMAR`, `Unknown`/`UNKNOWN`),
condition spelled many ways (`NEW`, `New`, `NEW/UNUSED`, `USED GOOD`…),
`QTY` like `"1 PC"`, `"6.0"`, `"10"`, `Availability` like `YES`/`Yes?`/`NO`.

### `Model_Group_code_look-up_sheet.xlsx` — sheet `Sheet1`, ~1,889 rows
Columns: `Standardized model`, `Model Group Code`. Maps each model to its group.

---

## 3. Core domain concept — Model Group interchangeability

Parts that physically fit several engine models share a **Model Group Code**
ending in `+`. Example: group `L20+` covers
`L20, L20(GAS), 4L20, 6L20, 8L20, 9L20, 6L200, IL-200/330, L20/L32/W32, ZL20`.
A spare listed under any of those models can serve a customer asking for any
other model in the same group. The inventory already carries a `Model Group
Code` column (~91% filled); fill the gaps from the lookup sheet via
`Standardized model`. Ignore the literal group `UNKNOWN` for interchangeability.

---

## 4. ETL — `build_db.py`

Read both spreadsheets, clean, score, and write `parts.db` (SQLite). Re-runnable.

**Cleaning rules**
- NaN-safe string clean: `series.fillna("").astype(str).str.strip()`, then treat
  `{"", "nan", "none", "?", "na", "n/a", "-"}` (case-insensitive) as blank.
  (Important: `astype(str)` alone does **not** stringify NaN reliably here — always
  `fillna("")` first.)
- `qty` = first number extracted from the `QTY` cell, else 0.
- `brand` = normalised `Std. Brand` (title-case, keep short all-caps acronyms),
  blank → `"Unknown"`.
- `condition` mapped to canonical labels: New, New / Unused, Unused,
  Reconditioned, Used, Used / Good, etc.
- `available` = 0 if `Availability` starts with "NO"; else 1 if qty>0.
- `photo` = 1 if `Photos Availabiltity` contains "YES".

**Info-richness score → `stars` (1–5)** — reward fields a buyer can act on:
```
points = 2.0*has(part_number) + 2.0*has(details) + 1.5*has(marking)
       + 1.5*photo_yes        + 1.0*has(genuinity) + 1.0*has(key_specs)
       + 0.5*has(sub_category) + 0.5*has(fittings) + 0.5*has(weight)
stars  = 5 if points>=6 ; 4 if >=4.5 ; 3 if >=3 ; 2 if >=1.5 ; else 1
```
(Calibrated against the real file this yields roughly 19/33/33/9/5 % for 1–5★.)

**Tables**
- `parts` — one row per spare with cleaned columns + `row_id` (1..N), `stars`,
  `available`, `qty`, `photo`.
- `model_groups(model, group_code)` — union of the lookup sheet and the
  (model, group) / (std_model, group) pairs observed in inventory; drop `UNKNOWN`.
- Indexes on `parts(group_code, model, brand, stars, available, category)` and
  on `model_groups(group_code, model)`.
- **FTS5** virtual table `parts_fts` over
  `part_name, brand, model, group_code, part_number, marking, details,
  category, sub_category` (external-content on `parts`, `content_rowid=row_id`)
  for fast prefix search.

---

## 5. App — `app.py` (Streamlit)

### Layout
- **Hero header** (teal gradient): title "⚓ AbhiMarine Parts Finder" + one-line
  description.
- **4 KPI tiles**: Total Parts · In Stock · Model Groups · Brands.
- **Sidebar filters**: Brand (multiselect), Condition, Category, Location,
  Minimum info ★ (slider 1–5), "In-stock only" checkbox (default on). Caption
  explaining what stars mean.
- **Search bar** (full width): single text box. Splits the query into terms and
  runs an FTS5 prefix `MATCH` (`"term"*`) across the indexed fields, then applies
  the sidebar filters in SQL. Order results by `stars DESC, available DESC,
  part_name`. Cap at 1,500.
- **Results**: render the first ~60 as **cards** — part name (+📷 if photo),
  meta line (brand · model · qty · location/rack · part no.), pill tags for
  group code / condition / availability, and a gold ★ rating. Show truncated
  `details`. Below, an expander with the full result set as a `st.dataframe`.

### Model-group summary bar (the key feature)
- Above the cards, a selectbox lists the top model groups present in the current
  results ("`L20+  (12 here)`"). Selecting one sets `st.session_state["group"]`.
- A **sticky bottom bar** (`position:fixed; bottom:0`, dark ink background,
  injected via `st.markdown(..., unsafe_allow_html=True)`) shows:
  `Model group L20+ — N spares in stock · total qty Q · interchangeable across
  M models`.
- An **expander** ("📦 L20+ — full spares list & export") reveals:
  - the interchangeable models as pills,
  - **⬇ Excel** and **⬇ PDF** download buttons,
  - the full group spares table (`st.dataframe`).

### Exports
- **Excel** (openpyxl): teal header row (`#0F9D8F`, white bold, centered),
  sensible column widths, frozen header, "In Stock" Yes/No. Filename
  `AbhiMarine_<GROUP>_spares.xlsx`.
- **PDF** (reportlab, landscape A4): teal title + interchangeable-model line +
  timestamp, then a striped table (part name, PN, brand, model, condition, qty,
  location, rack, ★). Wrap long cells in `Paragraph`. Filename
  `AbhiMarine_<GROUP>_spares.pdf`.

### Theme
`.streamlit/config.toml` sets `primaryColor="#0F9D8F"`,
`secondaryBackgroundColor="#F1FAF8"`, `textColor="#0B2E2A"`. Additional CSS via
`st.markdown` for cards, pills, stars, KPI tiles, the fixed footer bar, and
button styling. Add `padding-bottom:7rem` to `.block-container` so the fixed bar
never covers content.

### Performance
- `@st.cache_resource` for the SQLite connection (`check_same_thread=False`).
- `@st.cache_data` for filter option lists and per-group model lists.

---

## 6. Project layout
```
abhimarine_parts_finder/
├── app.py
├── build_db.py
├── parts.db                # generated
├── requirements.txt        # streamlit, pandas, openpyxl, reportlab
├── README.md
├── VIBE_CODE_SPEC.md       # this file
├── .streamlit/config.toml
└── data/
    ├── INVENTORY_FILE_30_06_2026.xlsx
    └── Model_Group_code_look-up_sheet.xlsx
```

---

## 7. Extensions

**Shipped in v2:**
- **Team login** — shared password from `st.secrets["TEAM_PASSWORD"]` (fallback
  constant), session-state gate, sign-out button.
- **Enquiry → quote shortlist** — ➕ on each result adds to `session_state
  ["shortlist"]` (row_id → fields + editable `qty_needed`); a second tab lists
  items, edits qty, removes lines, and exports a quote **Excel** + **PDF** and a
  **WhatsApp** (`wa.me/?text=`) / **Email** (`mailto:`) share message.
- **Fuzzy fallback search** — when FTS returns nothing, `rapidfuzz.process.extract`
  (WRatio, cutoff ~62) over `part_name+model+brand+part_number` returns closest
  matches; degrades to a SQL `LIKE` fallback if rapidfuzz is absent.

**Still open (nice-to-have):**
- Photo thumbnails — needs actual image files/paths added to the inventory; today
  only a yes/no `photo` flag exists.
- Multi-group export (combine several model groups into one workbook).
- Editable inventory (write-back) with an audit trail.
- Per-user accounts / roles instead of one shared password.
```
```
