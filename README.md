# ⚓ AbhiMarine Parts Finder

A Streamlit + SQLite web app for the AbhiMarine team to answer customer enquiries
fast — search 41,000+ marine engine / equipment spares by model, brand, part name
or part number, see **model-group interchangeability**, **info-richness star
ratings**, and export any model group's full spares list to **Excel or PDF**.

---

## Quick start

```bash
# 1. (optional) create a virtual environment
python -m venv .venv && source .venv/bin/activate    # Windows: .venv\Scripts\activate

# 2. install dependencies (requirements.txt lives at the repo root)
pip install -r requirements.txt

# 3. build the database from the Excel files in abhimarine_parts_finder/data
#    (the inventory spreadsheets are not in this repo — supply your own)
python abhimarine_parts_finder/build_db.py

# 4. run the app
streamlit run abhimarine_parts_finder/app.py
```

The app opens at <http://localhost:8501>.

---

## Deploying on Streamlit Community Cloud

- **Main file path:** `abhimarine_parts_finder/app.py`
- **requirements.txt** and **packages.txt** must sit at the repo root (they do) —
  Cloud installs Python deps from the former and the `tesseract-ocr` system
  package (for scanned-PDF / image OCR) from the latter.
- The repo is **private**, so Streamlit needs GitHub access to it: authorize the
  Streamlit app for private repositories in your GitHub settings.
- `.doc` parsing uses MS Word COM automation, which only works on Windows — on
  the Linux Cloud host, upload `.docx`/PDF/Excel instead.

---

## How it works

| File | Purpose |
|------|---------|
| `data/INVENTORY_FILE_30_06_2026.xlsx` | source inventory (41,225 rows) |
| `data/Model_Group_code_look-up_sheet.xlsx` | model → group-code mapping |
| `build_db.py` | cleans the data, scores each part 1–5 ★, builds `parts.db` (SQLite + FTS5) |
| `parts.db` | the prepared database the app reads (pre-built and included) |
| `app.py` | the Streamlit interface |
| `VIBE_CODE_SPEC.md` | full spec you can paste into Cursor / Claude Code to extend or rebuild |

### Updating inventory
Drop a new inventory or lookup file into `data/` (keep the same column headers),
re-run `python build_db.py`, and refresh the app. Nothing else changes.

### Interchangeability
Parts that fit several models share a **Model Group Code** (e.g. `L20+` covers
`L20, 4L20, 6L20, 8L20, 9L20, 6L200, IL-200/330, L20/L32/W32, ZL20`…). When you
pick a model group in the app, the bottom bar totals every in-stock spare across
all interchangeable models and lets you export the list.

### Info-richness stars (1–5 ★)
A listing earns stars for the detail a buyer can act on — part number, dimensions,
markings, genuinity, a photo, key specs. More complete listings rank higher in
search results.

---
