# Confidence Level Rules — parsed from `Copy of CL sorter -1.xlsx`

Source: `confidance_level/Copy of CL sorter -1.xlsx`, sheets `inv-column-CLmap`
(the rule matrix) + `how-to-read` (intent). This is a read-out of what the
sheet says, translated to our field names — nothing has been coded yet.

## Field name mapping (their inventory column → our `parts` column)

| Sheet column | Our field |
|---|---|
| Std. Brand | `brand` |
| Model Group Code | `group_code` |
| Model | `model` |
| PART NUMBER | `part_number` |
| PART NAME / Standard_Part_Name | `part_name` (we match against `description` from the customer line) |
| Condition | `condition` |
| QTY | `qty` |
| Photos Availabiltity | `photo` |
| Availability | `available` |

## Match-type vocabulary used in the sheet

- **Exact** ("Yes") — values equal.
- **Substring** — inventory value contains the customer's target value (one-directional).
- **Bidirectional substring** ("LHS contains ... or ... is a substring of LHS") — either side contains the other.
- **Stripped Match** — two MAN B&W part-number rules, ORed together as `_stripped_match`:
  - *Middle Bit* (`*` form): strip the middle segment, compare start+end (`773627-7*95` vs `773627-14*95` → match). `_middle_bit_match`.
  - *Zero-padding* (all-dash/slash form): split on `-` / `/`, drop leading zeros per segment, require **every** segment to agree (`90904-0029-169` vs `90904-29-0169` → match; vs `90904-0046-169` → **no** match, different part). `_zero_pad_match`.
- **Part name** — inventory column **K** `PART NAME` *or* column **L** `Standard_Part_Name` (the synonym column), bidirectional substring. `_part_name_ok`.
- **Model** — either query value (target model name *or* target model group code) against either inventory value (column **F** `Model Group Code` *or* column **G** `Model`), bidirectional substring. `_model_or_group_ok`.
- **Missing** — this tier only applies when the customer *didn't* supply a part number (CL2.5).

## The tiers, as implemented

CL4 keeps the merge with CL4.1: CL4.1's condition (`part_number` match alone)
is a strict subset of what CL4 additionally required, so ORing them collapses
to the broader part-number-only check.

**CL3 was rewritten** (2026-07-20) to require all three legs — model **and**
part number **and** part name. It is no longer part-number-only.

**CL2 (both variants), CL1 and CL0 were removed** (2026-07-20). Four tiers
remain. A row matching none of them scores `NaN`: it stays in the result list,
sorted to the bottom, with a blank confidence cell — visible as a near-miss but
never presented as a score.

| Tier | Conditions (ALL must hold unless noted) |
|---|---|
| **CL5** | `brand` exact **AND** `model` exact **AND** `part_number` exact |
| **CL4** | `part_number` exact-or-stripped-match |
| **CL3** | model/group **AND** `part_number` substring-or-stripped-match **AND** part name |
| **CL2.5** | *Only when customer gave no part number:* model/group **AND** part name |
| *(none)* | `NaN` — listed last, blank cell |

"model/group" and "part name" above are the `_model_or_group_ok` and
`_part_name_ok` definitions from the vocabulary section.

## Output-field guidance (separate from confidence logic)

The sheet also marks which columns to surface where — this isn't a scoring
rule, it's UI guidance:

- **First view** (compact/card): brand, group_code, model, part_number, condition, dispatch status, qty, photo flag
- **2nd view — long text description**: brand + modifier, group_code, model, category, sub-category, part_name, Standard_Part_Name, fittings, condition, dispatch status, details/dimension, marking, key specs, genuinity
- **2nd view — other**: qty, unit, photo flag, availability, box, rack, weight, details, weight (dup), temp location UAE

## Open questions before I implement this

1. **Scale mismatch** — the app currently shows a 1–5 ★ `confidence_stars()` I built from your earlier verbal spec (5/4/4/3/2/1). This sheet defines **8 tiers** (5, 4, 4.1, 3, 3.1, 2.5, 2×2, 0), two of which (4/4.1 and the two 2-variants) share the same star number but different rules. Options:
   - (a) keep 8 numeric levels as-is (`4.1`, `3.1` etc. shown as text/decimal, not a 5-star icon), or
   - (b) collapse to display stars but keep the finer tie-break order internally for sorting.
2. **`Std. Brand` row (J4) text says "model name"** in a column about Brand — looks like a copy-paste artifact from the Model row above it. I've read it as "brand name substring," but flag if that's wrong.
3. **CL2.5 / CL2-model** are only distinguished by *why* there's no part-number match (missing vs. present-but-wrong) — same match rule, different trigger condition. Confirmed that's intentional and not a sheet error?
4. Want the **output-field view guidance** (first view / 2nd view columns) implemented too, or just the confidence scoring for now?
