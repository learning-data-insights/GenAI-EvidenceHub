# IRR (Inter-Rater Reliability) calculation

Computes inter-rater reliability for the GenAI Evidence Hub coding workbook.
The method was **reverse-engineered from the full combined coding sheet and the
meeting notes that recorded the team's human-calculated IRR values**, then
calibrated until the script reproduced those reference numbers.

## The metric

**Krippendorff's alpha, nominal level of measurement**, computed as a single
**pooled** coefficient (not an average of per-question alphas):

| Element | Definition |
|---|---|
| reliability unit | one `(Paper ID, Research Question ID, Field)` cell |
| observer | a coder (the `Reviewer` column; spellings/aliases normalized) |
| coincidence matrix | **all** cells pooled into one matrix → **one** alpha |
| overall alpha | every cell |
| per-paper alpha | cells of that paper only |
| pairwise / "by team" | observers restricted to the two coders |
| missing | blank / `N/A` dropped; a unit needs ≥2 coders to contribute |

## What the script expects from the workbook

So it can be pointed at a new export without code changes, it assumes:

- **Each coded sheet has two header rows:** row 1 = field display names,
  row 2 = data types (`Selection`, `Multiselection`, `Numerical`, `String`,
  `True`/`False`). Data starts on row 3.
- **Required columns:** `Reviewer`, `Paper ID`, and (optionally)
  `Research Question ID`. Rows without a reviewer or paper id are skipped.
- **Which fields count toward IRR** is read from a sheet named
  **`FieldList+IRR`** (column A = field name, column B = `Y` to include). If
  that sheet is absent, a built-in fallback list is used. So the field set is
  data-driven — editing the workbook's `FieldList+IRR` sheet changes what the
  script scores, no code edit required.
- A coded sheet is auto-detected by the presence of a `Reviewer` column
  (`--all-sheets` uses this to skip empty template tabs).

## Two policy choices (the reverse-engineering result)

These are *why* the script matches the human numbers where a naive exact-match
tool does not:

1. **Free-text "list" fields are excluded** — e.g. `Research Question`,
   `Tested LLM Model & Version Used`, `Tested Model & Techniques List`,
   `Baseline list`. Coders write divergent prose ("gpt-4 (zero-shot)" vs
   "zs_nocot, zs_cot"); the team scores these by *concept agreement, not exact
   match*, which can't be automated reliably. **This was the main flaw in the
   original tool** — it exact-string-matched these and counted false
   disagreements, deflating alpha. The excluded names live in the
   `EXCLUDE_FREETEXT` constant at the top of `compute_irr.py`; adjust there if
   the template's free-text fields change.
2. **Metric numbers are normalized** so `85` == `85.0` == `"85"`, percentages
   are parsed, and multiselect cells are compared order-independently.

## Adapting coder names

Any reviewer name **not** in the alias table passes through unchanged, so the
tool already works for an arbitrary set of coders. Aliases matter only when one
person appears under multiple spellings (e.g. initials vs. full name). Two ways
to supply them:

- **`--aliases aliases.json`** — a JSON object mapping spellings to a canonical
  name, merged over (and overriding) the built-ins. No code edit:

  ```json
  { "jw": "John", "j.white": "John", "maggie b": "Maggie" }
  ```

- Or edit the `NAME_MAP` default near the top of `compute_irr.py`.

## Usage (CLI)

```bash
python compute_irr.py WORKBOOK.xlsx --sheet SHEET_NAME              # overall alpha
python compute_irr.py WORKBOOK.xlsx --sheet SHEET_NAME --pairs --per-paper
python compute_irr.py WORKBOOK.xlsx --sheet SHEET_NAME --per-field --csv out.csv
python compute_irr.py WORKBOOK.xlsx --all-sheets                    # every coded sheet
python compute_irr.py WORKBOOK.xlsx --all-sheets --aliases aliases.json
```

`--csv` writes a tidy `scope,key,alpha,n_units` table (scopes: `overall`,
`paper`, `pair`, `field`) for downstream use.

## Usage (as a library)

The two functions a host tool will want:

```python
from compute_irr import load_workbook_records, krippendorff_alpha

# records: list of (paper, rq, field, coder, value); missing values are None.
records, fields = load_workbook_records("WORKBOOK.xlsx", "SHEET_NAME")

overall, n_units, n_pairs = krippendorff_alpha(records)            # all cells
pair_alpha, *_           = krippendorff_alpha(records, coders={"John", "Maggie"})
paper_alpha, *_          = krippendorff_alpha(records, paper="12")

# custom alias map:
import json
records, _ = load_workbook_records(
    "WORKBOOK.xlsx", "SHEET_NAME",
    name_map=json.load(open("aliases.json")),   # or build the dict inline
)
```

`krippendorff_alpha` returns `(alpha, n_units, n_pairable_values)`; `alpha` is
`None` when fewer than two comparable values exist. Only depends on `openpyxl`.

## Validation notes

When checked against the recorded human-calculated pairwise numbers, the
cleanest coding round reproduced 4 of 5 pairs to within ~0.01. Some earlier
rounds read uniformly low under *every* field configuration because that data
was re-coded/consolidated after its meeting numbers were recorded — a
data-version effect, not a method error. Re-run validation against the current
notes whenever the workbook is re-exported.
