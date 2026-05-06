# CLAUDE.md — DAF Processor (dealdesk-tiles)

## What this project is

An internal Streamlit web app for RAK Ceramics' Deal Desk team. It processes Discount Approval Forms (DAFs) by accepting two Excel uploads, calculating per-SKU margins, applying an approval matrix, and producing a downloadable processed Excel file. Deployed on Streamlit Community Cloud from the GitHub repo `cherylpersonal/dealdesk-tiles`.

---

## Folder structure

```
Code files/
├── app.py                  # Streamlit UI — the only entry point
├── core/
│   ├── __init__.py
│   ├── parser.py           # Reads and normalises both Excel files
│   ├── validators.py       # Column presence + value checks
│   ├── calculator.py       # Merge + margin calculations
│   ├── approval.py         # Applies approval matrix, returns ApprovalResult
│   └── exporter.py         # Builds the downloadable Excel output
├── generate_samples.py     # One-off script to create sample test files
├── requirements.txt        # streamlit, pandas, openpyxl, xlsxwriter
├── README.md
└── CLAUDE.md               # This file
```

The flat-file copies of the core modules (approval.py, calculator.py, etc.) in the root are unused — the live imports all go through `core/`.

---

## What each file does

### `app.py`
Single-page Streamlit app. Flow:
1. Two `st.file_uploader` widgets (DAF + Cost Sheet)
2. Calls `parse_daf` / `parse_cost_sheet` → `validate_daf` / `validate_cost_sheet` → `merge_data` → `calculate_margins` → `evaluate_approval`
3. Renders approval badge, four summary metrics, a colour-coded per-SKU table with a total row, a Tracker Row Preview expander, and a download button
4. All processing is guarded with `st.stop()` on any error so the page fails gracefully

### `core/parser.py`
- `_normalise_columns`: strips whitespace, lowercases, replaces spaces with underscores
- `_find_header_row`: scans first 10 rows to find the row that best matches expected column names — handles Excel files with blank leading rows (the real cost sheet has one blank row before the header)
- `parse_daf`: reads DAF, coerces numeric columns, normalises `sku_code` to uppercase, aggregates duplicate SKUs (boxes summed, NEF weighted-averaged by boxes)
- `parse_cost_sheet`: reads cost sheet, coerces numeric columns, normalises `sku_code`, deduplicates keeping first occurrence

### `core/validators.py`
- `validate_daf`: checks required columns `[sku_code, boxes, nef]`, flags zero/negative values, warns on duplicate SKUs (not a hard stop — aggregation handles it)
- `validate_cost_sheet`: checks required columns `[sku_code, area_per_box, buying_price]`, flags zero/negative values. Error message includes actual columns found in the file to help diagnose name mismatches.
- `check_missing_skus`: returns list of DAF SKUs absent from cost sheet

### `core/calculator.py`
- `merge_data`: left-joins DAF onto cost sheet on `sku_code`; drops unmatched rows and warns
- `calculate_margins`: adds `total_area`, `revenue`, `cost`, `margin_value`, `margin_percent` columns; computes deal-level totals dict

### `core/approval.py`
- `ApprovalResult` dataclass: `level` (str), `reasons` (list), `sku_flags` (DataFrame)
- `evaluate_approval`: applies the three-tier approval matrix (see below)

### `core/exporter.py`
- `generate_output_file`: writes a two-sheet Excel file (Calculations + Tracker) using xlsxwriter; applies column widths, number formats, colour-coded total row and approval block
- `build_tracker_row`: extracts meta from the first DAF row, computes 6-hr SLA due date from `submitted_date`, calculates avg discount %, leaves manual fields as empty strings

---

## Approval matrix

| Level | Trigger |
|---|---|
| **Deal Desk** | All SKU margins ≥ 10% AND overall ≥ 20% |
| **PM Head** | All SKU margins ≥ 5% AND overall ≥ 10% (but not Deal Desk) |
| **CEO** | Any SKU margin < 5% OR overall < 10% |

---

## Margin calculations

```
total_area     = boxes × area_per_box
revenue        = nef × total_area
cost           = buying_price × total_area
margin_value   = revenue − cost
margin_percent = margin_value / revenue
```

---

## Design choices

### UI (app.py)
- **Layout**: `st.set_page_config(layout="wide")` — full-width
- **Page title**: "DAF Processor", icon 📋
- **Colour scheme** (approval badges + table highlights):
  - Deal Desk (green): background `#D4EDDA`, text `#155724`
  - PM Head (amber): background `#FFF3CD`, text `#856404`
  - CEO (red): background `#F8D7DA`, text `#721C24`
- **Approval badge**: custom HTML via `st.markdown(..., unsafe_allow_html=True)`, CSS class `.approval-badge` with rounded corners (`border-radius: 6px`), 20px bold font
- **Per-SKU table**: `st.dataframe` with pandas Styler — rows colour-coded by margin percent (red <5%, amber 5–10%, green ≥10%). Separate total row displayed below the styled table as a plain dataframe.
- **Metrics**: four `st.columns` with `st.metric` for Total Revenue, Total Cost, Total Margin, Overall Margin %
- **Currency**: Indian Rupee symbol ₹ used throughout

### Excel output (exporter.py)
- Sheet 1 "Calculations": header row in `#E8F0FE` (light blue), total row in `#FFF3CD` (amber), approval level cell colour-matched to the three-tier scheme
- Sheet 2 "Tracker": single structured row matching the Deal Desk tracking sheet columns
- Approval block written two rows below the data, not in a separate sheet
- Filename: `DAF_Processed_{daf_ref_no}.xlsx`

---

## Input file requirements

### DAF Excel
Required columns: `sku_code`, `boxes`, `nef`
Meta columns (optional but used for tracker): `daf_ref_no`, `lob`, `channel`, `state`, `project_name`, `developer_name`, `dealer_name`, `zonal_coordinator`, `submitted_date`, `list_value`
- Column names are normalised (lowercased, spaces→underscores), so "SKU Code" works as well as "sku_code"
- The parser auto-detects the header row, so blank leading rows are fine
- Duplicate SKU rows are aggregated, not rejected

### Cost Sheet Excel
Required columns: `sku_code`, `area_per_box`, `buying_price`
- Same column normalisation and header auto-detection as DAF
- Real cost sheet has one blank row before the header — handled by `_find_header_row`

---

## Deployment

- **Repo**: https://github.com/cherylpersonal/dealdesk-tiles
- **Branch**: `main`
- **Entry point**: `app.py`
- **Platform**: Streamlit Community Cloud (free tier)
- **Python version**: 3.14 (on Streamlit Cloud); 3.12 locally (Windows Store Python)
- Local Python path: `C:\Users\99202928\AppData\Local\Microsoft\WindowsApps\python.exe`
- Local packages install to user site-packages (not system), so run via `python -m streamlit run app.py`

---

## Known issues fixed

- `core/` package missing on initial push — modules were flat files; fixed by creating `core/__init__.py` and copying modules in
- `KeyError` on `df.duplicated("sku_code")` when column absent — fixed with `"sku_code" in df.columns` guard in both parsers
- Cost sheet columns read as `unnamed:_0` etc. — caused by blank row 1 in real file; fixed by `_find_header_row` auto-detection

---

## Test files

Located at: `C:\Users\99202928\OneDrive - RAK Ceramics PJSC\Cheryl\Others\Deal Desk\Test Files\`
- `daf_input.xlsx` — real DAF file; headers on row 1, first column is `daf_ref_nometa` (not `daf_ref_no`)
- `cost_sheet.xlsx` — real cost sheet; one blank row before headers

Sample files can be regenerated with `python generate_samples.py` (writes to `sample_data/`).
