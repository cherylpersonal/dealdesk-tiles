# DAF Processor

Internal tool to process Discount Approval Forms, calculate margins, and recommend approval levels.

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Generate sample data (optional — for testing)
python generate_samples.py

# 3. Run the app
streamlit run app.py
```

The app opens at http://localhost:8501

---

## What it does

1. Upload a DAF Excel file and a Cost Sheet Excel file
2. Joins them on `sku_code`
3. Calculates per-SKU: total_area, revenue, cost, margin_value, margin_percent
4. Calculates deal totals
5. Applies approval matrix → recommends Deal Desk / PM Head / CEO
6. Generates downloadable Excel with two sheets:
   - **Calculations** — all data + totals + approval recommendation
   - **Tracker** — one structured row for your tracking sheet

---

## Required Column Names (exact)

### DAF file
| Column | Type | Notes |
|---|---|---|
| `daf_ref_no` | string | DAF reference number |
| `lob` | string | Line of Business |
| `channel` | string | Sales channel |
| `state` | string | State |
| `project_name` | string | |
| `developer_name` | string | |
| `dealer_name` | string | |
| `zonal_coordinator` | string | |
| `submitted_date` | date | Triggers 6-hr SLA |
| `sku_code` | string | **Join key** — must match cost sheet |
| `boxes` | number | Quantity |
| `nef` | number | Net effective price per sq ft |
| `list_value` | number | Optional — used for avg discount % |

### Cost Sheet
| Column | Type | Notes |
|---|---|---|
| `sku_code` | string | **Join key** |
| `area_per_box` | number | Sq ft per box |
| `buying_price` | number | Cost per sq ft |

---

## Approval Matrix

| Level | SKU condition | Overall condition |
|---|---|---|
| Deal Desk | All SKUs ≥ 10% | Overall ≥ 20% |
| PM Head | All SKUs ≥ 5% | Overall ≥ 10% |
| CEO | Any SKU < 5% | OR overall < 10% |

---

## Project Structure

```
daf-processor/
├── app.py                  # Streamlit UI
├── core/
│   ├── parser.py           # parse_daf(), parse_cost_sheet()
│   ├── calculator.py       # merge_data(), calculate_margins()
│   ├── approval.py         # evaluate_approval()
│   ├── exporter.py         # generate_output_file(), build_tracker_row()
│   └── validators.py       # edge case checks
├── sample_data/
│   ├── sample_daf.xlsx
│   └── sample_cost_sheet.xlsx
├── generate_samples.py
├── requirements.txt
└── README.md
```
