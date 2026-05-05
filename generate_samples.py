"""
Generate sample DAF and cost sheet for testing.
Run: python generate_samples.py
"""
import pandas as pd
from datetime import date

daf_data = {
    "daf_ref_no":        ["DAF-2024-001"] * 5,
    "lob":               ["Residential"] * 5,
    "channel":           ["Dealer"] * 5,
    "state":             ["Maharashtra"] * 5,
    "project_name":      ["Sunshine Heights"] * 5,
    "developer_name":    ["Prestige Group"] * 5,
    "dealer_name":       ["ABC Tiles Pvt Ltd"] * 5,
    "zonal_coordinator": ["Rahul Mehta"] * 5,
    "submitted_date":    [date(2024, 6, 15)] * 5,
    "sku_code":          ["GVT6060A", "GVT8080B", "PGVT6060C", "PGVT8080D", "NAT6060E"],
    "boxes":             [200, 150, 100, 80, 120],
    "nef":               [38.0, 55.0, 72.0, 95.0, 28.0],
    "list_value":        [9500, 9800, 10200, 15000, 5800],
}

cost_data = {
    "sku_code":     ["GVT6060A", "GVT8080B", "PGVT6060C", "PGVT8080D", "NAT6060E"],
    "area_per_box": [11.16, 17.28, 11.16, 17.28, 11.16],
    "buying_price": [32.0, 48.0, 60.0, 78.0, 26.0],
}

pd.DataFrame(daf_data).to_excel("sample_data/sample_daf.xlsx", index=False)
pd.DataFrame(cost_data).to_excel("sample_data/sample_cost_sheet.xlsx", index=False)
print("Sample files written to sample_data/")
