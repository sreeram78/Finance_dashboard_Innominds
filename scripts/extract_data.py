#!/usr/bin/env python3
"""
extract_data.py
----------------
Reads the monthly Statement of Financial MIS workbook (.xlsb) and produces
a compact data/financials.json consumed by index.html.

Usage:
    python scripts/extract_data.py path/to/Statement_Of_Financial_MIS.xlsb

If no path is given, it looks for a single .xlsb file in the project root.

Source sheets expected in the workbook:
    - "Detailed SOP": monthly P&L detail, FY total, and prior-FY total
      (see ROW_MAP below for the exact row labels this script relies on)

Re-run this script whenever a new monthly MIS workbook is provisioned, then
commit the refreshed data/financials.json. index.html itself never changes.
"""

import sys
import json
import glob
import argparse
from pathlib import Path

from pyxlsb import open_workbook

# --------------------------------------------------------------------------
# Column layout on the "Detailed SOP" sheet (0-indexed, matches pyxlsb Cell.c)
# --------------------------------------------------------------------------
COL_FY_CURRENT = 3   # "FY 2026-27" (current fiscal year total to date)
COL_FY_PRIOR = 4     # "FY 25-26"   (prior fiscal year, full year total)
COL_MONTHS_START = 10   # first month column (April of the current FY)
N_MONTHS_ACTUAL = 5     # how many month columns currently have actuals loaded
FY_START_CALENDAR_YEAR = 26  # 2-digit year of the FY's April (FY2026-27 -> 26)
MONTH_LABELS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
                "Nov", "Dec", "Jan", "Feb", "Mar"]

# --------------------------------------------------------------------------
# Row labels on "Detailed SOP" -> row index. Update these if the workbook's
# row layout changes. Row indices are 0-based, matching pyxlsb's row order.
# --------------------------------------------------------------------------
ROW_MAP = {
    "product_rev": 15,
    "product_cogs": 22,
    "product_gp": 23,
    "service_rev": 28,
    "service_cogs": 31,
    "service_gm": 32,
    "sub_rev": 36,
    "sub_cogs": 39,
    "sub_gm": 40,
    "total_rev": 43,
    "total_cogs": 44,
    "total_gm": 45,
    "total_gm_pct": 46,
    "payroll_onsite": 49,
    "payroll_offshore": 50,
    "payroll_pss_nonbillable": 51,
    "third_party_contractors": 52,
    "sales_commission_bonus": 53,
    "travel_entertainment": 54,
    "marketing_campaign_events": 55,
    "communication": 56,
    "dues_subscriptions": 57,
    "rent_utilities": 58,
    "professional_fee": 59,
    "insurance": 60,
    "other_expenses_ga": 61,
    "total_sga": 63,
    "ebitda": 65,
    "ebitda_pct": 66,
    "finance_charges": 67,
    "depreciation": 68,
    "other_income_exp": 69,
    "pbt": 70,
    "pbt_pct": 71,
    "tax": 72,
    "pat": 73,
}

SGA_LABELS = {
    "sales_commission_bonus": "Sales commission & bonus",
    "payroll_onsite": "Payroll & benefits \u2013 onsite",
    "payroll_offshore": "Payroll & benefits \u2013 offshore",
    "payroll_pss_nonbillable": "Payroll PSS non-billable",
    "third_party_contractors": "Third-party contractors",
    "other_expenses_ga": "Other G&A expenses",
    "dues_subscriptions": "Dues & subscriptions",
    "travel_entertainment": "Travel & entertainment",
    "professional_fee": "Professional fees",
    "insurance": "Insurance",
    "marketing_campaign_events": "Marketing, campaigns & events",
    "rent_utilities": "Rent & utilities",
    "communication": "Communication",
}


def find_workbook(cli_path: str | None) -> Path:
    if cli_path:
        p = Path(cli_path)
        if not p.exists():
            sys.exit(f"File not found: {p}")
        return p
    candidates = glob.glob(str(Path(__file__).resolve().parents[1] / "*.xlsb"))
    if len(candidates) == 1:
        return Path(candidates[0])
    if not candidates:
        sys.exit("No .xlsb file found in project root. Pass a path explicitly.")
    sys.exit(f"Multiple .xlsb files found, pass one explicitly: {candidates}")


def read_sheet_rows(path: Path, sheet_name: str):
    with open_workbook(str(path)) as wb:
        with wb.get_sheet(sheet_name) as sheet:
            return list(sheet.rows())


def row_values(rows, row_idx: int, cols):
    """Return values for the given columns of a row, defaulting missing/None to 0."""
    cell_map = {c.c: c.v for c in rows[row_idx]}
    return [cell_map.get(c, 0) or 0 for c in cols]


def extract(path: Path) -> dict:
    rows = read_sheet_rows(path, "Detailed SOP")

    month_cols = list(range(COL_MONTHS_START, COL_MONTHS_START + N_MONTHS_ACTUAL))
    month_labels = []
    for i in range(N_MONTHS_ACTUAL):
        # MONTH_LABELS[0] = April (fiscal year start); rolls into Jan-Mar of the next calendar year
        year = FY_START_CALENDAR_YEAR + (1 if i >= 9 else 0)
        month_labels.append(f"{MONTH_LABELS[i % 12]}-{year:02d}")

    series = {}
    for key, row_idx in ROW_MAP.items():
        monthly = row_values(rows, row_idx, month_cols)
        ytd = row_values(rows, row_idx, [COL_FY_CURRENT])[0]
        prior_fy = row_values(rows, row_idx, [COL_FY_PRIOR])[0]
        series[key] = {"monthly": monthly, "ytd": ytd, "prior_fy_full": prior_fy}

    sga_breakdown = {
        label: series[key]["ytd"] for key, label in SGA_LABELS.items()
    }

    out = {
        "months": month_labels,
        "revenue": {
            "product": series["product_rev"]["monthly"],
            "service": series["service_rev"]["monthly"],
            "subscription": series["sub_rev"]["monthly"],
            "total": series["total_rev"]["monthly"],
        },
        "gm_pct": series["total_gm_pct"]["monthly"],
        "cogs": series["total_cogs"]["monthly"],
        "gross_margin": series["total_gm"]["monthly"],
        "sga": series["total_sga"]["monthly"],
        "ebitda": series["ebitda"]["monthly"],
        "ebitda_pct": series["ebitda_pct"]["monthly"],
        "pbt": series["pbt"]["monthly"],
        "finance_charges": series["finance_charges"]["monthly"],
        "depreciation": series["depreciation"]["monthly"],
        "other_income_exp": series["other_income_exp"]["monthly"],
        "ytd": {
            "revenue": series["total_rev"]["ytd"],
            "cogs": series["total_cogs"]["ytd"],
            "gross_margin": series["total_gm"]["ytd"],
            "gm_pct": series["total_gm_pct"]["ytd"],
            "sga": series["total_sga"]["ytd"],
            "ebitda": series["ebitda"]["ytd"],
            "ebitda_pct": series["ebitda_pct"]["ytd"],
            "finance_charges": series["finance_charges"]["ytd"],
            "depreciation": series["depreciation"]["ytd"],
            "other_income_exp": series["other_income_exp"]["ytd"],
            "pbt": series["pbt"]["ytd"],
            "pbt_pct": (series["pbt"]["ytd"] / series["total_rev"]["ytd"]
                        if series["total_rev"]["ytd"] else 0),
        },
        "fy_prior_full": {
            "revenue": series["total_rev"]["prior_fy_full"],
            "gross_margin": series["total_gm"]["prior_fy_full"],
            "gm_pct": (series["total_gm"]["prior_fy_full"] / series["total_rev"]["prior_fy_full"]
                       if series["total_rev"]["prior_fy_full"] else 0),
            "ebitda": series["ebitda"]["prior_fy_full"],
            "ebitda_pct": (series["ebitda"]["prior_fy_full"] / series["total_rev"]["prior_fy_full"]
                           if series["total_rev"]["prior_fy_full"] else 0),
            "pbt": series["pbt"]["prior_fy_full"],
        },
        "sga_breakdown_ytd": sga_breakdown,
    }
    return out


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", nargs="?", help="Path to the .xlsb MIS file")
    parser.add_argument("-o", "--output", default=None,
                         help="Output JSON path (default: <project_root>/data/financials.json)")
    args = parser.parse_args()

    path = find_workbook(args.workbook)
    data = extract(path)

    out_path = Path(args.output) if args.output else \
        Path(__file__).resolve().parents[1] / "data" / "financials.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(data, indent=2))

    print(f"Wrote {out_path} from {path.name}")
    print(f"  Revenue YTD: {data['ytd']['revenue']:,.0f}")
    print(f"  EBITDA YTD:  {data['ytd']['ebitda']:,.0f}")
    print(f"  PBT YTD:     {data['ytd']['pbt']:,.0f}")


if __name__ == "__main__":
    main()
