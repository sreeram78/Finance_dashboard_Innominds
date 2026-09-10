#!/usr/bin/env python3
"""
Nymi Inc MIS extractor
----------------------
Reads the XLSB MIS workbook and writes data/financials.json.

Sources:
  - Summary SOP
  - Detailed SOP
  - Balance sheet

The Detailed SOP extraction is header-driven rather than relying on fixed
row/column positions. This makes it safer when rows are inserted or moved
in the Excel workbook.

Usage:
    python scripts/extract_data.py "Statement Of Financial MIS Aug26 Provisional V3 1.xlsb"
"""

from pathlib import Path
from datetime import datetime, timedelta
import json
import math
import re
import sys

from pyxlsb import open_workbook


MONTHS = ["Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26",
          "Sep-26", "Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27", "Mar-27"]

# Excel/XLSB serial date conversion.
def excel_date(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        try:
            # Excel 1900 date system
            return datetime(1899, 12, 30) + timedelta(days=float(value))
        except Exception:
            return None
    return None


def clean(value):
    if value is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        # pyxlsb can expose malformed/future cells such as 0x17.
        if value.lower() in {"0x17", "#n/a", "#na", "n/a", "na", "-"}:
            return None
    return value


def number(value):
    value = clean(value)
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(",", "").replace("$", "").strip()
    s = s.replace("(", "-").replace(")", "")
    try:
        return float(s)
    except Exception:
        return 0.0


def row_values(sheet, row_num):
    """Return one XLSB row as a list of raw values."""
    with open_workbook(WORKBOOK) as wb:
        with wb.get_sheet(sheet) as ws:
            for i, row in enumerate(ws.rows(), start=1):
                if i == row_num:
                    return [c.v for c in row]
    return []


def read_sheet(sheet_name):
    with open_workbook(WORKBOOK) as wb:
        with wb.get_sheet(sheet_name) as ws:
            return [[c.v for c in row] for row in ws.rows()]


def find_header_row(rows, required_terms):
    """Find the row containing the highest number of required header terms."""
    best = None
    best_score = -1
    for i, row in enumerate(rows):
        texts = [str(clean(v) or "").strip().lower() for v in row]
        score = sum(
            any(term.lower() in text for text in texts)
            for term in required_terms
        )
        if score > best_score:
            best_score = score
            best = i
    return best


def normalize_header(value):
    value = clean(value)
    if value is None:
        return ""
    d = excel_date(value)
    if d:
        return d.strftime("%b-%y")
    return re.sub(r"\s+", " ", str(value)).strip()


def make_unique_headers(raw_headers):
    """Preserve duplicate labels while giving them unique internal names."""
    result = []
    seen = {}
    for h in raw_headers:
        h = normalize_header(h)
        if not h:
            h = "Blank"
        n = seen.get(h, 0)
        seen[h] = n + 1
        result.append(h if n == 0 else f"{h}__{n+1}")
    return result


def extract_detailed_sop():
    """
    Extract Detailed SOP using the worksheet's own headers.

    Output:
      {
        "headers": [...],
        "rows": [
          {
            "row": <1-based Excel row>,
            "particulars": "...",
            "values": [...]
          }
        ]
      }

    The values array follows the same header order as the workbook.
    """
    rows = read_sheet("Detailed SOP")

    # Find the row containing the financial period headers.
    header_idx = find_header_row(
        rows,
        ["FY 25-26", "FY 2026-27", "Apr", "May", "Jun", "Jul", "Aug"]
    )
    if header_idx is None:
        raise RuntimeError("Could not locate Detailed SOP period header row.")

    raw_header = rows[header_idx]
    headers = make_unique_headers(raw_header)

    # Find the particulars column. Usually column B, but discover it.
    particulars_col = 1
    for j, v in enumerate(raw_header):
        t = str(clean(v) or "").lower()
        if "particular" in t or "description" in t or "account" in t:
            particulars_col = j
            break

    extracted = []
    for excel_row, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if not row:
            continue

        # Extend short rows so indexes always match headers.
        values = list(row) + [None] * max(0, len(headers) - len(row))
        particulars = clean(values[particulars_col])

        # Ignore completely empty rows.
        numeric_or_text = [clean(v) for v in values]
        if particulars is None and not any(v is not None for v in numeric_or_text):
            continue

        # Keep section/subtotal rows as well as numeric rows.
        data = []
        for v in values:
            x = clean(v)
            data.append(number(x) if isinstance(x, (int, float)) or
                        (isinstance(x, str) and re.match(r"^[\(\)\-\$,\d\. ]+$", x))
                        else x)

        extracted.append({
            "row": excel_row,
            "particulars": str(particulars) if particulars is not None else "",
            "values": data
        })

    return {
        "headers": headers,
        "header_row": header_idx + 1,
        "particulars_column": particulars_col + 1,
        "rows": extracted
    }


def extract_summary_sop():
    """Same header-driven extraction for Summary SOP."""
    rows = read_sheet("Summary SOP")
    header_idx = find_header_row(
        rows,
        ["FY 25-26", "FY 2026-27", "Apr", "May", "Jun", "Jul", "Aug"]
    )
    if header_idx is None:
        raise RuntimeError("Could not locate Summary SOP period header row.")

    raw_header = rows[header_idx]
    headers = make_unique_headers(raw_header)

    particulars_col = 1
    for j, v in enumerate(raw_header):
        t = str(clean(v) or "").lower()
        if "particular" in t or "description" in t or "account" in t:
            particulars_col = j
            break

    extracted = []
    for excel_row, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        values = list(row) + [None] * max(0, len(headers) - len(row))
        if not values:
            continue
        particulars = clean(values[particulars_col])
        if particulars is None and not any(clean(v) is not None for v in values):
            continue

        data = []
        for v in values:
            x = clean(v)
            data.append(number(x) if isinstance(x, (int, float)) else x)

        extracted.append({
            "row": excel_row,
            "particulars": str(particulars) if particulars is not None else "",
            "values": data
        })

    return {
        "headers": headers,
        "header_row": header_idx + 1,
        "particulars_column": particulars_col + 1,
        "rows": extracted
    }


def extract_balance_sheet():
    rows = read_sheet("Balance sheet")
    if len(rows) < 4:
        return {}

    # The supplied MIS has period headers on row 4.
    header_idx = 3
    headers = make_unique_headers(rows[header_idx])
    particulars_col = 1

    result_rows = []
    for excel_row, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        values = list(row) + [None] * max(0, len(headers) - len(row))
        label = clean(values[particulars_col]) if len(values) > particulars_col else None
        if label is None:
            continue
        nums = [number(v) for v in values]
        result_rows.append({
            "row": excel_row,
            "particulars": str(label),
            "values": nums
        })

    # Keep only populated actual period columns.
    actual_headers = []
    actual_indices = []
    for i, h in enumerate(headers):
        if h in MONTHS or re.match(r"^(Jul|Aug)-\d\d$", h):
            col_vals = [r["values"][i] for r in result_rows if i < len(r["values"])]
            if any(abs(v) > 1e-9 for v in col_vals):
                actual_headers.append(h)
                actual_indices.append(i)

    def find(label):
        for r in result_rows:
            if r["particulars"].strip().lower() == label.lower():
                return r
        return None

    wanted = {
        "ppe": "Property, Plant & Equipment",
        "trade_receivables": "Trade Receivables",
        "inventory": "Inventory",
        "cash": "Cash & Bank",
        "other_assets": "Other Assets",
        "total_assets": "Total Assets",
        "shareholders_fund": "Shareholder's Fund",
        "debts": "Debts",
        "trade_payables": "Trade Payables",
        "deferred_revenue": "Deferred Revenue",
        "other_liabilities": "Other Liabilities",
        "total_equity_liability": "Total Eq. & Liability",
    }

    output = {"months": actual_headers}
    for key, label in wanted.items():
        r = find(label)
        output[key] = [
            r["values"][i] if r and i < len(r["values"]) else 0
            for i in actual_indices
        ]

    # These are single values in the current worksheet.
    for key, label in {
        "change_wc": "Change in WC",
        "change_capex": "Change in Capex",
        "fcf": "FCF",
    }.items():
        r = find(label)
        output[key] = r["values"][actual_indices[-1]] if r and actual_indices else 0

    return output


def main():
    global WORKBOOK

    if len(sys.argv) > 1:
        WORKBOOK = Path(sys.argv[1]).expanduser().resolve()
    else:
        # Default: workbook in project root.
        WORKBOOK = Path(__file__).resolve().parents[1] / \
            "Statement Of Financial MIS Aug26 Provisional V3 1.xlsb"

    if not WORKBOOK.exists():
        raise FileNotFoundError(f"Workbook not found: {WORKBOOK}")

    project_root = Path(__file__).resolve().parents[1]
    data_dir = project_root / "data"
    data_dir.mkdir(parents=True, exist_ok=True)

    detailed = extract_detailed_sop()
    summary = extract_summary_sop()
    bs = extract_balance_sheet()

    # Preserve the detailed/summary workbook data in a clean, explicit schema.
    payload = {
        "source": WORKBOOK.name,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "months": MONTHS,
        "summary_sop": summary,
        "detailed_sop": detailed,
        "balance_sheet": bs,
    }

    output = data_dir / "financials.json"
    output.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False, allow_nan=False),
        encoding="utf-8"
    )

    print(f"Workbook : {WORKBOOK}")
    print(f"Summary SOP rows  : {len(summary['rows'])}")
    print(f"Detailed SOP rows : {len(detailed['rows'])}")
    print(f"BS periods        : {bs.get('months', [])}")
    print(f"Wrote             : {output}")


if __name__ == "__main__":
    main()
