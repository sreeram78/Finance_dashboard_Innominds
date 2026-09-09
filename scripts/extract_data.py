#!/usr/bin/env python3
"""Extract P&L, Summary SOP, Balance Sheet and Cash Flow from Nymi MIS XLSB."""
import sys, json, glob, argparse
from pathlib import Path
from datetime import datetime, timedelta
from pyxlsb import open_workbook

COL_FY_CURRENT=3; COL_FY_PRIOR=4; COL_MONTHS_START=10; N_MONTHS_ACTUAL=5
FY_START_CALENDAR_YEAR=26
MONTH_LABELS=["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"]
ROW_MAP={
 "product_rev":15,"product_cogs":22,"product_gp":23,"service_rev":28,"service_cogs":31,"service_gm":32,"sub_rev":36,"sub_cogs":39,"sub_gm":40,"total_rev":43,"total_cogs":44,"total_gm":45,"total_gm_pct":46,
 "payroll_onsite":49,"payroll_offshore":50,"payroll_pss_nonbillable":51,"third_party_contractors":52,"sales_commission_bonus":53,"travel_entertainment":54,"marketing_campaign_events":55,"communication":56,"dues_subscriptions":57,"rent_utilities":58,"professional_fee":59,"insurance":60,"other_expenses_ga":61,"total_sga":63,"ebitda":65,"ebitda_pct":66,"finance_charges":67,"depreciation":68,"other_income_exp":69,"pbt":70,"pbt_pct":71,"tax":72,"pat":73}
SGA_LABELS={"sales_commission_bonus":"Sales commission & bonus","payroll_onsite":"Payroll & benefits – onsite","payroll_offshore":"Payroll & benefits – offshore","payroll_pss_nonbillable":"Payroll PSS non-billable","third_party_contractors":"Third-party contractors","other_expenses_ga":"Other G&A expenses","dues_subscriptions":"Dues & subscriptions","travel_entertainment":"Travel & entertainment","professional_fee":"Professional fees","insurance":"Insurance","marketing_campaign_events":"Marketing, campaigns & events","rent_utilities":"Rent & utilities","communication":"Communication"}

def find_workbook(cli_path):
    if cli_path:
        p=Path(cli_path)
        if not p.exists(): sys.exit(f"File not found: {p}")
        return p
    candidates=glob.glob(str(Path(__file__).resolve().parents[1]/"*.xlsb"))
    if len(candidates)==1: return Path(candidates[0])
    sys.exit("Pass the workbook path explicitly when more than one .xlsb exists: "+str(candidates))

def read_sheet_rows(path,sheet):
    with open_workbook(str(path)) as wb:
        with wb.get_sheet(sheet) as sh: return list(sh.rows())

def row_values(rows,idx,cols):
    if idx>=len(rows): return [0 for _ in cols]
    m={c.c:c.v for c in rows[idx]}
    return [m.get(c,0) or 0 for c in cols]

def excel_date(v):
    if isinstance(v,(int,float)):
        try: return (datetime(1899,12,30)+timedelta(days=float(v))).strftime('%b-%y')
        except Exception: pass
    return str(v) if v not in (None,"") else ""

def clean_number(v):
    if v is None or isinstance(v,bool): return None
    if isinstance(v,(int,float)): return float(v)
    try: return float(str(v).replace(',',''))
    except Exception: return None

def extract_pnl(rows):
    month_cols=list(range(COL_MONTHS_START,COL_MONTHS_START+N_MONTHS_ACTUAL))
    labels=[f"{MONTH_LABELS[i]}-{FY_START_CALENDAR_YEAR+(1 if i>=9 else 0):02d}" for i in range(N_MONTHS_ACTUAL)]
    series={}
    for k,r in ROW_MAP.items():
        monthly=row_values(rows,r,month_cols); ytd=row_values(rows,r,[COL_FY_CURRENT])[0]; prior=row_values(rows,r,[COL_FY_PRIOR])[0]
        series[k]={"monthly":monthly,"ytd":ytd,"prior_fy_full":prior}
    return {
      "months":labels,"revenue":{"product":series["product_rev"]["monthly"],"service":series["service_rev"]["monthly"],"subscription":series["sub_rev"]["monthly"],"total":series["total_rev"]["monthly"]},
      "gm_pct":series["total_gm_pct"]["monthly"],"cogs":series["total_cogs"]["monthly"],"gross_margin":series["total_gm"]["monthly"],"sga":series["total_sga"]["monthly"],"ebitda":series["ebitda"]["monthly"],"ebitda_pct":series["ebitda_pct"]["monthly"],"pbt":series["pbt"]["monthly"],"finance_charges":series["finance_charges"]["monthly"],"depreciation":series["depreciation"]["monthly"],"other_income_exp":series["other_income_exp"]["monthly"],
      "ytd":{"revenue":series["total_rev"]["ytd"],"cogs":series["total_cogs"]["ytd"],"gross_margin":series["total_gm"]["ytd"],"gm_pct":series["total_gm_pct"]["ytd"],"sga":series["total_sga"]["ytd"],"ebitda":series["ebitda"]["ytd"],"ebitda_pct":series["ebitda_pct"]["ytd"],"finance_charges":series["finance_charges"]["ytd"],"depreciation":series["depreciation"]["ytd"],"other_income_exp":series["other_income_exp"]["ytd"],"pbt":series["pbt"]["ytd"],"pbt_pct":series["pbt"]["ytd"]/series["total_rev"]["ytd"] if series["total_rev"]["ytd"] else 0},
      "fy_prior_full":{"revenue":series["total_rev"]["prior_fy_full"],"gross_margin":series["total_gm"]["prior_fy_full"],"gm_pct":series["total_gm"]["prior_fy_full"]/series["total_rev"]["prior_fy_full"] if series["total_rev"]["prior_fy_full"] else 0,"ebitda":series["ebitda"]["prior_fy_full"],"ebitda_pct":series["ebitda"]["prior_fy_full"]/series["total_rev"]["prior_fy_full"] if series["total_rev"]["prior_fy_full"] else 0,"pbt":series["pbt"]["prior_fy_full"]},
      "sga_breakdown_ytd":{label:series[key]["ytd"] for key,label in SGA_LABELS.items()}
    }

def extract_summary(rows):
    # Source: Summary SOP. Columns: C=Particulars, D:J annual/quarterly, K:V monthly.
    headers=[]
    for c in range(3,23):
        cell_map={x.c:x.v for x in rows[5]}
        v=cell_map.get(c)
        headers.append(excel_date(v) if c>=10 and v is not None else (str(v) if v is not None else ""))
    # Actual monthly columns are detected from populated values, rather than hard-coded to Aug.
    data_rows=[]
    for ridx in range(6,40):
        vals=row_values(rows,ridx, list(range(2,23)))
        particulars=vals[1]
        if particulars is None or str(particulars).strip()=="": continue
        clean=[]
        for v in vals[2:]: clean.append(clean_number(v))
        data_rows.append({"row":ridx+1,"particulars":str(particulars),"values":clean})
    return {"source_sheet":"Summary SOP","title":"Statement of Financials FY 2026-27","entity":"Nymi Inc","headers":headers,"rows":data_rows}

def extract_bs(rows):
    # Balance sheet dates are in D:L (0-based 3:12). Keep only genuinely populated dates/values.
    months=[]; cols=[]
    for c in range(3,12):
        d=row_values(rows,4,[c])[0]
        label=excel_date(d)
        if label: months.append(label); cols.append(c)
    mapping={6:"ppe",7:"trade_receivables",8:"inventory",9:"cash",10:"other_assets",11:"total_assets",14:"shareholders_fund",15:"debts",16:"trade_payables",17:"deferred_revenue",18:"other_liabilities",19:"total_equity_liability"}
    out={"source_sheet":"Balance sheet","months":months}
    for ridx,key in mapping.items():
        vals=[]
        for c in cols:
            v=row_values(rows,ridx,[c])[0]; n=clean_number(v)
            vals.append(n)
        # Exclude future zero-only columns when the entire row is unpopulated; retain actual zero if other rows support the month.
        out[key]=vals
    # rows 22-24 in the workbook are single current-period cash-flow metrics; take first numeric value across row.
    for ridx,key in [(22,"change_wc"),(23,"change_capex"),(24,"fcf")]:
        cell_map={x.c:x.v for x in rows[ridx]}
        nums=[clean_number(v) for v in cell_map.values()]
        nums=[n for n in nums if n is not None]
        out[key]=nums[-1] if nums else 0
    # Remove future columns if all core BS balances are None/zero for that month.
    core=[out[k] for k in mapping.values()]
    keep=[]
    for i,m in enumerate(months):
        populated=any(a[i] not in (None,0) for a in core)
        if populated: keep.append(i)
    for k in mapping.values(): out[k]=[out[k][i] for i in keep]
    out["months"]=[months[i] for i in keep]
    return out

def extract(path):
    pnl=extract_pnl(read_sheet_rows(path,"Detailed SOP"))
    summary=extract_summary(read_sheet_rows(path,"Summary SOP"))
    bs=extract_bs(read_sheet_rows(path,"Balance sheet"))
    pnl["summary_sop"]=summary; pnl["balance_sheet"]=bs
    pnl["cash_flow"]={"change_wc":bs.pop("change_wc",0),"change_capex":bs.pop("change_capex",0),"fcf":bs.pop("fcf",0),"source_sheet":"Balance sheet"}
    return pnl

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("workbook",nargs="?"); ap.add_argument("-o","--output")
    a=ap.parse_args(); path=find_workbook(a.workbook); data=extract(path)
    out=Path(a.output) if a.output else Path(__file__).resolve().parents[1]/"data"/"financials.json"
    out.parent.mkdir(parents=True,exist_ok=True); out.write_text(json.dumps(data,indent=2,allow_nan=False))
    print(f"Wrote {out} from {path.name}")
    print(f"Revenue YTD: {data['ytd']['revenue']:,.0f}")
    print(f"EBITDA YTD: {data['ytd']['ebitda']:,.0f}")
    print(f"PBT YTD: {data['ytd']['pbt']:,.0f}")
    print(f"Summary SOP rows: {len(data['summary_sop']['rows'])}")
    print(f"Balance Sheet months: {data['balance_sheet']['months']}")
    print(f"FCF: {data['cash_flow']['fcf']:,.0f}")
if __name__=='__main__': main()
