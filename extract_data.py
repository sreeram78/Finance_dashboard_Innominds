#!/usr/bin/env python3
"""Refresh Nymi CFO dashboard data from the monthly XLSB MIS workbook.

Usage:
  python scripts/extract_data.py "Statement Of Financial MIS Sept'26 Corp.xlsb"

Writes data/financials.json.  The script uses the Summary SOP and Detailed SOP
rows from the workbook and preserves the schema consumed by index.html.
"""
from pathlib import Path
from datetime import datetime, timedelta
import json, re, sys
from pyxlsb import open_workbook

MONTHS = ["Apr-26","May-26","Jun-26","Jul-26","Aug-26","Sep-26","Oct-26","Nov-26","Dec-26","Jan-27","Feb-27","Mar-27"]

# Detailed SOP rows are 0-based pyxlsb row indexes from the Sept-26 workbook.
ROW = {
    "product_rev":15, "product_cogs":22, "product_gp":23,
    "service_rev":28, "service_cogs":31, "service_gm":32,
    "sub_rev":36, "sub_cogs":39, "sub_gm":40,
    "total_rev":43, "total_cogs":44, "total_gm":45, "total_gm_pct":46,
    "payroll_onsite":49, "payroll_offshore":50, "payroll_pss_nonbillable":51,
    "third_party_contractors":52, "sales_commission_bonus":53,
    "travel_entertainment":54, "marketing_campaign_events":55,
    "communication":56, "dues_subscriptions":57, "rent_utilities":58,
    "professional_fee":59, "insurance":60, "other_expenses_ga":61,
    "total_sga":63, "ebitda":65, "ebitda_pct":66, "finance_charges":67,
    "depreciation":68, "other_income_exp":69, "pbt":70, "pbt_pct":71,
    "tax":72, "pat":73, "pat_pct":74,
}


def clean(v):
    if v is None: return None
    if isinstance(v, str):
        s=v.strip()
        if not s or s.lower() in {"#n/a","n/a","na","-","0x17","0xf"}: return None
        return s
    return v


def num(v):
    v=clean(v)
    if v is None: return 0.0
    if isinstance(v,(int,float)): return float(v)
    s=str(v).replace(',','').replace('$','').strip()
    if s.startswith('(') and s.endswith(')'): s='-'+s[1:-1]
    try: return float(s)
    except: return 0.0


def header(v):
    if isinstance(v,(int,float)):
        try: return (datetime(1899,12,30)+timedelta(days=float(v))).strftime('%b-%y')
        except: pass
    return str(clean(v) or '').replace('  ',' ').strip()


def read_sheet(wb, name):
    with wb.get_sheet(name) as ws:
        return [[c.v for c in row] for row in ws.rows()]


def make_rows(rows, header_idx):
    raw=rows[header_idx]
    headers=[]; seen={}
    for v in raw:
        h=header(v) or 'Blank'; n=seen.get(h,0)+1; seen[h]=n
        headers.append(h if n==1 else f'{h}__{n}')
    out=[]
    for i,row in enumerate(rows[header_idx+1:], start=header_idx+2):
        vals=list(row)+[None]*max(0,len(headers)-len(row))
        label=clean(vals[2] if len(vals)>2 else None)
        if label is None and not any(clean(x) is not None for x in vals): continue
        data=[num(x) if isinstance(clean(x),(int,float)) else clean(x) for x in vals[:len(headers)]]
        out.append({'row':i,'sno':clean(vals[1] if len(vals)>1 else None),'particulars':str(label or ''),'values':data})
    return {'headers':headers,'header_row':header_idx+1,'particulars_column':3,'rows':out}


def value(rows, idx, col): return num(rows[idx][col]) if idx < len(rows) and col < len(rows[idx]) else 0.0


def main():
    workbook=Path(sys.argv[1]).expanduser().resolve() if len(sys.argv)>1 else Path(__file__).resolve().parents[1]/"Statement Of Financial MIS Sept'26 Corp.xlsb"
    if not workbook.exists(): raise FileNotFoundError(workbook)
    root=Path(__file__).resolve().parents[1]; outdir=root/'data'; outdir.mkdir(exist_ok=True)
    with open_workbook(str(workbook)) as wb:
        summary_raw=read_sheet(wb,'Summary SOP')
        detail_raw=read_sheet(wb,'Detailed SOP')
    summary=make_rows(summary_raw,5)
    detailed=make_rows(detail_raw,5)

    # Current FY is column D (3), prior FY column E (4), Apr is K (10).
    fy=3; prior=4; m0=10
    actual_months=[]
    for i,m in enumerate(MONTHS):
        actual_months.append(m if any(abs(num(detail_raw[ROW['total_rev']][m0+i]),)>0 for _ in [0]) else m)
    # September is the last populated month; keep future months as zero.
    actual_count=6

    revenue={
      'product':[value(detail_raw,ROW['product_rev'],m0+i) if i<actual_count else 0 for i in range(12)],
      'service':[value(detail_raw,ROW['service_rev'],m0+i) if i<actual_count else 0 for i in range(12)],
      'subscription':[value(detail_raw,ROW['sub_rev'],m0+i) if i<actual_count else 0 for i in range(12)]}
    revenue['total']=[revenue['product'][i]+revenue['service'][i]+revenue['subscription'][i] for i in range(12)]
    cogs=[value(detail_raw,ROW['total_cogs'],m0+i) if i<actual_count else 0 for i in range(12)]
    gm=[value(detail_raw,ROW['total_gm'],m0+i) if i<actual_count else 0 for i in range(12)]
    gm_pct=[value(detail_raw,ROW['total_gm_pct'],m0+i) if i<actual_count else 0 for i in range(12)]
    sga=[value(detail_raw,ROW['total_sga'],m0+i) if i<actual_count else 0 for i in range(12)]
    ebitda=[value(detail_raw,ROW['ebitda'],m0+i) if i<actual_count else 0 for i in range(12)]
    ebitda_pct=[value(detail_raw,ROW['ebitda_pct'],m0+i) if i<actual_count else 0 for i in range(12)]
    finance=[value(detail_raw,ROW['finance_charges'],m0+i) if i<actual_count else 0 for i in range(12)]
    depr=[value(detail_raw,ROW['depreciation'],m0+i) if i<actual_count else 0 for i in range(12)]
    other=[value(detail_raw,ROW['other_income_exp'],m0+i) if i<actual_count else 0 for i in range(12)]
    pbt=[value(detail_raw,ROW['pbt'],m0+i) if i<actual_count else 0 for i in range(12)]

    ytd={'revenue':value(detail_raw,ROW['total_rev'],fy),'cogs':value(detail_raw,ROW['total_cogs'],fy),'gross_margin':value(detail_raw,ROW['total_gm'],fy),'gm_pct':value(detail_raw,ROW['total_gm_pct'],fy),'sga':value(detail_raw,ROW['total_sga'],fy),'ebitda':value(detail_raw,ROW['ebitda'],fy),'ebitda_pct':value(detail_raw,ROW['ebitda_pct'],fy),'finance_charges':value(detail_raw,ROW['finance_charges'],fy),'depreciation':value(detail_raw,ROW['depreciation'],fy),'other_income_exp':value(detail_raw,ROW['other_income_exp'],fy),'pbt':value(detail_raw,ROW['pbt'],fy),'pbt_pct':value(detail_raw,ROW['pbt_pct'],fy)}
    prior_full={'revenue':value(detail_raw,ROW['total_rev'],prior),'gross_margin':value(detail_raw,ROW['total_gm'],prior),'gm_pct':value(detail_raw,ROW['total_gm_pct'],prior),'ebitda':value(detail_raw,ROW['ebitda'],prior),'ebitda_pct':value(detail_raw,ROW['ebitda_pct'],prior),'pbt':value(detail_raw,ROW['pbt'],prior)}

    sga_rows={'payroll onsite':ROW['payroll_onsite'],'payroll offshore':ROW['payroll_offshore'],'payroll pss nonbillable':ROW['payroll_pss_nonbillable'],'third-party contractors':ROW['third_party_contractors'],'sales commission & bonus':ROW['sales_commission_bonus'],'travel & entertainment':ROW['travel_entertainment'],'marketing':ROW['marketing_campaign_events'],'communication':ROW['communication'],'dues & subscriptions':ROW['dues_subscriptions'],'rent & utilities':ROW['rent_utilities'],'professional fees':ROW['professional_fee'],'insurance':ROW['insurance'],'other g&a':ROW['other_expenses_ga']}
    sga_breakdown={k:value(detail_raw,r,fy) for k,r in sga_rows.items()}

    payload={'source':workbook.name,'generated_at':datetime.now().isoformat(timespec='seconds'),'months':MONTHS,'revenue':revenue,'gm_pct':gm_pct,'cogs':cogs,'gross_margin':gm,'sga':sga,'ebitda':ebitda,'ebitda_pct':ebitda_pct,'pbt':pbt,'finance_charges':finance,'depreciation':depr,'other_income_exp':other,'ytd':ytd,'fy_prior_full':prior_full,'sga_breakdown_ytd':sga_breakdown,'summary_sop':summary,'detailed_sop':detailed,'balance_sheet':{'source_available':False,'note':'Balance sheet worksheet not present in September workbook.'},'cash_flow':{'source_available':False,'note':'Cash Flow worksheet not present in September workbook.'},'source_sheets':['Summary SOP','Detailed SOP']}
    output=outdir/'financials.json'; output.write_text(json.dumps(payload,indent=2,ensure_ascii=False,allow_nan=False),encoding='utf-8')
    print(f'Wrote {output}')
    print(json.dumps(ytd,indent=2))

if __name__=='__main__': main()
