import pandas as pd
import re
import json
import os
import sys

EXCEL_PATH = "Data/cfpb_data.xlsx"
HTML_PATH = "cfpb_dashboard.html"

REGION_MAP = {
    "TX":"South","FL":"South","GA":"South","NC":"South","VA":"South","MD":"South",
    "DC":"South","TN":"South","AL":"South","SC":"South","MS":"South","AR":"South",
    "LA":"South","KY":"South","WV":"South","OK":"South","DE":"South",
    "CA":"West","WA":"West","AZ":"West","CO":"West","OR":"West","NV":"West",
    "UT":"West","NM":"West","HI":"West","AK":"West","ID":"West","MT":"West","WY":"West",
    "NY":"Northeast","PA":"Northeast","NJ":"Northeast","MA":"Northeast","CT":"Northeast",
    "RI":"Northeast","NH":"Northeast","VT":"Northeast","ME":"Northeast",
    "IL":"Midwest","OH":"Midwest","MI":"Midwest","MN":"Midwest","MO":"Midwest",
    "WI":"Midwest","IN":"Midwest","IA":"Midwest","KS":"Midwest","NE":"Midwest",
    "ND":"Midwest","SD":"Midwest",
}

MONTH_ABBR = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def fmt_num(n):
    return f"{int(n):,}"

def fmt_pct(p):
    return f"{p:.1f}%"

def fmt_days(d):
    return f"{d:.1f} days"

def excel_serial_to_date(s):
    return pd.to_datetime("1899-12-30") + pd.to_timedelta(s, unit="D")

def load_data():
    df = pd.read_excel(EXCEL_PATH)
    df["Date received"] = excel_serial_to_date(df["Date received"])
    df["Company_Response_Date"] = excel_serial_to_date(df["Company_Response_Date"])
    df["year"] = df["Date received"].dt.year
    df["month"] = df["Date received"].dt.month
    df["ym"] = df["year"] * 100 + df["month"]
    return df

def calc_kpis(df):
    total = len(df)
    timely_pct = (df["Timely response?"] == "Yes").mean() * 100
    avg_days = df["Response_Time_Days"].mean()
    peak_year = int(df["year"].value_counts().idxmax())
    date_min = df["Date received"].min()
    date_max = df["Date received"].max()
    date_range = f"{MONTH_ABBR[date_min.month]} {date_min.year} – {MONTH_ABBR[date_max.month]} {date_max.year}"
    top_state_code = df["State"].value_counts().idxmax()
    return {
        "total": fmt_num(total),
        "timely_pct": fmt_pct(timely_pct),
        "avg_days": fmt_days(avg_days),
        "peak_year": str(peak_year),
        "date_range": date_range,
        "top_state": top_state_code,
    }

def calc_monthly_rolling(df):
    grp = df.groupby(["year","month"]).size().reset_index(name="c")
    grp = grp.sort_values(["year","month"])
    items = [f"{{y:{r.year},m:{r.month},c:{r.c}}}" for _, r in grp.iterrows()]
    return "[" + ",".join(items) + "]"

def calc_yearly(df):
    grp = df.groupby("year").size()
    years = sorted(grp.index.tolist())
    labels = "[" + ",".join(f"'{y}'" for y in years) + "]"
    data = "[" + ",".join(str(grp[y]) for y in years) + "]"
    return labels, data

def calc_state_data(df):
    states = {}
    for state, sg in df.groupby("State"):
        if not isinstance(state, str) or len(state) != 2:
            continue
        total = len(sg)
        unresolved = int((sg["Company response to consumer"].str.contains("In progress", na=False)).sum())
        untimely = int((sg["Timely response?"] == "No").sum())
        avg_rt = float(sg["Response_Time_Days"].mean()) if "Response_Time_Days" in sg else 0.0
        timely_pct = float((sg["Timely response?"] == "Yes").mean() * 100)
        region = REGION_MAP.get(state, "Other")
        unresolved_pct = round(unresolved / total * 100, 1) if total > 0 else 0.0
        states[state] = {
            "total": total,
            "unresolved": unresolved,
            "untimely": untimely,
            "avg_rt": round(avg_rt, 1),
            "region": region,
            "timely_pct": round(timely_pct, 1),
            "unresolved_pct": unresolved_pct,
        }
    lines = []
    for k, v in sorted(states.items()):
        lines.append(f'  "{k}":{{"total":{v["total"]},"unresolved":{v["unresolved"]},"untimely":{v["untimely"]},"avg_rt":{v["avg_rt"]},"region":"{v["region"]}","timely_pct":{v["timely_pct"]},"unresolved_pct":{v["unresolved_pct"]}}}')
    return "{\n" + ",\n".join(lines) + "\n}"

def calc_subprod_data(df):
    result = {}
    for prod, pg in df.groupby("Product"):
        sp_counts = pg["Sub-product"].value_counts()
        sp_counts = sp_counts[sp_counts > 0]
        labels = list(sp_counts.index[:8])
        data = [int(x) for x in sp_counts.values[:8]]
        result[prod] = {"labels": labels, "data": data}
    lines = []
    for k, v in result.items():
        lbl = json.dumps(v["labels"])
        dat = json.dumps(v["data"])
        lines.append(f'  {json.dumps(k)}:{{labels:{lbl},data:{dat}}}')
    return "{\n" + ",\n".join(lines) + "\n}"

def calc_product_data(df):
    prod_counts = df["Product"].value_counts()
    products = prod_counts.index.tolist()
    totals = prod_counts.values.tolist()
    unresolved = []
    untimely = []
    for p in products:
        pg = df[df["Product"] == p]
        unresolved.append(int((pg["Company response to consumer"].str.contains("In progress", na=False)).sum()))
        untimely.append(int((pg["Timely response?"] == "No").sum()))
    prod_labels = json.dumps(products)
    prod_totals = "[" + ",".join(str(x) for x in totals) + "]"
    prod_unresolved = "[" + ",".join(str(x) for x in unresolved) + "]"
    prod_untimely = "[" + ",".join(str(x) for x in untimely) + "]"
    return products, prod_labels, prod_totals, prod_unresolved, prod_untimely

def calc_channel_data(df):
    ch = df["Submitted via"].value_counts()
    labels = json.dumps(ch.index.tolist())
    data = "[" + ",".join(str(x) for x in ch.values) + "]"
    return labels, data

def calc_resolution_data(df):
    res = df["Company response to consumer"].value_counts()
    labels = json.dumps(res.index.tolist())
    data = "[" + ",".join(str(x) for x in res.values) + "]"
    return labels, data

def calc_region_data(df):
    df2 = df.copy()
    df2["region"] = df2["State"].map(REGION_MAP).fillna("Other")
    rg = df2.groupby("region").size()
    order = ["South","West","Northeast","Midwest"]
    data = "[" + ",".join(str(rg.get(r, 0)) for r in order) + "]"
    return data

def calc_per_product_monthly(df, products):
    TARGET_PRODUCTS = {
        "Checking or savings account": "M_CHECK",
        "Credit card or prepaid card": "M_CC",
        "Credit reporting, credit repair services, or other personal consumer reports": "M_CR",
        "Mortgage": "M_MORT",
        "Money transfer, virtual currency, or money service": "M_MT",
    }
    ym_range = sorted(df["ym"].unique())
    result = {}
    for prod, var in TARGET_PRODUCTS.items():
        pg = df[df["Product"] == prod]
        counts = pg.groupby("ym").size()
        vals = [int(counts.get(ym, 0)) for ym in ym_range]
        result[var] = "[" + ",".join(str(v) for v in vals) + "]"
    return result

def calc_issues_data(df):
    top_issues = df["Issue"].value_counts().head(10)
    issue_names = json.dumps(top_issues.index.tolist())
    years = sorted(df["year"].unique())
    issue_year_data = []
    for issue in top_issues.index:
        ig = df[df["Issue"] == issue]
        year_counts = ig.groupby("year").size()
        vals = "[" + ",".join(str(int(year_counts.get(y, 0))) for y in years) + "]"
        issue_year_data.append(vals)
    return issue_names, issue_year_data, years

def replace_block(html, pattern, replacement, flags=re.DOTALL):
    new_html, n = re.subn(pattern, replacement, html, count=1, flags=flags)
    if n == 0:
        print(f"  WARNING: pattern not found: {pattern[:60]}")
    return new_html

def main():
    print("Loading Excel...")
    df = load_data()
    print(f"  {len(df):,} rows loaded")

    print("Calculating metrics...")
    kpis = calc_kpis(df)
    mr = calc_monthly_rolling(df)
    ry_labels, ry_data = calc_yearly(df)
    state_data = calc_state_data(df)
    subprod_data = calc_subprod_data(df)
    products, prod_labels, prod_totals, prod_unresolved, prod_untimely = calc_product_data(df)
    ch_labels, ch_data = calc_channel_data(df)
    res_labels, res_data = calc_resolution_data(df)
    region_data = calc_region_data(df)
    per_prod = calc_per_product_monthly(df, products)
    issue_names, issue_year_data, issue_years = calc_issues_data(df)

    print("Reading HTML...")
    html = open(HTML_PATH, encoding="utf-8").read()

    print("Patching KPIs...")
    html = replace_block(html, r'class="kpi-val[^"]*"[^>]*>\s*62,516\s*<', f'class="kpi-val fira">{kpis["total"]}<')
    html = replace_block(html, r'class="kpi-val[^"]*"[^>]*>\s*93\.8%\s*<', f'class="kpi-val fira">{kpis["timely_pct"]}<')
    html = replace_block(html, r'class="kpi-val[^"]*"[^>]*>\s*15\.1 days\s*<', f'class="kpi-val fira">{kpis["avg_days"]}<')
    html = replace_block(html, r'class="kpi-val[^"]*"[^>]*>\s*2022\s*<', f'class="kpi-val fira">{kpis["peak_year"]}<')

    print("Patching date range...")
    html = re.sub(
        r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{4}\s*[–\-]\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{4}',
        kpis["date_range"],
        html
    )

    print("Patching JS data blocks...")
    html = replace_block(html, r'const MR\s*=\s*\[[^\]]+\]', f'const MR = {mr}')
    html = replace_block(html, r'const RY_LABELS\s*=\s*\[[^\]]+\]', f'const RY_LABELS = {ry_labels}')
    html = replace_block(html, r'const STATE_DATA\s*=\s*\{[^;]+\}', f'const STATE_DATA = {state_data}')
    html = replace_block(html, r'const SUBPROD_DATA\s*=\s*\{[^;]+\}', f'const SUBPROD_DATA = {subprod_data}')

    for var, arr in per_prod.items():
        html = replace_block(html, rf'const {var}\s*=\s*\[[^\]]+\]', f'const {var} = {arr}')

    print("Writing HTML...")
    open(HTML_PATH, "w", encoding="utf-8").write(html)
    print("Done. Dashboard refreshed.")
    print(f"  Total: {kpis['total']} | Timely: {kpis['timely_pct']} | Avg days: {kpis['avg_days']} | Peak: {kpis['peak_year']}")
    print(f"  Date range: {kpis['date_range']}")

if __name__ == "__main__":
    main()
