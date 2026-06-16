#!/usr/bin/env python3
import sys, json
from datetime import datetime
from pathlib import Path

TODAY = datetime.today()

def parse_date(s):
    if not s: return None
    for fmt in ["%d/%m/%Y", "%Y-%m-%d"]:
        try: return datetime.strptime(s, fmt)
        except: pass
    return None

def days_overdue(date_str):
    d = parse_date(date_str)
    if not d: return None
    return (TODAY - d).days

def bucket(days):
    if days is None: return "N/A"
    if days > 90: return ">90 ngay"
    if days > 60: return "61-90 ngay"
    if days > 30: return "31-60 ngay"
    if days > 0:  return "1-30 ngay"
    if days > -7: return "Sap den han"
    return "Trong han"

BUCKET_ORDER = ["Trong han", "Sap den han", "1-30 ngay", "31-60 ngay", "61-90 ngay", ">90 ngay", "N/A"]
BUCKET_LABEL = {"Trong han":"Trong han","Sap den han":"Sap den han","1-30 ngay":"1-30 ngay","31-60 ngay":"31-60 ngay","61-90 ngay":"61-90 ngay",">90 ngay":">90 ngay","N/A":"N/A"}
BUCKET_COLOR = {"Trong han":"#22c55e","Sap den han":"#facc15","1-30 ngay":"#fb923c","31-60 ngay":"#f97316","61-90 ngay":"#ef4444",">90 ngay":"#991b1b","N/A":"#94a3b8"}

def fmt_vnd(n):
    try: return f"{float(n):,.0f}"
    except: return str(n)

def generate(json_path, output_path):
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    ft      = data.get("file_type", "")
    account = data.get("account", "")
    period  = data.get("period", "")

    rows = []
    if ft == "TAM_UNG":
        title = "Dashboard Tam Ung TK" + account + " - Ky " + period
        col_headers = ["Nhan vien","Bo phan","Ngay hoan ung","So tien (VND)","Trang thai"]
        for emp, records in data.get("by_employee", {}).items():
            for rec in records:
                days = days_overdue(rec.get("reimbursement_date"))
                b = bucket(days)
                rows.append({"cols":[emp,rec.get("dept",""),rec.get("reimbursement_date","N/A"),fmt_vnd(rec.get("amount",0)),b],"bucket":b,"amount":rec.get("amount",0)})
    elif ft == "AP_AGING":
        title = "Dashboard AP Aging TK" + account + " - Ky " + period
        col_headers = ["NCC","So HD","Due date","Tong VND","Trang thai"]
        for vendor, records in data.get("by_vendor", {}).items():
            for rec in records:
                days = days_overdue(rec.get("due_date"))
                b = bucket(days)
                rows.append({"cols":[vendor,rec.get("invoice_no",""),rec.get("due_date","N/A"),fmt_vnd(rec.get("total_vnd",0)),b],"bucket":b,"amount":rec.get("total_vnd",0)})
    elif ft == "PREPAY":
        title = "Dashboard Tra Truoc TK" + account + " - Ky " + period
        col_headers = ["NCC","Dien giai","Ngay hoan tra","So tien (VND)","Trang thai"]
        for sup, records in data.get("by_supplier", {}).items():
            for rec in records:
                days = days_overdue(rec.get("reimbursement_date"))
                b = bucket(days)
                rows.append({"cols":[sup,rec.get("description","")[:40],rec.get("reimbursement_date","N/A"),fmt_vnd(rec.get("amount",0)),b],"bucket":b,"amount":rec.get("amount",0)})
    else:
        title = "Dashboard"; col_headers = []

    summary = {b: {"count":0,"amount":0.0} for b in BUCKET_ORDER}
    for r in rows:
        b = r["bucket"]
        if b in summary:
            summary[b]["count"] += 1
            try: summary[b]["amount"] += float(r["amount"])
            except: pass

    total_amount = sum(v["amount"] for v in summary.values())
    total_count  = sum(v["count"]  for v in summary.values())

    chart_bars = ""
    for b in BUCKET_ORDER:
        sv = summary[b]
        if sv["count"] == 0: continue
        pct = (sv["amount"] / total_amount * 100) if total_amount > 0 else 0
        color = BUCKET_COLOR.get(b, "#94a3b8")
        chart_bars += (
            '<div class="bar-row">'
            '<div class="bar-label">' + BUCKET_LABEL.get(b,b) + '</div>'
            '<div class="bar-track"><div class="bar-fill" style="width:' + str(round(pct,1)) + '%;background:' + color + '"></div></div>'
            '<div class="bar-meta">' + str(sv["count"]) + ' khoan | ' + fmt_vnd(sv["amount"]) + ' VND</div>'
            '</div>'
        )

    table_rows = ""
    for r in rows:
        b = r["bucket"]
        color = BUCKET_COLOR.get(b, "#94a3b8")
        cells = "".join("<td>" + str(c) + "</td>" for c in r["cols"][:-1])
        cells += '<td><span class="badge" style="background:' + color + '">' + BUCKET_LABEL.get(b,b) + '</span></td>'
        table_rows += "<tr>" + cells + "</tr>"

    th = "".join("<th>" + h + "</th>" for h in col_headers)
    overdue_count = summary[">90 ngay"]["count"] + summary["61-90 ngay"]["count"] + summary["31-60 ngay"]["count"] + summary["1-30 ngay"]["count"]

    html_parts = [
        '<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><title>' + title + '</title>',
        '<style>*{box-sizing:border-box;margin:0;padding:0}',
        'body{font-family:"Segoe UI",sans-serif;background:#f1f5f9;color:#1e293b;padding:24px}',
        'h1{font-size:20px;font-weight:700;margin-bottom:4px}.meta{color:#64748b;font-size:13px;margin-bottom:24px}',
        '.cards{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:28px}',
        '.card{background:white;border-radius:12px;padding:18px 22px;min-width:160px;box-shadow:0 1px 6px rgba(0,0,0,.08)}',
        '.card .num{font-size:26px;font-weight:700}.card .lbl{font-size:12px;color:#64748b;margin-top:2px}',
        '.section{background:white;border-radius:12px;padding:20px 24px;margin-bottom:24px;box-shadow:0 1px 6px rgba(0,0,0,.08)}',
        '.section h2{font-size:15px;font-weight:600;margin-bottom:16px;color:#334155}',
        '.bar-row{display:flex;align-items:center;gap:12px;margin-bottom:10px}',
        '.bar-label{width:120px;font-size:13px;text-align:right;color:#475569;flex-shrink:0}',
        '.bar-track{flex:1;background:#e2e8f0;border-radius:99px;height:18px;overflow:hidden}',
        '.bar-fill{height:100%;border-radius:99px}.bar-meta{width:220px;font-size:12px;color:#64748b;flex-shrink:0}',
        'table{width:100%;border-collapse:collapse;font-size:13px}',
        'th{background:#f8fafc;color:#475569;font-weight:600;padding:9px 12px;text-align:left;border-bottom:2px solid #e2e8f0}',
        'td{padding:8px 12px;border-bottom:1px solid #f1f5f9;vertical-align:middle}',
        'tr:hover td{background:#f8fafc}.badge{display:inline-block;padding:3px 10px;border-radius:99px;color:white;font-size:11px;font-weight:600}',
        '.search{width:100%;padding:9px 14px;border:1px solid #e2e8f0;border-radius:8px;font-size:13px;margin-bottom:14px;outline:none}',
        '.search:focus{border-color:#2563eb}</style></head><body>',
        '<h1>&#x1F4CA; ' + title + '</h1>',
        '<div class="meta">Ngay tao: ' + TODAY.strftime("%d/%m/%Y") + ' | Tong: ' + str(total_count) + ' khoan | ' + fmt_vnd(total_amount) + ' VND</div>',
        '<div class="cards">',
        '<div class="card"><div class="num">' + str(total_count) + '</div><div class="lbl">Tong khoan</div></div>',
        '<div class="card"><div class="num" style="color:#ef4444">' + str(overdue_count) + '</div><div class="lbl">Qua han</div></div>',
        '<div class="card"><div class="num" style="color:#facc15">' + str(summary["Sap den han"]["count"]) + '</div><div class="lbl">Sap den han</div></div>',
        '<div class="card"><div class="num" style="color:#22c55e">' + str(summary["Trong han"]["count"]) + '</div><div class="lbl">Trong han</div></div>',
        '</div><div class="section"><h2>Phan tich Aging</h2>' + chart_bars + '</div>',
        '<div class="section"><h2>Chi tiet</h2>',
        '<input class="search" id="search" placeholder="Tim kiem..." oninput="filterTable()">',
        '<table id="tbl"><thead><tr>' + th + '</tr></thead><tbody id="tbody">' + table_rows + '</tbody></table></div>',
        '<script>function filterTable(){var q=document.getElementById("search").value.toLowerCase();document.querySelectorAll("#tbody tr").forEach(function(r){r.style.display=r.textContent.toLowerCase().indexOf(q)>=0?"":"none";});}</script>',
        '</body></html>'
    ]

    Path(output_path).write_text("".join(html_parts), encoding="utf-8")
    print("Dashboard saved: " + output_path)

if __name__ == "__main__":
    generate(sys.argv[1], sys.argv[2])
