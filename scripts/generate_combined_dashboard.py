#!/usr/bin/env python3
import sys, json
from datetime import datetime, date, timedelta
from pathlib import Path

def parse_date(s):
    if not s: return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try: return datetime.strptime(str(s).strip(), fmt).date()
        except: pass
    return None

def compute_overdue(due_str, today):
    d = parse_date(due_str)
    if d is None: return None
    return (today - d).days

def payment_sla(od):
    if od is None: return "Unknown"
    if od < 0:    return "Not Due"
    if od <= 30:  return "Within SLA"
    return "SLA Violation"

TYPE_LABEL = {
    "AR":       "AR Aging",
    "TAM_UNG":  "Advance Payment",
    "AP_AGING": "AP Aging",
    "PREPAY":   "Prepayment",
}

def normalize(sources, today):
    rows = []
    for src in sources:
        ft = src.get("file_type",""); sf = src.get("source_file",""); acc = src.get("account","")
        lbl = TYPE_LABEL.get(ft, ft)
        if ft == "AR":
            for pic, recs in src.get("by_pic", {}).items():
                real_pic = "" if pic in ("Chua phan cong","Chưa phân công") else pic
                for r in recs:
                    # AR: due_date luôn = invoice_date + 30 ngày
                    inv_date_str = r.get("invoice_date","")
                    inv_d = parse_date(inv_date_str)
                    due = (inv_d + timedelta(days=30)).strftime("%d/%m/%Y") if inv_d else inv_date_str
                    od  = compute_overdue(due, today)
                    rows.append({"loai":lbl,"pic":real_pic,
                        "entity":(r.get("name") or str(r.get("code",""))).strip(),
                        "invoice":r.get("invoice_no",""),"inv_date":r.get("invoice_date",""),
                        "due_date":due or "","desc":(r.get("description","") or "").strip(),
                        "amount":float(r.get("base_amount",0) or 0),"overdue_days":od,
                        "payment_sla":payment_sla(od),"email_sent":"","pic_resp":"Not Reminded","reminder_count":0,
                        "ghi_chu":"","source":sf,"account":acc})
        elif ft == "TAM_UNG":
            for emp, recs in src.get("by_employee", {}).items():
                for r in recs:
                    due = r.get("reimbursement_date"); od = compute_overdue(due, today)
                    rows.append({"loai":lbl,"pic":emp,"entity":emp,
                        "invoice":r.get("invoice_no",""),"inv_date":r.get("invoice_date",""),
                        "due_date":due or "","desc":(r.get("description","") or "").strip(),
                        "amount":float(r.get("amount",0) or 0),"overdue_days":od,
                        "payment_sla":payment_sla(od),"email_sent":"","pic_resp":"Not Reminded","reminder_count":0,
                        "ghi_chu":"","source":sf,"account":acc})
        elif ft == "AP_AGING":
            for vendor, recs in src.get("by_vendor", {}).items():
                for r in recs:
                    due = r.get("due_date"); od = compute_overdue(due, today)
                    if od is None:
                        if r.get("overdue_90plus",0): od=91
                        elif r.get("overdue_61_90",0): od=75
                        elif r.get("overdue_31_60",0): od=45
                        elif r.get("overdue_1_30",0):  od=15
                    rows.append({"loai":lbl,"pic":"","entity":vendor,
                        "invoice":r.get("invoice_no",""),"inv_date":r.get("invoice_date",""),
                        "due_date":due or "","desc":(r.get("description","") or "").strip(),
                        "amount":float(r.get("total_vnd",0) or 0),"overdue_days":od,
                        "payment_sla":payment_sla(od),"email_sent":"","pic_resp":"Not Reminded","reminder_count":0,
                        "ghi_chu":"","source":sf,"account":acc})
        elif ft == "PREPAY":
            for sup, recs in src.get("by_supplier", {}).items():
                for r in recs:
                    due = r.get("reimbursement_date"); od = compute_overdue(due, today)
                    rows.append({"loai":lbl,"pic":r.get("requester",""),"entity":sup,
                        "invoice":r.get("invoice_no",""),"inv_date":r.get("gl_date",""),
                        "due_date":due or "","desc":(r.get("description","") or "").strip(),
                        "amount":float(r.get("amount",0) or 0),"overdue_days":od,
                        "payment_sla":payment_sla(od),"email_sent":"","pic_resp":"Not Reminded","reminder_count":0,
                        "ghi_chu":"","source":sf,"account":acc})
    # Chỉ giữ lại khoản quá hạn (overdue_days >= 1)
    return [r for r in rows if r.get("overdue_days") is not None and r["overdue_days"] >= 1]

COLS = ["Loai","PIC","Doi tac / NV","Invoice","Ngay HD","Due date","Mo ta",
        "So tien (VND)","Qua han","Payment SLA","Ngay gui mail","PIC Response SLA","Ghi chu","Nguon file"]


def build_html(records_json, report_date, today_str):
    cols = ["Type","PIC","Entity / Staff","Invoice","Inv. Date","Due Date","Description",
            "Amount (VND)","Overdue (days)","Reminded","Email Sent","PIC Response","Notes","Source File"]
    ths = "".join(
        '<th onclick="sortBy(%d)">%s<span id="s%d"></span></th>' % (i, c, i)
        for i, c in enumerate(cols)
    )
    html = r"""<!DOCTYPE html>
<html lang="vi"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard Cong No Tong Hop</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Segoe UI",Arial,sans-serif;background:#f0f4f8;font-size:13px}
.hdr{background:#1e3a5f;color:white;padding:14px 24px;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
.hdr-left h1{font-size:19px;font-weight:700;margin:0}
.hdr-left .sub{font-size:11px;color:#93c5fd;margin-top:2px}
.hdr-right{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.hdr-meta{text-align:right;font-size:11px;color:#93c5fd;line-height:1.7;margin-right:12px}
.btn-hdr{border:none;border-radius:7px;padding:7px 16px;font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px;white-space:nowrap}
.btn-imp{background:#3b82f6;color:white}.btn-imp:hover{background:#2563eb}
.btn-exp{background:#10b981;color:white}.btn-exp:hover{background:#059669}
.body{padding:18px 24px}
.sec-title{display:flex;align-items:center;gap:8px;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px}
.sec-title::before{content:"";display:inline-block;width:3px;height:13px;background:#2563eb;border-radius:2px}
.cards{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}
.card{background:white;border-radius:10px;border:1.5px solid #e2e8f0;padding:12px 16px;min-width:140px;flex:1}
.card .num{font-size:26px;font-weight:700;color:#111;line-height:1.1}
.card .lbl{font-size:11px;color:#6b7280;margin-top:3px}
.card.green .num,.card.yellow .num,.card.red .num,.card.gray .num{color:var(--color-text-primary,#111)}
.bc{background:white;border-radius:10px;border:1.5px solid #e2e8f0;padding:14px 20px;margin-bottom:10px;display:flex;align-items:center;gap:28px;flex-wrap:wrap}
.bc .bc-blk{display:flex;flex-direction:column;gap:3px}
.bc .bc-lbl{font-size:11px;color:#6b7280}
.bc .bc-row{display:flex;align-items:baseline;gap:6px}
.bc .bc-num{font-size:34px;font-weight:700;color:#111;line-height:1}
.bc .bc-unit{font-size:13px;color:#6b7280}
.bc .bc-amt{font-size:26px;font-weight:700;color:#111;line-height:1}
.bc .bc-sep{width:1px;height:40px;background:#e2e8f0;flex-shrink:0}
.sc{border-radius:10px;padding:10px 12px;min-width:100px;flex:1}
.sc .sc-lbl{font-size:10px;margin:0 0 4px}
.sc .sc-top{display:flex;align-items:baseline;gap:4px;margin-bottom:5px}
.sc .sc-num{font-size:22px;font-weight:700;line-height:1}
.sc .sc-unit{font-size:10px;font-weight:600}
.sc .sc-div{border:none;border-top:1px solid;margin:0 0 5px}
.sc .sc-amt{font-size:12px;font-weight:700;margin:0}
.sc .sc-sub{font-size:10px;margin:2px 0 0}
.sc.r1{background:#fff5f5;border:1.5px solid #fecaca}.sc.r1 .sc-lbl,.sc.r1 .sc-sub{color:#fca5a5}.sc.r1 .sc-num,.sc.r1 .sc-unit,.sc.r1 .sc-amt{color:#ef4444}.sc.r1 .sc-div{border-color:#fecaca}
.sc.r2{background:#fff0f0;border:1.5px solid #fca5a5}.sc.r2 .sc-lbl,.sc.r2 .sc-sub{color:#f87171}.sc.r2 .sc-num,.sc.r2 .sc-unit,.sc.r2 .sc-amt{color:#dc2626}.sc.r2 .sc-div{border-color:#fca5a5}
.sc.r3{background:#fee8e8;border:1.5px solid #f87171}.sc.r3 .sc-lbl,.sc.r3 .sc-sub{color:#ef4444}.sc.r3 .sc-num,.sc.r3 .sc-unit,.sc.r3 .sc-amt{color:#b91c1c}.sc.r3 .sc-div{border-color:#f87171}
.sc.r4{background:#fdd8d8;border:1.5px solid #ef4444}.sc.r4 .sc-lbl,.sc.r4 .sc-sub{color:#dc2626}.sc.r4 .sc-num,.sc.r4 .sc-unit,.sc.r4 .sc-amt{color:#991b1b}.sc.r4 .sc-div{border-color:#ef4444}
.sc.r5{background:#fbc2c2;border:1.5px solid #dc2626}.sc.r5 .sc-lbl,.sc.r5 .sc-sub{color:#b91c1c}.sc.r5 .sc-num,.sc.r5 .sc-unit,.sc.r5 .sc-amt{color:#7f1d1d}.sc.r5 .sc-div{border-color:#dc2626}
.filters{background:white;border-radius:10px;border:1px solid #e2e8f0;padding:10px 14px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:10px}
.filters label{font-size:11px;color:#6b7280;font-weight:600}
.filters select,.filters input{border:1px solid #d1d5db;border-radius:6px;padding:4px 9px;font-size:12px;color:#374151;background:white}
.btn-r{background:#f1f5f9;border:1px solid #d1d5db;border-radius:6px;padding:4px 12px;font-size:12px;cursor:pointer}
.btn-r:hover{background:#e2e8f0}
.count{font-size:12px;color:#6b7280;margin-bottom:6px}
.tbl-wrap{background:white;border-radius:10px;border:1px solid #e2e8f0;overflow:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#1e3a5f;color:white;padding:9px 11px;text-align:left;white-space:nowrap;cursor:pointer;user-select:none}
th:hover{background:#1e40af}
td{padding:7px 11px;border-bottom:1px solid #f1f5f9;vertical-align:top}
tr:hover td{background:#f8fafc}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}
.b-ar{background:#dcfce7;color:#166534}
.b-tu{background:#fef9c3;color:#854d0e}
.b-ap{background:#ede9fe;color:#5b21b6}
.b-pp{background:#fee2e2;color:#991b1b}
.ok{color:#16a34a;font-weight:600}.warn{color:#b45309;font-weight:600}.bad{color:#dc2626;font-weight:600}.unk{color:#9ca3af}
.editable{cursor:pointer;border-bottom:1px dashed #cbd5e1;display:inline-block;min-width:40px}
.editable:hover{border-bottom-color:#2563eb;color:#2563eb}
/* Modal */
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;align-items:center;justify-content:center}
.overlay.show{display:flex}
.modal{background:white;border-radius:14px;width:620px;max-width:95vw;max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.25)}
.modal-hdr{padding:16px 22px;border-bottom:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center}
.modal-hdr h2{font-size:16px;color:#1e293b;margin:0}
.modal-close{background:none;border:none;font-size:20px;cursor:pointer;color:#9ca3af;line-height:1}
.modal-close:hover{color:#374151}
.modal-body{padding:20px 22px}
.modal-foot{padding:12px 22px;border-top:1px solid #e2e8f0;display:flex;justify-content:flex-end;gap:8px}
.drop-zone{border:2px dashed #cbd5e1;border-radius:10px;padding:32px;text-align:center;cursor:pointer;transition:.2s}
.drop-zone:hover,.drop-zone.over{border-color:#2563eb;background:#eff6ff}
.drop-zone .dz-icon{font-size:32px;margin-bottom:8px}
.drop-zone p{color:#6b7280;font-size:13px}
.file-list{margin-top:12px;display:flex;flex-wrap:wrap;gap:6px}
.ftag{display:inline-flex;align-items:center;gap:5px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:16px;padding:3px 10px;font-size:12px;color:#1d4ed8}
.imp-status{margin-top:14px;font-size:13px;min-height:20px}
.imp-status.ok{color:#16a34a}.imp-status.err{color:#dc2626}.imp-status.info{color:#6b7280}
.prog-bar{height:6px;background:#e2e8f0;border-radius:4px;margin-top:8px;overflow:hidden;display:none}
.prog-fill{height:100%;background:#2563eb;border-radius:4px;width:0;transition:width .3s}
.btn-main{background:#2563eb;color:white;border:none;border-radius:7px;padding:7px 18px;font-size:13px;font-weight:600;cursor:pointer}
.btn-main:hover{background:#1d4ed8}
.btn-main:disabled{background:#93c5fd;cursor:not-allowed}
.btn-sec{background:#f1f5f9;border:1px solid #d1d5db;border-radius:7px;padding:7px 18px;font-size:13px;cursor:pointer}
/* Toast */
.toast{position:fixed;bottom:24px;right:24px;padding:10px 18px;border-radius:8px;font-size:13px;font-weight:600;z-index:200;box-shadow:0 4px 12px rgba(0,0,0,.2);opacity:0;transition:opacity .3s;pointer-events:none}
.toast.show{opacity:1}
.toast.success{background:#16a34a;color:white}
.toast.error{background:#dc2626;color:white}
</style></head><body>
<div class="hdr">
  <div class="hdr-left">
    <h1>&#128202; Aging Dashboard</h1>
    
  </div>
  <div class="hdr-right">
    <div class="hdr-meta">
      <div>Report date: <span id="rdate">REPORT_DATE_PH</span></div>
      <div>Viewed: TODAY_PH</div>
    </div>
    <button class="btn-hdr btn-imp" onclick="openImport()">&#8673; Import</button>
    <button class="btn-hdr btn-exp" onclick="exportExcel()">&#8675; Export</button>
    <button class="btn-hdr" style="background:#64748b;color:white" onclick="clearData()">&#128465; Clear Data</button>
  </div>
</div>
<div class="body">
<script>
var ROWS=ROWS_PH;
(function(){try{var s=localStorage.getItem('cnresp');if(s){var m=JSON.parse(s);ROWS.forEach(function(r){var k=r.invoice||(r.entity+'|'+r.inv_date);if(m[k]){r.email_sent=m[k].e||'';r.pic_resp=m[k].p||r.pic_resp;r.ghi_chu=m[k].g||'';}});}}catch(e){}})();
</script>
<div class="sec-title" style="font-size:13px">PAYMENT SLA</div>
<div id="pc"></div>
<div class="sec-title" style="font-size:13px">PIC RESPONSE SLA</div>
<div class="cards" id="rc"></div>
<div class="filters">
  <span><label>Type:&nbsp;</label><select id="f-loai"><option value="">All</option></select></span>
  <span><label>PIC:&nbsp;</label><select id="f-pic"><option value="">All</option></select></span>
  <span><label>Reminded:&nbsp;</label><select id="f-psla"><option value="">All</option><option value="0">0/3</option><option value="1">1/3</option><option value="2">2/3</option><option value="3">3/3</option></select></span>
  <span><label>PIC Response:&nbsp;</label><select id="f-resp"><option value="">All</option></select></span>
  <span><label>Search:&nbsp;</label><input id="f-q" placeholder="Invoice, entity, description..." style="width:200px"></span>
  <button class="btn-r" onclick="doReset()">&#8635; Reset</button>
</div>
<div style="display:flex;align-items:center;justify-content:space-between;margin:8px 0 4px">
  <div class="count" id="cnt"></div>
  <div style="display:flex;align-items:center;gap:6px;font-size:12px;color:#6b7280">
    <label>Rows/page:</label>
    <select id="pg-size" onchange="setPageSize()" style="font-size:12px;padding:2px 6px;border:1px solid #d1d5db;border-radius:4px">
      <option value="10">10</option>
      <option value="50" selected>50</option>
      <option value="100">100</option>
      <option value="0">All</option>
    </select>
  </div>
</div>
<div class="tbl-wrap"><table><thead><tr>THS_PH</tr></thead><tbody id="tb"></tbody></table></div>
<div id="pager" style="display:flex;align-items:center;justify-content:center;gap:8px;margin:12px 0;font-size:13px"></div>
</div>

<!-- IMPORT MODAL -->
<div class="overlay" id="imp-ov" onclick="if(event.target===this)closeImport()">
<div class="modal">
  <div class="modal-hdr">
    <h2>&#8673; Import Debt Reports</h2>
    <button class="modal-close" onclick="closeImport()">&#x2715;</button>
  </div>
  <div class="modal-body">
    <p style="color:#6b7280;font-size:13px;margin-bottom:14px">Upload Excel debt report files (AR, AP Aging, Advance Payment, Prepayment). Dashboard will update automatically.</p>
    <div class="drop-zone" id="imp-dz" onclick="document.getElementById('imp-fi').click()">
      <div class="dz-icon">&#128196;</div>
      <p>Drag &amp; drop files or <strong>click to select</strong></p>
      <p style="font-size:11px;color:#9ca3af;margin-top:4px">Accepts: .xlsx, .xls</p>
      <input type="file" id="imp-fi" accept=".xlsx,.xls" multiple style="display:none">
    </div>
    <div class="file-list" id="imp-flist"></div>
    <div class="prog-bar" id="imp-pb"><div class="prog-fill" id="imp-pf"></div></div>
    <div class="imp-status info" id="imp-st"></div>
    <p style="font-size:11px;color:#9ca3af;margin-top:10px">&#9432; Requires server running at localhost:8080</p>
  </div>
  <div class="modal-foot">
    <button class="btn-sec" onclick="closeImport()">Close</button>
    <button class="btn-main" id="imp-btn" disabled onclick="doImport()">Process &amp; Update</button>
  </div>
</div>
</div>

<!-- TOAST -->
<div class="toast" id="toast"></div>

<script>
var sortCol=-1,sortDir=1,filtered=[],impFiles=[];
var LC={'AR Aging':'b-ar','Advance Payment':'b-tu','AP Aging':'b-ap','Prepayment':'b-pp'};
var SC={'Not Due':'ok','Within SLA':'warn','SLA Violation':'bad','Unknown':'unk'};
var RC={'Responded':'ok','Awaiting':'warn','No Response':'bad','Not Reminded':'unk'};

function fmt(n){return(n===null||n===undefined||n==='')?' ':Number(n).toLocaleString('vi-VN');}
function odCls(d){if(d===null||d===undefined)return'unk';if(d<0)return'ok';if(d<=30)return'warn';return'bad';}

function popSel(id,key){
  var s=document.getElementById(id);
  Array.from(new Set(ROWS.map(function(r){return r[key];}).filter(Boolean))).sort().forEach(function(v){
    var o=document.createElement('option');o.value=v;o.textContent=v;s.appendChild(o);
  });
}
function resetSels(){
  ['f-loai','f-pic','f-psla','f-resp'].forEach(function(id){
    var s=document.getElementById(id);
    while(s.options.length>1)s.remove(1);
  });
  popSel('f-loai','loai');popSel('f-pic','pic');popSel('f-resp','pic_resp');
}

function getFiltered(){
  var fl=document.getElementById('f-loai').value,fp=document.getElementById('f-pic').value,
      fs=document.getElementById('f-psla').value,fr=document.getElementById('f-resp').value,
      fq=document.getElementById('f-q').value.toLowerCase();
  return ROWS.filter(function(r){
    if(fl&&r.loai!==fl)return false;
    if(fp&&r.pic!==fp)return false;
    if(fs&&String(r.reminder_count)!==fs)return false;
    if(fr&&r.pic_resp!==fr)return false;
    if(fq&&!(r.entity+r.invoice+r.desc+r.pic+r.source+r.ghi_chu).toLowerCase().includes(fq))return false;
    return true;
  });
}

function card(n,lbl,cls,icon){return '<div class="card '+cls+'"><div class="num">'+n+'</div><div class="lbl">'+icon+' '+lbl+'</div></div>';}

function sc(n,v,lbl,cls){
  return '<div class="sc '+cls+'">'
    +'<div class="sc-lbl">'+lbl+'</div>'
    +'<div class="sc-top"><span class="sc-num">'+n+'</span><span class="sc-unit">items</span></div>'
    +'<hr class="sc-div">'
    +'<div class="sc-amt">'+fmt(v)+'</div>'
    +'<div class="sc-sub">VND</div>'
    +'</div>';
}
function renderCards(rows){
  var s0=0,v0=0, s1=0,v1=0, s2=0,v2=0, s3=0,v3=0, s4=0,v4=0, s5=0,v5=0;
  var ro=0,rw=0,rb=0,rn=0;
  rows.forEach(function(r){
    var od=r.overdue_days;
    if(od!==null&&od!==undefined&&od>=1){
      s0++;v0+=r.amount;
      if(od<=30){s1++;v1+=r.amount;}
      else if(od<=60){s2++;v2+=r.amount;}
      else if(od<=90){s3++;v3+=r.amount;}
      else if(od<=180){s4++;v4+=r.amount;}
      else{s5++;v5+=r.amount;}
    }
    var p=r.pic_resp;
    if(p==='Da phan hoi dung han')ro++;
    else if(p==='Dang cho phan hoi')rw++;
    else if(p==='Vi pham / Khong phan hoi')rb++;
    else rn++;
  });
  document.getElementById('pc').innerHTML=
    '<div class="bc">'
      +'<div class="bc-blk"><div class="bc-lbl">Total Overdue</div><div class="bc-row"><span class="bc-num">'+s0+'</span><span class="bc-unit">items</span></div></div>'
      +'<div class="bc-sep"></div>'
      +'<div class="bc-blk"><div class="bc-lbl">Total Overdue Amount</div><div class="bc-row"><span class="bc-amt">'+fmt(v0)+'</span><span class="bc-unit">VND</span></div></div>'
    +'</div>'
    +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
      +sc(s1,v1,'1 – 30 days','r1')
      +sc(s2,v2,'31 – 60 days','r2')
      +sc(s3,v3,'61 – 90 days','r3')
      +sc(s4,v4,'90 – 180 days','r4')
      +sc(s5,v5,'Over 180 days','r5')
    +'</div>';
  document.getElementById('rc').innerHTML=
    card(ro,'Responded','green','&#10003;')
    +card(rw,'Awaiting','yellow','&#8987;')
    +card(rb,'No Response','red','&#9679;')
    +card(rn,'Not Reminded','gray','&mdash;');
}

var currentPage=1, pageSize=50;
function setPageSize(){pageSize=parseInt(document.getElementById('pg-size').value)||0;currentPage=1;renderTable(filtered);}
function goPage(p){currentPage=p;renderTable(filtered);}

function renderTable(rows){
  var tb=document.getElementById('tb');
  if(!rows.length){
    tb.innerHTML='<tr><td colspan="14" style="text-align:center;color:#9ca3af;padding:24px">Khong co du lieu</td></tr>';
    document.getElementById('cnt').textContent='Showing 0 / '+ROWS.length+' items';
    document.getElementById('pager').innerHTML='';
    return;
  }
  var total=rows.length;
  var ps=pageSize>0?pageSize:total;
  var totalPages=Math.ceil(total/ps);
  if(currentPage>totalPages)currentPage=totalPages;
  var start=(currentPage-1)*ps, end=Math.min(start+ps,total);
  var pageRows=rows.slice(start,end);

  var h='';
  pageRows.forEach(function(r){
    var lc=LC[r.loai]||'',rc2=RC[r.pic_resp]||'';
    var od=(r.overdue_days!==null&&r.overdue_days!==undefined)?(r.overdue_days+' ng'):'N/A';
    var ri=ROWS.indexOf(r);
    h+='<tr>';
    h+='<td><span class="badge '+lc+'">'+r.loai+'</span></td>';
    h+='<td>'+r.pic+'</td>';
    h+='<td>'+r.entity+'</td>';
    h+='<td>'+r.invoice+'</td>';
    h+='<td style="white-space:nowrap">'+r.inv_date+'</td>';
    h+='<td style="white-space:nowrap">'+r.due_date+'</td>';
    h+='<td style="max-width:180px">'+r.desc+'</td>';
    h+='<td style="text-align:right">'+fmt(r.amount)+'</td>';
    h+='<td style="text-align:right">'+od+'</td>';
    var rc_str=(r.reminder_count||0)+'/3';
    var rc_cls=r.reminder_count>=3?'bad':r.reminder_count>=2?'warn':r.reminder_count>=1?'':'unk';
    h+='<td class="'+rc_cls+'">'+rc_str+'</td>';
    h+='<td><span class="editable" title="Click to edit" onclick="editCell('+ri+',\'email_sent\',this)">'+(r.email_sent||'—')+'</span></td>';
    h+='<td class="'+rc2+'"><span class="editable" title="Click to edit" onclick="editPicResp('+ri+',this)">'+r.pic_resp+'</span></td>';
    h+='<td><span class="editable" title="Click to edit" onclick="editCell('+ri+',\'ghi_chu\',this)">'+(r.ghi_chu||'—')+'</span></td>';
    h+='<td style="font-size:11px;color:#9ca3af">'+r.source+'</td>';
    h+='</tr>';
  });
  tb.innerHTML=h;
  document.getElementById('cnt').textContent='Showing '+(start+1)+'–'+end+' of '+total+' items (total '+ROWS.length+')';

  /* pager */
  var pg=document.getElementById('pager');
  if(totalPages<=1){pg.innerHTML='';return;}
  var btnStyle='style="padding:3px 10px;border:1px solid #d1d5db;border-radius:4px;cursor:pointer;font-size:12px;background:#fff"';
  var activStyle='style="padding:3px 10px;border:1px solid #2563eb;border-radius:4px;cursor:pointer;font-size:12px;background:#2563eb;color:#fff"';
  var ph='';
  ph+='<button '+btnStyle+' onclick="goPage('+Math.max(1,currentPage-1)+')" '+(currentPage===1?'disabled':'')+'>&#8249; Prev</button>';
  var from=Math.max(1,currentPage-2), to=Math.min(totalPages,currentPage+2);
  if(from>1)ph+='<button '+btnStyle+' onclick="goPage(1)">1</button>'+(from>2?'<span style="padding:0 4px">…</span>':'');
  for(var p=from;p<=to;p++){ph+='<button '+(p===currentPage?activStyle:btnStyle)+' onclick="goPage('+p+')">'+p+'</button>';}
  if(to<totalPages)ph+=(to<totalPages-1?'<span style="padding:0 4px">…</span>':'')+'<button '+btnStyle+' onclick="goPage('+totalPages+')">'+totalPages+'</button>';
  ph+='<button '+btnStyle+' onclick="goPage('+Math.min(totalPages,currentPage+1)+')" '+(currentPage===totalPages?'disabled':'')+'>Next &#8250;</button>';
  pg.innerHTML=ph;
}

/* inline edit */
function editCell(ri,field,el){
  var cur=ROWS[ri][field]||'';
  var inp=document.createElement('input');
  inp.value=cur;inp.style.cssText='border:1px solid #2563eb;border-radius:4px;padding:1px 6px;font-size:12px;width:120px';
  el.replaceWith(inp);inp.focus();
  var saved=false;
  function save(){if(saved)return;saved=true;ROWS[ri][field]=inp.value;saveLocal();refresh();}
  inp.addEventListener('blur',save);
  inp.addEventListener('keydown',function(e){if(e.key==='Enter')inp.blur();if(e.key==='Escape'){saved=true;refresh();}});
}
var ROPTS=['Not Reminded','Awaiting','Responded','No Response'];
function editPicResp(ri,el){
  var sel=document.createElement('select');
  sel.style.cssText='border:1px solid #2563eb;border-radius:4px;padding:2px 4px;font-size:12px';
  ROPTS.forEach(function(v){var o=document.createElement('option');o.value=v;o.textContent=v;if(v===ROWS[ri].pic_resp)o.selected=true;sel.appendChild(o);});
  el.replaceWith(sel);sel.focus();
  var saved=false;
  function save(){if(saved)return;saved=true;ROWS[ri].pic_resp=sel.value;saveLocal();refresh();}
  sel.addEventListener('change',save);sel.addEventListener('blur',save);
}
function saveLocal(){
  try{var m={};ROWS.forEach(function(r){if(r.email_sent||r.pic_resp!=='Chua gui mail'||r.ghi_chu){var k=r.invoice||(r.entity+'|'+r.inv_date);m[k]={e:r.email_sent,p:r.pic_resp,g:r.ghi_chu};}});localStorage.setItem('cnresp',JSON.stringify(m));}catch(e){}
}

/* sort */
function sortBy(col){
  if(sortCol===col)sortDir*=-1;else{sortCol=col;sortDir=1;}
  for(var i=0;i<14;i++){var el=document.getElementById('s'+i);if(el)el.textContent='';}
  var el2=document.getElementById('s'+col);if(el2)el2.textContent=sortDir===1?' ↑':' ↓';
  var keys=['loai','pic','entity','invoice','inv_date','due_date','desc','amount','overdue_days','payment_sla','email_sent','pic_resp','ghi_chu','source'];
  filtered.sort(function(a,b){var av=a[keys[col]],bv=b[keys[col]];if(av===null||av===undefined)av='';if(bv===null||bv===undefined)bv='';if(typeof av==='number'&&typeof bv==='number')return(av-bv)*sortDir;return String(av).localeCompare(String(bv),'vi')*sortDir;});
  renderTable(filtered);
}

function doReset(){['f-loai','f-pic','f-psla','f-resp'].forEach(function(id){document.getElementById(id).value='';});document.getElementById('f-q').value='';refresh();}
function refresh(){currentPage=1;filtered=getFiltered();renderCards(filtered);renderTable(filtered);}

/* ---- EXPORT ---- */
function exportExcel(){
  var rows=filtered.length?filtered:ROWS;
  var hdr=['Loai','PIC','Doi tac / NV','Invoice','Ngay HD','Due date','Mo ta','So tien (VND)','Qua han (ngay)','Payment SLA','Ngay gui mail','PIC Response','Ghi chu','Nguon file'];
  var data=[hdr];
  rows.forEach(function(r){
    data.push([r.loai,r.pic,r.entity,r.invoice,r.inv_date,r.due_date,r.desc,r.amount,
               r.overdue_days===null?'':r.overdue_days,r.payment_sla,r.email_sent,r.pic_resp,r.ghi_chu,r.source]);
  });
  var ws=XLSX.utils.aoa_to_sheet(data);
  ws['!cols']=[{wch:14},{wch:16},{wch:28},{wch:18},{wch:12},{wch:12},{wch:30},{wch:16},{wch:10},{wch:20},{wch:14},{wch:24},{wch:20},{wch:22}];
  var wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'CongNo');
  XLSX.writeFile(wb,'TrangThai_CongNo_'+new Date().toISOString().slice(0,10)+'.xlsx');
}

/* ---- IMPORT ---- */
function openImport(){document.getElementById('imp-ov').classList.add('show');impFiles=[];renderImpList();document.getElementById('imp-st').textContent='';document.getElementById('imp-btn').disabled=true;}
function closeImport(){document.getElementById('imp-ov').classList.remove('show');}

var dz=document.getElementById('imp-dz');
dz.addEventListener('dragover',function(e){e.preventDefault();dz.classList.add('over')});
dz.addEventListener('dragleave',function(){dz.classList.remove('over')});
dz.addEventListener('drop',function(e){e.preventDefault();dz.classList.remove('over');addImpFiles(e.dataTransfer.files)});
document.getElementById('imp-fi').addEventListener('change',function(e){addImpFiles(e.target.files);e.target.value='';});

function addImpFiles(fl){
  for(var i=0;i<fl.length;i++){if(!impFiles.find(function(x){return x.name===fl[i].name;}))impFiles.push(fl[i]);}
  renderImpList();
  document.getElementById('imp-btn').disabled=impFiles.length===0;
  document.getElementById('imp-st').textContent='';
}
function renderImpList(){
  document.getElementById('imp-flist').innerHTML=impFiles.map(function(f,i){
    return '<span class="ftag">'+f.name+' <span style="cursor:pointer" onclick="removeImpFile('+i+')">&#x2715;</span></span>';
  }).join('');
}
function removeImpFile(i){impFiles.splice(i,1);renderImpList();document.getElementById('imp-btn').disabled=impFiles.length===0;}

async function doImport(){
  if(!impFiles.length)return;
  var st=document.getElementById('imp-st'),pb=document.getElementById('imp-pb'),pf=document.getElementById('imp-pf');
  var btn=document.getElementById('imp-btn');
  btn.disabled=true;pb.style.display='block';pf.style.width='20%';
  st.className='imp-status info';st.textContent='Uploading to server...';

  // save current response status to re-apply after update
  var respMap={};
  ROWS.forEach(function(r){
    var k=r.invoice||(r.entity+'|'+r.inv_date);
    if(r.email_sent||r.pic_resp!=='Chua gui mail'||r.ghi_chu)
      respMap[k]={e:r.email_sent,p:r.pic_resp,g:r.ghi_chu};
  });

  var form=new FormData();
  for(var i=0;i<impFiles.length;i++)form.append('files',impFiles[i],impFiles[i].name);

  try{
    pf.style.width='50%';
    var resp=await fetch('http://localhost:8080/api/process',{method:'POST',body:form});
    pf.style.width='80%';
    if(!resp.ok){var err=await resp.json().catch(function(){return{detail:'Server error'};});throw new Error(err.detail||'Server error');}
    var data=await resp.json();
    pf.style.width='100%';
    // merge new rows, preserve response status
    ROWS=data.rows||[];
    resetSels();
    ROWS.forEach(function(r){
      var k=r.invoice||(r.entity+'|'+r.inv_date);
      if(respMap[k]){r.email_sent=respMap[k].e||'';r.pic_resp=respMap[k].p||r.pic_resp;r.ghi_chu=respMap[k].g||'';}
    });
    saveLocal();refresh();
    st.className='imp-status ok';st.textContent='Updated '+ROWS.length+' items.';
    setTimeout(closeImport,1200);
  }catch(e){
    st.className='imp-status err';st.textContent='Error: '+e.message;
    btn.disabled=false;
  }
}

async function clearData(){
  if(!confirm('Clear all dashboard data? Reminder history will be preserved.'))return;
  try{await fetch('http://localhost:8080/api/clear',{method:'POST'});}catch(e){}
  ROWS=[];refresh();showToast('Dashboard cleared','success');
}

document.addEventListener('DOMContentLoaded',function(){
  resetSels();
  ['f-loai','f-pic','f-psla','f-resp'].forEach(function(id){document.getElementById(id).addEventListener('change',refresh);});
  document.getElementById('f-q').addEventListener('input',refresh);
  refresh();
});
</script>
</body></html>"""
    return html \
        .replace("ROWS_PH", records_json) \
        .replace("REPORT_DATE_PH", report_date) \
        .replace("TODAY_PH", today_str) \
        .replace("THS_PH", ths)

def get_rows_json(json_path):
    """Return (rows_list, report_date_str) — used by /api/process."""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    today = date.today()
    sources = data.get("sources", [])
    if not sources and "file_type" in data:
        sources = [data]
    rows = normalize(sources, today)
    rdates = [s.get("report_date","") for s in sources if s.get("report_date","")]
    report_date = ", ".join(sorted(set(rdates))) if rdates else today.strftime("%d/%m/%Y")
    return rows, report_date


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: generate_combined_dashboard.py <json_path> [<output_html>|--json]", file=sys.stderr)
        sys.exit(1)

    json_path = sys.argv[1]

    if len(sys.argv) >= 3 and sys.argv[2] == "--json":
        rows, report_date = get_rows_json(json_path)
        print(json.dumps({"rows": rows, "report_date": report_date}, ensure_ascii=False))
    else:
        out_path = sys.argv[2] if len(sys.argv) >= 3 else "dashboard.html"
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        today = date.today()
        sources = data.get("sources", [])
        if not sources and "file_type" in data:
            sources = [data]
        rdates = [s.get("report_date","") for s in sources if s.get("report_date","")]
        report_date = ", ".join(sorted(set(rdates))) if rdates else today.strftime("%d/%m/%Y")
        today_str = datetime.now().strftime("%d/%m/%Y %H:%M")
        rows = normalize(sources, today)
        records_json = json.dumps(rows, ensure_ascii=False)
        html = build_html(records_json, report_date, today_str)
        Path(out_path).write_text(html, encoding="utf-8")
        print(f"Dashboard saved: {out_path}", file=sys.stderr)
