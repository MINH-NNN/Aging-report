#!/usr/bin/env python3
"""
FastAPI wrapper cho Aging Report Agent.
AgentBase Runtime yêu cầu server lắng nghe port 8080 và có GET /health.
"""

import os
import sys
import json
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

app = FastAPI(title="Aging Report Agent", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
SCRIPTS  = Path(__file__).parent / "scripts"
BASE_DIR = Path(__file__).parent

# Đọc .env
def _load_env():
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

_load_env()


def run_script(script_name, args):
    """Chạy script con, raise HTTPException nếu lỗi (không sys.exit)."""
    result = subprocess.run(
        [sys.executable, str(SCRIPTS / script_name)] + args,
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise HTTPException(
            status_code=500,
            detail=f"[{script_name}] {result.stderr.strip() or result.stdout.strip()}"
        )
    return result.stdout


def get_period(ar_path: str) -> str:
    out = run_script("parse_ar.py", [ar_path])
    data = json.loads(out)
    report_date = data.get("report_date", "")
    match = re.search(r"(\d{2})/(\d{4})", report_date)
    if match:
        return f"T{match.group(1)}_{match.group(2)}"
    return datetime.now().strftime("T%m_%Y")


def generate_email_drafts(ar_path: str, output_path: str):
    out = run_script("parse_ar.py", [ar_path])
    data = json.loads(out)
    report_date = data["report_date"]
    by_pic = data["by_pic"]

    AGING_STATUS = {
        "Over 180 ngày": lambda d: f"Quá hạn {d} ngày 🔴🔴",
        "91-180 ngày":   lambda d: f"Quá hạn {d} ngày 🔴",
        "61-90 ngày":    lambda d: f"Quá hạn {d} ngày ⚠️",
        "31-60 ngày":    lambda d: f"Quá hạn {d} ngày",
        "1-30 ngày":     lambda d: f"Quá hạn {d} ngày",
        "Current":       lambda d: "Chưa đến hạn",
        "N/A":           lambda d: "N/A",
    }

    lines = [f"# Email Drafts AR Report — {report_date}\n"]
    for pic, records in by_pic.items():
        lines.append(f"---\n\n## PIC: {pic}\n")
        lines.append(f"**To:** {pic}")
        lines.append(f"**Subject:** [AR] Thông báo công nợ chờ thanh toán — {report_date}\n")

        table = "| # | Khách hàng | Invoice | Ngày HĐ | Số tiền (VND) | Tình trạng | Mô tả |\n"
        table += "|---|---|---|---|---|---|---|\n"
        for i, r in enumerate(records, 1):
            fn = AGING_STATUS.get(r["aging_bucket"], lambda d: r["aging_bucket"])
            status = fn(r["over_day"])
            amount = f"{r['base_amount']:,.0f}"
            inv = f"#{r['invoice_no']}" if r["invoice_no"] else "N/A"
            name = f"{r['code']} - {r['name']}" if r["name"] else str(r["code"])
            desc = (r["description"] or "").strip()
            table += f"| {i} | {name} | {inv} | {r['invoice_date'] or 'N/A'} | {amount} | {status} | {desc} |\n"

        email = f"""Kính gửi anh/chị {pic},

Bộ phận AR xin gửi thông tin về các khoản công nợ hiện đang chờ thanh toán thuộc trách nhiệm theo dõi của anh/chị, tính đến ngày **{report_date}**:

{table}
Anh/chị vui lòng:
1. Đôn đốc khách hàng thanh toán các khoản trên trong thời gian sớm nhất
2. Phản hồi lý do chưa thu được tiền (nếu có vướng mắc) để bộ phận AR cập nhật vào hệ thống

Trân trọng cảm ơn,
[Ký tên]
"""
        lines.append(email)

    Path(output_path).write_text("\n".join(lines), encoding="utf-8")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def ui():
    return """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Aging Report Agent</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',sans-serif;background:#f0f4f8;min-height:100vh;display:flex;align-items:center;justify-content:center}
  .card{background:white;border-radius:16px;box-shadow:0 4px 24px rgba(0,0,0,.10);padding:40px;width:480px;max-width:96vw}
  h1{font-size:22px;color:#1e293b;margin-bottom:6px}
  .sub{color:#64748b;font-size:14px;margin-bottom:28px}
  label{display:block;font-size:13px;font-weight:600;color:#374151;margin-bottom:6px}
  .drop{border:2px dashed #cbd5e1;border-radius:10px;padding:32px 20px;text-align:center;cursor:pointer;transition:all .2s;margin-bottom:20px;background:#f8fafc}
  .drop:hover,.drop.over{border-color:#2563eb;background:#eff6ff}
  .drop .icon{font-size:36px;margin-bottom:8px}
  .drop p{color:#64748b;font-size:14px}
  .fname{color:#2563eb;font-weight:600;margin-top:6px;font-size:14px}
  input[type=file]{display:none}
  .btn{width:100%;padding:14px;border:none;border-radius:10px;background:#2563eb;color:white;font-size:16px;font-weight:600;cursor:pointer;transition:background .2s}
  .btn:hover:not(:disabled){background:#1d4ed8}
  .btn:disabled{background:#93c5fd;cursor:not-allowed}
  .prog{display:none;margin-top:20px}
  .bar-wrap{background:#e2e8f0;border-radius:99px;height:8px;overflow:hidden;margin-bottom:8px}
  .bar{height:100%;background:#2563eb;border-radius:99px;width:0;transition:width .3s}
  .st{font-size:13px;color:#64748b;text-align:center}
  .err{color:#dc2626;font-size:13px;margin-top:12px;background:#fef2f2;padding:10px 14px;border-radius:8px;display:none}
  .ok{color:#16a34a;font-size:13px;margin-top:12px;background:#f0fdf4;padding:10px 14px;border-radius:8px;display:none}
  hr{border:none;border-top:1px solid #e2e8f0;margin:24px 0}
  .info{font-size:12px;color:#94a3b8}
</style>
</head>
<body>
<div class="card">
  <h1>📊 Aging Report Agent</h1>
  <p class="sub">Upload file AR Excel để tạo báo cáo tự động</p>
  <label>File AR Excel</label>
  <div class="drop" id="dz" onclick="document.getElementById('fi').click()">
    <div class="icon">📁</div>
    <p>Kéo thả file vào đây hoặc <strong>nhấn để chọn</strong></p>
    <div class="fname" id="fn"></div>
  </div>
  <input type="file" id="fi" accept=".xlsx,.xls">
  <button class="btn" id="btn" disabled onclick="run()">▶ &nbsp;Tạo báo cáo</button>
  <div class="prog" id="prog">
    <div class="bar-wrap"><div class="bar" id="bar"></div></div>
    <div class="st" id="st">Đang xử lý...</div>
  </div>
  <div class="err" id="err"></div>
  <div class="ok" id="ok"></div>
  <hr>
  <p class="info">Kết quả trả về file ZIP gồm: AR Master, SLA Tracker, SLA Dashboard và Email Drafts cho từng PIC.</p>
</div>
<script>
let file=null;
const dz=document.getElementById('dz');
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over')});
dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');if(e.dataTransfer.files[0])setFile(e.dataTransfer.files[0])});
document.getElementById('fi').addEventListener('change',e=>{if(e.target.files[0])setFile(e.target.files[0])});
function setFile(f){file=f;document.getElementById('fn').textContent='✅ '+f.name;document.getElementById('btn').disabled=false;hide('err');hide('ok')}
async function run(){
  if(!file)return;
  document.getElementById('btn').disabled=true;
  show('prog');hide('err');hide('ok');anim();
  const form=new FormData();
  form.append('ar_file',file,file.name);
  try{
    setSt('⏳ Đang gửi file...');
    const r=await fetch('/process',{method:'POST',body:form});
    if(!r.ok){const t=await r.text();let m='Lỗi '+r.status;try{m=JSON.parse(t).detail||m}catch{}throw new Error(m)}
    setSt('⏳ Đang tải kết quả...');
    const blob=await r.blob();
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='AgingReport_output.zip';a.click();
    setBar(100);setSt('✅ Hoàn thành!');
    document.getElementById('ok').style.display='block';
    document.getElementById('ok').textContent='✅ Đã tải xong! Kiểm tra thư mục Downloads.';
  }catch(e){setSt('');document.getElementById('err').style.display='block';document.getElementById('err').textContent='❌ '+e.message}
  finally{document.getElementById('btn').disabled=false}
}
let iv;
function anim(){let w=0;setBar(0);iv=setInterval(()=>{w=Math.min(w+Math.random()*3,90);setBar(w)},400)}
function setBar(v){clearInterval(iv);document.getElementById('bar').style.width=v+'%'}
function setSt(t){document.getElementById('st').textContent=t}
function show(id){document.getElementById(id).style.display='block'}
function hide(id){document.getElementById(id).style.display='none'}
</script>
</body>
</html>"""


@app.post("/process")
async def process(
    ar_file: UploadFile = File(..., description="File AR Excel (.xlsx)"),
    tracker_file: UploadFile = File(None, description="Tracker cũ (tuỳ chọn)"),
):
    work_dir = tempfile.mkdtemp()
    try:
        # Lưu file upload
        ar_path = os.path.join(work_dir, "ar_input.xlsx")
        with open(ar_path, "wb") as f:
            f.write(await ar_file.read())

        output_dir = Path(work_dir) / "output"
        output_dir.mkdir()

        period = get_period(ar_path)

        master_path      = str(output_dir / f"AR_Master_{period}.xlsx")
        sla_tracker_path = str(output_dir / f"SLA_Tracker_{period}.xlsx")
        dashboard_path   = str(output_dir / f"SLA_Dashboard_{period}.html")
        email_path       = str(output_dir / f"EmailDrafts_{period}.md")

        run_script("generate_master.py",      [ar_path, master_path])
        run_script("generate_sla_tracker.py", [ar_path, sla_tracker_path])
        run_script("generate_sla_dashboard.py", [ar_path, sla_tracker_path, dashboard_path])
        generate_email_drafts(ar_path, email_path)

        zip_base = os.path.join(work_dir, f"AgingReport_{period}")
        shutil.make_archive(zip_base, "zip", str(output_dir))

        return FileResponse(
            path=zip_base + ".zip",
            media_type="application/zip",
            filename=f"AgingReport_{period}.zip",
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/send")
async def send_emails_endpoint(
    ar_file: UploadFile = File(..., description="File AR Excel (.xlsx)"),
    dry_run: bool = False,
):
    """Gửi email nhắc thanh toán cho từng PIC qua Gmail."""
    pic_emails_path = BASE_DIR / "pic_emails.json"
    if not pic_emails_path.exists():
        raise HTTPException(status_code=500, detail="Thiếu file pic_emails.json")

    work_dir = tempfile.mkdtemp()
    try:
        ar_path = os.path.join(work_dir, "ar_input.xlsx")
        with open(ar_path, "wb") as f:
            f.write(await ar_file.read())

        args = [ar_path, str(pic_emails_path)]
        if dry_run:
            args.append("--dry-run")

        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "send_emails.py")] + args,
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise HTTPException(status_code=500, detail=result.stderr.strip() or result.stdout.strip())

        return {"status": "ok", "detail": result.stdout.strip()}
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
