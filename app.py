#!/usr/bin/env python3
"""FastAPI wrapper cho Aging Report Agent — output: 4 combined files."""

import os, sys, json, re, shutil, subprocess, tempfile
from datetime import datetime
from pathlib import Path
from typing import List
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

app = FastAPI(title="Aging Report Agent", version="2.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

SCRIPTS  = Path(__file__).parent / "scripts"
BASE_DIR = Path(__file__).parent

# ── Đọc .env ──────────────────────────────────────────────────────────────────
def _load_env():
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

_load_env()

# ── Tracker (lưu lịch sử nhắc nợ) ────────────────────────────────────────────
_last_output_dir: str = ""
_dashboard_html: str = ""
TRACKER_PATH = BASE_DIR / "debt_tracker.json"

def load_tracker() -> dict:
    if TRACKER_PATH.exists():
        try: return json.loads(TRACKER_PATH.read_text(encoding="utf-8"))
        except: pass
    return {}

def save_tracker(tracker: dict):
    TRACKER_PATH.write_text(json.dumps(tracker, ensure_ascii=False, indent=2), encoding="utf-8")

def tracker_key(row: dict) -> str:
    inv = str(row.get("invoice") or "").strip()
    ent = str(row.get("entity") or "").strip()
    return f"{inv}|{ent}" if inv else f"{ent}|{row.get('inv_date','')}"

def merge_tracker(rows: list) -> list:
    tracker = load_tracker()
    new_tracker = {}
    for row in rows:
        key = tracker_key(row)
        rec = tracker.get(key, {"reminder_count": 0, "history": []})
        row["reminder_count"] = rec.get("reminder_count", 0)
        row["reminder_history"] = rec.get("history", [])
        new_tracker[key] = {
            "entity": row.get("entity",""),
            "invoice": row.get("invoice",""),
            "reminder_count": row["reminder_count"],
            "history": row["reminder_history"],
        }
    save_tracker(new_tracker)
    return rows

# ── Dashboard ─────────────────────────────────────────────────────────────────
import sys as _sys
_sys.path.insert(0, str(BASE_DIR / "scripts"))
from generate_combined_dashboard import build_html as _build_html
from datetime import datetime as _dt

def _empty_dashboard():
    today = _dt.now().strftime("%d/%m/%Y %H:%M")
    return _build_html("[]", "Chua co du lieu", today)


# ── Helper chạy script con ────────────────────────────────────────────────────
def run_script(script_name, args):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        [sys.executable, str(SCRIPTS / script_name)] + args,
        capture_output=True, text=True, encoding="utf-8", env=env
    )
    if result.returncode != 0:
        raise HTTPException(status_code=500,
            detail=f"[{script_name}] {result.stderr.strip() or result.stdout.strip()}")
    return result.stdout


# ── Email draft helpers ───────────────────────────────────────────────────────
def _generate_ar_email_lines(data: dict) -> list:
    report_date = data.get("report_date", "")
    by_pic = data.get("by_pic", {})
    NO_PIC_KEYS = {"Chua phan cong", "Chưa phân công"}
    pic_records = {k: v for k, v in by_pic.items() if k not in NO_PIC_KEYS}
    no_pic_records = next((v for k, v in by_pic.items() if k in NO_PIC_KEYS), [])
    lines = [f"\n---\n## AR — {report_date}\n"]
    AGING = {
        "Over 180 ngày": lambda d: f"Quá hạn {d} ngày 🔴🔴",
        "91-180 ngày":   lambda d: f"Quá hạn {d} ngày 🔴",
        "61-90 ngày":    lambda d: f"Quá hạn {d} ngày ⚠️",
        "31-60 ngày":    lambda d: f"Quá hạn {d} ngày",
        "1-30 ngày":     lambda d: f"Quá hạn {d} ngày",
        "Current":       lambda d: "Chưa đến hạn",
        "N/A":           lambda d: "N/A",
    }
    for pic, records in pic_records.items():
        rows = ""
        for i, r in enumerate(records, 1):
            fn   = AGING.get(r["aging_bucket"], lambda d: r["aging_bucket"])
            name = f"{r['code']} - {r['name']}" if r.get("name") else str(r.get("code",""))
            rows += f"| {i} | {name} | {r.get('invoice_no') or 'N/A'} | {r.get('invoice_date') or 'N/A'} | {r.get('base_amount',0):,.0f} | {fn(r.get('over_day',0))} | {(r.get('description') or '').strip()} |\n"
        table = "| # | Khách hàng | Invoice | Ngày HĐ | Số tiền | Tình trạng | Mô tả |\n|---|---|---|---|---|---|---|\n" + rows
        lines.append(f"---\n\n## PIC: {pic}\n\n**Subject:** [AR] Thông báo công nợ — {report_date}\n\n{table}\nVui lòng đôn đốc khách hàng thanh toán.\n\nTrân trọng,\n[Ký tên]\n")
    if no_pic_records:
        by_cust: dict = {}
        for r in no_pic_records:
            key = (r.get("code",""), r.get("name") or str(r.get("code","")))
            by_cust.setdefault(key, []).append(r)
        for (code, cname), recs in by_cust.items():
            total = sum(r.get("base_amount",0) for r in recs)
            rows = ""
            for r in recs:
                rows += f"| {cname} | {(r.get('description') or '').strip()} | {r.get('invoice_date') or 'N/A'} | {r.get('base_amount',0):,.0f} |\n"
            table = "| Customer Name | Description | Invoice Date | Amount |\n|---|---|---|---|\n" + rows + f"| **Tổng Cộng** | | | **{total:,.0f}** |\n"
            lines.append(f"---\n\n## {cname}\n\n**Subject:** [AR] THƯ NHẮC NỢ THANH TOÁN — {report_date}\n\n{table}\nVui lòng kiểm tra và thanh toán.\n\nTrân trọng,\n[Ký tên]\n")
    return lines


def _generate_new_email_lines(data: dict, period: str) -> list:
    account = data.get("account", "")
    lines = [f"\n---\n## TK{account} — Kỳ {period}\n"]
    if data.get("by_employee"):
        for emp, records in data["by_employee"].items():
            total = sum(r.get("amount",0) for r in records)
            rows = "".join(f"| {r.get('invoice_date') or 'N/A'} | {r.get('invoice_no','')} | {r.get('description','')} | {r.get('amount',0):,.0f} | {r.get('reimbursement_date') or 'Chưa có'} |\n" for r in records)
            table = "| Ngày CT | Số CT | Diễn giải | Số tiền (VND) | Ngày hoàn ứng |\n|---|---|---|---|---|\n" + rows + f"| | | **Tổng cộng** | **{total:,.0f}** | |\n"
            lines.append(f"---\n\n## {emp}\n\n**Subject:** [Nhắc hoàn ứng] Kỳ {period}\n\n{table}\nVui lòng hoàn tất hồ sơ.\n\nTrân trọng,\nBộ phận Kế toán\n")
    elif data.get("by_vendor"):
        for vendor, records in data["by_vendor"].items():
            total = sum(r.get("total_vnd",0) for r in records)
            rows = "".join(f"| {r.get('invoice_date') or 'N/A'} | {r.get('invoice_no','')} | {r.get('description','')} | {r.get('total_vnd',0):,.0f} |\n" for r in records)
            table = "| Ngày HĐ | Số HĐ | Diễn giải | Tổng VND |\n|---|---|---|---|\n" + rows + f"| | | **Tổng cộng** | **{total:,.0f}** |\n"
            lines.append(f"---\n\n## {vendor}\n\n**Subject:** [AP TK{account}] Công nợ kỳ {period}\n\n{table}\n\nTrân trọng,\nBộ phận Kế toán\n")
    elif data.get("by_supplier"):
        for sup, records in data["by_supplier"].items():
            total = sum(r.get("amount",0) for r in records)
            rows = "".join(f"| {r.get('gl_date') or 'N/A'} | {r.get('invoice_no','')} | {r.get('description','')} | {r.get('amount',0):,.0f} |\n" for r in records)
            table = "| Ngày GL | Số HĐ | Diễn giải | Số tiền |\n|---|---|---|---|\n" + rows + f"| | | **Tổng cộng** | **{total:,.0f}** |\n"
            lines.append(f"---\n\n## {sup}\n\n**Subject:** [TK{account}] Nhắc công nợ kỳ {period}\n\n{table}\n\nTrân trọng,\nBộ phận Kế toán\n")
    return lines


# ── Routes ────────────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def ui():
    return _dashboard_html if _dashboard_html else _empty_dashboard()


@app.post("/process")
async def process(files: List[UploadFile] = File(...)):
    """Upload 1+ file Excel, trả về ZIP gồm Master, SLA Tracker, Dashboard, Email Drafts."""
    global _last_output_dir, _dashboard_html
    work_dir = tempfile.mkdtemp()
    try:
        output_dir = Path(work_dir) / "output"
        output_dir.mkdir()
        bundle_period = datetime.now().strftime("T%m_%Y")
        all_parsed = []

        for upload in files:
            safe_name = re.sub(r'[^\w\-_. ]', '_', upload.filename)
            file_path = os.path.join(work_dir, safe_name)
            with open(file_path, "wb") as f:
                f.write(await upload.read())

            type_out  = run_script("detect_type.py", [file_path])
            type_info = json.loads(type_out)
            ftype   = type_info.get("type", "UNKNOWN")
            period  = type_info.get("period", bundle_period)
            account = type_info.get("account", "")

            if period in ("N/A", None, ""):
                period = bundle_period
            else:
                period = "T" + str(period).replace("/", "_")

            if ftype == "AR":
                parsed_data = json.loads(run_script("parse_ar.py", [file_path]))
                parsed_data.update({"file_type":"AR","period":period,"source_file":upload.filename,"account":"AR"})
                all_parsed.append(parsed_data)
            elif ftype in ("TAM_UNG", "AP_AGING", "PREPAY"):
                parser_map = {"TAM_UNG":"parse_tam_ung.py","AP_AGING":"parse_ap_aging.py","PREPAY":"parse_prepay.py"}
                parsed_data = json.loads(run_script(parser_map[ftype], [file_path]))
                parsed_data.update({"file_type":ftype,"period":period,"source_file":upload.filename})
                all_parsed.append(parsed_data)

        if not all_parsed:
            raise HTTPException(status_code=422, detail="Không nhận dạng được loại file nào.")

        combined_path = os.path.join(work_dir, "parsed_combined.json")
        Path(combined_path).write_text(
            json.dumps({"file_type":"COMBINED","period":bundle_period,"sources":all_parsed}, ensure_ascii=False),
            encoding="utf-8")

        run_script("generate_combined_master.py",      [combined_path, str(output_dir / f"Master_{bundle_period}.xlsx")])
        run_script("generate_combined_sla_tracker.py", [combined_path, str(output_dir / f"SLA_Tracker_{bundle_period}.xlsx")])
        run_script("generate_combined_dashboard.py",   [combined_path, str(output_dir / f"Dashboard_{bundle_period}.html")])

        email_lines = [f"# Email Drafts — {bundle_period}\n"]
        for src in all_parsed:
            if src.get("file_type") == "AR":
                email_lines.extend(_generate_ar_email_lines(src))
            else:
                email_lines.extend(_generate_new_email_lines(src, src.get("period", bundle_period)))
        Path(str(output_dir / f"EmailDrafts_{bundle_period}.md")).write_text("\n".join(email_lines), encoding="utf-8")

        _last_output_dir = str(output_dir)
        dash_path = str(output_dir / f"Dashboard_{bundle_period}.html")
        _dashboard_html = open(dash_path, encoding="utf-8").read()

        zip_base = os.path.join(work_dir, "ZionReports")
        shutil.make_archive(zip_base, "zip", str(output_dir))
        return FileResponse(path=zip_base+".zip", media_type="application/zip", filename="ZionReports_output.zip")

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/process")
async def api_process(files: List[UploadFile] = File(...)):
    """Giống /process nhưng trả về JSON rows để cập nhật dashboard trực tiếp."""
    global _last_output_dir, _dashboard_html
    work_dir = tempfile.mkdtemp()
    try:
        output_dir = Path(work_dir) / "output"
        output_dir.mkdir()
        bundle_period = datetime.now().strftime("T%m_%Y")
        all_parsed = []

        for upload in files:
            safe_name = re.sub(r'[^\w\-_. ]', '_', upload.filename)
            file_path = os.path.join(work_dir, safe_name)
            with open(file_path, "wb") as f:
                f.write(await upload.read())

            type_info = json.loads(run_script("detect_type.py", [file_path]))
            ftype   = type_info.get("type", "UNKNOWN")
            period  = type_info.get("period", bundle_period)
            if period in ("N/A", None, ""):
                period = bundle_period
            else:
                period = "T" + str(period).replace("/", "_")

            if ftype == "AR":
                parsed_data = json.loads(run_script("parse_ar.py", [file_path]))
                parsed_data.update({"file_type":"AR","period":period,"source_file":upload.filename,"account":"AR"})
                all_parsed.append(parsed_data)
            elif ftype in ("TAM_UNG", "AP_AGING", "PREPAY"):
                parser_map = {"TAM_UNG":"parse_tam_ung.py","AP_AGING":"parse_ap_aging.py","PREPAY":"parse_prepay.py"}
                parsed_data = json.loads(run_script(parser_map[ftype], [file_path]))
                parsed_data.update({"file_type":ftype,"period":period,"source_file":upload.filename})
                all_parsed.append(parsed_data)

        if not all_parsed:
            raise HTTPException(status_code=422, detail="Khong nhan dang duoc loai file nao.")

        combined_path = os.path.join(work_dir, "parsed_combined.json")
        Path(combined_path).write_text(
            json.dumps({"file_type":"COMBINED","period":bundle_period,"sources":all_parsed}, ensure_ascii=False),
            encoding="utf-8")

        try:
            run_script("generate_combined_master.py",      [combined_path, str(output_dir / f"Master_{bundle_period}.xlsx")])
            run_script("generate_combined_sla_tracker.py", [combined_path, str(output_dir / f"SLA_Tracker_{bundle_period}.xlsx")])
            run_script("generate_combined_dashboard.py",   [combined_path, str(output_dir / f"Dashboard_{bundle_period}.html")])
            email_lines = [f"# Email Drafts — {bundle_period}\n"]
            for src in all_parsed:
                if src.get("file_type") == "AR":
                    email_lines.extend(_generate_ar_email_lines(src))
                else:
                    email_lines.extend(_generate_new_email_lines(src, src.get("period", bundle_period)))
            Path(str(output_dir / f"EmailDrafts_{bundle_period}.md")).write_text("\n".join(email_lines), encoding="utf-8")
            _last_output_dir = str(output_dir)
            _dashboard_html = open(str(output_dir / f"Dashboard_{bundle_period}.html"), encoding="utf-8").read()
        except Exception:
            pass

        rows_data = json.loads(run_script("generate_combined_dashboard.py", [combined_path, "--json"]))
        rows_data["rows"] = merge_tracker(rows_data["rows"])
        return JSONResponse(rows_data)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/download")
def download():
    if not _last_output_dir:
        raise HTTPException(status_code=404, detail="Chua co du lieu. Upload file truoc.")
    zip_base = os.path.join(tempfile.mkdtemp(), "ZionReports")
    shutil.make_archive(zip_base, "zip", _last_output_dir)
    return FileResponse(path=zip_base+".zip", media_type="application/zip", filename="ZionReports_output.zip")


@app.post("/api/clear")
def api_clear():
    global _dashboard_html
    _dashboard_html = ""
    return {"status": "cleared"}


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
            capture_output=True, text=True, encoding="utf-8"
        )
        if result.returncode != 0:
            raise HTTPException(status_code=500, detail=result.stderr.strip() or result.stdout.strip())

        return {"status": "ok", "detail": result.stdout.strip()}
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
