#!/usr/bin/env python3
"""
Aging Report Agent — Entry point
Nhận file AR Excel, xuất:
  1. AR_Master_<period>.xlsx   — master file tổng hợp theo PIC
  2. SLA_Tracker_<period>.xlsx — tracker SLA (Payment + PIC Response)
  3. SLA_Dashboard_<period>.html — dashboard HTML
  4. EmailDrafts_<period>.md   — email draft cho từng PIC

Usage:
  python main.py --input <ar_file.xlsx> [--tracker <existing_tracker.xlsx>] [--output-dir <dir>]

Arguments:
  --input     : (bắt buộc) đường dẫn file AR Excel
  --tracker   : (tuỳ chọn) tracker cũ từ tháng trước để merge dữ liệu
  --output-dir: (tuỳ chọn) thư mục output, mặc định ./output
"""

import argparse
import sys
import json
import re
import subprocess
from pathlib import Path
from datetime import datetime

SCRIPTS = Path(__file__).parent / "scripts"


def run_script(script_name, args):
    result = subprocess.run(
        [sys.executable, str(SCRIPTS / script_name)] + args,
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"[ERROR] {script_name}:\n{result.stderr}", file=sys.stderr)
        sys.exit(1)
    return result.stdout


def get_period(ar_file):
    """Extract period string from report date in AR file."""
    out = run_script("parse_ar.py", [ar_file])
    data = json.loads(out)
    report_date = data.get("report_date", "")
    match = re.search(r"(\d{2})/(\d{4})", report_date)
    if match:
        return f"T{match.group(1)}_{match.group(2)}"
    return datetime.now().strftime("T%m_%Y")


def generate_email_drafts(ar_file, output_path):
    """Generate email drafts markdown from AR data."""
    out = run_script("parse_ar.py", [ar_file])
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


def main():
    parser = argparse.ArgumentParser(description="Aging Report Agent")
    parser.add_argument("--input", required=True, help="Đường dẫn file AR Excel")
    parser.add_argument("--tracker", default=None, help="Tracker cũ từ tháng trước (tuỳ chọn)")
    parser.add_argument("--output-dir", default="output", help="Thư mục output (mặc định: ./output)")
    args = parser.parse_args()

    ar_file = args.input
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Determine period label
    period = get_period(ar_file)
    print(f"📅 Kỳ báo cáo: {period}")

    # 1. Master Excel
    master_path = str(output_dir / f"AR_Master_{period}.xlsx")
    print(f"📊 Tạo master file: {master_path}")
    run_script("generate_master.py", [ar_file, master_path])

    # 2. SLA Tracker
    tracker_path = str(output_dir / f"SLA_Tracker_{period}.xlsx")
    tracker_input = args.tracker if args.tracker else tracker_path
    print(f"📋 Tạo SLA tracker: {tracker_path}")
    run_script("generate_sla_tracker.py", [ar_file, tracker_path])

    # 3. SLA Dashboard
    dashboard_path = str(output_dir / f"SLA_Dashboard_{period}.html")
    print(f"🖥️  Tạo SLA dashboard: {dashboard_path}")
    run_script("generate_sla_dashboard.py", [ar_file, tracker_path, dashboard_path])

    # 4. Email Drafts
    email_path = str(output_dir / f"EmailDrafts_{period}.md")
    print(f"✉️  Tạo email drafts: {email_path}")
    generate_email_drafts(ar_file, email_path)

    print(f"\n✅ Hoàn thành! Output tại: {output_dir.resolve()}/")
    print(f"   ├── AR_Master_{period}.xlsx")
    print(f"   ├── SLA_Tracker_{period}.xlsx")
    print(f"   ├── SLA_Dashboard_{period}.html")
    print(f"   └── EmailDrafts_{period}.md")


if __name__ == "__main__":
    main()
