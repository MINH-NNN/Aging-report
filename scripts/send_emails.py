#!/usr/bin/env python3
"""
Gửi email nhắc thanh toán cho từng PIC qua Gmail.

Usage:
  python send_emails.py <ar_file.xlsx> <pic_emails.json> [--dry-run]

  --dry-run: In ra email sẽ gửi mà không gửi thật
"""

import json
import os
import re
import smtplib
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# Đọc .env nếu có
def load_env():
    env_path = Path(__file__).parent.parent / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

load_env()

GMAIL_USER     = os.environ.get("GMAIL_USER", "")
GMAIL_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
FORM_BASE_URL  = "https://docs.google.com/forms/d/e/1FAIpQLSfxeYGq-JXp5iLgkOHLjOC6WscrWJbetzMoQsBwkYcCZ2Ck1Q/viewform"
ENTRY_HO_TEN   = "entry.1273518422"

def get_form_url(pic_name: str) -> str:
    """Tạo prefilled form URL với tên PIC điền sẵn."""
    from urllib.parse import urlencode
    params = urlencode({ENTRY_HO_TEN: pic_name})
    return f"{FORM_BASE_URL}?usp=pp_url&{params}"

AGING_STATUS = {
    "Over 180 ngày": lambda d: f"Quá hạn {d} ngày 🔴🔴",
    "91-180 ngày":   lambda d: f"Quá hạn {d} ngày 🔴",
    "61-90 ngày":    lambda d: f"Quá hạn {d} ngày ⚠️",
    "31-60 ngày":    lambda d: f"Quá hạn {d} ngày",
    "1-30 ngày":     lambda d: f"Quá hạn {d} ngày",
    "Current":       lambda d: "Chưa đến hạn",
    "N/A":           lambda d: "N/A",
}


def build_email_body(pic, records, report_date, form_url=None):
    if form_url is None:
        form_url = get_form_url(pic)
    rows = ""
    for i, r in enumerate(records, 1):
        fn = AGING_STATUS.get(r["aging_bucket"], lambda d: r["aging_bucket"])
        status = fn(r["over_day"])
        amount = f"{r['base_amount']:,.0f}"
        inv = f"#{r['invoice_no']}" if r["invoice_no"] else ""
        name = f"{r['code']} - {r['name']}" if r["name"] else str(r["code"])
        desc = (r["description"] or "").strip()
        inv_date = r["invoice_date"] if r["invoice_date"] else ""
        rc = int(r.get("reminder_count") or 0)
        rc_label = f"Lần {rc}" if rc > 0 else "1"
        rc_color = "#dc2626" if rc >= 3 else "#d97706" if rc == 2 else "#2563eb"
        rows += f"""
        <tr>
          <td>{i}</td>
          <td>{name}</td>
          <td>{inv}</td>
          <td>{inv_date}</td>
          <td style="text-align:right">{amount}</td>
          <td>{status}</td>
          <td>{desc}</td>
          <td style="text-align:center;color:{rc_color};font-weight:bold">{rc_label}</td>
        </tr>"""

    return f"""
<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
<p>Kính gửi anh/chị <strong>{pic}</strong>,</p>
<p>Bộ phận AR xin gửi thông tin về các khoản công nợ hiện đang chờ thanh toán
thuộc trách nhiệm theo dõi của anh/chị:</p>

<table border="1" cellpadding="6" cellspacing="0"
       style="border-collapse:collapse;width:100%;font-size:13px">
  <thead style="background:#2563eb;color:white">
    <tr>
      <th>#</th><th>Khách hàng</th><th>Invoice</th>
      <th>Ngày HĐ</th><th>Số tiền (VND)</th><th>Tình trạng</th><th>Mô tả</th>
      <th>Lần nhắc</th>
    </tr>
  </thead>
  <tbody>{rows}
  </tbody>
</table>

<p>Anh/chị vui lòng đôn đốc khách hàng thanh toán các khoản trên trong thời gian sớm nhất,
và nhấn nút bên dưới để cập nhật kế hoạch thanh toán / lý do (nếu có vướng mắc):</p>

<div style="text-align:center;margin:28px 0">
  <a href="{form_url}"
     style="background:#2563eb;color:white;padding:14px 32px;border-radius:8px;
            text-decoration:none;font-weight:bold;font-size:15px;display:inline-block;
            letter-spacing:.3px">
    &#128203; CẬP NHẬT KẾ HOẠCH THANH TOÁN
  </a>
</div>

<p style="font-size:12px;color:#9ca3af;border-top:1px solid #e5e7eb;padding-top:12px;margin-top:8px">
  Đây là email tự động từ hệ thống AR. Vui lòng không phản hồi email này —
  mọi phản hồi xin gửi qua form ở trên.
</p>

<p>Trân trọng,<br>
<strong>Bộ phận AR</strong></p>
</body></html>"""


def send_emails(ar_file, pic_emails_file, dry_run=False):
    # Parse AR data
    import subprocess
    scripts_dir = Path(__file__).parent
    result = subprocess.run(
        [sys.executable, str(scripts_dir / "parse_ar.py"), ar_file],
        capture_output=True, text=True, encoding="utf-8"
    )
    if result.returncode != 0:
        print(f"[ERROR] parse_ar.py: {result.stderr}", file=sys.stderr)
        sys.exit(1)

    data = json.loads(result.stdout)
    report_date = data["report_date"]
    by_pic = data["by_pic"]

    # Đọc mapping PIC → email
    with open(pic_emails_file, encoding="utf-8") as f:
        pic_emails = json.load(f)

    if not dry_run and (not GMAIL_USER or not GMAIL_PASSWORD):
        print("[ERROR] Chưa cấu hình GMAIL_USER / GMAIL_APP_PASSWORD trong .env", file=sys.stderr)
        sys.exit(1)

    # Kết nối Gmail
    smtp = None
    if not dry_run:
        smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        smtp.login(GMAIL_USER, GMAIL_PASSWORD)
        print(f"✅ Đăng nhập Gmail: {GMAIL_USER}")

    log = []
    not_found = []

    for pic, records in by_pic.items():
        to_email = pic_emails.get(pic)
        if not to_email:
            not_found.append(pic)
            print(f"  ⚠️  Không tìm thấy email cho PIC: {pic} — bỏ qua")
            continue

        subject = "[AR] Thông báo công nợ chờ thanh toán"
        body    = build_email_body(pic, records, report_date)

        if dry_run:
            print(f"\n{'─'*60}")
            print(f"  [DRY RUN] To: {to_email}")
            print(f"  Subject: {subject}")
            print(f"  Số invoice: {len(records)}")
        else:
            msg = MIMEMultipart("alternative")
            msg["From"]     = f"AR System (no-reply) <{GMAIL_USER}>"
            msg["To"]       = to_email
            msg["Subject"]  = subject
            msg["Reply-To"] = "no-reply@donotreply.invalid"
            msg.attach(MIMEText(body, "html", "utf-8"))
            smtp.sendmail(GMAIL_USER, to_email, msg.as_string())
            sent_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"  ✅ Đã gửi → {to_email} ({sent_at})")
            log.append({"pic": pic, "email": to_email, "sent_at": sent_at, "invoices": len(records)})

    if smtp:
        smtp.quit()

    if not_found:
        print(f"\n⚠️  PIC chưa có email: {', '.join(not_found)}")
        print(f"   Thêm vào file: {pic_emails_file}")

    return log


def _send_one(pic, records, to_email, report_date, dry_run=False):
    """Gửi 1 email cho 1 PIC, tạo SMTP connection riêng (dùng trong thread)."""
    subject = "[AR] Thông báo công nợ chờ thanh toán"
    body = build_email_body(pic, records, report_date)
    if dry_run:
        print(f"  [DRY RUN] To: {to_email} | {len(records)} invoice(s)")
        return {"pic": pic, "email": to_email, "sent_at": None, "invoices": len(records)}
    try:
        smtp = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
        smtp.starttls()
        smtp.login(GMAIL_USER, GMAIL_PASSWORD)
        msg = MIMEMultipart("alternative")
        msg["From"]     = f"AR System (no-reply) <{GMAIL_USER}>"
        msg["To"]       = to_email
        msg["Subject"]  = subject
        msg["Reply-To"] = "no-reply@donotreply.invalid"
        msg.attach(MIMEText(body, "html", "utf-8"))
        smtp.sendmail(GMAIL_USER, to_email, msg.as_string())
        smtp.quit()
        sent_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"  ✅ Đã gửi → {to_email} ({sent_at})")
        return {"pic": pic, "email": to_email, "sent_at": sent_at, "invoices": len(records)}
    except Exception as e:
        print(f"  ❌ Lỗi gửi {to_email}: {e}", file=sys.stderr)
        return None


def send_emails_from_json(json_file, pic_emails_file, dry_run=False):
    """Gửi email từ file JSON (by_pic đã lọc) — song song, mỗi PIC 1 connection."""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    with open(json_file, encoding="utf-8") as f:
        data = json.load(f)
    report_date = data.get("report_date", "N/A")
    by_pic = data.get("by_pic", {})

    with open(pic_emails_file, encoding="utf-8") as f:
        pic_emails = json.load(f)

    if not dry_run and (not GMAIL_USER or not GMAIL_PASSWORD):
        print("[ERROR] Chưa cấu hình GMAIL_USER / GMAIL_APP_PASSWORD", file=sys.stderr)
        sys.exit(1)

    tasks = []
    not_found = []
    for pic, records in by_pic.items():
        to_email = pic_emails.get(pic)
        if not to_email:
            not_found.append(pic)
            print(f"  ⚠️  Không tìm thấy email cho PIC: {pic} — bỏ qua")
            continue
        tasks.append((pic, records, to_email))

    log = []
    with ThreadPoolExecutor(max_workers=min(len(tasks), 5)) as executor:
        futures = {
            executor.submit(_send_one, pic, records, to_email, report_date, dry_run): pic
            for pic, records, to_email in tasks
        }
        for future in as_completed(futures):
            result = future.result()
            if result and result.get("sent_at"):
                log.append(result)

    if not_found:
        print(f"\n⚠️  PIC chưa có email: {', '.join(not_found)}")
    return log


if __name__ == "__main__":
    import argparse, sys as _sys
    _sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser()
    parser.add_argument("ar_file")
    parser.add_argument("pic_emails_file")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--from-json", action="store_true", help="ar_file la JSON by_pic thay vi xlsx")
    args = parser.parse_args()

    print("\n" + "="*60)
    print(("[DRY RUN] " if args.dry_run else "") + "Gui email AR Report")
    print("="*60)
    if args.from_json:
        log = send_emails_from_json(args.ar_file, args.pic_emails_file, dry_run=args.dry_run)
    else:
        log = send_emails(args.ar_file, args.pic_emails_file, dry_run=args.dry_run)

    if not args.dry_run:
        print(f"\nDa gui {len(log)} email thanh cong")
        print(json.dumps(log, ensure_ascii=False, indent=2))
