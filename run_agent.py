#!/usr/bin/env python3
"""
Aging Report Agent — Giao diện đơn giản
Chạy: python run_agent.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import urllib.request
import urllib.error
import threading
import os
import tempfile
import subprocess

ENDPOINT = "https://endpoint-27eb5189-b34d-4ce1-9a2c-7fdbc97411b2.agentbase-runtime.aiplatform.vngcloud.vn"


def call_api(ar_file_path, output_dir, status_var, btn):
    """Gọi API và lưu kết quả (chạy trong thread riêng)."""
    try:
        status_var.set("⏳ Đang xử lý...")

        # Dùng requests nếu có, fallback urllib
        try:
            import requests as req_lib
            filename = os.path.basename(ar_file_path)
            with open(ar_file_path, "rb") as f:
                resp = req_lib.post(
                    f"{ENDPOINT}/process",
                    files={"ar_file": (filename, f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                    timeout=120,
                )
            if resp.status_code != 200:
                raise Exception(f"Server lỗi {resp.status_code}: {resp.text[:300]}")
            zip_data = resp.content
        except ImportError:
            # Fallback: urllib với multipart thủ công
            with open(ar_file_path, "rb") as f:
                ar_data = f.read()
            filename = os.path.basename(ar_file_path)
            boundary = "AgentBoundaryXYZ789"
            body = (
                f"--{boundary}\r\n"
                f'Content-Disposition: form-data; name="ar_file"; filename="{filename}"\r\n'
                f"Content-Type: application/octet-stream\r\n\r\n"
            ).encode() + ar_data + f"\r\n--{boundary}--\r\n".encode()
            req = urllib.request.Request(
                f"{ENDPOINT}/process", data=body,
                headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=120) as r:
                zip_data = r.read()

        # Lưu file ZIP
        zip_name = "AgingReport_output.zip"
        zip_path = os.path.join(output_dir, zip_name)
        with open(zip_path, "wb") as f:
            f.write(zip_data)

        status_var.set(f"✅ Xong! Đã lưu: {zip_path}")

        # Hỏi có mở thư mục không
        if messagebox.askyesno("Hoàn thành", f"Đã tạo xong!\n\nFile lưu tại:\n{zip_path}\n\nMở thư mục ngay?"):
            subprocess.Popen(f'explorer /select,"{zip_path}"')

    except urllib.error.HTTPError as e:
        body = e.read().decode()
        status_var.set(f"❌ Lỗi server: {e.code}")
        messagebox.showerror("Lỗi", f"Server trả về lỗi {e.code}:\n{body[:300]}")
    except Exception as e:
        status_var.set(f"❌ Lỗi: {e}")
        messagebox.showerror("Lỗi", str(e))
    finally:
        btn.config(state="normal")


def browse_file(entry_var):
    path = filedialog.askopenfilename(
        title="Chọn file AR Excel",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if path:
        entry_var.set(path)


def browse_output(entry_var):
    path = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
    if path:
        entry_var.set(path)


def run(ar_var, out_var, status_var, btn):
    ar_path = ar_var.get().strip()
    out_path = out_var.get().strip()

    if not ar_path:
        messagebox.showwarning("Thiếu file", "Vui lòng chọn file AR Excel.")
        return
    if not os.path.exists(ar_path):
        messagebox.showerror("Lỗi", f"File không tồn tại:\n{ar_path}")
        return
    if not out_path:
        out_path = os.path.dirname(ar_path)
        out_var.set(out_path)

    btn.config(state="disabled")
    thread = threading.Thread(target=call_api, args=(ar_path, out_path, status_var, btn), daemon=True)
    thread.start()


# ── Giao diện ────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Aging Report Agent")
root.resizable(False, False)

pad = {"padx": 10, "pady": 6}

# Header
tk.Label(root, text="Aging Report Agent", font=("Segoe UI", 14, "bold")).pack(**pad)
tk.Label(root, text="Chọn file AR Excel để tạo báo cáo tự động", fg="gray").pack()

frame = tk.Frame(root, padx=15, pady=10)
frame.pack(fill="x")

# File AR
tk.Label(frame, text="File AR Excel:", anchor="w").grid(row=0, column=0, sticky="w", pady=4)
ar_var = tk.StringVar()
tk.Entry(frame, textvariable=ar_var, width=50).grid(row=0, column=1, padx=6)
tk.Button(frame, text="Chọn...", command=lambda: browse_file(ar_var)).grid(row=0, column=2)

# Thư mục lưu
tk.Label(frame, text="Lưu kết quả vào:", anchor="w").grid(row=1, column=0, sticky="w", pady=4)
out_var = tk.StringVar(value=os.path.expanduser("~/Desktop"))
tk.Entry(frame, textvariable=out_var, width=50).grid(row=1, column=1, padx=6)
tk.Button(frame, text="Chọn...", command=lambda: browse_output(out_var)).grid(row=1, column=2)

# Nút xử lý
status_var = tk.StringVar(value="Sẵn sàng")
btn = tk.Button(
    root, text="▶  Tạo báo cáo", font=("Segoe UI", 11, "bold"),
    bg="#2563EB", fg="white", padx=20, pady=8,
    command=lambda: run(ar_var, out_var, status_var, btn)
)
btn.pack(pady=(0, 6))

# Status
tk.Label(root, textvariable=status_var, fg="#374151").pack(pady=(0, 10))

root.mainloop()
