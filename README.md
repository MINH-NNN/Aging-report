# Aging Report Agent

Tự động xử lý file AR Excel (Aged Debtor Report) và tạo bộ báo cáo hoàn chỉnh gồm:

- **AR Master** — tổng hợp công nợ theo PIC
- **SLA Tracker** — theo dõi Payment SLA (NET30) và PIC Response SLA (3 ngày làm việc)
- **SLA Dashboard** — trang HTML trực quan về tình trạng SLA
- **Email Drafts** — email nhắc thanh toán cho từng PIC

## Sử dụng nhanh

Truy cập endpoint đã deploy:

```
https://endpoint-27eb5189-b34d-4ce1-9a2c-7fdbc97411b2.agentbase-runtime.aiplatform.vngcloud.vn
```

Upload file AR Excel → nhận file ZIP kết quả.

## Cấu trúc repo

```
├── app.py                  # FastAPI server (dùng để deploy lên cloud)
├── main.py                 # Chạy trực tiếp trên terminal (local)
├── Dockerfile              # Đóng gói thành Docker image
├── requirements.txt        # Thư viện Python cần thiết
├── deploy.py               # Script deploy lên GreenNode AgentBase
├── status.py               # Kiểm tra trạng thái runtime
├── AgingReport.html        # Giao diện HTML standalone (local)
├── run_agent.py            # Giao diện desktop (tkinter)
└── scripts/
    ├── parse_ar.py             # Đọc và phân tích file AR Excel
    ├── generate_master.py      # Tạo AR Master Excel
    ├── generate_sla_tracker.py # Tạo SLA Tracker Excel
    └── generate_sla_dashboard.py # Tạo SLA Dashboard HTML
```

## Chạy local

```bash
pip install -r requirements.txt
python main.py --input <ar_file.xlsx> --output-dir ./output
```

## Deploy lên GreenNode AgentBase

Yêu cầu: Docker đang chạy, có credentials GreenNode.

```bash
python deploy.py
```

Xem trạng thái sau khi deploy:

```bash
python status.py
```
