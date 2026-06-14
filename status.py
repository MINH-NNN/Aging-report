#!/usr/bin/env python3
"""Kiểm tra trạng thái AgentBase Runtime. Chạy: python status.py"""

import base64, json, time, urllib.request, urllib.error, urllib.parse

CLIENT_ID     = "4e15b882-af6e-4acb-afed-159172ec2998"
CLIENT_SECRET = "579ef815-5d01-4a54-883a-483fd1d9e049"
RUNTIME_NAME  = "aging-report-agent"
IAM_URL       = "https://iam.api.vngcloud.vn/accounts-api/v2/auth/token"
AB_API        = "https://agentbase.api.vngcloud.vn/runtime"

def http(method, url, token=None, form=None, basic=None):
    headers = {"Accept": "application/json"}
    data = None
    if basic:
        cred = base64.b64encode(f"{basic[0]}:{basic[1]}".encode()).decode()
        headers["Authorization"] = f"Basic {cred}"
    elif token:
        headers["Authorization"] = f"Bearer {token}"
    if form:
        data = urllib.parse.urlencode(form).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=15) as r:
        body = r.read().decode()
        return json.loads(body) if body.strip() else {}

# Lấy token
resp = http("POST", IAM_URL, basic=(CLIENT_ID, CLIENT_SECRET), form={"grant_type": "client_credentials"})
TOKEN = resp["access_token"]

# Tìm runtime
page = 1
RUNTIME_ID = None
while True:
    rlist = http("GET", f"{AB_API}/agent-runtimes?page={page}&size=100", token=TOKEN)
    for rt in (rlist.get("listData") or []):
        if rt.get("name") == RUNTIME_NAME:
            RUNTIME_ID = rt["id"]
            break
    if RUNTIME_ID or page >= (rlist.get("totalPage") or 1):
        break
    page += 1

if not RUNTIME_ID:
    print(f"❌ Không tìm thấy runtime '{RUNTIME_NAME}'")
    exit(1)

# Lấy chi tiết
rt = http("GET", f"{AB_API}/agent-runtimes/{RUNTIME_ID}", token=TOKEN)
data = rt.get("data") or rt
status = data.get("status", "?")

print(f"\n{'═'*50}")
print(f"Runtime:  {RUNTIME_NAME}")
print(f"ID:       {RUNTIME_ID}")
print(f"Status:   {status}")
print(f"Flavor:   {data.get('flavorId','?')}")

# Lấy endpoint
eps = http("GET", f"{AB_API}/agent-runtimes/{RUNTIME_ID}/endpoints", token=TOKEN)
ep_list = eps.get("listData") or (eps.get("data") or {}).get("listData") or []
for ep in ep_list:
    print(f"Endpoint: {ep.get('name','?')} → {ep.get('url','?')}")

print(f"Console:  https://aiplatform.console.vngcloud.vn/agent-runtime?tab=runtime")
print(f"{'═'*50}\n")

import sys
if "--delete" in sys.argv:
    confirm = input(f"\n⚠️  Xóa runtime '{RUNTIME_NAME}' (id={RUNTIME_ID})? [yes/N]: ").strip()
    if confirm.lower() == "yes":
        http("DELETE", f"{AB_API}/agent-runtimes/{RUNTIME_ID}", token=TOKEN)
        print("🗑️  Đã xóa runtime. Chạy lại: python deploy.py")
    else:
        print("Hủy.")
elif status == "CREATING":
    print("⏳ Runtime vẫn đang khởi động. Chạy lại sau 1-2 phút.")
elif status == "ACTIVE":
    print("✅ Runtime ACTIVE — sẵn sàng nhận request!")
elif status == "ERROR":
    print("❌ Runtime lỗi — kiểm tra logs trên console.")
    print(f"   Để xóa và deploy lại: python status.py --delete")
