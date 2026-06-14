#!/usr/bin/env python3
"""
AgentBase Deploy Script — Aging Report Agent
Chạy: python deploy.py
Yêu cầu: Docker đang chạy trên máy
"""

import argparse
import base64
import json
import os
import subprocess
import sys
import time
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime

parser = argparse.ArgumentParser()
parser.add_argument("--skip-build", action="store_true", help="Bỏ qua bước build & push (dùng image đã push)")
parser.add_argument("--image", default=None, help="Dùng image URL có sẵn (bỏ qua build)")
parser.add_argument("--status", action="store_true", help="Chỉ kiểm tra trạng thái runtime, không deploy")
parser.add_argument("--timeout", type=int, default=60, help="Số lần poll x10s (mặc định 60 = 10 phút)")
ARGS = parser.parse_args()

# Mode --status: chỉ check trạng thái, chạy sau khi đã có TOKEN và RUNTIME_ID
# (Xử lý ở cuối file sau khi định nghĩa các hàm helper)

# ── CONFIG ──────────────────────────────────────────────────────────────────
CLIENT_ID     = "4e15b882-af6e-4acb-afed-159172ec2998"
CLIENT_SECRET = "579ef815-5d01-4a54-883a-483fd1d9e049"

RUNTIME_NAME  = "aging-report-agent"
FLAVOR        = None   # tự động lấy từ API

IAM_URL       = "https://iam.api.vngcloud.vn/accounts-api/v2/auth/token"
AB_CR_API     = "https://agentbase.api.vngcloud.vn/cr/api/v1"   # AgentBase-integrated CR
VCR_HOST      = "vcr.vngcloud.vn"
AB_API        = "https://agentbase.api.vngcloud.vn/runtime"
# ────────────────────────────────────────────────────────────────────────────


def log(msg):  print(f"  {msg}")
def ok(msg):   print(f"  ✅ {msg}")
def err(msg):  print(f"  ❌ {msg}"); sys.exit(1)
def warn(msg): print(f"  ⚠️  {msg}")
def step(n, m): print(f"\n{'─'*60}\nStep {n}: {m}\n{'─'*60}")


def http(method, url, token=None, payload=None, form=None, basic=None, silent=False):
    data = None
    headers = {"Accept": "application/json"}
    if basic:
        cred = base64.b64encode(f"{basic[0]}:{basic[1]}".encode()).decode()
        headers["Authorization"] = f"Basic {cred}"
    elif token:
        headers["Authorization"] = f"Bearer {token}"
    if form:
        data = urllib.parse.urlencode(form).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    elif payload is not None:
        data = json.dumps(payload).encode()
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read().decode()
            return json.loads(body) if body.strip() else {}
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        if not silent:
            log(f"HTTP {e.code} {method} {url}")
            log(f"Response: {body[:400]}")
        raise


def run(cmd, check=True):
    log(f"$ {' '.join(cmd)}")
    r = subprocess.run(cmd, text=True, capture_output=True)
    if r.stdout.strip(): print("   " + r.stdout.strip()[:300])
    if r.returncode != 0 and check:
        if r.stderr.strip(): print("   ERR: " + r.stderr.strip()[:300])
        err(f"Lệnh thất bại (exit {r.returncode})")
    return r


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — IAM Token
# ══════════════════════════════════════════════════════════════════════════════
step(1, "Xác thực IAM")
try:
    resp = http("POST", IAM_URL,
                basic=(CLIENT_ID, CLIENT_SECRET),
                form={"grant_type": "client_credentials"})
    TOKEN = resp.get("access_token")
    if not TOKEN:
        err(f"Không lấy được token: {resp}")
    ok(f"Token OK (expires_in={resp.get('expires_in','?')}s)")
except Exception as e:
    err(f"Lỗi IAM: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — AgentBase Container Registry: lấy repo info
# ══════════════════════════════════════════════════════════════════════════════
step(2, "AgentBase Container Registry — lấy thông tin repo")

def ab_cr(method, path, **kw):
    return http(method, AB_CR_API + path, token=TOKEN, **kw)

repo_info = ab_cr("GET", "/repository")
repo_data = repo_info.get("data") or repo_info
REPO_BACKEND = repo_data.get("name") or repo_data.get("backendName") or repo_data.get("repoName")
REGISTRY_URL = repo_data.get("registryUrl") or VCR_HOST
ok(f"Repo: {REGISTRY_URL}/{REPO_BACKEND}")

LAST_IMAGE_FILE = ".last_image"

if ARGS.image:
    IMAGE_FULL = ARGS.image
elif ARGS.skip_build:
    # Dùng image đã push lần trước
    if os.path.exists(LAST_IMAGE_FILE):
        IMAGE_FULL = open(LAST_IMAGE_FILE).read().strip()
        ok(f"Dùng image từ lần push trước: {IMAGE_FULL}")
    else:
        print("  ❌ Không tìm thấy .last_image. Chạy lại không có --skip-build để build mới.")
        sys.exit(1)
else:
    IMAGE_TAG  = f"v{datetime.now().strftime('%Y%m%d%H%M%S')}"
    IMAGE_FULL = f"{REGISTRY_URL}/{REPO_BACKEND}/aging-report:{IMAGE_TAG}"

log(f"Image: {IMAGE_FULL}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Lấy CR credentials để docker login
# ══════════════════════════════════════════════════════════════════════════════
step(3, "Lấy Container Registry credentials")

cred_info  = ab_cr("GET", "/registry-credential")
cred_data  = cred_info.get("data") or cred_info
ROBOT_USER   = cred_data.get("username")
ROBOT_SECRET = cred_data.get("password") or cred_data.get("secret")

if not ROBOT_USER or not ROBOT_SECRET:
    err(f"Không lấy được CR credentials: {cred_info}")
ok(f"CR username: {ROBOT_USER}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4 — Docker login
# ══════════════════════════════════════════════════════════════════════════════
step(4, "Docker login vào vCR")
login = subprocess.run(
    ["docker", "login", REGISTRY_URL, "-u", ROBOT_USER, "--password-stdin"],
    input=ROBOT_SECRET, text=True, capture_output=True
)
if login.returncode != 0:
    err(f"Docker login thất bại:\n{login.stderr}")
ok(f"Docker login OK: {REGISTRY_URL}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5 — Build Docker image
# ══════════════════════════════════════════════════════════════════════════════
if ARGS.image:
    IMAGE_FULL = ARGS.image
    step(5, f"Dùng image có sẵn: {IMAGE_FULL}")
elif ARGS.skip_build:
    step(5, "Bỏ qua build (--skip-build)")
    log(f"Dùng image: {IMAGE_FULL}")
else:
    step(5, "Build Docker image")
    log(f"Platform: linux/amd64  |  Tag: {IMAGE_FULL}")
    build = subprocess.run(
        ["docker", "build", "--platform", "linux/amd64", "-t", IMAGE_FULL, "."],
        text=True
    )
    if build.returncode != 0:
        err("Docker build thất bại. Kiểm tra Dockerfile và thử lại.")
    ok("Build thành công")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 6 — Push image
# ══════════════════════════════════════════════════════════════════════════════
if ARGS.image or ARGS.skip_build:
    step(6, "Bỏ qua push (dùng image có sẵn)")
else:
    step(6, "Push image lên vCR")
    push = subprocess.run(["docker", "push", IMAGE_FULL], text=True)
    if push.returncode != 0:
        err("Docker push thất bại. Kiểm tra kết nối và quyền robot account.")
    # Lưu image URL để dùng lại với --skip-build
    open(LAST_IMAGE_FILE, "w").write(IMAGE_FULL)
    ok(f"Push OK: {IMAGE_FULL}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 7 — Tạo / Cập nhật AgentBase Runtime
# ══════════════════════════════════════════════════════════════════════════════
step(7, "Tạo hoặc cập nhật AgentBase Runtime")

def ab(method, path, **kw):
    return http(method, AB_API + path, token=TOKEN, **kw)

# Lấy danh sách flavor và chọn tự động
flavor_resp = ab("GET", "/flavors")
flavor_list = flavor_resp if isinstance(flavor_resp, list) else (flavor_resp.get("listData") or flavor_resp.get("data") or [])
# Ưu tiên flavor hỗ trợ agent-runtime (PUBLIC mode)
suitable = [
    f for f in flavor_list
    if "agent-runtime" in (f.get("supportedResourceTypes") or [])
    and "agent-runtime-vpc" not in (f.get("supportedResourceTypes") or [])
]
if not suitable:
    suitable = flavor_list  # fallback: dùng bất kỳ flavor nào có

if not suitable:
    err("Không tìm thấy flavor nào. Kiểm tra quyền hoặc liên hệ GreenNode support.")

# In danh sách để người dùng biết
log("Danh sách flavor khả dụng:")
for f in suitable[:5]:
    log(f"  - {f.get('id') or f.get('name')} | {f.get('description','')}")

FLAVOR = suitable[0].get("id") or suitable[0].get("name")
ok(f"Chọn flavor: {FLAVOR}")

image_auth = {"enabled": True, "username": ROBOT_USER, "password": ROBOT_SECRET}

# Tìm runtime đã tồn tại
RUNTIME_ID = None
page = 1
while True:
    rlist = ab("GET", f"/agent-runtimes?page={page}&size=100")
    for rt in (rlist.get("listData") or []):
        if rt.get("name") == RUNTIME_NAME:
            RUNTIME_ID = rt["id"]
            break
    if RUNTIME_ID or page >= (rlist.get("totalPage") or 1):
        break
    page += 1

if RUNTIME_ID:
    warn(f"Runtime '{RUNTIME_NAME}' đã tồn tại (id={RUNTIME_ID}) → UPDATE")
    ab("PATCH", f"/agent-runtimes/{RUNTIME_ID}", payload={
        "description": "Aging Report Agent — AR Excel processing",
        "imageUrl":    IMAGE_FULL,
        "imageAuth":   image_auth,
        "flavorId":    FLAVOR,
        "command":     [],
        "args":        [],
        "environmentVariables": {},
        "autoscaling": {
            "minReplicas":       1,
            "maxReplicas":       1,
            "cpuUtilization":    50,
            "memoryUtilization": 50,
        },
    })
else:
    log(f"Tạo runtime mới: {RUNTIME_NAME}")
    resp = ab("POST", "/agent-runtimes", payload={
        "name":        RUNTIME_NAME,
        "description": "Aging Report Agent — AR Excel processing",
        "imageUrl":    IMAGE_FULL,
        "imageAuth":   image_auth,
        "flavorId":    FLAVOR,
        "command":     [],
        "args":        [],
        "environmentVariables": {},
        "autoscaling": {
            "minReplicas":       1,
            "maxReplicas":       1,
            "cpuUtilization":    50,
            "memoryUtilization": 50,
        },
    })
    RUNTIME_ID = (resp.get("data") or resp).get("id")
    if not RUNTIME_ID:
        err(f"Không lấy được runtime ID: {resp}")

ok(f"Runtime ID: {RUNTIME_ID}")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 8 — Chờ ACTIVE
# ══════════════════════════════════════════════════════════════════════════════
step(8, f"Chờ runtime ACTIVE (tối đa {ARGS.timeout * 10 // 60} phút)")
for i in range(ARGS.timeout):
    time.sleep(10)
    rt = ab("GET", f"/agent-runtimes/{RUNTIME_ID}")
    status = (rt.get("data") or rt).get("status", "?")
    log(f"[{(i+1)*10}s] Status: {status}")
    if status == "ACTIVE":
        ok("Runtime ACTIVE!")
        break
    if status == "ERROR":
        err(f"Runtime lỗi: {rt}")
else:
    warn("Timeout — kiểm tra console: https://aiplatform.console.vngcloud.vn/agent-runtime?tab=runtime")


# ══════════════════════════════════════════════════════════════════════════════
# STEP 9 — Lấy endpoint URL
# ══════════════════════════════════════════════════════════════════════════════
step(9, "Lấy endpoint URL")
eps = ab("GET", f"/agent-runtimes/{RUNTIME_ID}/endpoints")
ep_list = eps.get("listData") or (eps.get("data") or {}).get("listData") or []
endpoint_url = None
for ep in ep_list:
    if ep.get("name") == "DEFAULT" or ep.get("type") == "DEFAULT":
        endpoint_url = ep.get("url")
        break
if not endpoint_url and ep_list:
    endpoint_url = ep_list[0].get("url", "")


# ══════════════════════════════════════════════════════════════════════════════
# KẾT QUẢ
# ══════════════════════════════════════════════════════════════════════════════
print(f"\n{'═'*60}")
print("🚀 DEPLOY HOÀN THÀNH!")
print(f"{'═'*60}")
print(f"  Runtime:    {RUNTIME_NAME}")
print(f"  Runtime ID: {RUNTIME_ID}")
print(f"  Image:      {IMAGE_FULL}")
print(f"  Endpoint:   {endpoint_url or '(lấy trên console)'}")
print(f"\n  Console: https://aiplatform.console.vngcloud.vn/agent-runtime?tab=runtime")
if endpoint_url:
    print(f"\n  Gọi API:")
    print(f"  curl -X POST {endpoint_url}/process \\")
    print(f"       -F 'ar_file=@<đường_dẫn_file_excel>'")
print(f"{'═'*60}\n")
