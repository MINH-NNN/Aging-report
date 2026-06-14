#!/usr/bin/env python3
"""Chuyển endpoint sang public (không yêu cầu xác thực). Chạy: python make_public.py"""

import base64, json, urllib.request, urllib.error, urllib.parse, sys

CLIENT_ID     = "4e15b882-af6e-4acb-afed-159172ec2998"
CLIENT_SECRET = "579ef815-5d01-4a54-883a-483fd1d9e049"
RUNTIME_NAME  = "aging-report-agent"
IAM_URL       = "https://iam.api.vngcloud.vn/accounts-api/v2/auth/token"
AB_API        = "https://agentbase.api.vngcloud.vn/runtime"

def http(method, url, token=None, payload=None, form=None, basic=None):
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
    elif payload is not None:
        data = json.dumps(payload).encode()
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            body = r.read().decode()
            return r.status, json.loads(body) if body.strip() else {}
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        print(f"  HTTP {e.code} {method} {url}")
        print(f"  Response: {body[:500]}")
        return e.code, {}

# ── Lấy token ────────────────────────────────────────────────────────────────
_, resp = http("POST", IAM_URL, basic=(CLIENT_ID, CLIENT_SECRET), form={"grant_type": "client_credentials"})
TOKEN = resp["access_token"]
print("✅ IAM token OK")

# ── Tìm runtime ───────────────────────────────────────────────────────────────
RUNTIME_ID = None
_, rlist = http("GET", f"{AB_API}/agent-runtimes?page=1&size=100", token=TOKEN)
for rt in (rlist.get("listData") or []):
    if rt.get("name") == RUNTIME_NAME:
        RUNTIME_ID = rt["id"]
        break

if not RUNTIME_ID:
    print(f"❌ Không tìm thấy runtime '{RUNTIME_NAME}'"); sys.exit(1)
print(f"✅ Runtime ID: {RUNTIME_ID}")

# ── Lấy danh sách endpoint ────────────────────────────────────────────────────
_, eps = http("GET", f"{AB_API}/agent-runtimes/{RUNTIME_ID}/endpoints", token=TOKEN)
print(f"\nEndpoints hiện tại:")
print(json.dumps(eps, indent=2, ensure_ascii=False))

ep_list = eps.get("listData") or (eps.get("data") or {}).get("listData") or []
if isinstance(eps, list):
    ep_list = eps

EP_ID      = None
EP_VERSION = None
for ep in ep_list:
    print(f"\n  [{ep.get('name')}] id={ep.get('id')} url={ep.get('url')}")
    print(f"  Full data: {json.dumps(ep, ensure_ascii=False)}")
    if ep.get("name") == "DEFAULT":
        EP_ID      = ep.get("id")
        EP_VERSION = ep.get("version") or ep.get("targetVersion")

if not EP_ID:
    print("❌ Không tìm thấy DEFAULT endpoint"); sys.exit(1)

# ── Thử PATCH endpoint sang public ───────────────────────────────────────────
print(f"\n🔄 Thử update endpoint {EP_ID} sang public...")

# Thử các payload khác nhau
payloads_to_try = [
    {"authEnabled": False},
    {"isPublic": True},
    {"authRequired": False},
    {"authentication": {"enabled": False}},
    {"accessControl": "public"},
]

for payload in payloads_to_try:
    url = f"{AB_API}/agent-runtimes/{RUNTIME_ID}/endpoints/{EP_ID}"
    if EP_VERSION:
        url += f"?version={EP_VERSION}"
    code, result = http("PATCH", url, token=TOKEN, payload=payload)
    print(f"\n  Payload: {payload}")
    print(f"  Result ({code}): {json.dumps(result, ensure_ascii=False)[:300]}")
    if code in (200, 201, 204):
        print(f"  ✅ Thành công với payload: {payload}")
        break
