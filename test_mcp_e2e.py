"""End-to-end MCP test: initialize, list tools, call ps_create_document."""
import subprocess
import json
import sys
import os
import time
import threading

python_exe = sys.executable
server_path = os.path.join(os.path.dirname(__file__), "server.py")

proc = subprocess.Popen(
    [python_exe, server_path],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=False,
    bufsize=0,
)

def send_msg(msg_dict):
    body = json.dumps(msg_dict, ensure_ascii=False).encode("utf-8")
    frame = f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body
    proc.stdin.write(frame)
    proc.stdin.flush()

responses = {}  # id -> response
notifications = []

def reader():
    buf = b""
    while True:
        try:
            chunk = proc.stdout.read(1)
        except Exception:
            break
        if not chunk:
            break
        buf += chunk
        while b"\r\n\r\n" in buf:
            header_end = buf.index(b"\r\n\r\n")
            header = buf[:header_end].decode("ascii", errors="replace")
            body_start = header_end + 4
            cl = None
            for line in header.split("\r\n"):
                if line.lower().startswith("content-length:"):
                    cl = int(line.split(":", 1)[1].strip())
                    break
            if cl is None:
                buf = buf[body_start:]
                break
            if len(buf) - body_start < cl:
                break
            body = buf[body_start:body_start + cl]
            buf = buf[body_start + cl:]
            try:
                data = json.loads(body.decode("utf-8"))
                if "id" in data:
                    responses[data["id"]] = data
                else:
                    notifications.append(data)
            except Exception:
                pass

t = threading.Thread(target=reader, daemon=True)
t.start()

def wait_for_response(req_id, timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        if req_id in responses:
            return responses[req_id]
        time.sleep(0.1)
    return None

# Step 1: Initialize
print("=== Step 1: Initialize ===")
send_msg({
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2024-11-05",
        "capabilities": {},
        "clientInfo": {"name": "test-client", "version": "1.0.0"}
    }
})
resp = wait_for_response(1, timeout=10)
if resp:
    info = resp.get("result", {}).get("serverInfo", {})
    print(f"  Server: {info.get('name', '?')} {info.get('version', '?')}")
    caps = resp.get("result", {}).get("capabilities", {})
    print(f"  Capabilities: {list(caps.keys())}")
else:
    print("  NO RESPONSE")

# Step 2: Initialized notification
send_msg({"jsonrpc": "2.0", "method": "notifications/initialized"})
time.sleep(0.3)

# Step 3: List tools
print("\n=== Step 2: List tools ===")
send_msg({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
resp = wait_for_response(2, timeout=5)
if resp and "result" in resp:
    tools = resp["result"].get("tools", [])
    print(f"  Total tools: {len(tools)}")
    ps_tools = [t for t in tools if t["name"].startswith("ps_")]
    ai_tools = [t for t in tools if t["name"].startswith("ai_")]
    print(f"  Photoshop tools ({len(ps_tools)}):")
    for t in ps_tools:
        print(f"    - {t['name']}")
    print(f"  Illustrator tools ({len(ai_tools)}):")
    for t in ai_tools:
        print(f"    - {t['name']}")
else:
    print("  NO RESPONSE")

# Step 4: Call ps_create_document
print("\n=== Step 3: Call ps_create_document(1024, 768) ===")
send_msg({
    "jsonrpc": "2.0",
    "id": 3,
    "method": "tools/call",
    "params": {
        "name": "ps_create_document",
        "arguments": {"width": 1024, "height": 768, "name": "MCP-E2E-Test"}
    }
})
resp = wait_for_response(3, timeout=15)
if resp and "result" in resp:
    content = resp["result"].get("content", [])
    print(f"  Response content:")
    for c in content:
        if c.get("type") == "text":
            # Try to parse as JSON for pretty printing
            try:
                data = json.loads(c["text"])
                print(f"    {json.dumps(data, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"    {c['text'][:200]}")
else:
    print("  NO RESPONSE or error")
    if resp:
        print(f"  {json.dumps(resp, ensure_ascii=False)[:300]}")

# Step 5: Call ps_add_text_layer
print("\n=== Step 4: Call ps_add_text_layer ===")
send_msg({
    "jsonrpc": "2.0",
    "id": 4,
    "method": "tools/call",
    "params": {
        "name": "ps_add_text_layer",
        "arguments": {
            "text": "Hello from TRAE MCP!",
            "size": 48,
            "x": 50,
            "y": 100,
            "r": 255, "g": 100, "b": 50
        }
    }
})
resp = wait_for_response(4, timeout=10)
if resp and "result" in resp:
    content = resp["result"].get("content", [])
    for c in content:
        if c.get("type") == "text":
            try:
                data = json.loads(c["text"])
                print(f"    {json.dumps(data, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"    {c['text'][:200]}")
else:
    print("  Error or no response")
    if resp:
        print(f"  {json.dumps(resp, ensure_ascii=False)[:300]}")

# Step 6: Call ps_get_active_info
print("\n=== Step 5: Call ps_get_active_info ===")
send_msg({
    "jsonrpc": "2.0",
    "id": 5,
    "method": "tools/call",
    "params": {"name": "ps_get_active_info", "arguments": {}}
})
resp = wait_for_response(5, timeout=10)
if resp and "result" in resp:
    content = resp["result"].get("content", [])
    for c in content:
        if c.get("type") == "text":
            try:
                data = json.loads(c["text"])
                print(f"    {json.dumps(data, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"    {c['text'][:200]}")

# Step 7: AI create document
print("\n=== Step 6: Call ai_create_document(800, 600) ===")
send_msg({
    "jsonrpc": "2.0",
    "id": 6,
    "method": "tools/call",
    "params": {
        "name": "ai_create_document",
        "arguments": {"width": 800, "height": 600}
    }
})
resp = wait_for_response(6, timeout=10)
if resp and "result" in resp:
    content = resp["result"].get("content", [])
    for c in content:
        if c.get("type") == "text":
            try:
                data = json.loads(c["text"])
                print(f"    {json.dumps(data, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"    {c['text'][:200]}")
else:
    print("  Error or no response")
    if resp:
        print(f"  {json.dumps(resp, ensure_ascii=False)[:300]}")

# Done
print("\n=== DONE ===")
proc.terminate()
try:
    proc.wait(timeout=3)
except Exception:
    proc.kill()
