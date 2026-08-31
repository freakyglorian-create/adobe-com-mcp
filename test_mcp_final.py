"""MCP E2E test - NDJSON (line-delimited JSON) format."""
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
    text=True,
    encoding="utf-8",
    bufsize=1,
)

responses = {}
notifications = []
stderr_lines = []

def stdout_reader():
    for line in proc.stdout:
        line = line.strip()
        if not line:
            continue
        try:
            data = json.loads(line)
            if "id" in data:
                responses[data["id"]] = data
            else:
                notifications.append(data)
        except Exception as e:
            notifications.append({"_raw": line[:200]})

def stderr_reader():
    for line in proc.stderr:
        stderr_lines.append(line.rstrip())

t_out = threading.Thread(target=stdout_reader, daemon=True)
t_err = threading.Thread(target=stderr_reader, daemon=True)
t_out.start()
t_err.start()

def send(msg_dict):
    line = json.dumps(msg_dict, ensure_ascii=False) + "\n"
    proc.stdin.write(line)
    proc.stdin.flush()

def wait_response(req_id, timeout=15):
    start = time.time()
    while time.time() - start < timeout:
        if req_id in responses:
            return responses[req_id]
        time.sleep(0.1)
    return None

time.sleep(2)
print(f"Server started. Stderr: {len(stderr_lines)} lines")
for l in stderr_lines:
    print(f"  {l}")

# 1. Initialize
print("\n=== 1. Initialize ===")
send({
    "jsonrpc": "2.0", "id": 1, "method": "initialize",
    "params": {
        "protocolVersion": "2024-11-05",
        "capabilities": {},
        "clientInfo": {"name": "test", "version": "1.0"}
    }
})
r = wait_response(1, 10)
if r and "result" in r:
    info = r["result"].get("serverInfo", {})
    print(f"  Server: {info.get('name','?')} {info.get('version','?')}")
else:
    print(f"  FAIL: {json.dumps(r, ensure_ascii=False)[:200] if r else 'timeout'}")

# 2. Initialized notification
send({"jsonrpc": "2.0", "method": "notifications/initialized"})
time.sleep(0.3)

# 3. List tools
print("\n=== 2. List tools ===")
send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
r = wait_response(2, 5)
if r and "result" in r:
    tools = r["result"].get("tools", [])
    ps = [t["name"] for t in tools if t["name"].startswith("ps_")]
    ai = [t["name"] for t in tools if t["name"].startswith("ai_")]
    print(f"  Total: {len(tools)} tools")
    print(f"  Photoshop ({len(ps)}): {', '.join(ps)}")
    print(f"  Illustrator ({len(ai)}): {', '.join(ai)}")
else:
    print(f"  FAIL")

# 4. ps_create_document
print("\n=== 3. ps_create_document(900, 600, 'E2E-Test') ===")
send({
    "jsonrpc": "2.0", "id": 3, "method": "tools/call",
    "params": {
        "name": "ps_create_document",
        "arguments": {"width": 900, "height": 600, "name": "E2E-Test"}
    }
})
r = wait_response(3, 15)
if r and "result" in r:
    for c in r["result"].get("content", []):
        if c.get("type") == "text":
            try:
                d = json.loads(c["text"])
                print(f"  {json.dumps(d, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"  {c['text'][:200]}")
else:
    print(f"  FAIL: {json.dumps(r, ensure_ascii=False)[:200] if r else 'timeout'}")

# 5. ps_add_text_layer
print("\n=== 4. ps_add_text_layer ===")
send({
    "jsonrpc": "2.0", "id": 4, "method": "tools/call",
    "params": {
        "name": "ps_add_text_layer",
        "arguments": {"text": "你好 TRAE MCP！", "size": 48, "x": 50, "y": 120, "r": 255, "g": 120, "b": 50}
    }
})
r = wait_response(4, 10)
if r and "result" in r:
    for c in r["result"].get("content", []):
        if c.get("type") == "text":
            try:
                d = json.loads(c["text"])
                print(f"  {json.dumps(d, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"  {c['text'][:200]}")
else:
    print(f"  FAIL")

# 6. ps_get_active_info
print("\n=== 5. ps_get_active_info ===")
send({
    "jsonrpc": "2.0", "id": 5, "method": "tools/call",
    "params": {"name": "ps_get_active_info", "arguments": {}}
})
r = wait_response(5, 5)
if r and "result" in r:
    for c in r["result"].get("content", []):
        if c.get("type") == "text":
            try:
                d = json.loads(c["text"])
                print(f"  {json.dumps(d, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"  {c['text'][:200]}")

# 7. ai_create_document
print("\n=== 6. ai_create_document ===")
send({
    "jsonrpc": "2.0", "id": 6, "method": "tools/call",
    "params": {"name": "ai_create_document", "arguments": {"width": 800, "height": 600}}
})
r = wait_response(6, 10)
if r and "result" in r:
    for c in r["result"].get("content", []):
        if c.get("type") == "text":
            try:
                d = json.loads(c["text"])
                print(f"  {json.dumps(d, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"  {c['text'][:200]}")
else:
    print(f"  FAIL")

# 8. ai_add_rectangle
print("\n=== 7. ai_add_rectangle ===")
send({
    "jsonrpc": "2.0", "id": 7, "method": "tools/call",
    "params": {
        "name": "ai_add_rectangle",
        "arguments": {"x": 100, "y": 300, "w": 200, "h": 150, "fr": 255, "fg": 100, "fb": 50}
    }
})
r = wait_response(7, 10)
if r and "result" in r:
    for c in r["result"].get("content", []):
        if c.get("type") == "text":
            try:
                d = json.loads(c["text"])
                print(f"  {json.dumps(d, ensure_ascii=False, indent=2)}")
            except Exception:
                print(f"  {c['text'][:200]}")

print("\n=== ALL DONE ===")
proc.terminate()
try:
    proc.wait(timeout=3)
except Exception:
    proc.kill()
