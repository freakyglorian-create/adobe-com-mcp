"""Proper MCP stdio test with Content-Length framing and initialize handshake."""
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
    text=False,  # binary mode for proper framing
    bufsize=0,
)

def send_request(method, params=None, request_id=None):
    """Send JSON-RPC request with Content-Length framing."""
    msg = {
        "jsonrpc": "2.0",
        "method": method,
    }
    if params is not None:
        "params" not in msg and msg.__setitem__("params", params)
        msg["params"] = params
    if request_id is not None:
        msg["id"] = request_id
    body = json.dumps(msg, ensure_ascii=False).encode("utf-8")
    frame = f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body
    proc.stdin.write(frame)
    proc.stdin.flush()

responses = []
def reader():
    buf = b""
    while True:
        chunk = proc.stdout.read(1)
        if not chunk:
            break
        buf += chunk
        # Parse Content-Length frames
        while b"\r\n\r\n" in buf:
            header_end = buf.index(b"\r\n\r\n")
            header = buf[:header_end].decode("ascii")
            body_start = header_end + 4
            # Find Content-Length
            cl = None
            for line in header.split("\r\n"):
                if line.lower().startswith("content-length:"):
                    cl = int(line.split(":", 1)[1].strip())
                    break
            if cl is None:
                buf = buf[body_start:]
                break
            if len(buf) - body_start < cl:
                # Need more data
                break
            body = buf[body_start:body_start + cl]
            buf = buf[body_start + cl:]
            try:
                responses.append(json.loads(body.decode("utf-8")))
            except Exception as e:
                responses.append({"_parse_error": str(e), "_raw": body[:200]})

t = threading.Thread(target=reader, daemon=True)
t.start()

# 1. Initialize
print("1. Sending initialize...")
send_request("initialize", {
    "protocolVersion": "2024-11-05",
    "capabilities": {},
    "clientInfo": {"name": "test", "version": "1.0"}
}, request_id=1)

time.sleep(2)
print(f"   Responses so far: {len(responses)}")
for r in responses:
    if "result" in r:
        caps = r["result"].get("capabilities", {})
        print(f"   Server: {r['result'].get('serverInfo', {})}")
        print(f"   Capabilities: {list(caps.keys())}")

# 2. Send initialized notification
print("\n2. Sending initialized notification...")
send_request("notifications/initialized")
time.sleep(0.5)

# 3. List tools
print("\n3. Sending tools/list...")
send_request("tools/list", request_id=2)
time.sleep(2)

# Find the tools/list response
tool_resp = None
for r in responses:
    if r.get("id") == 2 and "result" in r:
        tool_resp = r
        break

if tool_resp:
    tools = tool_resp["result"].get("tools", [])
    print(f"\n   Found {len(tools)} tools:")
    for t in tools:
        print(f"   - {t['name']}")
else:
    print(f"\n   No tools/list response found. Total responses: {len(responses)}")
    for r in responses:
        print(f"   {json.dumps(r, ensure_ascii=False)[:200]}")

proc.terminate()
try:
    proc.wait(timeout=3)
except Exception:
    proc.kill()
