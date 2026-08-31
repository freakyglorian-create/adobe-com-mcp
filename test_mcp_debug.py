"""Minimal MCP test - just initialize and read stderr."""
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

stderr_lines = []
def stderr_reader():
    while True:
        try:
            line = proc.stderr.readline()
        except Exception:
            break
        if not line:
            break
        stderr_lines.append(line.decode("utf-8", errors="replace").rstrip())

t_err = threading.Thread(target=stderr_reader, daemon=True)
t_err.start()

stdout_buf = b""
stdout_msgs = []
def stdout_reader():
    global stdout_buf
    while True:
        try:
            chunk = proc.stdout.read(1)
        except Exception:
            break
        if not chunk:
            break
        stdout_buf += chunk
        while b"\r\n\r\n" in stdout_buf:
            header_end = stdout_buf.index(b"\r\n\r\n")
            header = stdout_buf[:header_end].decode("ascii", errors="replace")
            body_start = header_end + 4
            cl = None
            for line in header.split("\r\n"):
                if line.lower().startswith("content-length:"):
                    cl = int(line.split(":", 1)[1].strip())
                    break
            if cl is None:
                stdout_buf = stdout_buf[body_start:]
                break
            if len(stdout_buf) - body_start < cl:
                break
            body = stdout_buf[body_start:body_start + cl]
            stdout_buf = stdout_buf[body_start + cl:]
            try:
                stdout_msgs.append(json.loads(body.decode("utf-8")))
            except Exception as e:
                stdout_msgs.append({"_parse_error": str(e)})

t_out = threading.Thread(target=stdout_reader, daemon=True)
t_out.start()

time.sleep(2)
print(f"=== STDERR after 2s ({len(stderr_lines)} lines) ===")
for line in stderr_lines:
    print(f"  {line}")

print(f"\n=== STDOUT messages after 2s: {len(stdout_msgs)} ===")
for m in stdout_msgs:
    print(f"  {json.dumps(m, ensure_ascii=False)[:200]}")

# Send initialize
print("\n=== Sending initialize ===")
body = json.dumps({
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2024-11-05",
        "capabilities": {},
        "clientInfo": {"name": "test", "version": "1.0"}
    }
}, ensure_ascii=False).encode("utf-8")
frame = f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body
print(f"  Frame size: {len(frame)} bytes, body: {len(body)} bytes")
proc.stdin.write(frame)
proc.stdin.flush()

time.sleep(3)

print(f"\n=== STDERR after initialize ({len(stderr_lines)} lines) ===")
for line in stderr_lines:
    print(f"  {line}")

print(f"\n=== STDOUT messages after initialize: {len(stdout_msgs)} ===")
for m in stdout_msgs:
    print(f"  {json.dumps(m, ensure_ascii=False)[:300]}")

# Also check raw stdout buffer
print(f"\n=== Raw stdout buffer tail: {len(stdout_buf)} bytes ===")
print(f"  {stdout_buf[-200:]}")

proc.terminate()
try:
    proc.wait(timeout=2)
except Exception:
    proc.kill()
