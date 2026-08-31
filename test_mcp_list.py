"""Test the MCP server by running it and sending a tools/list request via stdin."""
import subprocess
import json
import sys
import os

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

# Send tools/list request
request = {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "tools/list",
    "params": {}
}
proc.stdin.write(json.dumps(request, ensure_ascii=False) + "\n")
proc.stdin.flush()

# Read response (first non-empty line that starts with {)
import time
time.sleep(3)

# Read all available stdout
output_lines = []
try:
    proc.stdout.flush()
    import select
    # Simple approach: read line by line with a timeout
    import threading
    result = []
    def read_stdout():
        for line in proc.stdout:
            result.append(line.rstrip())
    t = threading.Thread(target=read_stdout, daemon=True)
    t.start()
    t.join(timeout=5)
except Exception as e:
    print(f"Read error: {e}")

proc.terminate()

stderr = proc.stderr.read() if proc.stderr else ""

print(f"=== STDOUT ({len(result)} lines) ===")
for line in result[:20]:
    # Try to parse as JSON and pretty-print the tool list part
    try:
        data = json.loads(line)
        if "result" in data and "tools" in data["result"]:
            tools = data["result"]["tools"]
            print(f"Found {len(tools)} tools:")
            for t in tools:
                print(f"  - {t['name']}: {t.get('description','')[:60]}")
        else:
            print(line[:200])
    except Exception:
        if line.strip():
            print(line[:200])

print(f"\n=== STDERR (first 1000 chars) ===")
print(stderr[:1000])
