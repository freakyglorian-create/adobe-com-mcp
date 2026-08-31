"""Quick test: connect to Photoshop via pywin32 and run DoJavaScript."""
import sys
import time
import pythoncom
import win32com.client

pythoncom.CoInitialize()

print("Connecting to Photoshop...")
try:
    app = win32com.client.GetActiveObject("Photoshop.Application")
    print("Connected via GetActiveObject")
except Exception as e:
    print(f"GetActiveObject failed: {e}")
    app = win32com.client.Dispatch("Photoshop.Application")
    print("Connected via Dispatch")

print(f"Version: {app.Version}")
print(f"Doc count: {app.Documents.Count}")

# Try DoJavaScript with retry
print("\nTrying DoJavaScript...")
js_code = """
    if (app.documents.length === 0) {
        var d = app.documents.add(800, 800, 72, 'Python-Test');
    } else {
        var d = app.activeDocument;
    }
    d.name + '|' + Math.round(d.width.value) + 'x' + Math.round(d.height.value);
"""

last_err = None
for i in range(15):
    try:
        result = app.DoJavaScript(js_code)
        print(f"Try {i+1} SUCCESS: {result}")
        break
    except Exception as e:
        last_err = e
        print(f"Try {i+1} failed: {str(e)[:80]}")
        time.sleep(0.3)
else:
    print(f"FAILED after 15 tries: {last_err}")
