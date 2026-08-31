"""Test: threading + COM in Python with win32com."""
import pythoncom
import win32com.client
import threading
import time

def ps_work(label):
    pythoncom.CoInitialize()
    print(f"[{label}] Connecting...")
    app = win32com.client.GetActiveObject("Photoshop.Application")
    print(f"[{label}] Connected. Calling DoJavaScript...")
    for i in range(5):
        try:
            r = app.DoJavaScript("app.documents.length.toString()")
            print(f"[{label}] Try {i+1}: {r}")
            break
        except Exception as e:
            print(f"[{label}] Try {i+1} failed: {str(e)[:60]}")
            time.sleep(0.3)
    pythoncom.CoUninitialize()

# Run from main thread
print("=== Main thread ===")
ps_work("main")

# Run from another thread
print("\n=== Worker thread ===")
t = threading.Thread(target=ps_work, args=("worker",))
t.start()
t.join()

print("\nDone")
