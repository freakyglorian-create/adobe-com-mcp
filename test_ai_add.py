"""Test AI Documents.Add with named args."""
import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("Illustrator.Application")

print("Trying Add with named args...")
try:
    doc = app.Documents.Add(Width=612, Height=792)
    print(f"SUCCESS (named): {doc.Name} {doc.Width}x{doc.Height}")
except Exception as e:
    print(f"Named args failed: {e}")

print("\nTrying Add() with no args...")
try:
    doc2 = app.Documents.Add()
    print(f"SUCCESS (no args): {doc2.Name} {doc2.Width}x{doc2.Height}")
except Exception as e:
    print(f"No args failed: {e}")
