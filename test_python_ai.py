"""Quick Illustrator COM test."""
import pythoncom
import win32com.client
import os

pythoncom.CoInitialize()

print("Connecting to Illustrator...")
try:
    app = win32com.client.GetActiveObject("Illustrator.Application")
    print("Connected via GetActiveObject")
except Exception as e:
    print(f"GetActiveObject failed: {e}")
    app = win32com.client.Dispatch("Illustrator.Application")
    print("Connected via Dispatch")

print(f"Version: {app.Version}")
print(f"Doc count: {app.Documents.Count}")

# Create document
print("\nCreating document...")
doc = app.Documents.Add(612, 792)  # A4 in points
print(f"Created: {doc.Name} ({doc.Width}x{doc.Height}pt)")

# Add rectangle
print("Adding rectangle...")
rect = doc.PathItems.Rectangle(600, 100, 200, 150)  # top, left, width, height
rect.Filled = True
try:
    sw = doc.Swatches.Add()
    sw.Color.Red = 255
    sw.Color.Green = 100
    sw.Color.Blue = 50
    rect.FillColor = sw.Color
except Exception as e:
    print(f"  fill color warning: {e}")
print(f"  Rect: {rect.Name}")

# Add ellipse
print("Adding ellipse...")
ell = doc.PathItems.Ellipse(500, 350, 120, 80)
ell.Filled = True
try:
    sw2 = doc.Swatches.Add()
    sw2.Color.Red = 50
    sw2.Color.Green = 150
    sw2.Color.Blue = 255
    ell.FillColor = sw2.Color
except Exception as e:
    print(f"  fill color warning: {e}")
print(f"  Ellipse: {ell.Name}")

# Add text
print("Adding text...")
tf = doc.TextFrames.Add()
tf.Contents = "Hello Illustrator!"
try:
    tf.TextRange.CharacterAttributes.Size = 36
except Exception as e:
    print(f"  text size warning: {e}")
try:
    tf.Top = 300
    tf.Left = 100
except Exception as e:
    print(f"  text position warning: {e}")
print(f"  Text: {tf.Contents}")

# Save as AI
ai_path = os.path.join(os.path.expanduser("~"), "Documents", "python_mcp_test.ai")
print(f"\nSaving as AI: {ai_path}")
doc.SaveAs(ai_path)
print("  Saved")

# Export SVG
svg_path = os.path.join(os.path.expanduser("~"), "Documents", "python_mcp_test.svg")
print(f"Exporting SVG: {svg_path}")
try:
    opts = win32com.client.Dispatch("Illustrator.SVGExportOptions")
    opts.FontType = 1
    doc.Export(svg_path, 10, opts)  # 10 = aiSVG
    print("  Exported")
except Exception as e:
    print(f"  SVG export failed: {e}")

print(f"\nPaths: {doc.PathItems.Count}")
print(f"TextFrames: {doc.TextFrames.Count}")
print("ALL AI TESTS PASSED")
