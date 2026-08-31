"""Full test of Photoshop COM via Python + DoJavaScript."""
import sys
import time
import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("Photoshop.Application")

def run_js(js):
    return app.DoJavaScript(js)

pass_count = 0
fail_count = 0

def test(name, fn):
    global pass_count, fail_count
    print(f"\n=== {name} ===")
    try:
        r = fn()
        print(f"  PASS: {r}")
        pass_count += 1
    except Exception as e:
        print(f"  FAIL: {e}")
        fail_count += 1

# 1. Create document
def t1():
    r = run_js("""
        var doc = app.documents.add(800, 600, 72, 'Python-MCP-Test');
        doc.name + '|' + Math.round(doc.width.value) + 'x' + Math.round(doc.height.value);
    """)
    return r

test("Create 800x600 document", t1)

# 2. Get info
def t2():
    r = run_js("""
        var d = app.activeDocument;
        d.name + '|' + Math.round(d.width.value) + '|' + Math.round(d.height.value) + '|' + d.mode + '|' + d.artLayers.length;
    """)
    return r

test("Get active doc info", t2)

# 3. Add layer
def t3():
    r = run_js("""
        var d = app.activeDocument;
        var L = d.artLayers.add();
        L.name = 'Test Layer';
        L.name + '|' + d.artLayers.length;
    """)
    return r

test("Add layer", t3)

# 4. Fill with color
def t4():
    r = run_js("""
        var d = app.activeDocument;
        var c = new SolidColor();
        c.rgb.red = 255; c.rgb.green = 150; c.rgb.blue = 50;
        app.foregroundColor = c;
        d.selection.selectAll();
        d.selection.fill(app.foregroundColor);
        d.selection.deselect();
        'filled';
    """)
    return r

test("Fill layer with orange", t4)

# 5. Add text
def t5():
    r = run_js("""
        var d = app.activeDocument;
        var tf = d.artLayers.add();
        tf.kind = LayerKind.TEXT;
        tf.textItem.contents = 'Hello Photoshop!';
        tf.textItem.size = 48;
        tf.textItem.position = [50/72, 100/72];
        var c = new SolidColor();
        c.rgb.red = 255; c.rgb.green = 255; c.rgb.blue = 255;
        tf.textItem.color = c;
        tf.name + '|' + tf.textItem.contents;
    """)
    return r

test("Add text layer", t5)

# 6. Duplicate layer
def t6():
    r = run_js("""
        var d = app.activeDocument;
        var L = d.activeLayer.duplicate();
        L.name + '|' + d.artLayers.length;
    """)
    return r

test("Duplicate active layer", t6)

# 7. Gaussian blur
def t7():
    r = run_js("""
        var d = app.activeDocument;
        d.activeLayer.applyGaussianBlur(3);
        'blurred';
    """)
    return r

test("Gaussian blur 3px", t7)

# 8. Resize
def t8():
    r = run_js("""
        var d = app.activeDocument;
        d.resizeImage(400, 300);
        Math.round(d.width.value) + 'x' + Math.round(d.height.value);
    """)
    return r

test("Resize to 400x300", t8)

# 9. Save as PNG
import os
png_path = os.path.join(os.path.expanduser("~"), "Documents", "python_mcp_test.png")
def t9():
    r = run_js(f"""
        var d = app.activeDocument;
        var opts = new PNGSaveOptions();
        opts.compression = 6;
        d.saveAs(new File('{png_path.replace(chr(92), chr(47))}'), opts, true);
        'saved';
    """)
    return r + " -> " + png_path

test("Save as PNG", t9)

# Summary
print(f"\n{'='*40}")
print(f"Results: {pass_count} passed, {fail_count} failed")
print(f"PNG saved to: {png_path}")
