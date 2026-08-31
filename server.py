#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Adobe COM MCP Server — Photoshop + Illustrator (Windows).
Drives both apps via pywin32 COM automation directly from this process.
No UXP / proxy / PowerShell middleman needed.

PS 2020+ and AI CS6+ should work.
"""
import json
import sys
import os
import time
import math
import threading
from typing import Any

import win32com.client
import pythoncom
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("adobe-com-mcp")

# ---- Thread-local COM object cache ----
# COM objects are apartment-threaded; must not be shared across threads.
_tls = threading.local()


def _co_init():
    """Ensure COM is initialized on current thread (STA mode)."""
    if not getattr(_tls, "co_init", False):
        pythoncom.CoInitialize()
        _tls.co_init = True


def _get_ps():
    """Get Photoshop Application COM object for the current thread.
    Uses DoJavaScript for operations — the most reliable COM interface."""
    _co_init()
    app = getattr(_tls, "ps_app", None)
    if app is not None:
        return app
    # Try running instance first; launch if not running
    try:
        app = win32com.client.GetActiveObject("Photoshop.Application")
    except Exception:
        app = win32com.client.Dispatch("Photoshop.Application")
        # Give PS a moment to warm up if we just launched it
        time.sleep(2)
    # Suppress dialogs for non-interactive automation
    _ps_set_quiet_mode(app)
    _tls.ps_app = app
    return app


def _ps_set_quiet_mode(app):
    """Configure PS for headless automation: suppress dialogs, set ruler units."""
    for _ in range(10):
        try:
            app.DisplayDialogs = 3  # psDisplayNoDialogs — suppress everything
            break
        except Exception:
            time.sleep(0.5)
    try:
        app.Preferences.RulerUnits = 1  # psInches
    except Exception:
        pass


def _is_busy_error(e: Exception) -> bool:
    """Check if a COM error is the 'application busy' retryable type."""
    msg = str(e)
    return any(marker in msg for marker in [
        "0x8001010A", "8001010A",
        "应用程序正在使用中",
        "消息筛选器显示应用程序正在使用中",
        "application is busy",
        "RPC_E_SERVERCALL_RETRYLATER",
        "SERVERCALL_RETRYLATER",
    ])


def _ps_js(js_code: str, max_retries: int = 40, retry_sleep: float = 0.25) -> str:
    """Execute ExtendScript in Photoshop via DoJavaScript.
    Automatically retries on 'application busy' errors.
    """
    app = _get_ps()
    last_err = None
    for i in range(max_retries):
        try:
            return app.DoJavaScript(js_code)
        except Exception as e:
            last_err = e
            if _is_busy_error(e):
                time.sleep(retry_sleep)
                continue
            raise
    raise last_err


def _get_ai():
    """Get Illustrator Application COM object for the current thread."""
    _co_init()
    app = getattr(_tls, "ai_app", None)
    if app is not None:
        return app
    try:
        app = win32com.client.GetActiveObject("Illustrator.Application")
    except Exception:
        app = win32com.client.Dispatch("Illustrator.Application")
        time.sleep(1)
    _tls.ai_app = app
    return app


def _gcd(a: int, b: int) -> int:
    a, b = abs(a), abs(b)
    while b:
        a, b = b, a % b
    return a or 1


# ============================================================
# PHOTOSHOP TOOLS
# ============================================================

@mcp.tool()
def ps_create_document(width: int, height: int = 0, name: str = "") -> dict:
    """
    在 Photoshop 中新建指定像素尺寸的文档。
    当用户说"新建一个 800x800 的 Photoshop 文档""建一个 1920x1080 的画布"时使用。
    - width: 宽度（像素）
    - height: 高度（像素，不填则等于 width）
    - name: 文档名（可选）
    """
    if height <= 0:
        height = width
    width = max(1, int(width))
    height = max(1, int(height))
    # gcd trick: exact integer pixels via integer inches + ppi = gcd
    g = _gcd(width, height)
    w_in = width / g
    h_in = height / g
    js = f"""
        var doc = app.documents.add({w_in}, {h_in}, {g}, {json.dumps(name)});
        doc.name + '|' + Math.round(doc.width.value) + '|' + Math.round(doc.height.value);
    """
    result = _ps_js(js)
    parts = result.strip().split('|')
    return {
        "result": "ok",
        "name": parts[0] if parts else name,
        "width": width,
        "height": height,
        "unit": "pixels",
    }


@mcp.tool()
def ps_get_active_info() -> dict:
    """获取 Photoshop 当前活动文档信息：名称、像素尺寸、分辨率、颜色模式、图层数、活动图层名、打开文档数。"""
    js = """
        if (app.documents.length === 0) {
            'NO_DOCS|' + app.documents.length;
        } else {
            var d = app.activeDocument;
            var ppi = d.resolution;
            var pxW = Math.round(d.width.value * ppi);
            var pxH = Math.round(d.height.value * ppi);
            d.name + '|' + pxW + '|' + pxH
                + '|' + ppi + '|' + d.mode + '|' + d.artLayers.length
                + '|' + d.activeLayer.name + '|' + app.documents.length;
        }
    """
    result = _ps_js(js).strip()
    parts = result.split('|')
    if parts[0] == 'NO_DOCS':
        return {"result": "ok", "active": False, "open": int(parts[1])}
    return {
        "result": "ok",
        "name": parts[0],
        "pixelWidth": int(float(parts[1])),
        "pixelHeight": int(float(parts[2])),
        "resolution": float(parts[3]),
        "mode": parts[4],
        "layers": int(parts[5]),
        "activeLayer": parts[6],
        "open": int(parts[7]),
    }


@mcp.tool()
def ps_list_documents() -> dict:
    """列出 Photoshop 当前打开的所有文档名称。"""
    js = """
        var names = [];
        for (var i = 0; i < app.documents.length; i++) {
            names.push(app.documents[i].name);
        }
        names.join('\\n');
    """
    result = _ps_js(js).strip()
    names = [n for n in result.split('\n') if n] if result else []
    return {"result": "ok", "count": len(names), "documents": [{"name": n} for n in names]}


@mcp.tool()
def ps_close_document() -> dict:
    """关闭 Photoshop 当前活动文档（不保存）。"""
    js = """
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else { var n = app.activeDocument.name; app.activeDocument.close(SaveOptions.DONOTSAVECHANGES); n; }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "ok", "closed": False}
    return {"result": "ok", "closed": True, "name": result}


@mcp.tool()
def ps_add_layer(name: str = "") -> dict:
    """在当前 Photoshop 文档中新建一个空白图层。name 可选。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var L;
            if ({json.dumps(name)} !== '') {{ L = d.artLayers.add({json.dumps(name)}); }}
            else {{ L = d.artLayers.add(); }}
            L.name + '|' + d.artLayers.length;
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    return {"result": "ok", "layer": parts[0], "totalLayers": int(parts[1])}


@mcp.tool()
def ps_duplicate_layer() -> dict:
    """复制当前活动图层，生成副本。"""
    js = """
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var src = d.activeLayer.name;
            var L = d.activeLayer.duplicate();
            src + '|' + L.name + '|' + d.artLayers.length;
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    return {"result": "ok", "from": parts[0], "new": parts[1], "totalLayers": int(parts[2])}


@mcp.tool()
def ps_delete_layer() -> dict:
    """删除当前活动图层。"""
    js = """
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var n = d.activeLayer.name;
            d.activeLayer.remove();
            n + '|' + d.artLayers.length;
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    return {"result": "ok", "deleted": parts[0], "totalLayers": int(parts[1])}


@mcp.tool()
def ps_set_layer_opacity(opacity: int) -> dict:
    """设置当前活动图层的不透明度（0-100）。"""
    opacity = max(0, min(100, int(opacity)))
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            d.activeLayer.opacity = {opacity};
            d.activeLayer.name + '|' + d.activeLayer.opacity;
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    return {"result": "ok", "layer": parts[0], "opacity": float(parts[1])}


@mcp.tool()
def ps_add_text_layer(text: str, size: float = 48, x: float = 50, y: float = 100,
                      r: int = 0, g: int = 0, b: int = 0) -> dict:
    """
    在当前 Photoshop 文档中添加文字图层。
    - text: 文字内容
    - size: 字号（像素），默认 48
    - x, y: 左上角位置（像素）
    - r, g, b: 文字颜色（0-255）
    """
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var ppi = d.resolution;
            var xi = {x} / ppi;
            var yi = {y} / ppi;
            var tf = d.artLayers.add();
            tf.kind = LayerKind.TEXT;
            tf.textItem.contents = {json.dumps(text)};
            tf.textItem.size = {size};
            try {{ tf.textItem.position = [xi, yi]; }} catch(e) {{}}
            try {{
                var c = new SolidColor();
                c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
                tf.textItem.color = c;
            }} catch(e) {{}}
            tf.name + '|' + tf.textItem.contents + '|' + tf.textItem.size.value;
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    return {"result": "ok", "layer": parts[0], "text": parts[1], "size": float(parts[2]), "x": x, "y": y}


@mcp.tool()
def ps_fill_layer(r: int, g: int, b: int) -> dict:
    """用指定颜色填充当前活动图层（全画布填充）。r/g/b 为 0-255。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var c = new SolidColor();
            c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
            app.foregroundColor = c;
            d.selection.selectAll();
            try {{
                d.selection.fill(app.foregroundColor);
            }} catch(e) {{
                // Background layer may be locked - create new fill layer
                var fillLyr = d.artLayers.add();
                d.activeLayer = fillLyr;
                d.selection.selectAll();
                d.selection.fill(app.foregroundColor);
            }}
            d.selection.deselect();
            'ok';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "r": r, "g": g, "b": b}


@mcp.tool()
def ps_set_foreground_color(r: int, g: int, b: int) -> dict:
    """设置 Photoshop 前景色（RGB）。r/g/b 取值 0-255。"""
    js = f"""
        var c = new SolidColor();
        c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
        app.foregroundColor = c;
        'ok';
    """
    _ps_js(js)
    return {"result": "ok", "r": r, "g": g, "b": b}


@mcp.tool()
def ps_apply_gaussian_blur(radius: float = 5.0) -> dict:
    """对当前活动图层应用高斯模糊滤镜。radius 为模糊半径（像素），默认 5。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            app.activeDocument.activeLayer.applyGaussianBlur({radius});
            'ok';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "radius": radius}


@mcp.tool()
def ps_apply_unsharp_mask(amount: float = 100, radius: float = 2.0, threshold: float = 0) -> dict:
    """对当前活动图层应用 USM 锐化（非锐化蒙版）。amount 数量% (1-500)，radius 半径像素，threshold 阈值(0-255)。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            app.activeDocument.activeLayer.applyUnSharpMask({amount}, {radius}, {threshold});
            'ok';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "amount": amount, "radius": radius, "threshold": threshold}


@mcp.tool()
def ps_resize_document(width: int = 0, height: int = 0) -> dict:
    """
    调整当前 Photoshop 文档的像素尺寸。
    给 width 或 height 任意一个，另一个按比例自动算；两个都给就按指定尺寸。
    单位：像素。
    """
    if width <= 0 and height <= 0:
        return {"result": "error", "error": "at least one of width/height must be positive"}
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var ppi = d.resolution;
            var curW = d.width.value * ppi;
            var curH = d.height.value * ppi;
            var targetW = {width};
            var targetH = {height};
            if (targetW <= 0) {{ targetW = Math.round(curW * (targetH / curH)); }}
            if (targetH <= 0) {{ targetH = Math.round(curH * (targetW / curW)); }}
            d.resizeImage(targetW / ppi, targetH / ppi, ppi);
            targetW + '|' + targetH;
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    return {"result": "ok", "pixelWidth": int(float(parts[0])), "pixelHeight": int(float(parts[1]))}


@mcp.tool()
def ps_save_as_png(path: str = "") -> dict:
    """将当前 Photoshop 文档导出为 PNG 图片。path 为保存路径（可选，默认存到文档目录）。"""
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.png")
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var opts = new PNGSaveOptions();
            opts.compression = 6;
            opts.interlaced = false;
            app.activeDocument.saveAs(new File({json.dumps(path)}), opts, true);
            'ok';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "path": path}


@mcp.tool()
def ps_save_as_jpg(path: str = "", quality: int = 8) -> dict:
    """将当前 Photoshop 文档导出为 JPG 图片。quality 范围 1-12，默认 8。"""
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.jpg")
    quality = max(1, min(12, int(quality)))
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var opts = new JPEGSaveOptions();
            opts.quality = {quality};
            opts.formatOptions = FormatOptions.STANDARDBASELINE;
            app.activeDocument.saveAs(new File({json.dumps(path)}), opts, true);
            'ok';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "path": path, "quality": quality}


@mcp.tool()
def ps_save_as_psd(path: str = "") -> dict:
    """将当前 Photoshop 文档另存为 PSD 格式。path 可选。"""
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.psd")
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var opts = new PhotoshopSaveOptions();
            opts.embedColorProfile = true;
            opts.alphaChannels = true;
            opts.layers = true;
            app.activeDocument.saveAs(new File({json.dumps(path)}), opts, true);
            'ok';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "path": path}


@mcp.tool()
def ps_open_document(path: str) -> dict:
    """在 Photoshop 中打开一个本地图片/PSD 文件。传入文件完整路径。"""
    if not os.path.exists(path):
        return {"result": "error", "error": f"file not found: {path}"}
    js = f"""
        var f = new File({json.dumps(path)});
        var d = app.open(f);
        d.name;
    """
    name = _ps_js(js).strip()
    return {"result": "ok", "name": name, "path": path}


@mcp.tool()
def ps_do_action(action: str, frm: str = "") -> dict:
    """在 Photoshop 中执行一个已安装的动作（Action）。action 为动作名，frm 为动作集名（可选）。"""
    js = f"""
        app.doAction({json.dumps(action)}, {json.dumps(frm) if frm else '""'});
        'ok';
    """
    _ps_js(js)
    return {"result": "ok", "action": action, "from": frm}


@mcp.tool()
def ps_list_text_layers() -> dict:
    """
    识别并列出当前 Photoshop 文档中所有文字图层的详细信息。
    返回每个文字图层的：名称、文字内容、字号、颜色（RGB）、位置。
    用户说"看看有哪些文字""列出所有文字图层""识别文字"时使用。
    """
    js = r"""
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var results = [];
            for (var i = 0; i < d.artLayers.length; i++) {
                var lyr = d.artLayers[i];
                if (lyr.kind === LayerKind.TEXT) {
                    var ti = lyr.textItem;
                    var contents = '';
                    try { contents = ti.contents; } catch(e) {}
                    var size = '';
                    try { size = ti.size.value; } catch(e) { size = '0'; }
                    var r='0', g='0', b='0';
                    try {
                        r = Math.round(ti.color.rgb.red);
                        g = Math.round(ti.color.rgb.green);
                        b = Math.round(ti.color.rgb.blue);
                    } catch(e) {}
                    var px='0', py='0';
                    try {
                        px = Math.round(ti.position[0].value * d.resolution);
                        py = Math.round(ti.position[1].value * d.resolution);
                    } catch(e) {}
                    var op = '100';
                    try { op = Math.round(lyr.opacity); } catch(e) {}
                    results.push(
                        lyr.name + '\t' + contents + '\t' + size + '\t' + r + ',' + g + ',' + b + '\t' + px + ',' + py + '\t' + op
                    );
                }
            }
            results.join('\n');
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if not result:
        return {"result": "ok", "count": 0, "textLayers": [], "message": "当前文档没有文字图层"}
    layers = []
    for line in result.split('\n'):
        if not line.strip():
            continue
        parts = line.split('\t')
        rgb = parts[3] if len(parts) > 3 else "0,0,0"
        pos = parts[4] if len(parts) > 4 else "0,0"
        rgb_parts = rgb.split(',') if rgb else ["0", "0", "0"]
        pos_parts = pos.split(',') if pos else ["0", "0"]
        layers.append({
            "name": parts[0],
            "text": parts[1],
            "size": float(parts[2]) if len(parts) > 2 and parts[2] else 0,
            "color": {"r": int(float(rgb_parts[0])) if rgb_parts[0] else 0,
                      "g": int(float(rgb_parts[1])) if len(rgb_parts) > 1 and rgb_parts[1] else 0,
                      "b": int(float(rgb_parts[2])) if len(rgb_parts) > 2 and rgb_parts[2] else 0},
            "position": {"x": int(float(pos_parts[0])) if pos_parts[0] else 0,
                         "y": int(float(pos_parts[1])) if len(pos_parts) > 1 and pos_parts[1] else 0},
            "opacity": float(parts[5]) if len(parts) > 5 and parts[5] else 100,
        })
    return {"result": "ok", "count": len(layers), "textLayers": layers}


@mcp.tool()
def ps_select_all_text_layers() -> dict:
    """
    一键全选当前 Photoshop 文档中的所有文字图层。
    选中后可以对它们进行批量操作（如移动、对齐、合并等）。
    用户说"全选文字图层""选中所有文字""一键选中文字"时使用。
    """
    js = r"""
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var textLayers = [];
            for (var i = 0; i < d.artLayers.length; i++) {
                if (d.artLayers[i].kind === LayerKind.TEXT) {
                    textLayers.push(d.artLayers[i]);
                }
            }
            if (textLayers.length === 0) { 'NONE'; }
            else {
                // PS 2020 compatible: select each text layer one by one
                // Multi-select via Action Descriptor with putIndex
                var desc = new ActionDescriptor();
                var ref = new ActionReference();
                for (var j = 0; j < textLayers.length; j++) {
                    ref.putIndex(charIDToTypeID('Lyr '), textLayers[j].itemIndex);
                }
                desc.putReference(charIDToTypeID('null'), ref);
                try {
                    executeAction(charIDToTypeID('slct'), desc, DialogModes.NO);
                } catch(e) {
                    // Fallback: just activate first text layer
                    d.activeLayer = textLayers[0];
                }
                textLayers.length + '|' + (function() {
                    var names = [];
                    for (var k = 0; k < textLayers.length; k++) names.push(textLayers[k].name);
                    return names.join(',');
                })();
            }
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'NONE':
        return {"result": "ok", "selected": 0, "layers": [], "message": "当前文档没有文字图层"}
    parts = result.split('|')
    count = int(parts[0])
    names = parts[1].split(',') if len(parts) > 1 else []
    return {"result": "ok", "selected": count, "layers": names}


@mcp.tool()
def ps_replace_text_in_all_layers(old_text: str, new_text: str) -> dict:
    """
    批量替换所有文字图层中的指定文字内容。
    - old_text: 要查找的旧文字
    - new_text: 替换成的新文字
    用户说"把所有文字图层里的'标题'改成'新标题'"时使用。
    """
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var changed = 0;
            for (var i = 0; i < d.artLayers.length; i++) {{
                var lyr = d.artLayers[i];
                if (lyr.kind === LayerKind.TEXT) {{
                    var c = lyr.textItem.contents;
                    if (c.indexOf({json.dumps(old_text)}) >= 0) {{
                        lyr.textItem.contents = c.split({json.dumps(old_text)}).join({json.dumps(new_text)});
                        changed++;
                    }}
                }}
            }}
            changed.toString();
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    count = int(result)
    return {"result": "ok", "replaced": count, "oldText": old_text, "newText": new_text,
            "message": f"在 {count} 个文字图层中进行了替换"}


@mcp.tool()
def ps_set_text_color_all(r: int, g: int, b: int) -> dict:
    """
    批量设置所有文字图层的颜色。
    - r, g, b: RGB 颜色值（0-255）
    用户说"把所有文字改成白色""所有文字图层设为红色"时使用。
    """
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var changed = 0;
            var c = new SolidColor();
            c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
            for (var i = 0; i < d.artLayers.length; i++) {{
                var lyr = d.artLayers[i];
                if (lyr.kind === LayerKind.TEXT) {{
                    try {{ lyr.textItem.color = c; changed++; }} catch(e) {{}}
                }}
            }}
            changed.toString();
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    count = int(result)
    return {"result": "ok", "changed": count, "color": {"r": r, "g": g, "b": b},
            "message": f"已将 {count} 个文字图层的颜色改为 RGB({r},{g},{b})"}


@mcp.tool()
def ps_set_text_size_all(size: float) -> dict:
    """
    批量设置所有文字图层的字号。
    - size: 字号（像素）
    用户说"把所有文字字号改成 36""所有文字图层统一字号为 48"时使用。
    """
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var changed = 0;
            for (var i = 0; i < d.artLayers.length; i++) {{
                var lyr = d.artLayers[i];
                if (lyr.kind === LayerKind.TEXT) {{
                    try {{ lyr.textItem.size = {size}; changed++; }} catch(e) {{}}
                }}
            }}
            changed.toString();
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    count = int(result)
    return {"result": "ok", "changed": count, "size": size,
            "message": f"已将 {count} 个文字图层的字号改为 {size}"}


@mcp.tool()
def ps_list_all_layers() -> dict:
    """
    列出当前 Photoshop 文档中所有图层（含图层组）的完整信息。
    返回每个图层的：名称、类型、可见性、是否锁定、不透明度、是否为文字图层。
    用户说"看看有哪些图层""列出图层""文档结构"时使用。
    """
    js = r"""
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var results = [];
            function walk(layers, indent) {
                for (var i = 0; i < layers.length; i++) {
                    var lyr = layers[i];
                    var kind = '';
                    try { kind = lyr.typename; } catch(e) { kind = 'unknown'; }
                    var vis = '';
                    try { vis = lyr.visible; } catch(e) {}
                    var locked = '';
                    try { locked = lyr.allLocked; } catch(e) {}
                    var op = '';
                    try { op = lyr.opacity; } catch(e) {}
                    var isText = '';
                    try { isText = (lyr.kind === LayerKind.TEXT); } catch(e) { isText = false; }
                    var name = '';
                    try { name = lyr.name; } catch(e) {}
                    results.push(indent + name + '\t' + kind + '\t' + vis + '\t' + locked + '\t' + op + '\t' + isText);
                    // Recurse into layer sets
                    if (kind === 'LayerSet') {
                        walk(lyr.layerSets, indent + '  ');
                        walk(lyr.artLayers, indent + '  ');
                    }
                }
            }
            walk(d.layerSets, '');
            walk(d.artLayers, '');
            results.join('\n');
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if not result:
        return {"result": "ok", "count": 0, "layers": []}
    layers = []
    for line in result.split('\n'):
        if not line.strip():
            continue
        parts = line.split('\t')
        is_text = parts[5] == 'true' if len(parts) > 5 and parts[5] else False
        layers.append({
            "name": parts[0] if parts else "",
            "type": parts[1] if len(parts) > 1 else "",
            "indent": len(line) - len(line.lstrip()),
            "visible": parts[2] == 'true' if len(parts) > 2 and parts[2] else True,
            "locked": parts[3] == 'true' if len(parts) > 3 and parts[3] else False,
            "opacity": float(parts[4]) if len(parts) > 4 and parts[4] else 100,
            "isText": is_text,
        })
    return {"result": "ok", "count": len(layers), "layers": layers}


@mcp.tool()
def ps_select_layer_by_name(name: str) -> dict:
    """
    按名称选中 Photoshop 图层。
    - name: 图层名（精确匹配）
    用户说"选中背景图层""选中标题文字"时使用。
    """
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var found = false;
            // Search art layers
            for (var i = 0; i < d.artLayers.length; i++) {{
                if (d.artLayers[i].name === {json.dumps(name)}) {{
                    d.activeLayer = d.artLayers[i];
                    found = true;
                    break;
                }}
            }}
            // Search in layer sets if not found
            if (!found) {{
                function searchLayerSets(sets) {{
                    for (var i = 0; i < sets.length; i++) {{
                        var ls = sets[i];
                        for (var j = 0; j < ls.artLayers.length; j++) {{
                            if (ls.artLayers[j].name === {json.dumps(name)}) {{
                                d.activeLayer = ls.artLayers[j];
                                return true;
                            }}
                        }}
                        if (searchLayerSets(ls.layerSets)) return true;
                    }}
                    return false;
                }}
                found = searchLayerSets(d.layerSets);
            }}
            found ? 'OK|' + d.activeLayer.name : 'NOT_FOUND';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result.startswith('NOT_FOUND'):
        return {"result": "error", "error": f"未找到名为 '{name}' 的图层"}
    parts = result.split('|')
    return {"result": "ok", "selected": parts[1] if len(parts) > 1 else name}


@mcp.tool()
def ps_toggle_layer_visibility(name: str = "") -> dict:
    """
    切换图层可见性（显示/隐藏）。
    - name: 图层名（不填则切换当前活动图层）
    用户说"隐藏背景图层""显示标题图层""切换图层可见"时使用。
    """
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var lyr;
            if ({json.dumps(name)} !== '') {{
                for (var i = 0; i < d.artLayers.length; i++) {{
                    if (d.artLayers[i].name === {json.dumps(name)}) {{ lyr = d.artLayers[i]; break; }}
                }}
            }}
            if (!lyr) {{ lyr = d.activeLayer; }}
            if (lyr) {{
                lyr.visible = !lyr.visible;
                lyr.name + '|' + lyr.visible;
            }} else {{
                'NO_LAYER';
            }}
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'NO_LAYER':
        return {"result": "error", "error": "no layer found"}
    parts = result.split('|')
    visible = parts[1] == 'true' if len(parts) > 1 else False
    return {"result": "ok", "layer": parts[0], "visible": visible,
            "message": f"图层 '{parts[0]}' 已{'显示' if visible else '隐藏'}"}


# ============================================================
# PHOTOSHOP 高级批量操作
# ============================================================

@mcp.tool()
def ps_batch_resize_folder(folder: str, target_width: int = 0, target_height: int = 0,
                           suffix: str = "", output_format: str = "jpg",
                           quality: int = 8) -> dict:
    """
    批量处理文件夹中的所有图片：自动打开、调整尺寸、导出。
    可用于一键缩放整批素材到统一规格（如电商主图、缩略图）。
    - folder: 图片所在文件夹路径
    - target_width: 目标宽度（像素，0=按高度等比缩放）
    - target_height: 目标高度（像素，0=按宽度等比缩放）
    - suffix: 输出文件名后缀（如 "_thumb"），不填则覆盖
    - output_format: 导出格式 "jpg" 或 "png"
    - quality: JPG 质量 1-12，默认 8
    用户说"把文件夹里所有图片缩放到 800 宽""批量处理这批图"时使用。
    """
    if not os.path.isdir(folder):
        return {"result": "error", "error": f"文件夹不存在: {folder}"}
    if target_width <= 0 and target_height <= 0:
        return {"result": "error", "error": "至少指定 target_width 或 target_height"}

    exts = {'.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.psd'}
    files = [f for f in os.listdir(folder)
             if os.path.splitext(f)[1].lower() in exts]
    if not files:
        return {"result": "error", "error": "文件夹中没有图片文件"}

    success = []
    failed = []
    for fname in files:
        src = os.path.join(folder, fname)
        base, _ = os.path.splitext(fname)
        out_name = f"{base}{suffix}.{output_format}"
        out_path = os.path.join(folder, out_name)

        w = target_width if target_width > 0 else 0
        h = target_height if target_height > 0 else 0
        js = f"""
            try {{
                var f = new File('{src.replace(chr(92), chr(47))}');
                var d = app.open(f);
                var ppi = d.resolution;
                var curW = d.width.value * ppi;
                var curH = d.height.value * ppi;
                var tw = {w};
                var th = {h};
                if (tw <= 0) tw = Math.round(curW * (th / curH));
                if (th <= 0) th = Math.round(curH * (tw / curW));
                d.resizeImage(tw / ppi, th / ppi, ppi);
                var opts;
                if ('{output_format}' === 'png') {{
                    opts = new PNGSaveOptions();
                    opts.compression = 6;
                }} else {{
                    opts = new JPEGSaveOptions();
                    opts.quality = {quality};
                }}
                d.saveAs(new File('{out_path.replace(chr(92), chr(47))}'), opts, true);
                d.close(SaveOptions.DONOTSAVECHANGES);
                'OK|' + '{fname}' + '|' + tw + 'x' + th;
            }} catch(e) {{
                'FAIL|' + '{fname}' + '|' + e.toString();
            }}
        """
        try:
            r = _ps_js(js, max_retries=20, retry_sleep=0.3).strip()
            if r.startswith('OK'):
                success.append(r.split('|')[1])
            else:
                failed.append(r)
        except Exception as e:
            failed.append(f"FAIL|{fname}|{e}")

    return {
        "result": "ok",
        "total": len(files),
        "success": len(success),
        "failed": len(failed),
        "successFiles": success,
        "failedFiles": failed,
        "message": f"处理完成: {len(success)}/{len(files)} 成功",
    }


@mcp.tool()
def ps_batch_watermark(folder: str, text: str = "WATERMARK", opacity: int = 30,
                        size: float = 72, r: int = 255, g: int = 255, b: int = 255,
                        output_folder: str = "") -> dict:
    """
    批量给文件夹中所有图片加水印文字。
    自动打开 → 加文字水印 → 导出。
    - folder: 图片所在文件夹路径
    - text: 水印文字内容
    - opacity: 不透明度 0-100，默认 30
    - size: 字号，默认 72
    - r, g, b: 水印颜色，默认白色
    - output_folder: 输出文件夹（不填则原地加 _wm 后缀）
    用户说"给这批图加上水印""批量打水印"时使用。
    """
    if not os.path.isdir(folder):
        return {"result": "error", "error": f"文件夹不存在: {folder}"}
    out_dir = output_folder if output_folder else folder
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    exts = {'.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff'}
    files = [f for f in os.listdir(folder)
             if os.path.splitext(f)[1].lower() in exts]
    if not files:
        return {"result": "error", "error": "文件夹中没有图片文件"}

    success = []
    failed = []
    for fname in files:
        src = os.path.join(folder, fname)
        base, ext = os.path.splitext(fname)
        out_path = os.path.join(out_dir, f"{base}_wm{ext}")

        js = f"""
            try {{
                var d = app.open(new File('{src.replace(chr(92), chr(47))}'));
                var ppi = d.resolution;
                var tf = d.artLayers.add();
                tf.kind = LayerKind.TEXT;
                tf.textItem.contents = {json.dumps(text)};
                tf.textItem.size = {size};
                tf.opacity = {opacity};
                var c = new SolidColor();
                c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
                tf.textItem.color = c;
                // Center the text
                var docW = d.width.value;
                var docH = d.height.value;
                tf.textItem.position = [docW / 2 - 2, docH / 2 - 2];
                // Flatten and save
                d.flatten();
                var opts = new JPEGSaveOptions();
                opts.quality = 10;
                d.saveAs(new File('{out_path.replace(chr(92), chr(47))}'), opts, true);
                d.close(SaveOptions.DONOTSAVECHANGES);
                'OK|{fname}';
            }} catch(e) {{
                'FAIL|{fname}|' + e.toString();
            }}
        """
        try:
            r = _ps_js(js, max_retries=20, retry_sleep=0.3).strip()
            if r.startswith('OK'):
                success.append(r.split('|')[1])
            else:
                failed.append(r)
        except Exception as e:
            failed.append(f"FAIL|{fname}|{e}")

    return {
        "result": "ok",
        "total": len(files),
        "success": len(success),
        "failed": len(failed),
        "successFiles": success,
        "message": f"水印完成: {len(success)}/{len(files)} 成功",
    }


@mcp.tool()
def ps_social_media_kit(folder: str, platforms: str = "instagram_post,instagram_story,facebook_cover,wechat_moment") -> dict:
    """
    将当前 PS 文档一键导出为多个社交媒体平台的适配尺寸。
    自动为每个平台创建对应尺寸的文档，复制内容，调整，导出。
    - folder: 输出文件夹路径
    - platforms: 平台列表（逗号分隔），可选值：
      instagram_post(1080x1080), instagram_story(1080x1920),
      facebook_cover(1640x856), wechat_moment(1080x1080),
      twitter_header(1500x500), youtube_thumbnail(1280x720),
      linkedin_banner(1584x396)
    用户说"导出社媒套图""生成各平台尺寸""一键多尺寸导出"时使用。
    """
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)

    size_map = {
        "instagram_post": (1080, 1080, "ins_post"),
        "instagram_story": (1080, 1920, "ins_story"),
        "facebook_cover": (1640, 856, "fb_cover"),
        "wechat_moment": (1080, 1080, "wx_moment"),
        "twitter_header": (1500, 500, "tw_header"),
        "youtube_thumbnail": (1280, 720, "yt_thumb"),
        "linkedin_banner": (1584, 396, "in_banner"),
    }
    plat_list = [p.strip() for p in platforms.split(',') if p.strip()]
    results = []
    for plat in plat_list:
        if plat not in size_map:
            results.append({"platform": plat, "status": "unknown"})
            continue
        tw, th, prefix = size_map[plat]
        out_path = os.path.join(folder, f"{prefix}.jpg")
        js = f"""
            try {{
                if (app.documents.length === 0) {{ 'FAIL|no source document'; }}
                else {{
                    // Duplicate current doc to new size
                    var src = app.activeDocument;
                    var ppi = src.resolution;
                    // Resize source image to target
                    src.resizeImage({tw} / ppi, {th} / ppi, ppi, ResampleMethod.BICUBIC);
                    var opts = new JPEGSaveOptions();
                    opts.quality = 10;
                    src.saveAs(new File('{out_path.replace(chr(92), chr(47))}'), opts, true);
                    // Undo the resize to restore original
                    src.activeHistoryState = src.historyStates[src.historyStates.length - 1];
                    // Find a state before resize (the second-to-last might be resize, go back further)
                    // Actually, let's just undo
                    try {{ app.executeAction(app.charIDToTypeID('undo'), undefined, DialogModes.NO); }} catch(e) {{}}
                    'OK|{plat}|{tw}x{th}|{prefix}';
                }}
            }} catch(e) {{
                'FAIL|{plat}|' + e.toString();
            }}
        """
        try:
            r = _ps_js(js, max_retries=15, retry_sleep=0.3).strip()
            parts = r.split('|')
            if parts[0] == 'OK':
                results.append({"platform": plat, "status": "ok",
                                "size": f"{tw}x{th}", "file": f"{prefix}.jpg"})
            else:
                results.append({"platform": plat, "status": "error",
                                "error": parts[2] if len(parts) > 2 else r})
        except Exception as e:
            results.append({"platform": plat, "status": "error", "error": str(e)})

    ok_count = sum(1 for r in results if r["status"] == "ok")
    return {
        "result": "ok",
        "total": len(plat_list),
        "success": ok_count,
        "platforms": results,
        "outputFolder": folder,
        "message": f"社媒套图导出: {ok_count}/{len(plat_list)} 成功",
    }


@mcp.tool()
def ps_contact_sheet(folder: str, columns: int = 4, thumb_size: int = 200,
                      output_path: str = "", spacing: int = 10,
                      bg_r: int = 255, bg_g: int = 255, bg_b: int = 255) -> dict:
    """
    自动生成联系表（Contact Sheet）：把文件夹中的所有图片排列成网格缩略图。
    常用于预览整批素材、制作产品目录页。
    - folder: 图片文件夹路径
    - columns: 每行排列几张图，默认 4
    - thumb_size: 每张缩略图的边长（像素），默认 200
    - output_path: 输出 PNG 路径（不填则存到文件夹下 contact_sheet.png）
    - spacing: 图与图之间的间距像素，默认 10
    - bg_r/g/b: 背景色，默认白色
    用户说"做个联系表""缩略图网格""排列预览这批图"时使用。
    """
    if not os.path.isdir(folder):
        return {"result": "error", "error": f"文件夹不存在: {folder}"}
    if not output_path:
        output_path = os.path.join(folder, "contact_sheet.png")

    exts = {'.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff'}
    files = [f for f in os.listdir(folder)
             if os.path.splitext(f)[1].lower() in exts]
    if not files:
        return {"result": "error", "error": "文件夹中没有图片文件"}

    n = len(files)
    rows = math.ceil(n / columns)
    canvas_w = columns * thumb_size + (columns + 1) * spacing
    canvas_h = rows * thumb_size + (rows + 1) * spacing

    files_json = json.dumps([os.path.join(folder, f).replace('\\', '/') for f in files])
    js = f"""
        try {{
            var files = {files_json};
            var cols = {columns};
            var thumb = {thumb_size};
            var gap = {spacing};
            var rows = {rows};
            var cw = {canvas_w};
            var ch = {canvas_h};

            // Create canvas
            var g = _gcd(cw, ch);
            var doc = app.documents.add(cw / g, ch / g, g, 'ContactSheet');
            var ppi = doc.resolution;

            // Fill background
            var bgc = new SolidColor();
            bgc.rgb.red = {bg_r}; bgc.rgb.green = {bg_g}; bgc.rgb.blue = {bg_b};
            doc.selection.selectAll();
            doc.selection.fill(bgc);
            doc.selection.deselect();

            function _gcd(a, b) {{ a = Math.abs(a); b = Math.abs(b); while(b) {{ var t = b; b = a % b; a = t; }} return a || 1; }}

            // Place each image
            for (var i = 0; i < files.length && i < cols * rows; i++) {{
                var col = i % cols;
                var row = Math.floor(i / cols);
                var px = gap + col * (thumb + gap);
                var py = gap + row * (thumb + gap);

                // Open image, duplicate layer to contact sheet
                var imgDoc = app.open(new File(files[i]));
                var imgW = imgDoc.width.value * imgDoc.resolution;
                var imgH = imgDoc.height.value * imgDoc.resolution;
                // Fit into thumb x thumb (contain)
                var scale = Math.min(thumb / imgW, thumb / imgH);
                var newW = imgW * scale;
                var newH = imgH * scale;
                imgDoc.resizeImage(newW / imgDoc.resolution, newH / imgDoc.resolution, imgDoc.resolution);
                imgDoc.selection.selectAll();
                imgDoc.selection.duplicate(doc.artLayers[0]);
                imgDoc.close(SaveOptions.DONOTSAVECHANGES);

                // Move the pasted layer
                var pasted = doc.activeLayer;
                pasted.translate((px / ppi), (py / ppi));
            }}

            // Save
            var opts = new PNGSaveOptions();
            opts.compression = 6;
            doc.saveAs(new File('{output_path.replace(chr(92), chr(47))}'), opts, true);
            'OK|' + files.length + '|' + cols + 'x' + rows + '|' + cw + 'x' + ch;
        }} catch(e) {{
            'FAIL|' + e.toString();
        }}
    """
    result = _ps_js(js, max_retries=20, retry_sleep=0.3).strip()
    parts = result.split('|')
    if parts[0] == 'OK':
        return {
            "result": "ok",
            "images": int(parts[1]),
            "grid": parts[2],
            "canvas": parts[3],
            "path": output_path,
            "message": f"联系表已生成: {parts[2]} 网格, {parts[3]}px",
        }
    return {"result": "error", "error": parts[1] if len(parts) > 1 else result}


@mcp.tool()
def ps_batch_rename_layers(prefix: str = "Layer", start: int = 1,
                             by_type: str = "") -> dict:
    """
    批量重命名所有图层（或特定类型图层）。
    - prefix: 名称前缀，如 "图层"
    - start: 起始编号，默认 1
    - by_type: 只重命名特定类型（"text"=仅文字图层, "image"=仅图像图层, ""=全部）
    用户说"批量重命名图层""把所有图层名改成统一格式"时使用。
    """
    type_filter = ""
    if by_type == "text":
        type_filter = "if (d.artLayers[i].kind === LayerKind.TEXT)"
    elif by_type == "image":
        type_filter = "if (d.artLayers[i].kind === LayerKind.NORMAL)"

    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var count = 0;
            var idx = {start};
            for (var i = 0; i < d.artLayers.length; i++) {{
                var lyr = d.artLayers[i];
                {type_filter if type_filter else 'if (true)'} {{
                    lyr.name = {json.dumps(prefix)} + ' ' + idx;
                    idx++;
                    count++;
                }}
            }}
            count.toString();
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    count = int(result)
    return {"result": "ok", "renamed": count, "prefix": prefix,
            "message": f"已重命名 {count} 个图层为 '{prefix} N' 格式"}


@mcp.tool()
def ps_auto_gradient_background(r1: int = 0, g1: int = 0, b1: int = 0,
                                 r2: int = 255, g2: int = 255, b2: int = 255,
                                 angle: float = 90) -> dict:
    """
    在当前文档背景图层上创建渐变填充。
    从颜色1渐变到颜色2，可指定角度。
    - r1,g1,b1: 起始色 RGB
    - r2,g2,b2: 终止色 RGB
    - angle: 渐变角度 0-360，默认 90（从下到上）
    用户说"做一个蓝色到白色的渐变背景""渐变底图"时使用。
    """
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            // Create gradient
            var desc = new ActionDescriptor();
            var gradDesc = new ActionDescriptor();
            var colors = new ActionList();

            // Color 1
            var c1 = new ActionDescriptor();
            c1.putEnumerated(charIDToTypeID('Clr '), charIDToTypeID('ClrS'), charIDToTypeID('RGBC'));
            c1.putDouble(charIDToTypeID('Rd  '), {r1});
            c1.putDouble(charIDToTypeID('Grn '), {g1});
            c1.putDouble(charIDToTypeID('Bl  '), {b1});

            // Color 2
            var c2 = new ActionDescriptor();
            c2.putEnumerated(charIDToTypeID('Clr '), charIDToTypeID('ClrS'), charIDToTypeID('RGBC'));
            c2.putDouble(charIDToTypeID('Rd  '), {r2});
            c2.putDouble(charIDToTypeID('Grn '), {g2});
            c2.putDouble(charIDToTypeID('Bl  '), {b2});

            colors.putObject(charIDToTypeID('Clrz'), c1);
            colors.putObject(charIDToTypeID('Clrz'), c2);

            gradDesc.putEnumerated(charIDToTypeID('Type'), charIDToTypeID('GrdT'), charIDToTypeID('Lnr '));
            gradDesc.putDouble(charIDToTypeID('Angl'), {angle});
            gradDesc.putObject(charIDToTypeID('Clrz'), c1);
            gradDesc.putList(charIDToTypeID('Clrs'), colors);

            desc.putObject(charIDToTypeID('Usng'), charIDToTypeID('GrFl'), gradDesc);
            desc.putEnumerated(charIDToTypeID('Opct'), charIDToTypeID('Opct'), charIDToTypeID('100'));
            desc.putBoolean(charIDToTypeID('AlPh'), true);

            // Select all and fill
            d.selection.selectAll();
            executeAction(charIDToTypeID('Fl  '), desc, DialogModes.NO);
            d.selection.deselect();
            'OK|{r1},{g1},{b1}->{r2},{g2},{b2}|{angle}';
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    return {"result": "ok", "from": parts[0], "angle": float(parts[1]) if len(parts) > 1 else angle}


@mcp.tool()
def ps_smart_replace_text_csv(csv_text: str, template_field: str = "TITLE") -> dict:
    """
    根据CSV数据批量生成多个版本的PS文档。
    在模板文档中找到包含指定占位符的文字图层，替换为CSV中每行的值，
    每行生成一个独立的导出图片。
    - csv_text: CSV文本（第一行为表头，后续每行数据，逗号分隔）
    - template_field: 模板中要替换的占位符字段名（对应CSV表头列名）
    用户说"根据数据批量生成100张海报""按CSV批量替换文字导出"时使用。

    示例CSV：
    TITLE,PRICE
    夏日特惠,99元
    秋季新品,128元
    """
    lines = csv_text.strip().split('\n')
    if len(lines) < 2:
        return {"result": "error", "error": "CSV至少需要表头+1行数据"}

    header = [h.strip() for h in lines[0].split(',')]
    if template_field not in header:
        return {"result": "error", "error": f"CSV表头中未找到字段 '{template_field}'，可用字段: {header}"}
    col_idx = header.index(template_field)

    data_rows = []
    for line in lines[1:]:
        cols = [c.strip() for c in line.split(',')]
        if col_idx < len(cols):
            data_rows.append(cols[col_idx])

    if not data_rows:
        return {"result": "error", "error": "CSV中没有数据行"}

    out_dir = os.path.join(os.path.expanduser("~"), "Documents", "batch_export")
    os.makedirs(out_dir, exist_ok=True)

    results = []
    for i, val in enumerate(data_rows):
        out_path = os.path.join(out_dir, f"export_{i+1}_{val}.png")
        js = f"""
            try {{
                if (app.documents.length === 0) {{ 'FAIL|no document'; }}
                else {{
                    var d = app.activeDocument;
                    var changed = false;
                    for (var j = 0; j < d.artLayers.length; j++) {{
                        var lyr = d.artLayers[j];
                        if (lyr.kind === LayerKind.TEXT) {{
                            if (lyr.textItem.contents.indexOf({json.dumps(template_field)}) >= 0) {{
                                lyr.textItem.contents = {json.dumps(val)};
                                changed = true;
                            }}
                        }}
                    }}
                    if (changed) {{
                        var opts = new PNGSaveOptions();
                        opts.compression = 6;
                        d.saveAs(new File('{out_path.replace(chr(92), chr(47))}'), opts, true);
                    }}
                    // Undo text replacement to restore template
                    try {{ app.executeAction(app.charIDToTypeID('undo'), undefined, DialogModes.NO); }} catch(e) {{}}
                    'OK|{val}|{out_path}';
                }}
            }} catch(e) {{
                'FAIL|{val}|' + e.toString();
            }}
        """
        try:
            r = _ps_js(js, max_retries=15, retry_sleep=0.3).strip()
            parts = r.split('|')
            if parts[0] == 'OK':
                results.append({"text": parts[1], "path": parts[2], "status": "ok"})
            else:
                results.append({"text": val, "status": "error",
                                "error": parts[2] if len(parts) > 2 else r})
        except Exception as e:
            results.append({"text": val, "status": "error", "error": str(e)})

    ok_count = sum(1 for r in results if r["status"] == "ok")
    return {
        "result": "ok",
        "total": len(data_rows),
        "success": ok_count,
        "outputFolder": out_dir,
        "results": results,
        "message": f"批量生成: {ok_count}/{len(data_rows)} 成功，输出到 {out_dir}",
    }


@mcp.tool()
def ps_extract_color_palette(num_colors: int = 6) -> dict:
    """
    从当前文档中提取主色调色盘。
    通过将图像缩小为缩略图后采样像素，返回最主要的几种颜色及其RGB/Hex值。
    - num_colors: 提取颜色数量，默认 6
    用户说"提取这张图的主色调""分析配色""生成调色板"时使用。
    """
    num_colors = max(2, min(12, int(num_colors)))

    # Export a small temp BMP, then analyze with Python's built-in approach
    import tempfile
    tmp_path = os.path.join(tempfile.gettempdir(), "palette_sample.bmp")

    # Step 1: Export a small flattened thumbnail from PS as BMP
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var dup = d.duplicate('palette_tmp');
            dup.flatten();
            var ppi = dup.resolution;
            var tw = 48;
            var th = Math.round(tw * (dup.height.value / dup.width.value));
            if (th < 1) th = 1;
            dup.resizeImage(tw / ppi, th / ppi, ppi);
            // Save as BMP (simplest uncompressed format)
            dup.saveAs(new File('{tmp_path.replace(chr(92), chr(47))}'), new BMPSaveOptions(), true);
            dup.close(SaveOptions.DONOTSAVECHANGES);
            'OK|' + tw + 'x' + th;
        }}
    """
    result = _ps_js(js, max_retries=15, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if not result.startswith('OK'):
        return {"result": "error", "error": f"export failed: {result}"}

    # Step 2: Read the BMP and analyze pixel colors in Python
    try:
        import struct

        with open(tmp_path, 'rb') as f:
            data = f.read()

        if data[:2] != b'BM':
            return {"result": "error", "error": "not a valid BMP file"}

        data_offset = struct.unpack('<I', data[10:14])[0]
        width = struct.unpack('<i', data[18:22])[0]
        height_raw = struct.unpack('<i', data[22:26])[0]
        bpp = struct.unpack('<H', data[28:30])[0]
        height = abs(height_raw)

        bytes_per_pixel = bpp // 8
        if bytes_per_pixel < 3:
            # Fallback: sample via JS if BMP format is unusual
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return _palette_via_js(num_colors)

        row_size = (width * bytes_per_pixel + 3) & ~3

        hist = {}
        for y in range(height):
            row_start = data_offset + y * row_size
            for x in range(width):
                px = row_start + x * bytes_per_pixel
                if px + 2 < len(data):
                    # BMP stores BGR
                    b_val = data[px]
                    g_val = data[px + 1]
                    r_val = data[px + 2]
                    qr = (r_val // 32) * 32
                    qg = (g_val // 32) * 32
                    qb = (b_val // 32) * 32
                    key = (qr, qg, qb)
                    hist[key] = hist.get(key, 0) + 1

        # Clean up temp file
        try:
            os.remove(tmp_path)
        except Exception:
            pass

        if not hist:
            return {"result": "error", "error": "no pixels found in BMP"}

        sorted_colors = sorted(hist.items(), key=lambda x: x[1], reverse=True)
        palette = []
        for (r, g, b), freq in sorted_colors[:num_colors]:
            palette.append({
                "r": r, "g": g, "b": b,
                "hex": f"#{r:02X}{g:02X}{b:02X}",
                "frequency": freq,
            })

        return {
            "result": "ok",
            "count": len(palette),
            "palette": palette,
            "message": f"提取了 {len(palette)} 个主色调",
        }
    except Exception as e:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        return {"result": "error", "error": f"color analysis failed: {e}"}


def _palette_via_js(num_colors: int) -> dict:
    """Fallback: sample pixel colors via DoJavaScript using histogram channels."""
    js = """
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            // Use histogram to find dominant brightness ranges per channel
            var rHist = d.histogram[0];  // Red channel
            var gHist = d.histogram[1];  // Green channel
            var bHist = d.histogram[2];  // Blue channel

            // Find top N brightness values per channel
            var rPeaks = [], gPeaks = [], bPeaks = [];
            for (var i = 0; i < 256; i++) {
                rPeaks.push([i, rHist[i]]);
                gPeaks.push([i, gHist[i]]);
                bPeaks.push([i, bHist[i]]);
            }
            rPeaks.sort(function(a,b) { return b[1] - a[1]; });
            gPeaks.sort(function(a,b) { return b[1] - a[1]; });
            bPeaks.sort(function(a,b) { return b[1] - a[1]; });

            // Take top peaks and combine
            var n = Math.min(6, arguments[0] || 6);
            var result = [];
            for (var j = 0; j < n; j++) {
                var r = rPeaks[j] ? rPeaks[j][0] : 0;
                var g = gPeaks[j] ? gPeaks[j][0] : 0;
                var b = bPeaks[j] ? bPeaks[j][0] : 0;
                result.push(r + ',' + g + ',' + b);
            }
            result.join(';');
        }
    """
    try:
        result = _ps_js(js, max_retries=10, retry_sleep=0.3).strip()
        if result == 'NO_DOCS':
            return {"result": "error", "error": "no document"}
        palette = []
        for entry in result.split(';'):
            if not entry.strip():
                continue
            rgb = entry.split(',')
            if len(rgb) == 3:
                r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
                palette.append({
                    "r": r, "g": g, "b": b,
                    "hex": f"#{r:02X}{g:02X}{b:02X}",
                })
        return {
            "result": "ok",
            "count": len(palette),
            "palette": palette[:num_colors],
            "message": f"提取了 {len(palette)} 个主色调",
        }
    except Exception as e:
        return {"result": "error", "error": f"palette fallback failed: {e}"}


@mcp.tool()
def ps_auto_layout_cards(rows: int = 3, cols: int = 3, card_w: int = 400,
                          card_h: int = 600, gap: int = 20,
                          bg_r: int = 240, bg_g: int = 240, bg_b: int = 240) -> dict:
    """
    自动批量排版卡片（如名片、优惠券、产品卡）。
    在一张大画布上自动排列 R行×C列 个卡片位，每个位可后续填充内容。
    - rows: 行数
    - cols: 列数
    - card_w, card_h: 单张卡片宽高（像素）
    - gap: 卡片间距
    - bg_r/g/b: 背景色
    用户说"排3x3的名片版""批量排版卡片"时使用。
    """
    total = rows * cols
    canvas_w = cols * card_w + (cols + 1) * gap
    canvas_h = rows * card_h + (rows + 1) * gap

    js = f"""
        try {{
            var g = 1; // gcd trick fallback - use ppi directly
            var cw = {canvas_w};
            var ch = {canvas_h};
            // Use simple approach: create doc at 100 ppi so 1 inch = 100px
            var ppi = 100;
            var doc = app.documents.add(cw / ppi, ch / ppi, ppi, 'CardLayout');

            // Fill background
            var bgc = new SolidColor();
            bgc.rgb.red = {bg_r}; bgc.rgb.green = {bg_g}; bgc.rgb.blue = {bg_b};
            doc.selection.selectAll();
            doc.selection.fill(bgc);
            doc.selection.deselect();

            // Draw card placeholders
            var cardW = {card_w};
            var cardH = {card_h};
            var gap = {gap};
            var rows = {rows};
            var cols = {cols};
            var cardColor = new SolidColor();
            cardColor.rgb.red = 255; cardColor.rgb.green = 255; cardColor.rgb.blue = 255;

            for (var r = 0; r < rows; r++) {{
                for (var c = 0; c < cols; c++) {{
                    var x1 = (gap + c * (cardW + gap)) / ppi;
                    var y1 = (gap + r * (cardH + gap)) / ppi;
                    var x2 = x1 + cardW / ppi;
                    var y2 = y1 + cardH / ppi;
                    doc.selection.select([[x1, y1], [x2, y1], [x2, y2], [x1, y2]]);
                    doc.selection.fill(cardColor);
                    // Add border
                    doc.selection.stroke(cardColor, 1);
                }}
            }}
            doc.selection.deselect();
            'OK|' + rows + 'x' + cols + '|' + cw + 'x' + ch + '|' + {total};
        }} catch(e) {{
            'FAIL|' + e.toString();
        }}
    """
    result = _ps_js(js, max_retries=15, retry_sleep=0.3).strip()
    parts = result.split('|')
    if parts[0] == 'OK':
        return {
            "result": "ok",
            "grid": parts[1],
            "canvas": parts[2],
            "totalCards": int(parts[3]),
            "cardSize": f"{card_w}x{card_h}px",
            "message": f"已创建 {parts[1]} 网格画布 ({parts[2]}px)，共 {parts[3]} 个卡片位",
        }
    return {"result": "error", "error": parts[1] if len(parts) > 1 else result}


@mcp.tool()
def ps_create_gif_from_layers(folder: str, delay: float = 0.5) -> dict:
    """
    将当前 PS 文档中的每个图层导出为帧，然后生成 GIF 动画。
    - folder: 输出文件夹（GIF 保存路径）
    - delay: 每帧延迟时间（秒），默认 0.5
    用户说"把图层做成GIF""图层转动画"时使用。
    """
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)
    gif_path = os.path.join(folder, "animation.gif")

    js = f"""
        try {{
            if (app.documents.length === 0) {{ 'FAIL|no document'; }}
            else {{
                var d = app.activeDocument;
                var layers = d.artLayers;
                var frameCount = layers.length;
                var outDir = '{folder.replace(chr(92), chr(47))}';
                var delay = {delay};

                // Export each layer as a frame
                for (var i = 0; i < frameCount; i++) {{
                    // Hide all layers
                    for (var j = 0; j < frameCount; j++) {{
                        layers[j].visible = (j === i);
                    }}
                    var frameFile = new File(outDir + '/frame_' + (i+1) + '.png');
                    var opts = new PNGSaveOptions();
                    opts.compression = 6;
                    d.saveAs(frameFile, opts, true);
                }}

                // Restore all visible
                for (var k = 0; k < frameCount; k++) {{
                    layers[k].visible = true;
                }}

                'OK|' + frameCount + ' frames exported';
            }}
        }} catch(e) {{
            'FAIL|' + e.toString();
        }}
    """
    result = _ps_js(js, max_retries=15, retry_sleep=0.3).strip()
    parts = result.split('|')
    if parts[0] == 'OK':
        return {
            "result": "ok",
            "frames": parts[1],
            "frameFolder": folder,
            "delay": delay,
            "message": f"已导出 {parts[1]}，各帧 PNG 存在 {folder}。可用 PS → 文件 → 导出 → 存储为Web格式 生成最终GIF。",
        }
    return {"result": "error", "error": parts[1] if len(parts) > 1 else result}


# ============================================================
# ILLUSTRATOR TOOLS
# ============================================================

@mcp.tool()
def ai_create_document(width: float = 612, height: float = 792) -> dict:
    """
    在 Illustrator 中新建一个文档，单位为点（pt，72pt = 1 英寸）。
    - width: 宽度（点），默认 612（A4 宽）
    - height: 高度（点），默认 792（A4 高）
    """
    app = _get_ai()
    doc = app.Documents.Add(Width=width, Height=height)
    return {"result": "ok", "name": doc.Name, "width": width, "height": height, "unit": "points"}


@mcp.tool()
def ai_get_active_info() -> dict:
    """获取 Illustrator 当前活动文档信息：名称、尺寸、路径数、文字框数、打开文档数。"""
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "ok", "active": False, "open": 0}
    d = app.ActiveDocument
    return {
        "result": "ok",
        "name": d.Name,
        "width": d.Width,
        "height": d.Height,
        "paths": d.PathItems.Count,
        "textFrames": d.TextFrames.Count,
        "open": app.Documents.Count,
    }


@mcp.tool()
def ai_list_documents() -> dict:
    """列出 Illustrator 当前打开的所有文档名称。"""
    app = _get_ai()
    names = [app.Documents.Item(i).Name for i in range(1, app.Documents.Count + 1)]
    return {"result": "ok", "count": len(names), "documents": [{"name": n} for n in names]}


@mcp.tool()
def ai_close_document() -> dict:
    """关闭 Illustrator 当前活动文档（不保存）。"""
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "ok", "closed": False}
    name = app.ActiveDocument.Name
    app.ActiveDocument.Close(2)  # aiDoNotSaveChanges
    return {"result": "ok", "closed": True, "name": name}


def _ai_rgb(app, r, g, b):
    """Create an RGBColor object for Illustrator."""
    c = app.ActiveDocument.Swatches.Add()
    c.Color.Red = r
    c.Color.Green = g
    c.Color.Blue = b
    return c.Color


@mcp.tool()
def ai_add_rectangle(x: float = 100, y: float = 400, w: float = 200, h: float = 150,
                     fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """
    在当前 Illustrator 文档中画一个矩形。
    - x, y: 左下角坐标（点）
    - w, h: 宽高（点）
    - fr, fg, fb: 填充色 RGB（0-255），不填则无填充
    """
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    rect = app.ActiveDocument.PathItems.Rectangle(y, x, w, h)  # top, left, width, height
    if fr >= 0 and fg >= 0 and fb >= 0:
        rect.Filled = True
        try:
            rgb = app.ActiveDocument.Swatches.Add()
            rgb.Color.Red = fr
            rgb.Color.Green = fg
            rgb.Color.Blue = fb
            rect.FillColor = rgb.Color
        except Exception:
            pass
    return {"result": "ok", "name": rect.Name, "x": x, "y": y, "width": w, "height": h}


@mcp.tool()
def ai_add_ellipse(x: float = 200, y: float = 500, w: float = 150, h: float = 100,
                   fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """
    在当前 Illustrator 文档中画一个椭圆。
    - x, y: 左上角坐标（点）
    - w, h: 宽高（点）
    - fr, fg, fb: 填充色 RGB，可选
    """
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    ell = app.ActiveDocument.PathItems.Ellipse(y, x, w, h)
    if fr >= 0 and fg >= 0 and fb >= 0:
        ell.Filled = True
        try:
            rgb = app.ActiveDocument.Swatches.Add()
            rgb.Color.Red = fr
            rgb.Color.Green = fg
            rgb.Color.Blue = fb
            ell.FillColor = rgb.Color
        except Exception:
            pass
    return {"result": "ok", "name": ell.Name, "x": x, "y": y, "width": w, "height": h}


@mcp.tool()
def ai_add_polygon(x: float = 300, y: float = 400, radius: float = 100, sides: int = 6,
                   fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """
    在当前 Illustrator 文档中画一个正多边形。
    - x, y: 中心点坐标（点）
    - radius: 半径（点）
    - sides: 边数（默认 6）
    - fr, fg, fb: 填充色 RGB，可选
    """
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    poly = app.ActiveDocument.PathItems.Polygon(y, x, radius, sides)
    if fr >= 0 and fg >= 0 and fb >= 0:
        poly.Filled = True
        try:
            rgb = app.ActiveDocument.Swatches.Add()
            rgb.Color.Red = fr
            rgb.Color.Green = fg
            rgb.Color.Blue = fb
            poly.FillColor = rgb.Color
        except Exception:
            pass
    return {"result": "ok", "name": poly.Name, "x": x, "y": y, "radius": radius, "sides": sides}


@mcp.tool()
def ai_add_text(text: str, size: float = 24, x: float = 100, y: float = 300,
                fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """
    在当前 Illustrator 文档中添加文字。
    - text: 文字内容
    - size: 字号（点），默认 24
    - x, y: 位置（点）
    - fr, fg, fb: 文字颜色 RGB，可选
    """
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    tf = app.ActiveDocument.TextFrames.Add()
    tf.Contents = text
    try:
        tf.TextRange.CharacterAttributes.Size = size
    except Exception:
        pass
    try:
        tf.Top = y
        tf.Left = x
    except Exception:
        pass
    try:
        if fr >= 0 and fg >= 0 and fb >= 0:
            rgb = app.ActiveDocument.Swatches.Add()
            rgb.Color.Red = fr
            rgb.Color.Green = fg
            rgb.Color.Blue = fb
            tf.TextRange.CharacterAttributes.FillColor = rgb.Color
    except Exception:
        pass
    return {"result": "ok", "text": text, "size": size, "x": x, "y": y}


@mcp.tool()
def ai_save_as_ai(path: str = "") -> dict:
    """将当前 Illustrator 文档保存为 .ai 格式。path 可选。"""
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.ai")
    app.ActiveDocument.SaveAs(path)
    return {"result": "ok", "path": path}


@mcp.tool()
def ai_export_svg(path: str = "") -> dict:
    """将当前 Illustrator 文档导出为 SVG 矢量图。path 可选。"""
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.svg")
    opts = win32com.client.Dispatch("Illustrator.SVGExportOptions")
    opts.FontType = 1  # aiSVGFontOutline
    app.ActiveDocument.Export(path, 10, opts)  # 10 = aiSVG
    return {"result": "ok", "path": path}


@mcp.tool()
def ai_export_png(path: str = "") -> dict:
    """将当前 Illustrator 文档导出为 PNG 位图（透明背景）。path 可选。"""
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.png")
    opts = win32com.client.Dispatch("Illustrator.PNGExportOptions")
    opts.AntiAliasing = True
    opts.Transparency = True
    app.ActiveDocument.Export(path, 4, opts)  # 4 = aiPNG
    return {"result": "ok", "path": path}


# ============================================================
# Advanced Batch Tools — 体力活自动化
# ============================================================

@mcp.tool()
def ps_export_all_layers_to_png(output_folder: str = "") -> dict:
    """
    一键导出当前文档中所有图层为单独的 PNG 文件。
    自动逐个显示/隐藏图层并导出，文件名用图层名命名。
    手动操作需要逐个隐藏其他图层→显示当前图层→导出→重复几十次。
    - output_folder: 输出文件夹路径，默认桌面新建 layers_export 文件夹
    用户说"导出所有图层""每层导出一张图""图层拆分导出"时使用。
    """
    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop", "layers_export")
    output_folder = output_folder.replace("\\", "/")
    # Create folder
    js = """
        var outFolder = new Folder('""" + output_folder + """');
        if (!outFolder.exists) outFolder.create();
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var layers = [];
            for (var i = 0; i < d.artLayers.length; i++) {
                layers.push(d.artLayers[i]);
            }
            if (layers.length === 0) { 'NONE'; }
            else {
                // Save current visibility states
                var visStates = [];
                for (var j = 0; j < layers.length; j++) {
                    visStates.push(layers[j].visible);
                }
                // Hide all layers
                for (var k = 0; k < layers.length; k++) {
                    layers[k].visible = false;
                }
                var exported = [];
                for (var m = 0; m < layers.length; m++) {
                    layers[m].visible = true;
                    var safeName = layers[m].name.replace(/[^a-zA-Z0-9_\\-]/g, '_');
                    if (safeName.length === 0) safeName = 'layer_' + m;
                    var fname = '""" + output_folder + """/' + safeName + '_' + (m+1) + '.png';
                    var saveFile = new File(fname);
                    var pngOpts = new PNGSaveOptions();
                    pngOpts.compression = 6;
                    pngOpts.interlaced = false;
                    d.saveAs(saveFile, pngOpts, true);
                    exported.push(safeName + '_' + (m+1) + '.png');
                    layers[m].visible = false;
                }
                // Restore visibility
                for (var n = 0; n < layers.length; n++) {
                    layers[n].visible = visStates[n];
                }
                exported.length + '|' + exported.join(';');
            }
        }
    """
    result = _ps_js(js, max_retries=30, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'NONE':
        return {"result": "error", "error": "no layers"}
    parts = result.split('|')
    count = int(parts[0])
    files = parts[1].split(';') if len(parts) > 1 else []
    return {"result": "ok", "exported": count, "files": files, "folder": output_folder,
            "message": f"已导出 {count} 个图层为单独 PNG 到 {output_folder}"}


@mcp.tool()
def ps_create_spritesheet(output_folder: str = "", columns: int = 0,
                          cell_w: int = 0, cell_h: int = 0) -> dict:
    """
    将当前文档所有图层自动拼成一张精灵表（Sprite Sheet）。
    每个图层是一个精灵帧，自动排列成网格。
    游戏开发常用，手动拼图极费时间。
    - output_folder: 输出路径（不含文件名），默认桌面
    - columns: 列数，0=自动计算（接近正方形）
    - cell_w, cell_h: 每格宽高，0=使用图层最大尺寸
    用户说"生成精灵表""sprite sheet""拼图集""帧动画图集"时使用。
    """
    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop")
    output_folder = output_folder.replace("\\", "/")
    js = """
        var outDir = '""" + output_folder + """';
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var layers = [];
            for (var i = 0; i < d.artLayers.length; i++) {
                layers.push(d.artLayers[i]);
            }
            if (layers.length === 0) { 'NONE'; }
            else {
                var count = layers.length;
                var cols = """ + str(columns) + """;
                if (cols <= 0) {
                    cols = Math.ceil(Math.sqrt(count));
                    if (cols < 1) cols = 1;
                }
                var rows = Math.ceil(count / cols);
                // Determine cell size: use document dimensions or provided values
                var docW = d.width.value;
                var docH = d.height.value;
                var cw = """ + str(cell_w) + """;
                var ch = """ + str(cell_h) + """;
                if (cw <= 0) cw = Math.round(docW);
                if (ch <= 0) ch = Math.round(docH);

                // Create new sprite sheet document
                var totalW = cols * cw;
                var totalH = rows * ch;
                var sheet = app.documents.add(totalW, totalH, d.resolution, 'SpriteSheet',
                                              NewDocumentMode.RGB, DocumentFill.TRANSPARENT);

                // Save visibility states
                var visStates = [];
                for (var j = 0; j < layers.length; j++) {
                    visStates.push(layers[j].visible);
                }

                // For each layer: show only it, duplicate merged to sheet
                for (var m = 0; m < count; m++) {
                    // Hide all
                    for (var k = 0; k < layers.length; k++) {
                        layers[k].visible = false;
                    }
                    layers[m].visible = true;
                    // Duplicate and flatten to get just this layer's pixels
                    var dup = d.duplicate('frame_tmp');
                    dup.flatten();
                    // Copy
                    dup.selection.selectAll();
                    try { dup.selection.copy(); } catch(e) {}
                    dup.close(SaveOptions.DONOTSAVECHANGES);
                    // Paste into sheet at correct position
                    sheet.activate();
                    var col = m % cols;
                    var row = Math.floor(m / cols);
                    // Select target region
                    var left = col * cw;
                    var top = row * ch;
                    var right = left + cw;
                    var bottom = top + ch;
                    var selRegion = [[left, top], [right, top], [right, bottom], [left, bottom]];
                    sheet.selection.select(selRegion);
                    try { sheet.paste(); } catch(e) {}
                    // Merge pasted layer down
                    try {
                        if (sheet.artLayers.length > 1) {
                            sheet.activeLayer = sheet.artLayers[sheet.artLayers.length - 1];
                            sheet.activeLayer.merge();
                        }
                    } catch(e) {}
                }
                // Restore visibility
                for (var n = 0; n < layers.length; n++) {
                    layers[n].visible = visStates[n];
                }
                // Save sprite sheet
                var saveFile = new File(outDir + '/spritesheet.png');
                var pngOpts = new PNGSaveOptions();
                pngOpts.compression = 6;
                sheet.saveAs(saveFile, pngOpts, true);
                count + '|' + cols + 'x' + rows + '|' + totalW + 'x' + totalH + '|' + outDir + '/spritesheet.png';
            }
        }
    """
    result = _ps_js(js, max_retries=40, retry_sleep=0.5).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'NONE':
        return {"result": "error", "error": "no layers"}
    parts = result.split('|')
    count = int(parts[0])
    grid = parts[1]
    dims = parts[2]
    path = parts[3]
    return {"result": "ok", "frames": count, "grid": grid, "pixel_size": dims,
            "path": path, "message": f"精灵表已生成：{count} 帧，网格 {grid}，尺寸 {dims}px"}


@mcp.tool()
def ps_distribute_layers_evenly(direction: str = "horizontal", gap: float = 0) -> dict:
    """
    将当前选中的多个图层均匀分布间距（等距排列）。
    手动要精确计算每个图层的位置并逐一移动，极其繁琐。
    - direction: "horizontal"（水平分布）或 "vertical"（垂直分布）
    - gap: 图层之间的间距（像素），0=自动均匀分布
    使用前请先选中多个图层。
    用户说"均匀分布图层""等距排列""图层间距统一"时使用。
    """
    direction = direction.lower().strip()
    if direction not in ("horizontal", "vertical", "h", "v"):
        direction = "horizontal"
    if direction in ("h", "horiz"):
        direction = "horizontal"
    if direction in ("v", "vert"):
        direction = "vertical"

    js = """
        var dir = '""" + direction + """';
        var gap = """ + str(gap) + """;
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            // Get selected layers
            var selected = [];
            var ref = new ActionReference();
            ref.putEnumerated(charIDToTypeID('Dcmn'), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
            var desc = executeActionGet(ref);
            // Count selected layers
            var selCount = 0;
            try {
                var list = desc.getList(charIDToTypeID('LyrI'));
                selCount = list.count;
            } catch(e) { selCount = 0; }
            if (selCount < 2) { 'TOO_FEW'; }
            else {
                // Get bounds of each selected layer
                var layerData = [];
                for (var i = 0; i < d.artLayers.length; i++) {
                    var lyr = d.artLayers[i];
                    // Check if this layer is selected
                    var lyrDesc = executeActionGet(_getLayerRef(lyr.itemIndex));
                    var isSelected = false;
                    try {
                        isSelected = lyrDesc.getBoolean(stringIDToTypeID('visible'));
                    } catch(e) {}
                    // Alternative: check via active layer cycling
                    layerData.push({
                        index: i,
                        bounds: lyr.bounds,
                        name: lyr.name
                    });
                }
                // Simpler approach: use activeLayer and cycling through selection
                // Actually, PS doesn't easily expose multi-selection via JS in 2020
                // Use Action Manager to get selected layer IDs
                var selIds = [];
                try {
                    var ref2 = new ActionReference();
                    ref2.putEnumerated(charIDToTypeID('Dcmn'), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                    var desc2 = executeActionGet(ref2);
                    var layerList = desc2.getList(stringIDToTypeID('targetLayers'));
                    for (var j = 0; j < layerList.count; j++) {
                        var layerRef = layerList.getReference(j);
                        var idx = layerRef.getIndex();
                        selIds.push(idx);
                    }
                } catch(e) {}

                if (selIds.length < 2) { 'TOO_FEW'; }
                else {
                    // Sort by position
                    var items = [];
                    for (var m = 0; m < selIds.length; m++) {
                        d.activeLayer = d.artLayers[selIds[m] - 1];
                        var b = d.activeLayer.bounds;
                        items.push({
                            idx: selIds[m],
                            left: b[0].value,
                            top: b[1].value,
                            right: b[2].value,
                            bottom: b[3].value,
                            w: b[2].value - b[0].value,
                            h: b[3].value - b[1].value
                        });
                    }
                    var distributed = items.length;
                    if (dir === 'horizontal') {
                        items.sort(function(a, b) { return a.left - b.left; });
                        var totalLeft = items[0].left;
                        var totalRight = items[items.length - 1].right;
                        var totalSpan = totalRight - totalLeft;
                        var totalItemW = 0;
                        for (var p = 0; p < items.length; p++) totalItemW += items[p].w;
                        var totalGap = totalSpan - totalItemW;
                        var eachGap = gap > 0 ? gap : totalGap / (items.length - 1);
                        var cursor = totalLeft;
                        for (var q = 0; q < items.length; q++) {
                            d.activeLayer = d.artLayers[items[q].idx - 1];
                            var dx = cursor - items[q].left;
                            d.activeLayer.translate(dx, 0);
                            cursor += items[q].w + eachGap;
                        }
                    } else {
                        items.sort(function(a, b) { return a.top - b.top; });
                        var totalTop = items[0].top;
                        var totalBottom = items[items.length - 1].bottom;
                        var totalSpan = totalBottom - totalTop;
                        var totalItemH = 0;
                        for (var r2 = 0; r2 < items.length; r2++) totalItemH += items[r2].h;
                        var totalGap = totalSpan - totalItemH;
                        var eachGap = gap > 0 ? gap : totalGap / (items.length - 1);
                        var cursor = totalTop;
                        for (var s = 0; s < items.length; s++) {
                            d.activeLayer = d.artLayers[items[s].idx - 1];
                            var dy = cursor - items[s].top;
                            d.activeLayer.translate(0, dy);
                            cursor += items[s].h + eachGap;
                        }
                    }
                    distributed + '|' + dir;
                }
            }
        }
        function _getLayerRef(idx) {
            var ref = new ActionReference();
            ref.putIndex(charIDToTypeID('Lyr '), idx);
            return ref;
        }
    """
    result = _ps_js(js, max_retries=20, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'TOO_FEW':
        return {"result": "error", "error": "请先选中至少 2 个图层"}
    parts = result.split('|')
    count = int(parts[0]) if parts[0].isdigit() else 0
    dir_out = parts[1] if len(parts) > 1 else direction
    return {"result": "ok", "distributed": count, "direction": dir_out,
            "message": f"已将 {count} 个图层沿{dir_out}方向均匀分布"}


@mcp.tool()
def ps_auto_trim(trim_away: str = "transparent") -> dict:
    """
    自动裁剪当前文档边缘（透明像素、白色或黑色区域）。
    手动要用裁剪工具仔细拖拽，批量处理尤其费时。
    - trim_away: 裁剪类型 "transparent"（透明）/ "white"（白色）/ "black"（黑色）
    用户说"自动裁边""去掉空白边缘""裁剪透明区域""trim"时使用。
    """
    trim_away = trim_away.lower().strip()
    if trim_away not in ("transparent", "white", "black", "top_left"):
        trim_away = "transparent"

    trim_id = {
        "transparent": "Trns",
        "white": "Wht ",
        "black": "Blck",
        "top_left": "TlLt"
    }.get(trim_away, "Trns")

    js = """
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var trimType = '""" + trim_id + """';
            var desc = new ActionDescriptor();
            desc.putEnumerated(charIDToTypeID('Tmr '), charIDToTypeID('TrmA'), charIDToTypeID(trimType));
            desc.putBoolean(charIDToTypeID('T  '), true);  // top
            desc.putBoolean(charIDToTypeID('B  '), true);  // bottom
            desc.putBoolean(charIDToTypeID('L  '), true);  // left
            desc.putBoolean(charIDToTypeID('R  '), true);  // right
            try {
                executeAction(charIDToTypeID('Trim'), desc, DialogModes.NO);
                var w = d.width.value;
                var h = d.height.value;
                'OK|' + w + 'x' + h;
            } catch(e) {
                'ERR|' + e.toString();
            }
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    if parts[0] == 'OK':
        dims = parts[1] if len(parts) > 1 else ""
        return {"result": "ok", "trimmed_to": dims,
                "message": f"已自动裁剪边缘，当前尺寸 {dims}"}
    return {"result": "error", "error": parts[1] if len(parts) > 1 else "trim failed"}


@mcp.tool()
def ps_batch_apply_action_to_layers(action: str, action_set: str = "",
                                     layer_filter: str = "") -> dict:
    """
    批量给当前文档的所有图层（或指定类型图层）应用同一个 PS 动作。
    手动操作需要逐个选中图层→执行动作→重复几十甚至上百次。
    - action: 动作名称
    - action_set: 动作集名称（默认第一个）
    - layer_filter: 筛选图层类型 "text"(仅文字层) "image"(仅像素层) ""(所有层)
    用户说"批量执行动作""给所有图层加效果""批量应用动作"时使用。
    """
    js = """
        var actName = '""" + action.replace("'", "\\'") + """';
        var actSet = '""" + action_set.replace("'", "\\'") + """';
        var filter = '""" + layer_filter + """';
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            // Get available action sets if not specified
            if (actSet === '') {
                try {
                    var sets = app.actionSets;
                    if (sets.length > 0) actSet = sets[0].name;
                } catch(e) {}
            }
            var applied = 0;
            var skipped = 0;
            var errors = [];
            for (var i = 0; i < d.artLayers.length; i++) {
                var lyr = d.artLayers[i];
                // Apply filter
                if (filter === 'text' && lyr.kind !== LayerKind.TEXT) {
                    skipped++;
                    continue;
                }
                if (filter === 'image' && lyr.kind !== LayerKind.NORMAL) {
                    skipped++;
                    continue;
                }
                try {
                    d.activeLayer = lyr;
                    app.doAction(actName, actSet);
                    applied++;
                } catch(e) {
                    skipped++;
                    errors.push(lyr.name + ': ' + e.toString().substring(0, 50));
                }
            }
            applied + '|' + skipped + (errors.length > 0 ? '|' + errors.slice(0,3).join(';') : '');
        }
    """
    result = _ps_js(js, max_retries=30, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    parts = result.split('|')
    applied = int(parts[0])
    skipped = int(parts[1]) if len(parts) > 1 else 0
    errs = parts[2] if len(parts) > 2 else ""
    msg = f"已对 {applied} 个图层执行动作 '{action}'"
    if skipped > 0:
        msg += f"，跳过 {skipped} 个"
    return {"result": "ok", "applied": applied, "skipped": skipped,
            "errors": errs, "message": msg}


@mcp.tool()
def ps_smart_object_replace_batch(folder: str = "",
                                  start_index: int = 1) -> dict:
    """
    批量替换当前文档中的智能对象内容。
    将文件夹中的图片依次替换到文档中的智能对象图层。
    手动操作：双击打开智能对象→替换内容→保存→关闭→重复几十次。
    - folder: 包含替换图片的文件夹路径
    - start_index: 从第几个图片开始（默认1）
    用户说"批量替换智能对象""智能对象替换内容""mockup批量生成"时使用。
    """
    folder = folder.replace("\\", "/")
    js = """
        var folderPath = '""" + folder + """';
        var startIdx = """ + str(start_index) + """;
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var srcFolder = new File(folderPath);
            if (!srcFolder.exists || !srcFolder.isDirectory()) { 'BAD_FOLDER'; }
            else {
                // Get all image files
                var files = srcFolder.getFiles(function(f) {
                    var ext = f.name.toLowerCase().match(/\\.(jpg|jpeg|png|tif|tiff|bmp|psd)$/);
                    return ext !== null;
                });
                if (files.length === 0) { 'NO_FILES'; }
                else {
                    // Sort files by name
                    files.sort(function(a, b) { return a.name < b.name ? -1 : 1; });
                    var start = startIdx - 1;
                    if (start < 0) start = 0;
                    if (start >= files.length) start = 0;

                    // Find all smart object layers
                    var soLayers = [];
                    for (var i = 0; i < d.artLayers.length; i++) {
                        try {
                            if (d.artLayers[i].kind === LayerKind.SMARTOBJECT) {
                                soLayers.push(d.artLayers[i]);
                            }
                        } catch(e) {}
                    }
                    // Also check layer sets
                    for (var j = 0; j < d.layerSets.length; j++) {
                        var setLayers = d.layerSets[j].artLayers;
                        for (var k = 0; k < setLayers.length; k++) {
                            try {
                                if (setLayers[k].kind === LayerKind.SMARTOBJECT) {
                                    soLayers.push(setLayers[k]);
                                }
                            } catch(e) {}
                        }
                    }

                    if (soLayers.length === 0) { 'NO_SMART_OBJECTS'; }
                    else {
                        var replaced = 0;
                        var errors = [];
                        for (var m = 0; m < soLayers.length; m++) {
                            var fileIdx = start + m;
                            if (fileIdx >= files.length) break;
                            try {
                                d.activeLayer = soLayers[m];
                                // Replace smart object content
                                var desc = new ActionDescriptor();
                                var ref = new ActionReference();
                                ref.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                                desc.putReference(charIDToTypeID('null'), ref);
                                desc.putPath(charIDToTypeID('On  '), new File(files[fileIdx].fsName));
                                desc.putBoolean(charIDToTypeID('Al  '), false);  // don't allow linked
                                executeAction(charIDToTypeID('Plc '), desc, DialogModes.NO);
                                replaced++;
                            } catch(e) {
                                errors.push(soLayers[m].name + ': ' + e.toString().substring(0, 60));
                            }
                        }
                        replaced + '|' + soLayers.length + (errors.length > 0 ? '|' + errors.slice(0,3).join(';') : '');
                    }
                }
            }
        }
    """
    result = _ps_js(js, max_retries=40, retry_sleep=0.5).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'BAD_FOLDER':
        return {"result": "error", "error": f"文件夹不存在: {folder}"}
    if result == 'NO_FILES':
        return {"result": "error", "error": "文件夹中没有图片文件"}
    if result == 'NO_SMART_OBJECTS':
        return {"result": "error", "error": "文档中没有智能对象图层"}
    parts = result.split('|')
    replaced = int(parts[0])
    total = int(parts[1]) if len(parts) > 1 else 0
    errs = parts[2] if len(parts) > 2 else ""
    return {"result": "ok", "replaced": replaced, "total_smart_objects": total,
            "errors": errs, "message": f"已替换 {replaced}/{total} 个智能对象"}


@mcp.tool()
def ps_auto_color_match(reference_layer_name: str = "") -> dict:
    """
    自动统一所有图层的色调——以指定图层为参考色调匹配其他图层。
    手动调色匹配需要逐个图层使用"匹配颜色"功能并反复调整参数。
    - reference_layer_name: 参考图层名称，空=使用当前活动图层
    用户说"统一色调""色调匹配""颜色统一""match color"时使用。
    """
    ref_name = reference_layer_name.replace("'", "\\'")
    js = """
        var refLayerName = '""" + ref_name + """';
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            // Find reference layer
            var refLayer = null;
            if (refLayerName !== '') {
                for (var i = 0; i < d.artLayers.length; i++) {
                    if (d.artLayers[i].name === refLayerName) {
                        refLayer = d.artLayers[i];
                        break;
                    }
                }
            }
            if (refLayer === null) refLayer = d.activeLayer;
            if (!refLayer || refLayer.kind === undefined) { 'NO_REF'; }
            else {
                var matched = 0;
                var skipped = 0;
                // Save current active layer
                var origActive = d.activeLayer;
                for (var j = 0; j < d.artLayers.length; j++) {
                    var lyr = d.artLayers[j];
                    if (lyr === refLayer) { skipped++; continue; }
                    // Skip adjustment layers and text layers
                    if (lyr.kind === LayerKind.TEXT || lyr.kind === LayerKind.ADJUSTMENT) {
                        skipped++;
                        continue;
                    }
                    try {
                        d.activeLayer = lyr;
                        // Match Color: use reference layer
                        var desc = new ActionDescriptor();
                        var ref = new ActionReference();
                        ref.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                        desc.putReference(charIDToTypeID('null'), ref);
                        // Source = reference layer
                        var srcRef = new ActionReference();
                        srcRef.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                        // Point to reference layer by index
                        srcRef.putIndex(charIDToTypeID('Lyr '), refLayer.itemIndex);
                        desc.putReference(charIDToTypeID('Srce'), srcRef);
                        desc.putInteger(charIDToTypeID('LneI'), 1);  // layer index
                        desc.putBoolean(charIDToTypeID('Fl  '), false);  // flatten?
                        desc.putInteger(charIDToTypeID('SmA '), 50);  // smoothing
                        desc.putInteger(charIDToTypeID('Cntr'), 0);   // contrast
                        desc.putInteger(charIDToTypeID('Brgh'), 0);   // brightness
                        desc.putInteger(charIDToTypeID('Intn'), 0);   // intensity
                        executeAction(stringIDToTypeID('matchColor'), desc, DialogModes.NO);
                        matched++;
                    } catch(e) {
                        skipped++;
                    }
                }
                // Restore active layer
                d.activeLayer = origActive;
                matched + '|' + skipped;
            }
        }
    """
    result = _ps_js(js, max_retries=20, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'NO_REF':
        return {"result": "error", "error": "找不到参考图层"}
    parts = result.split('|')
    matched = int(parts[0])
    skipped = int(parts[1]) if len(parts) > 1 else 0
    ref_desc = f"图层 '{reference_layer_name}'" if reference_layer_name else "当前活动图层"
    return {"result": "ok", "matched": matched, "skipped": skipped,
            "reference": ref_desc,
            "message": f"以{ref_desc}为参考，已匹配 {matched} 个图层色调"}


@mcp.tool()
def ps_auto_center_content(only_visible: bool = True) -> dict:
    """
    自动将文档中可见图层的整体内容居中到画布正中。
    手动需要测量内容边界、计算偏移、再逐个移动。
    - only_visible: 是否只处理可见图层，默认 True
    用户说"内容居中""居中到画布""居中所有内容"时使用。
    """
    only_vis = "true" if only_visible else "false"
    js = """
        var onlyVisible = """ + only_vis + """;
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var docW = d.width.value;
            var docH = d.height.value;
            // Find overall bounds of target layers
            var minLeft = 999999, minTop = 999999, maxRight = 0, maxBottom = 0;
            var found = false;
            for (var i = 0; i < d.artLayers.length; i++) {
                var lyr = d.artLayers[i];
                if (onlyVisible && !lyr.visible) continue;
                if (lyr.kind === LayerKind.ADJUSTMENT || lyr.kind === LayerKind.BACKGROUNDLAYER) continue;
                try {
                    var b = lyr.bounds;
                    if (b[0].value < minLeft) minLeft = b[0].value;
                    if (b[1].value < minTop) minTop = b[1].value;
                    if (b[2].value > maxRight) maxRight = b[2].value;
                    if (b[3].value > maxBottom) maxBottom = b[3].value;
                    found = true;
                } catch(e) {}
            }
            if (!found) { 'NONE'; }
            else {
                var contentW = maxRight - minLeft;
                var contentH = maxBottom - minTop;
                var targetLeft = (docW - contentW) / 2;
                var targetTop = (docH - contentH) / 2;
                var dx = targetLeft - minLeft;
                var dy = targetTop - minTop;
                // Move all target layers
                var moved = 0;
                for (var j = 0; j < d.artLayers.length; j++) {
                    var lyr2 = d.artLayers[j];
                    if (onlyVisible && !lyr2.visible) continue;
                    if (lyr2.kind === LayerKind.ADJUSTMENT || lyr2.kind === LayerKind.BACKGROUNDLAYER) continue;
                    try {
                        d.activeLayer = lyr2;
                        d.activeLayer.translate(dx, dy);
                        moved++;
                    } catch(e) {}
                }
                moved + '|' + Math.round(dx) + ',' + Math.round(dy);
            }
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'NONE':
        return {"result": "error", "error": "没有可处理的图层"}
    parts = result.split('|')
    moved = int(parts[0])
    offset = parts[1] if len(parts) > 1 else "0,0"
    return {"result": "ok", "moved_layers": moved, "offset": offset,
            "message": f"已将 {moved} 个图层居中（偏移 {offset}）"}


@mcp.tool()
def ps_auto_round_corners(radius: int = 20, apply_to: str = "all") -> dict:
    """
    批量给当前文档的图层添加圆角效果（电商商品图常用）。
    手动需要逐个图层用圆角矩形工具重画或添加蒙版。
    - radius: 圆角半径（像素），默认 20
    - apply_to: "all"（所有图层）或 "visible"（仅可见层）或 "selected"（选中层）
    用户说"批量加圆角""圆角处理""给图层加圆角"时使用。
    """
    apply_to = apply_to.lower().strip()
    if apply_to not in ("all", "visible", "selected"):
        apply_to = "all"

    js = """
        var r = """ + str(radius) + """;
        var mode = '""" + apply_to + """';
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var processed = 0;
            for (var i = 0; i < d.artLayers.length; i++) {
                var lyr = d.artLayers[i];
                if (mode === 'visible' && !lyr.visible) continue;
                if (lyr.kind === LayerKind.TEXT || lyr.kind === LayerKind.ADJUSTMENT) continue;
                try {
                    d.activeLayer = lyr;
                    var b = lyr.bounds;
                    var l = b[0].value, t = b[1].value;
                    var w = b[2].value - l, h = b[3].value - t;
                    // Create rounded rectangle selection
                    var sel = [[l, t+r], [l, t], [l+w, t], [l+w, t+r],
                               [l+w, t+h-r], [l+w, t+h], [l, t+h], [l, t+h-r]];
                    d.selection.select(sel);
                    // Feather the selection for rounded corners
                    d.selection.feather(r);
                    // Invert and delete outside
                    d.selection.invert();
                    d.selection.clear();
                    d.selection.deselect();
                    processed++;
                } catch(e) {}
            }
            processed + '';
        }
    """
    result = _ps_js(js, max_retries=20, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    processed = int(result) if result.isdigit() else 0
    return {"result": "ok", "processed": processed, "radius": radius,
            "message": f"已对 {processed} 个图层添加 {radius}px 圆角"}


@mcp.tool()
def ps_auto_layout_strip(direction: str = "horizontal",
                         gap: float = 10,
                         background: bool = True) -> dict:
    """
    自动将多个图层排列成连续条带（长图/长条）。
    常用于制作长图、漫画条、时间线图等。
    手动需要计算每个图层的尺寸、逐一排列、调整画布大小。
    - direction: "horizontal"(横向长图) 或 "vertical"(纵向长图)
    - gap: 图层间距（像素），默认 10
    - background: 是否新建背景画布，默认 True
    用户说"拼长图""图层排成一行""做成长条""strip layout"时使用。
    """
    direction = direction.lower().strip()
    if direction not in ("horizontal", "vertical", "h", "v"):
        direction = "horizontal"
    if direction in ("h",):
        direction = "horizontal"
    if direction in ("v",):
        direction = "vertical"

    js = """
        var dir = '""" + direction + """';
        var g = """ + str(gap) + """;
        var makeBg = """ + ("true" if background else "false") + """;
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var layers = [];
            for (var i = 0; i < d.artLayers.length; i++) {
                var lyr = d.artLayers[i];
                if (lyr.kind === LayerKind.ADJUSTMENT) continue;
                layers.push(lyr);
            }
            if (layers.length < 2) { 'TOO_FEW'; }
            else {
                // Calculate total dimensions
                var totalGap = (layers.length - 1) * g;
                var totalSize = 0;
                var maxSize = 0;
                var boundsList = [];
                for (var j = 0; j < layers.length; j++) {
                    var b = layers[j].bounds;
                    var w = b[2].value - b[0].value;
                    var h = b[3].value - b[1].value;
                    boundsList.push({w: w, h: h, b: b});
                    if (dir === 'horizontal') {
                        totalSize += w;
                        if (h > maxSize) maxSize = h;
                    } else {
                        totalSize += h;
                        if (w > maxSize) maxSize = w;
                    }
                }
                var canvasW, canvasH;
                if (dir === 'horizontal') {
                    canvasW = totalSize + totalGap;
                    canvasH = maxSize;
                } else {
                    canvasW = maxSize;
                    canvasH = totalSize + totalGap;
                }
                // Resize canvas
                var oldW = d.width.value;
                var oldH = d.height.value;
                d.resizeCanvas(canvasW, canvasH, AnchorPosition.TOPLEFT);
                // Move layers into position
                var cursor = 0;
                for (var m = 0; m < layers.length; m++) {
                    var info = boundsList[m];
                    var b2 = layers[m].bounds;
                    var curLeft = b2[0].value;
                    var curTop = b2[1].value;
                    if (dir === 'horizontal') {
                        var dx = cursor - curLeft;
                        layers[m].translate(dx, -curTop);
                        cursor += info.w + g;
                    } else {
                        var dy = cursor - curTop;
                        layers[m].translate(-curLeft, dy);
                        cursor += info.h + g;
                    }
                }
                canvasW + 'x' + canvasH;
            }
        }
    """
    result = _ps_js(js, max_retries=20, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'TOO_FEW':
        return {"result": "error", "error": "至少需要 2 个图层"}
    return {"result": "ok", "canvas_size": result,
            "message": f"已排列 {direction} 条带，画布尺寸 {result}"}


@mcp.tool()
def ps_export_layer_comps(output_folder: str = "") -> dict:
    """
    一键导出所有图层组合（Layer Comps）为单独文件。
    手动操作需要逐个切换图层组合→导出→重复多次。
    - output_folder: 输出文件夹路径，默认桌面
    用户说"导出图层组合""export layer comps""批量导出版本"时使用。
    """
    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop", "layer_comps_export")
    output_folder = output_folder.replace("\\", "/")
    js = """
        var outFolder = new Folder('""" + output_folder + """');
        if (!outFolder.exists) outFolder.create();
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var comps = d.layerComps;
            if (comps.length === 0) { 'NO_COMPS'; }
            else {
                var exported = 0;
                for (var i = 0; i < comps.length; i++) {
                    try {
                        comps[i].apply();
                        var safeName = comps[i].name.replace(/[^a-zA-Z0-9_\\-]/g, '_');
                        if (safeName.length === 0) safeName = 'comp_' + (i+1);
                        var fname = outFolder + '/' + safeName + '_' + (i+1) + '.png';
                        var saveFile = new File(fname);
                        var pngOpts = new PNGSaveOptions();
                        pngOpts.compression = 6;
                        d.saveAs(saveFile, pngOpts, true);
                        exported++;
                    } catch(e) {}
                }
                exported + '|' + comps.length;
            }
        }
    """
    result = _ps_js(js, max_retries=30, retry_sleep=0.3).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'NO_COMPS':
        return {"result": "error", "error": "文档中没有图层组合"}
    parts = result.split('|')
    exported = int(parts[0])
    total = int(parts[1]) if len(parts) > 1 else 0
    return {"result": "ok", "exported": exported, "total": total,
            "folder": output_folder,
            "message": f"已导出 {exported}/{total} 个图层组合"}


# ---- Illustrator 高级批量工具 ----

@mcp.tool()
def ai_batch_replace_text(find_text: str, replace_text: str,
                          case_sensitive: bool = False) -> dict:
    """
    在 Illustrator 中批量查找替换所有文字框中的文字。
    手动需要打开每个文字框逐一查找替换。
    - find_text: 要查找的文字
    - replace_text: 替换为的文字
    - case_sensitive: 是否区分大小写，默认 False
    用户说"AI批量替换文字""查找替换文字""批量改文字内容"时使用。
    """
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    doc = app.ActiveDocument
    text_frames = doc.TextFrames
    count = 0
    replaced = 0
    find_lower = find_text.lower() if not case_sensitive else find_text
    for i in range(1, text_frames.Count + 1):
        tf = text_frames.Item(i)
        content = tf.Contents
        count += 1
        if not case_sensitive:
            if find_lower in content.lower():
                new_content = content.replace(find_text, replace_text)  # Python replace
                tf.Contents = new_content
                replaced += 1
        else:
            if find_text in content:
                tf.Contents = content.replace(find_text, replace_text)
                replaced += 1
    return {"result": "ok", "total_frames": count, "replaced": replaced,
            "message": f"在 {replaced}/{count} 个文字框中替换了 '{find_text}' → '{replace_text}'"}


@mcp.tool()
def ai_export_all_artboards(output_folder: str = "", format: str = "png") -> dict:
    """
    一键导出 Illustrator 中所有画板为单独文件。
    手动需要逐个画板导出，文件多时极费时间。
    - output_folder: 输出文件夹路径，默认桌面
    - format: "png" 或 "svg" 或 "pdf"
    用户说"导出所有画板""批量导出画板""export artboards"时使用。
    """
    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop", "artboards_export")
    os.makedirs(output_folder, exist_ok=True)

    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    doc = app.ActiveDocument
    artboards = doc.Artboards
    total = artboards.Count
    if total == 0:
        return {"result": "error", "error": "no artboards"}

    exported = 0
    fmt = format.lower().strip()
    for i in range(1, total + 1):
        try:
            artboards.ActiveArtboardIndex = i - 1
            safe_name = f"artboard_{i:03d}"
            ext = "png" if fmt == "png" else ("svg" if fmt == "svg" else "pdf")
            path = os.path.join(output_folder, f"{safe_name}.{ext}")

            if fmt == "png":
                opts = win32com.client.Dispatch("Illustrator.PNGExportOptions")
                opts.AntiAliasing = True
                opts.Transparency = True
                doc.Export(path, 4, opts)  # 4 = aiPNG
            elif fmt == "svg":
                opts = win32com.client.Dispatch("Illustrator.SVGExportOptions")
                opts.FontType = 1
                doc.Export(path, 10, opts)  # 10 = aiSVG
            else:
                opts = win32com.client.Dispatch("Illustrator.PDFExportOptions")
                doc.Export(path, 5, opts)  # 5 = aiPDF
            exported += 1
        except Exception as e:
            pass

    return {"result": "ok", "exported": exported, "total": total,
            "folder": output_folder, "format": fmt,
            "message": f"已导出 {exported}/{total} 个画板为 {fmt.upper()} 文件"}


@mcp.tool()
def ai_auto_layout_grid(rows: int = 3, cols: int = 3, gap: float = 10) -> dict:
    """
    将 Illustrator 文档中所有对象自动排列成网格布局。
    手动排列几十个对象到精确网格位置极费时间。
    - rows: 行数
    - cols: 列数
    - gap: 间距（点）
    用户说"AI网格排列""自动排版网格""grid layout"时使用。
    """
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    doc = app.ActiveDocument
    items = doc.PageItems
    total = items.Count
    if total < 2:
        return {"result": "error", "error": "至少需要 2 个对象"}

    # Collect all items with their bounds
    item_info = []
    for i in range(1, total + 1):
        item = items.Item(i)
        try:
            left = item.Left
            top = item.Top
            w = item.Width
            h = item.Height
            item_info.append({"item": item, "left": left, "top": top, "w": w, "h": h})
        except Exception:
            pass

    if not item_info:
        return {"result": "error", "error": "无法读取对象信息"}

    # Find max dimensions
    max_w = max(info["w"] for info in item_info) if item_info else 100
    max_h = max(info["h"] for info in item_info) if item_info else 100

    cell_w = max_w + gap
    cell_h = max_h + gap
    start_left = 50.0
    start_top = 800.0  # AI Y axis goes down

    arranged = 0
    for idx, info in enumerate(item_info):
        row = idx // cols
        col = idx % cols
        target_left = start_left + col * cell_w
        target_top = start_top - row * cell_h  # Y inverted in AI
        try:
            dx = target_left - info["left"]
            dy = target_top - info["top"]
            info["item"].Left = target_left
            info["item"].Top = target_top
            arranged += 1
        except Exception:
            pass

    grid_desc = f"{rows}×{cols}"
    return {"result": "ok", "arranged": arranged, "grid": grid_desc,
            "cell_size": f"{cell_w:.1f}×{cell_h:.1f}",
            "message": f"已将 {arranged} 个对象排列为 {grid_desc} 网格"}


@mcp.tool()
def ai_align_objects(direction: str = "left") -> dict:
    """
    将 Illustrator 文档中所有选中对象按指定方式对齐。
    手动逐个调整对齐极费时间。
    - direction: "left" "right" "top" "bottom" "center_h" "center_v"
    用户说"AI对齐对象""左对齐""居中对齐""align"时使用。
    """
    app = _get_ai()
    if app.Documents.Count == 0:
        return {"result": "error", "error": "no document"}
    doc = app.ActiveDocument
    # Get selection
    try:
        selection = doc.Selection
        if selection.Count < 2:
            return {"result": "error", "error": "请先选中至少 2 个对象"}
    except Exception:
        return {"result": "error", "error": "无法获取选中对象"}

    direction = direction.lower().strip()
    items = []
    for i in range(1, selection.Count + 1):
        item = selection.Item(i)
        try:
            items.append({"item": item, "left": item.Left, "top": item.Top,
                          "w": item.Width, "h": item.Height})
        except Exception:
            pass

    if len(items) < 2:
        return {"result": "error", "error": "无法读取对象信息"}

    aligned = 0
    if direction == "left":
        target = min(info["left"] for info in items)
        for info in items:
            info["item"].Left = target
            aligned += 1
    elif direction == "right":
        max_right = max(info["left"] + info["w"] for info in items)
        for info in items:
            info["item"].Left = max_right - info["w"]
            aligned += 1
    elif direction == "top":
        target = min(info["top"] for info in items)
        for info in items:
            info["item"].Top = target
            aligned += 1
    elif direction == "bottom":
        max_bottom = max(info["top"] + info["h"] for info in items)
        for info in items:
            info["item"].Top = max_bottom - info["h"]
            aligned += 1
    elif direction in ("center_h", "center_horizontal"):
        avg = sum(info["left"] + info["w"] / 2 for info in items) / len(items)
        for info in items:
            info["item"].Left = avg - info["w"] / 2
            aligned += 1
    elif direction in ("center_v", "center_vertical"):
        avg = sum(info["top"] + info["h"] / 2 for info in items) / len(items)
        for info in items:
            info["item"].Top = avg - info["h"] / 2
            aligned += 1
    else:
        return {"result": "error", "error": f"不支持的对齐方式: {direction}"}

    return {"result": "ok", "aligned": aligned, "direction": direction,
            "message": f"已将 {aligned} 个对象{direction}对齐"}


# ============================================================
# Main
# ============================================================

if __name__ == "__main__":
    print("[adobe-com-mcp] starting...", file=sys.stderr, flush=True)
    # Eager connection to surface errors early
    try:
        _get_ps()
        print("[adobe-com-mcp] Photoshop connected", file=sys.stderr, flush=True)
    except Exception as e:
        print(f"[adobe-com-mcp] Photoshop not available: {e}", file=sys.stderr, flush=True)
    try:
        _get_ai()
        print("[adobe-com-mcp] Illustrator connected", file=sys.stderr, flush=True)
    except Exception as e:
        print(f"[adobe-com-mcp] Illustrator not available: {e}", file=sys.stderr, flush=True)
    print("[adobe-com-mcp] entering MCP loop...", file=sys.stderr, flush=True)
    mcp.run(transport="stdio")
