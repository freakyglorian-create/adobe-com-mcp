#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Adobe COM MCP Server — macOS 版（AppleScript 后端）

与 Windows 版 server.py 工具名、参数、返回结构完全一致（ps_* / ai_*），
客户端（TRAE）换台 Mac 只改启动路径即可无缝使用。

原理：通过 macOS 自带的 osascript 运行 AppleScript，再借 Photoshop / Illustrator
的 `do javascript` 命令执行 ExtendScript 完成自动化。无需安装任何额外依赖
（只需 Python 3.10+ 和 `pip install "mcp<2"`）。

⚠️ 本文件在 Windows 上无法运行/验证，请在 Mac 上实测调试。
   Illustrator 的 `do javascript` 在不同版本返回行为略有差异，如数据返回异常请反馈。
"""

import json
import math
import os
import subprocess
import sys
import tempfile
import time

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("adobe-mac-mcp")


# --------------------------------------------------------------------------- #
# 底层：AppleScript / ExtendScript 执行
# --------------------------------------------------------------------------- #

def _osascript(applescript: str) -> str:
    """运行一段 AppleScript，返回 stdout（去尾随换行）。失败抛 RuntimeError。"""
    proc = subprocess.run(
        ["osascript", "-e", applescript],
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "osascript 执行失败")
    return proc.stdout.strip()


def _ps_js(js_code: str, max_retries: int = 40, retry_sleep: float = 0.25) -> str:
    """
    通过 Photoshop 的 `do javascript` 执行 ExtendScript 字符串，返回结果。
    AppleScript 字符串里反斜杠是字面量（无需二次转义），因此只需：
      1) 真实换行 -> 字面 \\n（交给 ExtendScript 解释回换行）
      2) 双引号   -> \\"（AppleScript 字符串定界符转义）
    max_retries/retry_sleep 仅为兼容 Windows 版签名，此处忽略。
    """
    single_line = (
        js_code.replace("\r", "")
        .replace("\n", "\\n")
        .replace('"', '\\"')
    )
    script = (
        'tell application id "com.adobe.Photoshop"\n'
        f'\tdo javascript "{single_line}"\n'
        'end tell'
    )
    return _osascript(script)


def _ai_js(js: str) -> str:
    """
    通过 Illustrator 的 `do javascript` 执行 ExtendScript 文件，返回结果。
    Illustrator 的 AppleScript 接口要求传入 .jsx 文件路径，因此先写临时文件。
    """
    fd, path = tempfile.mkstemp(suffix=".jsx")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(js)
    script = (
        'tell application id "com.adobe.illustrator"\n'
        f'\tdo javascript (POSIX file "{path}")\n'
        'end tell'
    )
    try:
        return _osascript(script)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def _gcd(a: int, b: int) -> int:
    a, b = abs(a), abs(b)
    while b:
        a, b = b, a % b
    return a or 1


def _parse(result: str):
    """尽力把返回字符串解析成 JSON，失败则原样返回字符串。"""
    try:
        return json.loads(result)
    except (json.JSONDecodeError, TypeError):
        return result


# ============================================================
# PHOTOSHOP TOOLS（ExtendScript 与 Windows 版完全一致）
# ============================================================

@mcp.tool()
def ps_create_document(width: int, height: int = 0, name: str = "") -> dict:
    """新建指定像素尺寸的 Photoshop 文档。height 不填则等于 width。"""
    if height <= 0:
        height = width
    width = max(1, int(width))
    height = max(1, int(height))
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
    """获取当前文档信息：名称、像素尺寸、分辨率、颜色模式、图层数、活动图层名、打开文档数。"""
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
    """新建空白图层。name 可选。"""
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
    """设置当前活动图层不透明度（0-100）。"""
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
    """添加文字图层（文字内容、字号、位置、颜色）。"""
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
    """用指定颜色填充当前活动图层（全画布填充）。"""
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
    """设置 Photoshop 前景色（RGB 0-255）。"""
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
    """对当前活动图层应用高斯模糊滤镜。"""
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
    """对当前活动图层应用 USM 锐化。"""
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
    """调整当前文档像素尺寸（给任一值则等比缩放）。"""
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
    """将当前文档导出为 PNG。"""
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
    """将当前文档导出为 JPG（quality 1-12）。"""
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
    """将当前文档另存为 PSD。"""
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
    """打开本地图片/PSD 文件。"""
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
    """执行已安装的 Photoshop 动作（Action）。"""
    js = f"""
        app.doAction({json.dumps(action)}, {json.dumps(frm) if frm else '""'});
        'ok';
    """
    _ps_js(js)
    return {"result": "ok", "action": action, "from": frm}


@mcp.tool()
def ps_list_text_layers() -> dict:
    """列出当前文档所有文字图层（名称、内容、字号、颜色、位置、不透明度）。"""
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
    """一键全选当前文档中的所有文字图层。"""
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
                var desc = new ActionDescriptor();
                var ref = new ActionReference();
                for (var j = 0; j < textLayers.length; j++) {
                    ref.putIndex(charIDToTypeID('Lyr '), textLayers[j].itemIndex);
                }
                desc.putReference(charIDToTypeID('null'), ref);
                try {
                    executeAction(charIDToTypeID('slct'), desc, DialogModes.NO);
                } catch(e) {
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
    """批量替换所有文字图层中的指定文字。"""
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
    """批量设置所有文字图层的颜色。"""
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
    """批量设置所有文字图层的字号。"""
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
    """列出当前文档所有图层（名称、类型、可见性、锁定、不透明度、是否文字层）。"""
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
    """按名称选中图层。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var found = false;
            for (var i = 0; i < d.artLayers.length; i++) {{
                if (d.artLayers[i].name === {json.dumps(name)}) {{
                    d.activeLayer = d.artLayers[i];
                    found = true;
                    break;
                }}
            }}
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
    """切换图层可见性（不填则切换当前活动图层）。"""
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
# PHOTOSHOP 高级批量操作（ExtendScript 与 Windows 版一致）
# ============================================================

@mcp.tool()
def ps_batch_resize_folder(folder: str, target_width: int = 0, target_height: int = 0,
                           suffix: str = "", output_format: str = "jpg",
                           quality: int = 8) -> dict:
    """批量处理文件夹中所有图片：打开、调整尺寸、导出。"""
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
            r = _ps_js(js).strip()
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
    """批量给文件夹中所有图片加水印文字。"""
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
                var docW = d.width.value;
                var docH = d.height.value;
                tf.textItem.position = [docW / 2 - 2, docH / 2 - 2];
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
            r = _ps_js(js).strip()
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
    """将当前文档一键导出为多个社交媒体平台尺寸。"""
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
                    var src = app.activeDocument;
                    var ppi = src.resolution;
                    src.resizeImage({tw} / ppi, {th} / ppi, ppi, ResampleMethod.BICUBIC);
                    var opts = new JPEGSaveOptions();
                    opts.quality = 10;
                    src.saveAs(new File('{out_path.replace(chr(92), chr(47))}'), opts, true);
                    try {{ app.executeAction(app.charIDToTypeID('undo'), undefined, DialogModes.NO); }} catch(e) {{}}
                    'OK|{plat}|{tw}x{th}|{prefix}';
                }}
            }} catch(e) {{
                'FAIL|{plat}|' + e.toString();
            }}
        """
        try:
            r = _ps_js(js).strip()
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
    """生成联系表（缩略图网格）。"""
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

            var g = _gcd(cw, ch);
            var doc = app.documents.add(cw / g, ch / g, g, 'ContactSheet');
            var ppi = doc.resolution;

            var bgc = new SolidColor();
            bgc.rgb.red = {bg_r}; bgc.rgb.green = {bg_g}; bgc.rgb.blue = {bg_b};
            doc.selection.selectAll();
            doc.selection.fill(bgc);
            doc.selection.deselect();

            function _gcd(a, b) {{ a = Math.abs(a); b = Math.abs(b); while(b) {{ var t = b; b = a % b; a = t; }} return a || 1; }}

            for (var i = 0; i < files.length && i < cols * rows; i++) {{
                var col = i % cols;
                var row = Math.floor(i / cols);
                var px = gap + col * (thumb + gap);
                var py = gap + row * (thumb + gap);

                var imgDoc = app.open(new File(files[i]));
                var imgW = imgDoc.width.value * imgDoc.resolution;
                var imgH = imgDoc.height.value * imgDoc.resolution;
                var scale = Math.min(thumb / imgW, thumb / imgH);
                var newW = imgW * scale;
                var newH = imgH * scale;
                imgDoc.resizeImage(newW / imgDoc.resolution, newH / imgDoc.resolution, imgDoc.resolution);
                imgDoc.selection.selectAll();
                imgDoc.selection.duplicate(doc.artLayers[0]);
                imgDoc.close(SaveOptions.DONOTSAVECHANGES);

                var pasted = doc.activeLayer;
                pasted.translate((px / ppi), (py / ppi));
            }}

            var opts = new PNGSaveOptions();
            opts.compression = 6;
            doc.saveAs(new File('{output_path.replace(chr(92), chr(47))}'), opts, true);
            'OK|' + files.length + '|' + cols + 'x' + rows + '|' + cw + 'x' + ch;
        }} catch(e) {{
            'FAIL|' + e.toString();
        }}
    """
    result = _ps_js(js).strip()
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
    """批量重命名所有图层（或特定类型图层）。"""
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
    """在背景图层创建渐变填充（颜色1→颜色2）。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var desc = new ActionDescriptor();
            var gradDesc = new ActionDescriptor();
            var colors = new ActionList();

            var c1 = new ActionDescriptor();
            c1.putEnumerated(charIDToTypeID('Clr '), charIDToTypeID('ClrS'), charIDToTypeID('RGBC'));
            c1.putDouble(charIDToTypeID('Rd  '), {r1});
            c1.putDouble(charIDToTypeID('Grn '), {g1});
            c1.putDouble(charIDToTypeID('Bl  '), {b1});

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
    """根据 CSV 数据批量生成多个版本的 PS 文档。"""
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
                    try {{ app.executeAction(app.charIDToTypeID('undo'), undefined, DialogModes.NO); }} catch(e) {{}}
                    'OK|{val}|{out_path}';
                }}
            }} catch(e) {{
                'FAIL|{val}|' + e.toString();
            }}
        """
        try:
            r = _ps_js(js).strip()
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
    """从当前文档提取主色调色盘。"""
    num_colors = max(2, min(12, int(num_colors)))

    import struct
    tmp_path = os.path.join(tempfile.gettempdir(), "palette_sample.bmp")

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
            dup.saveAs(new File('{tmp_path.replace(chr(92), chr(47))}'), new BMPSaveOptions(), true);
            dup.close(SaveOptions.DONOTSAVECHANGES);
            'OK|' + tw + 'x' + th;
        }}
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if not result.startswith('OK'):
        return {"result": "error", "error": f"export failed: {result}"}

    try:
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
                    b_val = data[px]
                    g_val = data[px + 1]
                    r_val = data[px + 2]
                    qr = (r_val // 32) * 32
                    qg = (g_val // 32) * 32
                    qb = (b_val // 32) * 32
                    key = (qr, qg, qb)
                    hist[key] = hist.get(key, 0) + 1

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
    """Fallback: 通过直方图通道采样主色。"""
    num_colors = max(2, min(6, int(num_colors)))
    js = """
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var rHist = d.histogram[0];
            var gHist = d.histogram[1];
            var bHist = d.histogram[2];

            var rPeaks = [], gPeaks = [], bPeaks = [];
            for (var i = 0; i < 256; i++) {
                rPeaks.push([i, rHist[i]]);
                gPeaks.push([i, gHist[i]]);
                bPeaks.push([i, bHist[i]]);
            }
            rPeaks.sort(function(a,b) { return b[1] - a[1]; });
            gPeaks.sort(function(a,b) { return b[1] - a[1]; });
            bPeaks.sort(function(a,b) { return b[1] - a[1]; });

            var n = __N__;
            var result = [];
            for (var j = 0; j < n; j++) {
                var r = rPeaks[j] ? rPeaks[j][0] : 0;
                var g = gPeaks[j] ? gPeaks[j][0] : 0;
                var b = bPeaks[j] ? bPeaks[j][0] : 0;
                result.push(r + ',' + g + ',' + b);
            }
            result.join(';');
        }
    """.replace("__N__", str(num_colors))
    try:
        result = _ps_js(js).strip()
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
    """自动批量排版卡片（R行×C列卡片位）。"""
    total = rows * cols
    canvas_w = cols * card_w + (cols + 1) * gap
    canvas_h = rows * card_h + (rows + 1) * gap

    js = f"""
        try {{
            var cw = {canvas_w};
            var ch = {canvas_h};
            var ppi = 100;
            var doc = app.documents.add(cw / ppi, ch / ppi, ppi, 'CardLayout');

            var bgc = new SolidColor();
            bgc.rgb.red = {bg_r}; bgc.rgb.green = {bg_g}; bgc.rgb.blue = {bg_b};
            doc.selection.selectAll();
            doc.selection.fill(bgc);
            doc.selection.deselect();

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
                    doc.selection.stroke(cardColor, 1);
                }}
            }}
            doc.selection.deselect();
            'OK|' + rows + 'x' + cols + '|' + cw + 'x' + ch + '|' + {total};
        }} catch(e) {{
            'FAIL|' + e.toString();
        }}
    """
    result = _ps_js(js).strip()
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
    """将每个图层导出为帧（供生成 GIF 动画）。"""
    if not os.path.isdir(folder):
        os.makedirs(folder, exist_ok=True)

    js = f"""
        try {{
            if (app.documents.length === 0) {{ 'FAIL|no document'; }}
            else {{
                var d = app.activeDocument;
                var layers = d.artLayers;
                var frameCount = layers.length;
                var outDir = '{folder.replace(chr(92), chr(47))}';

                for (var i = 0; i < frameCount; i++) {{
                    for (var j = 0; j < frameCount; j++) {{
                        layers[j].visible = (j === i);
                    }}
                    var frameFile = new File(outDir + '/frame_' + (i+1) + '.png');
                    var opts = new PNGSaveOptions();
                    opts.compression = 6;
                    d.saveAs(frameFile, opts, true);
                }}

                for (var k = 0; k < frameCount; k++) {{
                    layers[k].visible = true;
                }}

                'OK|' + frameCount + ' frames exported';
            }}
        }} catch(e) {{
            'FAIL|' + e.toString();
        }}
    """
    result = _ps_js(js).strip()
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


@mcp.tool()
def ps_export_all_layers_to_png(output_folder: str = "") -> dict:
    """一键导出所有图层为单独 PNG。"""
    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop", "layers_export")
    output_folder = output_folder.replace("\\", "/")
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
                var visStates = [];
                for (var j = 0; j < layers.length; j++) {
                    visStates.push(layers[j].visible);
                }
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
                for (var n = 0; n < layers.length; n++) {
                    layers[n].visible = visStates[n];
                }
                exported.length + '|' + exported.join(';');
            }
        }
    """
    result = _ps_js(js).strip()
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
    """将当前文档所有图层拼成精灵表（Sprite Sheet）。"""
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
                var docW = d.width.value;
                var docH = d.height.value;
                var cw = """ + str(cell_w) + """;
                var ch = """ + str(cell_h) + """;
                if (cw <= 0) cw = Math.round(docW);
                if (ch <= 0) ch = Math.round(docH);

                var totalW = cols * cw;
                var totalH = rows * ch;
                var sheet = app.documents.add(totalW, totalH, d.resolution, 'SpriteSheet',
                                              NewDocumentMode.RGB, DocumentFill.TRANSPARENT);

                var visStates = [];
                for (var j = 0; j < layers.length; j++) {
                    visStates.push(layers[j].visible);
                }

                for (var m = 0; m < count; m++) {
                    for (var k = 0; k < layers.length; k++) {
                        layers[k].visible = false;
                    }
                    layers[m].visible = true;
                    var dup = d.duplicate('frame_tmp');
                    dup.flatten();
                    dup.selection.selectAll();
                    try { dup.selection.copy(); } catch(e) {}
                    dup.close(SaveOptions.DONOTSAVECHANGES);
                    sheet.activate();
                    var col = m % cols;
                    var row = Math.floor(m / cols);
                    var left = col * cw;
                    var top = row * ch;
                    var right = left + cw;
                    var bottom = top + ch;
                    var selRegion = [[left, top], [right, top], [right, bottom], [left, bottom]];
                    sheet.selection.select(selRegion);
                    try { sheet.paste(); } catch(e) {}
                    try {
                        if (sheet.artLayers.length > 1) {
                            sheet.activeLayer = sheet.artLayers[sheet.artLayers.length - 1];
                            sheet.activeLayer.merge();
                        }
                    } catch(e) {}
                }
                for (var n = 0; n < layers.length; n++) {
                    layers[n].visible = visStates[n];
                }
                var saveFile = new File(outDir + '/spritesheet.png');
                var pngOpts = new PNGSaveOptions();
                pngOpts.compression = 6;
                sheet.saveAs(saveFile, pngOpts, true);
                count + '|' + cols + 'x' + rows + '|' + totalW + 'x' + totalH + '|' + outDir + '/spritesheet.png';
            }
        }
    """
    result = _ps_js(js).strip()
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
    """将选中图层均匀分布间距（等距排列）。"""
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
            var selected = [];
            var ref = new ActionReference();
            ref.putEnumerated(charIDToTypeID('Dcmn'), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
            var desc = executeActionGet(ref);
            var selCount = 0;
            try {
                var list = desc.getList(charIDToTypeID('LyrI'));
                selCount = list.count;
            } catch(e) { selCount = 0; }
            if (selCount < 2) { 'TOO_FEW'; }
            else {
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
                        var totalSpan2 = totalBottom - totalTop;
                        var totalItemH = 0;
                        for (var r2 = 0; r2 < items.length; r2++) totalItemH += items[r2].h;
                        var totalGap2 = totalSpan2 - totalItemH;
                        var eachGap2 = gap > 0 ? gap : totalGap2 / (items.length - 1);
                        var cursor2 = totalTop;
                        for (var s = 0; s < items.length; s++) {
                            d.activeLayer = d.artLayers[items[s].idx - 1];
                            var dy = cursor2 - items[s].top;
                            d.activeLayer.translate(0, dy);
                            cursor2 += items[s].h + eachGap2;
                        }
                    }
                    distributed + '|' + dir;
                }
            }
        }
    """
    result = _ps_js(js).strip()
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
    """自动裁剪边缘（透明/白色/黑色）。"""
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
            desc.putBoolean(charIDToTypeID('T  '), true);
            desc.putBoolean(charIDToTypeID('B  '), true);
            desc.putBoolean(charIDToTypeID('L  '), true);
            desc.putBoolean(charIDToTypeID('R  '), true);
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
    """批量给所有图层（或指定类型）应用同一个动作。"""
    js = """
        var actName = '""" + action.replace("'", "\\'") + """';
        var actSet = '""" + action_set.replace("'", "\\'") + """';
        var filter = '""" + layer_filter + """';
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
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
    result = _ps_js(js).strip()
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
    """批量替换当前文档中智能对象的内容。"""
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
                var files = srcFolder.getFiles(function(f) {
                    var ext = f.name.toLowerCase().match(/\\.(jpg|jpeg|png|tif|tiff|bmp|psd)$/);
                    return ext !== null;
                });
                if (files.length === 0) { 'NO_FILES'; }
                else {
                    files.sort(function(a, b) { return a.name < b.name ? -1 : 1; });
                    var start = startIdx - 1;
                    if (start < 0) start = 0;
                    if (start >= files.length) start = 0;

                    var soLayers = [];
                    for (var i = 0; i < d.artLayers.length; i++) {
                        try {
                            if (d.artLayers[i].kind === LayerKind.SMARTOBJECT) {
                                soLayers.push(d.artLayers[i]);
                            }
                        } catch(e) {}
                    }
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
                                var desc = new ActionDescriptor();
                                var ref = new ActionReference();
                                ref.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                                desc.putReference(charIDToTypeID('null'), ref);
                                desc.putPath(charIDToTypeID('On  '), new File(files[fileIdx].fsName));
                                desc.putBoolean(charIDToTypeID('Al  '), false);
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
    result = _ps_js(js).strip()
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
    """以指定图层为参考，统一所有图层色调。"""
    ref_name = reference_layer_name.replace("'", "\\'")
    js = """
        var refLayerName = '""" + ref_name + """';
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
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
                var origActive = d.activeLayer;
                for (var j = 0; j < d.artLayers.length; j++) {
                    var lyr = d.artLayers[j];
                    if (lyr === refLayer) { skipped++; continue; }
                    if (lyr.kind === LayerKind.TEXT || lyr.kind === LayerKind.ADJUSTMENT) {
                        skipped++;
                        continue;
                    }
                    try {
                        d.activeLayer = lyr;
                        var desc = new ActionDescriptor();
                        var ref = new ActionReference();
                        ref.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                        desc.putReference(charIDToTypeID('null'), ref);
                        var srcRef = new ActionReference();
                        srcRef.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                        srcRef.putIndex(charIDToTypeID('Lyr '), refLayer.itemIndex);
                        desc.putReference(charIDToTypeID('Srce'), srcRef);
                        desc.putInteger(charIDToTypeID('LneI'), 1);
                        desc.putBoolean(charIDToTypeID('Fl  '), false);
                        desc.putInteger(charIDToTypeID('SmA '), 50);
                        desc.putInteger(charIDToTypeID('Cntr'), 0);
                        desc.putInteger(charIDToTypeID('Brgh'), 0);
                        desc.putInteger(charIDToTypeID('Intn'), 0);
                        executeAction(stringIDToTypeID('matchColor'), desc, DialogModes.NO);
                        matched++;
                    } catch(e) {
                        skipped++;
                    }
                }
                d.activeLayer = origActive;
                matched + '|' + skipped;
            }
        }
    """
    result = _ps_js(js).strip()
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
    """将可见图层的整体内容居中到画布正中。"""
    only_vis = "true" if only_visible else "false"
    js = """
        var onlyVisible = """ + only_vis + """;
        if (app.documents.length === 0) { 'NO_DOCS'; }
        else {
            var d = app.activeDocument;
            var docW = d.width.value;
            var docH = d.height.value;
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
    """批量给图层添加圆角效果。"""
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
                    var sel = [[l, t+r], [l, t], [l+w, t], [l+w, t+r],
                               [l+w, t+h-r], [l+w, t+h], [l, t+h], [l, t+h-r]];
                    d.selection.select(sel);
                    d.selection.feather(r);
                    d.selection.invert();
                    d.selection.clear();
                    d.selection.deselect();
                    processed++;
                } catch(e) {}
            }
            processed + '';
        }
    """
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    processed = int(result) if result.isdigit() else 0
    return {"result": "ok", "processed": processed, "radius": radius,
            "message": f"已对 {processed} 个图层添加 {radius}px 圆角"}


@mcp.tool()
def ps_auto_layout_strip(direction: str = "horizontal",
                         gap: float = 10,
                         background: bool = True) -> dict:
    """将多个图层排列成横向/纵向条带（长图）。"""
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
                var oldW = d.width.value;
                var oldH = d.height.value;
                d.resizeCanvas(canvasW, canvasH, AnchorPosition.TOPLEFT);
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
    result = _ps_js(js).strip()
    if result == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    if result == 'TOO_FEW':
        return {"result": "error", "error": "至少需要 2 个图层"}
    return {"result": "ok", "canvas_size": result,
            "message": f"已排列 {direction} 条带，画布尺寸 {result}"}


@mcp.tool()
def ps_export_layer_comps(output_folder: str = "") -> dict:
    """一键导出所有图层组合（Layer Comps）为单独文件。"""
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
    result = _ps_js(js).strip()
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


# ============================================================
# ILLUSTRATOR TOOLS（COM 调用翻译为 ExtendScript）
# ============================================================

@mcp.tool()
def ai_create_document(width: float = 612, height: float = 792) -> dict:
    """新建 Illustrator 文档（单位：点）。"""
    js = f"""
        var doc = app.documents.add(DocumentColorSpace.RGB, {width}, {height});
        doc.name + '|' + doc.width + '|' + doc.height;
    """
    r = _ai_js(js)
    parts = r.split('|')
    return {"result": "ok", "name": parts[0] if parts else "", "width": width, "height": height, "unit": "points"}


@mcp.tool()
def ai_get_active_info() -> dict:
    """获取当前文档信息：名称、尺寸、路径数、文字框数、打开文档数。"""
    js = """
        if (app.documents.length === 0) { 'NO_DOCS|0'; }
        else {
            var d = app.activeDocument;
            d.name + '|' + d.width + '|' + d.height + '|' + d.pathItems.length
                + '|' + d.textFrames.length + '|' + app.documents.length;
        }
    """
    r = _ai_js(js)
    parts = r.split('|')
    if parts[0] == 'NO_DOCS':
        return {"result": "ok", "active": False, "open": int(parts[1]) if len(parts) > 1 else 0}
    return {
        "result": "ok",
        "name": parts[0],
        "width": float(parts[1]) if len(parts) > 1 else 0,
        "height": float(parts[2]) if len(parts) > 2 else 0,
        "paths": int(float(parts[3])) if len(parts) > 3 else 0,
        "textFrames": int(float(parts[4])) if len(parts) > 4 else 0,
        "open": int(float(parts[5])) if len(parts) > 5 else 0,
    }


@mcp.tool()
def ai_list_documents() -> dict:
    """列出所有打开的 Illustrator 文档。"""
    js = """
        var out = [];
        for (var i = 0; i < app.documents.length; i++) out.push(app.documents[i].name);
        JSON.stringify(out);
    """
    names = _parse(_ai_js(js))
    if not isinstance(names, list):
        names = []
    return {"result": "ok", "count": len(names), "documents": [{"name": n} for n in names]}


@mcp.tool()
def ai_close_document() -> dict:
    """关闭当前 Illustrator 文档（不保存）。"""
    js = """
        if (app.documents.length === 0) { 'NONE'; }
        else { var n = app.activeDocument.name; app.activeDocument.close(SaveOptions.DONOTSAVECHANGES); n; }
    """
    r = _ai_js(js).strip()
    if r == 'NONE':
        return {"result": "ok", "closed": False}
    return {"result": "ok", "closed": True, "name": r}


@mcp.tool()
def ai_add_rectangle(x: float = 100, y: float = 400, w: float = 200, h: float = 150,
                     fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """画矩形（左下角 x,y；宽高 w,h；可选填充色）。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var rect = d.pathItems.rectangle({y}, {x}, {w}, {h});
            if ({fr} >= 0 && {fg} >= 0 && {fb} >= 0) {{
                rect.filled = true;
                var c = new RGBColor(); c.red = {fr}; c.green = {fg}; c.blue = {fb};
                rect.fillColor = c;
            }}
            rect.name;
        }}
    """
    r = _ai_js(js).strip()
    if r == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "name": r, "x": x, "y": y, "width": w, "height": h}


@mcp.tool()
def ai_add_ellipse(x: float = 200, y: float = 500, w: float = 150, h: float = 100,
                   fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """画椭圆（左上角 x,y；宽高 w,h；可选填充色）。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var ell = d.pathItems.ellipse({y}, {x}, {w}, {h});
            if ({fr} >= 0 && {fg} >= 0 && {fb} >= 0) {{
                ell.filled = true;
                var c = new RGBColor(); c.red = {fr}; c.green = {fg}; c.blue = {fb};
                ell.fillColor = c;
            }}
            ell.name;
        }}
    """
    r = _ai_js(js).strip()
    if r == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "name": r, "x": x, "y": y, "width": w, "height": h}


@mcp.tool()
def ai_add_polygon(x: float = 300, y: float = 400, radius: float = 100, sides: int = 6,
                   fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """画正多边形（中心点 x,y；半径 radius；边数 sides）。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var poly = d.pathItems.polygon({y}, {x}, {radius}, {sides});
            if ({fr} >= 0 && {fg} >= 0 && {fb} >= 0) {{
                poly.filled = true;
                var c = new RGBColor(); c.red = {fr}; c.green = {fg}; c.blue = {fb};
                poly.fillColor = c;
            }}
            poly.name;
        }}
    """
    r = _ai_js(js).strip()
    if r == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "name": r, "x": x, "y": y, "radius": radius, "sides": sides}


@mcp.tool()
def ai_add_text(text: str, size: float = 24, x: float = 100, y: float = 300,
                fr: int = -1, fg: int = -1, fb: int = -1) -> dict:
    """添加文字（内容、字号、位置、可选颜色）。"""
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var tf = d.textFrames.add();
            tf.contents = {json.dumps(text)};
            tf.textRange.characterAttributes.size = {size};
            tf.top = {y};
            tf.left = {x};
            if ({fr} >= 0 && {fg} >= 0 && {fb} >= 0) {{
                var c = new RGBColor(); c.red = {fr}; c.green = {fg}; c.blue = {fb};
                tf.textRange.characterAttributes.fillColor = c;
            }}
            'ok';
        }}
    """
    r = _ai_js(js).strip()
    if r == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "text": text, "size": size, "x": x, "y": y}


@mcp.tool()
def ai_save_as_ai(path: str = "") -> dict:
    """将当前文档保存为 .ai。"""
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.ai")
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{ app.activeDocument.saveAs(new File({json.dumps(path)})); 'ok'; }}
    """
    r = _ai_js(js).strip()
    if r == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "path": path}


@mcp.tool()
def ai_export_svg(path: str = "") -> dict:
    """导出 SVG 矢量图。"""
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.svg")
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var f = new File({json.dumps(path)});
            var o = new ExportOptionsSVG();
            o.fontType = SVGFontType.OUTLINEFONT;
            d.exportFile(f, ExportType.SVG, o);
            'ok';
        }}
    """
    r = _ai_js(js).strip()
    if r == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "path": path}


@mcp.tool()
def ai_export_png(path: str = "") -> dict:
    """导出 PNG（透明背景）。"""
    if not path:
        path = os.path.join(os.path.expanduser("~"), "Documents", "adobe_mcp_out.png")
    js = f"""
        if (app.documents.length === 0) {{ 'NO_DOCS'; }}
        else {{
            var d = app.activeDocument;
            var f = new File({json.dumps(path)});
            var o = new ExportOptionsPNG24();
            o.antiAliasing = true;
            o.transparency = true;
            d.exportFile(f, ExportType.PNG24, o);
            'ok';
        }}
    """
    r = _ai_js(js).strip()
    if r == 'NO_DOCS':
        return {"result": "error", "error": "no document"}
    return {"result": "ok", "path": path}


@mcp.tool()
def ai_batch_replace_text(find_text: str, replace_text: str,
                          case_sensitive: bool = False) -> dict:
    """批量查找替换所有文字框中的文字。"""
    cs = "true" if case_sensitive else "false"
    js = f"""
        var d = app.activeDocument;
        var count = 0;
        var replaced = 0;
        function ciReplace(s, find, rep) {{
            if (find === '') return s;
            var result = '';
            var lower = s.toLowerCase();
            var f = find.toLowerCase();
            var i = 0;
            while (i < s.length) {{
                var j = lower.indexOf(f, i);
                if (j < 0) {{ result += s.substring(i); break; }}
                result += s.substring(i, j) + rep;
                i = j + find.length;
            }}
            return result;
        }}
        for (var k = 0; k < d.textFrames.length; k++) {{
            var tf = d.textFrames[k];
            var content = tf.contents;
            count++;
            var newContent;
            if ({cs}) {{
                newContent = content.split({json.dumps(find_text)}).join({json.dumps(replace_text)});
            }} else {{
                newContent = ciReplace(content, {json.dumps(find_text)}, {json.dumps(replace_text)});
            }}
            if (newContent !== content) {{ tf.contents = newContent; replaced++; }}
        }}
        String(replaced) + '|' + String(count);
    """
    r = _ai_js(js).strip()
    parts = r.split('|')
    replaced = int(float(parts[0])) if parts and parts[0].strip() else 0
    total = int(float(parts[1])) if len(parts) > 1 and parts[1].strip() else 0
    return {"result": "ok", "total_frames": total, "replaced": replaced,
            "message": f"在 {replaced}/{total} 个文字框中替换了 '{find_text}' → '{replace_text}'"}


@mcp.tool()
def ai_export_all_artboards(output_folder: str = "", format: str = "png") -> dict:
    """一键导出所有画板为单独文件。"""
    if not output_folder:
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop", "artboards_export")
    os.makedirs(output_folder, exist_ok=True)
    fmt = format.lower().strip()
    output_folder = output_folder.replace("\\", "/")
    js = f"""
        var d = app.activeDocument;
        var total = d.artboards.length;
        var folder = new Folder('{output_folder}');
        if (!folder.exists) folder.create();
        var exported = 0;
        for (var i = 0; i < total; i++) {{
            try {{
                d.artboards.setActiveArtboardIndex(i);
                var n = ('000' + (i+1)).slice(-3);
                if ('{fmt}' === 'png') {{
                    var f = new File(folder.fsName + '/artboard_' + n + '.png');
                    var o = new ExportOptionsPNG24(); o.antiAliasing = true; o.transparency = true;
                    d.exportFile(f, ExportType.PNG24, o);
                }} else if ('{fmt}' === 'svg') {{
                    var f = new File(folder.fsName + '/artboard_' + n + '.svg');
                    var s = new ExportOptionsSVG(); s.fontType = SVGFontType.OUTLINEFONT;
                    d.exportFile(f, ExportType.SVG, s);
                }} else {{
                    var f = new File(folder.fsName + '/artboard_' + n + '.pdf');
                    var p = new ExportOptionsPDF();
                    d.exportFile(f, ExportType.PDF, p);
                }}
                exported++;
            }} catch(e) {{}}
        }}
        String(exported);
    """
    r = _ai_js(js).strip()
    exported = int(float(r)) if r.replace('.', '', 1).isdigit() else 0
    return {"result": "ok", "exported": exported, "total": 0, "folder": output_folder, "format": fmt,
            "message": f"已导出 {exported} 个画板为 {fmt.upper()} 文件"}


@mcp.tool()
def ai_auto_layout_grid(rows: int = 3, cols: int = 3, gap: float = 10) -> dict:
    """将文档中所有对象排列成网格布局。"""
    js = f"""
        var d = app.activeDocument;
        var items = d.pageItems;
        var infos = [];
        for (var i = 0; i < items.length; i++) {{
            var it = items[i];
            try {{ infos.push({{item: it, left: it.left, top: it.top, w: it.width, h: it.height}}); }} catch(e) {{}}
        }}
        if (infos.length < 2) {{ 'TOO_FEW'; }}
        else {{
            var maxW = 0, maxH = 0;
            for (var j = 0; j < infos.length; j++) {{
                if (infos[j].w > maxW) maxW = infos[j].w;
                if (infos[j].h > maxH) maxH = infos[j].h;
            }}
            var cellW = maxW + {gap};
            var cellH = maxH + {gap};
            var startLeft = 50;
            var startTop = 800;
            var cols = {cols};
            var arranged = 0;
            for (var k = 0; k < infos.length; k++) {{
                var row = Math.floor(k / cols);
                var col = k % cols;
                var tl = startLeft + col * cellW;
                var tt = startTop - row * cellH;
                try {{
                    infos[k].item.translate(tl - infos[k].left, tt - infos[k].top);
                    arranged++;
                }} catch(e) {{}}
            }}
            String(arranged);
        }}
    """
    r = _ai_js(js).strip()
    if r == 'TOO_FEW':
        return {"result": "error", "error": "至少需要 2 个对象"}
    arranged = int(float(r)) if r.replace('.', '', 1).isdigit() else 0
    grid_desc = f"{rows}×{cols}"
    return {"result": "ok", "arranged": arranged, "grid": grid_desc,
            "message": f"已将 {arranged} 个对象排列为 {grid_desc} 网格"}


@mcp.tool()
def ai_align_objects(direction: str = "left") -> dict:
    """将选中对象按指定方式对齐。"""
    direction = direction.lower().strip()
    js = f"""
        var d = app.activeDocument;
        var sel = d.selection;
        if (sel.length < 2) {{ 'TOO_FEW'; }}
        else {{
            var infos = [];
            for (var i = 0; i < sel.length; i++) {{
                var it = sel[i];
                try {{ infos.push({{item: it, left: it.left, top: it.top, w: it.width, h: it.height}}); }} catch(e) {{}}
            }}
            if (infos.length < 2) {{ 'TOO_FEW'; }}
            else {{
                var dir = '{direction}';
                var aligned = 0;
                if (dir === 'left') {{
                    var t = 999999; for (var a = 0; a < infos.length; a++) if (infos[a].left < t) t = infos[a].left;
                    for (var b = 0; b < infos.length; b++) {{ infos[b].item.translate(t - infos[b].left, 0); aligned++; }}
                }} else if (dir === 'right') {{
                    var mr = -999999; for (var c = 0; c < infos.length; c++) {{ var e = infos[c].left + infos[c].w; if (e > mr) mr = e; }}
                    for (var d2 = 0; d2 < infos.length; d2++) {{ infos[d2].item.translate((mr - infos[d2].w) - infos[d2].left, 0); aligned++; }}
                }} else if (dir === 'top') {{
                    var tp = 999999; for (var e2 = 0; e2 < infos.length; e2++) if (infos[e2].top < tp) tp = infos[e2].top;
                    for (var f = 0; f < infos.length; f++) {{ infos[f].item.translate(0, tp - infos[f].top); aligned++; }}
                }} else if (dir === 'bottom') {{
                    var mb = -999999; for (var g = 0; g < infos.length; g++) {{ var e3 = infos[g].top + infos[g].h; if (e3 > mb) mb = e3; }}
                    for (var h = 0; h < infos.length; h++) {{ infos[h].item.translate(0, (mb - infos[h].h) - infos[h].top); aligned++; }}
                }} else if (dir === 'center_h' || dir === 'center_horizontal') {{
                    var sum = 0; for (var m = 0; m < infos.length; m++) sum += infos[m].left + infos[m].w / 2;
                    var avg = sum / infos.length;
                    for (var n = 0; n < infos.length; n++) {{ infos[n].item.translate((avg - infos[n].w / 2) - infos[n].left, 0); aligned++; }}
                }} else if (dir === 'center_v' || dir === 'center_vertical') {{
                    var sum2 = 0; for (var p = 0; p < infos.length; p++) sum2 += infos[p].top + infos[p].h / 2;
                    var avg2 = sum2 / infos.length;
                    for (var q = 0; q < infos.length; q++) {{ infos[q].item.translate(0, (avg2 - infos[q].h / 2) - infos[q].top); aligned++; }}
                }} else {{
                    'BAD_DIR';
                }}
                String(aligned);
            }}
        }}
    """
    r = _ai_js(js).strip()
    if r == 'TOO_FEW':
        return {"result": "error", "error": "请先选中至少 2 个对象"}
    if r == 'BAD_DIR':
        return {"result": "error", "error": f"不支持的对齐方式: {direction}"}
    aligned = int(float(r)) if r.replace('.', '', 1).isdigit() else 0
    return {"result": "ok", "aligned": aligned, "direction": direction,
            "message": f"已将 {aligned} 个对象{direction}对齐"}


if __name__ == "__main__":
    print("[adobe-mac-mcp] entering MCP loop...", file=sys.stderr, flush=True)
    mcp.run(transport="stdio")
