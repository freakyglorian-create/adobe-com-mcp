#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Adobe COM MCP Server — macOS 版（AppleScript 后端）

与 Windows 版 server.py 工具名完全一致（ps_* / ai_*），客户端无需改配置。
原理：通过 macOS 自带的 osascript 运行 AppleScript，再借 Photoshop / Illustrator
的 `do javascript` 命令执行 ExtendScript 完成自动化。无需安装任何额外依赖
（只需 Python 3.10+ 和 `pip install "mcp<2"`）。

⚠️ 注意：本文件在 Windows 上无法运行/验证，请在 Mac 上实测调试。
"""

import json
import os
import subprocess
import tempfile

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("adobe-mac-mcp")


# --------------------------------------------------------------------------- #
# 底层：AppleScript / ExtendScript 执行
# --------------------------------------------------------------------------- #

def _osascript(applescript: str) -> str:
    """运行一段 AppleScript，返回 stdout（去尾随换行）。失败时抛 RuntimeError。"""
    proc = subprocess.run(
        ["osascript", "-e", applescript],
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "osascript 执行失败")
    return proc.stdout.strip()


def _js_str(s) -> str:
    """把 Python 字符串转义成 ExtendScript 单引号字符串字面量（去掉危险字符）。"""
    return (
        str(s)
        .replace("\\", "\\\\")
        .replace("'", "\\'")
        .replace("\n", "\\n")
        .replace("\r", "")
    )


def _ps_js(js: str) -> str:
    """
    通过 Photoshop 的 `do javascript` 执行 ExtendScript 字符串，返回结果。
    AppleScript 字符串里反斜杠是字面量（无需二次转义），所以只需：
      1) 真实换行 -> 字面 \\n（交给 ExtendScript 解释回换行）
      2) 双引号 -> \\"（AppleScript 字符串定界符；本文件 JS 统一用单引号，基本不触发）
    """
    single_line = (
        js.replace("\r", "")
        .replace("\n", "\\n")
        .replace('"', '\\"')
    )
    script = f'tell application id "com.adobe.Photoshop"\n\tdo javascript "{single_line}"\nend tell'
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
        "end tell"
    )
    try:
        return _osascript(script)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def _parse(result: str):
    """尽力把 ExtendScript 返回的字符串解析成 JSON，失败则原样返回字符串。"""
    try:
        return json.loads(result)
    except (json.JSONDecodeError, TypeError):
        return result


# --------------------------------------------------------------------------- #
# Photoshop 工具（ps_*）
# --------------------------------------------------------------------------- #

@mcp.tool()
def ps_create_document(width: int, height: int, name: str = "Untitled") -> dict:
    """新建指定像素尺寸的 Photoshop 文档。"""
    js = f"""
    app.preferences.rulerUnits = Units.PIXELS;
    var doc = app.documents.add({width}, {height}, 72, '{_js_str(name)}');
    doc.name + '|' + doc.width.value + '|' + doc.height.value;
    """
    r = _parse(_ps_js(js))
    if isinstance(r, str) and "|" in r:
        parts = r.split("|")
        return {"result": "ok", "name": parts[0], "width": parts[1], "height": parts[2]}
    return {"result": "ok", "raw": r}


@mcp.tool()
def ps_get_active_info() -> dict:
    """获取当前文档信息（名称/尺寸/图层数）。"""
    js = """
    var d = app.activeDocument;
    d.name + '|' + d.width.value + '|' + d.height.value + '|' + d.layers.length;
    """
    r = _parse(_ps_js(js))
    if isinstance(r, str) and "|" in r:
        parts = r.split("|")
        return {"result": "ok", "name": parts[0], "width": parts[1],
                "height": parts[2], "layers": parts[3]}
    return {"result": "ok", "raw": r}


@mcp.tool()
def ps_list_documents() -> dict:
    """列出所有打开的文档。"""
    js = """
    var out = [];
    for (var i = 0; i < app.documents.length; i++) out.push(app.documents[i].name);
    JSON.stringify(out);
    """
    return {"result": "ok", "documents": _parse(_ps_js(js))}


@mcp.tool()
def ps_close_document() -> dict:
    """关闭当前文档（不保存）。"""
    _ps_js("app.activeDocument.close(SaveOptions.DONOTSAVECHANGES); 'ok';")
    return {"result": "ok"}


@mcp.tool()
def ps_add_layer(name: str = "Layer") -> dict:
    """新建空白图层。"""
    js = f"""
    var l = app.activeDocument.artLayers.add();
    l.name = '{_js_str(name)}';
    'ok';
    """
    _ps_js(js)
    return {"result": "ok"}


@mcp.tool()
def ps_add_text_layer(text: str, size: int = 24, x: int = 50, y: int = 50,
                      r: int = 0, g: int = 0, b: int = 0) -> dict:
    """添加文字图层（可指定字号/颜色/位置）。"""
    js = f"""
    var d = app.activeDocument;
    var layer = d.artLayers.add();
    layer.kind = LayerKind.TEXT;
    var ti = layer.textItem;
    ti.contents = '{_js_str(text)}';
    ti.size = {size};
    ti.position = [{x}, {y}];
    var c = new SolidColor();
    c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
    ti.color = c;
    layer.name = '{_js_str(text)}';
    'ok';
    """
    _ps_js(js)
    return {"result": "ok"}


@mcp.tool()
def ps_fill_layer(r: int, g: int, b: int) -> dict:
    """用指定颜色填充当前图层。"""
    js = f"""
    var d = app.activeDocument;
    d.selection.selectAll();
    var c = new SolidColor();
    c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
    d.selection.fill(c);
    d.selection.deselect();
    'ok';
    """
    _ps_js(js)
    return {"result": "ok"}


@mcp.tool()
def ps_set_layer_opacity(opacity: int) -> dict:
    """设置当前图层不透明度（0-100）。"""
    _ps_js(f"app.activeDocument.activeLayer.opacity = {opacity}; 'ok';")
    return {"result": "ok"}


@mcp.tool()
def ps_set_foreground_color(r: int, g: int, b: int) -> dict:
    """设置前景色。"""
    js = f"""
    var c = new SolidColor();
    c.rgb.red = {r}; c.rgb.green = {g}; c.rgb.blue = {b};
    app.foregroundColor = c;
    'ok';
    """
    _ps_js(js)
    return {"result": "ok"}


@mcp.tool()
def ps_resize_document(width: int, height: int) -> dict:
    """调整文档像素尺寸。"""
    js = f"""
    app.preferences.rulerUnits = Units.PIXELS;
    app.activeDocument.resizeImage({width}, {height}, null, ResampleMethod.BICUBIC);
    'ok';
    """
    _ps_js(js)
    return {"result": "ok"}


@mcp.tool()
def ps_save_as_png(path: str) -> dict:
    """导出当前文档为 PNG。"""
    js = f"""
    var d = app.activeDocument;
    var f = new File('{_js_str(path)}');
    d.saveAs(f, new PNGSaveOptions(), true, Extension.LOWERCASE);
    'ok';
    """
    _ps_js(js)
    return {"result": "ok", "path": path}


@mcp.tool()
def ps_save_as_jpg(path: str, quality: int = 8) -> dict:
    """导出当前文档为 JPG（quality 1-12）。"""
    js = f"""
    var d = app.activeDocument;
    var f = new File('{_js_str(path)}');
    var o = new JPEGSaveOptions();
    o.quality = {quality};
    d.saveAs(f, o, true, Extension.LOWERCASE);
    'ok';
    """
    _ps_js(js)
    return {"result": "ok", "path": path}


@mcp.tool()
def ps_list_all_layers() -> dict:
    """列出文档所有图层（名称/可见性）。"""
    js = """
    var d = app.activeDocument;
    var out = [];
    for (var i = 0; i < d.layers.length; i++) {
        out.push(d.layers[i].name + '\\t' + d.layers[i].visible);
    }
    JSON.stringify(out);
    """
    raw = _parse(_ps_js(js))
    layers = []
    if isinstance(raw, list):
        for item in raw:
            parts = str(item).split("\t")
            layers.append({"name": parts[0], "visible": parts[1] == "true" if len(parts) > 1 else True})
    return {"result": "ok", "layers": layers}


@mcp.tool()
def ps_export_all_layers_to_png(folder: str) -> dict:
    """一键导出所有图层为单独 PNG（隐藏其它图层逐个导出）。"""
    js = """
    var d = app.activeDocument;
    var folder = new Folder('__FOLDER__');
    var vis = [];
    for (var i = 0; i < d.layers.length; i++) vis.push(d.layers[i].visible);
    var out = [];
    for (var i = 0; i < d.layers.length; i++) {
        for (var j = 0; j < d.layers.length; j++) d.layers[j].visible = false;
        d.layers[i].visible = true;
        var f = new File(folder.fsName + '/' + d.layers[i].name + '.png');
        d.saveAs(f, new PNGSaveOptions(), true, Extension.LOWERCASE);
        out.push(f.fsName);
    }
    for (var k = 0; k < d.layers.length; k++) d.layers[k].visible = vis[k];
    JSON.stringify(out);
    """.replace("__FOLDER__", _js_str(folder))
    r = _parse(_ps_js(js))
    return {"result": "ok", "files": r}


# --------------------------------------------------------------------------- #
# Illustrator 工具（ai_*）
# --------------------------------------------------------------------------- #

@mcp.tool()
def ai_create_document(width: int, height: int, name: str = "Untitled") -> dict:
    """新建指定尺寸（点）的 Illustrator 文档。"""
    js = f"""
    var doc = app.documents.add(DocumentColorSpace.RGB, {width}, {height});
    doc.name + '|' + doc.width + '|' + doc.height;
    """
    r = _parse(_ai_js(js))
    if isinstance(r, str) and "|" in r:
        parts = r.split("|")
        return {"result": "ok", "name": parts[0], "width": parts[1], "height": parts[2]}
    return {"result": "ok", "raw": r}


@mcp.tool()
def ai_get_active_info() -> dict:
    """获取当前文档信息。"""
    js = """
    var d = app.activeDocument;
    d.name + '|' + d.width + '|' + d.height + '|' + d.artboards.length;
    """
    r = _parse(_ai_js(js))
    if isinstance(r, str) and "|" in r:
        parts = r.split("|")
        return {"result": "ok", "name": parts[0], "width": parts[1],
                "height": parts[2], "artboards": parts[3]}
    return {"result": "ok", "raw": r}


@mcp.tool()
def ai_list_documents() -> dict:
    """列出所有打开的文档。"""
    js = """
    var out = [];
    for (var i = 0; i < app.documents.length; i++) out.push(app.documents[i].name);
    JSON.stringify(out);
    """
    return {"result": "ok", "documents": _parse(_ai_js(js))}


@mcp.tool()
def ai_close_document() -> dict:
    """关闭当前文档（不保存）。"""
    _ai_js("app.activeDocument.close(SaveOptions.DONOTSAVECHANGES); 'ok';")
    return {"result": "ok"}


@mcp.tool()
def ai_add_rectangle(x: int, y: int, width: int, height: int,
                     r: int = 0, g: int = 0, b: int = 0) -> dict:
    """画矩形（pathItems.rectangle 参数为 top/left/width/height）。"""
    js = f"""
    var doc = app.activeDocument;
    var rect = doc.pathItems.rectangle({y}, {x}, {width}, {height});
    rect.filled = true;
    var c = new RGBColor(); c.red = {r}; c.green = {g}; c.blue = {b};
    rect.fillColor = c;
    'ok';
    """
    _ai_js(js)
    return {"result": "ok"}


@mcp.tool()
def ai_add_ellipse(x: int, y: int, width: int, height: int,
                   r: int = 0, g: int = 0, b: int = 0) -> dict:
    """画椭圆。"""
    js = f"""
    var doc = app.activeDocument;
    var e = doc.pathItems.ellipse({y}, {x}, {width}, {height});
    e.filled = true;
    var c = new RGBColor(); c.red = {r}; c.green = {g}; c.blue = {b};
    e.fillColor = c;
    'ok';
    """
    _ai_js(js)
    return {"result": "ok"}


@mcp.tool()
def ai_add_polygon(cx: int, cy: int, radius: int, sides: int = 6,
                   r: int = 0, g: int = 0, b: int = 0) -> dict:
    """画正多边形。"""
    js = f"""
    var doc = app.activeDocument;
    var p = doc.pathItems.polygon({cx}, {cy}, {radius}, {sides});
    p.filled = true;
    var c = new RGBColor(); c.red = {r}; c.green = {g}; c.blue = {b};
    p.fillColor = c;
    'ok';
    """
    _ai_js(js)
    return {"result": "ok"}


@mcp.tool()
def ai_add_text(text: str, x: int = 0, y: int = 0, size: int = 24,
                r: int = 0, g: int = 0, b: int = 0) -> dict:
    """添加文字。"""
    js = f"""
    var doc = app.activeDocument;
    var tf = doc.textFrames.add();
    tf.contents = '{_js_str(text)}';
    tf.position = [{x}, {y}];
    var attrs = tf.textRange.characterAttributes;
    attrs.size = {size};
    var c = new RGBColor(); c.red = {r}; c.green = {g}; c.blue = {b};
    attrs.fillColor = c;
    'ok';
    """
    _ai_js(js)
    return {"result": "ok"}


@mcp.tool()
def ai_export_svg(path: str) -> dict:
    """导出 SVG 矢量图。"""
    js = f"""
    var doc = app.activeDocument;
    var f = new File('{_js_str(path)}');
    doc.exportFile(f, ExportType.SVG, new ExportOptionsSVG());
    'ok';
    """
    _ai_js(js)
    return {"result": "ok", "path": path}


@mcp.tool()
def ai_export_png(path: str) -> dict:
    """导出 PNG（透明背景）。"""
    js = f"""
    var doc = app.activeDocument;
    var f = new File('{_js_str(path)}');
    var o = new ExportOptionsPNG24();
    o.transparency = true;
    o.antiAliasing = true;
    doc.exportFile(f, ExportType.PNG24, o);
    'ok';
    """
    _ai_js(js)
    return {"result": "ok", "path": path}


@mcp.tool()
def ai_batch_replace_text(find: str, replace: str) -> dict:
    """批量查找替换所有文字框内容。"""
    js = f"""
    var doc = app.activeDocument;
    var count = 0;
    for (var i = 0; i < doc.textFrames.length; i++) {{
        var tf = doc.textFrames[i];
        if (tf.contents.indexOf('{_js_str(find)}') >= 0) {{
            tf.contents = tf.contents.split('{_js_str(find)}').join('{_js_str(replace)}');
            count++;
        }}
    }}
    String(count);
    """
    r = _parse(_ai_js(js))
    return {"result": "ok", "replaced": r}


if __name__ == "__main__":
    mcp.run()
