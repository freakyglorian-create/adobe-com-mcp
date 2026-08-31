#!/usr/bin/env node
// Adobe COM MCP Server — Photoshop + Illustrator (Windows).
// Drives both apps via COM automation. No UXP / proxy / upgrade needed.
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PS_SCRIPT = join(__dirname, "photoshop.ps1");
const AI_SCRIPT = join(__dirname, "illustrator.ps1");
const POWERSHELL = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";

function runPs(script, env) {
  const out = execFileSync(
    POWERSHELL,
    ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script],
    {
      encoding: "utf8",
      timeout: 120000,
      maxBuffer: 16 * 1024 * 1024,
      env: { ...process.env, ...env },
      windowsHide: true,
    }
  );
  return out.trim();
}

function num(v, dflt) {
  const n = Number(v);
  return Number.isFinite(n) ? n : dflt;
}
function int(v, dflt) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : dflt;
}
function str(v, dflt = "") {
  return v == null ? dflt : String(v);
}

const tools = [
  // =========================
  // PHOTOSHOP TOOLS (ps_ prefix)
  // =========================
  {
    name: "ps_create_document",
    description:
      "在 Photoshop 中新建指定像素尺寸的文档。" +
      "当用户说“新建一个 800x800 的 Photoshop 文档”“建一个 1920x1080 的画布”时使用。" +
      "width 宽度像素，height 高度像素（默认等于 width），name 文档名可选。",
    inputSchema: {
      type: "object",
      properties: {
        width: { type: "integer", description: "宽度（像素）" },
        height: { type: "integer", description: "高度（像素，默认等于宽度）" },
        name: { type: "string", description: "文档名，可选" },
      },
      required: ["width"],
    },
  },
  {
    name: "ps_get_active_info",
    description: "获取 Photoshop 当前活动文档信息：名称、像素尺寸、分辨率、颜色模式、图层数、活动图层名。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ps_list_documents",
    description: "列出 Photoshop 当前打开的所有文档名称。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ps_close_document",
    description: "关闭 Photoshop 当前活动文档（不保存）。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ps_open_document",
    description: "在 Photoshop 中打开一个本地图片/PSD 文件。传入文件完整路径。",
    inputSchema: {
      type: "object",
      properties: { path: { type: "string", description: "本地文件完整路径" } },
      required: ["path"],
    },
  },
  {
    name: "ps_save_as_psd",
    description: "将当前 Photoshop 文档另存为 PSD 格式。path 为保存路径，可选。",
    inputSchema: {
      type: "object",
      properties: { path: { type: "string", description: "保存路径（含 .psd 后缀），可选" } },
    },
  },
  {
    name: "ps_save_as_png",
    description: "将当前 Photoshop 文档导出为 PNG 图片。path 为保存路径，可选。",
    inputSchema: {
      type: "object",
      properties: { path: { type: "string", description: "保存路径（含 .png 后缀），可选" } },
    },
  },
  {
    name: "ps_save_as_jpg",
    description: "将当前 Photoshop 文档导出为 JPG 图片。quality 范围 1-12，默认 8。",
    inputSchema: {
      type: "object",
      properties: {
        path: { type: "string", description: "保存路径（含 .jpg 后缀），可选" },
        quality: { type: "integer", description: "质量 1-12，默认 8" },
      },
    },
  },
  {
    name: "ps_resize_document",
    description: "调整当前 Photoshop 文档的像素尺寸。给 width 或 height 中任意一个，另一个按比例自动算。",
    inputSchema: {
      type: "object",
      properties: {
        width: { type: "integer", description: "目标宽度（像素）" },
        height: { type: "integer", description: "目标高度（像素）" },
      },
    },
  },
  {
    name: "ps_add_layer",
    description: "在当前 Photoshop 文档中新建一个空白图层。name 可选。",
    inputSchema: {
      type: "object",
      properties: { name: { type: "string", description: "图层名，可选" } },
    },
  },
  {
    name: "ps_duplicate_layer",
    description: "复制当前活动图层，生成副本。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ps_delete_layer",
    description: "删除当前活动图层。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ps_set_layer_opacity",
    description: "设置当前活动图层的不透明度（0-100）。",
    inputSchema: {
      type: "object",
      properties: { opacity: { type: "integer", description: "不透明度 0-100" } },
      required: ["opacity"],
    },
  },
  {
    name: "ps_add_text_layer",
    description:
      "在当前 Photoshop 文档中添加文字图层。text 为文字内容，size 字号（像素），" +
      "x/y 为左上角位置（像素），r/g/b 为文字颜色 RGB。",
    inputSchema: {
      type: "object",
      properties: {
        text: { type: "string", description: "文字内容" },
        size: { type: "number", description: "字号（像素），默认 48" },
        x: { type: "number", description: "X 坐标（像素，左上角），默认 50" },
        y: { type: "number", description: "Y 坐标（像素，左上角），默认 100" },
        r: { type: "integer", description: "红色 0-255，默认 0" },
        g: { type: "integer", description: "绿色 0-255，默认 0" },
        b: { type: "integer", description: "蓝色 0-255，默认 0" },
      },
      required: ["text"],
    },
  },
  {
    name: "ps_set_foreground_color",
    description: "设置 Photoshop 前景色（RGB）。r/g/b 取值 0-255。",
    inputSchema: {
      type: "object",
      properties: {
        r: { type: "integer", description: "红 0-255" },
        g: { type: "integer", description: "绿 0-255" },
        b: { type: "integer", description: "蓝 0-255" },
      },
      required: ["r", "g", "b"],
    },
  },
  {
    name: "ps_fill_layer",
    description: "用指定颜色填充当前活动图层（全画布填充）。r/g/b 为 RGB 值。",
    inputSchema: {
      type: "object",
      properties: {
        r: { type: "integer", description: "红 0-255" },
        g: { type: "integer", description: "绿 0-255" },
        b: { type: "integer", description: "蓝 0-255" },
      },
      required: ["r", "g", "b"],
    },
  },
  {
    name: "ps_apply_gaussian_blur",
    description: "对当前活动图层应用高斯模糊滤镜。radius 为模糊半径（像素），默认 5。",
    inputSchema: {
      type: "object",
      properties: { radius: { type: "number", description: "模糊半径（像素），默认 5" } },
    },
  },
  {
    name: "ps_apply_unsharp_mask",
    description: "对当前活动图层应用 USM 锐化（非锐化蒙版）。amount 数量 1-500，radius 半径，threshold 阈值。",
    inputSchema: {
      type: "object",
      properties: {
        amount: { type: "number", description: "数量%，默认 100" },
        radius: { type: "number", description: "半径像素，默认 2" },
        threshold: { type: "number", description: "阈值色阶 0-255，默认 0" },
      },
    },
  },
  {
    name: "ps_do_action",
    description: "在 Photoshop 中执行一个已安装的动作（Action）。action 为动作名，from 为动作集名（可选）。",
    inputSchema: {
      type: "object",
      properties: {
        action: { type: "string", description: "动作名称" },
        from: { type: "string", description: "动作所在动作集名称，可选" },
      },
      required: ["action"],
    },
  },

  // =========================
  // ILLUSTRATOR TOOLS (ai_ prefix)
  // =========================
  {
    name: "ai_create_document",
    description: "在 Illustrator 中新建一个文档，单位为点（pt，72pt = 1 英寸）。width/height 为尺寸。",
    inputSchema: {
      type: "object",
      properties: {
        width: { type: "number", description: "宽度（点 pt）" },
        height: { type: "number", description: "高度（点 pt）" },
      },
    },
  },
  {
    name: "ai_get_active_info",
    description: "获取 Illustrator 当前活动文档信息：名称、尺寸、路径数、文字框数、打开文档数。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ai_list_documents",
    description: "列出 Illustrator 当前打开的所有文档名称。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ai_close_document",
    description: "关闭 Illustrator 当前活动文档（不保存）。",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "ai_add_rectangle",
    description:
      "在当前 Illustrator 文档中画一个矩形。x/y 为左下角坐标（点），w/h 宽高，" +
      "fr/fg/fb 为填充色 RGB（可选，不填则无填充）。",
    inputSchema: {
      type: "object",
      properties: {
        x: { type: "number", description: "左下角 X 坐标（点）" },
        y: { type: "number", description: "左下角 Y 坐标（点）" },
        w: { type: "number", description: "宽度（点）" },
        h: { type: "number", description: "高度（点）" },
        fr: { type: "integer", description: "填充色 R 0-255，可选" },
        fg: { type: "integer", description: "填充色 G 0-255，可选" },
        fb: { type: "integer", description: "填充色 B 0-255，可选" },
      },
    },
  },
  {
    name: "ai_add_ellipse",
    description: "在当前 Illustrator 文档中画一个椭圆。x/y 为左上角坐标（点），w/h 宽高，fr/fg/fb 填充色。",
    inputSchema: {
      type: "object",
      properties: {
        x: { type: "number", description: "左上角 X 坐标（点）" },
        y: { type: "number", description: "左上角 Y 坐标（点）" },
        w: { type: "number", description: "宽度（点）" },
        h: { type: "number", description: "高度（点）" },
        fr: { type: "integer", description: "填充色 R 0-255，可选" },
        fg: { type: "integer", description: "填充色 G 0-255，可选" },
        fb: { type: "integer", description: "填充色 B 0-255，可选" },
      },
    },
  },
  {
    name: "ai_add_polygon",
    description:
      "在当前 Illustrator 文档中画一个正多边形。x/y 为中心点，radius 半径，" +
      "sides 边数（默认 6）。fr/fg/fb 填充色可选。",
    inputSchema: {
      type: "object",
      properties: {
        x: { type: "number", description: "中心点 X 坐标（点）" },
        y: { type: "number", description: "中心点 Y 坐标（点）" },
        radius: { type: "number", description: "半径（点）" },
        sides: { type: "integer", description: "边数（默认 6）" },
        fr: { type: "integer", description: "填充色 R 0-255，可选" },
        fg: { type: "integer", description: "填充色 G 0-255，可选" },
        fb: { type: "integer", description: "填充色 B 0-255，可选" },
      },
    },
  },
  {
    name: "ai_add_text",
    description: "在当前 Illustrator 文档中添加文字。text 为文字内容，size 字号（点），x/y 为位置，fr/fg/fb 为颜色。",
    inputSchema: {
      type: "object",
      properties: {
        text: { type: "string", description: "文字内容" },
        size: { type: "number", description: "字号（点），默认 24" },
        x: { type: "number", description: "X 坐标（点）" },
        y: { type: "number", description: "Y 坐标（点）" },
        fr: { type: "integer", description: "颜色 R 0-255，可选" },
        fg: { type: "integer", description: "颜色 G 0-255，可选" },
        fb: { type: "integer", description: "颜色 B 0-255，可选" },
      },
      required: ["text"],
    },
  },
  {
    name: "ai_save_as_ai",
    description: "将当前 Illustrator 文档保存为 .ai 格式。path 可选。",
    inputSchema: {
      type: "object",
      properties: { path: { type: "string", description: "保存路径（含 .ai 后缀），可选" } },
    },
  },
  {
    name: "ai_export_svg",
    description: "将当前 Illustrator 文档导出为 SVG 矢量图。path 可选。",
    inputSchema: {
      type: "object",
      properties: { path: { type: "string", description: "导出路径（含 .svg 后缀），可选" } },
    },
  },
  {
    name: "ai_export_png",
    description: "将当前 Illustrator 文档导出为 PNG 位图（透明背景）。path 可选。",
    inputSchema: {
      type: "object",
      properties: { path: { type: "string", description: "导出路径（含 .png 后缀），可选" } },
    },
  },
];

const server = new Server(
  { name: "adobe-com-mcp", version: "1.0.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;
  let script, env;

  // --- Photoshop tools (ps_ prefix) ---
  if (name.startsWith("ps_")) {
    script = PS_SCRIPT;
    const tool = name.slice(3);
    switch (tool) {
      case "create_document": {
        const w = int(args.width, 800);
        const h = args.height != null ? int(args.height, w) : w;
        env = {
          PSMCP_TOOL: "create_document",
          PSMCP_W: String(w),
          PSMCP_H: String(h),
          PSMCP_NAME: str(args.name),
        };
        break;
      }
      case "get_active_info":
        env = { PSMCP_TOOL: "get_active_document_info" }; break;
      case "list_documents":
        env = { PSMCP_TOOL: "list_open_documents" }; break;
      case "close_document":
        env = { PSMCP_TOOL: "close_active_document" }; break;
      case "open_document":
        env = { PSMCP_TOOL: "open_document", PSMCP_PATH: str(args.path) }; break;
      case "save_as_psd":
        env = { PSMCP_TOOL: "save_as_psd", PSMCP_PATH: str(args.path) }; break;
      case "save_as_png":
        env = { PSMCP_TOOL: "save_as_png", PSMCP_PATH: str(args.path) }; break;
      case "save_as_jpg":
        env = { PSMCP_TOOL: "save_as_jpg", PSMCP_PATH: str(args.path), PSMCP_QUALITY: String(int(args.quality, 8)) }; break;
      case "resize_document":
        env = { PSMCP_TOOL: "resize_document",
          PSMCP_W: String(int(args.width, 0)),
          PSMCP_H: String(int(args.height, 0)) };
        break;
      case "add_layer":
        env = { PSMCP_TOOL: "add_layer", PSMCP_NAME: str(args.name) }; break;
      case "duplicate_layer":
        env = { PSMCP_TOOL: "duplicate_layer" }; break;
      case "delete_layer":
        env = { PSMCP_TOOL: "delete_layer" }; break;
      case "set_layer_opacity":
        env = { PSMCP_TOOL: "set_layer_opacity", PSMCP_OPACITY: String(int(args.opacity, 100)) }; break;
      case "add_text_layer":
        env = { PSMCP_TOOL: "add_text_layer",
          PSMCP_TEXT: str(args.text, "Text"),
          PSMCP_SIZE: String(num(args.size, 48)),
          PSMCP_X: String(num(args.x, 50)),
          PSMCP_Y: String(num(args.y, 100)),
          PSMCP_R: String(int(args.r, 0)),
          PSMCP_G: String(int(args.g, 0)),
          PSMCP_B: String(int(args.b, 0)) };
        break;
      case "set_foreground_color":
        env = { PSMCP_TOOL: "set_foreground_color",
          PSMCP_R: String(int(args.r, 0)),
          PSMCP_G: String(int(args.g, 0)),
          PSMCP_B: String(int(args.b, 0)) };
        break;
      case "fill_layer":
        env = { PSMCP_TOOL: "fill_layer",
          PSMCP_R: String(int(args.r, 0)),
          PSMCP_G: String(int(args.g, 0)),
          PSMCP_B: String(int(args.b, 0)) };
        break;
      case "apply_gaussian_blur":
        env = { PSMCP_TOOL: "apply_gaussian_blur", PSMCP_RADIUS: String(num(args.radius, 5)) }; break;
      case "apply_unsharp_mask":
        env = { PSMCP_TOOL: "apply_unsharp_mask",
          PSMCP_AMOUNT: String(num(args.amount, 100)),
          PSMCP_RADIUS: String(num(args.radius, 2)),
          PSMCP_THRESHOLD: String(num(args.threshold, 0)) };
        break;
      case "do_action":
        env = { PSMCP_TOOL: "do_action",
          PSMCP_ACTION: str(args.action),
          PSMCP_FROM: str(args.from) };
        break;
      default:
        return { content: [{ type: "text", text: JSON.stringify({ result: "error", error: `unknown ps tool: ${tool}` }) }] };
    }
  }

  // --- Illustrator tools (ai_ prefix) ---
  else if (name.startsWith("ai_")) {
    script = AI_SCRIPT;
    const tool = name.slice(3);
    switch (tool) {
      case "create_document":
        env = { AIMCP_TOOL: "ai_create_document",
          AIMCP_W: String(num(args.width, 612)),
          AIMCP_H: String(num(args.height, 792)) };
        break;
      case "get_active_info":
        env = { AIMCP_TOOL: "ai_get_active_info" }; break;
      case "list_documents":
        env = { AIMCP_TOOL: "ai_list_documents" }; break;
      case "close_document":
        env = { AIMCP_TOOL: "ai_close_document" }; break;
      case "add_rectangle":
        env = { AIMCP_TOOL: "ai_add_rectangle",
          AIMCP_X: String(num(args.x, 100)),
          AIMCP_Y: String(num(args.y, 400)),
          AIMCP_W: String(num(args.w, 200)),
          AIMCP_H: String(num(args.h, 150)),
          AIMCP_FR: String(int(args.fr, -1)),
          AIMCP_FG: String(int(args.fg, -1)),
          AIMCP_FB: String(int(args.fb, -1)) };
        break;
      case "add_ellipse":
        env = { AIMCP_TOOL: "ai_add_ellipse",
          AIMCP_X: String(num(args.x, 200)),
          AIMCP_Y: String(num(args.y, 500)),
          AIMCP_W: String(num(args.w, 150)),
          AIMCP_H: String(num(args.h, 100)),
          AIMCP_FR: String(int(args.fr, -1)),
          AIMCP_FG: String(int(args.fg, -1)),
          AIMCP_FB: String(int(args.fb, -1)) };
        break;
      case "add_polygon":
        env = { AIMCP_TOOL: "ai_add_polygon",
          AIMCP_X: String(num(args.x, 300)),
          AIMCP_Y: String(num(args.y, 400)),
          AIMCP_RADIUS: String(num(args.radius, 100)),
          AIMCP_SIDES: String(int(args.sides, 6)),
          AIMCP_FR: String(int(args.fr, -1)),
          AIMCP_FG: String(int(args.fg, -1)),
          AIMCP_FB: String(int(args.fb, -1)) };
        break;
      case "add_text":
        env = { AIMCP_TOOL: "ai_add_text",
          AIMCP_TEXT: str(args.text, "Text"),
          AIMCP_SIZE: String(num(args.size, 24)),
          AIMCP_X: String(num(args.x, 100)),
          AIMCP_Y: String(num(args.y, 300)),
          AIMCP_FR: String(int(args.fr, -1)),
          AIMCP_FG: String(int(args.fg, -1)),
          AIMCP_FB: String(int(args.fb, -1)) };
        break;
      case "save_as_ai":
        env = { AIMCP_TOOL: "ai_save_as_ai", AIMCP_PATH: str(args.path) }; break;
      case "export_svg":
        env = { AIMCP_TOOL: "ai_export_svg", AIMCP_PATH: str(args.path) }; break;
      case "export_png":
        env = { AIMCP_TOOL: "ai_export_png", AIMCP_PATH: str(args.path) }; break;
      default:
        return { content: [{ type: "text", text: JSON.stringify({ result: "error", error: `unknown ai tool: ${tool}` }) }] };
    }
  }

  else {
    return { content: [{ type: "text", text: JSON.stringify({ result: "error", error: `unknown tool: ${name}` }) }] };
  }

  let out;
  try {
    out = runPs(script, env);
  } catch (e) {
    out = JSON.stringify({ result: "error", error: String(e.message || e) });
  }
  return { content: [{ type: "text", text: out || '{"result":"error","error":"empty output"}' }] };
});

const transport = new StdioServerTransport();
await server.connect(transport);
