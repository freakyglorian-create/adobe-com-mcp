// Full end-to-end test for adobe-com-mcp (PS + AI).
import { spawn } from "node:child_process";
import fs from "node:fs";

const p = spawn("node", ["server.js"], { stdio: ["pipe", "pipe", "pipe"] });
let buf = "";
let nextId = 1;
const pending = {};
let pass = 0, fail = 0;

p.stdout.on("data", (d) => {
  buf += d.toString();
  let i;
  while ((i = buf.indexOf("\n")) >= 0) {
    const line = buf.slice(0, i);
    buf = buf.slice(i + 1);
    if (line.trim()) handle(line);
  }
});
p.stderr.on("data", (d) => process.stderr.write("[stderr] " + d.toString().slice(0, 200) + "\n"));
p.on("close", (c) => console.log(`[server exited code=${c}] pass=${pass} fail=${fail}`));

function send(method, params) {
  const id = nextId++;
  p.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n");
  return id;
}

function handle(line) {
  let m;
  try { m = JSON.parse(line); } catch { console.log("<< " + line.slice(0, 200)); return; }
  if (m.id && pending[m.id]) {
    const cb = pending[m.id];
    delete pending[m.id];
    cb(m);
  }
}

function waitFor(id) {
  return new Promise((res) => { pending[id] = res; });
}

function assert(label, cond, detail = "") {
  if (cond) { pass++; console.log(`  PASS  ${label}`); }
  else { fail++; console.log(`  FAIL  ${label}  ${detail}`); }
}

function textOf(result) {
  const t = result?.content?.[0]?.text;
  try { return JSON.parse(t); } catch { return t; }
}

async function runTest(name, fn) {
  console.log(`\n=== ${name} ===`);
  try { await fn(); }
  catch (e) { fail++; console.log(`  EXCEPTION: ${e.message}`); }
}

async function main() {
  const id1 = send("initialize", { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "test", version: "1.0" } });
  await waitFor(id1);
  send("notifications/initialized", {});

  const id2 = send("tools/list", {});
  const tl = await waitFor(id2);
  const toolNames = tl.result.tools.map((t) => t.name);
  console.log(`\nTotal tools: ${toolNames.length}`);
  console.log(`  PS tools: ${toolNames.filter(n => n.startsWith("ps_")).length}`);
  console.log(`  AI tools: ${toolNames.filter(n => n.startsWith("ai_")).length}`);

  // ---- PHOTOSHOP ----
  await runTest("PS: create 800x800 document", async () => {
    const id = send("tools/call", { name: "ps_create_document", arguments: { width: 800, height: 800, name: "MCP-PS-800" } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok", JSON.stringify(d));
    assert("width=800", d.width === 800);
    assert("height=800", d.height === 800);
  });

  await runTest("PS: get active info", async () => {
    const id = send("tools/call", { name: "ps_get_active_info", arguments: {} });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok");
    assert("has pixelWidth", typeof d.pixelWidth === "number");
    assert("has pixelHeight", typeof d.pixelHeight === "number");
    assert("has layers count", typeof d.layers === "number");
  });

  await runTest("PS: add text layer", async () => {
    const id = send("tools/call", { name: "ps_add_text_layer", arguments: { text: "Hello TRAE", size: 48, x: 100, y: 200, r: 255, g: 0, b: 0 } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok", JSON.stringify(d));
    assert("text matches", d.text === "Hello TRAE");
  });

  await runTest("PS: fill + duplicate + delete layer", async () => {
    let id = send("tools/call", { name: "ps_fill_layer", arguments: { r: 100, g: 200, b: 255 } });
    let r = await waitFor(id);
    let d = textOf(r.result);
    assert("fill_layer ok", d.result === "ok", JSON.stringify(d));
    id = send("tools/call", { name: "ps_duplicate_layer", arguments: {} });
    r = await waitFor(id);
    d = textOf(r.result);
    assert("duplicate_layer ok", d.result === "ok");
    id = send("tools/call", { name: "ps_delete_layer", arguments: {} });
    r = await waitFor(id);
    d = textOf(r.result);
    assert("delete_layer ok", d.result === "ok");
  });

  await runTest("PS: gaussian blur + save png", async () => {
    let id = send("tools/call", { name: "ps_apply_gaussian_blur", arguments: { radius: 3 } });
    let r = await waitFor(id);
    let d = textOf(r.result);
    assert("gaussian_blur ok", d.result === "ok", JSON.stringify(d));
    const outPath = process.env.USERPROFILE + "\\Documents\\MCP-PS-800.png";
    id = send("tools/call", { name: "ps_save_as_png", arguments: { path: outPath } });
    r = await waitFor(id);
    d = textOf(r.result);
    assert("save_as_png ok", d.result === "ok");
    assert("file exists", fs.existsSync(outPath), outPath);
  });

  await runTest("PS: resize document", async () => {
    const id = send("tools/call", { name: "ps_resize_document", arguments: { width: 400 } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("resize ok", d.result === "ok");
    assert("width=400", d.pixelWidth === 400);
    assert("height=400 (square preserved)", d.pixelHeight === 400);
  });

  // ---- ILLUSTRATOR ----
  await runTest("AI: create document", async () => {
    const id = send("tools/call", { name: "ai_create_document", arguments: { width: 800, height: 600 } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok", JSON.stringify(d));
    assert("width=800", d.width === 800);
    assert("height=600", d.height === 600);
  });

  await runTest("AI: add rectangle + fill", async () => {
    const id = send("tools/call", { name: "ai_add_rectangle", arguments: { x: 100, y: 400, w: 200, h: 150, fr: 255, fg: 100, fb: 50 } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok", JSON.stringify(d));
    assert("has name", !!d.name);
  });

  await runTest("AI: add ellipse", async () => {
    const id = send("tools/call", { name: "ai_add_ellipse", arguments: { x: 400, y: 500, w: 120, h: 120, fr: 50, fg: 150, fb: 255 } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok");
  });

  await runTest("AI: add polygon", async () => {
    const id = send("tools/call", { name: "ai_add_polygon", arguments: { x: 600, y: 450, radius: 80, sides: 5, fr: 255, fg: 255, fb: 0 } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok");
    assert("sides=5", d.sides === 5);
  });

  await runTest("AI: add text", async () => {
    const id = send("tools/call", { name: "ai_add_text", arguments: { text: "TRAE AI Test", size: 36, x: 150, y: 250, fr: 0, fg: 0, fb: 0 } });
    const r = await waitFor(id);
    const d = textOf(r.result);
    assert("result=ok", d.result === "ok", JSON.stringify(d));
    assert("text matches", d.text === "TRAE AI Test");
  });

  await runTest("AI: get info + save + export svg", async () => {
    let id = send("tools/call", { name: "ai_get_active_info", arguments: {} });
    let r = await waitFor(id);
    let d = textOf(r.result);
    assert("get_active_info ok", d.result === "ok");
    assert("has paths count", typeof d.paths === "number");
    assert("has textFrames count", typeof d.textFrames === "number");

    const aiPath = process.env.USERPROFILE + "\\Documents\\MCP-AI-Test.ai";
    id = send("tools/call", { name: "ai_save_as_ai", arguments: { path: aiPath } });
    r = await waitFor(id);
    d = textOf(r.result);
    assert("save_as_ai ok", d.result === "ok");

    const svgPath = process.env.USERPROFILE + "\\Documents\\MCP-AI-Test.svg";
    id = send("tools/call", { name: "ai_export_svg", arguments: { path: svgPath } });
    r = await waitFor(id);
    d = textOf(r.result);
    assert("export_svg ok", d.result === "ok");
    assert("svg file exists", fs.existsSync(svgPath), svgPath);
  });

  console.log(`\n========= TOTAL: pass=${pass} fail=${fail} =========`);
  setTimeout(() => process.exit(fail > 0 ? 1 : 0), 500);
}

main().catch((e) => { console.error("FATAL:", e); process.exit(1); });
