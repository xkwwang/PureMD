// -*- mode: js -*-
// 多窗口 e2e：用冻结在 md_preview_app/_new 的新 exe 验证
//   1) 两个实例各自注册进实例注册表
//   2) 跨进程 receive_tab IPC 能把标签注入目标窗口
//   3) --new-window 能再开一个独立窗口
//   4) 双击 .md 会并入「最近使用」的实例（不新建窗口）
// 不依赖真实 HTML5 拖拽（那部分由用户真机验证），只验证底层多实例/IPC 机制。
import { spawn } from "child_process";
import net from "net";
import { existsSync, readdirSync, readFileSync, rmSync, mkdirSync, writeFileSync } from "fs";
import os from "os";
import path from "path";
import { setTimeout as sleep } from "timers/promises";

const EXE = "D:\\00. APP\\PureMD\\md_preview_app\\PureMD\\PureMD.exe";
const INTERNAL = "D:\\00. APP\\PureMD\\md_preview_app\\PureMD\\_internal";
const INST_DIR = path.join(INTERNAL, "mdpreview", ".instances");
const SESS_DIR = path.join(INTERNAL, "mdpreview", ".sessions");
const BASE = path.join(os.tmpdir(), "mw_e2e_" + Date.now());
mkdirSync(BASE, { recursive: true });

const WS = globalThis.WebSocket;
let pass = 0, fail = 0;
function check(cond, msg) { if (cond) { pass++; console.log("PASS", msg); } else { fail++; console.log("FAIL", msg); } }

function killNewExes() {
  try {
    spawn("powershell", ["-NoProfile", "-Command",
      "Get-CimInstance Win32_Process -Filter \"ExecutablePath LIKE '%md_preview_app\\\\PureMD\\\\PureMD.exe'\" | " +
      "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"], { stdio: "ignore" });
  } catch (e) {}
}
async function fetchJSON(url) {
  for (let i = 0; i < 80; i++) {
    try { const r = await fetch(url); if (r.ok) return await r.json(); } catch (e) {}
    await sleep(250);
  }
  return null;
}
async function connect(port) {
  const list = await fetchJSON(`http://127.0.0.1:${port}/json/list`);
  if (!list) throw new Error("no CDP at " + port);
  const page = list.find(p => /PureMD/.test(p.title)) || list[0];
  const ws = new WS(page.webSocketDebuggerUrl);
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
  let id = 0; const pending = {};
  ws.onmessage = (e) => { const m = JSON.parse(e.data); if (m.id && pending[m.id]) { pending[m.id](m.result); delete pending[m.id]; } };
  const send = (method, params) => new Promise((res) => { const i = ++id; pending[i] = res; ws.send(JSON.stringify({ id: i, method, params })); });
  return {
    eval: async (expr) => {
      const r = await send("Runtime.evaluate", { expression: expr, returnByValue: true, awaitPromise: true });
      if (r.exceptionDetails) throw new Error(JSON.stringify(r.exceptionDetails));
      return r.result.value;
    },
    close: () => ws.close(),
  };
}
function tabCount(cdp) { return cdp.eval("document.querySelectorAll('#tabs .tab').length"); }

async function launchExe(args, cdpPort, appdata) {
  const env = { ...process.env, APPDATA: appdata, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${cdpPort}` };
  const p = spawn(EXE, args, { env, detached: true, stdio: "ignore" });
  await fetchJSON(`http://127.0.0.1:${cdpPort}/json/list`);
  return p;
}
function readInstances() {
  if (!existsSync(INST_DIR)) return [];
  return readdirSync(INST_DIR).filter(f => f.endsWith(".json")).map(f => {
    try { return JSON.parse(readFileSync(path.join(INST_DIR, f), "utf-8")); } catch (e) { return null; }
  }).filter(Boolean);
}
function tcpSend(port, obj) {
  return new Promise((resolve) => {
    const s = net.connect(port, "127.0.0.1");
    let buf = "";
    s.setTimeout(3000);
    s.on("connect", () => { s.write(JSON.stringify(obj)); s.end(); });
    s.on("data", (d) => { buf += d.toString(); });
    s.on("timeout", () => { s.destroy(); resolve(buf); });
    s.on("close", () => resolve(buf));
    s.on("error", () => resolve(buf));
  });
}

async function main() {
  // 清理上一次残留
  killNewExes();
  await sleep(800);
  try { rmSync(INST_DIR, { recursive: true, force: true }); } catch (e) {}
  try { rmSync(SESS_DIR, { recursive: true, force: true }); } catch (e) {}

  const aMd = path.join(BASE, "a.md"); writeFileSync(aMd, "# A\nhello\n");
  const cMd = path.join(BASE, "c.md"); writeFileSync(cMd, "# C\nworld\n");
  const isoA = path.join(BASE, "appA"), isoB = path.join(BASE, "appB");
  mkdirSync(isoA, { recursive: true }); mkdirSync(isoB, { recursive: true });

  console.log("== launch A (a.md) ==");
  const pA = await launchExe([aMd], 9241, isoA);
  await sleep(1500);
  const cdpA = await connect(9241);
  const nA0 = await tabCount(cdpA);
  check(nA0 === 1, "A 启动后应有 1 个标签 (got " + nA0 + ")");
  let inst = readInstances();
  check(inst.length === 1, "注册表应有 1 个实例 (got " + inst.length + ")");
  const portA = inst[0].port;

  console.log("== launch B (--new-window) ==");
  const pB = await launchExe(["--new-window"], 9242, isoB);
  await sleep(1500);
  const cdpB = await connect(9242);
  const nB0 = await tabCount(cdpB);
  check(nB0 === 0, "B(--new-window) 应为空窗口 (got " + nB0 + ")");
  inst = readInstances();
  check(inst.length === 2, "注册表应有 2 个实例 (got " + inst.length + ")");

  console.log("== receive_tab IPC → A ==");
  const resp = await tcpSend(portA, { __ctl: "receive_tab", tab: { name: "IPC", path: "", html: "<p>ipc</p>", scrollY: 0 } });
  await sleep(800);
  const nA1 = await tabCount(cdpA);
  check(nA1 === 2, "A 经 IPC 收到标签后应有 2 个 (got " + nA1 + ", resp=" + resp.trim() + ")");

  console.log("== double-click delegation: 启动第三个进程打开 c.md（应并入最近使用的 A） ==");
  // 放在 new_window() 之前：此时只有 A、B 两个实例，A 因 receive_tab 刚被激活 → 最活跃
  const instBefore = readInstances().length;
  const pC = await launchExe([cMd], 9243, isoA);
  await sleep(2500);
  const nA2 = await tabCount(cdpA);
  const instAfter = readInstances().length;
  check(nA2 === 3, "A 应并入 c.md → 共 3 个标签 (got " + nA2 + ")");
  check(instAfter === instBefore, "第三个进程应并入而非新建实例 (" + instBefore + "→" + instAfter + ")");

  console.log("== new_window() API on B ==");
  const before = readInstances().length;
  await cdpB.eval("window.pywebview.api.new_window()");
  await sleep(1500);
  const after = readInstances().length;
  check(after === before + 1, "new_window() 应再注册 1 个实例 (" + before + "→" + after + ")");

  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  cdpA.close(); cdpB.close();
  killNewExes();
  process.exit(fail ? 1 : 0);
}
main().catch(e => { console.error("E2E ERROR", e); killNewExes(); process.exit(2); });
