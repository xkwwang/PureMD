// -*- mode: js -*-
// 查找/替换功能 e2e（Ctrl+F / Ctrl+H）：用冻结在 md_preview_app/_build_test 的临时 exe 验证
//   1) Ctrl+F 打开查找栏；输入即高亮正文里的匹配（区分大小写可切换）
//   2) 匹配数 / 当前序号显示正确，Enter / Shift+Enter 前后跳转
//   3) 关掉查找栏会清除高亮
//   4) Ctrl+H 打开替换：单个替换 / 全部替换；保存后源文件正确写回且无高亮污染
// 单窗口即可，不触碰用户正在用的真实实例（独立 APPDATA + 独立 build 目录）。
import { spawn, execFileSync } from "child_process";
import { existsSync, readdirSync, readFileSync, rmSync, mkdirSync, writeFileSync } from "fs";
import os from "os";
import path from "path";
import { setTimeout as sleep } from "timers/promises";

const WS = globalThis.WebSocket;
let pass = 0, fail = 0;
function check(cond, msg) { if (cond) { pass++; console.log("PASS", msg); } else { fail++; console.log("FAIL", msg); } }

const EXE = process.env.PUREMD_EXE ||
  "D:\\00. APP\\PureMD\\md_preview_app\\_build_test\\PureMD\\PureMD.exe";
const INTERNAL = path.join(path.dirname(EXE), "_internal");
const INST_DIR = path.join(INTERNAL, "mdpreview", ".instances");
const SESS_DIR = path.join(INTERNAL, "mdpreview", ".sessions");

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
function killMine() {
  try {
    // 同步、按精确可执行路径杀掉本测试启动的 PureMD（含子进程树），
    // 避免上一轮残留 exe 占用 CDP 端口导致下一轮连错实例。
    // 仅匹配 EXE 自身路径，不会误杀用户正在用的同名实例。
    const exep = EXE.replace(/\//g, "\\");
    execFileSync("powershell", ["-NoProfile", "-Command",
      `$p='${exep.replace(/'/g, "''")}';` +
      `Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $p } | ` +
      `ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }`],
      { stdio: "ignore", timeout: 15000 });
  } catch (e) {}
}
async function launchExe(args, cdpPort, appdata) {
  const env = { ...process.env, APPDATA: appdata, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${cdpPort}` };
  const p = spawn(EXE, args, { env, detached: true, stdio: "ignore" });
  await fetchJSON(`http://127.0.0.1:${cdpPort}/json/list`);
  return p;
}

async function main() {
  killMine();
  await sleep(800);
  try { rmSync(INST_DIR, { recursive: true, force: true }); } catch (e) {}
  try { rmSync(SESS_DIR, { recursive: true, force: true }); } catch (e) {}
  const BASE = path.join(os.tmpdir(), "find_e2e_" + Date.now());
  mkdirSync(BASE, { recursive: true });
  const md = path.join(BASE, "doc.md");
  writeFileSync(md,
    "# Heading\n\n" +
    "apple apple apple\n\n" +
    "banana Apple apple\n");
  const iso = path.join(BASE, "app"); mkdirSync(iso, { recursive: true });

  console.log("== launch doc.md ==");
  const p = await launchExe([md], 9251, iso);
  await sleep(1500);
  const cdp = await connect(9251);

  // 等 iframe 正文渲染出来（且整篇内容就位：含末尾的 banana 段落）
  let ready = false;
  for (let i = 0; i < 40; i++) {
    ready = await cdp.eval("(()=>{var fr=document.querySelector('#views iframe');if(!fr||!fr.contentDocument)return false;var mb=fr.contentDocument.querySelector('.markdown-body');return !!mb && /banana/.test(mb.textContent);})()");
    if (ready) break;
    await sleep(250);
  }
  check(ready, "iframe 正文（.markdown-body）已完整渲染");
  await sleep(400); // 额外稳定时间，避免边渲染边查找

  // 查找栏初始隐藏
  const hidden0 = await cdp.eval("(()=>{var b=document.getElementById('findbar');return !b||b.style.display==='none';})()");
  check(hidden0, "查找栏初始应隐藏");

  console.log("== Ctrl+F 打开查找栏（模拟在 iframe 内按 Ctrl+F）==");
  await cdp.eval(`(()=>{
    var fr=document.querySelector('#views iframe');
    var d=fr.contentDocument;
    d.dispatchEvent(new KeyboardEvent('keydown',{key:'f',ctrlKey:true,bubbles:true}));
  })()`);
  await sleep(400);
  const shown = await cdp.eval("document.getElementById('findbar').style.display==='flex'");
  check(shown, "Ctrl+F 后查找栏应显示 (display=flex)");

  // 输入 apple（默认不区分大小写）→ 期望 5 个：apple×3 + apple×1 + Apple×1
  await cdp.eval(`(()=>{var i=document.getElementById('findInput');i.value='apple';i.dispatchEvent(new Event('input',{bubbles:true}));})()`);
  await sleep(300);
  const nApple = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(nApple === 5, "apple(不区分大小写) 应命中 5 处 (got " + nApple + ")");
  const cnt1 = await cdp.eval("document.getElementById('findCount').textContent");
  check(cnt1 === "1 / 5", "计数器应为 1 / 5 (got " + JSON.stringify(cnt1) + ")");

  // 下一个 → 2 / 5，且当前高亮正确
  await cdp.eval("findNext()");
  await sleep(200);
  const cnt2 = await cdp.eval("document.getElementById('findCount').textContent");
  check(cnt2 === "2 / 5", "findNext 后应为 2 / 5 (got " + JSON.stringify(cnt2) + ")");
  const curOk = await cdp.eval("(()=>{var ms=document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find');for(var i=0;i<ms.length;i++){if(ms[i].classList.contains('pm-find-current'))return i===1;}return false;})()");
  check(curOk, "第 2 个匹配应带 pm-find-current");

  // 上一个 → 回到 1 / 5
  await cdp.eval("findPrev()");
  await sleep(200);
  const cnt3 = await cdp.eval("document.getElementById('findCount').textContent");
  check(cnt3 === "1 / 5", "findPrev 后应为 1 / 5 (got " + JSON.stringify(cnt3) + ")");

  // 区分大小写：Apple（大写）→ 1 处
  await cdp.eval(`(()=>{var c=document.getElementById('findCs');if(!c.checked){c.checked=true;c.dispatchEvent(new Event('change',{bubbles:true}));}var i=document.getElementById('findInput');i.value='Apple';i.dispatchEvent(new Event('input',{bubbles:true}));})()`);
  await sleep(300);
  const nAppleCs = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(nAppleCs === 1, "Apple(区分大小写) 应命中 1 处 (got " + nAppleCs + ")");

  // 无匹配
  await cdp.eval(`(()=>{var i=document.getElementById('findInput');i.value='zzzzz';i.dispatchEvent(new Event('input',{bubbles:true}));})()`);
  await sleep(300);
  const noneTxt = await cdp.eval("document.getElementById('findCount').textContent");
  const noneCls = await cdp.eval("document.getElementById('findCount').classList.contains('none')");
  check(noneTxt === "无匹配" && noneCls, "无匹配应显示「无匹配」并标红 (got " + JSON.stringify(noneTxt) + ")");
  const nMarks = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(nMarks === 0, "无匹配时不应有高亮 (got " + nMarks + ")");

  // 清空 → 高亮清除
  await cdp.eval(`(()=>{var i=document.getElementById('findInput');i.value='';i.dispatchEvent(new Event('input',{bubbles:true}));})()`);
  await sleep(200);
  const cleared = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(cleared === 0, "清空输入框后高亮应清除 (got " + cleared + ")");

  // Esc 关闭查找栏并清高亮
  await cdp.eval(`(()=>{var i=document.getElementById('findInput');i.value='apple';i.dispatchEvent(new Event('input',{bubbles:true}));i.focus();})()`);
  await sleep(200);
  await cdp.eval(`(()=>{var i=document.getElementById('findInput');i.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));})()`);
  await sleep(300);
  const hidden1 = await cdp.eval("document.getElementById('findbar').style.display==='none'");
  const afterEsc = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(hidden1, "Esc 后查找栏应隐藏");
  check(afterEsc === 0, "Esc 关闭后高亮应清除 (got " + afterEsc + ")");

  // ===== 替换功能（Ctrl+H）=====
  console.log("== Ctrl+H 打开替换 ==");
  await cdp.eval(`(()=>{var fr=document.querySelector('#views iframe');fr.contentDocument.dispatchEvent(new KeyboardEvent('keydown',{key:'h',ctrlKey:true,bubbles:true}));})()`);
  await sleep(400);
  // 前面「区分大小写」步骤勾选后一直保留着，这里取消，恢复不区分大小写
  await cdp.eval(`(()=>{var c=document.getElementById('findCs');if(c.checked){c.checked=false;c.dispatchEvent(new Event('change',{bubbles:true}));}})()`);
  await sleep(200);
  const rrowOn = await cdp.eval("document.getElementById('findRrow').classList.contains('on')");
  check(rrowOn, "Ctrl+H 后替换行应展开");
  const barShown2 = await cdp.eval("document.getElementById('findbar').style.display==='flex'");
  check(barShown2, "Ctrl+H 后查找栏应显示");

  // 查找 apple → 5 处；替换为 pear
  await cdp.eval(`(()=>{var i=document.getElementById('findInput');i.value='apple';i.dispatchEvent(new Event('input',{bubbles:true}));})()`);
  await sleep(300);
  await cdp.eval(`(()=>{document.getElementById('findReplace').value='pear';})()`);
  const nBefore = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(nBefore === 5, "替换前应命中 5 处 (got " + nBefore + ")");

  // 单个替换 → 剩 4 处，第一段变 pear apple apple，且标记未保存
  await cdp.eval("document.getElementById('findReplOne').click()");
  await sleep(300);
  const nAfterOne = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(nAfterOne === 4, "单个替换后应剩 4 处 (got " + nAfterOne + ")");
  const p1 = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('.markdown-body p')[0].textContent");
  check(p1 === "pear apple apple", "第一段应为 pear apple apple (got " + JSON.stringify(p1) + ")");
  const dirty1 = await cdp.eval("EDITED[activeTab().id] === true");
  check(dirty1, "替换后应标记为未保存");

  // 全部替换 → 高亮清零，正文无 apple（含原大写 Apple → Pear）
  await cdp.eval("document.getElementById('findReplAll').click()");
  await sleep(300);
  const nAfterAll = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(nAfterAll === 0, "全部替换后应无高亮 (got " + nAfterAll + ")");
  const bodyTxt = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelector('.markdown-body').textContent");
  check(!/apple/i.test(bodyTxt), "全部替换后正文不应再有 apple");
  check(/banana pear pear/.test(bodyTxt), "正文应为 banana pear pear（替换文本按字面插入，同 Excel）");

  // 点「保存」→ 源文件写回 pear；且无 == / pm-find 高亮污染
  await cdp.eval("document.getElementById('saveBtn').click()");
  let savedTxt = "";
  for (let i = 0; i < 30; i++) {
    await sleep(300);
    try { savedTxt = readFileSync(md, "utf8"); } catch (e) {}
    if (/pear/i.test(savedTxt)) break;
  }
  check(/pear/i.test(savedTxt), "保存后源文件应含 pear");
  check(!/apple/i.test(savedTxt), "保存后源文件不应再有 apple");
  check(!savedTxt.includes("==") && !savedTxt.includes("pm-find"), "源文件不应有高亮标记污染");
  const expected = "# Heading\n\npear pear pear\n\nbanana pear pear\n";
  check(savedTxt.replace(/\r\n/g, "\n") === expected, "源文件内容应与预期一致 (got " + JSON.stringify(savedTxt) + ")");

  // ===== 工具栏最右侧搜索框 =====
  console.log("== 工具栏搜索框 ==");
  const tfVisible = await cdp.eval("!!document.getElementById('topFind')");
  check(tfVisible, "工具栏右侧应有搜索框");
  // 此刻正文已是被保存后的内容：pear×5（第一段 3 + 第二段 2）
  await cdp.eval(`(()=>{var t=document.getElementById('topFind');t.value='pear';t.dispatchEvent(new Event('input',{bubbles:true}));})()`);
  await sleep(300);
  const tfBarShown = await cdp.eval("document.getElementById('findbar').style.display==='flex'");
  check(tfBarShown, "搜索框输入后查找栏应显示");
  const tfMarks = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(tfMarks === 5, "搜索框输入 pear 应命中 5 处 (got " + tfMarks + ")");
  const tfSync = await cdp.eval("document.getElementById('findInput').value");
  check(tfSync === "pear", "查找栏应同步搜索框的词 (got " + JSON.stringify(tfSync) + ")");
  // Enter → 下一个
  await cdp.eval(`(()=>{var t=document.getElementById('topFind');t.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));})()`);
  await sleep(200);
  const tfCnt = await cdp.eval("document.getElementById('findCount').textContent");
  check(tfCnt === "2 / 5", "搜索框 Enter 后应 2 / 5 (got " + JSON.stringify(tfCnt) + ")");
  // Esc 关闭
  await cdp.eval(`(()=>{var t=document.getElementById('topFind');t.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));})()`);
  await sleep(200);
  const tfHidden = await cdp.eval("document.getElementById('findbar').style.display==='none'");
  check(tfHidden, "搜索框 Esc 后查找栏应隐藏");
  const tfMarks0 = await cdp.eval("document.querySelector('#views iframe').contentDocument.querySelectorAll('mark.pm-find').length");
  check(tfMarks0 === 0, "搜索框 Esc 后高亮应清除 (got " + tfMarks0 + ")");

  console.log(`\nRESULT: ${pass} passed, ${fail} failed`);
  cdp.close();
  killMine();
  process.exit(fail ? 1 : 0);
}
main().catch(e => { console.error("E2E ERROR", e); killMine(); process.exit(2); });
