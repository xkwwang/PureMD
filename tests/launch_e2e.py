# -*- coding: utf-8 -*-
"""通用 e2e 启动器：python launch_e2e.py <fixture目录名> <CDP端口>
例：launch_e2e.py pm_e2e6 9226
读该 fixture 隔离配置里的 active 路径（没有则退回 <base>/t.md），
带 CDP 端口启动冻结 exe，并 wait 住保持后台任务存活。"""
import json, os, subprocess, sys, time

EXE = r"D:\00. APP\PureMD\md_preview_app\PureMD\PureMD.exe"

def kill_all():
    ps = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "Get-CimInstance Win32_Process -Filter \"Name='PureMD.exe'\" | "
         "Select-Object -ExpandProperty ProcessId"],
        capture_output=True, text=True)
    pids = [int(x) for x in ps.stdout.split() if x.strip().isdigit()]
    for pid in pids:
        subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"], capture_output=True)
        print("killed PID", pid)
    return pids

def main():
    name = sys.argv[1]
    port = sys.argv[2] if len(sys.argv) > 2 else "9224"
    base = os.path.join(os.environ["TEMP"], name)
    iso = os.path.join(base, "appdata")
    path = None
    cfgp = os.path.join(iso, "PureMD", "config.json")
    if os.path.isfile(cfgp):
        with open(cfgp, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        path = (cfg.get("session") or {}).get("active")
    if not path:
        cand = os.path.join(base, "t.md")
        path = cand if os.path.isfile(cand) else None
    if not path or not os.path.isfile(path):
        print("FATAL: no doc for fixture", name, "->", path)
        raise SystemExit(1)
    print("fixture", name, "doc OK:", path)
    kill_all()
    time.sleep(0.6)
    env = dict(os.environ)
    env["APPDATA"] = iso
    env["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] = "--remote-debugging-port=" + port
    p = subprocess.Popen([EXE, path], env=env,
                         creationflags=0x00000008 | 0x00000200,
                         close_fds=True)
    print("launched PID", p.pid, flush=True)
    rc = p.wait()
    print("exe exited rc=", rc, flush=True)

if __name__ == "__main__":
    main()
