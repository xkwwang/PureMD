# -*- coding: utf-8 -*-
"""启动 e2e6 测试实例：隔离 APPDATA + CDP 9226，wait 住保持后台任务存活。
用法: python launch_e2e6.py   （前置: seed_e2e6.py 已造好 fixture）"""
import json, os, subprocess, time

EXE = r"D:\00. APP\PureMD\md_preview_app\PureMD\PureMD.exe"
BASE = os.path.join(os.environ["TEMP"], "pm_e2e6")
ISO_APPDATA = os.path.join(BASE, "appdata")
PORT = "9226"

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
    cfgp = os.path.join(ISO_APPDATA, "PureMD", "config.json")
    with open(cfgp, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    path = cfg["session"]["active"]
    if not os.path.isfile(path):
        print("FATAL: fixture doc missing:", path)
        raise SystemExit(1)
    print("fixture doc OK:", path)
    kill_all()
    time.sleep(0.6)
    env = dict(os.environ)
    env["APPDATA"] = ISO_APPDATA
    env["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] = "--remote-debugging-port=" + PORT
    p = subprocess.Popen([EXE, path], env=env,
                         creationflags=0x00000008 | 0x00000200,
                         close_fds=True)
    print("launched PID", p.pid, flush=True)
    rc = p.wait()
    print("exe exited rc=", rc, flush=True)

if __name__ == "__main__":
    main()
