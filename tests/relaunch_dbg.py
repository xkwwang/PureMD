# -*- coding: utf-8 -*-
"""复现真实配置下的侧栏定位问题：把真实 config.json 复制到隔离 APPDATA，
带 CDP 9222 端口启动 exe，wait 住保持后台任务存活。"""
import json, os, shutil, subprocess, sys, time

EXE = r"D:\00. APP\PureMD\md_preview_app\PureMD\PureMD.exe"
DBG = os.path.join(os.environ["TEMP"], "pm_dbg")
ISO_APPDATA = os.path.join(DBG, "appdata")

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
    real_cfg = os.path.join(os.environ["APPDATA"], "PureMD", "config.json")
    with open(real_cfg, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    path = cfg["session"]["active"]
    if not os.path.isfile(path):
        print("FATAL: active file missing:", path); sys.exit(1)
    # 隔离 APPDATA：复制真实配置，避免调试实例污染用户数据
    iso_dir = os.path.join(ISO_APPDATA, "PureMD")
    os.makedirs(iso_dir, exist_ok=True)
    shutil.copy2(real_cfg, os.path.join(iso_dir, "config.json"))
    print("iso config copied; active:", path)
    kill_all()
    time.sleep(0.6)
    env = dict(os.environ)
    env["APPDATA"] = ISO_APPDATA
    env["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] = "--remote-debugging-port=9222"
    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    p = subprocess.Popen([EXE, path], env=env,
                         creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
                         close_fds=True)
    print("launched PID", p.pid, flush=True)
    rc = p.wait()   # 同生共死：exe 退出任务才结束
    print("exe exited rc=", rc, flush=True)

if __name__ == "__main__":
    main()
