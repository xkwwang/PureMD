# -*- coding: utf-8 -*-
"""从 config.json 读取 active 路径，零手打重启 PureMD 用户实例。
先杀掉现有 PureMD 进程（按 PID 列表），再以 DETACHED_PROCESS 后台启动。"""
import json, os, subprocess, sys, time

CFG = os.path.join(os.environ["APPDATA"], "PureMD", "config.json")
EXE = r"D:\00. APP\PureMD\md_preview_app\PureMD\PureMD.exe"

def kill_puremd():
    ps = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "Get-CimInstance Win32_Process -Filter \"Name='PureMD.exe'\" | "
         "Select-Object -ExpandProperty ProcessId"],
        capture_output=True, text=True)
    pids = [int(x) for x in ps.stdout.split() if x.strip().isdigit()]
    for pid in pids:
        subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"],
                       capture_output=True)
        print("killed PID", pid)
    if not pids:
        print("no PureMD running")
    return pids

def main():
    with open(CFG, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    path = cfg["session"]["active"]
    if not os.path.isfile(path):
        print("FATAL: active file missing:", path)
        sys.exit(1)
    print("active path OK, bytes:", len(open(path, "rb").read()))
    kill_puremd()
    time.sleep(0.6)
    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    p = subprocess.Popen(
        [EXE, path],
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
        close_fds=True)
    print("launched PID", p.pid, flush=True)
    # 关键：wait 住 exe。若 python 先退出，后台任务的 job 会把 exe（孙子进程）
    # 一并收走，表现为"启动成功但秒没"。同生共死直到用户关窗。
    rc = p.wait()
    print("exe exited rc=", rc, flush=True)

if __name__ == "__main__":
    main()
