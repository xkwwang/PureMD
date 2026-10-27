# -*- coding: utf-8 -*-
"""反向对照：把 locateGen / markTreeFile 陈旧防护从【冻结版 hub.html】里摘掉，
验证 e2e_sidebar3 确实会失败（证明它真的覆盖了那个 bug）。
用法：python negctl.py on|off
"""
import shutil, sys, os

FROZEN = r"D:\00. APP\PureMD\md_preview_app\PureMD\_internal\mdpreview\hub.html"
BAK = FROZEN + ".bak"
SRC = r"D:\00. APP\PureMD\mdpreview\hub.html"

NEW_GEN = """      if (gen === undefined) gen = ++locateGen;
      else if (gen !== locateGen) return;      // 已有更新的定位请求 → 放弃陈旧重试"""
OLD_GEN = """      if (gen === undefined) gen = ++locateGen;
      else if (gen !== locateGen) { /* negctl: 防护已摘除 */ }"""

NEW_GUARD = """        if (path && typeof locatePath !== "undefined" && locatePath && path !== locatePath) return;"""
OLD_GUARD = """        /* negctl: 陈旧防护已摘除 */"""


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "off"
    if mode == "on":
        if not os.path.isfile(BAK):
            shutil.copy2(FROZEN, BAK)
        with open(FROZEN, "r", encoding="utf-8") as f:
            s = f.read()
        assert NEW_GEN in s and NEW_GUARD in s, "冻结版不含修复，无法做反向对照"
        s = s.replace(NEW_GEN, OLD_GEN).replace(NEW_GUARD, OLD_GUARD)
        with open(FROZEN, "w", encoding="utf-8") as f:
            f.write(s)
        print("NEGCTL ON: 修复已摘除")
    else:
        shutil.copy2(SRC, FROZEN)
        if os.path.isfile(BAK):
            os.remove(BAK)
        print("NEGCTL OFF: 冻结版已还原为源文件（含修复）")


if __name__ == "__main__":
    main()
