#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
unlock_file.py — 检测文件占用，并优雅解除（Excel / Word / WPS / 通用进程兜底）

用法:
    python unlock_file.py <文件路径>                  # 检测 + 优雅关闭（默认）
    python unlock_file.py <文件路径> --no-close       # 只检测，不关闭
    python unlock_file.py <文件路径> --save           # 有未保存修改时：保存后关闭（COM 路线）
    python unlock_file.py <文件路径> --force-kill     # 强杀占用进程（需用户明确同意）
    python unlock_file.py <文件路径> --json           # 机器可读输出

退出码:
    0  未占用 / 已成功解除
    2  文档有未保存修改，需要用户决定（未做任何关闭动作）
    3  文件被占用，但无法安全自动解除
    1  脚本/环境错误
"""

import argparse
import ctypes
import json
import os
import subprocess
import sys
import time
from ctypes import wintypes

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROGIDS = [
    ("Excel.Application", "workbook"),
    ("KET.Application", "workbook"),      # WPS 表格
    ("Word.Application", "document"),
    ("KWps.Application", "document"),     # WPS 文字
    ("KWPP.Application", "presentation"), # WPS 演示
]


# ---------------------------------------------------------------- 基础

def norm(path):
    return os.path.normcase(os.path.abspath(path))


def is_locked(path):
    """能否以写方式打开；被占用时为 True。"""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "a+b"):
            return False
    except PermissionError:
        return True
    except OSError as e:
        return getattr(e, "winerror", None) in (32, 33)  # sharing / lock violation


def lock_files(path):
    """返回同目录下的 Office 临时锁文件（~$xxx）。"""
    d, name = os.path.split(path)
    return [p for p in (os.path.join(d, "~$" + name),)
            if os.path.exists(p)]


# ---------------------------------------------------------------- Restart Manager

class RM_UNIQUE_PROCESS(ctypes.Structure):
    _fields_ = [("dwProcessId", wintypes.DWORD),
                ("ProcessStartTime", wintypes.FILETIME)]


class RM_PROCESS_INFO(ctypes.Structure):
    _fields_ = [("Process", RM_UNIQUE_PROCESS),
                ("strAppName", wintypes.WCHAR * 256),
                ("strServiceShortName", wintypes.WCHAR * 64),
                ("ApplicationType", ctypes.c_int),
                ("AppStatus", wintypes.ULONG),
                ("TSSessionId", wintypes.DWORD),
                ("bRestartable", wintypes.BOOL)]


def process_name(pid):
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return None
    try:
        buf = ctypes.create_unicode_buffer(1024)
        size = wintypes.DWORD(1024)
        if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value)
    finally:
        k32.CloseHandle(h)
    return None


def restart_manager_owners(path):
    """用 Restart Manager 查谁占着文件。"""
    try:
        rm = ctypes.windll.rstrtmgr
    except Exception:
        return []

    rm.RmStartSession.argtypes = [ctypes.POINTER(wintypes.DWORD),
                                  wintypes.DWORD, ctypes.POINTER(ctypes.c_wchar)]
    rm.RmRegisterResources.argtypes = [wintypes.DWORD, wintypes.UINT,
                                       ctypes.POINTER(ctypes.c_wchar_p),
                                       wintypes.UINT, ctypes.c_void_p,
                                       wintypes.UINT, ctypes.c_void_p]
    rm.RmGetList.argtypes = [wintypes.DWORD, ctypes.POINTER(wintypes.UINT),
                             ctypes.POINTER(wintypes.UINT),
                             ctypes.POINTER(RM_PROCESS_INFO),
                             ctypes.POINTER(wintypes.DWORD)]
    rm.RmEndSession.argtypes = [wintypes.DWORD]

    session = wintypes.DWORD()
    key = ctypes.create_unicode_buffer(256)
    if rm.RmStartSession(ctypes.byref(session), 0, key) != 0:
        return []
    try:
        files = (ctypes.c_wchar_p * 1)(os.path.abspath(path))
        if rm.RmRegisterResources(session.value, 1, files, 0, None, 0, None) != 0:
            return []
        need, count = wintypes.UINT(0), wintypes.UINT(0)
        reasons = wintypes.DWORD(0)
        rc = rm.RmGetList(session.value, ctypes.byref(need), ctypes.byref(count),
                          None, ctypes.byref(reasons))
        if rc == 234 and need.value:  # ERROR_MORE_DATA
            arr = (RM_PROCESS_INFO * need.value)()
            count = wintypes.UINT(need.value)
            rc = rm.RmGetList(session.value, ctypes.byref(need), ctypes.byref(count),
                              arr, ctypes.byref(reasons))
            if rc != 0:
                return []
            out = []
            for i in range(count.value):
                info = arr[i]
                out.append({
                    "pid": info.Process.dwProcessId,
                    "process": process_name(info.Process.dwProcessId),
                    "app": info.strAppName,
                })
            return out
        return []
    finally:
        rm.RmEndSession(session.value)


# ---------------------------------------------------------------- Office / WPS COM

def _same(path_a, path_b):
    try:
        return norm(path_a) == norm(path_b)
    except Exception:
        return False


def find_open_docs(path):
    """在正在运行的 Office/WPS 实例里找这个文件。"""
    try:
        import win32com.client as wc
    except Exception:
        return []

    found = []
    for progid, kind in PROGIDS:
        try:
            app = wc.GetActiveObject(progid)
        except Exception:
            continue
        try:
            if kind == "workbook":
                items = app.Workbooks
            elif kind == "document":
                items = app.Documents
            else:
                items = app.Presentations
            for item in items:
                try:
                    full = item.FullName
                except Exception:
                    continue
                if not _same(full, path):
                    continue
                try:
                    modified = bool(item.Saved) is False  # Saved=False → 有未保存修改
                except Exception:
                    modified = None
                found.append({
                    "progid": progid, "kind": kind,
                    "fullname": full, "modified": modified,
                    "_item": item, "_app": app,
                })
        except Exception:
            continue
    return found


def validate_doc(doc, save):
    """未修改 → 关闭；有修改且 save → 保存关闭；否则返回需要用户决定。"""
    item, app = doc["_item"], doc["_app"]
    try:
        alerts = getattr(app, "DisplayAlerts", None)
        try:
            app.DisplayAlerts = 0  # wdAlertsNone / xlAlertsNone
        except Exception:
            pass
        if doc["kind"] == "workbook":
            item.Close(SaveChanges=bool(save))
        elif doc["kind"] == "document":
            item.Close(SaveChanges=-1 if save else 0)
        else:
            item.Close()
        if alerts is not None:
            try:
                app.DisplayAlerts = alerts
            except Exception:
                pass
        return True, None
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def wait_unlocked(path, timeout=8.0):
    end = time.time() + timeout
    while time.time() < end:
        if not is_locked(path):
            return True
        time.sleep(0.4)
    return not is_locked(path)


# ---------------------------------------------------------------- 兜底：窗口 / 强杀

def close_document_windows(path, pids=()):
    """按窗口标题里的文件名找文档窗口并发 WM_CLOSE。

    WPS 是多进程架构：文件锁在 et.exe/wps.exe 子进程上，但文档窗口属于
    另一个进程，所以不能只按 PID 找。标题匹配 + 已知 Office 窗口类双重确认。
    发送 WM_CLOSE 等价于点关闭按钮：有未保存修改时应用会自己弹保存提示，
    不会静默丢数据。
    """
    user32 = ctypes.windll.user32
    WM_CLOSE = 0x0010
    base = os.path.basename(path).lower()
    stem = os.path.splitext(base)[0]
    office_classes = {"OpusApp", "XLMAIN", "PPAMain", "framework::CHILD",
                      "wps", "et", "wpp"}

    hits = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        title = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, title, 512)
        if not title.value.strip():
            return True
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        p = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(p))

        t = title.value.lower()
        title_match = base in t or (len(stem) > 3 and stem in t)
        if title_match and (cls.value in office_classes or p.value in pids
                            or "wps" in t or "office" in t):
            hits.append((hwnd, title.value))
        return True

    user32.EnumWindows(cb, 0)
    for hwnd, _ in hits:
        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
    return [h[1] for h in hits]


def kill_pid(pid):
    try:
        subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                       capture_output=True, timeout=10)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- 主流程

def run(path, save=False, no_close=False, force_kill=False):
    path = os.path.abspath(path)
    result = {
        "file": path,
        "status": "unknown",
        "locked": False,
        "lock_files": lock_files(path),
        "owners": [],
        "message": "",
        "actions": [],
    }

    if not os.path.exists(path):
        result.update(status="not_found", message="文件不存在")
        return result, 1

    if not is_locked(path):
        result.update(status="not_locked", message="文件未被占用")
        return result, 0

    result["locked"] = True
    result["owners"] = restart_manager_owners(path)

    # 1) 优先走 COM：可以只关这一个文档，且能判断有无未保存修改
    docs = find_open_docs(path)
    if docs:
        doc = docs[0]
        result["doc"] = {"progid": doc["progid"], "fullname": doc["fullname"],
                         "modified": doc["modified"]}
        if no_close:
            result.update(status="locked",
                          message="文件被占用（--no-close：仅检测，未关闭任何文档）")
            return result, 3
        if doc["modified"] is None and not save:
            result.update(
                status="needs_decision",
                message="文档正被打开，且无法判断是否有未保存修改；请先询问用户。")
            return result, 2
        if doc["modified"] is True and not save:
            result.update(
                status="needs_decision",
                message="文档有未保存修改。请询问用户：保存并关闭 / 放弃修改关闭 / 取消。")
            return result, 2
        if doc["modified"] is True and force_kill:
            result.update(status="needs_decision",
                          message="有未保存修改，拒绝强杀。请询问用户。")
            return result, 2
        ok, err = validate_doc(doc, save)
        if ok and wait_unlocked(path):
            result.update(status="closed", locked=False,
                          message=f"已通过 {doc['progid']} 优雅关闭文档"
                                  + ("（已保存）" if save else "（无未保存修改）"))
            result["actions"].append(f"COM close {doc['progid']}")
            return result, 0
        if not ok:
            result["actions"].append(f"COM close failed: {err}")

    # 2) 兜底：按窗口标题找文档窗口，优雅关闭（默认行为）
    pids = sorted({o["pid"] for o in result["owners"] if o.get("pid")})
    if no_close:
        result.update(status="locked",
                      message="文件被占用（--no-close：仅检测，未关闭任何文档）")
        return result, 3
    else:
        titles = close_document_windows(path, pids)
        if titles:
            result["actions"].append("WM_CLOSE -> " + " | ".join(titles))
            if wait_unlocked(path, 12):
                result.update(status="closed", locked=False,
                              message="已关闭文档窗口（若弹出过保存提示，说明已按提示处理）")
                return result, 0
            result.update(
                status="needs_save_confirm",
                message="已发送关闭请求，但文档仍未关闭：多半是弹出了“是否保存”对话框。"
                        "请让用户在 WPS/Office 里选择保存或放弃，然后重试写入。")
            return result, 2

    if force_kill:
        for pid in pids:
            kill_pid(pid)
            result["actions"].append(f"taskkill /F -> {pid}")
        if wait_unlocked(path):
            result.update(status="closed", locked=False, message="已强制结束占用进程（未保存修改会丢失）")
            return result, 0

    owners_txt = "、".join(
        f"{o.get('process') or '?'}(PID {o.get('pid')})" for o in result["owners"]
    ) or "未知进程"
    result.update(
        status="locked_unknown",
        message=f"文件仍被占用：{owners_txt}。可请用户关闭该文档，或改用 --force-kill（需用户明确同意）。")
    return result, 3


def main():
    ap = argparse.ArgumentParser(description="检测并优雅解除文件占用")
    ap.add_argument("path")
    ap.add_argument("--save", action="store_true", help="有未保存修改时保存后关闭")
    ap.add_argument("--no-close", action="store_true", help="只检测，不关闭任何文档")
    ap.add_argument("--force-kill", action="store_true", help="强杀占用进程（危险）")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    args = ap.parse_args()

    result, code = run(args.path, args.save, args.no_close, args.force_kill)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"[{result['status']}] {result['message']}")
        if result.get("doc"):
            print(f"  文档: {result['doc']['fullname']} ({result['doc']['progid']}, "
                  f"modified={result['doc']['modified']})")
        if result["owners"]:
            print("  占用进程: " + ", ".join(
                f"{o.get('process')}(PID {o.get('pid')})" for o in result["owners"]))
        if result["lock_files"]:
            print("  锁文件: " + ", ".join(result["lock_files"]))
    return code


if __name__ == "__main__":
    sys.exit(main())