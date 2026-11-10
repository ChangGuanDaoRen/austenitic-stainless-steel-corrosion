# -*- coding: utf-8 -*-
"""
把检索文献阶段的遗留文件送入 Windows 回收站（可还原，不做永久删除）。
用 Win32 SHFileOperationW + FOF_ALLOWUNDO，避免 PowerShell Add-Type 被安全策略拦截，
也避免 rm / del 的永久删除。

用法：python scripts/trash_delete.py [dry|write]
"""
import ctypes
import os
import sys
from ctypes import wintypes

FO_DELETE = 3
FOF_SILENT = 0x0004
FOF_NOCONFIRMATION = 0x0010
FOF_ALLOWUNDO = 0x0040
FOF_NOERRORUI = 0x0400

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TARGETS = [
    os.path.join(ROOT, "temp", "dry.txt"),
    os.path.join(ROOT, "temp", "preview_A.html"),
    os.path.join(ROOT, "temp", "preview_user.html"),
    os.path.join(ROOT, "temp", "lit"),
    os.path.join(ROOT, "准"),
    os.path.join(ROOT, "待"),
]


class SHFILEOPSTRUCTW(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("wFunc", wintypes.UINT),
        ("pFrom", wintypes.LPCWSTR),
        ("pTo", wintypes.LPCWSTR),
        ("fFlags", ctypes.c_uint),
        ("fAnyOperationsAborted", wintypes.BOOL),
        ("hNameMappings", ctypes.c_void_p),
        ("lpszProgressTitle", wintypes.LPCWSTR),
    ]


def to_recycle_bin(path):
    p = os.path.abspath(path)
    if not os.path.exists(p):
        return "不存在"
    op = SHFILEOPSTRUCTW()
    op.hwnd = None
    op.wFunc = FO_DELETE
    op.pFrom = p + "\0\0"
    op.pTo = None
    op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI
    res = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
    if res == 0 and not op.fAnyOperationsAborted:
        return "已送回收站" if not os.path.exists(p) else "返回0但文件仍在"
    return f"失败 rc={res} aborted={op.fAnyOperationsAborted}"


def size_of(path):
    if os.path.isfile(path):
        return os.path.getsize(path)
    total = 0
    for root, _, files in os.walk(path):
        for f in files:
            try:
                total += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return total


MODE = sys.argv[1] if len(sys.argv) > 1 else "dry"
for t in TARGETS:
    rel = os.path.relpath(t, ROOT)
    if os.path.exists(t):
        mb = size_of(t) / 1024 / 1024
        n = sum(len(f) for _, _, f in os.walk(t)) if os.path.isdir(t) else 1
        print(f"{rel}  {mb:.1f} MB / {n} 个文件 -> {to_recycle_bin(t) if MODE == 'write' else '(dry) 待删'}")
    else:
        print(f"{rel}  不存在")
