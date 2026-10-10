# -*- coding: utf-8 -*-
"""打包答辩专用插件为单文件 exe（供 .bat / .ps1 调用）。"""

from __future__ import annotations

import datetime
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EXE_BASE_NAME = "WJF答辩专用插件"
_VERSION_RE = re.compile(r'^APP_VERSION\s*=\s*"[^"]*"', re.M)


def stamp_app_version() -> str:
    """把 notes_overlay.py 的 APP_VERSION 改为今天的 8 位日期并返回。"""
    version = datetime.date.today().strftime("%Y%m%d")
    src_path = ROOT / "notes_overlay.py"
    src = src_path.read_text(encoding="utf-8")
    new_src, n = _VERSION_RE.subn(f'APP_VERSION = "{version}"', src, count=1)
    if n != 1:
        raise RuntimeError("notes_overlay.py 中未找到 APP_VERSION")
    if new_src != src:
        src_path.write_text(new_src, encoding="utf-8", newline="\n")
    return version

# 用不到却会被连带打包的模块：减小 exe 体积、加快单文件解压启动
EXCLUDE_MODULES = (
    # pywin32 的 MFC 界面库（仅 makepy 图形进度条使用）
    "win32ui",
    "pywin",
    # 网络加密（hashlib 缺 _hashlib 时自动回退到内置实现）
    "ssl",
    "_ssl",
    "_hashlib",
    # lxml 中 python-docx 不用的部分，及其连带的 bs4 / charset_normalizer
    "lxml.html",
    "lxml.objectify",
    "lxml.isoschematron",
    "lxml.cssselect",
    "lxml.html5parser",
    "bs4",
    "charset_normalizer",
    "html5lib",
    # 标准库中与本工具无关的大模块
    "unittest",
    "pydoc",
    "doctest",
    "xmlrpc",
    "asyncio",
    "multiprocessing",
    "http.server",
    "ftplib",
)


def run(cmd: list[str]) -> None:
    print("+", " ".join(cmd), flush=True)
    subprocess.check_call(cmd, cwd=ROOT)


def main() -> int:
    exe_name = f"{EXE_BASE_NAME}_v{stamp_app_version()}"
    dist = ROOT / "dist"
    # 清掉旧版本 exe，dist 里只留本次产物
    for old in dist.glob("*答辩专用插件*.exe"):
        old.unlink()
    run(
        [
            sys.executable,
            "-m",
            "pip",
            "install",
            "-r",
            "requirements.txt",
            "-r",
            "requirements-dev.txt",
        ]
    )
    run(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--onefile",
            "--windowed",
            "--name",
            exe_name,
            "--icon",
            str(Path("assets") / "icon.ico"),
            "--add-data",
            r"assets\icon.ico;.",
            "--add-data",
            r"assets\icon.png;.",
            "--hidden-import=win32com.client",
            "--hidden-import=pythoncom",
            "--hidden-import=pywintypes",
            "--hidden-import=win32timezone",
            "--hidden-import=docx",
            "--collect-submodules=docx",
            "--hidden-import=lxml.etree",
            "--hidden-import=win32event",
            *(f"--exclude-module={m}" for m in EXCLUDE_MODULES),
            "notes_overlay.py",
        ]
    )
    # 只分发 exe；config.json 首次运行时自动生成
    stale_cfg = dist / "config.json"
    if stale_cfg.exists():
        stale_cfg.unlink()
    exe = dist / f"{exe_name}.exe"
    print(f"\nDone: {exe}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as exc:
        print(f"Build failed: exit {exc.returncode}", file=sys.stderr)
        raise SystemExit(exc.returncode)
