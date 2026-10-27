# -*- coding: utf-8 -*-
"""
PPT答辩备注优化插件 — PowerPoint 浮动备注 + 排练计时

放映时置顶显示当前页备注；支持字号调节、备注导出 Word。
翻页使用 PowerPoint 原生快捷键。
"""

from __future__ import annotations

import json
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont
from tkinter import filedialog, messagebox, ttk

try:
    import win32com.client
    import pythoncom
    import win32event
except ImportError:
    print("缺少 pywin32，请先执行: pip install pywin32")
    sys.exit(1)


PP_PLACEHOLDER_BODY = 2
# 计时中刷新界面的间隔；未计时时放慢，降低低配电脑占用
CLOCK_ACTIVE_MS = 250
CLOCK_IDLE_MS = 1000
# 页码轮询只是兜底：放映翻页靠 COM 事件即时推送，编辑视图切页靠轮询
POLL_SHOW_SEC = 0.2
POLL_IDLE_SEC = 0.4
RECONNECT_SEC = 1.0
# 已挂上 PPT 但长时间没有演示文稿时，释放后重新 GetActiveObject
EMPTY_REATTACH_SEC = 1.5
# 8 位日期；scripts/build_exe.py 打包时自动改为当天
APP_VERSION = "20261022"
DEFAULT_APP_TITLE = "PPT答辩备注优化插件"
CONFIG_FILE_NAME = "config.json"
DEFAULT_COUNTDOWN_MIN = 30
COUNTDOWN_FG = "#b0b0b0"
COUNTDOWN_WARN_FG = "#ff5252"
# 勾选「背景透明」时的窗口不透明度；取消勾选为 1（完全不透明）
WINDOW_ALPHA = 0.92
_NO_NOTES = "（本页无备注）"
# 备注里 #…# 括起的内容导出为 Word 批注，且不计入字数（全角＃同样识别）
_COMMENT_NOTE_RE = re.compile(r"[#＃]([^#＃]*)[#＃]")


def resource_path(name: str) -> Path:
    """定位资源文件（开发态 assets/，打包态 PyInstaller 临时目录）。"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bases = [Path(sys._MEIPASS)]
    else:
        root = Path(__file__).resolve().parent
        bases = [root / "assets", root]
    for base in bases:
        candidate = base / name
        if candidate.exists():
            return candidate
    return bases[0] / name


def app_dir() -> Path:
    """可写配置所在目录：开发态为源码目录，打包态为 exe 同目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _dir_writable(d: Path) -> bool:
    probe = d / ".write_test"
    try:
        d.mkdir(parents=True, exist_ok=True)
        probe.write_text("", encoding="utf-8")
        probe.unlink()
        return True
    except Exception:
        return False


_config_path_cache: Path | None = None


def config_path() -> Path:
    """优先 exe 同目录；该目录不可写时退到 %APPDATA%\\PPT答辩备注优化插件。"""
    global _config_path_cache
    if _config_path_cache is not None:
        return _config_path_cache
    local = app_dir()
    appdata = Path(os.environ.get("APPDATA") or Path.home()) / "PPT答辩备注优化插件"
    chosen = local / CONFIG_FILE_NAME
    if not chosen.exists():
        if (appdata / CONFIG_FILE_NAME).exists() or not _dir_writable(local):
            chosen = appdata / CONFIG_FILE_NAME
    _config_path_cache = chosen
    return chosen


def load_config() -> dict:
    """读取 config.json；不存在则写入默认配置。"""
    cfg_path = config_path()
    defaults = {
        "app_title": DEFAULT_APP_TITLE,
        "countdown_min": DEFAULT_COUNTDOWN_MIN,
    }
    if not cfg_path.exists():
        try:
            cfg_path.parent.mkdir(parents=True, exist_ok=True)
            cfg_path.write_text(
                json.dumps(defaults, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except Exception:
            pass
        return dict(defaults)

    try:
        data = json.loads(cfg_path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return dict(defaults)
        merged = dict(defaults)
        merged.update(data)
        title = str(merged.get("app_title") or "").strip()
        merged["app_title"] = title or DEFAULT_APP_TITLE
        try:
            merged["countdown_min"] = max(0, int(merged.get("countdown_min", DEFAULT_COUNTDOWN_MIN)))
        except Exception:
            merged["countdown_min"] = DEFAULT_COUNTDOWN_MIN
        merged.pop("countdown_sec", None)
        return merged
    except Exception:
        return dict(defaults)


def save_config(cfg: dict) -> None:
    cfg_path = config_path()
    try:
        cfg_path.parent.mkdir(parents=True, exist_ok=True)
        cfg_path.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    except Exception:
        pass


def format_hms(seconds: float) -> str:
    s = max(0, int(seconds))
    h, rem = divmod(s, 3600)
    m, sec = divmod(rem, 60)
    if h > 0:
        return f"{h:d}:{m:02d}:{sec:02d}"
    return f"{m:02d}:{sec:02d}"


# Word/XML 不允许 NULL 与多数控制字符（保留 \t \n \r）
_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]")


def sanitize_xml_text(text: object) -> str:
    """清洗备注文本，避免 python-docx 报 XML compatible 错误。"""
    if text is None:
        return ""
    s = str(text).replace("\r\n", "\n").replace("\r", "\n")
    s = _ILLEGAL_XML_RE.sub("", s)
    # 去掉不成对的代理项等非法 unicode
    s = s.encode("utf-8", "ignore").decode("utf-8")
    return s


def split_note_comments(text: str) -> tuple[str, list[str]]:
    """拆出正文与 #…# 批注；正文中去掉批注内容。"""
    comments = [sanitize_xml_text(c).strip() for c in _COMMENT_NOTE_RE.findall(text)]
    comments = [c for c in comments if c]
    body = _COMMENT_NOTE_RE.sub("", text)
    return body, comments


def count_visible_chars(text: str) -> int:
    """字数：去掉 #…# 批注与空白后的字符数。"""
    body, _ = split_note_comments(text)
    return len(re.sub(r"\s+", "", body))


def enable_dpi_awareness() -> None:
    """开启 DPI 感知，减轻高分屏字体发虚。"""
    try:
        from ctypes import windll

        try:
            # 2 = PROCESS_PER_MONITOR_DPI_AWARE
            windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def extract_notes(slide) -> str:
    """优先走备注正文占位符，失败再扫一遍。"""
    try:
        ph = slide.NotesPage.Shapes.Placeholders(PP_PLACEHOLDER_BODY)
        text = str(ph.TextFrame.TextRange.Text).replace("\r", "\n").strip()
        if text:
            return text
    except Exception:
        pass

    try:
        shapes = slide.NotesPage.Shapes
        parts: list[str] = []
        for i in range(1, shapes.Count + 1):
            shape = shapes(i)
            try:
                if not shape.HasTextFrame:
                    continue
                text = str(shape.TextFrame.TextRange.Text).replace("\r", "\n").strip()
                if not text:
                    continue
                try:
                    if int(shape.PlaceholderFormat.Type) == PP_PLACEHOLDER_BODY:
                        return text
                except Exception:
                    parts.append(text)
            except Exception:
                continue
        if parts:
            return "\n".join(parts)
    except Exception:
        pass
    return _NO_NOTES


def next_page_notes(slide) -> tuple[bool, str]:
    """下一页是否存在，以及下一页备注正文（无备注时为空串）。"""
    try:
        index = int(slide.SlideIndex)
        # slide.Parent 是 Presentation，幻灯片集合在 .Slides
        slides = slide.Parent.Slides
        total = int(slides.Count)
        if index < 1 or index >= total:
            return False, ""
        raw = extract_notes(slides(index + 1))
        if raw == _NO_NOTES:
            return True, ""
        return True, raw
    except Exception:
        return False, ""


def set_slide_notes(slide, text: str) -> bool:
    """把备注写回幻灯片的备注正文占位符；找不到占位符返回 False。"""
    ppt_text = text.replace("\r\n", "\n").replace("\n", "\r")
    try:
        ph = slide.NotesPage.Shapes.Placeholders(PP_PLACEHOLDER_BODY)
        ph.TextFrame.TextRange.Text = ppt_text
        return True
    except Exception:
        pass
    try:
        shapes = slide.NotesPage.Shapes
        for i in range(1, shapes.Count + 1):
            shape = shapes(i)
            try:
                if int(shape.PlaceholderFormat.Type) == PP_PLACEHOLDER_BODY:
                    shape.TextFrame.TextRange.Text = ppt_text
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


_PAGE_HEADING_RE = re.compile(r"^\s*第\s*(\d+)\s*页\s*$")
_FALLBACK_COMMENT_RE = re.compile(r"^\s*〔批注〕(.*)$")
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _read_docx_comments(doc) -> dict[str, str]:
    """读取 Word 批注：{批注 id: 文本}。直接解析 XML，兼容不带批注 API 的 python-docx。"""
    from lxml import etree

    result: dict[str, str] = {}
    try:
        rels = doc.part.rels.values()
    except Exception:
        return result
    for rel in rels:
        if not str(rel.reltype).endswith("/comments") or rel.is_external:
            continue
        try:
            root = etree.fromstring(rel.target_part.blob)
        except Exception:
            continue
        for c in root.iter(f"{{{_W_NS}}}comment"):
            cid = c.get(f"{{{_W_NS}}}id")
            paras = []
            for p in c.iter(f"{{{_W_NS}}}p"):
                paras.append("".join(t.text or "" for t in p.iter(f"{{{_W_NS}}}t")))
            if cid is not None:
                result[cid] = "\n".join(paras).strip()
    return result


def parse_notes_docx(path: str) -> list[tuple[int, str]]:
    """解析「导出备注」格式的 Word，返回 [(页码, 备注文本)]；Word 批注还原为 #…#。"""
    from docx import Document

    doc = Document(path)
    comments = _read_docx_comments(doc)
    paragraphs = list(doc.paragraphs)

    def is_heading(p) -> bool:
        try:
            name = str(p.style.name or "").lower()
        except Exception:
            name = ""
        return name.startswith("heading") or name.startswith("标题")

    marks = [
        (i, int(m.group(1)))
        for i, p in enumerate(paragraphs)
        if (m := _PAGE_HEADING_RE.match(p.text or "")) and is_heading(p)
    ]
    if not marks:
        marks = [
            (i, int(m.group(1)))
            for i, p in enumerate(paragraphs)
            if (m := _PAGE_HEADING_RE.match(p.text or ""))
        ]

    pages: list[tuple[int, str]] = []
    for k, (start, page_no) in enumerate(marks):
        end = marks[k + 1][0] if k + 1 < len(marks) else len(paragraphs)
        lines: list[str] = []
        page_comments: list[str] = []
        for p in paragraphs[start + 1 : end]:
            text = p.text or ""
            fb = _FALLBACK_COMMENT_RE.match(text)
            if fb:
                page_comments.append(fb.group(1).strip())
                continue
            if text.strip() in ("※", "（本页无备注）"):
                text = ""
            for ref in p._p.iter(f"{{{_W_NS}}}commentReference"):
                cmt = comments.get(ref.get(f"{{{_W_NS}}}id") or "", "")
                if cmt:
                    page_comments.append(cmt)
            lines.append(text.rstrip())
        body = "\n".join(lines).strip()
        tail = " ".join(f"#{c}#" for c in page_comments if c)
        pages.append((page_no, f"{body}\n\n{tail}" if body and tail else body or tail))
    return pages


class RehearsalTimer:
    def __init__(self) -> None:
        self.running = False
        self.session_start: float | None = None
        self.page_start: float | None = None
        self.frozen_total = 0.0
        self.frozen_page = 0.0
        self.current_index: int | None = None
        self.page_history: dict[int, float] = {}

    def on_show_start(self, slide_index: int) -> None:
        now = time.monotonic()
        self.running = True
        self.session_start = now
        self.page_start = now
        self.frozen_total = 0.0
        self.frozen_page = 0.0
        self.current_index = slide_index
        self.page_history.clear()

    def on_show_end(self) -> None:
        if self.running:
            self.frozen_total = self.total_seconds()
            self.frozen_page = self.page_seconds()
            if self.current_index is not None:
                prev = self.page_history.get(self.current_index, 0.0)
                self.page_history[self.current_index] = prev + self.frozen_page
        self.running = False
        self.session_start = None
        self.page_start = None

    def on_slide_change(self, slide_index: int) -> None:
        if not self.running:
            self.current_index = slide_index
            return
        now = time.monotonic()
        if self.current_index is not None and self.page_start is not None:
            elapsed = now - self.page_start
            prev = self.page_history.get(self.current_index, 0.0)
            self.page_history[self.current_index] = prev + elapsed
        self.current_index = slide_index
        self.page_start = now
        self.frozen_page = 0.0

    def total_seconds(self) -> float:
        if self.running and self.session_start is not None:
            return time.monotonic() - self.session_start
        return self.frozen_total

    def page_seconds(self) -> float:
        if self.running and self.page_start is not None:
            return time.monotonic() - self.page_start
        return self.frozen_page

    def page_times(self) -> list[tuple[int, float]]:
        """各页累计秒数（含当前页正在走的时间），按页码排序。"""
        hist = dict(self.page_history)
        if self.running and self.current_index is not None:
            hist[self.current_index] = hist.get(self.current_index, 0.0) + self.page_seconds()
        return sorted(hist.items(), key=lambda x: x[0])


class NotesOverlayApp:
    def __init__(self) -> None:
        self.config = load_config()
        self.app_title = str(self.config.get("app_title") or DEFAULT_APP_TITLE)
        self.timer = RehearsalTimer()
        self._was_showing = False
        self._last_index: int | None = None
        self._last_total: int | None = None
        self._last_notes = ""
        self._last_has_next = False
        self._last_next_notes = ""
        self._next_full_text = "下页：—"
        self._stop_event = threading.Event()

        # UI 合并：只应用最新一帧，避免 after 队列堆积造成卡顿
        self._pending_snap: dict | None = None
        self._flush_scheduled = False
        self._ui_lock = threading.Lock()
        self._font_size = 18
        self._font_min = 12
        self._font_max = 36
        self._ui_scale = 1.0
        self._page_times_win: tk.Toplevel | None = None
        self._countdown_deadline: float | None = None
        self._countdown_warned = False
        # UI 不直接调 PPT COM（会与后台线程死锁）；命令交给 worker 执行
        self._ppt_cmd_q: queue.Queue = queue.Queue()
        # 后台线程空闲时阻塞等待：COM 事件消息或此事件（有新命令/退出）到达即唤醒
        self._wake_evt = win32event.CreateEvent(None, False, False, None)

        enable_dpi_awareness()
        self.root = tk.Tk()
        self.root.title(f"{self.app_title} v{APP_VERSION}")
        self._setup_dpi_and_geometry()
        self.root.resizable(True, True)
        self.root.attributes("-topmost", True)
        self._apply_window_alpha(True)
        self.root.configure(bg="#1e1e1e")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_window_icon()

        self._drag_offset = (0, 0)
        self._build_ui()
        self.root.update_idletasks()
        self._min_w = max(900, self._header.winfo_reqwidth())
        self.root.minsize(self._min_w, 240)
        if self.root.winfo_width() < self._min_w:
            self.root.geometry(f"{self._min_w}x{self.root.winfo_height()}")
        self._bind_drag()

        self.root.after(CLOCK_IDLE_MS, self._tick_clock)
        threading.Thread(target=self._ppt_worker_loop, name="ppt-realtime", daemon=True).start()

    def _setup_dpi_and_geometry(self) -> None:
        """按屏幕 DPI 校正控件比例，默认窗口约占屏幕 1/4（保持约 16:5）。"""
        self.root.update_idletasks()
        try:
            dpi = float(self.root.winfo_fpixels("1i"))
        except Exception:
            dpi = 96.0
        # Tk 默认按 72dpi 解释 point；按实际 DPI 校正，避免高分屏按钮失衡
        scale = max(1.0, dpi / 96.0)
        self._ui_scale = scale
        try:
            self.root.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass

        sw = int(self.root.winfo_screenwidth())
        sh = int(self.root.winfo_screenheight())
        # 加宽默认窗口，避免顶栏按钮挤掉「本页/总计」时间
        target_area = (sw * sh) / 3.2
        # 约 16:4.5 => w=16k, h=4.5k, area=72 k^2
        k = (target_area / 72.0) ** 0.5
        w = max(980, int(16 * k))
        h = max(260, int(4.5 * k))
        w = min(w, max(980, sw - 40))
        h = min(h, max(260, sh - 80))
        h = max(260, int(w * 4.5 / 16))
        if h > sh - 80:
            h = sh - 80
            w = max(980, int(h * 16 / 4.5))
        x = max(0, (sw - w) // 2)
        y = max(20, (sh - h) // 5)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(900, 240)

    def _set_window_icon(self) -> None:
        ico = resource_path("icon.ico")
        png = resource_path("icon.png")
        if not ico.exists():
            ico = resource_path("scu_logo.ico")
        if not png.exists():
            png = resource_path("scu_logo.png")
        try:
            if ico.exists():
                self.root.iconbitmap(default=str(ico))
        except Exception:
            pass
        try:
            if png.exists():
                self._icon_img = tk.PhotoImage(file=str(png))
                self.root.iconphoto(True, self._icon_img)
        except Exception:
            pass

    def _build_ui(self) -> None:
        s = self._ui_scale
        # 字号用相对“逻辑点”，配合 tk scaling，高分屏更稳
        title_pt = 10
        timer_pt = 14
        self.title_font = tkfont.Font(family="Microsoft YaHei UI", size=title_pt, weight="bold")
        self.timer_font = tkfont.Font(family="Consolas", size=timer_pt, weight="bold")
        self.body_font = tkfont.Font(family="Microsoft YaHei UI", size=self._font_size)
        self.caption_font = tkfont.Font(family="Microsoft YaHei UI", size=8)

        header = tk.Frame(self.root, bg="#2b2b2b")
        header.pack(fill="x", side="top")
        self._header = header

        row1 = tk.Frame(header, bg="#2b2b2b")
        row1.pack(fill="x", padx=int(10 * s), pady=(int(6 * s), 2))
        row2 = tk.Frame(header, bg="#2b2b2b")
        row2.pack(fill="x", padx=int(10 * s), pady=(2, int(8 * s)))

        left = tk.Frame(row1, bg="#2b2b2b")
        left.pack(side="left")

        self.status_var = tk.StringVar(value="页码 —")
        self.status_label = tk.Label(
            left,
            textvariable=self.status_var,
            fg="#ffffff",
            bg="#2b2b2b",
            font=self.title_font,
            anchor="w",
        )
        self.status_label.pack(anchor="w")

        self.state_var = tk.StringVar(value="计时未开始（请点「幻灯片放映」）")
        tk.Label(
            left,
            textvariable=self.state_var,
            fg="#c8c8c8",
            bg="#2b2b2b",
            font=self.title_font,
            anchor="w",
        ).pack(anchor="w")

        timers = tk.Frame(row1, bg="#2b2b2b")
        timers.pack(side="left", padx=(int(24 * s), 0))
        self.page_time_var = tk.StringVar(value="本页 00:00")
        self.total_time_var = tk.StringVar(value="总计 00:00")
        tk.Label(
            timers,
            textvariable=self.page_time_var,
            fg="#7fd99a",
            bg="#2b2b2b",
            font=self.timer_font,
        ).pack(side="left", padx=(0, int(20 * s)))
        tk.Label(
            timers,
            textvariable=self.total_time_var,
            fg="#7ec8ff",
            bg="#2b2b2b",
            font=self.timer_font,
        ).pack(side="left", padx=(0, int(20 * s)))

        cfg_min = int(self.config.get("countdown_min", DEFAULT_COUNTDOWN_MIN))
        self.countdown_display_var = tk.StringVar(value=f"倒计时 {format_hms(cfg_min * 60)}")
        self.countdown_label = tk.Label(
            timers,
            textvariable=self.countdown_display_var,
            fg=COUNTDOWN_FG,
            bg="#2b2b2b",
            font=self.timer_font,
        )
        self.countdown_label.pack(side="left")

        check_style = {
            "fg": "#eeeeee",
            "bg": "#2b2b2b",
            "selectcolor": "#2b2b2b",
            "activebackground": "#2b2b2b",
            "activeforeground": "#ffffff",
            "font": self.title_font,
        }
        self.pin_var = tk.BooleanVar(value=True)
        pin_btn = tk.Checkbutton(
            row1,
            text="置顶",
            variable=self.pin_var,
            command=self._toggle_topmost,
            **check_style,
        )
        pin_btn.pack(side="right", padx=(int(8 * s), 0))
        self.alpha_var = tk.BooleanVar(value=True)
        alpha_btn = tk.Checkbutton(
            row1,
            text="背景透明",
            variable=self.alpha_var,
            command=self._toggle_alpha,
            **check_style,
        )
        alpha_btn.pack(side="right", padx=(int(8 * s), 0))

        zoom_style = {
            "font": self.title_font,
            "relief": "flat",
            "bd": 0,
            "padx": max(6, int(8 * s)),
            "pady": max(2, int(3 * s)),
            "cursor": "hand2",
            "bg": "#3a3a3a",
            "fg": "#ffffff",
            "activebackground": "#505050",
            "activeforeground": "#ffffff",
        }
        action_style = {
            **zoom_style,
            "bg": "#2e7d32",
            "activebackground": "#388e3c",
        }
        stop_style = {
            **zoom_style,
            "bg": "#ef6c00",
            "activebackground": "#f57c00",
        }
        self._show_btn_style = action_style
        self._stop_btn_style = stop_style
        self._show_btn = tk.Button(
            row2, text="幻灯片放映", command=self._start_slideshow, **action_style
        )
        self._show_btn.pack(side="left", padx=(0, int(6 * s)))
        tk.Button(row2, text="本页时长", command=self._show_page_times_dialog, **zoom_style).pack(
            side="left", padx=(0, int(6 * s))
        )
        tk.Button(row2, text="导出备注", command=self._export_notes_txt, **zoom_style).pack(
            side="left", padx=(0, int(6 * s))
        )
        tk.Button(row2, text="导入备注", command=self._import_notes_docx, **zoom_style).pack(
            side="left", padx=(0, int(6 * s))
        )
        tk.Button(row2, text="清除备注", command=self._clear_notes_dialog, **zoom_style).pack(
            side="left", padx=(0, int(6 * s))
        )
        tk.Button(row2, text="A−", command=lambda: self._change_font(-2), **zoom_style).pack(
            side="left", padx=(0, int(4 * s))
        )
        tk.Button(row2, text="A+", command=lambda: self._change_font(+2), **zoom_style).pack(
            side="left", padx=(0, int(16 * s))
        )

        tk.Label(
            row2,
            text="倒计时",
            fg=COUNTDOWN_FG,
            bg="#2b2b2b",
            font=self.title_font,
        ).pack(side="left", padx=(0, int(4 * s)))
        self.countdown_min_var = tk.StringVar(value=str(cfg_min))
        self._countdown_min_entry = tk.Entry(
            row2,
            textvariable=self.countdown_min_var,
            width=4,
            font=self.title_font,
            justify="center",
            bg="#333333",
            fg="#ffffff",
            relief="flat",
            insertbackground="#ffffff",
            highlightthickness=1,
            highlightbackground="#444444",
            highlightcolor="#666666",
        )
        self._countdown_min_entry.pack(side="left")
        tk.Label(row2, text="分钟", fg="#aaaaaa", bg="#2b2b2b", font=self.title_font).pack(
            side="left", padx=(2, 0)
        )
        self._countdown_min_entry.bind("<FocusOut>", lambda _e: self._on_countdown_setting_change())
        self._countdown_min_entry.bind("<Return>", lambda _e: self._on_countdown_setting_change())
        tk.Button(row2, text="使用说明", command=self._show_help_dialog, **zoom_style).pack(
            side="right"
        )

        # 底部先 pack，再放中间备注区；否则 expand 的 body 会把底部挤没
        grip_font = tkfont.Font(family="Microsoft YaHei UI", size=8)
        grip = tk.Frame(self.root, bg="#1e1e1e")
        grip.pack(fill="x", side="bottom")
        tk.Label(
            grip,
            text="Ctrl+滚轮 调字号",
            fg="#888888",
            bg="#1e1e1e",
            font=grip_font,
            anchor="w",
        ).pack(side="left", padx=8, pady=1)

        resize_box = tk.Frame(grip, bg="#2b2b2b", cursor="size_nw_se")
        resize_box.pack(side="right", padx=4, pady=1)
        tk.Label(
            grip,
            text=f"v{APP_VERSION}",
            fg="#666666",
            bg="#1e1e1e",
            font=grip_font,
        ).pack(side="right", padx=(0, 4), pady=1)
        self._resize_icon = tk.Label(
            resize_box,
            text="↘ 拖拽改大小",
            fg="#cccccc",
            bg="#2b2b2b",
            font=grip_font,
            cursor="size_nw_se",
            padx=6,
            pady=1,
        )
        self._resize_icon.pack()
        for w in (resize_box, self._resize_icon):
            w.bind("<ButtonPress-1>", self._start_resize)
            w.bind("<B1-Motion>", self._do_resize)

        body = tk.Frame(self.root, bg="#1e1e1e")
        body.pack(fill="both", expand=True, padx=10, pady=(8, 4))

        self._next_holder = tk.Frame(body, bg="#1a1a1a")
        self._next_holder.pack(side="bottom", fill="x", pady=(6, 0))
        self.next_var = tk.StringVar(value=self._next_full_text)
        self.next_label = tk.Label(
            self._next_holder,
            textvariable=self.next_var,
            fg="#8d8d8d",
            bg="#1a1a1a",
            font=self.caption_font,
            anchor="w",
            padx=4,
        )
        self.next_label.pack(side="left", fill="x", expand=True, padx=(0, 22))
        self._next_holder.bind("<Configure>", self._fit_next_label)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Dark.Vertical.TScrollbar",
            background="#666666",
            troughcolor="#2a2a2a",
            bordercolor="#2a2a2a",
            lightcolor="#666666",
            darkcolor="#666666",
            arrowcolor="#ffffff",
            relief="flat",
            width=max(14, int(16 * s)),
        )
        style.map(
            "Dark.Vertical.TScrollbar",
            background=[("active", "#888888"), ("pressed", "#aaaaaa")],
            arrowcolor=[("active", "#ffffff")],
        )
        self._notes_scroll = ttk.Scrollbar(
            body,
            orient="vertical",
            style="Dark.Vertical.TScrollbar",
            command=self._notes_yview,
        )
        self._notes_scroll.pack(side="right", fill="y")

        self.notes = tk.Text(
            body,
            wrap="word",
            font=self.body_font,
            bg="#1a1a1a",
            fg="#ffffff",
            relief="flat",
            padx=12,
            pady=10,
            borderwidth=0,
            highlightthickness=0,
            cursor="arrow",
            insertwidth=0,
            takefocus=0,
            state="disabled",
            yscrollcommand=self._notes_scroll.set,
        )
        self.notes.pack(side="left", fill="both", expand=True)
        self.notes.tag_configure("body", foreground="#ffffff")
        self._set_notes_if_changed("正在连接 PowerPoint…")
        self.notes.bind("<MouseWheel>", self._on_notes_wheel)
        body.bind_all("<MouseWheel>", self._on_notes_wheel_global)

        self.root.bind("<Control-plus>", lambda _e: self._change_font(+2))
        self.root.bind("<Control-equal>", lambda _e: self._change_font(+2))
        self.root.bind("<Control-minus>", lambda _e: self._change_font(-2))
        self.root.bind("<Control-MouseWheel>", self._on_ctrl_wheel)
        self.notes.bind("<Control-MouseWheel>", self._on_ctrl_wheel)
        # 放映时焦点常在本窗口，把翻页键转发给 PPT
        for _key in (
            "Left",
            "Right",
            "Up",
            "Down",
            "Prior",
            "Next",
            "space",
            "Return",
            "BackSpace",
            "Escape",
        ):
            self.root.bind_all(f"<{_key}>", self._on_nav_key, add="+")

    def _notes_yview(self, *args) -> None:
        self.notes.yview(*args)

    def _on_notes_wheel(self, event) -> str:
        self.notes.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _on_notes_wheel_global(self, event) -> str | None:
        try:
            x, y = self.root.winfo_pointerxy()
            widget = self.root.winfo_containing(x, y)
            if widget is self.notes or str(widget).startswith(str(self.notes)):
                self.notes.yview_scroll(int(-1 * (event.delta / 120)), "units")
                return "break"
        except Exception:
            pass
        return None

    def _start_resize(self, event) -> None:
        self._resize_origin = (
            event.x_root,
            event.y_root,
            self.root.winfo_width(),
            self.root.winfo_height(),
        )

    def _do_resize(self, event) -> None:
        if not hasattr(self, "_resize_origin"):
            return
        x0, y0, w0, h0 = self._resize_origin
        w = max(self._min_w, w0 + (event.x_root - x0))
        h = max(240, h0 + (event.y_root - y0))
        self.root.geometry(f"{w}x{h}")

    def _change_font(self, delta: int) -> None:
        new_size = max(self._font_min, min(self._font_max, self._font_size + delta))
        if new_size == self._font_size:
            return
        self._font_size = new_size
        self.body_font.configure(size=self._font_size)

    def _on_ctrl_wheel(self, event) -> str:
        self._change_font(+2 if event.delta > 0 else -2)
        return "break"

    def _toggle_topmost(self) -> None:
        self.root.attributes("-topmost", bool(self.pin_var.get()))

    def _format_next_line(self, index: int, total: int, has_next: bool, notes: str) -> str:
        if index <= 0 or total <= 0:
            return "下页：—"
        if index >= total:
            return "下页：（最后一页）"
        if not has_next:
            return "下页：—"
        one = " ".join((notes or "").split())
        if not one:
            return "下页：（无备注）"
        return "下页：" + one

    def _show_next_notes(self, index: int, total: int, has_next: bool, notes: str) -> None:
        self._last_has_next = has_next
        self._last_next_notes = notes
        self._next_full_text = self._format_next_line(index, total, has_next, notes)
        self._fit_next_label()

    def _fit_next_label(self, event=None) -> None:
        raw = self._next_full_text or "下页：—"
        try:
            width = int(event.width) if event is not None else int(self._next_holder.winfo_width())
        except Exception:
            width = 0
        width = max(40, width - 28)
        text = raw
        if self.caption_font.measure(text) > width:
            ell = "…"
            lo, hi = 0, len(raw)
            while lo < hi:
                mid = (lo + hi + 1) // 2
                if self.caption_font.measure(raw[:mid] + ell) <= width:
                    lo = mid
                else:
                    hi = mid - 1
            text = raw[:lo] + ell if lo else ell
        self._set_var(self.next_var, text)

    def _apply_window_alpha(self, transparent: bool) -> None:
        try:
            self.root.attributes("-alpha", WINDOW_ALPHA if transparent else 1.0)
        except Exception:
            pass

    def _toggle_alpha(self) -> None:
        self._apply_window_alpha(bool(self.alpha_var.get()))

    def _show_help_dialog(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("使用说明")
        win.configure(bg="#1e1e1e")
        win.attributes("-topmost", True)
        s = self._ui_scale
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w = min(int(820 * s), sw - 40)
        h = min(int(760 * s), sh - 80)
        win.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")
        win.minsize(int(600 * s), int(480 * s))

        help_title_font = tkfont.Font(family="Microsoft YaHei UI", size=16, weight="bold")
        help_body_font = tkfont.Font(family="Microsoft YaHei UI", size=13)
        help_head_font = tkfont.Font(family="Microsoft YaHei UI", size=14, weight="bold")

        tk.Label(
            win,
            text=f"{self.app_title} v{APP_VERSION} · 功能说明",
            fg="#ffffff",
            bg="#1e1e1e",
            font=help_title_font,
            anchor="w",
        ).pack(fill="x", padx=int(16 * s), pady=(int(14 * s), int(8 * s)))

        text = tk.Text(
            win,
            wrap="word",
            bg="#252525",
            fg="#eeeeee",
            relief="flat",
            padx=int(16 * s),
            pady=int(12 * s),
            font=help_body_font,
            spacing1=int(2 * s),
            spacing3=int(2 * s),
            highlightthickness=0,
        )
        text.tag_configure("head", font=help_head_font, foreground="#ffffff")
        text.pack(fill="both", expand=True, padx=int(16 * s), pady=4)
        help_body = (
            "【幻灯片放映】\n"
            "启动当前 PPT 的幻灯片放映，并开始本页计时、总计时与倒计时。"
            "在 PowerPoint 里自己按 F5 开的放映不会自动计时，必须点本按钮。\n\n"
            "【本页时长】\n"
            "查看各页累计停留时间（最长页会高亮），可导出 CSV。"
            "放映结束后也会自动弹出。\n\n"
            "【导出备注】\n"
            "把当前演示文稿全部备注导出为 Word（英文 Times New Roman，中文宋体）。"
            "备注里用 #…# 括起的文字会变成 Word 批注，且不计入正文字数。\n\n"
            "【导入备注】\n"
            "选择与「导出备注」同格式的 Word，按「第 N 页」标题逐页覆盖当前 PPT 的备注；"
            "Word 批注会用 #…# 包裹，空一行放在该页备注最后。"
            "Word 里没有的页保持不变；写入后需自行保存 PPT。\n\n"
            "【清除备注】\n"
            "可清除当前页或全部页的备注；清除前会确认。"
            "操作不可撤销（除非 PPT 尚未保存且可从备份恢复），完成后请自行保存 PPT。\n\n"
            "【A− / A+】\n"
            "缩小 / 放大备注字号；也可用 Ctrl+滚轮。\n\n"
            "【置顶】\n"
            "窗口是否始终浮在最前，录屏时可视情况关闭。\n\n"
            "【背景透明】\n"
            "默认勾选，窗口略透明。取消勾选后窗口完全不透明。\n\n"
            "【下页】\n"
            "备注区底部用灰色显示下一页备注的开头，方便翻页前看一眼。\n\n"
            "【倒计时设置】\n"
            "在按钮行右侧设置答辩限时（分钟，会记住）。"
            "点击「幻灯片放映」后开始倒数，剩余时间显示在顶部「总计」右侧，"
            "最后 1 分钟变红，到 0 会提示。\n\n"
            "【本页 / 总计 / 倒计时】\n"
            "本页：当前页停留时长；总计：本次计时会话总时长；倒计时：剩余时长。\n\n"
            "使用建议：先打开带备注的 PPT → 设好倒计时 → 点「幻灯片放映」。"
            "翻页请用 PowerPoint 原生快捷键（→ / ← / 空格等）。"
        )
        for line in help_body.splitlines(keepends=True):
            is_head = line.startswith("【") and line.rstrip().endswith("】")
            text.insert("end", line, ("head",) if is_head else ())
        text.configure(state="disabled")

        tk.Button(
            win,
            text="知道了",
            command=win.destroy,
            font=help_head_font,
            relief="flat",
            bg="#3a3a3a",
            fg="#ffffff",
            activebackground="#505050",
            cursor="hand2",
            padx=int(20 * s),
            pady=int(6 * s),
        ).pack(pady=int(14 * s))
        win.focus_force()

    def _read_countdown_setting_secs(self) -> int:
        try:
            mins = max(0, int(str(self.countdown_min_var.get()).strip() or "0"))
        except Exception:
            mins = DEFAULT_COUNTDOWN_MIN
        self.countdown_min_var.set(str(mins))
        return mins * 60

    def _on_countdown_setting_change(self) -> None:
        total = self._read_countdown_setting_secs()
        self.config["countdown_min"] = total // 60
        save_config(self.config)
        if self._countdown_deadline is None and not self.timer.running:
            self._set_countdown_text(total)
            self._set_countdown_fg(COUNTDOWN_FG)

    @staticmethod
    def _set_var(var: tk.StringVar, value: str) -> None:
        """值没变就不写，避免 Tk 无谓重绘。"""
        if var.get() != value:
            var.set(value)

    def _set_countdown_text(self, seconds: float) -> None:
        self._set_var(self.countdown_display_var, f"倒计时 {format_hms(seconds)}")

    def _set_countdown_fg(self, color: str) -> None:
        if str(self.countdown_label.cget("fg")) != color:
            self.countdown_label.configure(fg=color)

    def _start_countdown(self) -> None:
        total = self._read_countdown_setting_secs()
        self._countdown_deadline = time.monotonic() + total
        self._countdown_warned = False
        self._set_countdown_text(total)
        self._set_countdown_fg(COUNTDOWN_FG)

    def _stop_countdown(self, reset_display: bool = True) -> None:
        self._countdown_deadline = None
        if reset_display:
            self._set_countdown_text(self._read_countdown_setting_secs())
            self._set_countdown_fg(COUNTDOWN_FG)

    def _update_countdown_display(self) -> None:
        if self._countdown_deadline is None:
            return
        remain = max(0, int(self._countdown_deadline - time.monotonic() + 0.999))
        self._set_countdown_text(remain)
        if remain <= 0:
            self._set_countdown_fg(COUNTDOWN_WARN_FG)
            if not self._countdown_warned:
                self._countdown_warned = True
                self._countdown_deadline = None
                try:
                    self.root.bell()
                except Exception:
                    pass
                self.state_var.set("倒计时结束！")
        elif remain <= 60:
            self._set_countdown_fg(COUNTDOWN_WARN_FG)
        else:
            self._set_countdown_fg(COUNTDOWN_FG)

    def _begin_timing_session(self) -> None:
        idx = self._last_index if self._last_index and self._last_index > 0 else 1
        self.timer.on_show_start(idx)
        self._start_countdown()
        self.state_var.set("排练计时中…")

    # 放映翻页：本窗口有焦点时转发到 PPT（与 PowerPoint 默认一致）
    _NAV_KEYSYMS = {
        "Left": "prev",
        "Up": "prev",
        "Prior": "prev",
        "BackSpace": "prev",
        "Right": "next",
        "Down": "next",
        "Next": "next",
        "space": "next",
        "Return": "next",
        "Escape": "exit",
    }

    def _focus_in_editable(self) -> bool:
        w = self.root.focus_get()
        if w is None:
            return False
        try:
            cls = w.winfo_class()
        except Exception:
            return False
        return cls in ("TSpinbox", "Spinbox", "Entry", "TEntry")

    def _should_forward_nav(self) -> bool:
        if not self._was_showing:
            return False
        if self._focus_in_editable():
            return False
        try:
            if self.root.grab_current() is not None:
                return False
        except Exception:
            pass
        w = self.root.focus_get()
        if w is not None:
            try:
                if w.winfo_toplevel() is not self.root:
                    return False
            except Exception:
                pass
        return True

    def _forward_slideshow_key(self, event) -> bool:
        """放映中把方向键等转发给 PPT；返回是否已处理。"""
        if not self._should_forward_nav():
            return False
        direction = self._NAV_KEYSYMS.get(event.keysym)
        if not direction:
            return False
        self._send_ppt_cmd({"op": "nav", "dir": direction})
        return True

    def _on_nav_key(self, event):
        if self._forward_slideshow_key(event):
            return "break"

    def _reactivate_slideshow(self) -> None:
        """把键盘焦点交回普通放映窗口。"""
        self._send_ppt_cmd({"op": "activate"})

    def _update_show_button(self, showing: bool) -> None:
        """放映中显示「停止放映」，否则「幻灯片放映」。"""
        if showing:
            self._show_btn.configure(
                text="停止放映", command=self._stop_slideshow, **self._stop_btn_style
            )
        else:
            self._show_btn.configure(
                text="幻灯片放映", command=self._start_slideshow, **self._show_btn_style
            )

    def _start_slideshow(self) -> None:
        self.state_var.set("正在启动幻灯片放映…")
        self._send_ppt_cmd({"op": "show"})

    def _stop_slideshow(self) -> None:
        self.state_var.set("正在停止放映…")
        self._send_ppt_cmd({"op": "nav", "dir": "exit"})

    def _post_ppt_cmd_result(self, result: dict) -> None:
        if self._stop_event.is_set():
            return
        try:
            self.root.after(0, lambda r=result: self._handle_ppt_cmd_result(r))
        except Exception:
            pass

    def _handle_ppt_cmd_result(self, result: dict) -> None:
        op = str(result.get("op") or "")
        if op in ("import_notes", "clear_notes"):
            title = "导入备注" if op == "import_notes" else "清除备注"
            if not result.get("ok"):
                messagebox.showerror(title, f"{title}失败：\n{result.get('error') or '操作失败'}")
                return
            if op == "import_notes":
                msg = f"已写入 {result.get('written', 0)} 页备注。"
            else:
                msg = f"已清除 {result.get('written', 0)} 页备注。"
            skipped = result.get("skipped") or []
            failed = result.get("failed") or []
            if skipped:
                msg += f"\n\n超出 PPT 页数、已忽略：第 {', '.join(map(str, skipped))} 页"
            if failed:
                msg += f"\n\n找不到备注占位符、未写入：第 {', '.join(map(str, failed))} 页"
            msg += "\n\n请在 PowerPoint 中保存演示文稿。"
            messagebox.showinfo(title, msg)
            return
        if not result.get("ok"):
            self.state_var.set("计时未开始（请点「幻灯片放映」）")
            messagebox.showwarning("幻灯片放映", str(result.get("error") or "操作失败"))
            return
        info = result.get("info")
        self._begin_timing_session()
        if info:
            self.state_var.set(str(info))
        # 普通放映：把激活权还给放映窗
        self._reactivate_slideshow()
        self.root.after(80, self._reactivate_slideshow)

    def _export_notes_txt(self) -> None:
        """逐页导出当前 PPT 全部备注为 Word（.docx）。"""
        was_top = bool(self.pin_var.get())
        if was_top:
            self.root.attributes("-topmost", False)

        try:
            from docx import Document
            from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
            from docx.oxml.ns import qn
            from docx.shared import Pt

            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                messagebox.showwarning("导出备注", "未检测到已打开的 PowerPoint，请先打开演示文稿。")
                return

            try:
                if app.Presentations.Count < 1:
                    messagebox.showwarning("导出备注", "当前没有打开的演示文稿。")
                    return
                pres = app.ActivePresentation
            except Exception as exc:
                messagebox.showerror("导出备注", f"无法读取演示文稿：\n{exc}")
                return

            try:
                title = sanitize_xml_text(pres.Name)
            except Exception:
                title = "PPT备注"
            stem = sanitize_xml_text(Path(title).stem) or "PPT备注"
            # 文件名再去掉 Windows 非法字符
            safe_stem = re.sub(r'[\\/:*?"<>|]', "_", stem).strip() or "PPT备注"
            default_name = f"{safe_stem}_备注.docx"

            path = filedialog.asksaveasfilename(
                title="导出逐页备注为 Word",
                defaultextension=".docx",
                initialfile=default_name,
                filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            )
            if not path:
                return

            doc = Document()
            latin_font, cjk_font = "Times New Roman", "宋体"

            def apply_fonts(font) -> None:
                font.name = latin_font
                rpr = font.element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
                rfonts.set(qn("w:eastAsia"), cjk_font)
                rfonts.set(qn("w:cs"), latin_font)
                # 主题字体属性优先级高于显式字体，必须去掉
                for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                    rfonts.attrib.pop(qn(attr), None)

            def apply_paragraph_format(pf) -> None:
                pf.line_spacing = 1.5
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)

            for style_name in ("Normal", "Heading 1"):
                try:
                    style = doc.styles[style_name]
                except KeyError:
                    continue
                apply_fonts(style.font)
                apply_paragraph_format(style.paragraph_format)

            def set_run_font(run, size_pt: float, bold: bool = False) -> None:
                run.bold = bold
                run.font.size = Pt(size_pt)
                apply_fonts(run.font)

            def add_text_run(paragraph, text: str, size_pt: float, bold: bool = False) -> None:
                clean = sanitize_xml_text(text)
                # 全角空格若落到 Times New Roman 会显示成方块，单独成段并整段用宋体
                parts = re.split(r"(\u3000+)", clean) if "\u3000" in clean else [clean]
                for part in parts:
                    if not part and len(parts) > 1:
                        continue
                    run = paragraph.add_run(part)
                    set_run_font(run, size_pt, bold=bold)
                    if part.startswith("\u3000"):
                        rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
                        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
                            rfonts.set(qn(attr), cjk_font)

            # 文档标题：一级标题加粗
            p_title = doc.add_heading("", level=1)
            add_text_run(p_title, stem, 16, bold=True)
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            total = int(pres.Slides.Count)
            pages_notes: list[str] = []
            for i in range(1, total + 1):
                pages_notes.append(sanitize_xml_text(extract_notes(pres.Slides(i))))
            char_count = sum(count_visible_chars(n) for n in pages_notes)
            comment_count = sum(len(split_note_comments(n)[1]) for n in pages_notes)

            # 副信息正文
            p_meta = doc.add_paragraph()
            add_text_run(
                p_meta,
                (
                    f"演示文稿：{title}    总页数：{total}"
                    f"    正文字数：{char_count}（不含 #…# 批注）"
                ),
                11,
                bold=False,
            )

            # 标题与正文之间空一行
            doc.add_paragraph()

            for i, notes in enumerate(pages_notes, start=1):
                # 页码：一级标题加粗
                p_h = doc.add_heading("", level=1)
                add_text_run(p_h, f"第 {i} 页", 14, bold=True)

                body, comments = split_note_comments(notes)
                body = body.strip()
                anchor_para = None
                if not body or notes.strip() == "（本页无备注）":
                    if comments:
                        p = doc.add_paragraph()
                        add_text_run(p, "※", 12, bold=False)
                        anchor_para = p
                    else:
                        p = doc.add_paragraph()
                        add_text_run(p, "（本页无备注）", 12, bold=False)
                else:
                    for para_text in body.split("\n"):
                        p = doc.add_paragraph()
                        add_text_run(p, para_text.rstrip(), 12, bold=False)
                        anchor_para = p

                if comments and anchor_para is not None and anchor_para.runs:
                    for cmt in comments:
                        try:
                            doc.add_comment(
                                runs=anchor_para.runs[-1],
                                text=cmt,
                                author="备注批注",
                                initials="注",
                            )
                        except Exception:
                            # 旧版 python-docx 无批注 API 时，退化为脚注式正文
                            p_fb = doc.add_paragraph()
                            add_text_run(p_fb, f"〔批注〕{cmt}", 10, bold=False)

                if i < total:
                    doc.add_paragraph()

            for p in doc.paragraphs:
                apply_paragraph_format(p.paragraph_format)
            doc.save(path)
            messagebox.showinfo(
                "导出备注",
                (
                    f"已导出 {total} 页备注到：\n{path}\n\n"
                    f"正文字数：{char_count}（#…# 内不计入）\n"
                    f"批注条数：{comment_count}"
                ),
            )
        except Exception as exc:
            messagebox.showerror("导出备注", f"导出失败：\n{exc}")
        finally:
            if was_top and bool(self.pin_var.get()):
                self.root.attributes("-topmost", True)

    def _import_notes_docx(self) -> None:
        """从「导出备注」格式的 Word 逐页写回当前 PPT 备注。"""
        was_top = bool(self.pin_var.get())
        if was_top:
            self.root.attributes("-topmost", False)
        try:
            path = filedialog.askopenfilename(
                title="选择要导入的备注 Word",
                filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
            )
            if not path:
                return
            try:
                pages = parse_notes_docx(path)
            except Exception as exc:
                messagebox.showerror("导入备注", f"无法读取 Word：\n{exc}")
                return
            if not pages:
                messagebox.showwarning(
                    "导入备注",
                    "没有找到「第 N 页」标题。\n请使用「导出备注」生成的 Word 格式。",
                )
                return
            nums = [n for n, _ in pages]
            if not messagebox.askyesno(
                "导入备注",
                (
                    f"Word 中共 {len(pages)} 页备注（第 {min(nums)}–{max(nums)} 页）。\n\n"
                    "将覆盖当前 PPT 对应页的备注，是否继续？"
                ),
            ):
                return
            self._send_ppt_cmd({"op": "import_notes", "pages": pages})
        finally:
            if was_top and bool(self.pin_var.get()):
                self.root.attributes("-topmost", True)

    def _clear_notes_dialog(self) -> None:
        """清除当前页或全部页备注（二次确认）。"""
        was_top = bool(self.pin_var.get())
        if was_top:
            self.root.attributes("-topmost", False)
        try:
            win = tk.Toplevel(self.root)
            win.title("清除备注")
            win.configure(bg="#1e1e1e")
            win.attributes("-topmost", True)
            win.resizable(False, False)
            win.grab_set()

            tk.Label(
                win,
                text="请选择要清除的范围（不可撤销）",
                fg="#eeeeee",
                bg="#1e1e1e",
                font=self.title_font,
            ).pack(padx=20, pady=(16, 10))

            btns = tk.Frame(win, bg="#1e1e1e")
            btns.pack(padx=20, pady=(0, 16))

            def choose(scope: str) -> None:
                win.destroy()
                if scope == "all":
                    if not messagebox.askyesno(
                        "清除备注",
                        "将清空当前演示文稿全部页的备注。\n\n"
                        "此操作不可撤销（除非 PPT 尚未保存）。是否继续？",
                        icon="warning",
                    ):
                        return
                elif scope == "current":
                    idx = self._last_index if self._last_index and self._last_index > 0 else 0
                    if idx < 1:
                        messagebox.showwarning("清除备注", "未检测到当前页，请先打开演示文稿。")
                        return
                    if not messagebox.askyesno(
                        "清除备注",
                        f"将清空第 {idx} 页的备注。\n\n是否继续？",
                    ):
                        return
                self._send_ppt_cmd({"op": "clear_notes", "scope": scope})

            style = {
                "font": self.title_font,
                "relief": "flat",
                "bd": 0,
                "padx": 14,
                "pady": 6,
                "cursor": "hand2",
                "fg": "#ffffff",
                "activeforeground": "#ffffff",
            }
            tk.Button(
                btns,
                text="当前页",
                command=lambda: choose("current"),
                bg="#6d4c41",
                activebackground="#795548",
                **style,
            ).pack(side="left", padx=(0, 8))
            tk.Button(
                btns,
                text="全部页",
                command=lambda: choose("all"),
                bg="#c62828",
                activebackground="#d32f2f",
                **style,
            ).pack(side="left", padx=(0, 8))
            tk.Button(
                btns,
                text="取消",
                command=win.destroy,
                bg="#3a3a3a",
                activebackground="#505050",
                **style,
            ).pack(side="left")
            win.focus_force()
        finally:
            if was_top and bool(self.pin_var.get()):
                self.root.attributes("-topmost", True)

    def _show_page_times_dialog(self) -> None:
        rows = self.timer.page_times()
        if not rows:
            messagebox.showinfo(
                "本页时长",
                "还没有排练计时数据。\n请先点击「幻灯片放映」开始计时。",
            )
            return

        if self._page_times_win is not None and self._page_times_win.winfo_exists():
            self._refresh_page_times_tree()
            self._page_times_win.lift()
            self._page_times_win.focus_force()
            return

        win = tk.Toplevel(self.root)
        self._page_times_win = win
        win.title("本页累计时长")
        win.configure(bg="#1e1e1e")
        win.attributes("-topmost", True)
        win.geometry("420x600")
        win.minsize(360, 420)

        total = sum(sec for _, sec in rows) or 1.0
        ranked = sorted(rows, key=lambda x: x[1], reverse=True)
        top_pages = {idx for idx, _ in ranked[:3] if ranked[0][1] > 0}

        tk.Label(
            win,
            text=f"总计 {format_hms(self.timer.total_seconds())} · 共 {len(rows)} 页有停留",
            fg="#dddddd",
            bg="#1e1e1e",
            font=self.title_font,
            anchor="w",
        ).pack(fill="x", padx=12, pady=(10, 6))

        cols = ("page", "time", "pct")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=22)
        tree.heading("page", text="页码")
        tree.heading("time", text="累计时长")
        tree.heading("pct", text="占比")
        tree.column("page", width=70, anchor="center")
        tree.column("time", width=120, anchor="center")
        tree.column("pct", width=80, anchor="center")
        tree.pack(fill="both", expand=True, padx=12, pady=4)
        self._page_times_tree = tree
        self._page_times_top = top_pages

        style = ttk.Style(win)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Times.Treeview",
            background="#2a2a2a",
            foreground="#ffffff",
            fieldbackground="#2a2a2a",
            rowheight=max(22, int(24 * self._ui_scale)),
        )
        style.configure("Times.Treeview.Heading", background="#3a3a3a", foreground="#ffffff")
        tree.configure(style="Times.Treeview")
        tree.tag_configure("hot", foreground="#ffb74d")

        for idx, sec in rows:
            pct = f"{sec / total * 100:.1f}%"
            tags = ("hot",) if idx in top_pages else ()
            tree.insert("", "end", values=(f"第 {idx} 页", format_hms(sec), pct), tags=tags)

        btns = tk.Frame(win, bg="#1e1e1e")
        btns.pack(fill="x", padx=12, pady=10)
        tk.Button(
            btns,
            text="导出列表",
            command=self._export_page_times,
            font=self.title_font,
            relief="flat",
            bg="#3a3a3a",
            fg="#ffffff",
            activebackground="#505050",
            cursor="hand2",
            padx=10,
            pady=4,
        ).pack(side="left")
        tk.Button(
            btns,
            text="关闭",
            command=win.destroy,
            font=self.title_font,
            relief="flat",
            bg="#3a3a3a",
            fg="#ffffff",
            activebackground="#505050",
            cursor="hand2",
            padx=10,
            pady=4,
        ).pack(side="right")

        def _on_close() -> None:
            self._page_times_win = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _on_close)

    def _refresh_page_times_tree(self) -> None:
        tree = getattr(self, "_page_times_tree", None)
        if tree is None or not tree.winfo_exists():
            return
        rows = self.timer.page_times()
        total = sum(sec for _, sec in rows) or 1.0
        ranked = sorted(rows, key=lambda x: x[1], reverse=True)
        top_pages = {idx for idx, _ in ranked[:3] if ranked and ranked[0][1] > 0}
        for item in tree.get_children():
            tree.delete(item)
        for idx, sec in rows:
            pct = f"{sec / total * 100:.1f}%"
            tags = ("hot",) if idx in top_pages else ()
            tree.insert("", "end", values=(f"第 {idx} 页", format_hms(sec), pct), tags=tags)

    def _export_page_times(self) -> None:
        rows = self.timer.page_times()
        if not rows:
            messagebox.showinfo("导出列表", "没有可导出的时长数据。")
            return
        was_top = bool(self.pin_var.get())
        if was_top:
            self.root.attributes("-topmost", False)
        try:
            path = filedialog.asksaveasfilename(
                title="导出本页累计时长",
                defaultextension=".csv",
                initialfile="排练本页时长.csv",
                filetypes=[("CSV", "*.csv"), ("文本", "*.txt"), ("所有文件", "*.*")],
            )
            if not path:
                return
            total = sum(sec for _, sec in rows) or 1.0
            lines = ["页码,累计秒,时长,占比"]
            for idx, sec in rows:
                lines.append(
                    f"{idx},{sec:.1f},{format_hms(sec)},{sec / total * 100:.1f}%"
                )
            lines.append(f"合计,{total:.1f},{format_hms(total)},100%")
            Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
            messagebox.showinfo("导出列表", f"已导出到：\n{path}")
        except Exception as exc:
            messagebox.showerror("导出列表", f"导出失败：\n{exc}")
        finally:
            if was_top and bool(self.pin_var.get()):
                self.root.attributes("-topmost", True)

    def _bind_drag(self) -> None:
        def start(event):
            self._drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())

        def move(event):
            x = event.x_root - self._drag_offset[0]
            y = event.y_root - self._drag_offset[1]
            self.root.geometry(f"+{x}+{y}")

        for widget in (self.root, self.status_label):
            widget.bind("<Button-1>", start)
            widget.bind("<B1-Motion>", move)

    def _tick_clock(self) -> None:
        if self._stop_event.is_set():
            return
        self._set_var(self.page_time_var, f"本页 {format_hms(self.timer.page_seconds())}")
        self._set_var(self.total_time_var, f"总计 {format_hms(self.timer.total_seconds())}")
        self._update_countdown_display()
        active = self.timer.running or self._countdown_deadline is not None
        self.root.after(CLOCK_ACTIVE_MS if active else CLOCK_IDLE_MS, self._tick_clock)

    def _post_snap(self, snap: dict) -> None:
        """后台线程调用：只保留最新快照，合并到一次 UI 刷新。"""
        with self._ui_lock:
            self._pending_snap = snap
            need_schedule = not self._flush_scheduled
            if need_schedule:
                self._flush_scheduled = True
        if need_schedule and not self._stop_event.is_set():
            try:
                self.root.after(0, self._flush_snap)
            except Exception:
                pass

    def _flush_snap(self) -> None:
        with self._ui_lock:
            snap = self._pending_snap
            self._pending_snap = None
            self._flush_scheduled = False
        if snap is None or self._stop_event.is_set():
            return
        self._apply_snap(snap)

    def _set_notes_if_changed(self, text: str) -> None:
        if text == self._last_notes:
            return
        self._last_notes = text
        self.notes.configure(state="normal", cursor="arrow")
        self.notes.delete("1.0", "end")
        self.notes.insert("1.0", text, "body")
        self.notes.configure(state="disabled", cursor="arrow")
        self.notes.yview_moveto(0)
        self.notes.see("1.0")

    def _apply_snap(self, snap: dict) -> None:
        showing = bool(snap.get("showing"))
        index = int(snap.get("index") or 0)
        total = int(snap.get("total") or 0)
        notes = str(snap.get("notes") or "")
        force_notes = bool(snap.get("force_notes"))
        show_ended = False

        # 计时只由「幻灯片放映」按钮启动；这里只在放映结束时收尾
        if showing and not self._was_showing:
            self._update_show_button(True)
            if self.timer.running:
                self.state_var.set("排练计时中…")
            else:
                self.state_var.set("放映中（未计时，请点「幻灯片放映」）")
        elif (not showing) and self._was_showing:
            self._update_show_button(False)
            had_timing = self.timer.running or bool(self.timer.page_history)
            if self.timer.running:
                self.timer.on_show_end()
                self._stop_countdown(reset_display=True)
                self.state_var.set("放映已结束（计时已暂停）")
                show_ended = had_timing
            else:
                self.state_var.set("放映已结束")
        self._was_showing = showing

        if self.timer.running and index > 0 and index != self.timer.current_index:
            self.timer.on_slide_change(index)

        if index != self._last_index or total != self._last_total:
            self._last_index = index
            self._last_total = total
            if index > 0 and total > 0:
                self.status_var.set(f"第 {index} / {total} 页")
            elif snap.get("error"):
                self.status_var.set("读取失败")
            else:
                self.status_var.set("未连接 / 无幻灯片")

        if force_notes or notes:
            self._set_notes_if_changed(notes)

        if "has_next" in snap or "next_notes" in snap or force_notes:
            self._show_next_notes(
                index,
                total,
                bool(snap.get("has_next")),
                str(snap.get("next_notes") or ""),
            )
        if show_ended and self.timer.page_history:
            self.root.after(120, self._show_page_times_dialog)

        if not showing and not self.timer.running and self.timer.frozen_total <= 0:
            self.state_var.set("计时未开始（请点「幻灯片放映」）")

    def _ppt_worker_loop(self) -> None:
        pythoncom.CoInitialize()
        app = None
        events = None
        last_index: int | None = None
        last_showing: bool | None = None

        # 事件回调里只投递，不在回调深处做重活以外的事；备注仍当场读以保证实时
        outer = self

        def snap_from_slide(slide, showing: bool) -> dict:
            try:
                idx = int(slide.SlideIndex)
                total = int(slide.Parent.Slides.Count)
            except Exception:
                idx, total = 0, 0
            has_next, next_notes = next_page_notes(slide)
            return {
                "showing": showing,
                "index": idx,
                "total": total,
                "notes": extract_notes(slide),
                "has_next": has_next,
                "next_notes": next_notes,
                "force_notes": True,
                "error": None,
            }

        class AppEvents:
            def OnSlideShowBegin(self, Wn):
                try:
                    outer._post_snap(snap_from_slide(Wn.View.Slide, True))
                except Exception:
                    pass

            def OnSlideShowNextSlide(self, Wn):
                try:
                    outer._post_snap(snap_from_slide(Wn.View.Slide, True))
                except Exception:
                    pass

            def OnSlideShowEnd(self, Pres):
                outer._post_snap(
                    {
                        "showing": False,
                        "index": outer._last_index or 0,
                        "total": outer._last_total or 0,
                        "notes": outer._last_notes,
                        "has_next": outer._last_has_next,
                        "next_notes": outer._last_next_notes,
                        "force_notes": False,
                        "error": None,
                    }
                )

        def connect():
            """只挂接已在运行的 PowerPoint，绝不 Dispatch 新建进程。"""
            nonlocal app, events
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                app = None
                events = None
                return False
            try:
                events = win32com.client.DispatchWithEvents(app, AppEvents)
            except Exception:
                events = None
            return True

        def release():
            nonlocal app, events
            events = None
            app = None

        def run_ppt_command(cmd: dict) -> None:
            """仅在本 COM 线程执行放映命令，避免与 UI 线程交叉调用卡死 PPT。"""
            nonlocal last_index
            op = str(cmd.get("op") or "")
            try:
                if app is None and not connect():
                    raise RuntimeError("未检测到已打开的 PowerPoint，请先打开演示文稿。")
                assert app is not None
                if int(app.Presentations.Count) < 1:
                    raise RuntimeError("当前没有打开的演示文稿。")
                pres = app.ActivePresentation
                info = None
                if op == "activate":
                    if int(app.SlideShowWindows.Count) > 0:
                        app.SlideShowWindows(1).Activate()
                    return
                if op == "nav":
                    if int(app.SlideShowWindows.Count) < 1:
                        return
                    view = app.SlideShowWindows(1).View
                    direction = str(cmd.get("dir") or "")
                    if direction == "next":
                        view.Next()
                    elif direction == "prev":
                        view.Previous()
                    elif direction == "exit":
                        view.Exit()
                    return
                if op == "import_notes":
                    slide_count = int(pres.Slides.Count)
                    written = 0
                    skipped: list[int] = []
                    failed: list[int] = []
                    for page_no, text in cmd.get("pages") or []:
                        if page_no < 1 or page_no > slide_count:
                            skipped.append(page_no)
                        elif set_slide_notes(pres.Slides(page_no), text):
                            written += 1
                        else:
                            failed.append(page_no)
                    # 让主循环重新读取当前页备注
                    last_index = None
                    outer._post_ppt_cmd_result(
                        {
                            "ok": True,
                            "op": op,
                            "written": written,
                            "skipped": skipped,
                            "failed": failed,
                        }
                    )
                    return
                if op == "clear_notes":
                    slide_count = int(pres.Slides.Count)
                    scope = str(cmd.get("scope") or "all")
                    pages: list[int] = []
                    if scope == "current":
                        page_no = 0
                        try:
                            if int(app.SlideShowWindows.Count) > 0:
                                page_no = int(app.SlideShowWindows(1).View.Slide.SlideIndex)
                            else:
                                page_no = int(app.ActiveWindow.View.Slide.SlideIndex)
                        except Exception:
                            page_no = int(outer._last_index or 0)
                        if page_no < 1 or page_no > slide_count:
                            raise RuntimeError("未检测到当前页，请先打开演示文稿并选中一页。")
                        pages = [page_no]
                    else:
                        pages = list(range(1, slide_count + 1))
                    written = 0
                    failed: list[int] = []
                    for page_no in pages:
                        if set_slide_notes(pres.Slides(page_no), ""):
                            written += 1
                        else:
                            failed.append(page_no)
                    last_index = None
                    outer._post_ppt_cmd_result(
                        {
                            "ok": True,
                            "op": op,
                            "written": written,
                            "failed": failed,
                        }
                    )
                    return
                if op == "show":
                    if int(app.SlideShowWindows.Count) < 1:
                        # Run() 可能较慢，但仍须在本线程调用
                        pres.SlideShowSettings.Run()
                    try:
                        if int(app.SlideShowWindows.Count) > 0:
                            app.SlideShowWindows(1).Activate()
                    except Exception:
                        pass
                else:
                    raise RuntimeError(f"未知命令：{op}")
                outer._post_ppt_cmd_result({"ok": True, "op": op, "info": info})
            except Exception as exc:
                outer._post_ppt_cmd_result({"ok": False, "op": op, "error": str(exc)})

        empty_since: float | None = None

        def pump() -> None:
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass

        def idle_wait(seconds: float) -> None:
            """阻塞到超时；期间 COM 事件消息到达就处理，有新命令或退出则立即返回。"""
            deadline = time.monotonic() + seconds
            while not self._stop_event.is_set():
                remain_ms = int((deadline - time.monotonic()) * 1000)
                if remain_ms <= 0:
                    return
                rc = win32event.MsgWaitForMultipleObjects(
                    [self._wake_evt], False, remain_ms, win32event.QS_ALLINPUT
                )
                if rc == win32event.WAIT_OBJECT_0:
                    return
                pump()

        try:
            while not self._stop_event.is_set():
                # 先处理 UI 下发的放映命令
                try:
                    while True:
                        cmd = self._ppt_cmd_q.get_nowait()
                        run_ppt_command(cmd)
                        pump()
                except queue.Empty:
                    pass

                # 处理 COM 事件队列（放映翻页事件）
                pump()

                if app is None:
                    if not connect():
                        self._post_snap(
                            {
                                "showing": False,
                                "index": 0,
                                "total": 0,
                                "notes": (
                                    "正在等待 PowerPoint…\n"
                                    "请先打开 PowerPoint 并打开演示文稿。"
                                ),
                                "force_notes": True,
                                "error": None,
                            }
                        )
                        idle_wait(RECONNECT_SEC)
                        continue
                    empty_since = None
                    last_index = None
                    last_showing = None

                try:
                    showing = False
                    try:
                        showing = app.SlideShowWindows.Count > 0
                    except Exception:
                        release()
                        empty_since = None
                        continue

                    try:
                        pres_count = int(app.Presentations.Count)
                    except Exception:
                        release()
                        empty_since = None
                        continue

                    # 已连接但无演示文稿：短暂等待后释放并重挂，避免粘在错误/空实例上
                    if pres_count < 1 and not showing:
                        now = time.monotonic()
                        if empty_since is None:
                            empty_since = now
                        elif now - empty_since >= EMPTY_REATTACH_SEC:
                            release()
                            empty_since = None
                            last_index = None
                            last_showing = None
                            self._post_snap(
                                {
                                    "showing": False,
                                    "index": 0,
                                    "total": 0,
                                    "notes": (
                                        "未检测到已打开的演示文稿，正在重新连接…\n"
                                        "请确认 PPT 文件已在 PowerPoint 中打开。"
                                    ),
                                    "force_notes": True,
                                    "error": None,
                                }
                            )
                            continue

                        if last_index != 0 or last_showing is not False:
                            last_index = 0
                            last_showing = False
                            self._post_snap(
                                {
                                    "showing": False,
                                    "index": 0,
                                    "total": 0,
                                    "notes": (
                                        "已连接到 PowerPoint，但还没有打开的演示文稿。\n"
                                        "请打开 PPT 文件后再试。"
                                    ),
                                    "force_notes": True,
                                    "error": None,
                                }
                            )
                    else:
                        empty_since = None

                    index = 0
                    total = 0
                    slide = None
                    try:
                        if showing:
                            slide = app.SlideShowWindows(1).View.Slide
                        elif pres_count > 0:
                            try:
                                slide = app.ActiveWindow.View.Slide
                            except Exception:
                                try:
                                    slide = app.ActivePresentation.Slides(1)
                                except Exception:
                                    slide = None
                        if slide is not None:
                            index = int(slide.SlideIndex)
                            total = int(slide.Parent.Slides.Count)
                    except Exception:
                        index, total, slide = 0, 0, None

                    changed = (index != last_index) or (showing != last_showing)
                    if changed:
                        if slide is not None:
                            payload = snap_from_slide(slide, showing)
                        else:
                            payload = {
                                "showing": showing,
                                "index": index,
                                "total": total,
                                "notes": (
                                    "未检测到 PowerPoint 演示文稿。\n"
                                    "请先打开 PPT，再开始放映。"
                                ),
                                "force_notes": True,
                                "error": None,
                            }
                        last_index = index
                        last_showing = showing
                        self._post_snap(payload)
                except Exception as exc:
                    release()
                    empty_since = None
                    last_index = None
                    last_showing = None
                    self._post_snap(
                        {
                            "showing": False,
                            "index": 0,
                            "total": 0,
                            "notes": f"读取出错，正在重连…\n{exc}",
                            "force_notes": True,
                            "error": str(exc),
                        }
                    )

                idle_wait(POLL_SHOW_SEC if last_showing else POLL_IDLE_SEC)
        finally:
            events = None
            app = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _send_ppt_cmd(self, cmd: dict) -> None:
        self._ppt_cmd_q.put(cmd)
        win32event.SetEvent(self._wake_evt)

    def _on_close(self) -> None:
        self._stop_event.set()
        win32event.SetEvent(self._wake_evt)
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()
        self._stop_event.set()
        win32event.SetEvent(self._wake_evt)


def main() -> None:
    NotesOverlayApp().run()


if __name__ == "__main__":
    main()
