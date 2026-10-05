# -*- coding: utf-8 -*-
"""
答辩专用插件 — PowerPoint 浮动备注 + 排练计时

放映/录制时置顶显示当前页备注；支持字号调节、备注导出 Word。
翻页使用 PowerPoint 原生快捷键。
"""

from __future__ import annotations

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
except ImportError:
    print("缺少 pywin32，请先执行: pip install pywin32")
    sys.exit(1)


PP_PLACEHOLDER_BODY = 2
CLOCK_MS = 100
# 轻量页码轮询（兜底）；真正翻页靠事件 + 立刻读备注
INDEX_POLL_SEC = 0.05


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
    return "（本页无备注）"


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
                self.page_history[self.current_index] = self.frozen_page
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


class NotesOverlayApp:
    def __init__(self) -> None:
        self.timer = RehearsalTimer()
        self._was_showing = False
        self._last_index: int | None = None
        self._last_total: int | None = None
        self._last_notes = ""
        self._stop_event = threading.Event()

        # UI 合并：只应用最新一帧，避免 after 队列堆积造成卡顿
        self._pending_snap: dict | None = None
        self._flush_scheduled = False
        self._ui_lock = threading.Lock()
        self._font_size = 16
        self._font_min = 12
        self._font_max = 36
        self._ui_scale = 1.0

        enable_dpi_awareness()
        self.root = tk.Tk()
        self.root.title("答辩专用插件")
        self._setup_dpi_and_geometry()
        self.root.resizable(True, True)
        self.root.attributes("-topmost", True)
        try:
            self.root.attributes("-alpha", 0.92)
        except Exception:
            pass
        self.root.configure(bg="#1e1e1e")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_window_icon()

        self._drag_offset = (0, 0)
        self._build_ui()
        self._bind_drag()

        self.root.after(CLOCK_MS, self._tick_clock)
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
        target_area = (sw * sh) / 4.0
        # 16:5 => w=16k, h=5k, area=80 k^2
        k = (target_area / 80.0) ** 0.5
        w = max(640, int(16 * k))
        h = max(200, int(5 * k))
        w = min(w, max(640, sw - 40))
        h = min(h, max(200, sh - 80))
        # 再次贴合 16:5
        h = max(200, int(w * 5 / 16))
        if h > sh - 80:
            h = sh - 80
            w = max(640, int(h * 16 / 5))
        x = max(0, (sw - w) // 2)
        y = max(20, (sh - h) // 5)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.minsize(560, 180)

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

        header_h = max(48, int(self.title_font.metrics("linespace") * 2.4 + 12))
        header = tk.Frame(self.root, bg="#2b2b2b", height=header_h)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        left = tk.Frame(header, bg="#2b2b2b")
        left.pack(side="left", padx=int(12 * s), pady=int(6 * s))

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

        self.state_var = tk.StringVar(value="计时未开始（开始放映后自动计时）")
        tk.Label(
            left,
            textvariable=self.state_var,
            fg="#c8c8c8",
            bg="#2b2b2b",
            font=self.title_font,
            anchor="w",
        ).pack(anchor="w")

        self.pin_var = tk.BooleanVar(value=True)
        pin_btn = tk.Checkbutton(
            header,
            text="置顶",
            variable=self.pin_var,
            command=self._toggle_topmost,
            fg="#eeeeee",
            bg="#2b2b2b",
            selectcolor="#2b2b2b",
            activebackground="#2b2b2b",
            activeforeground="#ffffff",
            font=self.title_font,
        )
        pin_btn.pack(side="right", padx=(0, int(10 * s)), pady=int(6 * s))

        zoom = tk.Frame(header, bg="#2b2b2b")
        zoom.pack(side="right", padx=(0, int(6 * s)), pady=int(6 * s))
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
        tk.Button(zoom, text="导出备注", command=self._export_notes_txt, **zoom_style).pack(
            side="left", padx=(0, int(6 * s))
        )
        tk.Button(zoom, text="A−", command=lambda: self._change_font(-2), **zoom_style).pack(
            side="left", padx=(0, int(4 * s))
        )
        tk.Button(zoom, text="A+", command=lambda: self._change_font(+2), **zoom_style).pack(side="left")

        timers = tk.Frame(header, bg="#2b2b2b")
        timers.pack(side="right", padx=int(12 * s), pady=int(4 * s))

        self.page_time_var = tk.StringVar(value="本页 00:00")
        self.total_time_var = tk.StringVar(value="总计 00:00")

        tk.Label(
            timers,
            textvariable=self.page_time_var,
            fg="#7fd99a",
            bg="#2b2b2b",
            font=self.timer_font,
        ).pack(side="left", padx=(0, int(18 * s)))
        tk.Label(
            timers,
            textvariable=self.total_time_var,
            fg="#7ec8ff",
            bg="#2b2b2b",
            font=self.timer_font,
        ).pack(side="left")

        body = tk.Frame(self.root, bg="#1e1e1e")
        body.pack(fill="both", expand=True, padx=10, pady=(8, 4))

        # 先放右侧滚动条（始终显示，避免“显示不全也没有条”）
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
            insertbackground="#ffffff",
            relief="flat",
            padx=12,
            pady=10,
            borderwidth=0,
            highlightthickness=0,
            yscrollcommand=self._notes_scroll.set,
        )
        self.notes.pack(side="left", fill="both", expand=True)
        self.notes.insert("1.0", "正在连接 PowerPoint…")
        # 用 normal + 禁编辑，保证滚动条/滚轮在 Windows 上可靠
        self.notes.bind("<Key>", lambda _e: "break")
        self.notes.bind("<<Paste>>", lambda _e: "break")
        self.notes.bind("<Button-1>", lambda _e: self.notes.focus_set())
        self.notes.bind("<MouseWheel>", self._on_notes_wheel)
        body.bind_all("<MouseWheel>", self._on_notes_wheel_global)

        # 底部：提示 + 明显的拖拽改大小图标
        grip = tk.Frame(self.root, bg="#1e1e1e")
        grip.pack(fill="x", side="bottom")
        tk.Label(
            grip,
            text="Ctrl+滚轮 调字号",
            fg="#888888",
            bg="#1e1e1e",
            font=self.title_font,
            anchor="w",
        ).pack(side="left", padx=10, pady=2)

        resize_box = tk.Frame(grip, bg="#2b2b2b", cursor="size_nw_se")
        resize_box.pack(side="right", padx=6, pady=4)
        self._resize_icon = tk.Label(
            resize_box,
            text=" ↘ ",
            fg="#dddddd",
            bg="#2b2b2b",
            font=tkfont.Font(family="Segoe UI Symbol", size=14, weight="bold"),
            cursor="size_nw_se",
            padx=8,
            pady=2,
        )
        self._resize_icon.pack()
        tip = tk.Label(
            resize_box,
            text="拖拽改大小",
            fg="#aaaaaa",
            bg="#2b2b2b",
            font=self.title_font,
            cursor="size_nw_se",
        )
        tip.pack(pady=(0, 2))
        for w in (resize_box, self._resize_icon, tip):
            w.bind("<ButtonPress-1>", self._start_resize)
            w.bind("<B1-Motion>", self._do_resize)

        self.root.bind("<Control-plus>", lambda _e: self._change_font(+2))
        self.root.bind("<Control-equal>", lambda _e: self._change_font(+2))
        self.root.bind("<Control-minus>", lambda _e: self._change_font(-2))
        self.root.bind("<Control-MouseWheel>", self._on_ctrl_wheel)
        self.notes.bind("<Control-MouseWheel>", self._on_ctrl_wheel)

    def _notes_yview(self, *args) -> None:
        self.notes.yview(*args)

    def _on_notes_wheel(self, event) -> str:
        self.notes.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _on_notes_wheel_global(self, event) -> str | None:
        # 仅当鼠标在备注区域上时滚动
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
        w = max(560, w0 + (event.x_root - x0))
        h = max(180, h0 + (event.y_root - y0))
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

            def set_run_font(run, size_pt: float, bold: bool = False) -> None:
                run.bold = bold
                run.font.size = Pt(size_pt)
                run.font.name = "宋体"
                r = run._element
                r.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

            def add_text_run(paragraph, text: str, size_pt: float, bold: bool = False) -> None:
                clean = sanitize_xml_text(text)
                # 空串也允许，避免控制字符清洗后变空导致异常
                run = paragraph.add_run(clean)
                set_run_font(run, size_pt, bold=bold)

            # 文档标题：一级标题加粗
            p_title = doc.add_heading("", level=1)
            add_text_run(p_title, stem, 16, bold=True)
            p_title.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            # 副信息正文
            p_meta = doc.add_paragraph()
            add_text_run(
                p_meta,
                f"演示文稿：{title}　　总页数：{pres.Slides.Count}",
                11,
                bold=False,
            )

            # 标题与正文之间空一行
            doc.add_paragraph()

            total = int(pres.Slides.Count)
            for i in range(1, total + 1):
                slide = pres.Slides(i)
                notes = sanitize_xml_text(extract_notes(slide))

                # 页码：一级标题加粗
                p_h = doc.add_heading("", level=1)
                add_text_run(p_h, f"第 {i} 页", 14, bold=True)

                # 备注正文；多段用换行分割，不加横线
                body = notes.strip()
                if not body or body == "（本页无备注）":
                    p = doc.add_paragraph()
                    add_text_run(p, "（本页无备注）", 12, bold=False)
                else:
                    for para_text in body.split("\n"):
                        p = doc.add_paragraph()
                        add_text_run(p, para_text.rstrip(), 12, bold=False)

                # 页与页之间空一行分隔
                if i < total:
                    doc.add_paragraph()

            doc.save(path)
            messagebox.showinfo("导出备注", f"已导出 {total} 页备注到：\n{path}")
        except Exception as exc:
            messagebox.showerror("导出备注", f"导出失败：\n{exc}")
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
        self.page_time_var.set(f"本页 {format_hms(self.timer.page_seconds())}")
        self.total_time_var.set(f"总计 {format_hms(self.timer.total_seconds())}")
        self.root.after(CLOCK_MS, self._tick_clock)

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
        self.notes.configure(state="normal")
        self.notes.delete("1.0", "end")
        self.notes.insert("1.0", text)
        self.notes.yview_moveto(0)
        self.notes.see("1.0")

    def _apply_snap(self, snap: dict) -> None:
        showing = bool(snap.get("showing"))
        index = int(snap.get("index") or 0)
        total = int(snap.get("total") or 0)
        notes = str(snap.get("notes") or "")
        force_notes = bool(snap.get("force_notes"))

        if showing and not self._was_showing:
            self.timer.on_show_start(index if index > 0 else 1)
            self.state_var.set("排练计时中…")
        elif (not showing) and self._was_showing:
            self.timer.on_show_end()
            self.state_var.set("放映已结束（计时已暂停）")
        self._was_showing = showing

        if showing and index > 0 and index != self.timer.current_index:
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

        if not showing and not self.timer.running and self.timer.frozen_total <= 0:
            self.state_var.set("计时未开始（开始放映后自动计时）")

    def _ppt_worker_loop(self) -> None:
        pythoncom.CoInitialize()
        app = None
        events = None
        last_index: int | None = None
        last_showing: bool | None = None

        # 事件回调里只投递，不在回调深处做重活以外的事；备注仍当场读以保证实时
        outer = self

        class AppEvents:
            def OnSlideShowBegin(self, Wn):
                try:
                    slide = Wn.View.Slide
                    idx = int(slide.SlideIndex)
                    total = int(slide.Parent.Slides.Count)
                    notes = extract_notes(slide)
                    outer._post_snap(
                        {
                            "showing": True,
                            "index": idx,
                            "total": total,
                            "notes": notes,
                            "force_notes": True,
                            "error": None,
                        }
                    )
                except Exception:
                    pass

            def OnSlideShowNextSlide(self, Wn):
                try:
                    slide = Wn.View.Slide
                    idx = int(slide.SlideIndex)
                    total = int(slide.Parent.Slides.Count)
                    notes = extract_notes(slide)
                    outer._post_snap(
                        {
                            "showing": True,
                            "index": idx,
                            "total": total,
                            "notes": notes,
                            "force_notes": True,
                            "error": None,
                        }
                    )
                except Exception:
                    pass

            def OnSlideShowEnd(self, Pres):
                outer._post_snap(
                    {
                        "showing": False,
                        "index": outer._last_index or 0,
                        "total": outer._last_total or 0,
                        "notes": outer._last_notes,
                        "force_notes": False,
                        "error": None,
                    }
                )

        def connect():
            nonlocal app, events
            try:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            except Exception:
                try:
                    app = win32com.client.Dispatch("PowerPoint.Application")
                    app.Visible = True
                except Exception:
                    app = None
                    events = None
                    return False
            try:
                events = win32com.client.DispatchWithEvents(app, AppEvents)
            except Exception:
                events = None
            return True

        try:
            while not self._stop_event.is_set():
                # 处理 COM 事件队列（放映翻页事件）
                try:
                    pythoncom.PumpWaitingMessages()
                except Exception:
                    pass

                if app is None:
                    connect()
                    if app is None:
                        self._post_snap(
                            {
                                "showing": False,
                                "index": 0,
                                "total": 0,
                                "notes": "正在连接 PowerPoint…",
                                "force_notes": True,
                                "error": None,
                            }
                        )
                        self._stop_event.wait(0.2)
                        continue

                try:
                    # 先轻量看页码/是否放映；变化时立刻读备注
                    showing = False
                    try:
                        showing = app.SlideShowWindows.Count > 0
                    except Exception:
                        # 连接失效，重连
                        app = None
                        events = None
                        continue

                    index = 0
                    total = 0
                    slide = None
                    try:
                        if showing:
                            slide = app.SlideShowWindows(1).View.Slide
                        elif app.Presentations.Count > 0:
                            try:
                                slide = app.ActiveWindow.View.Slide
                            except Exception:
                                slide = None
                        if slide is not None:
                            index = int(slide.SlideIndex)
                            total = int(slide.Parent.Slides.Count)
                    except Exception:
                        index, total, slide = 0, 0, None

                    changed = (index != last_index) or (showing != last_showing)
                    if changed:
                        notes = extract_notes(slide) if slide is not None else (
                            "未检测到 PowerPoint 演示文稿。\n请先打开 PPT，再开始放映或录制。"
                        )
                        last_index = index
                        last_showing = showing
                        self._post_snap(
                            {
                                "showing": showing,
                                "index": index,
                                "total": total,
                                "notes": notes,
                                "force_notes": True,
                                "error": None,
                            }
                        )
                except Exception as exc:
                    app = None
                    events = None
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

                # 短睡并泵消息，兼顾事件实时性
                t_end = time.monotonic() + INDEX_POLL_SEC
                while not self._stop_event.is_set() and time.monotonic() < t_end:
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:
                        pass
                    time.sleep(0.01)
        finally:
            events = None
            app = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _on_close(self) -> None:
        self._stop_event.set()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()
        self._stop_event.set()


def main() -> None:
    NotesOverlayApp().run()


if __name__ == "__main__":
    main()
