"""Окно: текстовая выгрузка SQL Server → книга Excel."""

from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from alfa_excel.convert import (
    Cancelled,
    ExportError,
    Layout,
    Progress,
    Settings,
    convert,
    describe_delimiter,
    inspect,
    plan_folder,
    unescape_terminator,
)


BG = "#12141c"
SURFACE = "#1b1f2a"
INPUT = "#10131a"
BORDER = "#323748"
TEXT = "#f4f0e8"
MUTED = "#9b958b"
ACCENT = "#e2a56a"
ACCENT_HOVER = "#f0c08a"
DANGER = "#8f4540"
OK = "#8fbf9f"

STAGES = (
    ("probe", "Анализ", "Весь файл", "#a9b4d0"),
    ("read", "Чтение", "Байты с диска", "#7eb8da"),
    ("parse", "Разбор", "Поля строки", "#d6c07a"),
    ("form", "Формирование", "Типы ячеек", "#e2a56a"),
    ("write", "Запись", "Книга Excel", "#8fbf9f"),
)

ENCODINGS = (
    ("auto", "Авто"),
    ("utf-8", "UTF-8"),
    ("utf-8-sig", "UTF-8 с BOM"),
    ("utf-16", "UTF-16"),
    ("cp1251", "Windows-1251"),
    ("cp866", "OEM 866"),
    ("cp1252", "Windows-1252"),
)
MODES = (
    ("auto", "Авто"),
    ("delimited", "Текст bcp — разделитель"),
    ("report", "Отчёт sqlcmd / SSMS"),
    ("csv", "CSV с кавычками"),
    ("fmt", "Форматный файл bcp"),
)
DELIMITERS = (
    ("auto", "Авто"),
    ("\t", "Табуляция"),
    (",", "Запятая"),
    (";", "Точка с запятой"),
    ("|", "Вертикальная черта"),
    ("custom", "Свой"),
)
QUALIFIERS = (
    ("", "Нет"),
    ('"', 'Кавычки "'),
    ("'", "Апостроф '"),
)
HEADERS = (
    ("auto", "Авто"),
    ("yes", "Первая строка — имена"),
    ("no", "Без имён"),
)


def _prepare_tcl() -> None:
    """venv не видит каталоги Tcl рядом с базовым Python, пока их не указать."""
    if os.environ.get("TCL_LIBRARY") and os.environ.get("TK_LIBRARY"):
        return
    root = os.path.join(sys.base_prefix, "tcl")
    tcl = os.path.join(root, "tcl8.6")
    tk_dir = os.path.join(root, "tk8.6")
    if os.path.isdir(tcl):
        os.environ.setdefault("TCL_LIBRARY", tcl)
    if os.path.isdir(tk_dir):
        os.environ.setdefault("TK_LIBRARY", tk_dir)


def _enable_dpi() -> None:
    """Разный масштаб у мониторов: окно должно быть чётким на том экране, где его открыли."""
    try:
        import ctypes

        awareness = ctypes.windll.user32.SetProcessDpiAwarenessContext
        awareness.argtypes = [ctypes.c_void_p]
        awareness.restype = ctypes.c_bool
        if awareness(ctypes.c_void_p(-4)):
            return
    except Exception:
        pass
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        return


def _monitor_under_cursor() -> tuple[int, int, int, int, int]:
    """Рабочая область и DPI монитора, на котором сейчас курсор."""
    import ctypes

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    class MONITORINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_ulong),
            ("rcMonitor", RECT),
            ("rcWork", RECT),
            ("dwFlags", ctypes.c_ulong),
        ]

    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    user32 = ctypes.windll.user32
    point = POINT()
    user32.GetCursorPos(ctypes.byref(point))
    monitor = user32.MonitorFromPoint(point, 2)
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(MONITORINFO)
    if not user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
        raise OSError("monitor")
    dpi_x = ctypes.c_uint()
    dpi_y = ctypes.c_uint()
    ctypes.windll.shcore.GetDpiForMonitor(monitor, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y))
    work = info.rcWork
    return work.left, work.top, work.right, work.bottom, int(dpi_x.value or 96)


def _monitor_of_window(hwnd: int) -> tuple[int, int, int, int, int]:
    import ctypes

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    class MONITORINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_ulong),
            ("rcMonitor", RECT),
            ("rcWork", RECT),
            ("dwFlags", ctypes.c_ulong),
        ]

    user32 = ctypes.windll.user32
    parent = user32.GetParent(hwnd) or hwnd
    monitor = user32.MonitorFromWindow(parent, 2)
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(MONITORINFO)
    if not user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
        raise OSError("monitor")
    dpi_x = ctypes.c_uint()
    dpi_y = ctypes.c_uint()
    ctypes.windll.shcore.GetDpiForMonitor(monitor, 0, ctypes.byref(dpi_x), ctypes.byref(dpi_y))
    work = info.rcWork
    return work.left, work.top, work.right, work.bottom, int(dpi_x.value or 96)


def _format_int(value: int) -> str:
    return f"{value:,}".replace(",", " ")


def _format_bytes(value: int) -> str:
    amount = float(value)
    for unit in ("Б", "КБ", "МБ", "ГБ", "ТБ"):
        if amount < 1024 or unit == "ТБ":
            if unit == "Б":
                return f"{int(amount)} {unit}"
            return f"{amount:.1f} {unit}"
        amount /= 1024
    return f"{value} Б"


def _with_file(event: Progress, index: int, total: int, name: str) -> Progress:
    event.message = f"{index}/{total} {name}"
    return event


class _FolderDone:
    def __init__(
        self,
        rows: int,
        files: list[str],
        elapsed: float,
        truncated_cells: int,
        skipped_long: int,
        ragged_rows: int,
        columns: int,
        failed: list[str],
        sources: int,
    ) -> None:
        self.rows = rows
        self.files = files
        self.elapsed = elapsed
        self.truncated_cells = truncated_cells
        self.skipped_long = skipped_long
        self.ragged_rows = ragged_rows
        self.columns = columns
        self.failed = failed
        self.sources = sources


class DarkButton(tk.Button):
    def __init__(self, parent: tk.Misc, text: str, command, kind: str = "ghost") -> None:
        palette = {
            "accent": (ACCENT, "#1a120c", ACCENT_HOVER, "#c48b52"),
            "danger": (DANGER, "#f8e4e1", "#a85c56", "#6e3430"),
            "ghost": ("#2c3142", TEXT, "#3c4258", "#232838"),
        }
        bg, fg, hover, press = palette[kind]
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=press,
            activeforeground=fg,
            disabledforeground="#6f6a62",
            relief="flat",
            bd=0,
            padx=16,
            pady=8,
            font=("Segoe UI Semibold", 10),
            cursor="hand2",
            highlightthickness=0,
        )
        self._bg = bg
        self._fg = fg
        self._hover = hover
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event) -> None:
        if str(self.cget("state")) != "disabled":
            self.configure(bg=self._hover)

    def _on_leave(self, _event) -> None:
        if str(self.cget("state")) != "disabled":
            self.configure(bg=self._bg)

    def set_enabled(self, enabled: bool) -> None:
        self.configure(state="normal" if enabled else "disabled", bg=self._bg if enabled else "#2a2d36")


class WrapBar(tk.Frame):
    """Кнопки в ряд, а если не влезают — переносятся на следующую строку."""

    def __init__(self, parent: tk.Misc, bg: str) -> None:
        super().__init__(parent, bg=bg)
        self._widgets: list[tk.Widget] = []
        self.bind("<Configure>", self._relayout)

        self._last_width = 0

    def add(self, widget: tk.Widget) -> None:
        self._widgets.append(widget)
        self._relayout()

    def _relayout(self, event: tk.Event | None = None) -> None:
        width = event.width if event is not None and event.width > 1 else self.winfo_width()
        if width < 40 or abs(width - self._last_width) < 2:
            return
        self._last_width = width
        for widget in self._widgets:
            widget.grid_forget()
        row = 0
        column = 0
        used = 0
        for widget in self._widgets:
            needed = widget.winfo_reqwidth() + 8
            if column and used + needed > width:
                row += 1
                column = 0
                used = 0
            widget.grid(row=row, column=column, sticky="w", padx=(0, 8), pady=3)
            used += needed
            column += 1


class StageBar(tk.Frame):
    def __init__(self, parent: tk.Misc, title: str, caption: str, color: str) -> None:
        super().__init__(parent, bg=SURFACE)
        self.color = color
        self.fraction = 0.0
        title_row = tk.Frame(self, bg=SURFACE)
        title_row.pack(fill="x")
        tk.Label(title_row, text=title, bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 10)).pack(side="left")
        self.percent = tk.Label(title_row, text="0%", bg=SURFACE, fg=MUTED, font=("Segoe UI", 9))
        self.percent.pack(side="right")
        self.caption = tk.Label(
            self,
            text=caption,
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=240,
        )
        self.caption.pack(fill="x", pady=(2, 0))
        self.canvas = tk.Canvas(self, height=8, bg=SURFACE, highlightthickness=0, bd=0)
        self.canvas.pack(fill="x", pady=(6, 0))
        self.bind("<Configure>", self._fit_caption)
        self.canvas.bind("<Configure>", lambda _event: self.redraw())

    def _fit_caption(self, event: tk.Event) -> None:
        if event.width > 20:
            self.caption.configure(wraplength=event.width)

    def set(self, fraction: float, caption: str) -> None:
        self.fraction = max(0.0, min(1.0, fraction))
        self.caption.configure(text=caption)
        self.percent.configure(text=f"{round(self.fraction * 100)}%")
        self.redraw()

    def redraw(self) -> None:
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        if width < 4:
            return
        self.canvas.delete("all")
        self._round(1, 1, width - 1, height - 1, "#232838")
        fill = max(0, int((width - 2) * self.fraction))
        if fill > 0:
            self._round(1, 1, 1 + max(fill, height), height - 1, self.color)

    def _round(self, x1: int, y1: int, x2: int, y2: int, color: str) -> None:
        radius = max(1, (y2 - y1) // 2)
        self.canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2, fill=color, outline=color)
        self.canvas.create_oval(x1, y1, x1 + radius * 2, y2, fill=color, outline=color)
        self.canvas.create_oval(x2 - radius * 2, y1, x2, y2, fill=color, outline=color)


class App:
    def __init__(self) -> None:
        _prepare_tcl()
        _enable_dpi()
        self.root = tk.Tk()
        self.root.title("ALFA — TXT в Excel")
        self.root.configure(bg=BG)
        self.queue: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self.worker: threading.Thread | None = None
        self.closing = False
        self._screen_key: tuple[int, int, int] | None = None
        self._dpi = 96
        self._fit_job: str | None = None
        self._stacked: bool | None = None
        self._hints: list[tk.Label] = []
        self._apply_screen_scale()
        self._build_styles()
        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Configure>", self._schedule_screen_sync)
        self.root.bind_all("<MouseWheel>", self._on_wheel)
        self.root.after(30, lambda: self._place_on_screen(follow_cursor=True))
        self.root.after(80, self._poll)

    def _apply_screen_scale(self) -> None:
        try:
            _left, _top, _right, _bottom, dpi = _monitor_under_cursor()
        except Exception:
            dpi = 96
        self._dpi = max(96, dpi)
        self.root.tk.call("tk", "scaling", self._dpi / 72.0)

    def _place_on_screen(self, follow_cursor: bool = False) -> None:
        self.root.update_idletasks()
        try:
            if follow_cursor:
                left, top, right, bottom, dpi = _monitor_under_cursor()
            else:
                left, top, right, bottom, dpi = _monitor_of_window(self.root.winfo_id())
        except Exception:
            left, top = 0, 0
            right, bottom = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            dpi = self._dpi
        self._screen_key = (left, top, max(96, dpi))
        work_w = max(320, right - left)
        work_h = max(240, bottom - top)
        width = work_w - 24
        height = work_h - 32
        if work_w >= 1500:
            width = min(width, int(work_w * 0.72))
        if work_h >= 980:
            height = min(height, int(work_h * 0.8))
        width = max(320, min(width, work_w - 8))
        height = max(240, min(height, work_h - 8))
        x = left + max(0, (work_w - width) // 2)
        y = top + max(0, (work_h - height) // 2)
        self.root.minsize(min(520, width), min(360, height))
        if self.root.state() != "zoomed":
            self.root.geometry(f"{int(width)}x{int(height)}+{int(x)}+{int(y)}")
        self._screen_key = (left, top, self._dpi)
        self._sync_scroll()

    def _adopt_dpi(self, dpi: int, work_h: int) -> None:
        dpi = max(96, dpi)
        if dpi != self._dpi:
            self._dpi = dpi
            self.root.tk.call("tk", "scaling", self._dpi / 72.0)
        self._fit_density(work_h)

    def _fit_density(self, work_h: int) -> None:
        compact = work_h < 860
        if hasattr(self, "title_label"):
            self.title_label.configure(font=("Segoe UI Semibold", 16 if compact else 22))
        if hasattr(self, "log"):
            self.log.configure(height=4 if compact else 6)
        if hasattr(self, "tree"):
            self.tree.configure(height=5 if compact else 7)

    def _schedule_screen_sync(self, event: tk.Event) -> None:
        if event.widget is not self.root or self.closing:
            return
        if self._fit_job is not None:
            self.root.after_cancel(self._fit_job)
        self._fit_job = self.root.after(200, self._sync_screen)

    def _sync_screen(self) -> None:
        self._fit_job = None
        if self.closing:
            return
        try:
            left, top, right, bottom, dpi = _monitor_of_window(self.root.winfo_id())
        except Exception:
            return
        key = (left, top, dpi)
        if key == self._screen_key:
            self._sync_scroll()
            return
        self._screen_key = key
        if self.root.state() == "zoomed":
            self._sync_scroll()
            return
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        origin_x = self.root.winfo_x()
        origin_y = self.root.winfo_y()
        fitted_w = min(width, max(320, right - left - 8))
        fitted_h = min(height, max(240, bottom - top - 8))
        fitted_x = min(max(origin_x, left), right - fitted_w)
        fitted_y = min(max(origin_y, top), bottom - fitted_h)
        self.root.minsize(min(520, fitted_w), min(360, fitted_h))
        if (fitted_w, fitted_h, fitted_x, fitted_y) != (width, height, origin_x, origin_y):
            self.root.geometry(f"{int(fitted_w)}x{int(fitted_h)}+{int(fitted_x)}+{int(fitted_y)}")
        self._sync_scroll()

    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "Dark.TCombobox",
            fieldbackground=INPUT,
            background=SURFACE,
            foreground=TEXT,
            arrowcolor=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=6,
        )
        style.map(
            "Dark.TCombobox",
            fieldbackground=[("readonly", INPUT), ("disabled", "#1a1e28")],
            foreground=[("readonly", TEXT), ("disabled", MUTED)],
            selectbackground=[("readonly", INPUT)],
            selectforeground=[("readonly", TEXT)],
        )
        style.configure(
            "Preview.Treeview",
            background=INPUT,
            fieldbackground=INPUT,
            foreground=TEXT,
            borderwidth=0,
            rowheight=24,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Preview.Treeview.Heading",
            background="#2a3142",
            foreground=TEXT,
            relief="flat",
            font=("Segoe UI Semibold", 9),
        )
        style.map("Preview.Treeview", background=[("selected", "#3d3428")], foreground=[("selected", TEXT)])
        style.configure(
            "Dark.Vertical.TScrollbar",
            background="#2c3142",
            troughcolor=INPUT,
            bordercolor=INPUT,
            arrowcolor=TEXT,
        )
        self.root.option_add("*TCombobox*Listbox.background", INPUT)
        self.root.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#3d3428")
        self.root.option_add("*TCombobox*Listbox.selectForeground", TEXT)
        self.root.option_add("*TCombobox*Listbox.font", "Segoe UI 10")

    def _build(self) -> None:
        header = tk.Frame(self.root, bg=BG)
        header.pack(fill="x", padx=18, pady=(14, 6))
        tk.Label(header, text="ALFA", bg=BG, fg=ACCENT, font=("Segoe UI Semibold", 12)).pack(anchor="w")
        self.title_label = tk.Label(
            header,
            text="Текстовая выгрузка в Excel",
            bg=BG,
            fg=TEXT,
            font=("Segoe UI Semibold", 18),
            justify="left",
            anchor="w",
            wraplength=280,
        )
        self.title_label.pack(anchor="w", pady=(2, 0))
        subtitle = tk.Label(
            header,
            text="Файл или целая папка выгрузок. Книги из папки сохраняются рядом с исходными файлами.",
            bg=BG,
            fg=MUTED,
            font=("Segoe UI", 10),
            justify="left",
            anchor="w",
        )
        subtitle.pack(anchor="w", fill="x", pady=(4, 0))
        self._hints.append(self.title_label)
        self._hints.append(subtitle)

        middle = tk.Frame(self.root, bg=BG)
        middle.pack(fill="both", expand=True)
        self.vscroll = tk.Scrollbar(
            middle,
            orient="vertical",
            command=lambda *args: self.viewport.yview(*args),
            bg="#2c3142",
            troughcolor=BG,
            activebackground="#3c4258",
            highlightthickness=0,
            bd=0,
        )
        self.viewport = tk.Canvas(middle, bg=BG, highlightthickness=0, bd=0, yscrollcommand=self._on_scroll)
        self.viewport.pack(side="left", fill="both", expand=True)
        self.body = tk.Frame(self.viewport, bg=BG)
        self.viewport_window = self.viewport.create_window((0, 0), window=self.body, anchor="nw")
        self.body.columnconfigure(0, weight=5)
        self.body.columnconfigure(1, weight=6)
        self.body.rowconfigure(1, weight=1)
        self.body.bind("<Configure>", lambda _event: self._sync_scroll())
        self.viewport.bind("<Configure>", self._on_viewport_configure)
        self.viewport.bind("<MouseWheel>", self._on_wheel)

        self.source_card = self._card(self.body)
        self.preview_card = self._card(self.body)
        self.pipeline_card = self._card(self.body)
        self._build_source(self.source_card)
        self._build_preview(self.preview_card)
        self._build_pipeline(self.pipeline_card)
        self._reflow(1100)
        self._build_actions()

    def _on_viewport_configure(self, event: tk.Event) -> None:
        self.viewport.itemconfigure(self.viewport_window, width=event.width)
        self._reflow(event.width)
        self._sync_scroll()

    def _on_scroll(self, first: str, last: str) -> None:
        self.vscroll.set(first, last)
        if float(first) <= 0 and float(last) >= 1:
            if self.vscroll.winfo_ismapped():
                self.vscroll.pack_forget()
        elif not self.vscroll.winfo_ismapped():
            self.vscroll.pack(side="right", fill="y")

    def _sync_scroll(self) -> None:
        bbox = self.viewport.bbox("all")
        if bbox:
            self.viewport.configure(scrollregion=bbox)

    def _on_wheel(self, event: tk.Event) -> None:
        widget = event.widget
        while widget is not None:
            if widget in (getattr(self, "log", None), getattr(self, "tree", None)):
                return
            widget = getattr(widget, "master", None)
        if self.vscroll.winfo_ismapped():
            self.viewport.yview_scroll(int(-event.delta / 120), "units")

    def _reflow(self, width: int) -> None:
        scale = float(self.root.tk.call("tk", "scaling") or 1.3)
        card_min = int(360 * scale)
        stacked = width < card_min * 2 + 48
        if stacked == self._stacked:
            self._wrap_hints(width)
            return
        self._stacked = stacked
        pad = 12
        self.body.columnconfigure(0, weight=1, uniform="cards", minsize=0)
        self.body.columnconfigure(1, weight=1, uniform="cards", minsize=0)
        if stacked:
            self.body.columnconfigure(1, weight=0, uniform="")
            self.source_card.grid(row=0, column=0, columnspan=2, sticky="ew", padx=pad, pady=(8, 8))
            self.preview_card.grid(row=1, column=0, columnspan=2, sticky="ew", padx=pad, pady=8)
            self.pipeline_card.grid(row=2, column=0, columnspan=2, sticky="ew", padx=pad, pady=(8, 8))
        else:
            self.source_card.grid(row=0, column=0, columnspan=1, sticky="nsew", padx=(pad, 6), pady=(8, 8))
            self.preview_card.grid(row=0, column=1, columnspan=1, sticky="nsew", padx=(6, pad), pady=(8, 8))
            self.pipeline_card.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=pad, pady=(8, 8))
        self._wrap_hints(width)

    def _wrap_hints(self, width: int = 0) -> None:
        limit = max(160, (width or self.viewport.winfo_width()) - 48)
        if self._stacked:
            each = limit
        else:
            each = max(160, (limit - 24) // 2)
        for label in self._hints:
            parent_width = label.master.winfo_width()
            label.configure(wraplength=max(140, parent_width - 8 if parent_width > 48 else each))

    def _card(self, parent: tk.Misc) -> tk.Frame:
        outer = tk.Frame(parent, bg=BORDER)
        inner = tk.Frame(outer, bg=SURFACE)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        outer.inner = inner  # type: ignore[attr-defined]
        return outer

    def _label(self, parent: tk.Misc, text: str) -> None:
        tk.Label(parent, text=text, bg=SURFACE, fg=MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 4))

    def _entry(self, parent: tk.Misc, variable: tk.StringVar) -> tk.Entry:
        return tk.Entry(
            parent,
            textvariable=variable,
            bg=INPUT,
            fg=TEXT,
            insertbackground=TEXT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
            font=("Segoe UI", 10),
            width=12,
            bd=6,
        )

    def _combo(self, parent: tk.Misc, variable: tk.StringVar, pairs: tuple[tuple[str, str], ...]) -> ttk.Combobox:
        labels = [label for _key, label in pairs]
        self._combo_maps = getattr(self, "_combo_maps", {})
        mapping = {label: key for key, label in pairs}
        reverse = {key: label for key, label in pairs}
        widget = ttk.Combobox(
            parent,
            textvariable=variable,
            values=labels,
            state="readonly",
            style="Dark.TCombobox",
            font=("Segoe UI", 10),
            width=18,
        )
        variable.set(reverse[pairs[0][0]] if variable.get() not in reverse else reverse.get(variable.get(), labels[0]))
        if variable.get() not in labels:
            variable.set(labels[0])
        widget._alfa_map = mapping  # type: ignore[attr-defined]
        widget._alfa_reverse = reverse  # type: ignore[attr-defined]
        widget.pack(fill="x")
        return widget

    def _build_source(self, outer: tk.Frame) -> None:
        card = outer.inner  # type: ignore[attr-defined]
        pad = tk.Frame(card, bg=SURFACE)
        pad.pack(fill="both", expand=True, padx=16, pady=14)
        tk.Label(pad, text="Источник", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 13)).pack(anchor="w")
        source_hint = tk.Label(
            pad,
            text="Один файл или папка целиком. Из папки берутся txt, rpt, csv, dat и bcp.",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
            justify="left",
            anchor="w",
            wraplength=280,
        )
        source_hint.pack(anchor="w", fill="x", pady=(2, 10))
        self._hints.append(source_hint)

        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.custom_delim_var = tk.StringVar(value="|")
        self.fmt_var = tk.StringVar()
        self.rows_var = tk.StringVar(value="1000000")

        self._label(pad, "Файл или папка")
        self._entry(pad, self.source_var).pack(fill="x", ipady=2)
        source_buttons = WrapBar(pad, SURFACE)
        source_buttons.pack(fill="x", pady=(6, 8))
        source_buttons.add(DarkButton(source_buttons, "Файл", self._pick_source, "ghost"))
        source_buttons.add(DarkButton(source_buttons, "Папка", self._pick_folder, "ghost"))

        self._label(pad, "Кодировка")
        self.encoding_var = tk.StringVar()
        self.encoding_box = self._combo(pad, self.encoding_var, ENCODINGS)
        tk.Frame(pad, bg=SURFACE, height=8).pack()

        self._label(pad, "Формат")
        self.mode_var = tk.StringVar()
        self.mode_box = self._combo(pad, self.mode_var, MODES)
        tk.Frame(pad, bg=SURFACE, height=8).pack()

        self._label(pad, "Разделитель")
        self.delim_var = tk.StringVar()
        self.delim_box = self._combo(pad, self.delim_var, DELIMITERS)
        tk.Frame(pad, bg=SURFACE, height=8).pack()
        self._label(pad, "Кавычки")
        self.qual_var = tk.StringVar()
        self.qual_box = self._combo(pad, self.qual_var, QUALIFIERS)
        tk.Frame(pad, bg=SURFACE, height=8).pack()

        self._label(pad, "Свой разделитель, если выбран «Свой». Можно писать \\t, \\n и \\r")
        self._entry(pad, self.custom_delim_var).pack(fill="x")
        tk.Frame(pad, bg=SURFACE, height=8).pack()

        self._label(pad, "Имена колонок")
        self.header_var = tk.StringVar()
        self.header_box = self._combo(pad, self.header_var, HEADERS)
        tk.Frame(pad, bg=SURFACE, height=8).pack()

        self._label(pad, "Форматный файл .fmt — только для режима «Форматный файл bcp»")
        self._entry(pad, self.fmt_var).pack(fill="x", ipady=2)
        fmt_buttons = WrapBar(pad, SURFACE)
        fmt_buttons.pack(fill="x", pady=(6, 0))
        fmt_buttons.add(DarkButton(fmt_buttons, "FMT", self._pick_fmt, "ghost"))

    def _build_preview(self, outer: tk.Frame) -> None:
        card = outer.inner  # type: ignore[attr-defined]
        pad = tk.Frame(card, bg=SURFACE)
        pad.pack(fill="both", expand=True, padx=16, pady=14)
        tk.Label(pad, text="Книга и образец", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 13)).pack(anchor="w")
        preview_hint = tk.Label(
            pad,
            text="Лист Excel вмещает 1 048 575 строк данных. Дальше начинается следующий файл.",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
            justify="left",
            anchor="w",
            wraplength=280,
        )
        preview_hint.pack(anchor="w", fill="x", pady=(2, 10))
        self._hints.append(preview_hint)

        self.output_caption = tk.Label(
            pad,
            text="Куда сохранить .xlsx",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
        )
        self.output_caption.pack(anchor="w", fill="x", pady=(0, 4))
        self._hints.append(self.output_caption)
        self._entry(pad, self.output_var).pack(fill="x", ipady=2, pady=(0, 6))
        output_buttons = WrapBar(pad, SURFACE)
        output_buttons.pack(fill="x", pady=(0, 8))
        output_buttons.add(DarkButton(output_buttons, "Обзор", self._pick_output, "ghost"))

        self._label(pad, "Строк данных в одном файле")
        self._entry(pad, self.rows_var).pack(fill="x", ipady=2)
        tk.Frame(pad, bg=SURFACE, height=8).pack()

        self.summary = tk.Label(
            pad,
            text="Образец появится после разбора файла.",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
            justify="left",
            anchor="w",
        )
        self.summary.pack(fill="x", pady=(0, 8))
        self._hints.append(self.summary)

        tree_wrap = tk.Frame(pad, bg=BORDER, height=168)
        tree_wrap.pack(fill="both", expand=True)
        tree_wrap.pack_propagate(False)
        inner = tk.Frame(tree_wrap, bg=INPUT)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        self.tree = ttk.Treeview(inner, show="headings", height=6, style="Preview.Treeview")
        scroll_y = ttk.Scrollbar(inner, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(inner, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        inner.rowconfigure(0, weight=1)
        inner.columnconfigure(0, weight=1)
        self.tree["columns"] = ("empty",)
        self.tree.heading("empty", text="Образец")
        self.tree.column("empty", width=120, minwidth=48, stretch=True, anchor="w")

    def _build_pipeline(self, outer: tk.Frame) -> None:
        card = outer.inner  # type: ignore[attr-defined]
        pad = tk.Frame(card, bg=SURFACE)
        pad.pack(fill="both", expand=True, padx=16, pady=14)
        tk.Label(pad, text="Ход работы", bg=SURFACE, fg=TEXT, font=("Segoe UI Semibold", 13)).pack(anchor="w")
        self.status = tk.Label(
            pad,
            text="Ожидание",
            bg=SURFACE,
            fg=ACCENT,
            font=("Segoe UI", 10),
            anchor="w",
            justify="left",
            wraplength=280,
        )
        self.status.pack(fill="x", pady=(2, 0))
        self._hints.append(self.status)
        pace = tk.Label(
            pad,
            text="Один проход: строка читается, разбирается, превращается в ячейки и сразу пишется.",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 9),
            anchor="w",
            justify="left",
            wraplength=280,
        )
        pace.pack(anchor="w", fill="x", pady=(2, 10))
        self._hints.append(pace)
        self.bars: dict[str, StageBar] = {}
        for key, title, caption, color in STAGES:
            bar = StageBar(pad, title, caption, color)
            bar.pack(fill="x", pady=4)
            self.bars[key] = bar

        self.log = tk.Text(
            pad,
            height=6,
            bg=INPUT,
            fg="#d9d3c8",
            insertbackground=TEXT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            font=("Consolas", 9),
            wrap="word",
            padx=8,
            pady=6,
        )
        self.log.pack(fill="both", expand=True, pady=(10, 0))
        self.log.configure(state="disabled")

    def _build_actions(self) -> None:
        self.actions = WrapBar(self.root, BG)
        self.actions.pack(fill="x", padx=18, pady=12)
        self.probe_button = DarkButton(self.actions, "Разобрать образец", self._probe, "ghost")
        self.stop_button = DarkButton(self.actions, "Стоп", self._stop, "danger")
        self.run_button = DarkButton(self.actions, "Собрать Excel", self._start, "accent")
        folder_button = DarkButton(self.actions, "Открыть папку", self._open_folder, "ghost")
        self.stop_button.set_enabled(False)
        for button in (folder_button, self.probe_button, self.stop_button, self.run_button):
            self.actions.add(button)

    def _pick_source(self) -> None:
        path = filedialog.askopenfilename(
            title="Текстовая выгрузка",
            filetypes=(
                ("Текст SQL Server", "*.txt *.rpt *.csv *.dat *.bcp"),
                ("Все файлы", "*.*"),
            ),
        )
        if not path:
            return
        self.source_var.set(path)
        self.output_caption.configure(text="Куда сохранить .xlsx")
        stem = path.rsplit(".", 1)[0]
        self.output_var.set(stem + ".xlsx")
        self._probe()

    def _pick_folder(self) -> None:
        path = filedialog.askdirectory(title="Папка с текстовыми выгрузками")
        if not path:
            return
        self.source_var.set(path)
        self.output_var.set(path)
        self.output_caption.configure(text="Книги .xlsx сохраняются в эту же папку")
        self._probe()

    def _pick_output(self) -> None:
        source = self.source_var.get().strip()
        if source and __import__("pathlib").Path(source).is_dir():
            messagebox.showinfo("ALFA", "Для папки книги сохраняются рядом с каждым исходным файлом.")
            return
        path = filedialog.asksaveasfilename(
            title="Книга Excel",
            defaultextension=".xlsx",
            filetypes=(("Книга Excel", "*.xlsx"),),
        )
        if path:
            self.output_var.set(path)

    def _pick_fmt(self) -> None:
        path = filedialog.askopenfilename(
            title="Форматный файл bcp",
            filetypes=(("Формат bcp", "*.fmt"), ("Все файлы", "*.*")),
        )
        if path:
            self.fmt_var.set(path)
            self.mode_var.set("Форматный файл bcp")

    def _settings(self) -> Settings:
        delimiter = self.delim_box._alfa_map[self.delim_var.get()]  # type: ignore[attr-defined]
        if delimiter == "custom":
            delimiter = unescape_terminator(self.custom_delim_var.get())
        rows_text = self.rows_var.get().strip().replace(" ", "")
        if not rows_text.isdigit():
            raise ExportError("Число строк в файле должно быть целым.")
        rows = int(rows_text)
        if rows < 10000:
            raise ExportError("Минимум 10 000 строк на файл, иначе книг будет слишком много.")
        if rows > 1048575:
            raise ExportError("В одном листе Excel не больше 1 048 575 строк данных.")
        return Settings(
            encoding=self.encoding_box._alfa_map[self.encoding_var.get()],  # type: ignore[attr-defined]
            mode=self.mode_box._alfa_map[self.mode_var.get()],  # type: ignore[attr-defined]
            delimiter=delimiter,
            qualifier=self.qual_box._alfa_map[self.qual_var.get()],  # type: ignore[attr-defined]
            header=self.header_box._alfa_map[self.header_var.get()],  # type: ignore[attr-defined]
            fmt_path=self.fmt_var.get().strip(),
            rows_per_file=rows,
        )

    def _set_busy(self, busy: bool) -> None:
        self.run_button.set_enabled(not busy)
        self.probe_button.set_enabled(not busy)
        self.stop_button.set_enabled(busy)

    def _probe(self) -> None:
        path = self.source_var.get().strip()
        if not path:
            messagebox.showwarning("ALFA", "Сначала выберите файл или папку.")
            return
        if Path(path).is_dir():
            self._probe_folder(path)
            return
        try:
            settings = self._settings()
        except ExportError as exc:
            messagebox.showwarning("ALFA", str(exc))
            return
        self._log("Разбор образца…")
        self.status.configure(text="Смотрю начало файла")
        self.root.update_idletasks()
        try:
            layout = inspect(path, settings)
        except ExportError as exc:
            self._log(str(exc))
            messagebox.showerror("ALFA", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            self._log(str(exc))
            messagebox.showerror("ALFA", str(exc))
            return
        self._show_layout(layout)

    def _probe_folder(self, path: str) -> None:
        try:
            settings = self._settings()
            pairs = plan_folder(path)
        except ExportError as exc:
            self._log(str(exc))
            messagebox.showerror("ALFA", str(exc))
            return
        self._log(f"В папке файлов: {_format_int(len(pairs))}. Книги будут рядом с ними.")
        self.summary.configure(
            text=f"Папка · {_format_int(len(pairs))} файлов · образец: {Path(pairs[0][0]).name}"
        )
        self.status.configure(text="Смотрю первый файл")
        self.root.update_idletasks()
        try:
            layout = inspect(pairs[0][0], settings)
        except ExportError as exc:
            self._log(f"{Path(pairs[0][0]).name}: {exc}")
            messagebox.showerror("ALFA", str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            self._log(str(exc))
            messagebox.showerror("ALFA", str(exc))
            return
        self._show_layout(layout)
        self.summary.configure(
            text=f"Папка · {_format_int(len(pairs))} файлов · образец: {Path(pairs[0][0]).name} · {self.summary.cget('text')}"
        )

    def _show_layout(self, layout: Layout) -> None:
        for note in layout.notes:
            self._log(note)
        header = layout.preview_header or layout.column_names
        sample = layout.sample
        mode_name = {"delimited": "разделитель", "report": "отчёт", "fixed": "фиксированная ширина"}.get(
            layout.mode, layout.mode
        )
        self.summary.configure(
            text=(
                f"{layout.encoding} · {mode_name} · {describe_delimiter(layout.delimiter)} · "
                f"{len(header)} кол. · в образце {_format_int(len(sample))} строк"
            )
        )
        self._fill_tree(header, sample)
        self.status.configure(text="Образец готов")

    def _fill_tree(self, header: list[str], rows: list[list[str]]) -> None:
        self.tree.delete(*self.tree.get_children())
        if not header:
            self.tree["columns"] = ("empty",)
            self.tree.heading("empty", text="Пусто")
            self.tree.column("empty", width=120, minwidth=48, stretch=True, anchor="w")
            return
        columns = [f"c{index}" for index in range(len(header))]
        self.tree["columns"] = columns
        for column, name in zip(columns, header):
            self.tree.heading(column, text=name)
            self.tree.column(column, width=140, minwidth=48, stretch=True, anchor="w")
        for row in rows:
            values = list(row) + [""] * (len(header) - len(row))
            self.tree.insert("", "end", values=values[: len(header)])

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        source = self.source_var.get().strip()
        output = self.output_var.get().strip()
        if not source:
            messagebox.showwarning("ALFA", "Выберите файл или папку.")
            return
        source_path = Path(source)
        folder_pairs: list[tuple[str, str]] | None = None
        if source_path.is_dir():
            try:
                folder_pairs = plan_folder(source)
            except ExportError as exc:
                messagebox.showerror("ALFA", str(exc))
                return
            existing = [Path(target).name for _src, target in folder_pairs if Path(target).exists()]
            if existing:
                preview = "\n".join(existing[:8])
                more = "" if len(existing) <= 8 else f"\nи ещё {len(existing) - 8}"
                if not messagebox.askyesno(
                    "ALFA",
                    f"В папке уже есть книги, они будут перезаписаны:\n{preview}{more}",
                ):
                    return
        else:
            if not output:
                messagebox.showwarning("ALFA", "Укажите, куда сохранить книгу Excel.")
                return
            if not output.lower().endswith(".xlsx"):
                output += ".xlsx"
                self.output_var.set(output)
            if Path(output).exists():
                if not messagebox.askyesno("ALFA", f"Файл уже есть и будет перезаписан:\n{Path(output).name}"):
                    return
        try:
            settings = self._settings()
        except ExportError as exc:
            messagebox.showwarning("ALFA", str(exc))
            return
        self.cancel.clear()
        self._reset_bars()
        self._set_busy(True)
        self.status.configure(text="Идёт сборка")
        self._log(f"Старт: {source}")
        self.worker = threading.Thread(
            target=self._worker,
            args=(source, output, settings, folder_pairs),
            daemon=False,
            name="alfa-convert",
        )
        self.worker.start()

    def _worker(self, source: str, output: str, settings: Settings, folder_pairs: list[tuple[str, str]] | None) -> None:
        if folder_pairs:
            self._worker_folder(folder_pairs, settings)
            return
        try:
            result = convert(
                source,
                output,
                settings,
                lambda event: self.queue.put(("progress", event)),
                self.cancel.is_set,
                on_preview=lambda header, rows: self.queue.put(("preview", (header, rows))),
            )
            self.queue.put(("done", result))
        except Cancelled as exc:
            self.queue.put(("cancelled", str(exc)))
        except ExportError as exc:
            self.queue.put(("error", str(exc)))
        except Exception as exc:  # noqa: BLE001
            self.queue.put(("error", str(exc)))

    def _worker_folder(self, pairs: list[tuple[str, str]], settings: Settings) -> None:
        import time

        started = time.perf_counter()
        made: list[str] = []
        failed: list[str] = []
        rows = 0
        truncated = 0
        skipped = 0
        ragged = 0
        columns = 0
        total = len(pairs)
        try:
            for index, (src, target) in enumerate(pairs, start=1):
                if self.cancel.is_set():
                    raise Cancelled("Остановлено.")
                name = Path(src).name
                self.queue.put(("reset", None))
                self.queue.put(("log", f"{index} из {total}: {name}"))
                try:
                    result = convert(
                        src,
                        target,
                        settings,
                        lambda event, index=index, total=total, name=name: self.queue.put(
                            ("progress", _with_file(event, index, total, name))
                        ),
                        self.cancel.is_set,
                        on_preview=lambda header, rows, index=index: self.queue.put(("preview", (header, rows)))
                        if index == 1
                        else None,
                    )
                except Cancelled:
                    raise Cancelled(f"Остановлено. Уже собрано книг: {len(made)}.")
                except ExportError as exc:
                    failed.append(f"{name}: {exc}")
                    self.queue.put(("log", f"{name}: {exc}"))
                    continue
                except Exception as exc:  # noqa: BLE001
                    failed.append(f"{name}: {exc}")
                    self.queue.put(("log", f"{name}: {exc}"))
                    continue
                rows += result.rows
                made.extend(result.files)
                truncated += result.truncated_cells
                skipped += result.skipped_long
                ragged += result.ragged_rows
                columns = result.columns
            elapsed = time.perf_counter() - started
            self.queue.put(
                (
                    "done",
                    _FolderDone(rows, made, elapsed, truncated, skipped, ragged, columns, failed, total),
                )
            )
        except Cancelled as exc:
            self.queue.put(("cancelled", str(exc)))
        except Exception as exc:  # noqa: BLE001
            self.queue.put(("error", str(exc)))

    def _stop(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancel.set()
            self.status.configure(text="Останавливаю…")
            self._log("Остановка после текущей пачки. Книга будет закрыта.")

    def _reset_bars(self) -> None:
        for key, _title, caption, _color in STAGES:
            self.bars[key].set(0, caption)

    def _apply_progress(self, event: Progress) -> None:
        bar = self.bars.get(event.stage)
        if bar is not None and event.stage != "done":
            detail = event.message
            if event.rows:
                detail = f"{_format_int(event.rows)} строк"
                if event.rows_per_sec:
                    detail += f" · { _format_int(int(event.rows_per_sec)) }/с"
            if event.stage == "read" and event.file_size:
                detail = f"{_format_bytes(event.bytes_read)} из {_format_bytes(event.file_size)}"
            if event.stage == "write" and event.files_done:
                detail = f"{event.message} · файлов {event.files_done}"
                if event.rows:
                    detail = f"{_format_int(event.rows)} строк · файлов {event.files_done}"
            bar.set(1 if event.stage == "probe" and event.fraction >= 1 else event.fraction, detail)
        if event.stage == "done":
            for stage in self.bars.values():
                stage.set(1, stage.caption.cget("text"))
            self.status.configure(text="Готово")
        elif event.stage == "probe" and event.fraction >= 1:
            self._log(event.message)
        self.status.configure(text=event.message or self.status.cget("text"))

    def _poll(self) -> None:
        try:
            if self.closing or not self.root.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    self._apply_progress(payload)
                elif kind == "preview":
                    header, rows = payload
                    self._fill_tree(header, rows)
                    self.summary.configure(text=f"Первые строки файла · колонок {len(header)}")
                elif kind == "log":
                    self._log(str(payload))
                elif kind == "reset":
                    self._reset_bars()
                elif kind == "done":
                    self._finish_ok(payload)
                elif kind == "cancelled":
                    self._finish_stop(str(payload))
                elif kind == "error":
                    self._finish_error(str(payload))
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _finish_ok(self, result) -> None:
        self._set_busy(False)
        for stage in self.bars.values():
            stage.set(1, stage.caption.cget("text"))
        self._log(
            f"Готово за {result.elapsed:.1f} с. Строк {_format_int(result.rows)}, "
            f"колонок {result.columns}, файлов {len(result.files)}."
        )
        if result.truncated_cells:
            self._log(f"Обрезано ячеек длиннее 32 767 символов: {_format_int(result.truncated_cells)}.")
        if result.skipped_long:
            self._log(f"Пропущено слишком длинных строк: {_format_int(result.skipped_long)}.")
        if result.ragged_rows:
            self._log(f"Строк с другим числом полей: {_format_int(result.ragged_rows)}.")
        failed = getattr(result, "failed", None) or []
        if failed:
            self._log(f"Не собраны: {_format_int(len(failed))} из {getattr(result, 'sources', len(failed))}.")
            for item in failed[:30]:
                self._log(item)
        shown = result.files[:12]
        self._log(", ".join(shown) + ("…" if len(result.files) > 12 else ""))
        self.status.configure(text="Готово" if not failed else "Готово, есть ошибки")
        extra = f"\nС ошибкой: {len(failed)}" if failed else ""
        messagebox.showinfo(
            "ALFA",
            f"Собрано строк: {_format_int(result.rows)}\nКниг: {len(result.files)}{extra}",
        )

    def _finish_stop(self, message: str) -> None:
        self._set_busy(False)
        self._log(message)
        self.status.configure(text="Остановлено")

    def _finish_error(self, message: str) -> None:
        self._set_busy(False)
        self._log(message)
        self.status.configure(text="Ошибка")
        messagebox.showerror("ALFA", message)

    def _log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        line_count = int(self.log.index("end-1c").split(".")[0])
        if line_count > 200:
            self.log.delete("1.0", f"{line_count - 200}.0")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _open_folder(self) -> None:
        import os
        from pathlib import Path

        raw = self.output_var.get().strip() or self.source_var.get().strip()
        if not raw:
            return
        folder = Path(raw).parent if Path(raw).suffix else Path(raw)
        if folder.is_file():
            folder = folder.parent
        if folder.exists():
            os.startfile(folder)  # type: ignore[attr-defined]

    def _on_close(self) -> None:
        if self._fit_job is not None:
            self.root.after_cancel(self._fit_job)
            self._fit_job = None
        self.root.unbind_all("<MouseWheel>")
        if self.worker and self.worker.is_alive():
            self.cancel.set()
            self.closing = True
            self.status.configure(text="Закрываю после текущей пачки…")
            self.root.after(200, self._wait_close)
            return
        self.closing = True
        self.root.destroy()

    def _wait_close(self) -> None:
        if self.worker and self.worker.is_alive():
            self.root.after(200, self._wait_close)
            return
        self.root.destroy()


def smoke_ui() -> None:
    app = App()
    app.root.update_idletasks()
    app.root.update()
    assert app.run_button.winfo_exists()
    assert set(app.bars) == {"probe", "read", "parse", "form", "write"}
    app.closing = True
    app.root.destroy()


def main() -> None:
    if "--self-check" in sys.argv:
        from alfa_excel.convert import run_self_check

        run_self_check()
        smoke_ui()
        print("SELF_CHECK_OK", flush=True)
        return
    try:
        app = App()
        app.root.after(200, lambda: print("UI_READY", flush=True))
        app.root.mainloop()
    except Exception:
        import traceback

        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()
