from __future__ import annotations

import sys
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

from .engine import (FIELDS, PDF_FIELDS, Metadata, backup, capture_times, load,
                     restore_times, set_filesystem_times, write_office, write_pdf)

OPEN_PATTERNS = ("*.docx", "*.xlsx", "*.pptx", "*.docm", "*.xlsm", "*.pptm",
                 "*.dotx", "*.xltx", "*.potx", "*.dotm", "*.xltm", "*.potm",
                 "*.ppsx", "*.ppsm", "*.pdf")


def _pick(families: set[str], candidates: tuple[str, ...], fallback: str) -> str:
    """返回第一个系统里真实存在的字体名。

    macOS 没有 Microsoft YaHei UI / Consolas；直接写死会让 Tk 静默回退到
    .AppleSystemUIFont，界面虽然能显示但字重与间距和设计稿不一致。
    """
    for name in candidates:
        if name in families:
            return name
    return fallback


class Editor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("文件元数据编辑")
        families = set(tkfont.families(self))
        if sys.platform == "darwin":
            self.ui_font = _pick(families, ("PingFang SC", "Hiragino Sans GB", "Heiti SC"), "Helvetica")
            self.mono_font = _pick(families, ("Menlo", "SF Mono", "Monaco"), "Courier")
        else:
            self.ui_font = _pick(families, ("Microsoft YaHei UI", "Microsoft YaHei", "SimHei"), "Helvetica")
            self.mono_font = _pick(families, ("Consolas", "Courier New"), "Courier")
        self.after_idle(self._fit_initial_window)
        self.meta: Metadata | None = None
        self.vars: dict[str, tk.StringVar] = {}
        self.backup_var = tk.BooleanVar(value=True)
        self.path_var = tk.StringVar()
        self._style()
        self._layout()
        self._menu()
        self._shortcuts()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._activate()

    def _activate(self):
        """把窗口提到前台。

        从 .app 或 shell 启动时，Tk 在 macOS 上常以后台进程身份出现，
        窗口画好了却排在其他应用后面，看起来像「点了没反应」。
        """
        if sys.platform != "darwin":
            return
        self.lift()
        self.attributes("-topmost", True)
        self.after_idle(self.attributes, "-topmost", False)
        self.focus_force()

    def _fit_initial_window(self):
        """Size the first window to the actual DPI-aware monitor dimensions."""
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        # macOS 的 winfo_screenheight 已包含菜单栏与 Dock 之外的整块屏幕，
        # 可用高度还要再扣掉约 24px 菜单栏和 28px 标题栏，否则窗口会顶到屏幕外。
        usable_height = screen_height - (52 if sys.platform == "darwin" else 0)
        width = min(1240, max(760, int(screen_width * 0.90)))
        height = min(780, max(500, int(usable_height * 0.82)))
        self.minsize(min(900, max(680, int(screen_width * 0.62))), min(600, max(440, int(usable_height * 0.58))))
        self.geometry(f"{width}x{height}")

    def _style(self):
        style = ttk.Style(self)
        # aqua 主题下 ttk.Button/Labelframe 的 background 配置会被忽略，
        # 原设计的卡片式配色依赖 clam，因此所有平台统一用 clam。
        style.theme_use("clam")
        style.configure("TFrame", background="#f7f8fa")
        style.configure("Card.TLabelframe", background="#ffffff", bordercolor="#dfe3e8", relief="solid")
        style.configure("Card.TLabelframe.Label", background="#ffffff", foreground="#243447", font=(self.ui_font, 10, "bold"))
        style.configure("Title.TLabel", background="#f7f8fa", foreground="#172b4d", font=(self.ui_font, 15, "bold"))
        style.configure("Muted.TLabel", background="#f7f8fa", foreground="#667085", font=(self.ui_font, 9))
        style.configure("TLabel", background="#f7f8fa", font=(self.ui_font, 10))
        style.configure("TLabelframe.Label", font=(self.ui_font, 10, "bold"))
        style.configure("TButton", font=(self.ui_font, 10), padding=(10, 5))
        style.configure("TRadiobutton", background="#f7f8fa", font=(self.ui_font, 10))
        style.configure("TEntry", font=(self.ui_font, 10))
        style.configure("Treeview", font=(self.ui_font, 10), rowheight=22)
        style.configure("Treeview.Heading", font=(self.ui_font, 10, "bold"))
        style.configure("Accent.TButton", font=(self.ui_font, 10, "bold"), foreground="white", background="#2563eb", padding=(16, 7))
        style.map("Accent.TButton", background=[("active", "#1d4ed8")])

    def _layout(self):
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="文件元数据编辑", style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="本地编辑 · Office 与 PDF · 保存前可自动备份", style="Muted.TLabel").pack(anchor="w", pady=(2, 14))
        chooser = ttk.Frame(outer)
        chooser.pack(fill="x", pady=(0, 12))
        ttk.Label(chooser, text="文件", width=6).pack(side="left")
        ttk.Entry(chooser, textvariable=self.path_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(chooser, text="选择文件", command=self.choose).pack(side="left")
        body = ttk.Panedwindow(outer, orient="horizontal")
        left = ttk.Labelframe(body, text=" 文件信息 ", style="Card.TLabelframe", padding=8)
        right = ttk.Labelframe(body, text=" 编辑属性 ", style="Card.TLabelframe", padding=14)
        body.add(left, weight=1); body.add(right, weight=2)
        self.tree = ttk.Treeview(left, columns=("value",), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="属性"); self.tree.heading("value", text="值")
        self.tree.column("#0", width=145, stretch=False); self.tree.column("value", width=225)
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        self.form_canvas = tk.Canvas(right, highlightthickness=0, background="#ffffff")
        form_scroll = ttk.Scrollbar(right, orient="vertical", command=self.form_canvas.yview)
        self.form = ttk.Frame(self.form_canvas, padding=(2, 2, 10, 2))
        self.form.bind("<Configure>", lambda _e: self.form_canvas.configure(scrollregion=self.form_canvas.bbox("all")))
        self.form_canvas.create_window((0, 0), window=self.form, anchor="nw", tags="form")
        self.form_canvas.bind("<Configure>", lambda e: self.form_canvas.itemconfigure("form", width=e.width))
        self.form_canvas.configure(yscrollcommand=form_scroll.set)
        self.form_canvas.pack(side="left", fill="both", expand=True); form_scroll.pack(side="right", fill="y")
        # Pack the non-shrinking footer first.  The body is the only region
        # allowed to lose height while the user makes the window smaller.
        footer = ttk.Frame(outer); footer.pack(fill="x", side="bottom", pady=(12, 0))
        body.pack(fill="both", expand=True)
        ttk.Label(footer, text="保存前备份：").pack(side="left")
        ttk.Radiobutton(footer, text="是", variable=self.backup_var, value=True).pack(side="left", padx=(2, 8))
        ttk.Radiobutton(footer, text="否", variable=self.backup_var, value=False).pack(side="left")
        self.status = ttk.Label(footer, text="请选择一个文件开始", style="Muted.TLabel"); self.status.pack(side="left", padx=16)
        ttk.Button(footer, text="清除嵌入属性", command=self.clear).pack(side="right")
        ttk.Button(footer, text="保存修改", style="Accent.TButton", command=self.save).pack(side="right", padx=(0, 8))
        # Let the wheel work naturally even when the pointer is above an entry,
        # label, or nested form widget rather than only above the scrollbar.
        self.bind_all("<MouseWheel>", self._mousewheel, add="+")
        # 部分 Tk 构建（尤其是 X11 后端）用 Button-4/5 表示滚轮。
        self.bind_all("<Button-4>", lambda e: self._mousewheel(_FakeWheel(e, 120)), add="+")
        self.bind_all("<Button-5>", lambda e: self._mousewheel(_FakeWheel(e, -120)), add="+")

    def _menu(self):
        """macOS 必须有 Edit 菜单，否则输入框里的 ⌘C / ⌘V / ⌘X 全部无效。"""
        bar = tk.Menu(self)
        file_menu = tk.Menu(bar, tearoff=0)
        file_menu.add_command(label="打开文件…", command=self.choose, accelerator="⌘O" if sys.platform == "darwin" else "Ctrl+O")
        file_menu.add_command(label="保存修改", command=self.save, accelerator="⌘S" if sys.platform == "darwin" else "Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="退出" if sys.platform == "darwin" else "关闭", command=self.destroy,
                              accelerator="⌘Q" if sys.platform == "darwin" else "")
        bar.add_cascade(label="文件", menu=file_menu)

        edit_menu = tk.Menu(bar, tearoff=0)
        for label, event, accel in (("剪切", "<<Cut>>", "⌘X"), ("复制", "<<Copy>>", "⌘C"),
                                    ("粘贴", "<<Paste>>", "⌘V"), ("全选", "<<SelectAll>>", "⌘A")):
            edit_menu.add_command(label=label, accelerator=accel if sys.platform == "darwin" else "",
                                  command=lambda ev=event: self._emit(ev))
        bar.add_cascade(label="编辑", menu=edit_menu)
        self.config(menu=bar)

    def _emit(self, event: str):
        widget = self.focus_get()
        if widget is not None:
            widget.event_generate(event)
        return "break"

    def _shortcuts(self):
        mod = "Command" if sys.platform == "darwin" else "Control"
        self.bind_all(f"<{mod}-o>", lambda _e: (self.choose(), "break")[1], add="+")
        self.bind_all(f"<{mod}-s>", lambda _e: (self.save(), "break")[1], add="+")

    @staticmethod
    def _inside(widget, ancestor):
        while widget is not None:
            if widget == ancestor:
                return True
            parent = widget.winfo_parent()
            widget = widget.nametowidget(parent) if parent else None
        return False

    def _mousewheel(self, event):
        under_pointer = self.winfo_containing(event.x_root, event.y_root)
        magnitude = abs(event.delta)
        # Windows 每格滚轮是 ±120 的整数倍；macOS 触控板给出的是 ±1..±15 的小值，
        # 原式 magnitude // 120 在 macOS 上恒等于 0，快速滚动会丢失加速。
        steps = max(1, magnitude // 120) if magnitude >= 120 else 1
        steps = -steps if event.delta > 0 else steps
        if under_pointer and self._inside(under_pointer, self.tree):
            self.tree.yview_scroll(steps, "units")
            return "break"
        if under_pointer and self._inside(under_pointer, self.form_canvas):
            self.form_canvas.yview_scroll(steps, "units")
            return "break"

    def choose(self):
        path = filedialog.askopenfilename(
            filetypes=[("支持的文件", OPEN_PATTERNS), ("所有文件", "*")],
            defaultextension=".docx",
        )
        if path: self.open(path)

    def open(self, path: str):
        try:
            self.meta = load(path); self.path_var.set(path); self._render(); self.status.configure(text=f"已读取：{self.meta.kind} 文档")
        except Exception as e: messagebox.showerror("无法打开", str(e))

    def _render(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for child in self.form.winfo_children(): child.destroy()
        assert self.meta
        fs = self.tree.insert("", "end", text="文件系统时间", open=True)
        self.tree.insert(fs, "end", text="创建时间", values=(self.meta.values.get("filesystem_created", ""),))
        self.tree.insert(fs, "end", text="修改时间", values=(self.meta.values.get("filesystem_modified", ""),))
        props = self.tree.insert("", "end", text=f"{self.meta.kind} 属性", open=True)
        for key, label in (PDF_FIELDS if self.meta.kind == "PDF" else FIELDS): self.tree.insert(props, "end", text=label, values=(self.meta.values.get(key, ""),))
        if self.meta.kind == "Office":
            custom_node = self.tree.insert("", "end", text="自定义属性", open=True)
            for key, value in self.meta.custom.items(): self.tree.insert(custom_node, "end", text=key, values=(value,))
        self.vars = {}
        self._section("文件系统时间", [("filesystem_created", "创建时间"), ("filesystem_modified", "修改时间")], readonly=False)
        groups = [("来源与描述", [x for x in (PDF_FIELDS if self.meta.kind == "PDF" else FIELDS) if x[0] not in {"created", "modified", "lastPrinted", "totalTime"}]), ("文档元数据时间", [x for x in (PDF_FIELDS if self.meta.kind == "PDF" else FIELDS) if x[0] in {"created", "modified", "lastPrinted", "totalTime"}])]
        for title, fields in groups: self._section(title, fields)
        if self.meta.kind == "Office":
            box = ttk.Labelframe(self.form, text=" 自定义属性（每行：名称 = 值） ", style="Card.TLabelframe", padding=10); box.pack(fill="x", pady=(0, 12))
            self.custom_text = tk.Text(box, height=5, font=(self.mono_font, 10), relief="solid", borderwidth=1)
            self.custom_text.pack(fill="x"); self.custom_text.insert("1.0", "\n".join(f"{k} = {v}" for k, v in self.meta.custom.items()))
        else:
            ttk.Label(self.form, text="PDF 采用增量更新保存。加密 PDF 与已签名 PDF 不建议修改。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)

    def _section(self, title, fields, readonly=False):
        box = ttk.Labelframe(self.form, text=f" {title} ", style="Card.TLabelframe", padding=10); box.pack(fill="x", pady=(0, 12))
        for row, (key, label) in enumerate(fields):
            ttk.Label(box, text=label, width=18).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            var = tk.StringVar(value=self.meta.values.get(key, "")); self.vars[key] = var
            ttk.Entry(box, textvariable=var).grid(row=row, column=1, sticky="ew", pady=5)
        box.columnconfigure(1, weight=1)

    def _custom(self):
        result = {}
        for line in self.custom_text.get("1.0", "end").splitlines():
            if not line.strip(): continue
            if "=" not in line: raise ValueError("自定义属性每行必须使用“名称 = 值”格式。")
            key, value = line.split("=", 1); key = key.strip()
            if not key: raise ValueError("自定义属性名称不能为空。")
            result[key] = value.strip()
        return result

    def _system_times(self):
        assert self.meta
        set_filesystem_times(self.meta.path, self.vars["filesystem_created"].get().strip(), self.vars["filesystem_modified"].get().strip())

    def save(self):
        if not self.meta: return
        try:
            if self.backup_var.get(): backup(self.meta.path)
            values = {key: var.get().strip() for key, var in self.vars.items()}
            if self.meta.kind == "Office": write_office(self.meta, values, self._custom(), False)
            else: write_pdf(self.meta, values, False)
            self._system_times(); self.open(str(self.meta.path)); self.status.configure(text="保存成功，已重新读取验证。")
        except Exception as e: messagebox.showerror("保存失败", str(e))

    def clear(self):
        if not self.meta: return
        if not messagebox.askyesno("确认清除", "将清空文档内嵌属性和自定义属性；不会更改文件系统时间。是否继续？"): return
        try:
            if self.backup_var.get(): backup(self.meta.path)
            # 写入本身会重置文件系统时间（Office 换文件、PDF 刷新修改时间），
            # 因此先存快照、写完再还原，才能让提示语里的承诺成立。
            preserved = capture_times(self.meta.path)
            values = {key: "" for key in self.vars}
            if self.meta.kind == "Office": write_office(self.meta, values, {}, True)
            else: write_pdf(self.meta, values, True)
            if not restore_times(self.meta.path, preserved):
                self.status.configure(text="嵌入属性已清除，但文件系统时间未能保留。")
            else:
                self.open(str(self.meta.path)); self.status.configure(text="嵌入属性已清除。")
        except Exception as e: messagebox.showerror("清除失败", str(e))


class _FakeWheel:
    """把 Button-4/5 事件适配成 _mousewheel 需要的 delta 形状。"""

    def __init__(self, event, delta: int):
        self.delta = delta
        self.x_root = event.x_root
        self.y_root = event.y_root


def launch():
    Editor().mainloop()
