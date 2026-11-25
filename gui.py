# -*- coding: utf-8 -*-
"""
图形界面 (Tkinter)。由 skykiwi_scraper.run_gui() 调用; core 为主模块 (抓取 / 筛选 / 导出)。

结构:
  App          顶部工具栏 (抓取 / 停止 / 导出)、家庭住址栏、两个标签页、底部状态栏; 后台任务一次只跑一个
  CategoryTab  一个类别 (租房 / 招聘) 的筛选条件、结果表、排序和详情区, 两页互不影响
  MultiSelect  可多选的下拉框: 勾选时不收起, 点外面或按 Esc 关闭
"""

from __future__ import annotations

import os
import queue
import threading
import webbrowser

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog

import geo

DEBOUNCE_MS = 300      # 输入框停止输入后多久刷新结果

# 表格只显示关键列 (其余字段在下方详情区和导出的 Excel 里): (字段, 表头, 宽度单位, 对齐)
VIEW_COLUMNS = {
    "rent": [("post_date", "日期", 9, "w"), ("region", "区域", 7, "w"), ("title", "标题", 40, "w"),
             ("price", "周租金", 7, "e"), ("distance", "距离 km", 9, "e"), ("rent_type", "出租类型", 12, "w"),
             ("house_type", "房屋类型", 12, "w"), ("loc_desc", "定位依据", 20, "w")],
    "job": [("post_date", "日期", 9, "w"), ("region", "岗位类别", 9, "w"), ("title", "标题", 34, "w"),
            ("position", "职位", 12, "w"), ("nature", "性质", 6, "w"), ("salary_num", "时薪", 6, "e"),
            ("distance", "距离 km", 9, "e"), ("loc_desc", "定位依据", 20, "w")],
}
DEFAULT_SORTS = {"rent": {"出租"}, "job": {"招聘"}}     # 默认只看出租 / 招聘, 不看求租 / 求职


def enable_dpi_awareness() -> bool:
    """声明进程支持高 DPI (Windows)。否则在 125%/150% 缩放下系统会把窗口位图拉伸, 文字发虚。"""
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # PROCESS_SYSTEM_DPI_AWARE
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()
        return True
    except (AttributeError, OSError):
        return False


# ---------------------------------------------------------------- 多选下拉框 ----

class MultiSelect(ttk.Frame):
    def __init__(self, master, options, on_change, width=24, all_text="全部"):
        super().__init__(master)
        self.options = list(options)
        self.vars = {o: tk.BooleanVar(value=False) for o in self.options}
        self.on_change = on_change
        self.all_text = all_text
        self.text = tk.StringVar()
        self.button = ttk.Button(self, textvariable=self.text, width=width, command=self.toggle)
        self.button.pack(fill="x")
        self.popup = None
        self._update_text()
        self.winfo_toplevel().bind("<Button-1>", self._click_outside, add="+")

    def get(self) -> list[str]:
        return [o for o in self.options if self.vars[o].get()]

    def set(self, values):
        for o in self.options:
            self.vars[o].set(o in values)
        self._update_text()

    def _update_text(self):
        sel = self.get()
        if not sel:
            label = self.all_text
        elif len(sel) <= 2:
            label = ", ".join(sel)
        else:
            label = f"{sel[0]} 等 {len(sel)} 项"
        self.text.set(label + "  ▾")

    def _changed(self):
        self._update_text()
        self.on_change()

    def _clear(self):
        self.set([])
        self.on_change()

    def toggle(self):
        if self.popup:
            self.close()
        else:
            self.open()

    def open(self):
        top = tk.Toplevel(self)
        top.overrideredirect(True)
        top.transient(self.winfo_toplevel())
        frm = ttk.Frame(top, padding=6, relief="solid", borderwidth=1)
        frm.pack(fill="both", expand=True)
        ttk.Button(frm, text="全部 (清除选择)", command=self._clear).pack(fill="x", pady=(0, 4))
        for o in self.options:
            ttk.Checkbutton(frm, text=o, variable=self.vars[o], command=self._changed).pack(anchor="w", pady=1)
        ttk.Button(frm, text="完成", command=self.close).pack(fill="x", pady=(4, 0))
        self.button.update_idletasks()      # 刚切换标签页时按钮坐标可能还没更新
        x = self.button.winfo_rootx()
        y = self.button.winfo_rooty() + self.button.winfo_height()
        top.geometry(f"+{x}+{y}")
        top.bind("<Escape>", lambda e: self.close())
        self.popup = top
        top.focus_set()

    def close(self):
        if self.popup is not None:
            self.popup.destroy()
            self.popup = None

    def _click_outside(self, e):
        # 主窗口里点了别处 (不是本按钮) -> 收起; 下拉框里的点击不会走到这里
        if self.popup is not None and e.widget is not self.button:
            self.close()


# ---------------------------------------------------------------- 一个类别的标签页 ----

class CategoryTab:
    def __init__(self, app: "App", category: str):
        core = app.core
        self.app, self.core, self.cat = app, core, category
        self.cfg = core.CATEGORIES[category]
        self.rent = category == "rent"
        self.items: list = []
        self.shown: list = []
        self.sort_key: str | None = None
        self.sort_rev = False
        self._after = None

        self.frame = ttk.Frame(app.notebook, padding=(6, 6, 6, 0))
        app.notebook.add(self.frame, text=f"  {self.cfg['name']}  ")

        # ---- 筛选条件 (每页独立) ----
        self.min_var, self.max_var = tk.StringVar(), tk.StringVar()
        self.keep_np_var = tk.BooleanVar(value=True)
        self.addr_var, self.kw_var, self.ex_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.rent_type_var, self.house_type_var = tk.StringVar(), tk.StringVar()
        self.nature_var, self.pos_var = tk.StringVar(value="全部"), tk.StringVar()
        self.noads_var = tk.BooleanVar(value=False)
        self.taken_var = tk.BooleanVar(value=True)
        self.sort_vars = {n: tk.BooleanVar(value=n in DEFAULT_SORTS[category])
                          for n in self.cfg["sortids"].values()}

        box = ttk.LabelFrame(self.frame, text="筛选（修改后立即生效）", padding=(8, 4))
        box.pack(fill="x")
        r0 = ttk.Frame(box)
        r0.pack(fill="x")
        ttk.Label(r0, text=self.cfg["type_label"] + ":").pack(side="left")
        self.type_select = MultiSelect(r0, self.cfg["typeids"].values(), self.refresh_soon)
        self.type_select.pack(side="left", padx=(4, 14))
        ttk.Label(r0, text="类型:").pack(side="left")
        for n, v in self.sort_vars.items():
            ttk.Checkbutton(r0, text=n, variable=v, command=self.refresh_soon).pack(side="left", padx=2)
        if not self.rent:
            ttk.Label(r0, text="    工作性质:").pack(side="left")
            cb = ttk.Combobox(r0, width=6, textvariable=self.nature_var, state="readonly",
                              values=["全部", "全职", "兼职"])
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_soon())
        ttk.Checkbutton(r0, text="排除已租/已招满" if self.rent else "排除已招满",
                        variable=self.taken_var, command=self.refresh_soon).pack(side="left", padx=(14, 2))
        ttk.Checkbutton(r0, text="排除商家广告", variable=self.noads_var,
                        command=self.refresh_soon).pack(side="left", padx=2)
        ttk.Button(r0, text="重置筛选", command=self.reset).pack(side="right")

        r1 = ttk.Frame(box)
        r1.pack(fill="x", pady=(6, 0))
        ttk.Label(r1, text="周租金 NZD:" if self.rent else "时薪 NZD:").pack(side="left")
        self.min_entry = ttk.Entry(r1, width=7, textvariable=self.min_var)
        self.min_entry.pack(side="left")
        ttk.Label(r1, text="~").pack(side="left")
        self.max_entry = ttk.Entry(r1, width=7, textvariable=self.max_var)
        self.max_entry.pack(side="left")
        ttk.Checkbutton(r1, text="保留无价格的帖子", variable=self.keep_np_var,
                        command=self.refresh_soon).pack(side="left", padx=(6, 14))
        ttk.Label(r1, text="地址包含:" if self.rent else "工作地点包含:").pack(side="left")
        ttk.Entry(r1, width=14, textvariable=self.addr_var).pack(side="left", padx=(2, 14))
        if self.rent:
            ttk.Label(r1, text="出租类型:").pack(side="left")
            c1 = ttk.Combobox(r1, width=10, textvariable=self.rent_type_var,
                              values=["", "整租", "分租", "单间", "Flatmates", "短租"])
            c1.pack(side="left", padx=(2, 14))
            ttk.Label(r1, text="房屋类型:").pack(side="left")
            c2 = ttk.Combobox(r1, width=10, textvariable=self.house_type_var,
                              values=["", "独立屋", "公寓", "单元房", "Townhouse", "Unit"])
            c2.pack(side="left", padx=2)
            for c in (c1, c2):
                c.bind("<<ComboboxSelected>>", lambda e: self.refresh_soon())
        else:
            ttk.Label(r1, text="职位包含:").pack(side="left")
            ttk.Entry(r1, width=14, textvariable=self.pos_var).pack(side="left", padx=2)

        r2 = ttk.Frame(box)
        r2.pack(fill="x", pady=(6, 2))
        ttk.Label(r2, text="关键词 (空格分隔, 须全部包含):").pack(side="left")
        ttk.Entry(r2, width=26, textvariable=self.kw_var).pack(side="left", padx=(2, 14))
        ttk.Label(r2, text="排除词:").pack(side="left")
        ttk.Entry(r2, width=18, textvariable=self.ex_var).pack(side="left", padx=2)

        for v in (self.min_var, self.max_var, self.addr_var, self.kw_var, self.ex_var,
                  self.rent_type_var, self.house_type_var, self.pos_var):
            v.trace_add("write", lambda *a: self.refresh_soon())

        # ---- 结果表 + 详情区 (可拖动分隔条调整高度) ----
        pw = ttk.Panedwindow(self.frame, orient="vertical")
        pw.pack(fill="both", expand=True, pady=(6, 0))
        tf = ttk.Frame(pw)
        self.tree = ttk.Treeview(tf, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        df = ttk.Frame(pw)
        self.detail = tk.Text(df, height=6, wrap="word", relief="flat", padx=8, pady=6,
                              background="#FAFAFA", state="disabled")
        dsb = ttk.Scrollbar(df, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.pack(side="left", fill="both", expand=True)
        dsb.pack(side="right", fill="y")
        pw.add(tf, weight=4)
        pw.add(df, weight=1)

        ui = app.ui
        self.columns = VIEW_COLUMNS[category]
        self.tree["columns"] = [c[0] for c in self.columns]
        for key, head, w, anchor in self.columns:
            self.tree.heading(key, text=head, anchor=anchor, command=lambda kk=key: self.sort_by(kk))
            self.tree.column(key, width=int(w * 8 * ui), minwidth=int(36 * ui), anchor=anchor,
                             stretch=(key == "title"))
        self.tree.tag_configure("coarse", foreground="#8A8A8A")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.open_selected)
        self.tree.bind("<Return>", self.open_selected)
        self.tree.bind("<Control-c>", self.copy_selected)
        self.set_detail("抓取后在这里显示选中帖子的详细信息。双击或按回车在浏览器打开原帖。")

    # ---- 数据 ----
    def set_items(self, items):
        self.items = list(items)
        self.refresh()

    def refresh_soon(self, *_):
        if self._after:
            self.frame.after_cancel(self._after)
        self._after = self.frame.after(DEBOUNCE_MS, self.refresh)

    def _num(self, var, entry):
        s = var.get().strip()
        if not s:
            entry.configure(style="TEntry")
            return None
        try:
            v = float(s)
            entry.configure(style="TEntry")
            return v
        except ValueError:
            entry.configure(style="Bad.TEntry")   # 非法数字: 标红, 该条件不生效
            return None

    def criteria(self):
        app = self.app
        nature = self.nature_var.get()
        home = app.home
        return self.core.Criteria(
            regions=self.type_select.get(),
            sorts=[n for n, v in self.sort_vars.items() if v.get()],
            min_price=self._num(self.min_var, self.min_entry),
            max_price=self._num(self.max_var, self.max_entry),
            keep_no_price=self.keep_np_var.get(),
            keywords=self.kw_var.get().split(), exclude=self.ex_var.get().split(),
            address=self.addr_var.get().strip(),
            rent_type=self.rent_type_var.get().strip() if self.rent else "",
            house_type=self.house_type_var.get().strip() if self.rent else "",
            nature="" if self.rent or nature in ("", "全部") else nature,
            position=self.pos_var.get().strip() if not self.rent else "",
            exclude_ads=self.noads_var.get(), exclude_taken=self.taken_var.get(),
            max_distance=app.max_distance() if home else None,
            sort_distance=app.sortdist_var.get() and bool(home) and self.sort_key is None)

    def refresh(self):
        self._after = None
        core = self.core
        geo.compute_distances(self.items, self.app.home)
        crit = self.criteria()
        shown = core.apply_filters(self.items, crit)
        if self.sort_key:
            shown = core.sort_listings(shown, self.sort_key, self.sort_rev)
        self.shown = shown
        self._fill(shown)
        self._update_headings(crit.sort_distance)
        self.app.update_count()

    def _cell(self, it, key):
        if key == "distance":
            return self.core.format_distance(it)
        v = getattr(it, key, "")
        if v is None:
            return ""
        if isinstance(v, float):
            return f"{v:g}"
        return v

    def _fill(self, rows):
        sel = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        keys = [c[0] for c in self.columns]
        for it in rows:
            tags = ("coarse",) if it.loc_level in ("区域", "") else ()
            self.tree.insert("", "end", iid=str(it.tid), values=[self._cell(it, k) for k in keys], tags=tags)
        keep = [i for i in sel if self.tree.exists(i)]
        if keep:
            self.tree.selection_set(keep)

    def _update_headings(self, distance_sorted):
        for key, head, _, _ in self.columns:
            arrow = ""
            if key == self.sort_key:
                arrow = " ▼" if self.sort_rev else " ▲"
            elif key == "distance" and distance_sorted:
                arrow = " ▲"
            self.tree.heading(key, text=head + arrow)

    def sort_by(self, key):
        if self.sort_key == key:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_key, self.sort_rev = key, False
        self.refresh()

    def clear_manual_sort(self):
        self.sort_key, self.sort_rev = None, False

    def reset(self):
        for v in (self.min_var, self.max_var, self.addr_var, self.kw_var, self.ex_var,
                  self.rent_type_var, self.house_type_var, self.pos_var):
            v.set("")
        self.nature_var.set("全部")
        self.keep_np_var.set(True)
        self.noads_var.set(False)
        self.taken_var.set(True)
        for n, v in self.sort_vars.items():
            v.set(n in DEFAULT_SORTS[self.cat])
        self.type_select.set([])
        self.clear_manual_sort()
        self.refresh()

    # ---- 选中 / 打开 / 复制 ----
    def _find(self, tid):
        return next((it for it in self.items if str(it.tid) == str(tid)), None)

    def set_detail(self, text):
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", text)
        self.detail.configure(state="disabled")

    def on_select(self, _e=None):
        sel = self.tree.selection()
        it = self._find(sel[0]) if sel else None
        if not it:
            return
        home = self.app.home
        dist = (f"{it.distance:g} km（直线）" if it.distance is not None
                else ("-" if home else "未设置家庭住址"))
        lines = [it.title, it.url, f"发布时间: {it.post_time or it.post_date}    发帖人: {it.author}    "
                                   f"回复 {it.replies} / 查看 {it.views}",
                 f"预估距离: {dist}    定位依据: {it.loc_desc}"]
        if self.rent:
            lines.append(f"租金: {it.price_text or '-'}    出租类型: {it.rent_type}    房屋类型: {it.house_type}    "
                         f"地址: {it.address}")
            if it.facilities:
                lines.append(f"设施: {it.facilities}")
        else:
            lines.append(f"公司: {it.company}    职位: {it.position}    薪金: {it.salary}    性质: {it.nature}    "
                         f"地点: {it.location}    时间: {it.work_time}")
        lines.append(f"电话: {it.phone or '-'}    微信: {it.wechat or '-'}"
                     + (f"    邮箱: {it.email}" if getattr(it, "email", "") else ""))
        for k, v in it.extra.items():
            lines.append(f"{k}: {v}")
        lines += ["", it.content]
        self.set_detail("\n".join(lines))

    def open_selected(self, _e=None):
        for iid in self.tree.selection()[:5]:
            it = self._find(iid)
            if it:
                webbrowser.open(it.url)

    def copy_selected(self, _e=None):
        rows = [self._find(i) for i in self.tree.selection()]
        text = "\n".join(f"{it.title}\t{it.url}" for it in rows if it)
        if text:
            self.tree.clipboard_clear()
            self.tree.clipboard_append(text)
            self.app.status.set(f"已复制 {len(rows)} 条 (标题 + 链接)")
        return "break"


# ---------------------------------------------------------------- 主窗口 ----

class App:
    def __init__(self, core, root: tk.Tk):
        self.core, self.root = core, root
        self.ui = root.winfo_fpixels("1i") / 96.0          # 屏幕缩放比例 (100% = 1.0)
        self.msgq: queue.Queue = queue.Queue()
        self.task: str | None = None
        self.stop_event: threading.Event | None = None
        self._poll_id = None
        self.log_path = core.setup_logging()

        cfg = geo.load_config()
        self.home = ((cfg["home_lat"], cfg["home_lon"], cfg.get("home_desc", ""))
                     if cfg.get("home_lat") is not None else None)

        self._setup_window()
        self._build_toolbar()
        self._build_home(cfg)
        # 底部状态栏先于标签页分配空间: 窗口变矮时只压缩结果表
        self._build_statusbar()
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(4, 0))
        self.tabs = {"rent": CategoryTab(self, "rent"), "job": CategoryTab(self, "job")}
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self._on_tab_changed())

        root.bind("<F5>", lambda e: self.start_fetch())
        root.bind("<Control-e>", lambda e: self.export())
        root.bind("<Control-E>", lambda e: self.export())
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Destroy>", self._on_destroy, add="+")
        root.report_callback_exception = self._on_callback_error
        root.app = self   # 便于测试

        self.show_home_status()
        self._on_tab_changed()
        self._poll()

    # ---- 布局 ----
    def _setup_window(self):
        root, ui = self.root, self.ui
        root.title(f"天维网 租房 / 招聘 信息抓取工具 v{self.core.__version__}")
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        W, H = min(int(1280 * ui), int(sw * 0.94)), min(int(820 * ui), int(sh * 0.86))
        root.geometry(f"{W}x{H}+{max(0, (sw - W) // 2)}+{max(0, (sh - H) // 3)}")
        root.minsize(int(820 * ui), int(440 * ui))
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        line = tkfont.nametofont("TkDefaultFont").metrics("linespace")
        style.configure("Treeview", rowheight=int(line * 1.45))
        style.configure("Bad.TEntry", foreground="#C00000", fieldbackground="#FFE8E8")
        style.configure("Accent.TButton", font=(tkfont.nametofont("TkDefaultFont").actual("family"), 10, "bold"))

    def _build_toolbar(self):
        bar = ttk.Frame(self.root, padding=(8, 8, 8, 4))
        bar.pack(fill="x")
        ttk.Label(bar, text="最近天数:").pack(side="left")
        self.days_var = tk.IntVar(value=10)
        ttk.Spinbox(bar, from_=1, to=60, width=4, textvariable=self.days_var).pack(side="left", padx=(2, 10))
        self.detail_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="读取帖子详情 (租金 / 薪金等)", variable=self.detail_var).pack(side="left")
        self.btn_fetch = ttk.Button(bar, text="开始抓取 (F5)", style="Accent.TButton", command=self.start_fetch)
        self.btn_fetch.pack(side="left", padx=(12, 4))
        self.btn_stop = ttk.Button(bar, text="停止", state="disabled", command=self.stop_task)
        self.btn_stop.pack(side="left", padx=4)
        self.btn_export = ttk.Button(bar, text="导出 Excel (Ctrl+E)", command=self.export)
        self.btn_export.pack(side="left", padx=4)
        self.pbar = ttk.Progressbar(bar, length=int(180 * self.ui), mode="determinate")
        self.pbar.pack(side="left", padx=10)
        self.status = tk.StringVar(value="就绪")
        ttk.Label(bar, textvariable=self.status, foreground="#555").pack(side="left", fill="x", expand=True)

    def _build_home(self, cfg):
        box = ttk.LabelFrame(self.root, text="家庭住址 / 距离（租房和招聘通用）", padding=(8, 4))
        box.pack(fill="x", padx=8)
        self.home_var = tk.StringVar(value=cfg.get("home_text", ""))
        self.maxdist_var = tk.StringVar(value=cfg.get("max_distance", ""))
        self.sortdist_var = tk.BooleanVar(value=bool(cfg.get("sort_distance", True)))
        self.online_var = tk.BooleanVar(value=bool(cfg.get("online_geocode", True)))
        self.btn_here = ttk.Button(box, text="◎ 一键定位当前位置", command=self.locate_here)
        self.btn_here.pack(side="left", padx=(0, 10))
        ttk.Label(box, text="或输入地址:").pack(side="left")
        e = ttk.Entry(box, width=32, textvariable=self.home_var)
        e.pack(side="left")
        e.bind("<Return>", lambda ev: self.locate_home())
        self.btn_home = ttk.Button(box, text="按地址定位", command=self.locate_home)
        self.btn_home.pack(side="left", padx=4)
        ttk.Label(box, text="   最大距离 km:").pack(side="left")
        self.maxdist_entry = ttk.Entry(box, width=6, textvariable=self.maxdist_var)
        self.maxdist_entry.pack(side="left")
        self.maxdist_var.trace_add("write", lambda *a: self._on_maxdist())
        ttk.Checkbutton(box, text="按距离排序 (近→远)", variable=self.sortdist_var,
                        command=self._on_sortdist).pack(side="left", padx=10)
        ttk.Checkbutton(box, text="联网解析帖子地址 (更准)", variable=self.online_var,
                        command=self._on_online_toggle).pack(side="left")
        self.home_status = tk.StringVar()
        ttk.Label(self.root, textvariable=self.home_status, foreground="#2E6E2E",
                  padding=(16, 2)).pack(fill="x")

    def _build_statusbar(self):
        bar = ttk.Frame(self.root, padding=(10, 4))
        bar.pack(side="bottom", fill="x")
        self.count_var = tk.StringVar()
        ttk.Label(bar, textvariable=self.count_var).pack(side="left")
        ttk.Label(bar, text="双击或回车打开帖子 · Ctrl+C 复制 · 点表头排序 · 灰色行 = 只知道大区, 距离为粗略估算",
                  foreground="#777").pack(side="right")

    # ---- 当前页 / 计数 ----
    def current_tab(self) -> CategoryTab:
        sel = self.notebook.select()
        return next((t for t in self.tabs.values() if str(t.frame) == sel), self.tabs["rent"])

    def _on_tab_changed(self):
        tab = self.current_tab()
        self.btn_fetch.configure(text=f"抓取{tab.cfg['name']} (F5)")
        self.update_count()

    def update_count(self):
        if not hasattr(self, "tabs"):
            return
        tab = self.current_tab()
        if not tab.items:
            self.count_var.set(f"{tab.cfg['name']}: 尚未抓取, 点「抓取{tab.cfg['name']}」开始")
            return
        extra = ""
        if self.home:
            n = sum(1 for it in tab.shown if it.distance is not None)
            extra = f" (其中 {n} 条有预估距离)"
        self.count_var.set(f"{tab.cfg['name']}: 共 {len(tab.items)} 条, 显示 {len(tab.shown)} 条{extra}")

    def refresh_all(self):
        for t in self.tabs.values():
            t.refresh()

    # ---- 距离设置 ----
    def max_distance(self):
        s = self.maxdist_var.get().strip()
        if not s:
            self.maxdist_entry.configure(style="TEntry")
            return None
        try:
            v = float(s)
            self.maxdist_entry.configure(style="TEntry")
            return v
        except ValueError:
            self.maxdist_entry.configure(style="Bad.TEntry")
            return None

    def _on_maxdist(self):
        self.save_prefs()
        for t in self.tabs.values():
            t.refresh_soon()

    def _on_sortdist(self):
        for t in self.tabs.values():
            t.clear_manual_sort()
        self.save_prefs()
        self.refresh_all()

    def show_home_status(self):
        h = self.home
        if h:
            self.home_status.set(f"✔ 住址: {h[2][:90]}  ({h[0]:.5f}, {h[1]:.5f})   —— 距离为直线距离; "
                                 f"灰色行只知道大区, 距离带「≈」")
        else:
            self.home_status.set("未设置家庭住址: 在家时点「◎ 一键定位当前位置」, 或输入街道地址 / 街区名后点「按地址定位」")

    def save_prefs(self):
        c = geo.load_config()
        c.update(home_text=self.home_var.get().strip(), max_distance=self.maxdist_var.get().strip(),
                 sort_distance=self.sortdist_var.get(), online_geocode=self.online_var.get())
        if self.home:
            c.update(home_lat=self.home[0], home_lon=self.home[1], home_desc=self.home[2])
        else:
            for k in ("home_lat", "home_lon", "home_desc"):
                c.pop(k, None)
        try:
            geo.save_config(c)
        except OSError:
            self.core.log.exception("保存配置失败")

    # ---- 后台任务: 一次只运行一个 ----
    def begin_task(self, name: str, stoppable: bool = False) -> threading.Event | None:
        if self.task:
            self.status.set(f"请等待「{self.task}」完成")
            return None
        self.task = name
        self.stop_event = threading.Event()
        for b in (self.btn_fetch, self.btn_here, self.btn_home):
            b.configure(state="disabled")
        self.btn_stop.configure(state="normal" if stoppable else "disabled")
        self.pbar.configure(mode="indeterminate", value=0)
        self.pbar.start(12)
        self.status.set(f"{name} ...")
        return self.stop_event

    def end_task(self):
        self.task = None
        for b in (self.btn_fetch, self.btn_here, self.btn_home):
            b.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.pbar.stop()
        self.pbar.configure(mode="determinate", value=0)

    def stop_task(self):
        if self.stop_event:
            self.stop_event.set()
            self.status.set("正在停止 ...")

    def _run(self, name, fn):
        """在后台线程运行 fn(); 异常记日志并回到主线程提示。"""
        def wrapper():
            try:
                fn()
            except Exception as e:  # noqa: BLE001
                self.core.log.exception("%s 失败", name)
                self.msgq.put(("error", (name, f"{type(e).__name__}: {e}")))
        threading.Thread(target=wrapper, name=name, daemon=True).start()

    # ---- 抓取 ----
    def start_fetch(self):
        tab = self.current_tab()
        try:
            days = max(1, min(60, int(self.days_var.get())))
        except (tk.TclError, ValueError):
            days = 10
        self.days_var.set(days)
        with_detail, online = self.detail_var.get(), self.online_var.get()   # 在主线程读取
        stop = self.begin_task(f"抓取{tab.cfg['name']}", stoppable=True)
        if stop is None:
            return
        cat = tab.cat

        def work():
            sc = self.core.Scraper(log=lambda m: self.msgq.put(("log", m)), stop_event=stop)
            items = sc.scrape(cat, days, with_detail=with_detail,
                              progress=lambda d, t: self.msgq.put(("prog", (d, t))),
                              online_geocode=online)
            self.msgq.put(("fetched", (cat, items, stop.is_set())))
        self._run(self.task, work)

    # ---- 住址 ----
    def locate_home(self):
        text = self.home_var.get().strip()
        if not text:
            self.home = None
            self.save_prefs()
            self.show_home_status()
            self.refresh_all()
            return
        online = self.online_var.get()
        if self.begin_task("定位住址") is None:
            return
        self.home_status.set("正在定位住址 ...")

        def work():
            r = geo.resolve_home(text, geo.get_geocoder() if online else None)
            self.msgq.put(("home", (text, r)))
        self._run(self.task, work)

    def locate_here(self):
        online = self.online_var.get()
        if self.begin_task("一键定位") is None:
            return
        self.home_status.set("正在通过 Windows 定位服务获取当前位置 (约 5-15 秒) ...")

        def work():
            try:
                lat, lon, acc = geo.locate_device()
            except geo.LocateError as e:
                self.msgq.put(("here_err", str(e)))
                return
            addr = ""
            if online:   # 关闭联网解析时不把坐标发给地图服务
                try:
                    addr = geo.get_geocoder().reverse(lat, lon)
                except Exception:  # noqa: BLE001
                    self.core.log.exception("反查地址失败")
            self.msgq.put(("here", (lat, lon, acc, addr)))
        self._run(self.task, work)

    def _on_online_toggle(self):
        self.save_prefs()
        items = self.tabs["rent"].items + self.tabs["job"].items
        if not self.online_var.get() or not items:
            return
        stop = self.begin_task("联网解析帖子地址", stoppable=True)
        if stop is None:
            return

        def work():
            self.core.locate_items(items, True, log=lambda m: self.msgq.put(("log", m)),
                                   progress=lambda d, t: self.msgq.put(("prog", (d, t))), stop_event=stop)
            self.msgq.put(("relocated", None))
        self._run(self.task, work)

    # ---- 导出 ----
    def export(self):
        tab = self.current_tab()
        if not tab.shown:
            messagebox.showinfo("提示", "当前页没有可导出的数据")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile=self.core.default_filename(tab.cat),
                                            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        try:
            self.core.export_excel(tab.shown, tab.cat, path)
        except PermissionError:
            messagebox.showerror("导出失败", "文件被占用, 请先关闭 Excel 中已打开的同名文件")
            return
        except Exception as e:  # noqa: BLE001
            self.core.log.exception("导出 Excel 失败: %s", path)
            messagebox.showerror("导出失败", f"{type(e).__name__}: {e}\n\n详细信息已写入日志:\n{self.log_path}")
            return
        if messagebox.askyesno("导出完成", f"已导出 {len(tab.shown)} 条到\n{path}\n\n是否立即打开?"):
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except OSError:
                pass

    # ---- 后台消息 ----
    def _poll(self):
        try:
            while True:
                kind, data = self.msgq.get_nowait()
                self._handle(kind, data)
        except queue.Empty:
            pass
        self._poll_id = self.root.after(150, self._poll)

    def _handle(self, kind, data):
        if kind == "log":
            self.status.set(data)
        elif kind == "prog":
            d, t = data
            if t:
                if str(self.pbar["mode"]) != "determinate":
                    self.pbar.stop()
                    self.pbar.configure(mode="determinate")
                self.pbar.configure(maximum=t, value=d)
                self.status.set(f"{self.task or ''} {d}/{t}")
        elif kind == "fetched":
            cat, items, stopped = data
            self.end_task()
            tab = self.tabs[cat]
            if items or not tab.items:
                tab.set_items(items)
            name = tab.cfg["name"]
            self.status.set(f"已停止, 保留 {len(items)} 条{name}" if stopped else f"完成, 抓取 {len(items)} 条{name}")
            self.update_count()
        elif kind == "home":
            self.end_task()
            text, r = data
            if r:
                self.home = r
                self.save_prefs()
                self.refresh_all()
            else:
                messagebox.showwarning("定位失败", "无法解析该住址, 原来的住址保持不变。\n可以试试: 更完整的街道地址、"
                                                  "街区名 (如 Mount Albert), 或直接输入 '纬度,经度'")
                self.home_var.set(geo.load_config().get("home_text", ""))
            self.show_home_status()
        elif kind == "here":
            self.end_task()
            lat, lon, acc, addr = data
            note = f"当前位置 (精度约 ±{acc} m)"
            if acc > 2000:
                note += " ⚠ 精度较低, 建议连接 Wi-Fi 后重试或手动输入地址"
            if not (-47.6 <= lat <= -34.0 and 166.0 <= lon <= 179.0):
                note += " ⚠ 当前位置不在新西兰"
            self.home = (lat, lon, f"{note}: {addr or f'{lat:.5f}, {lon:.5f}'}")
            self.home_var.set(addr or f"{lat:.5f}, {lon:.5f}")
            self.save_prefs()
            self.show_home_status()
            self.refresh_all()
        elif kind == "here_err":
            self.end_task()
            self.show_home_status()
            messagebox.showwarning("一键定位失败", data)
        elif kind == "relocated":
            self.end_task()
            self.status.set("帖子位置已更新")
            self.refresh_all()
        elif kind == "error":
            name, msg = data
            self.end_task()
            self.show_home_status()
            self.status.set(f"{name}失败")
            messagebox.showerror(f"{name}失败", f"{msg}\n\n详细信息已写入日志:\n{self.log_path}")

    # ---- 错误 / 关闭 ----
    def _on_callback_error(self, exc_type, exc, tb):
        self.core.log.error("界面操作出错", exc_info=(exc_type, exc, tb))
        messagebox.showerror("程序出错", f"{exc_type.__name__}: {exc}\n\n详细信息已写入日志:\n{self.log_path}")

    def _on_destroy(self, e):
        if e.widget is not self.root:
            return
        if self._poll_id:                           # 关窗后不再触发轮询和延迟刷新
            self.root.after_cancel(self._poll_id)
            self._poll_id = None
        for t in getattr(self, "tabs", {}).values():
            if t._after:
                t.frame.after_cancel(t._after)
                t._after = None

    def on_close(self):
        if self.stop_event:
            self.stop_event.set()
        if geo._geocoder is not None:
            geo._geocoder.save()          # 每 10 条才自动保存一次, 关窗时补存
        self.save_prefs()
        self.root.destroy()


def run(core):
    enable_dpi_awareness()          # 必须在创建窗口之前
    root = tk.Tk()
    App(core, root)
    root.mainloop()
