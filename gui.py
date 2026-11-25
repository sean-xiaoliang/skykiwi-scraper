# -*- coding: utf-8 -*-
"""
图形界面 (Tkinter)。由 skykiwi_scraper.run_gui() 调用; core 为主模块 (抓取 / 筛选 / 导出)。

结构:
  App            顶部工具栏 (刷新 / 停止 / 导出 / 设置)、家庭住址栏、两个标签页、底部状态栏; 后台任务一次只跑一个
  CategoryTab    一个类别 (租房 / 招聘) 的筛选条件、结果表、排序和详情区, 两页互不影响; 招聘页另有简历和邮件申请
  MultiSelect    可多选的下拉框: 勾选时不收起, 点外面或按 Esc 关闭
  SettingsDialog 软件设置: 显示天数、是否读取详情、邮件方式和模板、列宽、清空缓存
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
import mail
import resume
import store

DEBOUNCE_MS = 300      # 输入框停止输入后多久刷新结果
DEFAULT_DAYS = 10      # 默认显示最近 10 天

# 表格只显示关键列 (其余字段在下方详情区和导出的 Excel 里): (字段, 表头, 宽度单位, 对齐)
# 所有列都不自动拉伸: 拖动表头边缘调整的宽度会被记住
VIEW_COLUMNS = {
    "rent": [("post_date", "日期", 9, "w"), ("region", "区域", 7, "w"), ("title", "标题", 26, "w"),
             ("price", "周租金", 7, "e"), ("distance", "距离 km", 9, "e"), ("rent_type", "出租类型", 12, "w"),
             ("house_type", "房屋类型", 12, "w"), ("loc_desc", "定位依据", 20, "w")],
    "job": [("post_date", "日期", 9, "w"), ("region", "岗位类别", 9, "w"), ("title", "标题", 24, "w"),
            ("position", "职位", 12, "w"), ("nature", "性质", 6, "w"), ("salary_num", "时薪", 6, "e"),
            ("mail", "邮箱", 5, "center"), ("distance", "距离 km", 9, "e"), ("loc_desc", "定位依据", 20, "w")],
}
MATCH_COLUMN = ("match", "匹配", 6, "e")     # 上传简历后加到招聘页
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

        # ---- 招聘页: 简历分析 + 邮件申请 ----
        if not self.rent:
            act = ttk.Frame(self.frame, padding=(0, 6, 0, 0))
            act.pack(fill="x")
            ttk.Button(act, text="上传简历…", command=app.choose_resume).pack(side="left")
            self.resume_label = tk.StringVar(value="未上传简历 (上传后按匹配度给岗位打分, 只在本机分析)")
            ttk.Label(act, textvariable=self.resume_label, foreground="#555").pack(side="left", padx=8)
            ttk.Button(act, text="查看分析", command=app.show_resume_report).pack(side="left")
            self.btn_mail = ttk.Button(act, text="✉ 邮件申请选中岗位", state="disabled", command=self.send_mail)
            self.btn_mail.pack(side="right")

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
                              background="#FAFAFA", state="disabled", font="TkDefaultFont")
        dsb = ttk.Scrollbar(df, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.pack(side="left", fill="both", expand=True)
        dsb.pack(side="right", fill="y")
        pw.add(tf, weight=4)
        pw.add(df, weight=1)

        self.match_enabled = False
        self.setup_columns()
        self.tree.tag_configure("coarse", foreground="#8A8A8A")
        self.tree.bind("<ButtonRelease-1>", self._maybe_save_widths, add="+")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", self.open_selected)
        self.tree.bind("<Return>", self.open_selected)
        self.tree.bind("<Control-c>", self.copy_selected)
        self.set_detail("抓取后在这里显示选中帖子的详细信息。双击或按回车在浏览器打开原帖。")

    # ---- 列 ----
    def view_columns(self):
        cols = list(VIEW_COLUMNS[self.cat])
        if self.match_enabled:
            cols.insert(3, MATCH_COLUMN)
        return cols

    def setup_columns(self):
        ui = self.app.ui
        saved = (geo.load_config().get("col_widths") or {}).get(self.cat, {})
        self.columns = self.view_columns()
        self.tree["columns"] = [c[0] for c in self.columns]
        for key, head, w, anchor in self.columns:
            width = saved.get(key, w * 8)          # 保存的是 100% 缩放下的像素
            self.tree.heading(key, text=head, anchor=anchor, command=lambda kk=key: self.sort_by(kk))
            self.tree.column(key, width=int(width * ui), minwidth=int(30 * ui), anchor=anchor, stretch=False)

    def _maybe_save_widths(self, e):
        if self.tree.identify_region(e.x, e.y) == "separator":
            self.save_column_widths()

    def save_column_widths(self):
        cfg = geo.load_config()
        widths = cfg.setdefault("col_widths", {}).setdefault(self.cat, {})
        for key, *_ in self.columns:
            widths[key] = round(self.tree.column(key, "width") / self.app.ui)
        geo.save_config(cfg)

    def reset_column_widths(self):
        cfg = geo.load_config()
        (cfg.get("col_widths") or {}).pop(self.cat, None)
        geo.save_config(cfg)
        self.setup_columns()
        self.refresh()

    def enable_match_column(self, on: bool):
        if on != self.match_enabled:
            self.match_enabled = on
            self.setup_columns()

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
        if not self.rent:
            prof = self.app.profile
            for it in self.items:
                it.mail = "✉" if mail.job_emails(it) else ""
                it.match = resume.match(prof, it) if prof else None
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

    def selected(self):
        sel = self.tree.selection()
        return self._find(sel[0]) if sel else None

    def on_select(self, _e=None):
        it = self.selected()
        if not self.rent:
            self.btn_mail.configure(state="normal" if it and mail.job_emails(it) else "disabled")
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
        emails = mail.job_emails(it) if not self.rent else []
        lines.append(f"电话: {it.phone or '-'}    微信: {it.wechat or '-'}"
                     + (f"    邮箱: {', '.join(emails)}" if emails else ""))
        if not self.rent and self.app.profile:
            lines.append(f"简历匹配度: {resume.match(self.app.profile, it)} / 100")
        for k, v in it.extra.items():
            lines.append(f"{k}: {v}")
        lines += ["", it.content]
        self.set_detail("\n".join(lines))

    def open_selected(self, _e=None):
        for iid in self.tree.selection()[:5]:
            it = self._find(iid)
            if it:
                webbrowser.open(it.url)

    def send_mail(self):
        it = self.selected()
        emails = mail.job_emails(it) if it else []
        if not emails:
            messagebox.showinfo("邮件申请", "选中的帖子没有留邮箱, 请用帖子里的电话或微信联系")
            return
        cfg = geo.load_config()
        client = mail.detect_mail_client()
        method = mail.choose_method(cfg.get("mail_method", "auto"), client, cfg.get("webmail", "gmail"))
        summary = self.app.profile.summary() if self.app.profile else ""
        subject = mail.fill(cfg.get("mail_subject") or mail.DEFAULT_SUBJECT, it, summary)
        body = mail.fill(cfg.get("mail_body") or mail.DEFAULT_BODY, it, summary)
        try:
            mail.compose(emails[0], subject, body, method)
        except OSError:
            method = cfg.get("webmail", "gmail")              # 客户端打不开时退回网页邮箱
            mail.compose(emails[0], subject, body, method)
        where = mail.client_name(client) if method == "client" else mail.WEBMAIL[method]
        msg = f"已在{where}打开写信窗口 (收件人 {emails[0]})"
        rp = self.app.resume_path
        if rp and os.path.exists(rp) and cfg.get("reveal_resume", True):
            mail.reveal_file(rp)
            msg += "; 已在资源管理器中选中简历, 请拖进邮件作为附件"
        if len(emails) > 1:
            msg += f"; 帖子里还有其他邮箱: {', '.join(emails[1:])}"
        self.app.status.set(msg)

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
        self.cache = store.PostCache()
        self.profile = None                 # 简历分析结果
        self.resume_path = ""
        self.settings_dialog = None

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
        self._load_resume_silently(cfg.get("resume_path", ""))
        self.load_cached()
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
        self.btn_fetch = ttk.Button(bar, text="刷新 (F5)", style="Accent.TButton", command=self.start_fetch)
        self.btn_fetch.pack(side="left", padx=(0, 4))
        self.btn_stop = ttk.Button(bar, text="停止", state="disabled", command=self.stop_task)
        self.btn_stop.pack(side="left", padx=4)
        self.btn_export = ttk.Button(bar, text="导出 Excel (Ctrl+E)", command=self.export)
        self.btn_export.pack(side="left", padx=4)
        ttk.Button(bar, text="⚙ 设置", command=self.open_settings).pack(side="left", padx=4)
        self.days_label = tk.StringVar()
        ttk.Label(bar, textvariable=self.days_label, foreground="#2E6E2E").pack(side="left", padx=(8, 0))
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
        self.btn_fetch.configure(text=f"刷新{tab.cfg['name']} (F5)")
        self.update_count()

    # ---- 设置 / 缓存 ----
    def display_days(self) -> int:
        try:
            return max(1, min(60, int(geo.load_config().get("display_days", DEFAULT_DAYS))))
        except (TypeError, ValueError):
            return DEFAULT_DAYS

    def read_detail(self) -> bool:
        return bool(geo.load_config().get("read_detail", True))

    def load_cached(self):
        days = self.display_days()
        self.days_label.set(f"显示最近 {days} 天 (在设置中修改)")
        for cat, tab in self.tabs.items():
            tab.set_items(self.cache.load(cat, days))
        last = [self.cache.get_meta(f"last_refresh_{c}") for c in self.tabs]
        if any(last):
            when = max(x for x in last if x).replace("T", " ")[:16]
            self.status.set(f"已载入本地缓存 (最近刷新 {when}); 点「刷新」只抓新帖")

    def open_settings(self):
        if self.settings_dialog is not None and self.settings_dialog.winfo_exists():
            self.settings_dialog.lift()
            return
        self.settings_dialog = SettingsDialog(self)

    # ---- 简历 ----
    def choose_resume(self):
        path = filedialog.askopenfilename(title="选择简历", filetypes=[
            ("简历 (PDF / Word / TXT)", "*.pdf *.docx *.txt"), ("所有文件", "*.*")])
        if path:
            self.load_resume(path)

    def load_resume(self, path: str, show_report: bool = True) -> bool:
        try:
            prof = resume.analyze(resume.extract_text(path))
        except resume.ResumeError as e:
            messagebox.showwarning("简历", str(e))
            return False
        self.profile, self.resume_path = prof, path
        cfg = geo.load_config()
        cfg["resume_path"] = path
        geo.save_config(cfg)
        job = self.tabs["job"]
        if prof.keywords:
            job.resume_label.set(f"简历: {os.path.basename(path)} · 方向: {'、'.join(prof.categories[:3])}")
        else:
            job.resume_label.set(f"简历: {os.path.basename(path)} · 没有识别出技能关键词, 匹配度可能不准")
        job.enable_match_column(True)
        if job.sort_key is None:
            job.sort_key, job.sort_rev = "match", True      # 上传后按匹配度从高到低
        job.refresh()
        if show_report:
            self.show_resume_report()
        return True

    def _load_resume_silently(self, path):
        if path and os.path.exists(path):
            try:
                self.profile = resume.analyze(resume.extract_text(path))
                self.resume_path = path
            except resume.ResumeError:
                self.core.log.warning("启动时读取简历失败: %s", path)

    def show_resume_report(self):
        prof = self.profile
        if not prof:
            messagebox.showinfo("简历分析", "还没有上传简历")
            return
        job = self.tabs["job"]
        top = sorted((it for it in job.items if getattr(it, "match", None)), key=lambda it: -it.match)[:10]
        lines = [f"简历文件: {self.resume_path}", "（分析只在本机进行, 不会上传）", ""]
        if prof.cat_hits:
            lines.append("技能方向:")
            lines += [f"  · {c}: {'、'.join(prof.cat_hits[c][:8])}" for c in prof.categories]
        else:
            lines.append("技能方向: 没有识别出来 (可以在简历里写明职位, 如 厨师 / 清洁 / 会计)")
        lines += ["", f"语言: {'、'.join(prof.languages) or '未写明'}",
                  f"证照: {'、'.join(prof.licences) or '未写明'}",
                  f"签证: {prof.visa or '未写明'}",
                  f"工作年限: {f'{prof.years} 年' if prof.years else '未写明'}",
                  "", "匹配度计算: 岗位类别吻合 +35, 每个技能关键词 +15 (最多 45), "
                      "满足 / 不满足 英语和驾照要求 ±10",
                  "", f"当前招聘页匹配度最高的 {len(top)} 个岗位:"]
        lines += [f"  {it.match:>3}  {it.title}" for it in top] or ["  (先刷新招聘页)"]
        win = tk.Toplevel(self.root)
        win.title("简历分析")
        win.transient(self.root)
        txt = tk.Text(win, width=70, height=26, wrap="word", padx=10, pady=8)
        txt.insert("1.0", "\n".join(lines))
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=6)

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
        days, with_detail, online = self.display_days(), self.read_detail(), self.online_var.get()
        stop = self.begin_task(f"刷新{tab.cfg['name']}", stoppable=True)
        if stop is None:
            return
        cat, cache = tab.cat, self.cache

        def work():
            sc = self.core.Scraper(log=lambda m: self.msgq.put(("log", m)), stop_event=stop)
            items, n_new = sc.refresh(cat, days, cache, with_detail=with_detail,
                                      progress=lambda d, t: self.msgq.put(("prog", (d, t))),
                                      online_geocode=online)
            self.msgq.put(("fetched", (cat, items, stop.is_set(), n_new)))
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

        cache = self.cache

        def work():
            self.core.locate_items(items, True, log=lambda m: self.msgq.put(("log", m)),
                                   progress=lambda d, t: self.msgq.put(("prog", (d, t))), stop_event=stop)
            cache.save(items)                   # 定位结果写回缓存
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
            cat, items, stopped, n_new = data
            self.end_task()
            tab = self.tabs[cat]
            tab.set_items(items)
            name = tab.cfg["name"]
            head = "已停止" if stopped else "刷新完成"
            self.status.set(f"{head}: 新帖 {n_new} 条, 最近 {self.display_days()} 天共 {len(items)} 条{name}")
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


class SettingsDialog(tk.Toplevel):
    def __init__(self, app: App):
        super().__init__(app.root)
        self.app = app
        self.title("设置")
        self.transient(app.root)
        self.resizable(False, False)
        cfg = geo.load_config()
        pad = {"padx": 10, "pady": 4}

        data = ttk.LabelFrame(self, text="数据", padding=8)
        data.pack(fill="x", **pad)
        ttk.Label(data, text="显示最近").grid(row=0, column=0, sticky="w")
        self.days_var = tk.IntVar(value=app.display_days())
        ttk.Spinbox(data, from_=1, to=60, width=5, textvariable=self.days_var).grid(row=0, column=1, padx=4)
        ttk.Label(data, text="天的帖子 (默认 10; 调大后点「刷新」会补抓更早的帖子)").grid(row=0, column=2, sticky="w")
        self.detail_var = tk.BooleanVar(value=bool(cfg.get("read_detail", True)))
        ttk.Checkbutton(data, text="读取帖子详情 (租金 / 薪金 / 地址 / 邮箱等, 关闭后刷新更快但信息少)",
                        variable=self.detail_var).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
        counts = "、".join(f"{t.cfg['name']} {app.cache.count(c)} 条" for c, t in app.tabs.items())
        ttk.Label(data, text=f"本地缓存: {counts} (保存最近 60 天)", foreground="#555").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(data, text="清空本地缓存", command=self.clear_cache).grid(row=2, column=2, sticky="e", pady=(6, 0))

        m = ttk.LabelFrame(self, text="邮件申请", padding=8)
        m.pack(fill="x", **pad)
        client = mail.detect_mail_client()
        ttk.Label(m, text="发送方式:").grid(row=0, column=0, sticky="w")
        self.method_var = tk.StringVar(value=mail.METHODS.get(cfg.get("mail_method", "auto"), mail.METHODS["auto"]))
        ttk.Combobox(m, width=40, state="readonly", textvariable=self.method_var,
                     values=list(mail.METHODS.values())).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(m, text=f"检测到的邮件客户端: {mail.client_name(client)}", foreground="#555").grid(
            row=1, column=1, sticky="w", padx=4)
        ttk.Label(m, text="没有客户端时用:").grid(row=2, column=0, sticky="w")
        self.webmail_var = tk.StringVar(value=mail.WEBMAIL.get(cfg.get("webmail", "gmail")))
        ttk.Combobox(m, width=40, state="readonly", textvariable=self.webmail_var,
                     values=list(mail.WEBMAIL.values())).grid(row=2, column=1, sticky="w", padx=4)
        ttk.Label(m, text="邮件主题:").grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.subject_var = tk.StringVar(value=cfg.get("mail_subject") or mail.DEFAULT_SUBJECT)
        ttk.Entry(m, width=60, textvariable=self.subject_var).grid(row=3, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(m, text="邮件正文:").grid(row=4, column=0, sticky="nw", pady=(6, 0))
        self.body_text = tk.Text(m, width=60, height=7, wrap="word")
        self.body_text.insert("1.0", cfg.get("mail_body") or mail.DEFAULT_BODY)
        self.body_text.grid(row=4, column=1, sticky="w", padx=4, pady=(6, 0))
        ttk.Label(m, text="可用占位符: {职位} {公司} {标题} {链接} {简历摘要}", foreground="#555").grid(
            row=5, column=1, sticky="w", padx=4)
        self.reveal_var = tk.BooleanVar(value=bool(cfg.get("reveal_resume", True)))
        ttk.Checkbutton(m, text="打开写信窗口时在资源管理器中选中简历 (邮件链接无法自动带附件)",
                        variable=self.reveal_var).grid(row=6, column=1, sticky="w", padx=4, pady=(4, 0))

        t = ttk.LabelFrame(self, text="表格", padding=8)
        t.pack(fill="x", **pad)
        ttk.Label(t, text="拖动表头边缘可调整列宽, 程序会记住。").pack(side="left")
        ttk.Button(t, text="恢复默认列宽", command=self.reset_widths).pack(side="right")

        b = ttk.Frame(self, padding=8)
        b.pack(fill="x")
        ttk.Button(b, text="确定", command=self.apply).pack(side="right", padx=4)
        ttk.Button(b, text="取消", command=self.destroy).pack(side="right")
        self.bind("<Escape>", lambda e: self.destroy())

    def _key(self, mapping, label, default):
        return next((k for k, v in mapping.items() if v == label), default)

    def apply(self):
        try:
            days = max(1, min(60, int(self.days_var.get())))
        except (tk.TclError, ValueError):
            days = DEFAULT_DAYS
        cfg = geo.load_config()
        old_days = cfg.get("display_days", DEFAULT_DAYS)
        cfg.update(display_days=days, read_detail=self.detail_var.get(),
                   mail_method=self._key(mail.METHODS, self.method_var.get(), "auto"),
                   webmail=self._key(mail.WEBMAIL, self.webmail_var.get(), "gmail"),
                   mail_subject=self.subject_var.get().strip() or mail.DEFAULT_SUBJECT,
                   mail_body=self.body_text.get("1.0", "end").rstrip() + "\n",
                   reveal_resume=self.reveal_var.get())
        geo.save_config(cfg)
        self.destroy()
        if days != old_days:
            self.app.load_cached()
            if days > old_days:
                self.app.status.set(f"已改为显示最近 {days} 天; 点「刷新」补抓更早的帖子")

    def clear_cache(self):
        if messagebox.askyesno("清空本地缓存", "删除所有已缓存的帖子? 下次刷新会重新抓取。", parent=self):
            self.app.cache.clear()
            self.app.load_cached()
            self.destroy()

    def reset_widths(self):
        for tab in self.app.tabs.values():
            tab.reset_column_widths()


def run(core):
    enable_dpi_awareness()          # 必须在创建窗口之前
    root = tk.Tk()
    App(core, root)
    root.mainloop()
