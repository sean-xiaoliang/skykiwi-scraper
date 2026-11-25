# -*- coding: utf-8 -*-
"""界面测试: 会短暂打开窗口 (约 1 秒/项), 不联网, 使用临时数据目录。运行: python test_gui.py"""
import os
import sys
import tempfile

os.environ["SKYKIWI_APP_DIR"] = tempfile.mkdtemp(prefix="skykiwi_test_")

import tkinter as tk  # noqa: E402
from tkinter import ttk, messagebox, filedialog  # noqa: E402

import skykiwi_scraper as k  # noqa: E402


def walk(w):
    yield w
    for c in w.winfo_children():
        yield from walk(c)


def find(root, cls, text):
    for w in walk(root):
        if isinstance(w, cls) and text in str(w.cget("text")):
            return w
    raise LookupError(text)


def with_gui(check):
    """打开主窗口, 在事件循环里执行 check(root), 然后关闭。返回 check 的结果或抛出其异常。"""
    box = {}
    orig = tk.Tk.mainloop

    def ml(self):
        def run():
            try:
                box["result"] = check(self)
            except BaseException as e:  # noqa: BLE001
                box["error"] = e
            finally:
                self.destroy()
        self.after(300, run)
        orig(self)

    tk.Tk.mainloop = ml
    try:
        k.run_gui()
    finally:
        tk.Tk.mainloop = orig
    if "error" in box:
        raise box["error"]
    return box.get("result")


# ---- 窗口变矮时导出按钮仍可见 ----
def test_export_button_visible_when_window_short():
    def check(root):
        root.geometry("1100x480")
        root.update()
        btn = find(root, ttk.Button, "导出")
        y_bottom = btn.winfo_rooty() + btn.winfo_height()
        win_bottom = root.winfo_rooty() + root.winfo_height()
        assert btn.winfo_ismapped() and btn.winfo_height() > 10, "导出按钮被挤掉"
        assert y_bottom <= win_bottom + 1, "导出按钮在窗口外"
    with_gui(check)


# ---- 高分屏: 进程声明 DPI 感知 ----
def test_process_is_dpi_aware():
    import ctypes

    def check(root):
        user32 = ctypes.windll.user32
        user32.GetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        user32.GetAwarenessFromDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        ctx = user32.GetThreadDpiAwarenessContext()
        return user32.GetAwarenessFromDpiAwarenessContext(ctx)
    awareness = with_gui(check)
    assert awareness in (1, 2), f"DPI awareness = {awareness} (0 = 不感知, 文字会被拉伸发虚)"


# ---- 导出失败要提示用户 ----
def test_export_failure_shows_error():
    shown = []
    orig_err, orig_ask, orig_exp = messagebox.showerror, filedialog.asksaveasfilename, k.export_excel
    messagebox.showerror = lambda *a, **kw: shown.append(a)
    filedialog.asksaveasfilename = lambda **kw: os.path.join(os.environ["SKYKIWI_APP_DIR"], "x.xlsx")

    def boom(*a, **kw):
        raise ValueError("模拟导出失败")
    k.export_excel = boom
    try:
        def check(root):
            # 放入一条数据, 让导出有内容
            k.Scraper.refresh = lambda self, cat, days, cache, **kw: ([
                k.Listing(tid=1, category=cat, title="t", sort="出租" if cat == "rent" else "招聘")], 1)
            find(root, ttk.Button, "刷新").invoke()
            for _ in range(40):
                root.update()
                root.after(50)
                root.update()
                if any(isinstance(w, ttk.Treeview) and w.get_children() for w in walk(root)):
                    break
            find(root, ttk.Button, "导出").invoke()
        with_gui(check)
    finally:
        messagebox.showerror, filedialog.asksaveasfilename, k.export_excel = orig_err, orig_ask, orig_exp
    assert shown and "模拟导出失败" in " ".join(map(str, shown[0])), shown


# ---- 界面回调中的异常写入日志 ----
def test_callback_exception_logged():
    shown = []
    orig_err = messagebox.showerror
    messagebox.showerror = lambda *a, **kw: shown.append(a)
    try:
        def check(root):
            try:
                raise RuntimeError("回调里的意外错误")
            except RuntimeError as e:
                root.report_callback_exception(type(e), e, e.__traceback__)
        with_gui(check)
    finally:
        messagebox.showerror = orig_err
    log = os.path.join(os.environ["SKYKIWI_APP_DIR"], "app.log")
    assert os.path.exists(log), "没有日志文件"
    with open(log, encoding="utf-8") as f:
        assert "回调里的意外错误" in f.read()
    assert shown, "没有提示用户"


# ==================== v1.1.0 ====================

def pump(root, ms=500):
    """运行事件循环 ms 毫秒 (等待防抖等延迟任务)。"""
    import time
    end = time.time() + ms / 1000
    while time.time() < end:
        root.update()
        time.sleep(0.02)


def rent_items():
    a = k.Listing(tid=1, category="rent", title="Albany 双人房", sort="出租", region="北岸", price=300.0)
    b = k.Listing(tid=2, category="rent", title="Howick 单间", sort="出租", region="东区", price=200.0)
    c = k.Listing(tid=3, category="rent", title="求租 CBD", sort="求租", region="市中心", price=250.0)
    d = k.Listing(tid=4, category="rent", title="已租", sort="出租", region="东区", price=220.0)
    return [a, b, c, d]


def rows(tab):
    return [int(i) for i in tab.tree.get_children()]


def test_tabs_keep_separate_filters():
    def check(root):
        app = root.app
        app.tabs["rent"].min_var.set("200")
        app.notebook.select(app.tabs["job"].frame)
        pump(root, 100)
        assert app.tabs["job"].min_var.get() == "", "招聘页沿用了租房页的价格条件"
        assert app.tabs["rent"].min_var.get() == "200"
    with_gui(check)


def test_filters_apply_immediately_without_button():
    def check(root):
        assert not [w for w in walk(root) if isinstance(w, ttk.Button) and "应用筛选" in str(w.cget("text"))]
        tab = root.app.tabs["rent"]
        tab.set_items(rent_items())
        tab.kw_var.set("Howick")
        pump(root, 600)
        assert rows(tab) == [2], rows(tab)
    with_gui(check)


def test_defaults_rental_only_and_exclude_taken():
    def check(root):
        tab = root.app.tabs["rent"]
        tab.set_items(rent_items())
        pump(root, 100)
        assert sorted(rows(tab)) == [1, 2], rows(tab)     # 不含求租, 不含"已租"
        assert root.app.tabs["job"].sort_vars["招聘"].get() is True
    with_gui(check)


def test_invalid_number_marked():
    def check(root):
        tab = root.app.tabs["rent"]
        tab.min_var.set("abc")
        pump(root, 600)
        assert tab.min_entry.cget("style") == "Bad.TEntry", tab.min_entry.cget("style")
        tab.min_var.set("150")
        pump(root, 600)
        assert tab.min_entry.cget("style") != "Bad.TEntry"
    with_gui(check)


def test_job_key_columns_visible():
    def check(root):
        app = root.app
        app.notebook.select(app.tabs["job"].frame)
        tab = app.tabs["job"]
        tab.set_items([k.Listing(tid=1, category="job", title="招聘厨房帮工", sort="招聘", nature="兼职")])
        pump(root, 200)
        cols = list(tab.tree["columns"])
        assert "nature" in cols and "salary_num" in cols, cols
        x = 0
        for c in cols:
            w = tab.tree.column(c, "width")
            if c in ("nature", "salary_num"):
                assert x + w <= tab.tree.winfo_width() + 2, f"{c} 列在可视区域外"
            x += w
        assert tab.tree.column("title", "width") >= 180, tab.tree.column("title", "width")
    with_gui(check)


def test_multiselect_stays_open_while_choosing():
    def check(root):
        ms = root.app.tabs["rent"].type_select
        ms.open()
        pump(root, 100)
        boxes = [w for w in walk(ms.popup) if isinstance(w, ttk.Checkbutton)]
        boxes[1].invoke()
        boxes[2].invoke()
        pump(root, 100)
        assert ms.popup is not None and ms.popup.winfo_exists(), "勾选一项后下拉框就关了"
        assert len(ms.get()) == 2, ms.get()
        ms.close()
    with_gui(check)


def test_sort_indicator_and_persistence():
    def check(root):
        tab = root.app.tabs["rent"]
        root.app.sortdist_var.set(False)
        tab.set_items(rent_items())
        tab.sort_by("price")
        pump(root, 100)
        assert "▲" in tab.tree.heading("price", "text"), tab.tree.heading("price", "text")
        assert rows(tab) == [2, 1], rows(tab)
        tab.ex_var.set("不存在的词")   # 触发一次刷新, 不改变结果
        pump(root, 600)
        assert rows(tab) == [2, 1], "刷新后排序丢失"
    with_gui(check)


def test_detail_is_readonly():
    def check(root):
        assert str(root.app.tabs["rent"].detail.cget("state")) == "disabled"
    with_gui(check)


def test_keyboard_shortcuts_bound():
    def check(root):
        assert root.bind("<F5>"), "F5 未绑定"
        assert root.bind("<Control-e>"), "Ctrl+E 未绑定"
    with_gui(check)


# ==================== v1.2.0 ====================
import geo  # noqa: E402
import store  # noqa: E402
import mail  # noqa: E402


def test_title_column_narrow_and_all_resizable():
    def check(root):
        app = root.app
        for tab in app.tabs.values():
            assert tab.tree.column("title", "width") <= int(260 * app.ui), tab.tree.column("title", "width")
            for c in tab.tree["columns"]:
                assert not tab.tree.column(c, "stretch"), f"{c} 会自动拉伸, 用户调好的宽度会被改掉"
    with_gui(check)


def test_column_widths_remembered():
    def first(root):
        tab = root.app.tabs["job"]
        tab.tree.column("title", width=333)
        tab.save_column_widths()
    with_gui(first)

    def second(root):
        tab = root.app.tabs["job"]
        w = tab.tree.column("title", "width")
        tab.reset_column_widths()          # 清理, 不影响其他测试
        assert abs(w - 333) <= 2, w
    with_gui(second)


def test_settings_dialog_changes_days():
    def check(root):
        app = root.app
        app.open_settings()
        dlg = app.settings_dialog
        dlg.days_var.set(15)
        dlg.apply()
        assert geo.load_config()["display_days"] == 15
        assert app.display_days() == 15
        assert not dlg.winfo_exists()
    with_gui(check)


def test_cached_posts_shown_on_startup():
    store.PostCache().save([k.Listing(tid=777, category="rent", title="缓存里的房源", sort="出租",
                                      post_date=__import__("datetime").date.today().isoformat())])

    def check(root):
        assert "777" in root.app.tabs["rent"].tree.get_children()
    with_gui(check)


def test_fetch_uses_incremental_refresh():
    called = []
    orig = k.Scraper.refresh

    def fake(self, cat, days, cache, **kw):
        called.append((cat, days))
        return [], 0
    k.Scraper.refresh = fake
    try:
        def check(root):
            root.app.start_fetch()
            pump(root, 800)
        with_gui(check)
    finally:
        k.Scraper.refresh = orig
    assert called and called[0][0] == "rent", called


def job_items():
    a = k.Listing(tid=1, category="job", title="中餐厅招厨师", sort="招聘", region="服务人员", position="厨师",
                  company="好吃餐厅", email="hr@food.co.nz", content="需要炒锅经验", url="https://x/1")
    b = k.Listing(tid=2, category="job", title="招清洁工", sort="招聘", region="清洁/家政", position="清洁",
                  content="家庭清洁, 电话联系", url="https://x/2")
    return [a, b]


def test_email_button_follows_selection():
    def check(root):
        app = root.app
        app.notebook.select(app.tabs["job"].frame)
        tab = app.tabs["job"]
        tab.set_items(job_items())
        tab.tree.selection_set("1")
        pump(root, 100)
        assert str(tab.btn_mail.cget("state")) == "normal"
        tab.tree.selection_set("2")
        pump(root, 100)
        assert str(tab.btn_mail.cget("state")) == "disabled"
    with_gui(check)


def test_email_compose_uses_webmail_when_no_client():
    sent, orig = [], (mail.compose, mail.detect_mail_client, mail.reveal_file)
    mail.compose = lambda to, subject, body, method: sent.append((to, subject, body, method))
    mail.detect_mail_client = lambda: None
    mail.reveal_file = lambda p: None
    try:
        def check(root):
            app = root.app
            tab = app.tabs["job"]
            tab.set_items(job_items())
            tab.tree.selection_set("1")
            pump(root, 100)
            tab.send_mail()
        with_gui(check)
    finally:
        mail.compose, mail.detect_mail_client, mail.reveal_file = orig
    assert sent, "没有打开写信窗口"
    to, subject, body, method = sent[0]
    assert to == "hr@food.co.nz" and "厨师" in subject and method == "gmail", sent[0]
    assert "https://x/1" in body


def test_resume_scores_jobs():
    import tempfile as tf
    p = os.path.join(tf.mkdtemp(), "cv.txt")
    with open(p, "w", encoding="utf-8") as f:
        f.write("厨师 5 年经验，熟悉炒锅。英语流利，有驾照。")

    def check(root):
        app = root.app
        tab = app.tabs["job"]
        tab.set_items(job_items())
        app.load_resume(p)
        pump(root, 100)
        assert "match" in tab.tree["columns"]
        col = list(tab.tree["columns"]).index("match")
        cook = int(tab.tree.item("1", "values")[col])
        clean = int(tab.tree.item("2", "values")[col])
        assert cook > clean, (cook, clean)
        assert geo.load_config()["resume_path"] == p
    with_gui(check)


if __name__ == "__main__":
    fails = 0
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_")]
    for n, f in tests:
        try:
            f()
            print("PASS", n)
        except Exception as e:  # noqa: BLE001
            fails += 1
            print("FAIL", n, "->", type(e).__name__, e)
    print(f"\n{len(tests) - fails}/{len(tests)} passed")
    sys.exit(1 if fails else 0)
