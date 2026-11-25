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
            k.Scraper.scrape = lambda self, cat, days=10, **kw: [k.Listing(tid=1, category=cat, title="t")]
            find(root, ttk.Button, "开始抓取").invoke()
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
