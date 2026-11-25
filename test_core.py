# -*- coding: utf-8 -*-
"""核心逻辑测试 (不联网、不打开窗口)。运行: python test_core.py"""
import os
import sys
import tempfile

# 测试使用独立的数据目录, 不读写用户的 %APPDATA%\SkykiwiScraper
os.environ["SKYKIWI_APP_DIR"] = tempfile.mkdtemp(prefix="skykiwi_test_")

import geo  # noqa: E402
import skykiwi_scraper as k  # noqa: E402


# ---- 测试隔离 ----
def test_app_dir_overridable_by_env():
    assert geo.APP_DIR == os.environ["SKYKIWI_APP_DIR"], geo.APP_DIR


# ---- 租金识别 ----
def test_rent_picks_plausible_number():
    assert k.parse_rent("380 NZD") == 380
    assert k.parse_rent("2房 $500") == 500
    assert k.parse_rent("$1,200/周") == 1200


def test_rent_rejects_placeholders():
    assert k.parse_rent("88888") is None
    assert k.parse_rent("面议") is None
    assert k.parse_rent("") is None


# ---- 时薪识别: 按单位换算, 不合理的丢弃 ----
def approx(a, b, tol=0.6):
    return a is not None and abs(a - b) <= tol


def test_wage_hourly():
    assert approx(k.parse_wage("$25/小时"), 25)
    assert approx(k.parse_wage("时薪28起"), 28)
    assert approx(k.parse_wage("25-30/hr"), 25)


def test_wage_annual_weekly_monthly():
    assert approx(k.parse_wage("年薪 60000"), 60000 / 2080)
    assert approx(k.parse_wage("$60,000 per annum"), 60000 / 2080)
    assert approx(k.parse_wage("40k-50k"), 40000 / 2080)
    assert approx(k.parse_wage("$1000/周"), 1000 / 40)
    assert approx(k.parse_wage("月薪5000"), 5000 / 173.3)


def test_wage_unknown_is_none():
    assert k.parse_wage("高于新西兰法定最低工资") is None
    assert k.parse_wage("面议") is None
    assert k.parse_wage("2个人") is None


def test_listing_wage_and_rent_use_new_parsers():
    it = k.Listing(tid=1, category="job", salary="年薪 60000")
    k.Scraper._post_process(it)
    assert approx(it.salary_num, 60000 / 2080), it.salary_num
    it = k.Listing(tid=2, category="rent", price_text="2房 $500")
    k.Scraper._post_process(it)
    assert it.price == 500, it.price


# ---- 导出 Excel ----
def _export_one(**fields):
    from openpyxl import load_workbook
    it = k.Listing(tid=1, category="rent", url="https://example.invalid/t", **fields)
    path = os.path.join(os.environ["SKYKIWI_APP_DIR"], "t.xlsx")
    k.export_excel([it], "rent", path)
    ws = load_workbook(path).active
    head = [c.value for c in ws[1]]
    return {h: ws.cell(row=2, column=i + 1) for i, h in enumerate(head)}


def test_export_formula_like_text_stays_text():
    cells = _export_one(title='=HYPERLINK("http://x","点我")')
    c = cells["标题"]
    assert c.data_type == "s", c.data_type
    assert c.value == '=HYPERLINK("http://x","点我")', c.value


def test_export_strips_illegal_control_chars():
    cells = _export_one(title="好\x01房", content="a\x02b")
    assert cells["标题"].value == "好房"
    assert cells["正文摘要"].value == "ab"


# ---- 错误日志 ----
def test_setup_logging_writes_file():
    log_path = k.setup_logging()
    import logging
    logging.getLogger("skykiwi").error("测试日志 %s", "xyz")
    for h in logging.getLogger("skykiwi").handlers:
        h.flush()
    assert os.path.dirname(log_path) == geo.APP_DIR
    with open(log_path, encoding="utf-8") as f:
        assert "测试日志 xyz" in f.read()


# ==================== v1.1.0 ====================

# ---- 默认排除已租 / 已招满 ----
def test_exclude_taken_posts():
    items = [k.Listing(tid=1, category="rent", title="Albany 双人房出租"),
             k.Listing(tid=2, category="rent", title="已租"),
             k.Listing(tid=3, category="job", title="招聘服务员（已招到）")]
    kept = k.apply_filters(items, k.Criteria(exclude_taken=True))
    assert [it.tid for it in kept] == [1], [it.tid for it in kept]
    assert len(k.apply_filters(items, k.Criteria(exclude_taken=False))) == 3


# ---- 距离显示: 精度不同显示方式不同 ----
def _loc(level, d):
    it = k.Listing(tid=1, category="rent")
    it.loc_level, it.distance = level, d
    return it


def test_format_distance():
    assert k.format_distance(_loc("地址", 2.34)) == "2.3"
    assert k.format_distance(_loc("街道", 0.42)) == "0.4"
    assert k.format_distance(_loc("街区", 0.0)) == "<1"
    assert k.format_distance(_loc("街区", 3.26)) == "3.3"
    assert k.format_distance(_loc("区域", 4.5)) == "≈5"
    assert k.format_distance(_loc("区域", 0.3)) == "≈1"
    assert k.format_distance(_loc("", None)) == ""


# ---- 排序: 指定列, None 永远在最后 ----
def test_sort_listings_by_price_none_last():
    a, b, c = (k.Listing(tid=i, category="rent") for i in range(3))
    a.price, b.price, c.price = 300.0, None, 200.0
    assert [x.tid for x in k.sort_listings([a, b, c], "price", False)] == [2, 0, 1]
    assert [x.tid for x in k.sort_listings([a, b, c], "price", True)] == [0, 2, 1]


# ---- 命令行入口判断 ----
def test_is_cli():
    assert k.is_cli(["app.exe", "--cli", "--category", "rent"])
    assert k.is_cli(["app.exe", "--version"])
    assert not k.is_cli(["app.exe"])
    assert not k.is_cli(["app.exe", r"C:\some\dragged file.txt"])


# ---- 抓取: 4xx 不重试; 列表页中途失败保留已抓的页; 停止时保留数据 ----
class _Resp:
    def __init__(self, status, text=""):
        self.status_code, self.text, self.encoding, self.apparent_encoding = status, text, "utf-8", "utf-8"

    def raise_for_status(self):
        import requests
        if self.status_code >= 400:
            raise requests.HTTPError(f"{self.status_code}", response=self)


def test_get_does_not_retry_client_errors():
    sc = k.Scraper(delay=0)
    calls = []

    class S:
        headers = {}

        def get(self, url, timeout):
            calls.append(url)
            return _Resp(404)
    sc._local.s = S()
    try:
        sc.get("https://example.invalid/x")
    except Exception:  # noqa: BLE001
        pass
    assert len(calls) == 1, f"404 被重试了 {len(calls)} 次"


def _list_page(tids, date):
    rows = "".join(
        f'<tbody id="normalthread_{t}"><tr><th><a href="forum.php?mod=viewthread&tid={t}" class="xst">帖子{t}</a></th>'
        f'<td class="by"><cite><a>u</a></cite><em><span>{date}</span></em></td>'
        f'<td class="num"><a>1</a><em>2</em></td></tr></tbody>' for t in tids)
    return f"<html><body><table>{rows}</table></body></html>"


def test_list_failure_keeps_earlier_pages():
    import datetime as dt
    today = dt.date.today().isoformat()
    sc = k.Scraper(delay=0, log=lambda m: None)
    pages = {1: _list_page([11, 12], today)}

    def fake_get(url, retries=3):
        page = int(url.rsplit("page=", 1)[1])
        if page in pages:
            return pages[page]
        raise RuntimeError("请求失败: 模拟第 2 页出错")
    sc.get = fake_get
    items = sc.list_threads("rent", 10)
    assert sorted(it.tid for it in items) == [11, 12], items


def test_stop_during_list_keeps_data():
    import datetime as dt
    today = dt.date.today().isoformat()
    sc = k.Scraper(delay=0, log=lambda m: None)

    def fake_get(url, retries=3):
        page = int(url.rsplit("page=", 1)[1])
        if page == 1:
            return _list_page([21, 22], today)
        raise InterruptedError("已停止")
    sc.get = fake_get
    items = sc.list_threads("rent", 10)
    assert sorted(it.tid for it in items) == [21, 22]


# ---- 地图服务: 全进程共享一个实例 ----
def test_geocoder_singleton():
    assert geo.get_geocoder() is geo.get_geocoder()


# ---- 配置: 原子写入, 文件损坏时备份而不是静默清空 ----
def test_config_atomic_and_corrupt_backup():
    geo.save_config({"home_text": "A"})
    assert geo.load_config()["home_text"] == "A"
    with open(geo.CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write("{坏的 json")
    assert geo.load_config() == {}
    assert os.path.exists(geo.CONFIG_FILE + ".bad"), "损坏的配置没有备份"
    assert not [n for n in os.listdir(geo.APP_DIR) if n.endswith(".tmp")], "残留临时文件"


# ---- 英文描述中的 close / view 不是街道名 ----
def test_ambiguous_street_words():
    assert geo.match_place("Nice room in Takapuna close to the beach") == "Takapuna"
    assert geo.match_place("Albany view apartment") == "Albany"
    assert geo.match_place("CBD市中心 Nelson Street") == "Auckland CBD"   # 真正的街道名仍跳过


# ---- 打包的 exe 是 引导进程 -> Python 进程 两层: 要能找到祖父进程 (启动它的 cmd) ----
def test_parent_pid_lookup():
    assert k._parent_pid(os.getpid()) == os.getppid()
    assert k._parent_pid(0x7FFFFFF0) is None


if __name__ == "__main__":
    import traceback
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
