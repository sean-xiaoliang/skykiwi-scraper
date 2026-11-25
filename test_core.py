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
