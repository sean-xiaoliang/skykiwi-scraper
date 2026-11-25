# -*- coding: utf-8 -*-
"""v1.2 新功能测试: 本地缓存与增量抓取、简历分析与匹配、邮件申请。运行: python test_features.py"""
import datetime as dt
import os
import sys
import tempfile
import zipfile

os.environ["SKYKIWI_APP_DIR"] = tempfile.mkdtemp(prefix="skykiwi_test_")

import skykiwi_scraper as k  # noqa: E402
import store  # noqa: E402
import resume  # noqa: E402
import mail  # noqa: E402

TODAY = dt.date.today()


def day(n):
    return (TODAY - dt.timedelta(days=n)).isoformat()


def L(tid, cat="rent", date=None, **kw):
    it = k.Listing(tid=tid, category=cat, title=kw.pop("title", f"帖子{tid}"), post_date=date or day(0), **kw)
    return it


def new_cache():
    return store.PostCache(os.path.join(tempfile.mkdtemp(), "posts.db"))


# ==================== 本地缓存 ====================

def test_cache_roundtrip_keeps_fields():
    c = new_cache()
    a = L(1, price=300.0, address="1 Queen St", detail_ok=True, extra={"微信": "abc"})
    a.lat, a.lon, a.loc_level, a.loc_name, a.loc_src = -36.8, 174.7, "地址", "1 queen st", "地址"
    c.save([a])
    b = c.load("rent", 10)[0]
    assert (b.tid, b.price, b.address, b.extra, b.lat, b.loc_level) == (1, 300.0, "1 Queen St", {"微信": "abc"},
                                                                         -36.8, "地址")
    assert b.detail_ok is True


def test_cache_load_respects_days_and_category():
    c = new_cache()
    c.save([L(1, date=day(0)), L(2, date=day(9)), L(3, date=day(15)), L(4, cat="job", date=day(1))])
    assert sorted(x.tid for x in c.load("rent", 10)) == [1, 2]
    assert sorted(x.tid for x in c.load("rent", 20)) == [1, 2, 3]
    assert [x.tid for x in c.load("job", 10)] == [4]


def test_cache_prune_old_posts():
    c = new_cache()
    c.save([L(1, date=day(5)), L(2, date=day(90))])
    c.prune(keep_days=60)
    assert sorted(c.known("rent")) == [1]


def test_cache_meta():
    c = new_cache()
    assert c.get_meta("x") is None
    c.set_meta("x", "2026-01-01")
    assert c.get_meta("x") == "2026-01-01"


# ==================== 增量抓取 ====================

def _page(tids_dates):
    rows = "".join(
        f'<tbody id="normalthread_{t}"><tr><th><a href="forum.php?mod=viewthread&tid={t}" class="xst">帖子{t}</a></th>'
        f'<td class="by"><cite><a>u</a></cite><em><span>{d}</span></em></td>'
        f'<td class="num"><a>3</a><em>9</em></td></tr></tbody>' for t, d in tids_dates)
    return f"<html><body><table>{rows}</table></body></html>"


def _fake_scraper(pages, requested):
    sc = k.Scraper(delay=0, log=lambda m: None)

    def get(url, retries=3):
        p = int(url.rsplit("page=", 1)[1])
        requested.append(p)
        return pages.get(p, _page([]))
    sc.get = get
    return sc


def test_list_stops_at_fully_known_page():
    pages = {1: _page([(13, day(0)), (11, day(0))]), 2: _page([(12, day(1)), (10, day(2))]),
             3: _page([(9, day(3))])}
    req = []
    sc = _fake_scraper(pages, req)
    items = sc.list_threads("rent", 10, known={10, 11, 12}, covered_since=dt.date.fromisoformat(day(20)))
    assert req == [1, 2], f"多请求了列表页: {req}"
    assert 13 in [it.tid for it in items]
    assert sc.last_list_complete


def test_list_continues_when_coverage_insufficient():
    pages = {1: _page([(11, day(0))]), 2: _page([(10, day(5))]), 3: _page([(9, day(12))])}
    req = []
    sc = _fake_scraper(pages, req)
    sc.list_threads("rent", 10, known={11, 10}, covered_since=None)
    assert req == [1, 2, 3], req


def test_refresh_fetches_details_only_for_new_posts():
    c = new_cache()
    old = L(11, date=day(0), detail_ok=True, price=200.0)
    c.save([old])
    c.set_meta("covered_since_rent", day(30))
    pages = {1: _page([(13, day(0)), (11, day(0))]), 2: _page([(11, day(0))])}
    sc = _fake_scraper(pages, [])
    fetched = []

    def fake_detail(it):
        fetched.append(it.tid)
        it.detail_ok = True
        it.price_text = "350 NZD"
        k.Scraper._post_process(it)
        return it
    sc.fetch_detail = fake_detail
    items, n_new = sc.refresh("rent", 10, c, with_detail=True, online_geocode=False)
    assert fetched == [13], fetched
    assert n_new == 1
    got = {it.tid: it for it in items}
    assert got[11].price == 200.0 and got[13].price == 350.0
    assert got[11].replies == 3, "已缓存帖子的回复数应更新"


def test_refresh_saves_coverage_after_complete_run():
    c = new_cache()
    pages = {1: _page([(5, day(0))]), 2: _page([(4, day(40))])}
    sc = _fake_scraper(pages, [])
    sc.fetch_detail = lambda it: it
    sc.refresh("rent", 10, c, with_detail=False, online_geocode=False)
    assert c.get_meta("covered_since_rent") == day(9)


# ==================== 简历 ====================

RESUME_TEXT = """张三  电话 021 000 0000
求职意向：厨师 / 厨房帮工
工作经历：在奥克兰中餐厅担任厨师 5 年，熟悉炒锅、寿司制作。
语言：普通话、英语（雅思 6）
证照：新西兰正式驾照 Full licence
签证：Open work visa 开放工签
"""


def test_extract_text_txt_gbk(tmp=tempfile.mkdtemp()):
    p = os.path.join(tmp, "cv.txt")
    with open(p, "w", encoding="gbk") as f:
        f.write(RESUME_TEXT)
    assert "厨师" in resume.extract_text(p)


def test_extract_text_docx():
    p = os.path.join(tempfile.mkdtemp(), "cv.docx")
    body = "".join(f"<w:p><w:r><w:t>{line}</w:t></w:r></w:p>" for line in RESUME_TEXT.splitlines())
    xml = ('<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/'
           f'wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>')
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("word/document.xml", xml)
    text = resume.extract_text(p)
    assert "厨师" in text and "Open work visa" in text


def test_extract_text_pdf():
    from pypdf import PdfWriter
    from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject
    w = PdfWriter()
    page = w.add_blank_page(300, 200)
    font = DictionaryObject({NameObject("/Type"): NameObject("/Font"), NameObject("/Subtype"): NameObject("/Type1"),
                             NameObject("/BaseFont"): NameObject("/Helvetica")})
    page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): DictionaryObject(
        {NameObject("/F1"): w._add_object(font)})})
    s = DecodedStreamObject()
    s.set_data(b"BT /F1 12 Tf 20 100 Td (Chef with 5 years experience, forklift licence) Tj ET")
    page[NameObject("/Contents")] = w._add_object(s)
    p = os.path.join(tempfile.mkdtemp(), "cv.pdf")
    with open(p, "wb") as f:
        w.write(f)
    assert "Chef" in resume.extract_text(p)


def test_extract_text_rejects_unknown_type():
    p = os.path.join(tempfile.mkdtemp(), "cv.png")
    with open(p, "wb") as f:
        f.write(b"x")
    try:
        resume.extract_text(p)
    except resume.ResumeError:
        return
    raise AssertionError("应提示不支持的格式")


def test_analyze_resume():
    prof = resume.analyze(RESUME_TEXT)
    assert prof.categories[0] == "服务人员", prof.categories
    assert "英语" in prof.languages and "普通话" in prof.languages
    assert "驾照" in prof.licences
    assert prof.visa and "工签" in prof.visa
    assert prof.years == 5
    assert "厨师" in prof.keywords


def test_match_prefers_relevant_job():
    prof = resume.analyze(RESUME_TEXT)
    cook = L(1, "job", title="中餐厅招厨师", region="服务人员", position="厨师", content="需要有经验的炒锅师傅")
    clean = L(2, "job", title="招清洁工", region="清洁/家政", position="清洁", content="家庭清洁")
    sc, sl = resume.match(prof, cook), resume.match(prof, clean)
    assert 0 <= sl < sc <= 100, (sc, sl)


def test_match_penalises_missing_requirement():
    prof = resume.analyze("厨师 3 年经验")          # 没写英语
    job = L(1, "job", title="招厨师", region="服务人员", position="厨师", content="要求英语流利")
    prof_en = resume.analyze("厨师 3 年经验，英语流利")
    assert resume.match(prof_en, job) > resume.match(prof, job)


def test_match_ignores_job_seeker_posts():
    prof = resume.analyze(RESUME_TEXT)
    seeker = L(1, "job", title="厨师求职, 5年经验", region="服务人员", sort="求职", content="会炒锅")
    assert resume.match(prof, seeker) == 0


# ==================== 邮件 ====================

def test_extract_emails():
    text = "请发简历至 HR@abc.co.nz 或 jobs@abc.co.nz，抄送 hr@abc.co.nz。微信 abc123"
    assert mail.extract_emails(text) == ["hr@abc.co.nz", "jobs@abc.co.nz"]
    assert mail.extract_emails("没有邮箱") == []


def test_job_email_from_field_or_content():
    it = L(1, "job", email="", content="Send CV to cv@shop.nz thanks")
    assert mail.job_emails(it) == ["cv@shop.nz"]
    it.email = "boss@shop.nz"
    assert mail.job_emails(it)[0] == "boss@shop.nz"


def test_mailto_url_encodes_chinese_and_newlines():
    u = mail.mailto_url("hr@abc.co.nz", "应聘：厨师", "您好\n谢谢")
    assert u.startswith("mailto:hr@abc.co.nz?")
    assert "%E5%BA%94%E8%81%98" in u and "%0D%0A" in u and " " not in u


def test_webmail_urls():
    g = mail.webmail_url("gmail", "hr@abc.co.nz", "应聘", "正文")
    o = mail.webmail_url("outlook", "hr@abc.co.nz", "应聘", "正文")
    assert g.startswith("https://mail.google.com/mail/?view=cm") and "to=hr%40abc.co.nz" in g
    assert o.startswith("https://outlook.live.com/") and "subject=" in o


def test_choose_method():
    assert mail.choose_method("auto", client="Outlook.URL.mailto.15", webmail="gmail") == "client"
    assert mail.choose_method("auto", client=None, webmail="outlook") == "outlook"
    assert mail.choose_method("gmail", client="X", webmail="outlook") == "gmail"
    assert mail.choose_method("client", client=None, webmail="gmail") == "gmail"   # 没有客户端时退回网页


def test_client_friendly_name():
    assert mail.client_name("MailMaster.FNVLZXW3IDB4GMBTS6UBB2VZSI") == "网易邮箱大师"
    assert mail.client_name("Outlook.URL.mailto.15") == "Outlook"
    assert mail.client_name(r'"C:\Program Files\Mozilla Thunderbird\thunderbird.exe" -osint -compose "%1"') == "Thunderbird"
    assert mail.client_name("MSEdgeHTM") == "浏览器 (Edge)"
    assert mail.client_name("SomethingElse.Mail") == "SomethingElse.Mail"
    assert mail.client_name(None) == "没有"


def test_fill_template():
    it = L(1, "job", title="招厨师", position="厨师", company="好吃餐厅", url="https://x/1")
    s = mail.fill("应聘：{职位}（{公司}）{链接}", it, summary="5 年经验")
    assert s == "应聘：厨师（好吃餐厅）https://x/1"
    assert mail.fill("{简历摘要}", it, summary="5 年经验") == "5 年经验"
    assert mail.fill("{职位}", L(2, "job", title="招聘洗碗工"), summary="") == "招聘洗碗工"   # 没有职位时用标题


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
