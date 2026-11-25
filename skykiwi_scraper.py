# -*- coding: utf-8 -*-
"""
天维网 (bbs.skykiwi.com) 租房 / 招聘信息抓取工具

功能:
  * 抓取「租房中心」(fid=19) 或「求职招聘」(fid=55) 最近 N 天(默认 10 天)发布的帖子
  * 进入每个帖子读取结构化字段(租金、房型、地址、电话 / 公司、职位、薪金、工作性质 等)
  * 按区域、价格(周租金)、类型、关键词等条件筛选;招聘可按岗位类别、工作性质、地点筛选
  * 导出 Excel(含超链接、筛选器、冻结表头)
  * 图形界面 (tkinter) + 命令行两种用法,可用 PyInstaller 打包成 Windows exe

命令行示例:
  python skykiwi_scraper.py --cli --category rent --days 10 --region 东区 北岸 --min-price 200 --max-price 450 -o 租房.xlsx
  python skykiwi_scraper.py --cli --category job --job-type 服务人员 --nature 兼职 --keyword 奥克兰 -o 招聘.xlsx
不带参数运行则打开图形界面。
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import queue
import re
import sys
import threading
import time
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field

import requests
from bs4 import BeautifulSoup

import geo

__version__ = "1.0.0"

BASE = "https://bbs.skykiwi.com/"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0 Safari/537.36"
)

# 版块配置: 区域/类别(typeid) 与 分类信息(sortid)
CATEGORIES = {
    "rent": {
        "name": "租房",
        "fid": 19,
        "typeids": {  # 区域
            33: "市中心", 34: "东区", 35: "西区", 36: "南区",
            37: "北岸", 38: "中区", 39: "其他地区", 505: "商家广告",
        },
        "sortids": {287: "出租", 298: "求租"},
        "type_label": "区域",
    },
    "job": {
        "name": "招聘",
        "fid": 55,
        "typeids": {  # 岗位类别
            470: "技术类工种", 471: "销售/市场", 482: "前台/行政", 483: "公关/顾问",
            484: "财会/金融", 485: "编辑/文职", 486: "清洁/家政", 487: "教育/培训",
            488: "服务人员", 489: "其它", 498: "商家广告",
        },
        "sortids": {231: "招聘", 232: "求职"},
        "type_label": "岗位类别",
    },
}

# Excel 列: (字段key, 表头, 列宽)
RENT_COLUMNS = [
    ("post_time", "发布时间", 17), ("region", "区域", 9), ("sort", "类型", 7),
    ("title", "标题", 50), ("price", "周租金(NZD)", 12), ("distance", "预估距离(km)", 12),
    ("loc_desc", "定位依据", 22), ("rent_type", "出租类型", 16),
    ("house_type", "房屋类型", 16), ("address", "地址", 24), ("facilities", "屋内设施", 30),
    ("phone", "电话", 15), ("wechat", "微信", 15), ("author", "发帖人", 14),
    ("replies", "回复", 6), ("views", "查看", 7), ("content", "正文摘要", 60), ("url", "链接", 45),
]
JOB_COLUMNS = [
    ("post_time", "发布时间", 17), ("region", "岗位类别", 12), ("sort", "类型", 7),
    ("title", "标题", 46), ("position", "职位", 18), ("company", "公司名称", 24),
    ("location", "工作地点", 16), ("distance", "预估距离(km)", 12), ("loc_desc", "定位依据", 22),
    ("salary", "薪金", 20), ("salary_num", "时薪(估算)", 10),
    ("nature", "工作性质", 10), ("work_time", "工作时间", 22), ("phone", "电话", 15),
    ("wechat", "微信", 15), ("email", "邮箱", 22), ("author", "发帖人", 14),
    ("replies", "回复", 6), ("views", "查看", 7), ("content", "正文摘要", 60), ("url", "链接", 45),
]

# 帖子里结构化字段的中文名 -> 内部 key
FIELD_MAP = {
    "出租类型": "rent_type", "求租类型": "rent_type", "房屋类型": "house_type",
    "每周租金": "price_text", "租金": "price_text", "预算": "price_text", "每周预算": "price_text",
    "屋内设施": "facilities", "微信": "wechat", "联系电话": "phone", "电话": "phone",
    "地址": "address", "区域": "address", "期望区域": "address",
    "公司名称": "company", "工作地点": "location", "职位": "position", "薪金": "salary",
    "工作性质": "nature", "工作时间": "work_time", "邮箱": "email", "电子邮箱": "email",
    "Email": "email", "联系邮箱": "email",
}


@dataclass
class Listing:
    tid: int
    category: str
    title: str = ""
    url: str = ""
    region: str = ""        # typeid 对应的名称(租房=区域,招聘=岗位类别)
    sort: str = ""          # 出租/求租 或 招聘/求职
    author: str = ""
    post_date: str = ""     # 列表页上的日期 YYYY-MM-DD
    post_time: str = ""     # 详情页上的完整时间
    replies: int = 0
    views: int = 0
    # 租房
    price: float | None = None
    price_text: str = ""
    rent_type: str = ""
    house_type: str = ""
    address: str = ""
    facilities: str = ""
    # 招聘
    company: str = ""
    location: str = ""
    position: str = ""
    salary: str = ""
    salary_num: float | None = None
    nature: str = ""
    work_time: str = ""
    email: str = ""
    # 通用
    phone: str = ""
    wechat: str = ""
    content: str = ""
    extra: dict = field(default_factory=dict)
    detail_ok: bool = False
    # 位置 / 距离
    lat: float | None = None
    lon: float | None = None
    loc_level: str = ""     # 地址 / 街道 / 街区 / 区域 (精度由高到低)
    loc_name: str = ""
    loc_src: str = ""       # 依据来自: 地址 / 标题 / 正文 / 分区
    distance: float | None = None  # 与家庭住址的直线距离 km

    @property
    def loc_desc(self) -> str:
        if not self.loc_level:
            return "无法定位"
        note = {"区域": " (仅知大区, 误差大)"}.get(self.loc_level, "")
        src = f"·{self.loc_src}" if self.loc_src and self.loc_src != self.loc_level else ""
        return f"{self.loc_level}{src}: {self.loc_name}{note}"


# ---------------------------------------------------------------- 抓取 ----

class Scraper:
    def __init__(self, delay: float = 0.4, workers: int = 4, timeout: int = 20,
                 log=print, stop_event: threading.Event | None = None):
        self.delay = delay
        self.workers = max(1, workers)
        self.timeout = timeout
        self.log = log
        self.stop_event = stop_event or threading.Event()
        self._local = threading.local()

    def _session(self) -> requests.Session:
        s = getattr(self._local, "s", None)
        if s is None:
            s = requests.Session()
            s.headers.update({"User-Agent": USER_AGENT,
                              "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"})
            self._local.s = s
        return s

    def get(self, url: str, retries: int = 3) -> str:
        last = None
        for i in range(retries):
            if self.stop_event.is_set():
                raise InterruptedError("已停止")
            try:
                r = self._session().get(url, timeout=self.timeout)
                r.raise_for_status()
                r.encoding = r.apparent_encoding if not r.encoding or r.encoding.lower() == "iso-8859-1" else r.encoding
                time.sleep(self.delay)
                return r.text
            except requests.RequestException as e:
                last = e
                time.sleep(1.5 * (i + 1))
        raise RuntimeError(f"请求失败: {url} ({last})")

    # ---- 列表页 ----
    def list_threads(self, category: str, days: int, max_pages: int = 30,
                     progress=None) -> list[Listing]:
        cfg = CATEGORIES[category]
        cutoff = (dt.date.today() - dt.timedelta(days=days - 1)) if days > 0 else dt.date.min
        results: dict[int, Listing] = {}
        for page in range(1, max_pages + 1):
            url = (f"{BASE}forum.php?mod=forumdisplay&fid={cfg['fid']}"
                   f"&filter=author&orderby=dateline&page={page}")
            self.log(f"读取列表第 {page} 页 ...")
            html = self.get(url)
            rows, normal_dates = self._parse_list(html, category)
            for it in rows:
                d = _to_date(it.post_date)
                if d and d >= cutoff and it.tid not in results:
                    results[it.tid] = it
            if progress:
                progress(page, None)
            # 按发帖时间倒序: 本页普通帖全部早于截止日期 -> 停止翻页
            if not normal_dates or all(d < cutoff for d in normal_dates):
                break
            if self.stop_event.is_set():
                break
        self.log(f"共找到 {len(results)} 条最近 {days} 天的帖子")
        return sorted(results.values(), key=lambda x: (x.post_date, x.tid), reverse=True)

    def _parse_list(self, html: str, category: str):
        soup = BeautifulSoup(html, "html.parser")
        items, normal_dates = [], []
        for tb in soup.select('tbody[id^="normalthread_"], tbody[id^="stickthread_"]'):
            tid = int(tb["id"].split("_")[1])
            th = tb.find("th")
            a = th.select_one("a.xst") if th else None
            if not a:
                continue
            it = Listing(tid=tid, category=category, title=a.get_text(strip=True),
                         url=f"{BASE}forum.php?mod=viewthread&tid={tid}")
            for link in th.select("em > a"):
                href = link.get("href", "")
                m = re.search(r"typeid=(\d+)", href)
                if m:
                    it.region = CATEGORIES[category]["typeids"].get(int(m.group(1)), link.get_text(strip=True))
                m = re.search(r"sortid=(\d+)", href)
                if m:
                    it.sort = CATEGORIES[category]["sortids"].get(int(m.group(1)), link.get_text(strip=True))
            bys = tb.select("td.by")
            if bys:
                cite = bys[0].find("cite")
                it.author = cite.get_text(strip=True) if cite else ""
                em = bys[0].find("em")
                if em:
                    sp = em.find("span")
                    txt = (sp.get("title") if sp and sp.get("title") else em.get_text(" ", strip=True))
                    it.post_date = _norm_date(txt)
            num = tb.select_one("td.num")
            if num:
                a2, em2 = num.find("a"), num.find("em")
                it.replies = _to_int(a2.get_text() if a2 else "")
                it.views = _to_int(em2.get_text() if em2 else "")
            items.append(it)
            if tb["id"].startswith("normalthread_"):
                d = _to_date(it.post_date)
                if d:
                    normal_dates.append(d)
        return items, normal_dates

    # ---- 详情页 ----
    def fetch_detail(self, it: Listing) -> Listing:
        html = self.get(it.url)
        soup = BeautifulSoup(html, "html.parser")
        first = soup.select_one('table[id^="pid"]') or soup
        tbl = first.select_one("div.table-container table")
        if tbl:
            for tr in tbl.find_all("tr"):
                t = tr.select_one("td.td-title")
                c = tr.select_one("td.td-content")
                if not t or not c:
                    continue
                k = t.get_text(strip=True).rstrip(":：")
                v = re.sub(r"\s+", " ", c.get_text(" ", strip=True)).strip()
                key = FIELD_MAP.get(k)
                if key:
                    if not getattr(it, key):
                        setattr(it, key, v)
                elif v:
                    it.extra[k] = v
        msg = first.select_one('td[id^="postmessage_"]')
        if msg:
            # 去掉脚本/样式/表格/广告 以及 Discuz 防复制干扰字符 (font.jammer, display:none)
            for bad in msg.select("script, style, i.pstatus, .table-container, center, font.jammer, "
                                  "span[style*='display:none'], span[style*='display: none'], ignore_js_op, "
                                  ".tip, .attach_nopermission, .aimg_tip"):
                bad.decompose()
            text = re.sub(r"\s+", " ", msg.get_text(" ", strip=True))
            text = re.sub(r"^本帖最后由 .+? 编辑\s*", "", text)
            text = re.sub(r"\S*\s*(?:\d+\s*(?:秒|分钟|小时|天)前|\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{2})\s*上传\s*下载附件\s*\([^)]*\)\s*", " ", text)
            it.content = text[:800]
        em = first.select_one('em[id^="authorposton"]')
        if em:
            sp = em.find("span")
            txt = sp.get("title") if sp and sp.get("title") else em.get_text(" ", strip=True)
            m = re.search(r"\d{4}-\d{1,2}-\d{1,2}(\s+\d{1,2}:\d{2}(:\d{2})?)?", txt)
            if m:
                it.post_time = m.group(0)
        it.post_time = it.post_time or it.post_date
        it.detail_ok = True
        self._post_process(it)
        return it

    @staticmethod
    def _post_process(it: Listing):
        if it.category == "rent":
            it.price = parse_price(it.price_text) or guess_price_from_text(it.title)
        else:
            # 薪金字段多为 "高于最低工资" 这类档位;优先取其中数字,否则从标题/正文猜时薪
            it.salary_num = (parse_price(it.salary) or guess_wage_from_text(it.title)
                             or guess_wage_from_text(it.content))
            if not it.location:
                it.location = it.extra.get("地点", "")

    def scrape(self, category: str, days: int = 10, with_detail: bool = True,
               progress=None, online_geocode: bool = True) -> list[Listing]:
        items = self._scrape(category, days, with_detail, progress)
        locate_items(items, online_geocode, log=self.log, progress=progress, stop_event=self.stop_event)
        return items

    def _scrape(self, category, days, with_detail, progress) -> list[Listing]:
        items = self.list_threads(category, days, progress=progress)
        if not with_detail or not items:
            for it in items:
                self._post_process(it)
            return items
        total, done = len(items), 0
        self.log(f"开始读取 {total} 个帖子详情 (并发 {self.workers}) ...")
        with ThreadPoolExecutor(max_workers=self.workers) as ex:
            futs = {ex.submit(self.fetch_detail, it): it for it in items}
            for f in as_completed(futs):
                done += 1
                try:
                    f.result()
                except InterruptedError:
                    pass
                except Exception as e:  # noqa: BLE001
                    it = futs[f]
                    self._post_process(it)
                    self.log(f"  [跳过] {it.tid}: {e}")
                if progress:
                    progress(done, total)
                if self.stop_event.is_set():
                    for x in futs:
                        x.cancel()
                    self.log("已停止,保留已抓取的数据")
                    break
        for it in items:
            if not it.detail_ok:
                self._post_process(it)
        return items


def locate_items(items: list[Listing], online_geocode: bool = True, log=print, progress=None,
                 stop_event: threading.Event | None = None):
    """为帖子定位(离线街区匹配 + 可选联网解析门牌地址)。"""
    gc = geo.Geocoder() if online_geocode else None
    need_net = [it for it in items if gc and geo._address_queries(
        (it.address if it.category == "rent" else it.location) or "")]
    if need_net:
        log(f"解析帖子位置 (含 {len(need_net)} 个街道地址需联网, 已解析过的会直接用缓存) ...")
    for i, it in enumerate(items, 1):
        use_gc = gc if (gc and not (stop_event and stop_event.is_set())) else None
        try:
            geo.locate(it, use_gc)
        except Exception:  # noqa: BLE001
            geo.locate(it, None)
        if progress and need_net:
            progress(i, len(items))
    if gc:
        gc.save()
    ok = sum(1 for it in items if it.loc_level)
    log(f"定位完成: {ok}/{len(items)} 条可估算距离")


# ---------------------------------------------------------------- 工具函数 ----

def _to_int(s: str) -> int:
    m = re.search(r"\d+", s or "")
    return int(m.group(0)) if m else 0


def _norm_date(s: str) -> str:
    s = s or ""
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    today = dt.date.today()
    if "前天" in s:
        return str(today - dt.timedelta(days=2))
    if "昨天" in s:
        return str(today - dt.timedelta(days=1))
    m = re.search(r"(\d+)\s*天前", s)
    if m:
        return str(today - dt.timedelta(days=int(m.group(1))))
    if re.search(r"(小时|分钟|秒)前|刚刚", s):
        return str(today)
    return ""


def _to_date(s: str):
    try:
        return dt.date.fromisoformat(s)
    except (TypeError, ValueError):
        return None


def parse_price(text: str) -> float | None:
    """从 '380 NZD' / '$25-30/小时' 之类的文本中取第一个数字。"""
    if not text:
        return None
    m = re.search(r"(\d+(?:\.\d+)?)", text.replace(",", ""))
    if not m:
        return None
    v = float(m.group(1))
    return v if v > 0 else None


def guess_price_from_text(text: str) -> float | None:
    """字段缺失时从标题猜周租金,例如 '$250/周'、'280一周'、'300pw'。"""
    if not text:
        return None
    t = text.replace(",", "")
    for pat in (r"[\$＄]\s*(\d{2,4})", r"(\d{2,4})\s*(?:刀|纽币|nzd)?\s*(?:/|每|一)?\s*(?:周|pw|p/w|week)"):
        m = re.search(pat, t, re.I)
        if m:
            v = float(m.group(1))
            if 50 <= v <= 5000:
                return v
    return None


def guess_wage_from_text(text: str) -> float | None:
    """从文本中猜时薪,例如 '$25/小时'、'时薪28'、'25-30/hr'、'$26 per hour'。"""
    if not text:
        return None
    t = text.replace(",", "")
    pats = (
        r"[\$＄]\s*(\d{2}(?:\.\d{1,2})?)\s*(?:-\s*\d{2}(?:\.\d{1,2})?)?\s*(?:/|每|一|per\s*)?\s*(?:小时|hour|hr|h\b|时)",
        r"(?:时薪|每小时|hourly(?:\s*rate)?)\s*[:：]?\s*[\$＄]?\s*(\d{2}(?:\.\d{1,2})?)",
        r"(\d{2}(?:\.\d{1,2})?)\s*(?:-\s*\d{2}(?:\.\d{1,2})?)?\s*(?:刀|纽币|nzd)?\s*(?:/|每|一|per\s*)\s*(?:小时|hour|hr|h\b)",
    )
    for pat in pats:
        m = re.search(pat, t, re.I)
        if m:
            v = float(m.group(1))
            if 15 <= v <= 200:
                return v
    return None


# ---------------------------------------------------------------- 筛选 ----

@dataclass
class Criteria:
    regions: list[str] = field(default_factory=list)    # typeid 名称(租房区域 / 岗位类别),空=全部
    sorts: list[str] = field(default_factory=list)      # 出租/求租 或 招聘/求职,空=全部
    min_price: float | None = None                      # 租房: 周租金; 招聘: 薪金数值
    max_price: float | None = None
    keep_no_price: bool = True                          # 价格为空时是否保留
    keywords: list[str] = field(default_factory=list)   # 全部包含 (标题/地址/正文/职位...)
    exclude: list[str] = field(default_factory=list)    # 任一包含则排除
    address: str = ""                                   # 地址 / 工作地点 包含
    rent_type: str = ""                                 # 出租类型包含 (整租/分租/单间...)
    house_type: str = ""                                # 房屋类型包含
    nature: str = ""                                    # 工作性质包含 (全职/兼职...)
    position: str = ""                                  # 职位包含
    exclude_ads: bool = False                           # 排除商家广告
    max_distance: float | None = None                   # 距家最大距离 km (需先设置家庭住址)
    sort_distance: bool = False                         # 按距离由近到远排序


def _contains(hay: str, needle: str) -> bool:
    return needle.strip().lower() in (hay or "").lower()


def apply_filters(items: list[Listing], c: Criteria) -> list[Listing]:
    out = []
    for it in items:
        if c.regions and it.region not in c.regions:
            continue
        if c.exclude_ads and it.region == "商家广告":
            continue
        if c.sorts and it.sort not in c.sorts:
            continue
        val = it.price if it.category == "rent" else it.salary_num
        if c.min_price is not None or c.max_price is not None:
            if val is None:
                if not c.keep_no_price:
                    continue
            else:
                if c.min_price is not None and val < c.min_price:
                    continue
                if c.max_price is not None and val > c.max_price:
                    continue
        blob = " ".join(str(x) for x in (
            it.title, it.address, it.location, it.position, it.company, it.content,
            it.rent_type, it.house_type, it.facilities, it.nature, it.salary, it.work_time,
            " ".join(it.extra.values()))).lower()
        if any(k and k.lower() not in blob for k in c.keywords):
            continue
        if any(k and k.lower() in blob for k in c.exclude):
            continue
        if c.address and not (_contains(it.address, c.address) or _contains(it.location, c.address)
                              or _contains(it.title, c.address)):
            continue
        if c.rent_type and not _contains(it.rent_type, c.rent_type):
            continue
        if c.house_type and not _contains(it.house_type, c.house_type):
            continue
        if c.nature and not (_contains(it.nature, c.nature) or _contains(it.title, c.nature)):
            continue
        if c.position and not (_contains(it.position, c.position) or _contains(it.title, c.position)):
            continue
        if c.max_distance is not None and (it.distance is None or it.distance > c.max_distance):
            continue
        out.append(it)
    if c.sort_distance:
        # 位置可靠的(地址/街道/街区)在前, 只知大区的估算排在后面, 无法定位的最后
        out.sort(key=lambda x: (x.distance is None, x.loc_level == "区域",
                                x.distance if x.distance is not None else 0))
    return out


# ---------------------------------------------------------------- 导出 ----

def columns_for(category: str):
    return RENT_COLUMNS if category == "rent" else JOB_COLUMNS


def export_excel(items: list[Listing], category: str, path: str) -> str:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    cols = columns_for(category)
    wb = Workbook()
    ws = wb.active
    ws.title = CATEGORIES[category]["name"]
    head_fill = PatternFill("solid", fgColor="1F4E78")
    ws.append([h for _, h, _ in cols])
    for i, (_, _, w) in enumerate(cols, 1):
        cell = ws.cell(row=1, column=i)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    link_col = [k for k, _, _ in cols].index("url") + 1
    for it in items:
        row = []
        for k, _, _ in cols:
            v = getattr(it, k, "")
            row.append("" if v is None else v)
        ws.append(row)
        c = ws.cell(row=ws.max_row, column=link_col)
        c.hyperlink = it.url
        c.font = Font(color="0563C1", underline="single")
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(vertical="top", wrap_text=False)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # 第二个工作表: 额外字段
    extras = [(it, k, v) for it in items for k, v in it.extra.items()]
    if extras:
        ws2 = wb.create_sheet("其他字段")
        ws2.append(["tid", "标题", "字段", "内容"])
        for it, k, v in extras:
            ws2.append([it.tid, it.title, k, v])
    wb.save(path)
    return path


def default_filename(category: str) -> str:
    return f"天维网{CATEGORIES[category]['name']}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"


# ---------------------------------------------------------------- 命令行 ----

def run_cli(argv=None):
    p = argparse.ArgumentParser(description="天维网租房/招聘信息抓取")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    p.add_argument("--cli", action="store_true", help="使用命令行模式")
    p.add_argument("--category", choices=["rent", "job"], default="rent", help="rent=租房, job=招聘")
    p.add_argument("--days", type=int, default=10, help="最近 N 天 (默认 10)")
    p.add_argument("--region", nargs="*", default=[],
                   help="租房区域: 市中心 东区 西区 南区 北岸 中区 其他地区 商家广告")
    p.add_argument("--job-type", nargs="*", default=[], help="岗位类别, 如 服务人员 技术类工种")
    p.add_argument("--sort", nargs="*", default=[], help="出租/求租 或 招聘/求职")
    p.add_argument("--min-price", type=float, help="最低周租金 / 最低薪金数值")
    p.add_argument("--max-price", type=float, help="最高周租金 / 最高薪金数值")
    p.add_argument("--drop-no-price", action="store_true", help="设置价格条件时丢弃没有价格的帖子")
    p.add_argument("--keyword", nargs="*", default=[], help="必须包含的关键词(全部满足)")
    p.add_argument("--exclude", nargs="*", default=[], help="排除关键词")
    p.add_argument("--address", default="", help="地址 / 工作地点 包含")
    p.add_argument("--rent-type", default="", help="出租类型包含, 如 整租 分租 单间")
    p.add_argument("--house-type", default="", help="房屋类型包含, 如 公寓 独立屋 Unit")
    p.add_argument("--nature", default="", help="工作性质包含, 如 全职 兼职")
    p.add_argument("--position", default="", help="职位包含")
    p.add_argument("--no-ads", action="store_true", help="排除商家广告")
    p.add_argument("--no-detail", action="store_true", help="只抓列表,不进帖子(更快,但没有租金等字段)")
    p.add_argument("--here", action="store_true", help="用本机当前位置 (Windows 定位服务) 作为家庭住址")
    p.add_argument("--home", help="家庭住址 (街道地址 / 街区名 / '纬度,经度'); 不填则使用图形界面保存的住址")
    p.add_argument("--max-distance", type=float, help="距家最大距离 km")
    p.add_argument("--sort-distance", action="store_true", help="按距离由近到远排序")
    p.add_argument("--no-geocode", action="store_true", help="不联网解析门牌地址, 只用离线街区匹配")
    p.add_argument("--workers", type=int, default=4)
    p.add_argument("--delay", type=float, default=0.4)
    p.add_argument("-o", "--output", help="导出 Excel 文件名")
    a = p.parse_args(argv)

    sc = Scraper(delay=a.delay, workers=a.workers)
    last = [0.0]

    def prog(done, total):
        if total and (time.time() - last[0] > 2 or done == total):
            last[0] = time.time()
            print(f"  进度 {done}/{total}")

    items = sc.scrape(a.category, a.days, with_detail=not a.no_detail, progress=prog,
                      online_geocode=not a.no_geocode)
    home = None
    cfg = geo.load_config()
    if a.here:
        try:
            lat, lon, acc = geo.locate_device()
            addr = geo.Geocoder().reverse(lat, lon)
            home = (lat, lon, f"当前位置 (精度约 ±{acc} m): {addr or f'{lat:.5f}, {lon:.5f}'}")
            cfg.update(home_text=addr or f"{lat:.5f}, {lon:.5f}", home_lat=lat, home_lon=lon, home_desc=home[2])
            geo.save_config(cfg)
        except geo.LocateError as e:
            print(f"一键定位失败: {e}")
    elif a.home:
        home = geo.resolve_home(a.home, None if a.no_geocode else geo.Geocoder())
        if home:
            cfg.update(home_text=a.home, home_lat=home[0], home_lon=home[1], home_desc=home[2])
            geo.save_config(cfg)
        else:
            print(f"无法解析家庭住址: {a.home}")
    elif cfg.get("home_lat") is not None:
        home = (cfg["home_lat"], cfg["home_lon"], cfg.get("home_desc", ""))
    if home:
        print(f"家庭住址: {home[2]} ({home[0]:.5f}, {home[1]:.5f})")
    elif a.max_distance or a.sort_distance:
        print("未设置家庭住址, 距离筛选/排序无效 (用 --home 设置)")
    geo.compute_distances(items, home)
    crit = Criteria(
        regions=(a.region if a.category == "rent" else a.job_type), sorts=a.sort,
        min_price=a.min_price, max_price=a.max_price, keep_no_price=not a.drop_no_price,
        keywords=a.keyword, exclude=a.exclude, address=a.address, rent_type=a.rent_type,
        house_type=a.house_type, nature=a.nature, position=a.position, exclude_ads=a.no_ads,
        max_distance=a.max_distance if home else None, sort_distance=a.sort_distance and bool(home))
    res = apply_filters(items, crit)
    print(f"\n抓取 {len(items)} 条, 筛选后 {len(res)} 条\n")
    for it in res[:30]:
        dist = f"{it.distance:5.1f}km" if it.distance is not None else "   -   "
        if a.category == "rent":
            price = f"${it.price:g}" if it.price else "$-"
            print(f"{it.post_date} {dist} [{it.region}][{it.sort}] {price} {it.title[:36]}  ({it.loc_desc})")
        else:
            print(f"{it.post_date} {dist} [{it.region}][{it.sort}] {it.position or '-'} | {it.salary or '-'} | "
                  f"{it.nature or '-'} | {it.title[:30]}  ({it.loc_desc})")
    if len(res) > 30:
        print(f"... 其余 {len(res) - 30} 条见 Excel")
    out = a.output or default_filename(a.category)
    export_excel(res, a.category, out)
    print(f"\n已导出: {os.path.abspath(out)}")


# ---------------------------------------------------------------- 图形界面 ----

def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    root = tk.Tk()
    root.title(f"天维网 租房 / 招聘 信息抓取工具 v{__version__}")
    root.geometry("1280x780")
    try:
        root.tk.call("tk", "scaling", 1.25)
    except tk.TclError:
        pass
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")

    state = {"items": {"rent": [], "job": []}, "shown": [], "running": False, "stop": None, "home": None}
    msgq: queue.Queue = queue.Queue()
    cfg = geo.load_config()
    if cfg.get("home_lat") is not None:
        state["home"] = (cfg["home_lat"], cfg["home_lon"], cfg.get("home_desc", ""))

    cat_var = tk.StringVar(value="rent")
    days_var = tk.IntVar(value=10)
    detail_var = tk.BooleanVar(value=True)

    # ---- 顶部: 抓取 ----
    top = ttk.Frame(root, padding=8)
    top.pack(fill="x")
    ttk.Label(top, text="类别:").pack(side="left")
    for k in ("rent", "job"):
        ttk.Radiobutton(top, text=CATEGORIES[k]["name"], value=k, variable=cat_var,
                        command=lambda: on_cat_change()).pack(side="left", padx=4)
    ttk.Label(top, text="   最近天数:").pack(side="left")
    ttk.Spinbox(top, from_=1, to=60, width=5, textvariable=days_var).pack(side="left")
    ttk.Checkbutton(top, text="读取帖子详情(租金/薪金等)", variable=detail_var).pack(side="left", padx=10)
    btn_fetch = ttk.Button(top, text="开始抓取", command=lambda: start_fetch())
    btn_fetch.pack(side="left", padx=4)
    btn_stop = ttk.Button(top, text="停止", state="disabled", command=lambda: stop_fetch())
    btn_stop.pack(side="left", padx=4)
    pbar = ttk.Progressbar(top, length=220, mode="determinate")
    pbar.pack(side="left", padx=10)
    status = tk.StringVar(value="就绪")
    ttk.Label(top, textvariable=status, foreground="#555").pack(side="left", padx=6)

    # ---- 筛选区 ----
    fbox = ttk.LabelFrame(root, text="筛选条件 (抓取后可随时调整,点「应用筛选」)", padding=8)
    fbox.pack(fill="x", padx=8)

    type_frame = ttk.Frame(fbox)
    type_frame.grid(row=0, column=0, columnspan=12, sticky="w")
    type_vars: dict[str, tk.BooleanVar] = {}
    sort_vars: dict[str, tk.BooleanVar] = {}

    min_var, max_var = tk.StringVar(), tk.StringVar()
    keep_np_var = tk.BooleanVar(value=True)
    kw_var, ex_var, addr_var = tk.StringVar(), tk.StringVar(), tk.StringVar()
    rt_var, ht_var = tk.StringVar(), tk.StringVar()
    nature_var, pos_var = tk.StringVar(value="全部"), tk.StringVar()
    type_label_var = tk.StringVar(value="全部")
    noads_var = tk.BooleanVar(value=False)

    row2 = ttk.Frame(fbox)
    row2.grid(row=1, column=0, columnspan=12, sticky="w", pady=(6, 0))
    price_label = ttk.Label(row2, text="周租金 NZD:")
    price_label.pack(side="left")
    ttk.Entry(row2, width=7, textvariable=min_var).pack(side="left")
    ttk.Label(row2, text="~").pack(side="left")
    ttk.Entry(row2, width=7, textvariable=max_var).pack(side="left")
    ttk.Checkbutton(row2, text="保留无价格帖", variable=keep_np_var).pack(side="left", padx=6)
    addr_label = ttk.Label(row2, text="  地址包含:")
    addr_label.pack(side="left")
    ttk.Entry(row2, width=14, textvariable=addr_var).pack(side="left")
    ttk.Label(row2, text="  关键词(空格分隔,全部包含):").pack(side="left")
    ttk.Entry(row2, width=18, textvariable=kw_var).pack(side="left")
    ttk.Label(row2, text="  排除词:").pack(side="left")
    ttk.Entry(row2, width=12, textvariable=ex_var).pack(side="left")

    row3 = ttk.Frame(fbox)
    row3.grid(row=2, column=0, columnspan=12, sticky="w", pady=(6, 0))
    rent_opts = ttk.Frame(row3)
    ttk.Label(rent_opts, text="出租类型包含:").pack(side="left")
    ttk.Combobox(rent_opts, width=10, textvariable=rt_var,
                 values=["", "整租", "分租", "单间", "Flatmates", "短租"]).pack(side="left")
    ttk.Label(rent_opts, text="  房屋类型包含:").pack(side="left")
    ttk.Combobox(rent_opts, width=10, textvariable=ht_var,
                 values=["", "独立屋", "公寓", "单元房", "Townhouse", "Unit"]).pack(side="left")
    job_opts = ttk.Frame(row3)
    ttk.Label(job_opts, text="职位包含:").pack(side="left")
    ttk.Entry(job_opts, width=14, textvariable=pos_var).pack(side="left")
    ttk.Checkbutton(row3, text="排除商家广告", variable=noads_var).pack(side="right", padx=10)
    ttk.Button(row3, text="应用筛选", command=lambda: refresh_table()).pack(side="right", padx=4)
    ttk.Button(row3, text="重置筛选", command=lambda: reset_filters()).pack(side="right", padx=4)

    # ---- 家庭住址 / 距离 (租房和招聘通用) ----
    hbox = ttk.LabelFrame(root, text="家庭住址 / 距离 (租房和招聘通用)", padding=8)
    hbox.pack(fill="x", padx=8, pady=(6, 0))
    home_var = tk.StringVar(value=cfg.get("home_text", ""))
    maxdist_var = tk.StringVar(value=cfg.get("max_distance", ""))
    sortdist_var = tk.BooleanVar(value=bool(cfg.get("sort_distance", True)))
    online_var = tk.BooleanVar(value=bool(cfg.get("online_geocode", True)))
    btn_here = ttk.Button(hbox, text="📍 一键定位当前位置", command=lambda: locate_here())
    btn_here.pack(side="left", padx=(0, 10))
    ttk.Label(hbox, text="或输入地址:").pack(side="left")
    home_entry = ttk.Entry(hbox, width=34, textvariable=home_var)
    home_entry.pack(side="left")
    btn_home = ttk.Button(hbox, text="按地址定位", command=lambda: locate_home())
    btn_home.pack(side="left", padx=4)
    ttk.Label(hbox, text="  最大距离 km:").pack(side="left")
    ttk.Entry(hbox, width=6, textvariable=maxdist_var).pack(side="left")
    ttk.Checkbutton(hbox, text="按距离排序(近→远)", variable=sortdist_var,
                    command=lambda: (save_prefs(), refresh_table())).pack(side="left", padx=8)
    ttk.Checkbutton(hbox, text="联网解析帖子门牌地址(更准)", variable=online_var,
                    command=lambda: on_online_toggle()).pack(side="left")
    home_status = tk.StringVar()
    ttk.Label(root, textvariable=home_status, foreground="#2E6E2E", padding=(16, 2)).pack(fill="x")

    # ---- 结果表 ----
    mid = ttk.Frame(root, padding=(8, 6))
    mid.pack(fill="both", expand=True)
    tree = ttk.Treeview(mid, show="headings", selectmode="extended")
    vsb = ttk.Scrollbar(mid, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(mid, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    mid.rowconfigure(0, weight=1)
    mid.columnconfigure(0, weight=1)

    detail = tk.Text(root, height=6, wrap="word")
    detail.pack(fill="x", padx=8)

    bottom = ttk.Frame(root, padding=8)
    bottom.pack(fill="x")
    count_var = tk.StringVar(value="")
    ttk.Label(bottom, textvariable=count_var).pack(side="left")
    ttk.Button(bottom, text="导出 Excel (当前筛选结果)", command=lambda: do_export()).pack(side="right")
    ttk.Label(bottom, text="双击行在浏览器打开帖子;点击表头排序   ", foreground="#777").pack(side="right")

    # ---- 逻辑 ----
    def cell(it, k):
        v = getattr(it, k, "")
        if v is None:
            return ""
        if k == "distance" and it.loc_level == "区域":
            return f"≈{v:g}"          # 只知大区, 距离为粗略估算
        if isinstance(v, float):
            return f"{v:g}"
        return v

    def fill_tree(rows):
        tree.delete(*tree.get_children())
        keys = [k for k, _, _ in view_cols()]
        for it in rows:
            tags = ("coarse",) if it.loc_level in ("区域", "") else ()
            tree.insert("", "end", iid=str(it.tid), values=[cell(it, k) for k in keys], tags=tags)

    def show_home_status():
        h = state["home"]
        if h:
            home_status.set(f"✔ 住址已定位: {h[2][:90]}  ({h[0]:.5f}, {h[1]:.5f})   "
                            f"—— 预估距离为直线距离, 定位依据: 地址 > 街区 > 区域中心")
        else:
            home_status.set("未设置家庭住址: 在家时点「📍 一键定位当前位置」即可; 也可输入街道地址 "
                            "(如 1 Ti Rakau Drive, Pakuranga)、街区名或 '纬度,经度' 后点「按地址定位」")

    def save_prefs():
        c = geo.load_config()
        c.update(home_text=home_var.get().strip(), max_distance=maxdist_var.get().strip(),
                 sort_distance=sortdist_var.get(), online_geocode=online_var.get())
        if state["home"]:
            c.update(home_lat=state["home"][0], home_lon=state["home"][1], home_desc=state["home"][2])
        else:
            for k in ("home_lat", "home_lon", "home_desc"):
                c.pop(k, None)
        try:
            geo.save_config(c)
        except OSError:
            pass

    def locate_home():
        text = home_var.get().strip()
        if not text:
            state["home"] = None
            save_prefs()
            show_home_status()
            refresh_table()
            return
        btn_home.config(state="disabled")
        home_status.set("正在定位住址 ...")
        use_net = online_var.get()

        def work():
            try:
                r = geo.resolve_home(text, geo.Geocoder() if use_net else None)
            except Exception:  # noqa: BLE001
                r = None
            msgq.put(("home", r))
        threading.Thread(target=work, daemon=True).start()

    home_entry.bind("<Return>", lambda e: locate_home())

    def locate_here():
        btn_here.config(state="disabled")
        btn_home.config(state="disabled")
        home_status.set("正在通过 Windows 定位服务获取当前位置 (约 5-15 秒) ...")

        def work():
            try:
                lat, lon, acc = geo.locate_device()
            except geo.LocateError as e:
                msgq.put(("here_err", str(e)))
                return
            except Exception as e:  # noqa: BLE001
                msgq.put(("here_err", f"定位失败: {e}"))
                return
            try:
                addr = geo.Geocoder().reverse(lat, lon)
            except Exception:  # noqa: BLE001
                addr = ""
            msgq.put(("here", (lat, lon, acc, addr)))
        threading.Thread(target=work, daemon=True).start()

    def on_online_toggle():
        save_prefs()
        if online_var.get() and not state["running"] and any(state["items"].values()):
            status.set("联网解析已抓取帖子的门牌地址 ...")
            items_all = state["items"]["rent"] + state["items"]["job"]

            def work():
                locate_items(items_all, True, log=lambda m: msgq.put(("log", m)))
                msgq.put(("relocated", None))
            threading.Thread(target=work, daemon=True).start()

    def view_cols():
        cols = [c for c in columns_for(cat_var.get()) if c[0] not in ("content", "url", "facilities", "wechat")]
        return cols

    def setup_columns():
        cols = view_cols()
        tree["columns"] = [k for k, _, _ in cols]
        for k, h, w in cols:
            tree.heading(k, text=h, command=lambda kk=k: sort_by(kk))
            tree.column(k, width=int(w * 8), minwidth=40, stretch=(k == "title"))

    def build_type_checks():
        for w in type_frame.winfo_children():
            w.destroy()
        type_vars.clear()
        sort_vars.clear()
        cfg = CATEGORIES[cat_var.get()]
        ttk.Label(type_frame, text=cfg["type_label"] + ":").pack(side="left")
        if cat_var.get() == "rent":
            for name in cfg["typeids"].values():
                v = tk.BooleanVar(value=False)
                type_vars[name] = v
                ttk.Checkbutton(type_frame, text=name, variable=v).pack(side="left", padx=2)
        else:
            mb = ttk.Menubutton(type_frame, textvariable=type_label_var, width=26)
            menu = tk.Menu(mb, tearoff=0)

            def clear_types():
                for x in type_vars.values():
                    x.set(False)
                update_type_label()
                refresh_table()
            menu.add_command(label="全部 (清除选择)", command=clear_types)
            menu.add_separator()
            for name in cfg["typeids"].values():
                v = tk.BooleanVar(value=False)
                type_vars[name] = v
                menu.add_checkbutton(label=name, variable=v,
                                     command=lambda: (update_type_label(), refresh_table()))
            mb["menu"] = menu
            mb.pack(side="left", padx=2)
            update_type_label()
            ttk.Label(type_frame, text="   工作性质:").pack(side="left")
            cb = ttk.Combobox(type_frame, width=6, textvariable=nature_var, state="readonly",
                              values=["全部", "全职", "兼职"])
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: refresh_table())
        ttk.Label(type_frame, text="    类型:").pack(side="left")
        for name in cfg["sortids"].values():
            v = tk.BooleanVar(value=False)
            sort_vars[name] = v
            ttk.Checkbutton(type_frame, text=name, variable=v).pack(side="left", padx=2)
        ttk.Label(type_frame, text="  (不勾选=全部)", foreground="#777").pack(side="left")

    def update_type_label():
        sel = [k for k, v in type_vars.items() if v.get()]
        type_label_var.set("全部" if not sel else (", ".join(sel) if len(sel) <= 2 else f"已选 {len(sel)} 类: "
                                                   + ", ".join(sel[:2]) + " ..."))

    def on_cat_change():
        build_type_checks()
        if cat_var.get() == "rent":
            job_opts.pack_forget()
            rent_opts.pack(side="left")
            price_label.config(text="周租金 NZD:")
            addr_label.config(text="  地址包含:")
        else:
            rent_opts.pack_forget()
            job_opts.pack(side="left")
            price_label.config(text="时薪 NZD:")
            addr_label.config(text="  工作地点包含:")
        setup_columns()
        refresh_table()

    def _num(s):
        s = s.strip()
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None

    def current_criteria() -> Criteria:
        rent = cat_var.get() == "rent"
        return Criteria(
            regions=[k for k, v in type_vars.items() if v.get()],
            sorts=[k for k, v in sort_vars.items() if v.get()],
            min_price=_num(min_var.get()), max_price=_num(max_var.get()),
            keep_no_price=keep_np_var.get(),
            keywords=kw_var.get().split(), exclude=ex_var.get().split(),
            address=addr_var.get().strip(),
            rent_type=rt_var.get().strip() if rent else "",
            house_type=ht_var.get().strip() if rent else "",
            nature="" if rent or nature_var.get() in ("", "全部") else nature_var.get(),
            position=pos_var.get().strip() if not rent else "",
            exclude_ads=noads_var.get(),
            max_distance=_num(maxdist_var.get()) if state["home"] else None,
            sort_distance=sortdist_var.get() and bool(state["home"]))

    def reset_filters():
        for v in list(type_vars.values()) + list(sort_vars.values()):
            v.set(False)
        for v in (min_var, max_var, kw_var, ex_var, addr_var, rt_var, ht_var, pos_var):
            v.set("")
        nature_var.set("全部")
        update_type_label()
        keep_np_var.set(True)
        noads_var.set(False)
        refresh_table()

    def refresh_table():
        cat = cat_var.get()
        items = state["items"][cat]
        geo.compute_distances(items, state["home"])
        shown = apply_filters(items, current_criteria())
        state["shown"] = shown
        fill_tree(shown)
        extra = ""
        if state["home"] and items:
            n = sum(1 for it in shown if it.distance is not None)
            extra = f"  (其中 {n} 条有预估距离)"
        count_var.set(f"{CATEGORIES[cat]['name']}: 共抓取 {len(items)} 条,筛选后 {len(shown)} 条{extra}")

    sort_state = {"key": None, "rev": False}

    def sort_by(key):
        rev = not sort_state["rev"] if sort_state["key"] == key else False
        sort_state.update(key=key, rev=rev)

        def kf(it):
            v = getattr(it, key, "")
            if isinstance(v, (int, float)) or v is None:
                return (0, v if v is not None else float("-inf"))
            return (1, str(v))
        if key == "distance":   # 无距离的始终排在最后
            state["shown"].sort(key=lambda x: (x.distance is None, x.loc_level == "区域",
                                               (x.distance or 0) * (-1 if rev else 1)))
        else:
            state["shown"].sort(key=kf, reverse=rev)
        fill_tree(state["shown"])

    def find_item(tid):
        for it in state["items"][cat_var.get()]:
            if str(it.tid) == tid:
                return it
        return None

    def on_select(_e=None):
        sel = tree.selection()
        if not sel:
            return
        it = find_item(sel[0])
        if not it:
            return
        detail.delete("1.0", "end")
        lines = [it.title, it.url]
        dist = f"{it.distance:g} km (直线)" if it.distance is not None else ("-" if state["home"] else "未设置家庭住址")
        lines.append(f"预估距离: {dist}   定位依据: {it.loc_desc}")
        if it.category == "rent":
            lines.append(f"租金: {it.price_text or '-'}  出租类型: {it.rent_type}  房屋类型: {it.house_type}  "
                         f"地址: {it.address}  电话: {it.phone}  微信: {it.wechat}")
            if it.facilities:
                lines.append(f"设施: {it.facilities}")
        else:
            lines.append(f"公司: {it.company}  职位: {it.position}  薪金: {it.salary}  性质: {it.nature}  "
                         f"地点: {it.location}  时间: {it.work_time}  电话: {it.phone}  微信: {it.wechat}  邮箱: {it.email}")
        for k, v in it.extra.items():
            lines.append(f"{k}: {v}")
        lines.append(it.content)
        detail.insert("1.0", "\n".join(lines))

    def on_double(_e=None):
        sel = tree.selection()
        if sel:
            it = find_item(sel[0])
            if it:
                webbrowser.open(it.url)

    tree.tag_configure("coarse", foreground="#8A8A8A")   # 只能定位到大区 / 无法定位的行显示为灰色
    tree.bind("<<TreeviewSelect>>", on_select)
    tree.bind("<Double-1>", on_double)

    def do_export():
        cat = cat_var.get()
        if not state["shown"]:
            messagebox.showinfo("提示", "当前没有可导出的数据")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_filename(cat),
                                            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        try:
            export_excel(state["shown"], cat, path)
        except PermissionError:
            messagebox.showerror("导出失败", "文件被占用,请先关闭 Excel 中已打开的同名文件")
            return
        if messagebox.askyesno("导出完成", f"已导出 {len(state['shown'])} 条到\n{path}\n\n是否立即打开?"):
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except Exception:  # noqa: BLE001
                pass

    def start_fetch():
        if state["running"]:
            return
        try:
            days = int(days_var.get())
        except (tk.TclError, ValueError):
            days = 10
        cat = cat_var.get()
        stop = threading.Event()
        state.update(running=True, stop=stop)
        btn_fetch.config(state="disabled")
        btn_stop.config(state="normal")
        pbar.config(mode="indeterminate", value=0)
        pbar.start(12)

        def worker():
            sc = Scraper(log=lambda m: msgq.put(("log", m)), stop_event=stop)
            try:
                items = sc.scrape(cat, days, with_detail=detail_var.get(),
                                  progress=lambda d, t: msgq.put(("prog", (d, t))),
                                  online_geocode=online_var.get())
                msgq.put(("done", (cat, items)))
            except InterruptedError:
                msgq.put(("done", (cat, [])))
            except Exception as e:  # noqa: BLE001
                msgq.put(("error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def stop_fetch():
        if state["stop"]:
            state["stop"].set()
            status.set("正在停止 ...")

    def poll():
        try:
            while True:
                kind, data = msgq.get_nowait()
                if kind == "log":
                    status.set(data)
                elif kind == "prog":
                    d, t = data
                    if t:
                        if str(pbar["mode"]) != "determinate":
                            pbar.stop()
                            pbar.config(mode="determinate")
                        pbar.config(maximum=t, value=d)
                        status.set(f"处理中 {d}/{t}")
                elif kind == "here":
                    btn_here.config(state="normal")
                    btn_home.config(state="normal")
                    lat, lon, acc, addr = data
                    note = f"当前位置 (精度约 ±{acc} m)"
                    if acc > 2000:
                        note += " ⚠ 精度较低, 建议连接 Wi-Fi 后重试或手动输入地址"
                    if not (-47.6 <= lat <= -34.0 and 166.0 <= lon <= 179.0):
                        note += " ⚠ 当前位置不在新西兰"
                    state["home"] = (lat, lon, f"{note}: {addr or f'{lat:.5f}, {lon:.5f}'}")
                    home_var.set(addr or f"{lat:.5f}, {lon:.5f}")
                    save_prefs()
                    show_home_status()
                    refresh_table()
                elif kind == "here_err":
                    btn_here.config(state="normal")
                    btn_home.config(state="normal")
                    show_home_status()
                    messagebox.showwarning("一键定位失败", data)
                elif kind == "home":
                    btn_here.config(state="normal")
                    btn_home.config(state="normal")
                    if data:
                        state["home"] = data
                    else:
                        messagebox.showwarning("定位失败", "无法解析该住址。\n可以试试: 更完整的街道地址、街区名 "
                                               "(如 Mount Albert), 或直接输入 '纬度,经度'")
                    save_prefs()
                    show_home_status()
                    refresh_table()
                elif kind == "relocated":
                    status.set("帖子位置已更新")
                    refresh_table()
                elif kind in ("done", "error"):
                    pbar.stop()
                    pbar.config(mode="determinate", value=0)
                    state["running"] = False
                    btn_fetch.config(state="normal")
                    btn_stop.config(state="disabled")
                    if kind == "error":
                        status.set("出错")
                        messagebox.showerror("抓取失败", data)
                    else:
                        cat, items = data
                        if items:
                            state["items"][cat] = items
                        status.set(f"完成,抓取 {len(items)} 条")
                        if cat == cat_var.get():
                            refresh_table()
        except queue.Empty:
            pass
        root.after(150, poll)

    show_home_status()
    on_cat_change()
    poll()
    root.mainloop()


def main():
    # 打包成 --windowed exe 时没有控制台, stdout/stderr 为 None, 防止 print 报错
    for name in ("stdout", "stderr"):
        if getattr(sys, name) is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
        else:
            try:  # 中文 Windows 控制台为 GBK, 标题里的 emoji 等字符替换成 ? 而不是报错
                getattr(sys, name).reconfigure(errors="replace")
            except (AttributeError, ValueError):
                pass
    if len(sys.argv) > 1:
        run_cli()
    else:
        run_gui()


if __name__ == "__main__":
    main()
