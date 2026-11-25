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
import logging
import datetime as dt
import os
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field

import requests
from bs4 import BeautifulSoup

import geo

from version import __version__  # noqa: E402

log = logging.getLogger("skykiwi")
log.addHandler(logging.NullHandler())   # 未调用 setup_logging 时 (如测试) 不输出到终端

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
    def __init__(self, delay: float = 0.3, workers: int = 4, timeout: int = 20,
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

    # 所有线程共享的请求节流: 相邻两次请求至少间隔 delay 秒 (对论坛友好)
    _rate_lock = threading.Lock()
    _last_request = 0.0

    def _throttle(self):
        with Scraper._rate_lock:
            wait = self.delay - (time.time() - Scraper._last_request)
            if wait > 0:
                time.sleep(wait)
            Scraper._last_request = time.time()

    def get(self, url: str, retries: int = 3) -> str:
        last = None
        for i in range(retries):
            if self.stop_event.is_set():
                raise InterruptedError("已停止")
            self._throttle()
            try:
                r = self._session().get(url, timeout=self.timeout)
                r.raise_for_status()
                r.encoding = r.apparent_encoding if not r.encoding or r.encoding.lower() == "iso-8859-1" else r.encoding
                return r.text
            except requests.HTTPError as e:
                code = getattr(e.response, "status_code", 0) or 0
                if 400 <= code < 500 and code != 429:   # 404/403 等重试也没用
                    raise RuntimeError(f"请求失败: {url} (HTTP {code})") from e
                last = e
            except requests.RequestException as e:
                last = e
            time.sleep(1.5 * (i + 1))
        raise RuntimeError(f"请求失败: {url} ({last})")

    # ---- 列表页 ----
    def list_threads(self, category: str, days: int, max_pages: int = 40,
                     progress=None) -> list[Listing]:
        cfg = CATEGORIES[category]
        days = max(1, int(days))
        cutoff = dt.date.today() - dt.timedelta(days=days - 1)
        results: dict[int, Listing] = {}
        for page in range(1, max_pages + 1):
            url = (f"{BASE}forum.php?mod=forumdisplay&fid={cfg['fid']}"
                   f"&filter=author&orderby=dateline&page={page}")
            self.log(f"读取列表第 {page} 页 ...")
            try:
                html = self.get(url)
            except InterruptedError:
                self.log("已停止, 保留已读取的列表")
                break
            except RuntimeError as e:
                if not results:
                    raise
                log.warning("列表第 %d 页读取失败, 保留前 %d 条: %s", page, len(results), e)
                self.log(f"第 {page} 页读取失败, 保留已读取的 {len(results)} 条")
                break
            if page == max_pages:
                self.log(f"已达到最多 {max_pages} 页的上限, 更早的帖子未读取")
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
            it.price = parse_rent(it.price_text) or guess_price_from_text(it.title)
        else:
            # 薪金字段多为 "高于最低工资" 这类档位;能按单位换算出时薪就用, 否则从标题/正文猜时薪
            it.salary_num = (parse_wage(it.salary) or guess_wage_from_text(it.title)
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
    gc = geo.get_geocoder() if online_geocode else None
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


RENT_RANGE = (50, 5000)      # 合理的周租金 (NZD)
WAGE_RANGE = (15, 200)       # 合理的时薪 (NZD)
_HOURS = {"year": 2080, "month": 173.3, "week": 40, "day": 8, "hour": 1}


def parse_rent(text: str) -> float | None:
    """周租金字段 -> 数值。优先取 $ 后面的数字, 只接受合理区间, 过滤 '88888' 之类占位数。"""
    if not text:
        return None
    t = text.replace(",", "")
    nums = [(m.group(1) is not None, float(m.group(2)))
            for m in re.finditer(r"([\$＄]\s*)?(\d+(?:\.\d+)?)", t)]
    ok = [(dollar, v) for dollar, v in nums if RENT_RANGE[0] <= v <= RENT_RANGE[1]]
    for dollar, v in ok:
        if dollar:
            return v
    return ok[0][1] if ok else None


def parse_wage(text: str) -> float | None:
    """薪金文本 -> 估算时薪。识别 年/月/周/天/小时 和 k, 换算后只接受合理区间。"""
    if not text:
        return None
    t = text.lower().replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*(k)?(?![\d.])", t)
    if not m:
        return None
    v = float(m.group(1)) * (1000 if m.group(2) else 1)
    if m.group(2) or re.search(r"年薪|/\s*年|每年|per\s*annum|(?<![a-z])p\.?a\.?(?![a-z])|annual|year|/\s*yr", t):
        unit = "year"
    elif re.search(r"月薪|/\s*月|每月|month", t):
        unit = "month"
    elif re.search(r"周薪|/\s*周|每周|week|(?<![a-z])p/?w(?![a-z])", t):
        unit = "week"
    elif re.search(r"日薪|/\s*天|每天|/\s*day|per\s*day|daily", t):
        unit = "day"
    else:
        unit = "hour"
    hourly = v / _HOURS[unit]
    return round(hourly, 1) if WAGE_RANGE[0] <= hourly <= WAGE_RANGE[1] else None


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
    exclude_taken: bool = False                         # 排除标题注明已租 / 已招满的帖子
    max_distance: float | None = None                   # 距家最大距离 km (需先设置家庭住址)
    sort_distance: bool = False                         # 按距离由近到远排序


TAKEN_RE = re.compile(r"已租|已出租|已经租出|租出去了|已招到|已招满|招满|已满员|已找到|已结束|已转让")


def _contains(hay: str, needle: str) -> bool:
    return needle.strip().lower() in (hay or "").lower()


def _distance_key(x):
    # 位置可靠的(地址/街道/街区)在前, 只知大区的估算排在后面, 无法定位的最后
    return (x.distance is None, x.loc_level == "区域", x.distance if x.distance is not None else 0)


def sort_listings(items, key: str, reverse: bool = False) -> list:
    """按列排序; 空值永远排在最后。distance 列按可靠程度分组。"""
    if key == "distance":
        known = [x for x in items if x.distance is not None]
        known.sort(key=lambda x: (x.loc_level == "区域", -x.distance if reverse else x.distance))
        return known + [x for x in items if x.distance is None]
    vals = [(getattr(x, key, None), x) for x in items]
    present = [(v, x) for v, x in vals if v not in (None, "")]
    missing = [x for v, x in vals if v in (None, "")]
    present.sort(key=lambda p: p[0] if isinstance(p[0], (int, float)) else str(p[0]), reverse=reverse)
    return [x for _, x in present] + missing


def format_distance(it) -> str:
    """表格中的距离: 地址/街道 1 位小数; 街区 <1 km 显示 '<1'; 只知大区显示 '≈整数'。"""
    d = it.distance
    if d is None:
        return ""
    if it.loc_level == "区域":
        return f"≈{max(1, int(d + 0.5))}"   # 四舍五入 (round 是银行家舍入, 4.5 会变 4)
    if it.loc_level == "街区" and d < 1:
        return "<1"
    return f"{d:.1f}"


def apply_filters(items: list[Listing], c: Criteria) -> list[Listing]:
    out = []
    for it in items:
        if c.regions and it.region not in c.regions:
            continue
        if c.exclude_ads and it.region == "商家广告":
            continue
        if c.exclude_taken and TAKEN_RE.search(it.title or ""):
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
        out.sort(key=_distance_key)
    return out


# ---------------------------------------------------------------- 导出 ----

def columns_for(category: str):
    return RENT_COLUMNS if category == "rent" else JOB_COLUMNS


def _excel_safe(v):
    """去掉 Excel 不允许的控制字符; None -> 空字符串。"""
    if v is None:
        return ""
    if isinstance(v, str):
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
        return ILLEGAL_CHARACTERS_RE.sub("", v)
    return v


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
            row.append(_excel_safe(v))
        ws.append(row)
        for c in ws[ws.max_row]:   # 帖子内容来自网友: 以 = 开头的也只当文本, 不当公式
            if isinstance(c.value, str) and c.value.startswith("="):
                c.data_type = "s"
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
            ws2.append([it.tid, _excel_safe(it.title), _excel_safe(k), _excel_safe(v)])
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
            addr = "" if a.no_geocode else geo.get_geocoder().reverse(lat, lon)
            home = (lat, lon, f"当前位置 (精度约 ±{acc} m): {addr or f'{lat:.5f}, {lon:.5f}'}")
            cfg.update(home_text=addr or f"{lat:.5f}, {lon:.5f}", home_lat=lat, home_lon=lon, home_desc=home[2])
            geo.save_config(cfg)
        except geo.LocateError as e:
            print(f"一键定位失败: {e}")
    elif a.home:
        home = geo.resolve_home(a.home, None if a.no_geocode else geo.get_geocoder())
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
    """图形界面见 gui.py; 把本模块传过去, 避免以脚本运行时被重复导入。"""
    import gui
    gui.run(sys.modules[__name__])


def setup_logging() -> str:
    """日志写到 <数据目录>/app.log (最多约 3 MB), 并记录所有未处理的异常。返回日志路径。"""
    from logging.handlers import RotatingFileHandler
    path = os.path.join(geo.APP_DIR, "app.log")
    if not any(getattr(h, "baseFilename", None) == os.path.abspath(path) for h in log.handlers):
        os.makedirs(geo.APP_DIR, exist_ok=True)
        h = RotatingFileHandler(path, maxBytes=1_000_000, backupCount=2, encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s"))
        log.addHandler(h)
        log.setLevel(logging.INFO)

        def on_exc(exc_type, exc, tb):
            log.error("未处理的异常", exc_info=(exc_type, exc, tb))
            sys.__excepthook__(exc_type, exc, tb)
        sys.excepthook = on_exc
        threading.excepthook = lambda a: log.error(
            "后台线程未处理的异常", exc_info=(a.exc_type, a.exc_value, a.exc_traceback))
    return path


def is_cli(argv) -> bool:
    """有以 - 开头的参数才进命令行模式 (拖文件到 exe 上仍打开图形界面)。"""
    return any(a.startswith("-") for a in argv[1:])


def _parent_pid(pid: int) -> int | None:
    """用 Toolhelp 快照查进程的父进程 ID (Windows); 找不到返回 None。"""
    try:
        import ctypes
        from ctypes import wintypes
    except ImportError:
        return None

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
                    ("th32ProcessID", wintypes.DWORD), ("th32DefaultHeapID", ctypes.c_void_p),
                    ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
                    ("th32ParentProcessID", wintypes.DWORD), ("pcPriClassBase", ctypes.c_long),
                    ("dwFlags", wintypes.DWORD), ("szExeFile", ctypes.c_wchar * 260)]
    k32 = ctypes.windll.kernel32
    k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    snap = k32.CreateToolhelp32Snapshot(0x2, 0)          # TH32CS_SNAPPROCESS
    if not snap or snap == wintypes.HANDLE(-1).value:
        return None
    try:
        e = PROCESSENTRY32W()
        e.dwSize = ctypes.sizeof(e)
        ok = k32.Process32FirstW(snap, ctypes.byref(e))
        while ok:
            if e.th32ProcessID == pid:
                return e.th32ParentProcessID
            ok = k32.Process32NextW(snap, ctypes.byref(e))
        return None
    finally:
        k32.CloseHandle(snap)


def attach_parent_console():
    """--windowed 打包的 exe 没有控制台: 命令行模式下挂到启动它的 cmd / PowerShell 窗口, 让输出可见。"""
    if not getattr(sys, "frozen", False):
        return
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        # onefile exe 有两层进程: cmd -> 引导进程(无控制台) -> 本进程, 所以先试祖父进程
        grand = _parent_pid(os.getppid())
        if not ((grand and k32.AttachConsole(grand)) or k32.AttachConsole(-1)):
            log.info("命令行模式: 没有可连接的控制台, 输出不可见 (请在 cmd 中用 start /wait 运行)")
            return
        enc = f"cp{ctypes.windll.kernel32.GetConsoleOutputCP()}"
        sys.stdout = open("CONOUT$", "w", encoding=enc, errors="replace", buffering=1)
        sys.stderr = open("CONOUT$", "w", encoding=enc, errors="replace", buffering=1)
        print()
        log.info("命令行模式: 已连接父控制台 (代码页 %s)", enc)
    except (AttributeError, OSError, LookupError):
        pass


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
    setup_logging()
    log.info("启动 v%s %s", __version__, " ".join(sys.argv[1:]))
    if is_cli(sys.argv):
        attach_parent_console()
        run_cli()
    else:
        run_gui()


if __name__ == "__main__":
    main()
