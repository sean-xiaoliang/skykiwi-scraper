# -*- coding: utf-8 -*-
"""
位置解析与距离估算

每条帖子的定位顺序:
  1. 地址 / 街道 : 帖子的"地址"(招聘为"工作地点")字段含门牌号或街道名时, 联网 (OpenStreetMap Nominatim)
                  解析; 同名街道取与上下文 (帖子提到的街区 / 所在分区) 最接近的候选, 太远则丢弃
  2. 街区       : 在 地址、标题、正文 中识别内置的奥克兰街区 / 地标 / 新西兰城市 (离线, 支持中文与繁体)
                  - 跳过 "Nelson Street" 这类街道名、"去/到/离/直达 X" 这类路程描述、"排除 X" 这类否定
                  - 与帖子所在分区 (东区/北岸...) 明显矛盾的地名会被丢弃
                  - 只有路程描述时 (如 "开车去 Ormiston mall 3 分钟") 作为 "附近" 使用
  3. 区域       : 只知道大区 (东区、北岸...), 使用该区中心点, 误差较大
距离为与家庭住址之间的直线距离 (km)。
"""

from __future__ import annotations

import json
import math
import os
import re
import threading
import time
import unicodedata
from typing import NamedTuple

import requests

from nz_places import PLACES
from version import __version__

# 数据目录 (配置、缓存、日志); 环境变量 SKYKIWI_APP_DIR 可覆盖 (测试用)
APP_DIR = os.environ.get("SKYKIWI_APP_DIR") or os.path.join(
    os.environ.get("APPDATA") or os.path.expanduser("~"), "SkykiwiScraper")
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
CACHE_FILE = os.path.join(APP_DIR, "geocache2.json")
NOMINATIM = "https://nominatim.openstreetmap.org/search"
GEO_UA = f"SkykiwiScraper/{__version__} (desktop app)"

# 别名 -> PLACES 中的标准名 (中文叫法、拼写变体、地标)。键会经过 _norm (小写、繁转简、去撇号)
ALIASES = {
    "梅西大学": "Massey University Albany", "梅西": "Massey University Albany",
    "massey university": "Massey University Albany", "massey uni": "Massey University Albany",
    "奥克兰大学": "University of Auckland", "奥大": "University of Auckland",
    "university of auckland": "University of Auckland", "uoa": "University of Auckland",
    "奥克兰理工": "AUT", "理工大学": "AUT", "奥理工": "AUT", "aut": "AUT",
    "市中心": "Auckland CBD", "cbd": "Auckland CBD", "city centre": "Auckland CBD",
    "auckland central": "Auckland CBD", "auckland city": "Auckland CBD", "皇后街": "Auckland CBD",
    "queen street": "Auckland CBD", "britomart": "Auckland CBD", "wynyard quarter": "Auckland CBD",
    "机场": "Auckland Airport", "auckland airport": "Auckland Airport", "airport": "Auckland Airport",
    "sylvia park": "Sylvia Park Mall", "西尔维亚": "Sylvia Park Mall",
    "botany town centre": "Botany Town Centre", "botany": "Botany Downs",
    "flatbush": "Flat Bush", "佛光山": "Flat Bush", "meadowland": "Meadowlands", "new market": "Newmarket",
    "newlynn": "New Lynn", "lynnmall": "New Lynn", "lynmall": "New Lynn",
    "st lukes": "Mount Albert", "eastridge": "Mission Bay", "te atatu": "Te Atatu Peninsula",
    "wairau park": "Wairau Valley", "westlake": "Forrest Hill",
    "dominion rd": "Dominion Road", "dominion 路": "Dominion Road", "dominion路": "Dominion Road", "倒霉路": "Dominion Road", "多米尼路": "Dominion Road",
    "unitec": "Unitec", "植物园": "Auckland Botanic Gardens", "botanic gardens": "Auckland Botanic Gardens",
    "豪威克": "Howick", "帕库兰加": "Pakuranga", "新市场": "Newmarket", "奥巴尼": "Albany",
    "阿尔巴尼": "Albany", "塔卡普纳": "Takapuna", "博塔尼": "Botany Downs", "保塔尼": "Botany Downs",
    "亨德森": "Henderson", "马努考": "Manukau", "曼努考": "Manukau", "帕帕库拉": "Papakura",
    "伊甸山": "Mount Eden", "艾伯特山": "Mount Albert", "阿尔伯特山": "Mount Albert",
    "罗斯基尔": "Mount Roskill", "埃普森": "Epsom", "雷姆艾拉": "Remuera", "格伦菲尔德": "Glenfield",
    "惠灵顿山": "Mount Wellington", "奥尼洪加": "Onehunga", "新林": "New Lynn", "德文港": "Devonport",
    "汉密尔顿": "Hamilton", "哈密尔顿": "Hamilton", "hamiltom": "Hamilton", "惠灵顿": "Wellington",
    "基督城": "Christchurch", "皇后镇": "Queenstown", "陶朗加": "Tauranga", "陶朗阿": "Tauranga",
    "tuaranga": "Tauranga", "丹尼丁": "Dunedin", "但尼丁": "Dunedin", "罗托鲁阿": "Rotorua",
    "北帕": "Palmerston North", "陶波": "Taupo", "芒格努伊": "Mount Maunganui", "纳尔逊": "Nelson",
    "旺格雷": "Whangarei", "下哈特": "Lower Hutt", "瓦纳卡": "Wanaka", "阿什伯顿": "Ashburton",
}
# "市中心" 之类是相对说法: 帖子不在奥克兰时不能当作奥克兰 CBD
GENERIC = {"市中心", "cbd", "city centre", "auckland central", "auckland city", "皇后街", "queen street"}
# 本身就是道路的地名, 后面跟 road 不算街道名误判
ROAD_PLACES = {"Dominion Road"}

# 奥克兰大区: 中心点 + 合理半径 (km), 用于粗略定位和一致性检查
DISTRICTS = {
    "市中心": ("Auckland CBD", 12), "中区": ("Mount Roskill", 15), "东区": ("Botany Downs", 18),
    "西区": ("Henderson", 30), "南区": ("Manukau", 35), "北岸": ("Glenfield", 35),
}
REGION_WORDS = {
    "北岸": "北岸", "北区": "北岸", "north shore": "北岸", "东区": "东区", "东奥克兰": "东区",
    "east auckland": "东区", "西区": "西区", "西奥克兰": "西区", "west auckland": "西区",
    "南区": "南区", "南奥克兰": "南区", "south auckland": "南区", "中区": "中区",
    "市区": "市中心",
}

# 繁体 -> 简体 (地名中常见的字)
_TRAD = str.maketrans({
    "奧": "奥", "東": "东", "區": "区", "灣": "湾", "島": "岛", "學": "学", "園": "园", "機": "机",
    "場": "场", "頓": "顿", "爾": "尔", "蘭": "兰", "維": "维", "亞": "亚", "鎮": "镇", "庫": "库",
    "華": "华", "門": "门", "達": "达", "邊": "边", "樓": "楼", "歐": "欧", "羅": "罗", "濱": "滨",
    "際": "际", "寧": "宁", "陽": "阳", "國": "国", "陸": "陆", "車": "车", "離": "离", "開": "开",
    "駕": "驾", "鐘": "钟", "內": "内", "納": "纳", "遜": "逊", "漢": "汉", "爾": "尔", "鄰": "邻",
    "近": "近", "倫": "伦", "華": "华", "裡": "里", "峽": "峡", "廣": "广", "級": "级", "線": "线",
})
STREET_WORDS = (r"street|st|road|rd|avenue|ave|drive|dr|place|pl|crescent|cres|terrace|tce|lane|ln|"
                r"highway|hwy|way|close|grove|parade|boulevard|blvd|court|ct|rise|view|esplanade")
_STREET_AFTER = re.compile(r"^\s*(" + STREET_WORDS + r")(?![a-z])")
_STREET_IN = re.compile(r"(?<![a-z])(" + STREET_WORDS + r")(?![a-z])")
# 这些词也是普通英文 (close to / sea view / on the way): 后面还跟着英文单词时不算街道名
_AMBIGUOUS_STREET = {"close", "view", "way", "rise", "grove", "court", "ct", "parade", "place", "pl"}


def _is_street_suffix(after: str) -> bool:
    m = _STREET_AFTER.match(after)
    if not m:
        return False
    if m.group(1) in _AMBIGUOUS_STREET and re.match(r"\s+[a-z]", after[m.end():]):
        return False
    return True
_TRAVEL_BEFORE = re.compile(r"(去|到|离|距|至|往|直达|通往|开车|驾车|步行|走路|坐车|搭车|车程|(?<![a-z])(to|from)\s*$)")
_TRAVEL_AFTER = re.compile(r"^\s*(mall|商场|购物中心|校区)?\s*(只需|仅需|大概|大约|约|也就|才)?\s*\d+\s*(分钟|min)")
_NEGATED = re.compile(r"(排除|除了|除去|不要|不考虑|不含|不包括|except|not\s)")
AKL_BOX = (-37.45, -36.25, 174.2, 175.35)  # lat_min, lat_max, lon_min, lon_max


def _norm(text: str) -> str:
    t = unicodedata.normalize("NFKC", text or "").translate(_TRAD)
    t = "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c))
    t = t.lower().replace("’", "").replace("'", "").replace("‘", "")
    t = re.sub(r"(?<![a-z])mt\.?\s*(?=[a-z])", "mount ", t)
    t = re.sub(r"(?<![a-z])pt\.?\s+(?=[a-z])", "point ", t)
    t = re.sub(r"(?<![a-z])st\.\s*(?=[a-z])", "st ", t)
    return re.sub(r"\s+", " ", t)


def _compile(key: str):
    if re.search(r"[a-z]", key):
        return re.compile(r"(?<![a-z])" + re.escape(key) + r"(?![a-z])")
    return re.compile(re.escape(key))


def _build_patterns():
    entries = {name.lower(): name for name in PLACES}
    for k, v in ALIASES.items():
        if v in PLACES:
            entries[_norm(k)] = v
    return [(_compile(k), k, v) for k, v in sorted(entries.items(), key=lambda kv: -len(kv[0]))]


_PATTERNS = _build_patterns()
_REGION_PATTERNS = [(_compile(_norm(k)), v) for k, v in sorted(REGION_WORDS.items(), key=lambda kv: -len(kv[0]))]


class Match(NamedTuple):
    pos: int
    name: str
    generic: bool
    street: bool
    travel: bool
    negated: bool


def find_places(text: str) -> list[Match]:
    """找出文本中所有(不重叠的)地名, 并标记: 街道名 / 路程描述 / 否定 / 相对说法。"""
    t = _norm(text)
    if not t:
        return []
    raw = []
    for rx, key, name in _PATTERNS:
        for m in rx.finditer(t):
            raw.append((m.start(), m.end(), key, name))
    raw.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    out, last_end = [], -1
    for s, e, key, name in raw:
        if s < last_end:
            continue
        last_end = e
        before = t[max(0, s - 6):s]
        out.append(Match(
            pos=s, name=name, generic=key in GENERIC,
            street=name not in ROAD_PLACES and _is_street_suffix(t[e:]),
            travel=bool(_TRAVEL_BEFORE.search(before)) or bool(_TRAVEL_AFTER.match(t[e:e + 14])),
            negated=bool(_NEGATED.search(t[max(0, s - 8):s]))))
    return out


def match_place(text: str) -> str | None:
    """文本中最可信的地名: 跳过街道名和否定; 路程描述只在没有别的地名时使用。"""
    ms = [m for m in find_places(text) if not m.street and not m.negated]
    for m in ms:
        if not m.travel:
            return m.name
    return ms[0].name if ms else None


def match_district(text: str) -> str | None:
    """文本中最早出现的大区 (东区/北岸...), 跳过否定。"""
    t = _norm(text)
    best = None
    for rx, d in _REGION_PATTERNS:
        for m in rx.finditer(t):
            if _NEGATED.search(t[max(0, m.start() - 8):m.start()]):
                continue
            if best is None or m.start() < best[0]:
                best = (m.start(), d)
            break
    return best[1] if best else None


def match_region(text: str) -> str | None:
    """兼容旧接口: 返回大区代表街区名。"""
    d = match_district(text)
    return DISTRICTS[d][0] if d else None


def haversine(a, b) -> float:
    lat1, lon1 = map(math.radians, a[:2])
    lat2, lon2 = map(math.radians, b[:2])
    h = math.sin((lat2 - lat1) / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2
    return 2 * 6371.0 * math.asin(math.sqrt(h))


def parse_latlon(s: str):
    m = re.fullmatch(r"\s*(-?\d{1,2}\.\d+)\s*[,， ]\s*(-?\d{2,3}\.\d+)\s*", s or "")
    if m:
        lat, lon = float(m.group(1)), float(m.group(2))
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return lat, lon
    return None


def in_auckland(lat, lon) -> bool:
    return AKL_BOX[0] <= lat <= AKL_BOX[1] and AKL_BOX[2] <= lon <= AKL_BOX[3]


# ---------------------------------------------------------------- 配置 ----

def _read_json(path: str) -> dict:
    """读取 JSON; 文件损坏时改名为 .bad 备份 (而不是静默丢弃), 返回空字典。"""
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except (OSError, ValueError):
        try:
            os.replace(path, path + ".bad")
        except OSError:
            pass
        return {}


def _write_json(path: str, data, **kw):
    """原子写入: 先写临时文件再替换, 程序中途被关也不会留下写了一半的文件。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = f"{path}.{os.getpid()}.{threading.get_ident()}.tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, **kw)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def load_config() -> dict:
    return _read_json(CONFIG_FILE)


def save_config(cfg: dict):
    _write_json(CONFIG_FILE, cfg, indent=2)


# ---------------------------------------------------------------- 联网解析 ----

class Geocoder:
    """OpenStreetMap Nominatim: 每秒最多 1 次请求, 每次最多取 5 个候选, 结果缓存在本地。"""

    def __init__(self, timeout: int = 15):
        self.timeout = timeout
        self._lock = threading.RLock()   # save() 可能在 search() 持锁时调用
        self._last = 0.0
        self._session = requests.Session()
        self._session.headers.update({"User-Agent": GEO_UA, "Accept-Language": "en"})
        self.cache = _read_json(CACHE_FILE)
        self._dirty = 0

    def search(self, query: str) -> list[tuple]:
        """返回候选 [(lat, lon, display_name), ...]。"""
        key = query.strip().lower()
        if not key:
            return []
        if key in self.cache:
            return [tuple(x) for x in self.cache[key]]
        with self._lock:
            if key in self.cache:
                return [tuple(x) for x in self.cache[key]]
            wait = 1.1 - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait)
            try:
                r = self._session.get(NOMINATIM, params={
                    "q": query, "format": "json", "limit": 5, "countrycodes": "nz"}, timeout=self.timeout)
                self._last = time.time()
                r.raise_for_status()
                js = r.json()
            except (requests.RequestException, ValueError):
                self._last = time.time()
                return []  # 网络错误不缓存, 下次重试
            val = [[float(x["lat"]), float(x["lon"]), x.get("display_name", "")] for x in js]
            self.cache[key] = val
            self._dirty += 1
            if self._dirty >= 10:
                self.save()
            return [tuple(x) for x in val]

    def geocode(self, query: str):
        res = self.search(query)
        return res[0] if res else None

    def reverse(self, lat: float, lon: float) -> str:
        """坐标 -> 简短地址 (门牌 街道, 街区, 城市); 失败返回空字符串。"""
        with self._lock:
            wait = 1.1 - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait)
            try:
                r = self._session.get(NOMINATIM.replace("/search", "/reverse"), params={
                    "lat": f"{lat:.6f}", "lon": f"{lon:.6f}", "format": "json", "zoom": 18,
                    "addressdetails": 1}, timeout=self.timeout)
                self._last = time.time()
                r.raise_for_status()
                js = r.json()
            except (requests.RequestException, ValueError):
                self._last = time.time()
                return ""
        a = js.get("address") or {}
        street = " ".join(x for x in (a.get("house_number"), a.get("road")) if x)
        suburb = a.get("suburb") or a.get("neighbourhood") or a.get("hamlet") or a.get("village") or ""
        city = a.get("city") or a.get("town") or a.get("county") or ""
        parts = [p for p in (street, suburb, city) if p]
        return ", ".join(dict.fromkeys(parts)) or js.get("display_name", "")

    def save(self):
        try:
            with self._lock:
                snapshot = dict(self.cache)
            _write_json(CACHE_FILE, snapshot)
            self._dirty = 0
        except OSError:
            pass


_geocoder: Geocoder | None = None
_geocoder_lock = threading.Lock()


def get_geocoder() -> Geocoder:
    """全进程共享一个 Geocoder: 限速 (每秒 1 次) 和缓存只有一份, 不会互相覆盖。"""
    global _geocoder
    with _geocoder_lock:
        if _geocoder is None:
            _geocoder = Geocoder()
        return _geocoder


class LocateError(Exception):
    pass


# 通过 Windows 定位服务 (Wi-Fi / GPS) 获取本机位置; 用系统自带的 PowerShell + .NET, 无需额外依赖
_PS_LOCATE = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Device
$w = New-Object System.Device.Location.GeoCoordinateWatcher([System.Device.Location.GeoPositionAccuracy]::High)
[void]$w.TryStart($false, [TimeSpan]::FromSeconds(%(t)d))
$end = (Get-Date).AddSeconds(%(t)d)
while (($w.Status -ne 'Ready') -and ($w.Permission -ne 'Denied') -and ((Get-Date) -lt $end)) { Start-Sleep -Milliseconds 200 }
$c = $w.Position.Location
if ($w.Permission -eq 'Denied') { 'DENIED' }
elseif ($c.IsUnknown) { 'UNKNOWN ' + $w.Status }
else { 'OK {0} {1} {2}' -f $c.Latitude.ToString([Globalization.CultureInfo]::InvariantCulture), $c.Longitude.ToString([Globalization.CultureInfo]::InvariantCulture), [int]$c.HorizontalAccuracy }
$w.Stop()
"""

LOCATION_HELP = ("请打开 Windows 定位:\n设置 → 隐私和安全性 → 位置 → 打开「定位服务」和「允许桌面应用访问你的位置」,"
                 "然后重试。\n\n也可以直接在输入框里填写地址后点「按地址定位」。")


def locate_device(timeout: int = 15):
    """返回 (纬度, 经度, 精度米)。失败抛出 LocateError(原因)。"""
    import subprocess
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        out = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-Command", _PS_LOCATE % {"t": timeout}],
            capture_output=True, text=True, timeout=timeout + 15, creationflags=flags).stdout.strip()
    except FileNotFoundError as e:
        raise LocateError("找不到 PowerShell, 无法调用 Windows 定位服务") from e
    except subprocess.TimeoutExpired as e:
        raise LocateError("定位超时。\n\n" + LOCATION_HELP) from e
    last = out.splitlines()[-1] if out else ""
    if last.startswith("OK "):
        _, lat, lon, acc = last.split()
        return float(lat), float(lon), int(acc)
    if last == "DENIED":
        raise LocateError("Windows 拒绝了位置访问。\n\n" + LOCATION_HELP)
    raise LocateError(f"无法获取当前位置 ({last or '无输出'})。\n\n" + LOCATION_HELP)


def resolve_home(text: str, geocoder=None):
    """解析家庭住址: '纬度,经度'、街道地址(联网)、街区名(离线)。返回 (lat, lon, 说明) 或 None。"""
    text = (text or "").strip()
    if not text:
        return None
    ll = parse_latlon(text)
    if ll:
        return ll[0], ll[1], f"坐标 {ll[0]:.5f}, {ll[1]:.5f}"
    if geocoder:
        q = text if re.search(r"new zealand|\bnz\b|新西兰", text, re.I) else f"{text}, New Zealand"
        res = geocoder.search(q)
        if not res and not re.search(r"auckland", text, re.I):
            res = geocoder.search(f"{text}, Auckland, New Zealand")
        geocoder.save()
        if res:
            lat, lon, name = res[0]
            others = [r for r in res[1:] if haversine(r, res[0]) > 2]
            if others:
                name += f"   ⚠ 有 {len(others) + 1} 处同名地址, 若不是这里请加上街区名 (如 ..., Mount Eden)"
            return lat, lon, name
    name = match_place(text)
    if name:
        lat, lon = PLACES[name]
        return lat, lon, f"街区 {name} (离线估算)"
    d = match_district(text)
    if d:
        lat, lon = PLACES[DISTRICTS[d][0]]
        return lat, lon, f"{d} 中心 (离线粗略估算)"
    return None


# ---------------------------------------------------------------- 帖子定位 ----

def _clean_query(addr: str) -> str:
    q = _norm(addr)
    q = re.sub(r"[^\x00-\x7f]+", " ", q)                                  # 去掉中文等
    q = re.sub(r"\(.*?\)", " ", q)
    q = re.sub(r"(?<![a-z])(?:level|lvl|l)\s*\d+\s*[/,]\s*", " ", q)       # L3/135 -> 135
    q = re.sub(r"(?<![a-z])(?:flat|unit|apt|apartment|room)\s*\d+[a-z]?\s*[,/]?\s*", " ", q)
    q = re.sub(r"(\d)([a-z]{3,})", r"\1 \2", q)                          # 11union -> 11 union
    q = re.sub(r"\s+", " ", q).strip(" ,.;-/")
    return q


def _address_queries(addr: str):
    """地址字段 -> [(查询词, 是否含门牌号)]; 只有像街道地址时才联网。"""
    q = _clean_query(addr or "")
    if len(q) < 4 or not re.search(r"[a-z]{3,}", q):
        return []
    has_num = bool(re.search(r"(?<![a-z0-9])\d+[a-z]?(/\d+[a-z]?)?\s+[a-z]", q))
    if not has_num and not _STREET_IN.search(q):
        return []           # 只是街区名, 用离线匹配即可
    qs = [(q, has_num)]
    if has_num:             # 门牌号查不到时退回到街道
        street_only = re.sub(r"(?<![a-z0-9])\d+[a-z]?(/\d+[a-z]?)?\s+(?=[a-z])", "", q, count=1).strip()
        if street_only != q and _STREET_IN.search(street_only):
            qs.append((street_only, False))
    return qs


def _context(it):
    """帖子的分区上下文: (大区名或 None, 是否在奥克兰以外)。"""
    if it.category == "rent":
        if it.region == "其他地区":
            return None, True
        return (it.region if it.region in DISTRICTS else None), False
    loc = it.location or ""
    if "其他地区" in loc:
        return None, True
    return match_district(loc), False


def locate(it, geocoder=None):
    """为一条帖子设置 lat / lon / loc_level / loc_name / loc_src。"""
    addr = it.address if it.category == "rent" else it.location
    district, outside = _context(it)
    dcenter, dradius = (PLACES[DISTRICTS[district][0]], DISTRICTS[district][1]) if district else (None, None)

    def consistent(p):
        if dcenter is not None:
            return haversine(p, dcenter) <= dradius
        return True

    # ---- 文本中的地名候选 ----
    primary = travel = None
    for src, text in (("地址", addr), ("标题", it.title), ("正文", (it.content or "")[:600])):
        for m in find_places(text):
            if m.street or m.negated or (m.generic and outside):
                continue
            if not consistent(PLACES[m.name]):
                continue
            if m.travel:
                travel = travel or (src, m.name)
                continue
            primary = (src, m.name)
            break
        if primary:
            break

    # ---- 1) 联网解析地址 / 街道 ----
    if geocoder and addr:
        hint = PLACES[primary[1]] if primary else dcenter
        radius = 3.5 if primary else dradius
        for q, has_num in _address_queries(addr):
            if not outside and "auckland" not in q:
                q += ", auckland"
            res = geocoder.search(q + ", new zealand")
            if not res:
                continue
            best = min(res, key=lambda r: haversine(r, hint)) if hint else res[0]
            if hint is not None and haversine(best, hint) > radius:
                continue
            if hint is None and not outside and not in_auckland(best[0], best[1]):
                continue
            it.lat, it.lon = best[0], best[1]
            it.loc_level = "地址" if has_num else "街道"
            it.loc_name = _clean_query(addr)[:40]
            it.loc_src = "地址"
            return it

    # ---- 2) 街区 ----
    for cand, suffix in ((primary, ""), (travel, " (附近)")):
        if cand:
            it.lat, it.lon = PLACES[cand[1]]
            it.loc_level, it.loc_name, it.loc_src = "街区", cand[1] + suffix, cand[0]
            return it

    # ---- 3) 大区 ----
    for src, text in (("地址", addr), ("标题", it.title), ("正文", (it.content or "")[:600])):
        d = match_district(text)
        if d and not outside:
            it.lat, it.lon = PLACES[DISTRICTS[d][0]]
            it.loc_level, it.loc_name, it.loc_src = "区域", f"{d}中心", src
            return it
    if district:
        it.lat, it.lon = dcenter
        it.loc_level, it.loc_name, it.loc_src = "区域", f"{district}中心", "分区"
        return it
    it.lat = it.lon = None
    it.loc_level = it.loc_name = it.loc_src = ""
    return it


def compute_distances(items, home):
    """home = (lat, lon, ...) 或 None; 设置每条帖子的 distance (km, 1 位小数)。"""
    for it in items:
        if home and it.lat is not None and it.lon is not None:
            it.distance = round(haversine(home, (it.lat, it.lon)), 1)
        else:
            it.distance = None
