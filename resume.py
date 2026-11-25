# -*- coding: utf-8 -*-
"""
简历分析与岗位匹配 (完全在本机进行, 不上传任何内容)。

- extract_text(path): 读取 PDF / Word(.docx) / TXT 简历的文字
- analyze(text) -> Profile: 识别技能方向、关键词、语言、证照、签证、工作年限
- match(profile, listing) -> 0~100: 岗位与简历的匹配度
"""

from __future__ import annotations

import os
import re
import zipfile
from dataclasses import dataclass, field


class ResumeError(Exception):
    pass


# ---------------------------------------------------------------- 读取文字 ----

def extract_text(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if not os.path.exists(path):
        raise ResumeError(f"找不到文件: {path}")
    try:
        if ext == ".pdf":
            from pypdf import PdfReader
            text = "\n".join((p.extract_text() or "") for p in PdfReader(path).pages)
        elif ext == ".docx":
            with zipfile.ZipFile(path) as z:
                xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            xml = re.sub(r"</w:p>|<w:br/>|<w:tab/>", "\n", xml)
            text = re.sub(r"<[^>]+>", "", xml)
            text = (text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
                    .replace("&quot;", '"').replace("&apos;", "'"))
        elif ext in (".txt", ".md"):
            with open(path, "rb") as f:
                raw = f.read()
            for enc in ("utf-8-sig", "gb18030", "big5", "latin-1"):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
        elif ext == ".doc":
            raise ResumeError("不支持旧版 .doc 格式, 请在 Word 中另存为 .docx 或 PDF")
        else:
            raise ResumeError(f"不支持的格式 {ext or '(无扩展名)'}; 支持 PDF、Word (.docx)、TXT")
    except ResumeError:
        raise
    except Exception as e:  # noqa: BLE001
        raise ResumeError(f"读取简历失败: {type(e).__name__}: {e}") from e
    text = re.sub(r"[ \t　]+", " ", text)
    if len(text.strip()) < 10:
        raise ResumeError("简历里读不到文字 (可能是扫描图片版 PDF), 请换成可复制文字的 PDF 或 Word")
    return text.strip()


# ---------------------------------------------------------------- 分析 ----

# 技能方向 -> 关键词 (键名与天维网岗位类别一致的, 匹配时有类别加分)
SKILLS = {
    "服务人员": ["服务员", "waiter", "waitress", "barista", "咖啡", "前厅", "kitchen hand", "厨房", "帮厨", "厨师",
             "chef", "cook", "炒锅", "寿司", "sushi", "奶茶", "餐厅", "restaurant", "cafe", "收银", "cashier",
             "零售", "retail", "店员", "理货", "超市"],
    "清洁/家政": ["清洁", "cleaner", "cleaning", "家政", "保洁", "housekeeping", "阿姨", "育儿", "nanny", "洗碗"],
    "技术类工种": ["电工", "electrician", "水工", "plumber", "木工", "carpenter", "builder", "建筑", "装修", "批灰",
              "plaster", "油漆", "painter", "瓷砖", "tiler", "焊", "weld", "汽修", "mechanic", "叉车", "forklift",
              "司机", "driver", "卡车", "truck", "挖机", "仓库", "warehouse", "picker", "packer", "物流", "工厂",
              "factory", "美容", "美甲", "按摩", "massage", "针灸", "理发", "hairdress"],
    "销售/市场": ["销售", "sales", "市场", "marketing", "客户", "customer", "电商", "运营", "直播", "推广"],
    "前台/行政": ["前台", "reception", "行政", "admin", "文员", "clerk", "助理", "assistant", "办公", "office"],
    "财会/金融": ["会计", "accountant", "accounting", "财务", "finance", "出纳", "xero", "myob", "审计", "audit",
              "报税", "bookkeep"],
    "编辑/文职": ["编辑", "editor", "文案", "copywrit", "翻译", "translat", "设计", "design", "photoshop"],
    "教育/培训": ["教师", "老师", "teacher", "tutor", "教育", "培训", "teaching", "幼教", "early childhood"],
    "公关/顾问": ["顾问", "consultant", "公关", "移民", "留学", "保险", "insurance", "房产", "real estate"],
    "IT / 软件": ["python", "java", "javascript", "sql", "developer", "开发", "程序", "软件", "software",
                "it support", "网络", "全栈", "前端", "后端"],
}
LANGUAGES = {"英语": ["英语", "英文", "english", "ielts", "雅思", "toefl"], "普通话": ["普通话", "mandarin", "中文"],
             "粤语": ["粤语", "广东话", "cantonese"]}
LICENCES = {"驾照": ["驾照", "驾驶执照", "driver licence", "drivers licence", "driver's licence", "full licence",
                   "restricted licence"],
            "叉车证": ["叉车证", "forklift licence", "forklift license", "f endorsement"],
            "安全证": ["site safe", "工地安全"]}
VISAS = [("公民/永居", ["公民", "永居", "citizen", "resident visa", "permanent resident", "绿卡", " pr "]),
         ("开放工签", ["开放工签", "open work visa", "open work"]),
         ("工签", ["工签", "工作签证", "work visa", "aewv"]),
         ("打工度假", ["打工度假", "whv", "working holiday"]),
         ("学签", ["学签", "学生签证", "student visa"])]


@dataclass
class Profile:
    categories: list = field(default_factory=list)        # 技能方向, 按命中数排序
    keywords: list = field(default_factory=list)          # 命中的技能关键词
    languages: list = field(default_factory=list)
    licences: list = field(default_factory=list)
    visa: str = ""
    years: int | None = None
    cat_hits: dict = field(default_factory=dict)

    def summary(self) -> str:
        """用于邮件正文的一句话简介。"""
        parts = []
        if self.categories:
            exp = f"{self.years} 年" if self.years else ""
            parts.append(f"我有{exp}{'、'.join(self.categories[:2])}方面的工作经验")
        if self.languages:
            parts.append(f"会说{'、'.join(self.languages)}")
        if self.licences:
            parts.append(f"持有{'、'.join(self.licences)}")
        if self.visa:
            parts.append(f"签证状态为{self.visa}")
        return ("，".join(parts) + "。") if parts else ""


def _norm(text: str) -> str:
    return " " + re.sub(r"\s+", " ", (text or "").lower()) + " "


def _has(t: str, words) -> list:
    return [w for w in words if w in t]


def analyze(text: str) -> Profile:
    t = _norm(text)
    p = Profile()
    for cat, words in SKILLS.items():
        hits = _has(t, words)
        if hits:
            p.cat_hits[cat] = hits
            p.keywords += [w for w in hits if w not in p.keywords]
    p.categories = sorted(p.cat_hits, key=lambda c: -len(p.cat_hits[c]))
    p.languages = [name for name, words in LANGUAGES.items() if _has(t, words)]
    p.licences = [name for name, words in LICENCES.items() if _has(t, words)]
    p.visa = next((name for name, words in VISAS if _has(t, words)), "")
    yrs = [int(m.group(1)) for m in re.finditer(r"(\d{1,2})\s*\+?\s*(?:年|years?|yrs?)", t)]
    yrs = [y for y in yrs if 0 < y <= 45]
    p.years = max(yrs) if yrs else None
    return p


# ---------------------------------------------------------------- 匹配 ----

_NEEDS_ENGLISH = re.compile(r"英语|英文|english", re.I)
_NEEDS_LICENCE = re.compile(r"驾照|有车|开车|licence|license|driver", re.I)


def match(profile: Profile, it) -> int:
    """0~100。类别吻合 +35, 每个技能关键词 +15 (最多 45), 满足/不满足语言和驾照要求 ±10。"""
    if not profile or not profile.keywords:
        return 0
    if getattr(it, "sort", "") == "求职":          # 求职帖是别人在找工作, 不是岗位
        return 0
    job = _norm(" ".join(str(x or "") for x in (it.title, it.position, it.company, it.content)))
    score = 0
    if it.region in profile.cat_hits:
        score += 35
    hits = [w for w in profile.keywords if w in job]
    score += min(45, 15 * len(hits))
    if _NEEDS_ENGLISH.search(job):
        score += 10 if "英语" in profile.languages else -10
    if _NEEDS_LICENCE.search(job):
        score += 10 if "驾照" in profile.licences else -10
    if hits or it.region in profile.cat_hits:
        score += 10
    return max(0, min(100, score))
