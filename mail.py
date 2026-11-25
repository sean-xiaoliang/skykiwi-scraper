# -*- coding: utf-8 -*-
"""
邮件申请: 从招聘帖找邮箱, 生成写信链接, 打开邮件客户端或网页邮箱。

mailto 链接不支持附件, 所以打开写信窗口时可同时在资源管理器中选中简历文件, 方便拖进邮件。
"""

from __future__ import annotations

import os
import re
import subprocess
import webbrowser
from urllib.parse import quote, urlencode

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}")

WEBMAIL = {"gmail": "Gmail 网页版", "outlook": "Outlook.com 网页版"}
METHODS = {"auto": "自动 (有邮件客户端就用客户端, 否则用网页邮箱)", "client": "邮件客户端",
           "gmail": "Gmail 网页版", "outlook": "Outlook.com 网页版"}

DEFAULT_SUBJECT = "应聘：{职位}"
DEFAULT_BODY = ("您好，\n\n我在天维网看到贵司招聘「{职位}」的信息（{链接}），对这个职位很感兴趣，特来应聘。\n\n"
                "{简历摘要}\n\n我的简历请见附件，期待您的回复。谢谢！\n\n")


def extract_emails(text: str) -> list[str]:
    seen, out = set(), []
    for m in EMAIL_RE.finditer(text or ""):
        e = m.group(0).strip(".").lower()
        if e not in seen:
            seen.add(e)
            out.append(e)
    return out


def job_emails(it) -> list[str]:
    """帖子"邮箱"字段优先, 再从正文、其他字段里找。"""
    found = extract_emails(getattr(it, "email", "") or "")
    for text in [it.content] + list((it.extra or {}).values()):
        for e in extract_emails(text):
            if e not in found:
                found.append(e)
    return found


def fill(template: str, it, summary: str = "") -> str:
    values = {"职位": it.position or it.title, "公司": it.company or "", "标题": it.title,
              "链接": it.url, "简历摘要": summary or ""}
    out = template
    for k, v in values.items():
        out = out.replace("{" + k + "}", v)
    return out


def mailto_url(to: str, subject: str, body: str) -> str:
    body = body.replace("\r\n", "\n").replace("\n", "\r\n")
    return f"mailto:{quote(to, safe='@')}?subject={quote(subject, safe='')}&body={quote(body, safe='')}"


def webmail_url(provider: str, to: str, subject: str, body: str) -> str:
    if provider == "outlook":
        return "https://outlook.live.com/mail/0/deeplink/compose?" + urlencode(
            {"to": to, "subject": subject, "body": body}, quote_via=quote)
    return "https://mail.google.com/mail/?" + urlencode(
        {"view": "cm", "fs": "1", "to": to, "su": subject, "body": body}, quote_via=quote)


def detect_mail_client() -> str | None:
    """返回系统里处理 mailto 链接的程序标识; 没有则返回 None。"""
    try:
        import winreg
    except ImportError:
        return None
    for root, path, name in (
            (winreg.HKEY_CURRENT_USER,
             r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\mailto\UserChoice", "ProgId"),
            (winreg.HKEY_CLASSES_ROOT, r"mailto\shell\open\command", "")):
        try:
            with winreg.OpenKey(root, path) as key:
                val = str(winreg.QueryValueEx(key, name)[0]).strip()
                if val:
                    return val
        except OSError:
            continue
    return None


_CLIENT_NAMES = [("mailmaster", "网易邮箱大师"), ("foxmail", "Foxmail"), ("thunderbird", "Thunderbird"),
                 ("outlook", "Outlook"), ("windowsmail", "Windows 邮件"), ("microsoft.windowscommunicationsapps",
                                                                            "Windows 邮件"),
                 ("msedge", "浏览器 (Edge)"), ("chrome", "浏览器 (Chrome)"), ("firefox", "浏览器 (Firefox)")]


def client_name(client: str | None) -> str:
    """把注册表里的程序标识变成好认的名字。"""
    if not client:
        return "没有"
    low = client.lower()
    return next((name for key, name in _CLIENT_NAMES if key in low), client)


def choose_method(setting: str, client: str | None, webmail: str) -> str:
    """返回 'client' / 'gmail' / 'outlook'。"""
    if setting in WEBMAIL:
        return setting
    if client:
        return "client"
    return webmail if webmail in WEBMAIL else "gmail"


def compose(to: str, subject: str, body: str, method: str):
    """打开写信窗口。method 由 choose_method 得到。"""
    if method == "client":
        os.startfile(mailto_url(to, subject, body))  # type: ignore[attr-defined]
    else:
        webbrowser.open(webmail_url(method, to, subject, body))


def reveal_file(path: str):
    """在资源管理器中选中文件 (方便拖进邮件作为附件)。"""
    if path and os.path.exists(path):
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
