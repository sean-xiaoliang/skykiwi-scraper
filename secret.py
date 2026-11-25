# -*- coding: utf-8 -*-
"""
用 Windows DPAPI 加密保存敏感信息 (如 API Key)。

加密结果只有当前 Windows 用户在本机能解开; 配置文件被复制到别处也读不出明文。
"""

from __future__ import annotations

import base64
import ctypes
from ctypes import wintypes

PREFIX = "dpapi:"


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _call(func, data: bytes) -> bytes:
    buf = ctypes.create_string_buffer(data, len(data))
    inp = _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    if not func(ctypes.byref(inp), None, None, None, None, 0x1, ctypes.byref(out)):   # CRYPTPROTECT_UI_FORBIDDEN
        raise OSError(ctypes.WinError())
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)


def protect(text: str) -> str:
    if not text:
        return ""
    crypt32 = ctypes.windll.crypt32
    return PREFIX + base64.b64encode(_call(crypt32.CryptProtectData, text.encode("utf-8"))).decode("ascii")


def unprotect(value: str) -> str:
    if not value:
        return ""
    if not value.startswith(PREFIX):
        return value
    crypt32 = ctypes.windll.crypt32
    try:
        return _call(crypt32.CryptUnprotectData, base64.b64decode(value[len(PREFIX):])).decode("utf-8")
    except (OSError, ValueError):
        return ""        # 换了用户 / 电脑, 无法解密: 视为未设置
