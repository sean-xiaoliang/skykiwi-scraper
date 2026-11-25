# -*- coding: utf-8 -*-
"""
本地帖子缓存 (SQLite, 存在数据目录的 posts.db)。

- 每个帖子以 JSON 保存全部字段 (含定位结果), 按 帖子ID 去重
- meta 表保存 "某类别已完整抓取到哪一天" 等信息, 用于增量刷新时判断何时停止翻页
- 每次操作单独开连接, 可在后台线程中安全使用
"""

from __future__ import annotations

import dataclasses
import datetime as dt
import json
import os
import sqlite3
import threading
from contextlib import contextmanager

import geo

_lock = threading.Lock()


def default_path() -> str:
    return os.path.join(geo.APP_DIR, "posts.db")


class PostCache:
    def __init__(self, path: str | None = None):
        self.path = path or default_path()
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with self._db() as con:
            con.execute("""CREATE TABLE IF NOT EXISTS posts (
                               tid INTEGER PRIMARY KEY, category TEXT NOT NULL, post_date TEXT,
                               detail_ok INTEGER, data TEXT NOT NULL, updated_at TEXT)""")
            con.execute("CREATE INDEX IF NOT EXISTS idx_posts_cat_date ON posts(category, post_date)")
            con.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")

    @contextmanager
    def _db(self):
        """打开连接 -> 事务 (成功提交 / 出错回滚) -> 关闭连接。"""
        con = sqlite3.connect(self.path, timeout=10)
        try:
            with con:
                yield con
        finally:
            con.close()

    # ---- 帖子 ----
    @staticmethod
    def _encode(it) -> str:
        d = dataclasses.asdict(it)
        d.pop("distance", None)              # 距离随家庭住址变化, 不缓存
        return json.dumps(d, ensure_ascii=False)

    @staticmethod
    def _decode(s: str):
        from skykiwi_scraper import Listing
        d = json.loads(s)
        names = {f.name for f in dataclasses.fields(Listing)}
        return Listing(**{k: v for k, v in d.items() if k in names})

    def save(self, items):
        now = dt.datetime.now().isoformat(timespec="seconds")
        rows = [(it.tid, it.category, it.post_date, int(bool(it.detail_ok)), self._encode(it), now) for it in items]
        if not rows:
            return
        with _lock, self._db() as con:
            con.executemany("INSERT OR REPLACE INTO posts VALUES (?,?,?,?,?,?)", rows)

    def load(self, category: str, days: int) -> list:
        cutoff = (dt.date.today() - dt.timedelta(days=max(1, int(days)) - 1)).isoformat()
        with self._db() as con:
            rows = con.execute("SELECT data FROM posts WHERE category=? AND post_date>=? "
                               "ORDER BY post_date DESC, tid DESC", (category, cutoff)).fetchall()
        return [self._decode(r[0]) for r in rows]

    def known(self, category: str) -> dict:
        """{帖子ID: 是否已读取过详情}"""
        with self._db() as con:
            return {tid: bool(ok) for tid, ok in
                    con.execute("SELECT tid, detail_ok FROM posts WHERE category=?", (category,))}

    def update_counts(self, rows):
        """更新已缓存帖子的 回复/查看 数: rows = [(tid, replies, views)]"""
        with _lock, self._db() as con:
            for tid, replies, views in rows:
                r = con.execute("SELECT data FROM posts WHERE tid=?", (tid,)).fetchone()
                if not r:
                    continue
                d = json.loads(r[0])
                d["replies"], d["views"] = replies, views
                con.execute("UPDATE posts SET data=? WHERE tid=?", (json.dumps(d, ensure_ascii=False), tid))

    def prune(self, keep_days: int = 60):
        cutoff = (dt.date.today() - dt.timedelta(days=keep_days)).isoformat()
        with _lock, self._db() as con:
            con.execute("DELETE FROM posts WHERE post_date < ?", (cutoff,))

    def clear(self):
        with _lock, self._db() as con:
            con.execute("DELETE FROM posts")
            con.execute("DELETE FROM meta")

    def count(self, category: str) -> int:
        with self._db() as con:
            return con.execute("SELECT COUNT(*) FROM posts WHERE category=?", (category,)).fetchone()[0]

    # ---- 元信息 ----
    def get_meta(self, key: str):
        with self._db() as con:
            r = con.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return r[0] if r else None

    def set_meta(self, key: str, value: str):
        with _lock, self._db() as con:
            con.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", (key, value))
