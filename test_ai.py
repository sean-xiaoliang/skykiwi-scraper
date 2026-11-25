# -*- coding: utf-8 -*-
"""AI (Gemini) 与距离检索测试。不联网: 用假的 HTTP 会话模拟 Gemini 的返回。运行: python test_ai.py"""
import json
import os
import sys
import tempfile

os.environ["SKYKIWI_APP_DIR"] = tempfile.mkdtemp(prefix="skykiwi_test_")

import skykiwi_scraper as k  # noqa: E402
import ai  # noqa: E402
import secret  # noqa: E402
import store  # noqa: E402
import resume  # noqa: E402

KEY = "AIzaSyTEST-not-a-real-key-1234567890"


class FakeResp:
    def __init__(self, status, data):
        self.status_code, self._data = status, data
        self.text = json.dumps(data, ensure_ascii=False)

    def json(self):
        return self._data


class FakeSession:
    """按顺序返回预设响应, 并记录每次请求。"""

    def __init__(self, *responses):
        self.responses = list(responses)
        self.calls = []
        self.headers = {}

    def _next(self, method, url, **kw):
        self.calls.append({"method": method, "url": url, **kw})
        return self.responses.pop(0)

    def post(self, url, **kw):
        return self._next("POST", url, **kw)

    def get(self, url, **kw):
        return self._next("GET", url, **kw)


def ok_json(obj):
    return FakeResp(200, {"candidates": [{"content": {"parts": [{"text": json.dumps(obj, ensure_ascii=False)}]},
                                          "finishReason": "STOP"}]})


def err(code, status, message):
    return FakeResp(code, {"error": {"code": code, "status": status, "message": message}})


# ==================== API Key 加密保存 ====================

def test_secret_roundtrip_not_plaintext():
    s = secret.protect(KEY)
    assert KEY not in s and "AIza" not in s
    assert secret.unprotect(s) == KEY


def test_api_key_saved_encrypted_in_config():
    ai.save_api_key(KEY)
    with open(os.path.join(os.environ["SKYKIWI_APP_DIR"], "config.json"), encoding="utf-8") as f:
        raw = f.read()
    assert KEY not in raw, "配置文件里出现了明文 API Key"
    assert ai.get_api_key() == KEY
    ai.save_api_key("")
    assert ai.get_api_key() == ""


# ==================== 隐私: 发送前去掉联系方式 ====================

def test_redact_personal_info():
    text = "张三 电话 021 123 4567 / +64 21 1234567，邮箱 zs@example.com，微信: zhang_san88。厨师 5 年。"
    r = ai.redact(text)
    for s in ("021 123 4567", "1234567", "zs@example.com", "zhang_san88"):
        assert s not in r, (s, r)
    assert "厨师 5 年" in r


# ==================== 选择模型 ====================

def test_pick_newest_flash_model():
    names = ["gemini-2.5-flash", "gemini-3.6-flash", "gemini-3.8-flash", "gemini-3.8-flash-lite",
             "gemini-3.8-pro", "gemini-embedding-001", "gemini-3.9-flash-image-preview", "gemini-3.9-flash-tts"]
    assert ai.pick_model(names) == "gemini-3.8-flash"
    assert ai.pick_model(["gemini-3.1-pro"]) == "gemini-3.1-pro"      # 没有 flash 时退而求其次
    assert ai.pick_model([]) is None


def test_list_models_parses_generate_content_models():
    s = FakeSession(FakeResp(200, {"models": [
        {"name": "models/gemini-3.8-flash", "supportedGenerationMethods": ["generateContent", "countTokens"]},
        {"name": "models/gemini-embedding-001", "supportedGenerationMethods": ["embedContent"]}]}))
    g = ai.Gemini(KEY, session=s)
    assert g.list_models() == ["gemini-3.8-flash"]
    assert KEY not in s.calls[0]["url"]


# ==================== 请求格式 / 返回解析 / 错误 ====================

def test_generate_json_request_shape():
    s = FakeSession(ok_json({"a": 1}))
    g = ai.Gemini(KEY, model="gemini-3.8-flash", session=s)
    assert g.generate_json("你好", {"type": "object"}) == {"a": 1}
    c = s.calls[0]
    assert c["url"].endswith("/v1beta/models/gemini-3.8-flash:generateContent"), c["url"]
    assert KEY not in c["url"], "API Key 不应出现在 URL 中 (会被写进日志)"
    assert c["headers"]["x-goog-api-key"] == KEY
    body = c["json"]
    assert body["contents"][0]["parts"][0]["text"] == "你好"
    assert body["generationConfig"]["responseMimeType"] == "application/json"
    assert body["generationConfig"]["responseJsonSchema"] == {"type": "object"}


def test_generate_json_auto_picks_model():
    s = FakeSession(FakeResp(200, {"models": [{"name": "models/gemini-3.8-flash",
                                               "supportedGenerationMethods": ["generateContent"]}]}),
                    ok_json({"x": 1}))
    g = ai.Gemini(KEY, model="", session=s)
    g.generate_json("hi", {"type": "object"})
    assert g.model == "gemini-3.8-flash" and ":generateContent" in s.calls[1]["url"]


def test_generate_json_handles_fences_and_parts():
    resp = FakeResp(200, {"candidates": [{"content": {"parts": [
        {"text": "思考过程", "thought": True}, {"text": '```json\n{"ok": true}\n```'}]}, "finishReason": "STOP"}]})
    g = ai.Gemini(KEY, model="m", session=FakeSession(resp))
    assert g.generate_json("x", {}) == {"ok": True}


def _error_message(resp):
    g = ai.Gemini(KEY, model="gemini-3.8-flash", session=FakeSession(resp))
    try:
        g.generate_json("x", {})
    except ai.GeminiError as e:
        assert KEY not in str(e)
        return str(e)
    raise AssertionError("应抛出 GeminiError")


def test_error_messages_are_friendly():
    assert "API Key" in _error_message(err(400, "INVALID_ARGUMENT", "API key not valid. Please pass a valid API key."))
    assert "额度" in _error_message(err(429, "RESOURCE_EXHAUSTED", "Quota exceeded"))
    assert "模型" in _error_message(err(404, "NOT_FOUND", "models/gemini-3.8-flash is not found"))
    assert "暂时不可用" in _error_message(err(503, "UNAVAILABLE", "overloaded"))
    assert "拦截" in _error_message(FakeResp(200, {"promptFeedback": {"blockReason": "SAFETY"}}))
    assert "不完整" in _error_message(FakeResp(200, {"candidates": [{"content": {"parts": [{"text": '{"a":'}]},
                                                                      "finishReason": "MAX_TOKENS"}]}))


# ==================== AI 简历分析 ====================

AI_RESUME = {"summary": "有 5 年中餐厨师经验，熟悉炒锅和寿司，持开放工签。", "desired_roles": ["厨师", "厨房帮工"],
             "categories": ["服务人员", "不存在的类别"], "skills": ["炒锅", "Sushi", "厨师"],
             "languages": ["英语", "普通话"], "licences": ["驾照"], "visa": "开放工签", "years": 5,
             "strengths": ["经验丰富"], "suggestions": ["补充英文简历"]}


def test_analyze_resume_with_ai():
    s = FakeSession(ok_json(AI_RESUME))
    g = ai.Gemini(KEY, model="m", session=s)
    prof = ai.analyze_resume(g, "张三 电话 021 123 4567 在中餐厅做厨师 5 年")
    sent = s.calls[0]["json"]["contents"][0]["parts"][0]["text"]
    assert "021 123 4567" not in sent, "简历电话被发送出去了"
    assert prof.source == "AI"
    assert prof.categories == ["服务人员"], prof.categories          # 只保留论坛已有的类别
    assert "sushi" in prof.keywords and "炒锅" in prof.keywords
    assert prof.years == 5 and prof.visa == "开放工签"
    assert prof.summary().startswith("有 5 年中餐厨师经验")
    assert prof.suggestions == ["补充英文简历"]


# ==================== AI 评估岗位 ====================

def _jobs(n):
    return [k.Listing(tid=i, category="job", title=f"岗位{i}", sort="招聘", region="服务人员",
                      position="厨师", content="x" * 1000) for i in range(1, n + 1)]


def test_score_jobs_in_batches():
    def batch(ids):
        return ok_json({"results": [{"id": i, "score": 50 + i, "reason": f"理由{i}"} for i in ids]})
    s = FakeSession(batch(range(1, 21)), batch(range(21, 41)), batch([41, 42, 43, 44, 999]))
    g = ai.Gemini(KEY, model="m", session=s)
    prof = resume.analyze("厨师 5 年")
    res = ai.score_jobs(g, prof, _jobs(45), batch_size=20)
    assert len(s.calls) == 3, len(s.calls)
    assert res[1] == (51, "理由1") and res[44] == (94, "理由44")
    assert 45 not in res and 999 not in res, "缺失的或多余的 id 要忽略"
    assert all(0 <= v[0] <= 100 for v in res.values())
    assert len(s.calls[0]["json"]["contents"][0]["parts"][0]["text"]) < 20000, "岗位正文应截断"


def test_ai_scores_cached_in_store():
    c = store.PostCache(os.path.join(tempfile.mkdtemp(), "p.db"))
    c.save_ai_scores("prof1", {1: (80, "合适"), 2: (20, "不合适")})
    assert c.load_ai_scores("prof1") == {1: (80, "合适"), 2: (20, "不合适")}
    assert c.load_ai_scores("other") == {}


# ==================== 匹配度星级 ====================

def test_match_stars():
    assert resume.stars(95) == "★★★★★"
    assert resume.stars(72) == "★★★★☆"
    assert resume.stars(55) == "★★★☆☆"
    assert resume.stars(40) == "★★☆☆☆"
    assert resume.stars(20) == "★☆☆☆☆"
    assert resume.stars(5) == "☆☆☆☆☆"
    assert resume.stars(None) == ""


# ==================== 距离检索 ====================

def _loc(tid, d, level):
    it = k.Listing(tid=tid, category="rent", title=str(tid), sort="出租")
    it.distance, it.loc_level = d, level
    return it


def test_precise_only_filter():
    items = [_loc(1, 2.0, "地址"), _loc(2, 3.0, "街区"), _loc(3, 1.0, "区域"), _loc(4, None, "")]
    got = k.apply_filters(items, k.Criteria(precise_only=True))
    assert [it.tid for it in got] == [1, 2]
    got = k.apply_filters(items, k.Criteria(max_distance=2.5, precise_only=False))
    assert sorted(it.tid for it in got) == [1, 3]


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
