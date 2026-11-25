# -*- coding: utf-8 -*-
"""定位回归测试 (用例来自真实帖子审计)。运行: python test_geo.py"""
from types import SimpleNamespace

import geo

CBD = (-36.8472, 174.7655)
HAMILTON = (-37.7879, 175.2818)


def L(category="rent", region="", title="", address="", location="", content=""):
    return SimpleNamespace(category=category, region=region, title=title, address=address,
                           location=location, content=content, lat=None, lon=None,
                           loc_level="", loc_name="", distance=None)


class FakeGeocoder:
    """按查询词中包含的关键字返回预设候选 [(lat, lon, name), ...]。"""

    def __init__(self, table):
        self.table = table
        self.queries = []

    def search(self, q):
        self.queries.append(q)
        for key, res in self.table.items():
            if key in q.lower():
                return res
        return []

    def save(self):
        pass


def km(it, p):
    assert it.lat is not None, f"未定位: {it.loc_level} {it.loc_name}"
    return geo.haversine((it.lat, it.lon), p)


# ---- 1. 街道名不能当成同名城市/街区 ----
def test_street_name_not_city():
    it = geo.locate(L(region="市中心", title="CBD市中心2b1b低价短租 可续租", address="Nelson Street"))
    assert it.loc_name != "Nelson", it.loc_name
    assert km(it, CBD) < 2


def test_numbered_street_geocoded_with_auckland():
    gc = FakeGeocoder({"nelson street": [(-36.8510, 174.7598, "72 Nelson Street, City Centre, Auckland")]})
    it = geo.locate(L(region="市中心", title="市中心Nelson St两房一卫公寓出租！", address="L4/72 Nelson Street"), gc)
    assert it.loc_level == "地址", (it.loc_level, it.loc_name)
    assert km(it, CBD) < 1
    assert all("auckland" in q.lower() for q in gc.queries), gc.queries


# ---- 2. 正文中的 去/到/直达/排除 等不是房源位置 ----
def test_travel_reference_ignored_when_better_exists():
    it = geo.locate(L(region="其他地区", title="位置，Hamiltom Rototuna", address="Rototuna",
                      content="有安全停车位,附近有多路公交车站,直达大学中学,市中心,生活学习很方便"))
    assert km(it, HAMILTON) < 10, it.loc_name


def test_negated_place_ignored():
    it = geo.locate(L(region="北岸", title="求租单间，可合租",
                      content="北岸、南区、东区等都可以考虑,但排除cbd区域。"))
    assert it.loc_name != "Auckland CBD"
    assert km(it, CBD) > 5


def test_travel_only_reference_used_as_fallback():
    it = geo.locate(L(region="东区", title="已租",
                      content="已租 房子在无尾静街,从家里开车去ormiston mall大概3分钟,步行去公车站"))
    assert "Ormiston" in it.loc_name, it.loc_name


def test_apostrophe_st_lukes_title_beats_travel_in_content():
    it = geo.locate(L(region="中区", title="16号高速st Luke’s 出口两分钟一独立出入独立卫浴出租",
                      content="开车去city 也就七分钟,去st lukes mall 三分钟"))
    assert it.loc_name == "Mount Albert", it.loc_name


# ---- 3. 同名街道: 选与上下文最接近的候选 ----
def test_ambiguous_street_picks_candidate_near_context():
    gc = FakeGeocoder({"liverpool": [(-36.8887, 174.7783, "11 Liverpool Street, Epsom"),
                                     (-36.8565, 174.7620, "11 Liverpool Street, City Centre")]})
    it = geo.locate(L(region="市中心", title="(女)*****高層海景雙人房-奧大,AUT @ 11 Liverpool St",
                      address="11 Liverpool st Auckland"), gc)
    assert it.loc_level == "地址"
    assert km(it, CBD) < 1.5


def test_street_without_number_geocoded():
    gc = FakeGeocoder({"may road": [(-36.9135, 174.7430, "May Road, Mount Roskill")],
                       "may rd": [(-36.9135, 174.7430, "May Road, Mount Roskill")]})
    it = geo.locate(L(region="中区", title="MAY RD 有主人房（双人）出租", address="may rd"), gc)
    assert it.loc_level == "街道", (it.loc_level, it.loc_name)
    assert km(it, (-36.9135, 174.7430)) < 0.5


def test_geocode_far_from_district_rejected():
    gc = FakeGeocoder({"hogans": [(-45.0, 170.0, "Hogans Road, Otago")]})
    it = geo.locate(L(region="北岸", title="双人大间招合租", address="hogans Rd"), gc)
    assert it.loc_level == "区域", (it.loc_level, it.loc_name)


def test_query_cleaning():
    assert "l3" not in geo._clean_query("L3/135 Hobson Street").lower()
    assert "135 hobson street" in geo._clean_query("L3/135 Hobson Street").lower()
    assert "11 union street" in geo._clean_query("11union street").lower()
    assert "30 kensington ave" in geo._clean_query("flat3, 30 Kensington Ave, Mount Eden").lower()


# ---- 4. 繁体字 / 补充地名 / 坐标修正 ----
def test_traditional_chinese():
    assert geo.match_region("東區租房") == "Botany Downs"
    assert geo.match_place("高層海景雙人房-奧大") == "University of Auckland"


def test_new_places_and_aliases():
    assert geo.match_place("地址:brookby 奥克兰东区") == "Brookby"
    assert geo.match_place("倒霉路附近3房1厅整租") == "Dominion Road"
    assert geo.match_place("Dominion Rd 益家超市附近") == "Dominion Road"
    assert geo.match_place("Dominion 路亚洲超市近旁") == "Dominion Road"
    assert geo.match_place("中区Unitec旁阳光双人房") == "Unitec"
    assert geo.match_place("植物园附近") == "Auckland Botanic Gardens"


def test_hill_park_coordinate():
    assert geo.haversine(geo.PLACES["Hill Park"], geo.PLACES["Manurewa"]) < 3


# ---- 5. 与所在分区明显矛盾的匹配要拒绝 ----
def test_district_consistency():
    it = geo.locate(L(region="北岸", title="招聘市场推广", content="南岛城际交通: 基督城 – 皇后镇等热门旅游目的地"))
    assert it.loc_name not in ("Christchurch", "Queenstown"), it.loc_name


def test_job_location_field():
    it = geo.locate(L(category="job", location="奥克兰北岸", title="奧克蘭北岸注册针灸师",
                      content="albany healthcare centre是新西兰认证雇主"))
    assert it.loc_name == "Albany", it.loc_name
    it = geo.locate(L(category="job", location="其他地区", title="中餐厅", content="位于市中心,交通方便"))
    assert it.loc_name != "Auckland CBD", it.loc_name


# ---- 6. 家庭住址同名提示 ----
def test_home_ambiguity_warning():
    gc = FakeGeocoder({"queen street": [(-36.9451, 174.8437, "12 Queen Street, Otahuhu"),
                                        (-36.8485, 174.7650, "12 Queen Street, City Centre")]})
    r = geo.resolve_home("12 Queen Street, Auckland", gc)
    assert r and "同名" in r[2], r


# ---- 7. 真实 Geocoder 的接口 (不联网) ----
def test_real_geocoder_interface():
    g = geo.Geocoder()
    for name in ("search", "geocode", "reverse", "save"):
        assert callable(getattr(g, name, None)), f"Geocoder 缺少 {name}()"
    assert callable(geo.locate_device)


if __name__ == "__main__":
    import sys
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
