# -*- coding: utf-8 -*-
"""API 冒烟测试。运行：py -m pytest tests -q"""
from fastapi.testclient import TestClient

from main import app

client = TestClient(app)
BASE = "/api/v1"
AID = "artifact-00143910"


def _data(resp):
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["code"] == 0, body
    return body


def test_root_and_health():
    assert client.get("/health").json()["status"] == "ok"
    body = client.get("/").json()
    assert body["docs"] == "/docs"


def test_list_artifacts():
    body = _data(client.get(f"{BASE}/artifacts"))
    assert len(body["data"]) == 1
    art = body["data"][0]
    assert art["name"] == "耀州窑青釉刻花瓶"
    assert art["dynasty"] == "北宋"
    assert art["dimensions_cm"]["height"] == 19.9
    assert art["record_count"] == 87


def test_artifact_detail():
    body = _data(client.get(f"{BASE}/artifacts/{AID}"))
    detail = body["data"]
    assert len(detail["records_by_section"]) == 3
    assert detail["process_steps"][0] == "采料"
    assert len(detail["process_steps"]) == 17
    assert detail["burn_temperature_c"] == {"min": 1280.0, "max": 1300.0}
    assert client.get(f"{BASE}/artifacts/nope").status_code == 404


def test_records_pagination_and_filter():
    body = _data(client.get(f"{BASE}/artifacts/{AID}/records?size=10"))
    assert body["meta"]["total"] == 87
    assert len(body["data"]["items"]) == 10
    assert body["data"]["pages"] == 9

    body = _data(client.get(f"{BASE}/artifacts/{AID}/records?category=工艺流程"))
    assert body["meta"]["total"] == 11

    body = _data(client.get(f"{BASE}/artifacts/{AID}/records?section=文物本体参数"))
    assert body["meta"]["total"] == 23


def test_records_keyword():
    body = _data(client.get(f"{BASE}/artifacts/{AID}/records?q=牡丹"))
    assert body["meta"]["total"] >= 5


def test_record_detail_and_404():
    body = _data(client.get(f"{BASE}/artifacts/{AID}/records/24"))
    assert body["data"]["field"] == "总工序数"
    assert client.get(f"{BASE}/artifacts/{AID}/records/9999").status_code == 404


def test_categories():
    body = _data(client.get(f"{BASE}/artifacts/{AID}/categories"))
    cats = {c["category"] for c in body["data"]}
    assert {"基本信息", "工艺流程", "鉴定特征"} <= cats
    assert sum(c["count"] for c in body["data"]) == 87


def test_search():
    body = _data(client.get(f"{BASE}/search?q=缠枝牡丹"))
    assert len(body["data"]) >= 1
    assert body["data"][0]["artifact_id"] == AID


def test_stats():
    body = _data(client.get(f"{BASE}/stats"))
    stats = body["data"]
    assert stats["record_count"] == 87
    assert stats["artifact_count"] == 1
    assert {s["section"] for s in stats["sections"]} == {
        "文物本体参数",
        "工艺·纹饰·历史背景",
        "对比·鉴定·价值",
    }
    assert stats["top_sources"], "来源分布不应为空"
