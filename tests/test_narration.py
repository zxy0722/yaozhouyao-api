# -*- coding: utf-8 -*-
"""智能导览（narrate / tts / qr / guide 页面）接口测试。

LLM 未配置时验证模板兜底；tts 仅测试参数校验（不联网）。
运行：py -m pytest tests -q
"""
from fastapi.testclient import TestClient

from main import app

client = TestClient(app)
BASE = "/api/v1"


def _data(resp):
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["code"] == 0, body
    return body


def test_narrate_template_mode():
    body = _data(client.get(f"{BASE}/narrate?q=釉色"))
    data = body["data"]
    assert data["question"] == "釉色"
    assert data["hit_count"] == 9
    assert data["mode"] in {"template", "llm"}
    if data["mode"] == "template":
        assert "橄榄绿" in data["script"]
        assert "为您整理了 9 条知识" in data["script"]
    assert data["sources"], "应返回知识来源"


def test_narrate_no_hit():
    body = _data(client.get(f"{BASE}/narrate?q=zzz不存在的关键词qqq"))
    data = body["data"]
    assert data["hit_count"] == 0
    assert data["script"] == ""


def test_narrate_limit():
    body = _data(client.get(f"{BASE}/narrate?q=窑&limit=3"))
    assert body["data"]["hit_count"] == 3


def test_narrate_validation():
    assert client.get(f"{BASE}/narrate").status_code == 422
    assert client.get(f"{BASE}/narrate?q=釉色&limit=99").status_code == 422


def test_tts_validation():
    # 空文本与非法音色应在联网前被拦截（无需真实合成）
    r = client.post(f"{BASE}/tts", json={"text": "   ", "voice": "zh-CN-XiaoxiaoNeural"})
    assert r.status_code == 400
    r = client.post(f"{BASE}/tts", json={"text": "你好", "voice": "bad-voice"})
    assert r.status_code == 400
    r = client.post(f"{BASE}/tts", json={"voice": "zh-CN-XiaoxiaoNeural"})
    assert r.status_code == 422
    r = client.post(f"{BASE}/tts", json={"text": "好" * 3000})
    assert r.status_code == 422


def test_qr_info_and_png():
    body = _data(client.get(f"{BASE}/qr-info"))
    data = body["data"]
    assert data["url"].startswith("http://")
    assert "qr_png_base64" in data and len(data["qr_png_base64"]) > 100
    resp = client.get(f"{BASE}/qr")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("image/png")
    assert resp.content[:8].startswith(b"\x89PNG")


def test_guide_page():
    resp = client.get("/guide")
    assert resp.status_code == 200
    assert "耀州窑" in resp.text


def test_root_reports_guide():
    body = client.get("/").json()
    assert body["guide"] == "/guide"
    assert "llm_narration" in body
