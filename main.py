# -*- coding: utf-8 -*-
"""耀州窑青釉刻花瓶知识 API（数据来源：故宫博物院等公开资料）。

启动：py -m uvicorn main:app --reload
文档：http://127.0.0.1:8000/docs
"""
import base64
import hashlib
import io
import json
import math
import os
import socket
from pathlib import Path
from typing import Optional

import edge_tts
import httpx
import qrcode
from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response

from schemas import (
    Artifact,
    ArtifactDetail,
    ArtifactSummary,
    CategoryCount,
    CategoryTotal,
    Envelope,
    Narration,
    QRInfo,
    Record,
    RecordPage,
    SearchHit,
    SectionCount,
    SourceCount,
    Stats,
    TTSRequest,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "dataset.json"
GUIDE_PAGE_FILE = BASE_DIR / "guide.html"

# ---------- 可选配置：大模型讲解（任意 OpenAI 兼容接口） ----------
# 设置 LLM_BASE_URL / LLM_API_KEY / LLM_MODEL 三个环境变量后自动启用；
# 未设置时使用知识库模板兜底，服务照常可用。
LLM_BASE_URL = os.getenv("LLM_BASE_URL", "").rstrip("/")
LLM_API_KEY = os.getenv("LLM_API_KEY", "")
LLM_MODEL = os.getenv("LLM_MODEL", "")

# ---------- 语音合成（微软 Edge 神经语音，免密钥） ----------
TTS_VOICES = {
    "zh-CN-XiaoxiaoNeural": "晓晓（女声 · 温暖亲切）",
    "zh-CN-YunxiNeural": "云希（男声 · 沉稳清晰）",
    "zh-CN-YunyangNeural": "云扬（男声 · 播音浑厚）",
}
_TTS_CACHE: dict[str, bytes] = {}
_TTS_CACHE_MAX = 50

# ---------- 数据加载与校验 ----------
_raw = json.loads(DATA_FILE.read_text(encoding="utf-8"))
ARTIFACT = Artifact(**_raw["artifact"])
RECORDS: list[Record] = [Record(**r) for r in _raw["records"]]
ARTIFACTS: dict[str, Artifact] = {ARTIFACT.id: ARTIFACT}
RECORDS_BY_ARTIFACT: dict[str, list[Record]] = {}
for _r in RECORDS:
    RECORDS_BY_ARTIFACT.setdefault(_r.artifact_id, []).append(_r)

app = FastAPI(
    title="耀州窑青釉刻花瓶知识 API",
    description=(
        "基于《耀州窑青釉刻花瓶》Excel 数据集生成的文物知识 REST API。\n\n"
        "数据来源：故宫博物院、陕西省地方志办公室、耀州窑考古研究等公开资料。"
    ),
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def add_cache_headers(request, call_next):
    response = await call_next(request)
    if request.method == "GET" and response.status_code == 200:
        response.headers.setdefault("Cache-Control", "public, max-age=300")
    return response


@app.exception_handler(HTTPException)
async def http_error_envelope(request, exc: HTTPException):
    return JSONResponse(
        status_code=exc.status_code,
        content={"code": exc.status_code, "message": str(exc.detail), "data": None, "meta": None},
    )


def ok(data, meta=None):
    return {"code": 0, "message": "ok", "data": data, "meta": meta}


def get_artifact_or_404(aid: str) -> Artifact:
    artifact = ARTIFACTS.get(aid)
    if artifact is None:
        raise HTTPException(status_code=404, detail=f"artifact not found: {aid}")
    return artifact


def artifact_records(aid: str) -> list[Record]:
    return RECORDS_BY_ARTIFACT.get(aid, [])


@app.get("/", tags=["meta"], summary="服务信息")
def root():
    return {
        "name": app.title,
        "version": app.version,
        "docs": "/docs",
        "health": "/health",
        "guide": "/guide",
        "llm_narration": bool(LLM_BASE_URL and LLM_API_KEY and LLM_MODEL),
    }


@app.get("/health", tags=["meta"], summary="健康检查")
def health():
    return {"status": "ok"}


@app.get(
    "/api/v1/artifacts",
    response_model=Envelope[list[ArtifactSummary]],
    tags=["artifacts"],
    summary="文物列表（含尺寸摘要）",
)
def list_artifacts():
    items = [
        ArtifactSummary(
            id=a.id,
            name=a.name,
            dynasty=a.dynasty,
            kiln=a.kiln,
            museum=a.museum,
            dimensions_cm=a.dimensions_cm,
            record_count=a.record_count,
        )
        for a in ARTIFACTS.values()
    ]
    return ok(items)


@app.get(
    "/api/v1/artifacts/{aid}",
    response_model=Envelope[ArtifactDetail],
    tags=["artifacts"],
    summary="文物完整档案（三大板块聚合）",
)
def artifact_detail(aid: str):
    a = get_artifact_or_404(aid)
    by_section: dict[str, list[Record]] = {}
    for r in artifact_records(aid):
        by_section.setdefault(r.section, []).append(r)
    return ok(ArtifactDetail(**a.model_dump(), records_by_section=by_section))


@app.get(
    "/api/v1/artifacts/{aid}/records",
    response_model=Envelope[RecordPage],
    tags=["records"],
    summary="记录检索（分页 / 板块 / 类别 / 关键词）",
)
def list_records(
    aid: str,
    section: Optional[str] = Query(None, description="按板块筛选"),
    category: Optional[str] = Query(None, description="按数据类别筛选"),
    q: Optional[str] = Query(None, description="关键词（匹配字段/值/类别/来源）"),
    page: int = Query(1, ge=1, description="页码，从 1 开始"),
    size: int = Query(20, ge=1, le=100, description="每页条数"),
):
    get_artifact_or_404(aid)
    items = artifact_records(aid)
    if section:
        items = [r for r in items if r.section == section]
    if category:
        items = [r for r in items if r.category == category]
    if q:
        kw = q.lower()
        items = [
            r
            for r in items
            if kw in r.category.lower()
            or kw in r.field.lower()
            or kw in r.value.lower()
            or kw in (r.source or "").lower()
        ]
    total = len(items)
    pages = max(1, math.ceil(total / size))
    start = (page - 1) * size
    page_obj = RecordPage(
        items=items[start:start + size], total=total, page=page, size=size, pages=pages
    )
    return ok(page_obj, meta={"total": total, "page": page, "size": size, "pages": pages})


@app.get(
    "/api/v1/artifacts/{aid}/records/{rid}",
    response_model=Envelope[Record],
    tags=["records"],
    summary="单条记录（含来源）",
)
def get_record(aid: str, rid: int):
    get_artifact_or_404(aid)
    for r in artifact_records(aid):
        if r.id == rid:
            return ok(r)
    raise HTTPException(status_code=404, detail=f"record not found: {rid}")


@app.get(
    "/api/v1/artifacts/{aid}/categories",
    response_model=Envelope[list[CategoryCount]],
    tags=["records"],
    summary="类别字典与计数",
)
def list_categories(aid: str):
    get_artifact_or_404(aid)
    counts: dict[tuple[str, str], int] = {}
    order: list[tuple[str, str]] = []
    for r in artifact_records(aid):
        key = (r.section, r.category)
        if key not in counts:
            order.append(key)
            counts[key] = 0
        counts[key] += 1
    items = [CategoryCount(section=s, category=c, count=counts[(s, c)]) for s, c in order]
    return ok(items)


@app.get(
    "/api/v1/search",
    response_model=Envelope[list[SearchHit]],
    tags=["search"],
    summary="跨板块全文搜索",
)
def search(
    q: str = Query(..., min_length=1, description="关键词"),
    section: Optional[str] = Query(None, description="限定板块"),
    limit: int = Query(20, ge=1, le=100, description="返回条数上限"),
):
    kw = q.lower()
    hits: list[SearchHit] = []
    for r in RECORDS:
        if section and r.section != section:
            continue
        haystack = f"{r.category} {r.field} {r.value} {r.source or ''}".lower()
        if kw in haystack:
            hits.append(SearchHit(**r.model_dump()))
            if len(hits) >= limit:
                break
    return ok(hits, meta={"q": q, "count": len(hits), "limit": limit})


@app.get("/api/v1/stats", response_model=Envelope[Stats], tags=["stats"], summary="数据集统计")
def stats():
    sections: dict[str, int] = {}
    categories: dict[str, int] = {}
    sources: dict[str, int] = {}
    for r in RECORDS:
        sections[r.section] = sections.get(r.section, 0) + 1
        categories[r.category] = categories.get(r.category, 0) + 1
        src = r.source or "未注明"
        sources[src] = sources.get(src, 0) + 1
    top = sorted(sources.items(), key=lambda kv: kv[1], reverse=True)[:10]
    return ok(
        Stats(
            artifact_count=len(ARTIFACTS),
            record_count=len(RECORDS),
            sections=[SectionCount(section=k, count=v) for k, v in sections.items()],
            categories=[CategoryTotal(category=k, count=v) for k, v in categories.items()],
            top_sources=[SourceCount(source=k, count=v) for k, v in top],
        )
    )


# ==================== 智能导览：讲解 / 语音 / 扫码部署 ====================

NARRATOR_SYSTEM_PROMPT = (
    "你是博物馆的金牌讲解员，正在展厅为游客面对面讲解一件文物。"
    "请只依据给出的知识库检索资料进行讲解，语言口语化、生动、有画面感，"
    "像讲故事一样自然衔接；不使用列表、序号或小标题，不编造资料以外的信息，"
    "全文控制在 350 字以内，结尾自然收束。"
)

_SECTION_OPENING = {
    "文物本体参数": "先从这件文物的形貌说起——",
    "工艺·纹饰·历史背景": "再看它的工艺与故事——",
    "对比·鉴定·价值": "最后从鉴定与价值的角度——",
}


def _llm_ready() -> bool:
    return bool(LLM_BASE_URL and LLM_API_KEY and LLM_MODEL)


def _template_narrate(question: str, hits: list[Record]) -> str:
    """未配置大模型（或调用失败）时的知识库讲解词模板。"""
    lines = [f"关于「{question}」，为您整理了 {len(hits)} 条知识。", ""]
    by_section: dict[str, list[Record]] = {}
    for r in hits:
        by_section.setdefault(r.section, []).append(r)
    for section, items in by_section.items():
        lines.append(f"【{section}】{_SECTION_OPENING.get(section, '')}")
        for i, r in enumerate(items, 1):
            lines.append(f"{i}. {r.value}（{r.category}·{r.field}）")
        lines.append("")
    lines.append("以上内容来源于公开资料整理的知识库。")
    return "\n".join(lines)


async def _llm_narrate(question: str, hits: list[Record]) -> str:
    """调用 OpenAI 兼容接口，把检索结果润色为口语化讲解词。"""
    context = "\n".join(
        f"- [{r.section}｜{r.category}] {r.field}：{r.value}" for r in hits
    )
    user = (
        "文物：耀州窑青釉刻花瓶（北宋 · 耀州窑 · 故宫博物院藏）\n"
        f"游客想了解：{question}\n\n"
        f"知识库检索结果：\n{context}"
    )
    payload = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": NARRATOR_SYSTEM_PROMPT},
            {"role": "user", "content": user},
        ],
        "temperature": 0.6,
        "max_tokens": 800,
    }
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            f"{LLM_BASE_URL}/chat/completions",
            json=payload,
            headers={"Authorization": f"Bearer {LLM_API_KEY}"},
        )
        resp.raise_for_status()
        data = resp.json()
    content = str(data["choices"][0]["message"]["content"]).strip()
    if not content:
        raise RuntimeError("LLM returned empty content")
    return content


@app.get(
    "/api/v1/narrate",
    response_model=Envelope[Narration],
    tags=["narration"],
    summary="AI 讲解词生成（配置 LLM 后为大模型润色，否则模板兜底）",
)
async def narrate(
    q: str = Query(..., min_length=1, description="游客问题 / 关键词"),
    limit: int = Query(20, ge=1, le=50, description="引用知识条数上限"),
):
    kw = q.lower()
    hits: list[Record] = []
    for r in RECORDS:
        haystack = f"{r.category} {r.field} {r.value} {r.source or ''}".lower()
        if kw in haystack:
            hits.append(r)
            if len(hits) >= limit:
                break
    if not hits:
        return ok(Narration(question=q, script="", mode="template", hit_count=0, sources=[]))
    sources: list[str] = []
    for r in hits:
        if r.source and r.source not in sources:
            sources.append(r.source)
    mode = "template"
    script = _template_narrate(q, hits)
    if _llm_ready():
        try:
            script = await _llm_narrate(q, hits)
            mode = "llm"
        except Exception:
            pass  # 大模型失败时静默回退模板，讲解不中断
    return ok(
        Narration(
            question=q,
            script=script,
            mode=mode,
            hit_count=len(hits),
            sources=sources[:5],
        )
    )


@app.post(
    "/api/v1/tts",
    tags=["narration"],
    summary="讲解词语音合成（微软神经语音，返回 mp3，免密钥）",
)
async def tts(req: TTSRequest):
    text = req.text.strip()
    if not text:
        raise HTTPException(status_code=400, detail="text is required")
    if req.voice not in TTS_VOICES:
        raise HTTPException(status_code=400, detail=f"voice must be one of {list(TTS_VOICES)}")
    key = hashlib.md5(f"{req.voice}\n{text}".encode("utf-8")).hexdigest()
    audio = _TTS_CACHE.get(key)
    if audio is None:
        try:
            communicate = edge_tts.Communicate(text, req.voice)
            buf = bytearray()
            async for chunk in communicate.stream():
                if chunk["type"] == "audio":
                    buf.extend(chunk["data"])
        except Exception as exc:
            raise HTTPException(status_code=502, detail=f"tts failed: {exc}")
        if not buf:
            raise HTTPException(status_code=502, detail="tts returned empty audio")
        audio = bytes(buf)
        if len(_TTS_CACHE) >= _TTS_CACHE_MAX:
            _TTS_CACHE.pop(next(iter(_TTS_CACHE)))
        _TTS_CACHE[key] = audio
    return Response(
        content=audio,
        media_type="audio/mpeg",
        headers={"Cache-Control": "public, max-age=86400"},
    )


def _lan_ip() -> str:
    """探测本机局域网 IP（不会真正发包）。"""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()


def _guide_url(request: Request, path: str) -> tuple[str, int]:
    port = request.url.port or (443 if request.url.scheme == "https" else 80)
    return f"http://{_lan_ip()}:{port}{path}", port


def _qr_png_bytes(url: str) -> bytes:
    img = qrcode.make(url)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


@app.get(
    "/api/v1/qr-info",
    response_model=Envelope[QRInfo],
    tags=["narration"],
    summary="局域网访问地址 + 二维码（展厅扫码用）",
)
def qr_info(request: Request, path: str = Query("/guide", description="目标路径")):
    url, port = _guide_url(request, path)
    png = _qr_png_bytes(url)
    return ok(
        QRInfo(
            url=url,
            lan_ip=_lan_ip(),
            port=port,
            qr_png_base64=base64.b64encode(png).decode("ascii"),
        )
    )


@app.get("/api/v1/qr", tags=["narration"], summary="导览页二维码 PNG")
def qr_png(request: Request, path: str = Query("/guide", description="目标路径")):
    url, _ = _guide_url(request, path)
    return Response(content=_qr_png_bytes(url), media_type="image/png")


@app.get("/guide", include_in_schema=False, summary="智能导览页面")
def guide_page():
    return FileResponse(GUIDE_PAGE_FILE, media_type="text/html")
