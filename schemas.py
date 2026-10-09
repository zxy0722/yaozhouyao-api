# -*- coding: utf-8 -*-
"""Pydantic 数据模型与统一响应封装。"""
from typing import Any, Generic, Optional, TypeVar

from pydantic import BaseModel, Field

T = TypeVar("T")


class Dimensions(BaseModel):
    height: Optional[float] = Field(None, description="通高（厘米）")
    mouth: Optional[float] = Field(None, description="口径（厘米）")
    foot: Optional[float] = Field(None, description="足径（厘米）")


class BurnTemperature(BaseModel):
    min: Optional[float] = None
    max: Optional[float] = None


class Record(BaseModel):
    id: int
    artifact_id: str
    section: str = Field(description="板块（对应源表）")
    category: str = Field(description="数据类别")
    field: str = Field(description="字段")
    value: str = Field(description="值/描述")
    source: Optional[str] = Field(None, description="来源")


class Artifact(BaseModel):
    id: str
    name: str
    inventory_number: Optional[str] = None
    dynasty: Optional[str] = None
    kiln: Optional[str] = None
    museum: Optional[str] = None
    dimensions_cm: Optional[Dimensions] = None
    process_steps: list[str] = []
    burn_temperature_c: Optional[BurnTemperature] = None
    record_count: int = 0


class ArtifactSummary(BaseModel):
    id: str
    name: str
    dynasty: Optional[str] = None
    kiln: Optional[str] = None
    museum: Optional[str] = None
    dimensions_cm: Optional[Dimensions] = None
    record_count: int = 0


class ArtifactDetail(Artifact):
    records_by_section: dict[str, list[Record]] = {}


class RecordPage(BaseModel):
    items: list[Record]
    total: int
    page: int
    size: int
    pages: int


class CategoryCount(BaseModel):
    section: str
    category: str
    count: int


class CategoryTotal(BaseModel):
    category: str
    count: int


class SectionCount(BaseModel):
    section: str
    count: int


class SourceCount(BaseModel):
    source: str
    count: int


class SearchHit(BaseModel):
    id: int
    artifact_id: str
    section: str
    category: str
    field: str
    value: str
    source: Optional[str] = None


class Narration(BaseModel):
    question: str
    script: str = Field("", description="讲解词；无命中时为空字符串")
    mode: str = Field("template", description="llm=大模型润色 / template=知识库模板")
    hit_count: int = 0
    sources: list[str] = []


class TTSRequest(BaseModel):
    text: str = Field(min_length=1, max_length=2000, description="要合成语音的文本")
    voice: str = Field("zh-CN-XiaoxiaoNeural", description="音色标识（见 TTS_VOICES）")


class QRInfo(BaseModel):
    url: str = Field(description="局域网访问地址")
    lan_ip: str
    port: int
    qr_png_base64: str = Field(description="二维码 PNG 的 base64")


class Stats(BaseModel):
    artifact_count: int
    record_count: int
    sections: list[SectionCount]
    categories: list[CategoryTotal]
    top_sources: list[SourceCount]


class Envelope(BaseModel, Generic[T]):
    """统一响应封装。"""

    code: int = 0
    message: str = "ok"
    data: Optional[T] = None
    meta: Optional[dict[str, Any]] = None
