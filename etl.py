# -*- coding: utf-8 -*-
"""ETL：解析《耀州窑青釉刻花瓶_数据集.xlsx》→ data/dataset.json（统一记录模型 + 派生字段）。

仅使用标准库（zipfile + xml.etree），无需安装 openpyxl。
用法：py etl.py
"""
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SOURCE = BASE_DIR / "耀州窑青釉刻花瓶_数据集.xlsx"
OUT_FILE = BASE_DIR / "data" / "dataset.json"

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS = {"m": M[1:-1]}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

DIM_FIELDS = {"通高": "height", "口径": "mouth", "足径": "foot"}


def _col_index(ref: str) -> int:
    letters = re.match(r"([A-Z]+)", ref or "A").group(1)
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def load_workbook_rows(path: Path) -> list[tuple[str, list[list[str]]]]:
    """返回 [(sheet名, 按行序的二维字符串数组)]，按 workbook 中的顺序。"""
    with zipfile.ZipFile(path) as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter(M + "t")))

        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rid_target = {rel.get("Id"): rel.get("Target") for rel in rels}

        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = []
        for sh in wb.iter(M + "sheet"):
            target = rid_target.get(sh.get(RID), "").lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheets.append((sh.get("name"), target))

        result = []
        for name, target in sheets:
            ws = ET.fromstring(zf.read(target))
            rows: dict[int, list[str]] = {}
            for row in ws.iter(M + "row"):
                cells: dict[int, str] = {}
                for c in row.iter(M + "c"):
                    idx = _col_index(c.get("r", "A1"))
                    v = c.find(M + "v")
                    is_el = c.find(M + "is")
                    if c.get("t") == "s" and v is not None:
                        val = shared[int(v.text)]
                    elif c.get("t") == "inlineStr" and is_el is not None:
                        val = "".join(t.text or "" for t in is_el.iter(M + "t"))
                    elif v is not None:
                        val = v.text or ""
                    else:
                        val = ""
                    cells[idx] = val.strip()
                if cells:
                    width = max(cells) + 1
                    rows[int(row.get("r"))] = [cells.get(i, "") for i in range(width)]
            result.append((name, [rows[k] for k in sorted(rows)]))
    return result


def section_from_sheet_name(name: str) -> str:
    m = re.match(r"^表\d+·(.+)$", name)
    return m.group(1) if m else name


def parse_records(sheets) -> list[dict]:
    records = []
    for sheet_name, rows in sheets:
        section = section_from_sheet_name(sheet_name)
        header_idx = next(
            (i for i, row in enumerate(rows) if row and row[0] == "序号"), None
        )
        if header_idx is None:
            continue
        for row in rows[header_idx + 1:]:
            if not row or not any(row):
                continue
            no_raw, category, field, value, source = [v.strip() for v in (row + [""] * 5)[:5]]
            m = re.search(r"\d+", no_raw)
            records.append({
                "id": int(m.group()) if m else len(records) + 1,
                "artifact_id": "",
                "section": section,
                "category": category,
                "field": field,
                "value": value,
                "source": source,
            })
    return records


def _find(records, category: str, field: str) -> str:
    for r in records:
        if r["category"] == category and r["field"] == field:
            return r["value"]
    return ""


def derive_artifact(records) -> dict:
    dims = {}
    for r in records:
        if r["category"] == "尺寸参数" and r["field"] in DIM_FIELDS:
            m = re.search(r"\d+(?:\.\d+)?", r["value"])
            if m:
                dims[DIM_FIELDS[r["field"]]] = float(m.group())

    steps = []
    chain = _find(records, "工艺流程", "总工序数")
    if chain:
        steps = [s.strip() for s in chain.split("：", 1)[-1].split("→") if s.strip()]

    temp = None
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", _find(records, "工艺流程", "烧成温度"))]
    if nums:
        temp = {"min": min(nums), "max": max(nums)}

    digits = re.findall(r"\d+", _find(records, "基本信息", "文物编号"))
    return {
        "id": "artifact-" + (digits[0] if digits else "unknown"),
        "name": _find(records, "基本信息", "文物名称"),
        "inventory_number": _find(records, "基本信息", "文物编号"),
        "dynasty": _find(records, "基本信息", "年代"),
        "kiln": _find(records, "基本信息", "窑口"),
        "museum": _find(records, "基本信息", "收藏机构"),
        "dimensions_cm": dims or None,
        "process_steps": steps,
        "burn_temperature_c": temp,
        "record_count": len(records),
    }


def main():
    sheets = load_workbook_rows(SOURCE)
    records = parse_records(sheets)
    artifact = derive_artifact(records)
    for r in records:
        r["artifact_id"] = artifact["id"]

    OUT_FILE.parent.mkdir(exist_ok=True)
    OUT_FILE.write_text(
        json.dumps(
            {
                "generated_at": date.today().isoformat(),
                "source_file": SOURCE.name,
                "artifact": artifact,
                "records": records,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"OK: {len(records)} records -> {OUT_FILE.relative_to(BASE_DIR)}")


if __name__ == "__main__":
    main()
