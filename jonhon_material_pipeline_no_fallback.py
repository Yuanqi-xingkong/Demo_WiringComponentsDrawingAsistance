#!/usr/bin/env python3
"""图纸 OCR 到物料推荐和接线表回填的可运行框架。

当前版本先跑通端到端流程：
1. 读取 PaddleOCR 服务生成的 ocr_result.json。
2. 归一化 OCR 文本和 HTML 表格。
3. 预留 LLM 抽取接口；未配置 API 时使用规则兜底抽取。
4. 从治理工作簿中生成候选物料并排序。
5. 基于模板创建新 xlsx 文件并回填关键表。

后续拿到 LLM API 后，只需要配置环境变量或替换 LLMExtractor.call_llm。
"""

from __future__ import annotations

import argparse
from copy import copy
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import asdict, dataclass, field
from difflib import SequenceMatcher
from html import escape, unescape
from io import StringIO
from pathlib import Path
from xml.etree import ElementTree as ET
from typing import Any

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OCR_JSON = BASE_DIR / "ocr_result.json"
DEFAULT_KNOWLEDGE_XLSX = BASE_DIR / "中航光电线束物料知识治理结果_final.xlsx"
DEFAULT_TEMPLATE_XLS = BASE_DIR / "线缆组件产品接线表格式模板20260420.xls"
DEFAULT_OUTPUT_DIR = BASE_DIR / "治理输出" / "auto_pipeline"
DEFAULT_ENV_FILE = BASE_DIR / ".env"
DASHSCOPE_COMPAT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DOCX_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


@dataclass
class WireRequirement:
    section_mm2: float | None = None
    shielded: bool | None = None
    color: str | None = None
    wire_type: str | None = None
    temperature: str | None = None
    withstand_voltage: str | None = None
    insulation_resistance: str | None = None


@dataclass
class ProtectionRequirement:
    name: str | None = None
    color: str | None = None
    color_code: str | None = None
    max_outer_diameter_mm: float | None = None


@dataclass
class ConnectionRow:
    from_connector: str = ""
    from_pin: str = ""
    from_signal: str = ""
    to_connector: str = ""
    to_pin: str = ""
    to_signal: str = ""
    wire_section_mm2: float | None = None
    wire_type: str = ""
    color: str = ""
    length: str = ""
    remark: str = ""


@dataclass
class DrawingMaterial:
    name: str = ""
    code: str = ""
    material: str = ""
    quantity: str = ""
    remark: str = ""
    category: str = ""


@dataclass
class DrawingRequirements:
    wire: WireRequirement = field(default_factory=WireRequirement)
    protection: ProtectionRequirement = field(default_factory=ProtectionRequirement)
    connector_models: list[str] = field(default_factory=list)
    tie_clip_models: list[str] = field(default_factory=list)
    connection_rows: list[ConnectionRow] = field(default_factory=list)
    materials_from_drawing: list[DrawingMaterial] = field(default_factory=list)
    raw_summary: dict[str, Any] = field(default_factory=dict)


def norm_text(value: Any) -> str:
    return str(value or "").strip()


def normalize_model(value: str) -> str:
    return re.sub(r"[\s\-_~/]", "", value.lower())


def parse_float(value: Any) -> float | None:
    if value is None:
        return None
    match = re.search(r"\d+(?:\.\d+)?", str(value))
    return float(match.group()) if match else None


def strip_html(value: str) -> str:
    text = re.sub(r"<br\s*/?>", "\n", value, flags=re.I)
    text = re.sub(r"</t[dh]>", "\t", text, flags=re.I)
    text = re.sub(r"</tr>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def load_ocr_blocks(ocr_json: Path) -> list[dict[str, Any]]:
    data = json.loads(ocr_json.read_text(encoding="utf-8"))
    result = data.get("result", data)
    if isinstance(result, dict) and isinstance(result.get("parsing_res_list"), list):
        return result["parsing_res_list"]
    if isinstance(data.get("all_results"), dict):
        for item in data["all_results"].values():
            if isinstance(item, dict) and isinstance(item.get("parsing_res_list"), list):
                return item["parsing_res_list"]
    raise ValueError(f"无法在 {ocr_json} 中找到 parsing_res_list")


def blocks_to_ocr_text(blocks: list[dict[str, Any]]) -> str:
    def sort_key(block: dict[str, Any]) -> tuple[float, float, float]:
        order = block.get("block_order")
        bbox = block.get("block_bbox") or [0, 0, 0, 0]
        if order is None:
            order = 10_000 + float(bbox[1])
        return float(order), float(bbox[1]), float(bbox[0])

    parts: list[str] = []
    for block in sorted(blocks, key=sort_key):
        label = block.get("block_label", "")
        content = norm_text(block.get("block_content"))
        if not content:
            continue
        bbox = block.get("block_bbox", "")
        if label == "table":
            parts.append(f"[table bbox={bbox}]\n{content}\n[table_text]\n{strip_html(content)}")
        else:
            parts.append(f"[{label} bbox={bbox}] {content}")
    return "\n\n".join(parts)


def extract_html_tables(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    try:
        import pandas as pd
    except Exception:
        return tables
    for block in blocks:
        if block.get("block_label") != "table":
            continue
        html = norm_text(block.get("block_content"))
        if "<table" not in html.lower():
            continue
        try:
            dfs = pd.read_html(StringIO(html))
        except Exception:
            continue
        for idx, df in enumerate(dfs, start=1):
            df = df.fillna("")
            rows = [[norm_text(v) for v in row] for row in df.values.tolist()]
            tables.append(
                {
                    "block_id": block.get("block_id"),
                    "bbox": block.get("block_bbox"),
                    "table_index": idx,
                    "columns": [norm_text(c) for c in df.columns],
                    "rows": rows,
                    "plain_text": strip_html(html),
                }
            )
    return tables


def load_docx_inputs(docx_path: Path) -> tuple[str, list[dict[str, Any]]]:
    """Read native Word text/tables and return pipeline-compatible inputs."""
    if not docx_path.exists():
        raise FileNotFoundError(f"docx 文件不存在: {docx_path}")
    try:
        with zipfile.ZipFile(docx_path) as archive:
            document_xml = archive.read("word/document.xml")
    except KeyError as exc:
        raise ValueError(f"{docx_path} 不是有效的 docx 文件，缺少 word/document.xml") from exc

    root = ET.fromstring(document_xml)
    body = root.find("w:body", DOCX_NS)
    if body is None:
        raise ValueError(f"{docx_path} 中没有可解析的 Word 正文")

    text_parts: list[str] = []
    tables: list[dict[str, Any]] = []
    paragraph_index = 0
    table_index = 0

    for child in body:
        if child.tag == qname("p"):
            text = docx_element_text(child)
            if text:
                paragraph_index += 1
                text_parts.append(f"[docx_text index={paragraph_index}] {text}")
        elif child.tag == qname("tbl"):
            rows = docx_table_rows(child)
            if not rows:
                continue
            table_index += 1
            plain_text = table_rows_to_plain_text(rows)
            text_parts.append(f"[docx_table index={table_index}]\n{plain_text}")
            tables.append(
                {
                    "block_id": f"docx_table_{table_index}",
                    "bbox": [],
                    "table_index": table_index,
                    "columns": rows[0],
                    "rows": rows,
                    "plain_text": plain_text,
                    "html": table_rows_to_html(rows),
                    "source": "docx",
                }
            )

    return "\n\n".join(text_parts), tables


def qname(local_name: str) -> str:
    return f"{{{DOCX_NS['w']}}}{local_name}"


def docx_element_text(element: ET.Element) -> str:
    chunks: list[str] = []
    for node in element.iter():
        if node.tag == qname("t") and node.text:
            chunks.append(node.text)
        elif node.tag == qname("tab"):
            chunks.append("\t")
        elif node.tag in {qname("br"), qname("cr")}:
            chunks.append("\n")
    text = "".join(chunks)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def docx_table_rows(table: ET.Element) -> list[list[str]]:
    rows: list[list[str]] = []
    max_cols = 0
    for tr in table.findall("w:tr", DOCX_NS):
        row = [docx_element_text(tc) for tc in tr.findall("w:tc", DOCX_NS)]
        if any(cell for cell in row):
            max_cols = max(max_cols, len(row))
            rows.append(row)
    if max_cols:
        rows = [row + [""] * (max_cols - len(row)) for row in rows]
    return rows


def table_rows_to_plain_text(rows: list[list[str]]) -> str:
    return "\n".join("\t".join(norm_text(cell) for cell in row) for row in rows)


def table_rows_to_html(rows: list[list[str]]) -> str:
    html_rows = []
    for row in rows:
        cells = "".join(f"<td>{escape(norm_text(cell))}</td>" for cell in row)
        html_rows.append(f"<tr>{cells}</tr>")
    return "<table>" + "".join(html_rows) + "</table>"


class LLMExtractor:
    """OpenAI-compatible LLM adapter with rule fallback.

    可配置环境变量：
    - JONHON_LLM_API_URL: 例如 http://host/v1/chat/completions
    - JONHON_LLM_API_KEY: Bearer token，可为空
    - JONHON_LLM_MODEL: 模型名
    - DASHSCOPE_API_KEY: DashScope API key
    - DASHSCOPE_BASE_URL: 默认 https://dashscope.aliyuncs.com/compatible-mode/v1
    - DASHSCOPE_MODEL: 默认 qwen-plus
    """

    def __init__(self, api_url: str | None = None, api_key: str | None = None, model: str | None = None):
        self.api_key = api_key or os.getenv("JONHON_LLM_API_KEY") or os.getenv("DASHSCOPE_API_KEY")
        raw_url = (
            api_url
            or os.getenv("JONHON_LLM_API_URL")
            or os.getenv("DASHSCOPE_API_URL")
            or os.getenv("DASHSCOPE_BASE_URL")
        )
        if not raw_url and self.api_key:
            raw_url = DASHSCOPE_COMPAT_BASE_URL
        self.api_url = normalize_chat_completions_url(raw_url) if raw_url else None
        self.model = model or os.getenv("JONHON_LLM_MODEL") or os.getenv("DASHSCOPE_MODEL", "qwen-plus")

    def extract(self, ocr_text: str, source_type: str = "ocr") -> dict[str, Any] | None:
        if not self.api_url:
            return None
        source_label = "DOCX原生解析内容" if source_type == "docx" else "OCR内容"
        prompt = (
            "你是线束/光缆组件图纸和技术要求文档的信息抽取助手。请只返回 JSON，不要解释。字段："
            "wire{section_mm2,shielded,color,wire_type,temperature,withstand_voltage,insulation_resistance},"
            "protection{name,color,color_code,max_outer_diameter_mm},"
            "connector_models[], tie_clip_models[], materials_from_drawing[], connection_rows[]。"
            "connection_rows 字段包括 from_connector,from_pin,from_signal,to_connector,to_pin,to_signal,"
            "wire_section_mm2,wire_type,color,length,remark。"
            "materials_from_drawing 字段包括 name,code,material,quantity,remark,category。\n\n"
            f"{source_label}：\n{ocr_text[:30000]}"
        )
        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": "只输出合法 JSON。"},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0,
        }
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        session = requests.Session()
        session.trust_env = False
        response = session.post(self.api_url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        obj = response.json()
        content = obj
        if isinstance(obj, dict) and obj.get("choices"):
            content = obj["choices"][0]["message"]["content"]
        if isinstance(content, str):
            match = re.search(r"\{.*\}", content, flags=re.S)
            return json.loads(match.group() if match else content)
        if isinstance(content, dict):
            return content
        raise ValueError("LLM 响应无法解析为 JSON")


def normalize_chat_completions_url(url: str) -> str:
    cleaned = url.rstrip("/")
    if cleaned.endswith("/chat/completions"):
        return cleaned
    if cleaned.endswith("/v1"):
        return f"{cleaned}/chat/completions"
    return cleaned


def rule_extract_requirements(ocr_text: str, tables: list[dict[str, Any]]) -> DrawingRequirements:
    plain = strip_html(ocr_text)
    req = DrawingRequirements()

    mm2_patterns = [
        r"(\d+(?:\.\d+)?)\s*mm\s*\^\s*\{\{?2\}?\}",
        r"(\d+(?:\.\d+)?)\s*mm²",
        r"(\d+(?:\.\d+)?)\s*平方",
    ]
    for pattern in mm2_patterns:
        match = re.search(pattern, plain, flags=re.I)
        if match:
            req.wire.section_mm2 = float(match.group(1))
            break
    if "屏蔽线缆" in plain or "屏蔽要求" in plain:
        req.wire.shielded = True
        req.wire.wire_type = "单芯屏蔽线缆" if "单芯屏蔽线缆" in plain else "屏蔽线缆"
    if "橙色" in plain:
        req.wire.color = "橙色"
    temp = re.search(r"-40.*?125\s*(?:℃|\\circ|C)", plain)
    if temp:
        req.wire.temperature = "-40~125℃"
    voltage = re.search(r"耐电压[:：]?\s*([0-9]+V\s*AC)", plain, flags=re.I)
    if voltage:
        req.wire.withstand_voltage = voltage.group(1).replace(" ", "")
    insulation = re.search(r"绝缘电阻.*?([≥>=]+\s*100M(?:Ω|\\Omega))", plain)
    if insulation:
        req.wire.insulation_resistance = insulation.group(1)

    if "波纹管" in plain:
        req.protection.name = "波纹管"
    if "RAL2003" in plain.upper():
        req.protection.color_code = "RAL2003"
        req.protection.color = "橙色"
    od = re.search(r"外径\s*[≤<=]\s*(\d+(?:\.\d+)?)\s*mm", plain, flags=re.I)
    if od:
        req.protection.max_outer_diameter_mm = float(od.group(1))

    connector_patterns = [
        r"\b\d{3}-\d{4}-\d{4}\b",
        r"\b\d-\d{6}-\d\b",
        r"\b\d{7}-\d\b",
    ]
    connectors: list[str] = []
    for pattern in connector_patterns:
        connectors.extend(re.findall(pattern, plain))
    connectors.extend(extract_contextual_connector_models(plain))
    req.connector_models = unique_keep_order(
        [
            m
            for m in connectors
            if is_plausible_connector_model(m) and not m.startswith("156-") and not m.startswith("5156-")
        ]
    )

    tie_models = re.findall(r"\bT\d{2,}[A-Z][A-Z0-9~.\-]+", plain)
    req.tie_clip_models = unique_keep_order([m.replace("~", "-") for m in tie_models])

    req.connection_rows = extract_connection_rows_from_tables(tables)
    req.materials_from_drawing = extract_materials_from_tables(tables)
    req.raw_summary = {
        "extractor": "rule_fallback",
        "ocr_text_chars": len(ocr_text),
        "table_count": len(tables),
    }
    return req


def extract_materials_from_tables(tables: list[dict[str, Any]]) -> list[DrawingMaterial]:
    materials: list[DrawingMaterial] = []
    seen: set[str] = set()
    for table in tables:
        rows = table.get("rows", [])
        header_idx = find_material_table_header(rows)
        if header_idx is None:
            continue
        header = [norm_text(v) for v in rows[header_idx]]
        name_cols = [idx for idx, value in enumerate(header) if value == "名称"]
        code_cols = [idx for idx, value in enumerate(header) if value == "代号"]
        material_col = first_index(header, "材料")
        quantity_col = first_index(header, "数量")
        remark_col = first_index(header, "备注")
        candidate_rows = rows[:header_idx] + rows[header_idx + 1 :]
        for row in candidate_rows:
            item = material_from_table_row(row, name_cols, code_cols, material_col, quantity_col, remark_col)
            if not item or not item.name:
                continue
            if item.name in {"名称", "图样名称Part Name", "3D 数模 3D Model"}:
                continue
            key = normalize_model("|".join([item.code, item.name, item.material, item.quantity]))
            if key and key not in seen:
                seen.add(key)
                materials.append(item)
    return materials


def find_material_table_header(rows: list[list[str]]) -> int | None:
    for idx, row in enumerate(rows):
        cells = [norm_text(v) for v in row]
        if "序号" in cells and "名称" in cells and "数量" in cells and ("材料" in cells or "代号" in cells):
            return idx
    return None


def first_index(values: list[str], target: str) -> int | None:
    try:
        return values.index(target)
    except ValueError:
        return None


def material_from_table_row(
    row: list[str],
    name_cols: list[int],
    code_cols: list[int],
    material_col: int | None,
    quantity_col: int | None,
    remark_col: int | None,
) -> DrawingMaterial | None:
    name = choose_best_cell(row, name_cols)
    if not name:
        return None
    code = choose_best_cell(row, code_cols)
    material = norm_text(row[material_col]) if material_col is not None and material_col < len(row) else ""
    quantity = norm_text(row[quantity_col]) if quantity_col is not None and quantity_col < len(row) else ""
    remark = norm_text(row[remark_col]) if remark_col is not None and remark_col < len(row) else ""
    if quantity and not re.fullmatch(r"\d+(?:\.\d+)?", quantity):
        return None
    if not quantity and not code and not material:
        return None
    return DrawingMaterial(
        name=name,
        code=code,
        material=material,
        quantity=quantity,
        remark=remark,
        category=classify_drawing_material(name, code),
    )


def choose_best_cell(row: list[str], indices: list[int]) -> str:
    values = [norm_text(row[idx]) for idx in indices if idx < len(row) and norm_text(row[idx])]
    if not values:
        return ""
    return max(values, key=len)


def classify_drawing_material(name: str, code: str = "") -> str:
    text = f"{name} {code}".upper()
    if "连接器" in name or re.fullmatch(r"\d{3}-\d{4}-\d{4}", code) or re.fullmatch(r"\d-\d{6}-\d", code):
        return "connector"
    if "线缆" in name or "导线" in name or "MM" in text and "屏蔽" in name:
        return "wire"
    if "波纹管" in name or "护套" in name:
        return "protection"
    if "扎带" in name or "卡扣" in name or re.search(r"\bT\d{2,}[A-Z0-9~.\-]+", text):
        return "tie_clip"
    return "unmatched_material"


def unique_keep_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        v = value.strip()
        key = normalize_model(v)
        if v and key not in seen:
            seen.add(key)
            result.append(v)
    return result


def extract_contextual_connector_models(text: str) -> list[str]:
    patterns = [
        r"([A-Z][A-Z0-9]*\d[A-Z0-9/.\-]{3,})(?:插头|连接器|插座)",
        r"(?:插头|连接器|插座)[:：]?\s*([A-Z][A-Z0-9]*\d[A-Z0-9/.\-]{3,})",
    ]
    models: list[str] = []
    for pattern in patterns:
        models.extend(re.findall(pattern, text, flags=re.I))
    return [m.rstrip(".-") for m in models]


def is_plausible_connector_model(model: str) -> bool:
    normalized = normalize_model(model)
    return len(normalized) >= 8 and bool(re.search(r"\d", normalized))


def extract_connection_rows_from_tables(tables: list[dict[str, Any]]) -> list[ConnectionRow]:
    rows: list[ConnectionRow] = []
    for table in tables:
        if not is_connection_table(table):
            continue
        last_connectors: list[str] = []
        for r in table.get("rows", []):
            joined = "\t".join(r)
            cells = [c for c in r if c]
            connectors = connector_models_from_cells(cells)
            if len(connectors) >= 2:
                last_connectors = connectors
            elif last_connectors:
                connectors = last_connectors
            if len(connectors) < 2:
                continue
            pins = [c for c in cells if re.fullmatch(r"[A-Za-z]?\d+[A-Za-z]?", c)]
            signals = [c for c in cells if is_signal_cell(c, connectors, pins)]
            section = None
            for c in cells:
                m = re.search(r"(\d+(?:\.\d+)?)\s*(?:mm\s*\^?\s*\{?2\}?|mm²|平方)", c, flags=re.I)
                if m:
                    section = float(m.group(1))
            rows.append(
                ConnectionRow(
                    from_connector=connectors[0],
                    from_pin=pins[0] if pins else "",
                    from_signal=signals[0] if signals else "",
                    to_connector=connectors[1],
                    to_pin=pins[-1] if len(pins) > 1 else "",
                    to_signal=signals[-1] if len(signals) > 1 else signals[0] if signals else "",
                    wire_section_mm2=section,
                    wire_type="单芯屏蔽线缆" if "单芯屏蔽" in joined else "屏蔽线缆",
                    color="橙色" if "橙" in joined else "",
                )
            )
    return rows


def is_connection_table(table: dict[str, Any]) -> bool:
    header_cells: list[str] = [norm_text(c) for c in table.get("columns", [])]
    for row in table.get("rows", [])[:5]:
        header_cells.extend(norm_text(c) for c in row)
    joined = "\t".join(header_cells)
    return (
        "接线" in joined
        or "对应关系" in joined
        or ("连接器" in joined and any(token in joined for token in ["孔位", "端子", "针脚", "点位"]))
    )


def connector_models_from_cells(cells: list[str]) -> list[str]:
    models: list[str] = []
    for cell in cells:
        models.extend(re.findall(r"\b\d{3}-\d{4}-\d{3,4}\b", cell))
        models.extend(re.findall(r"\b\d-\d{6}-\d\b", cell))
        models.extend(re.findall(r"\b\d{7}-\d\b", cell))
        models.extend(extract_contextual_connector_models(cell))
        models.extend(
            m.rstrip(".-")
            for m in re.findall(
                r"(?<![A-Z0-9/.\-])([A-Z][A-Z0-9]*\d[A-Z0-9/.\-]{3,})(?=$|[^A-Z0-9/.\-])",
                cell,
                flags=re.I,
            )
        )
    return unique_keep_order([model for model in models if is_plausible_connector_model(model)])


def is_signal_cell(value: str, connectors: list[str], pins: list[str]) -> bool:
    text = norm_text(value)
    if not text or text in connectors or text in pins:
        return False
    if any(connector and connector in text for connector in connectors):
        return False
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return False
    if re.search(r"(?:mm\s*\^?\s*\{?2\}?|mm²|平方|℃|V|Ω|ohm)", text, flags=re.I):
        return False
    return bool(re.search(r"[A-Za-z+\-]", text))


def requirements_from_llm_payload(payload: dict[str, Any]) -> DrawingRequirements:
    req = DrawingRequirements()
    wire = payload.get("wire") or {}
    protection = payload.get("protection") or {}
    req.wire = WireRequirement(
        section_mm2=parse_float(wire.get("section_mm2")),
        shielded=wire.get("shielded"),
        color=wire.get("color"),
        wire_type=wire.get("wire_type"),
        temperature=wire.get("temperature"),
        withstand_voltage=wire.get("withstand_voltage"),
        insulation_resistance=wire.get("insulation_resistance"),
    )
    req.protection = ProtectionRequirement(
        name=protection.get("name"),
        color=protection.get("color"),
        color_code=protection.get("color_code"),
        max_outer_diameter_mm=parse_float(protection.get("max_outer_diameter_mm")),
    )
    req.connector_models = unique_keep_order([norm_text(x) for x in payload.get("connector_models", [])])
    req.tie_clip_models = unique_keep_order([norm_text(x) for x in payload.get("tie_clip_models", [])])
    req.materials_from_drawing = [
        DrawingMaterial(
            name=norm_text(x.get("name")),
            code=norm_text(x.get("code")),
            material=norm_text(x.get("material")),
            quantity=norm_text(x.get("quantity")),
            remark=norm_text(x.get("remark")),
            category=norm_text(x.get("category")) or classify_drawing_material(norm_text(x.get("name")), norm_text(x.get("code"))),
        )
        for x in payload.get("materials_from_drawing", [])
        if isinstance(x, dict)
    ]
    req.connection_rows = [
        ConnectionRow(
            from_connector=norm_text(x.get("from_connector")),
            from_pin=norm_text(x.get("from_pin")),
            from_signal=norm_text(x.get("from_signal")),
            to_connector=norm_text(x.get("to_connector")),
            to_pin=norm_text(x.get("to_pin")),
            to_signal=norm_text(x.get("to_signal")),
            wire_section_mm2=parse_float(x.get("wire_section_mm2")),
            wire_type=norm_text(x.get("wire_type")),
            color=norm_text(x.get("color")),
            length=norm_text(x.get("length")),
            remark=norm_text(x.get("remark")),
        )
        for x in payload.get("connection_rows", [])
        if isinstance(x, dict)
    ]
    req.raw_summary = {"extractor": "llm", "llm_payload_keys": list(payload.keys())}
    return req


def build_requirements_from_text(
    input_text: str,
    tables: list[dict[str, Any]],
    *,
    use_llm_extract: bool = False,
    extractor: LLMExtractor | None = None,
    source_type: str = "ocr",
) -> DrawingRequirements:
    """Extract drawing requirements with optional LLM understanding.

    The rule path remains the deterministic fallback. When LLM is enabled, LLM
    results are used as the primary interpretation and rule extraction fills
    missing structured table/material fields.
    """
    rule_req = rule_extract_requirements(input_text, tables)
    if not use_llm_extract:
        return ensure_table_fallbacks(rule_req, tables)

    llm_payload = None
    llm_error = ""
    if extractor:
        try:
            llm_payload = extractor.extract(input_text, source_type=source_type)
        except Exception as exc:
            llm_error = str(exc)
    if not llm_payload:
        rule_req.raw_summary["llm_enabled"] = True
        rule_req.raw_summary["llm_used"] = False
        if llm_error:
            rule_req.raw_summary["llm_error"] = llm_error
        return ensure_table_fallbacks(rule_req, tables)

    llm_req = requirements_from_llm_payload(llm_payload)
    merged = merge_requirements(llm_req, rule_req)
    merged.raw_summary = {
        "extractor": "llm_with_rule_fallback",
        "source_type": source_type,
        "llm_payload_keys": list(llm_payload.keys()) if isinstance(llm_payload, dict) else [],
        "rule_fallback": rule_req.raw_summary,
    }
    return ensure_table_fallbacks(merged, tables)


def ensure_table_fallbacks(req: DrawingRequirements, tables: list[dict[str, Any]]) -> DrawingRequirements:
    if not req.connection_rows:
        req.connection_rows = extract_connection_rows_from_tables(tables)
    if not req.materials_from_drawing:
        req.materials_from_drawing = extract_materials_from_tables(tables)
    return req


def merge_requirements(primary: DrawingRequirements, fallback: DrawingRequirements) -> DrawingRequirements:
    merged = primary
    fill_wire_requirement(merged.wire, fallback.wire)
    fill_protection_requirement(merged.protection, fallback.protection)
    merged.connector_models = unique_keep_order(merged.connector_models + fallback.connector_models)
    merged.tie_clip_models = unique_keep_order(merged.tie_clip_models + fallback.tie_clip_models)
    merged.connection_rows = merge_connection_rows(merged.connection_rows, fallback.connection_rows)
    merged.materials_from_drawing = merge_drawing_materials(
        merged.materials_from_drawing, fallback.materials_from_drawing
    )
    return merged


def fill_wire_requirement(target: WireRequirement, fallback: WireRequirement) -> None:
    for field_name in [
        "section_mm2",
        "shielded",
        "color",
        "wire_type",
        "temperature",
        "withstand_voltage",
        "insulation_resistance",
    ]:
        if getattr(target, field_name) in (None, "") and getattr(fallback, field_name) not in (None, ""):
            setattr(target, field_name, getattr(fallback, field_name))


def fill_protection_requirement(target: ProtectionRequirement, fallback: ProtectionRequirement) -> None:
    for field_name in ["name", "color", "color_code", "max_outer_diameter_mm"]:
        if getattr(target, field_name) in (None, "") and getattr(fallback, field_name) not in (None, ""):
            setattr(target, field_name, getattr(fallback, field_name))


def merge_connection_rows(primary: list[ConnectionRow], fallback: list[ConnectionRow]) -> list[ConnectionRow]:
    result: list[ConnectionRow] = []
    seen: set[str] = set()
    for row in primary + fallback:
        key = normalize_model(
            "|".join([row.from_connector, row.from_pin, row.to_connector, row.to_pin, row.from_signal, row.to_signal])
        )
        if key and key not in seen:
            seen.add(key)
            result.append(row)
    return result


def merge_drawing_materials(primary: list[DrawingMaterial], fallback: list[DrawingMaterial]) -> list[DrawingMaterial]:
    result: list[DrawingMaterial] = []
    seen: set[str] = set()
    for item in primary + fallback:
        key = normalize_model("|".join([item.code, item.name, item.material, item.quantity, item.category]))
        if key and key not in seen:
            seen.add(key)
            result.append(item)
    return result


def load_sheet_rows(workbook_path: Path, sheet_name: str) -> list[dict[str, Any]]:
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    iterator = ws.iter_rows(values_only=True)
    headers = [norm_text(c) for c in next(iterator)]
    return [dict(zip(headers, row)) for row in iterator]


def load_wire_standard_rows(workbook_path: Path) -> list[dict[str, Any]]:
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    ws = wb["导线标准线径"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    normalized_headers = make_unique_headers(rows[0])
    parsed: list[dict[str, Any]] = []
    current_section = ""
    current_headers: list[str] | None = None

    for excel_row, raw_values in enumerate(rows[1:], start=2):
        values = list(raw_values)
        if not any(value not in (None, "") for value in values):
            continue
        first = norm_text(values[0] if values else "")
        if first.startswith("###"):
            current_section = first.replace("###", "").strip()
            current_headers = None
            continue
        if first == "来源图片":
            continue
        if is_wire_standard_embedded_header(values):
            current_headers = make_unique_headers(values)
            continue
        if current_section and current_headers:
            item = {
                "_sheet": "导线标准线径",
                "_excel_row": excel_row,
                "_record_type": "embedded_row",
                "表标题": current_section,
            }
            for header, value in zip(current_headers, values):
                if value not in (None, ""):
                    item[header] = value
            enrich_wire_standard_embedded_row(item)
            if item.get("导体标称截面"):
                parsed.append(item)
            continue
        item = {
            "_sheet": "导线标准线径",
            "_excel_row": excel_row,
            "_record_type": "normalized_row",
        }
        for header, value in zip(normalized_headers, values):
            if value not in (None, ""):
                item[header] = value
        if item.get("导体标称截面"):
            parsed.append(item)
    return parsed


def make_unique_headers(values: tuple[Any, ...] | list[Any]) -> list[str]:
    headers: list[str] = []
    seen: dict[str, int] = {}
    for idx, value in enumerate(values, start=1):
        header = norm_text(value) or f"未命名列{idx}"
        header = normalize_header_name(header)
        count = seen.get(header, 0) + 1
        seen[header] = count
        if count > 1:
            header = f"{header}_{count}"
        headers.append(header)
    return headers


def normalize_header_name(value: str) -> str:
    return (
        value.replace("\n", " / ")
        .replace("  ", " ")
        .replace("Amm", "mm")
        .replace("EO:Bmm", "mm")
        .replace("mm20", "mm")
        .replace("mm2 / 最大", "mm / 最大")
        .strip()
    )


def is_wire_standard_embedded_header(values: list[Any]) -> bool:
    nonempty = [norm_text(v) for v in values if v not in (None, "")]
    if len(nonempty) < 4:
        return False
    first = nonempty[0]
    joined = " ".join(nonempty)
    return (
        first.startswith("规格")
        or first.startswith("ISO导体尺寸")
        or ("导体" in joined and ("绝缘" in joined or "护套" in joined) and not looks_like_numeric(first))
    )


def looks_like_numeric(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?[a-zA-Z]*", value.strip()))


def enrich_wire_standard_embedded_row(item: dict[str, Any]) -> None:
    aliases = {
        "导体标称截面": [
            "规格 / 导体截面积mm / 标称",
            "ISO导体尺寸mm2",
            "ISO导体尺寸mm",
        ],
        "绝缘外径最小": [
            "绝缘 / 外径mm / 最小",
            "绝缘 R / 外径mm / 最小",
            "绝缘R / 外径mm / 最小",
            "单芯直径mm / Min.",
            "单芯直径mm/ Min.",
            "薄壁 / 电缆外径mm / Min.",
        ],
        "绝缘外径最大": [
            "绝缘 / 外径mm / 最大",
            "绝缘 R / 外径mm / 最大",
            "绝缘R / 外径mm / 最大",
            "单芯直径mm / Max.",
            "薄壁 / 电缆外径mm / Max.b",
        ],
        "护套外径最小": [
            "护套 / 外径mm / 最小",
            "护套 eB4 / 外径mm / 最小",
            "电缆外径mm / Min.",
            "厚壁 / 电缆外径mm / Min",
        ],
        "护套外径最大": [
            "护套 / 外径mm / 最大",
            "护套 eB4 / 外径mm / 最大",
            "电缆外径mm / Max.",
            "厚壁 / 电缆外径mm / Max.b",
        ],
    }
    for normalized_key, source_keys in aliases.items():
        for source_key in source_keys:
            if source_key in item and item[source_key] not in (None, ""):
                item[normalized_key] = item[source_key]
                break


def rank_wire_candidates(
    req: DrawingRequirements, rows: list[dict[str, Any]], standard: dict[str, Any] | None = None
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    if not standard:
        return out
    target_spec = f"1*{format_number(req.wire.section_mm2)}" if req.wire.section_mm2 else ""
    level_score = {"一类": 300, "二类": 200, "三类": 100}
    standard_name = get_wire_standard_name(standard) if standard else ""
    for row in rows:
        score = 0
        reasons: list[str] = []
        spec = norm_text(row.get("物料规格"))
        shield = norm_text(row.get("屏蔽"))
        color = norm_text(row.get("颜色"))
        level = norm_text(row.get("选用等级_治理"))
        if level not in level_score:
            continue
        if target_spec and spec == target_spec:
            score += 100
            reasons.append(f"规格精确匹配 {target_spec}")
        else:
            continue
        if standard:
            score += 30
            reasons.append(f"导线标准线径匹配 {standard_name}")
        if req.wire.shielded is True:
            if shield and shield not in {"非", "/"}:
                score += 30
                reasons.append(f"屏蔽满足 {shield}")
            else:
                score -= 80
                reasons.append("图纸要求屏蔽，候选为非屏蔽")
        if req.wire.color and req.wire.color in color:
            score += 15
            reasons.append(f"颜色匹配 {color}")
        score += level_score.get(level, 0)
        if level:
            reasons.append(f"选用等级 {level}")
        candidate = dict(row)
        candidate["_score"] = score
        candidate["_reasons"] = reasons
        out.append(candidate)
    return sorted(out, key=lambda x: x["_score"], reverse=True)


def augment_wire_candidates_with_standard_sizes(
    req: DrawingRequirements, candidates: list[dict[str, Any]], standard_rows: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    standard = select_wire_standard_size(req, standard_rows)
    if not standard:
        return candidates
    for candidate in candidates:
        std_insulation = standard_range(standard, "绝缘外径", "绝缘外径最小", "绝缘外径最大")
        std_sheath = standard_range(standard, "护套外径", "护套外径最小", "护套外径最大")
        if std_insulation:
            candidate["标准绝缘外径"] = std_insulation
        if std_sheath:
            candidate["标准护套外径"] = std_sheath
        if not norm_text(candidate.get("绝缘外径")) and std_insulation:
            candidate["绝缘外径"] = std_insulation
            candidate.setdefault("_reasons", []).append("绝缘外径由导线标准线径补充")
        if not norm_text(candidate.get("护套外径")) and std_sheath:
            candidate["护套外径"] = std_sheath
            candidate.setdefault("_reasons", []).append("护套外径由导线标准线径补充")
        candidate["标准线径依据"] = get_wire_standard_name(standard)
    return candidates


def standard_range(row: dict[str, Any], direct_key: str, min_key: str, max_key: str) -> str:
    direct = norm_text(row.get(direct_key))
    if direct:
        return direct
    min_value = norm_text(row.get(min_key))
    max_value = norm_text(row.get(max_key))
    if min_value and max_value:
        return f"{min_value}-{max_value}"
    return min_value or max_value


def select_wire_standard_size(req: DrawingRequirements, standard_rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    if req.wire.section_mm2 is None:
        return None
    section = float(req.wire.section_mm2)
    matches: list[dict[str, Any]] = []
    for row in standard_rows:
        row_section = parse_float(row.get("导体标称截面"))
        if row_section is None or abs(row_section - section) > 1e-6:
            continue
        standard_name = get_wire_standard_name(row)
        if not is_valid_wire_standard_name(standard_name):
            continue
        if req.wire.shielded is True and ("屏蔽" not in standard_name or "非屏蔽" in standard_name):
            continue
        if req.wire.shielded is False and "非屏蔽" not in standard_name:
            continue
        matches.append(row)
    if matches:
        return matches[0]
    for row in standard_rows:
        row_section = parse_float(row.get("导体标称截面"))
        if (
            row_section is not None
            and abs(row_section - section) <= 1e-6
            and is_valid_wire_standard_name(get_wire_standard_name(row))
        ):
            return row
    return None


def get_wire_standard_name(row: dict[str, Any] | None) -> str:
    if not row:
        return ""
    return norm_text(row.get("标准") or row.get("表标题"))


def is_valid_wire_standard_name(value: str) -> bool:
    return bool(value and ("导线" in value or "电缆" in value or "LV" in value.upper() or "ISO" in value.upper() or "QC/T" in value.upper()))


def rank_protection_candidates(
    req: DrawingRequirements, rows: list[dict[str, Any]], selected_wire: dict[str, Any] | None
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    if not has_protection_requirement(req):
        return out
    required_id = required_protection_inner_diameter(req, selected_wire)
    max_od = req.protection.max_outer_diameter_mm
    level_score = {"优选": 40, "可选": 25, "专用": 10, "限用": -30, "禁用": -100}
    for row in rows:
        score = 0
        reasons: list[str] = []
        name = norm_text(row.get("名称"))
        if req.protection.name and req.protection.name not in name:
            continue
        color = norm_text(row.get("颜色"))
        material = norm_text(row.get("材料"))
        inner_d = parse_float(row.get("内径"))
        outer_d = parse_float(row.get("规格外径")) or parse_outer_from_spec(row.get("规格"))
        if req.protection.color and "橙" in req.protection.color:
            if "橙" in color:
                score += 25
                reasons.append(f"颜色匹配 {color}")
            else:
                continue
        if max_od and outer_d and outer_d > max_od:
            continue
        if max_od and outer_d:
            score += 20
            reasons.append(f"外径 {outer_d:g}mm <= {max_od:g}mm")
        if required_id and inner_d:
            if inner_d > required_id:
                score += 25
                reasons.append(f"内径 {inner_d:g}mm > 线束包络外径约 {required_id:g}mm")
            else:
                continue
        if "PP" in material:
            score += 10
            reasons.append(f"材料优先 {material}")
        level = norm_text(row.get("选用等级"))
        score += level_score.get(level, 0)
        if level:
            reasons.append(f"选用等级 {level}")
        candidate = dict(row)
        candidate["_score"] = score
        candidate["_reasons"] = reasons
        out.append(candidate)
    return sorted(out, key=lambda x: x["_score"], reverse=True)


def has_protection_requirement(req: DrawingRequirements) -> bool:
    protection = req.protection
    if any(
        [
            protection.name,
            protection.color,
            protection.color_code,
            protection.max_outer_diameter_mm is not None,
        ]
    ):
        return True
    return any(item.category == "protection" for item in req.materials_from_drawing)


def required_protection_inner_diameter(req: DrawingRequirements, selected_wire: dict[str, Any] | None) -> float | None:
    if not selected_wire:
        return None
    wire_od = diameter_upper_bound(selected_wire.get("护套外径") or selected_wire.get("标准护套外径"))
    if wire_od is None:
        return None
    quantity = drawing_wire_quantity_from_requirements(req) or 1
    return wire_od * quantity


def drawing_wire_quantity_from_requirements(req: DrawingRequirements) -> int:
    for item in req.materials_from_drawing:
        if item.category == "wire":
            value = parse_float(item.quantity)
            if value:
                return int(value)
    if req.connection_rows:
        return max(1, len(req.connection_rows))
    return 1


def diameter_upper_bound(value: Any) -> float | None:
    text = norm_text(value)
    if not text or text == "/":
        return None
    numbers = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", text)]
    if not numbers:
        return None
    if "±" in text and len(numbers) >= 2:
        return numbers[0] + numbers[1]
    return max(numbers)


def rank_connector_candidates(req: DrawingRequirements, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    models = {normalize_model(m): m for m in req.connector_models}
    for row in rows:
        model = norm_text(row.get("供应商连接器型号"))
        if normalize_model(model) in models:
            candidate = dict(row)
            candidate["_score"] = 100
            candidate["_reasons"] = [f"供应商型号精确匹配 {models[normalize_model(model)]}"]
            out.append(candidate)
    return out


def expand_connector_bom_items(req: DrawingRequirements, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    seen_codes: set[str] = set()
    requested = {normalize_model(m): m for m in req.connector_models}
    for row in rows:
        model = norm_text(row.get("供应商连接器型号"))
        model_key = normalize_model(model)
        if model_key not in requested:
            continue
        main = dict(row)
        main["_bom_role"] = "主连接器"
        main["_parent_model"] = requested[model_key]
        add_unique_material(items, seen_codes, main)
        seq = norm_text(row.get("序号"))
        if seq:
            child_prefix = f"{seq}."
            for child in rows:
                child_seq = norm_text(child.get("序号"))
                if child_seq.startswith(child_prefix):
                    child_item = dict(child)
                    child_item["_bom_role"] = "连接器子件"
                    child_item["_parent_model"] = requested[model_key]
                    add_unique_material(items, seen_codes, child_item)
    return items


def add_unique_material(items: list[dict[str, Any]], seen_codes: set[str], item: dict[str, Any]) -> None:
    code = norm_text(item.get("中航物料号")) or norm_text(item.get("物料编号")) or norm_text(item.get("供应商连接器型号"))
    key = normalize_model(code)
    if key and key not in seen_codes:
        seen_codes.add(key)
        items.append(item)


def rank_tie_candidates(req: DrawingRequirements, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    seen_materials: set[str] = set()
    for model in req.tie_clip_models:
        model_norm = normalize_model(model)
        if not is_specific_match_query(model):
            continue
        best_hits: list[dict[str, Any]] = []
        for row in rows:
            hay = " ".join(norm_text(row.get(k)) for k in ["物料编号", "物料名称", "物料牌号/零件号", "物料规格", "备注"])
            part_norm = normalize_model(norm_text(row.get("物料牌号/零件号")))
            hay_norm = normalize_model(hay)
            fuzzy = SequenceMatcher(None, model_norm, part_norm).ratio() if model_norm and part_norm else 0
            if model_norm and (model_norm in hay_norm or fuzzy >= 0.82):
                candidate = dict(row)
                candidate["_score"] = 100 if model_norm in hay_norm else int(fuzzy * 100)
                candidate["_query_model"] = model
                if model_norm in hay_norm:
                    candidate["_reasons"] = [f"型号匹配 {model}"]
                else:
                    candidate["_reasons"] = [f"OCR型号 {model} 与知识库型号 {row.get('物料牌号/零件号')} 相似度 {fuzzy:.2f}"]
                best_hits.append(candidate)
        if best_hits:
            best = max(best_hits, key=lambda item: item["_score"])
            material_no = normalize_model(norm_text(best.get("物料编号")))
            if material_no and material_no not in seen_materials:
                seen_materials.add(material_no)
                out.append(best)
    return out


def enrich_requirements_with_governance(req: DrawingRequirements, tie_rows: list[dict[str, Any]]) -> None:
    tie_queries = list(req.tie_clip_models)
    for model in req.connector_models:
        if find_tie_match(model, tie_rows):
            tie_queries.append(model)
    req.connector_models = [model for model in req.connector_models if not find_tie_match(model, tie_rows)]
    for item in req.materials_from_drawing:
        if item.category in {"wire", "protection"}:
            continue
        queries = [item.code, item.name, item.remark]
        match_query = next((query for query in queries if find_tie_match(query, tie_rows)), "")
        if match_query:
            item.category = "tie_clip"
            tie_queries.append(match_query)
    req.tie_clip_models = unique_keep_order(tie_queries)


def find_tie_match(query: Any, rows: list[dict[str, Any]]) -> dict[str, Any] | None:
    query_text = norm_text(query)
    if not is_specific_match_query(query_text):
        return None
    query_norm = normalize_model(query_text)
    if not query_norm:
        return None
    for row in rows:
        hay = " ".join(norm_text(row.get(k)) for k in ["物料编号", "物料名称", "物料牌号/零件号", "物料规格", "备注"])
        part = norm_text(row.get("物料牌号/零件号"))
        hay_norm = normalize_model(hay)
        part_norm = normalize_model(part)
        fuzzy = SequenceMatcher(None, query_norm, part_norm).ratio() if part_norm else 0
        if query_norm in hay_norm or fuzzy >= 0.82:
            return row
    return None


def is_specific_match_query(query: Any) -> bool:
    query_norm = normalize_model(norm_text(query))
    return len(query_norm) >= 5 and not query_norm.isdigit()


def parse_outer_from_spec(value: Any) -> float | None:
    numbers = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", norm_text(value))]
    if len(numbers) >= 2:
        return max(numbers)
    return numbers[0] if numbers else None


def format_number(value: float | None) -> str:
    if value is None:
        return ""
    return str(int(value)) if float(value).is_integer() else str(value)


def build_recommendations(req: DrawingRequirements, knowledge_xlsx: Path) -> dict[str, Any]:
    wire_rows = load_sheet_rows(knowledge_xlsx, "导线电缆知识")
    wire_standard_rows = load_wire_standard_rows(knowledge_xlsx)
    protection_rows = load_sheet_rows(knowledge_xlsx, "防护辅材知识")
    tie_rows = load_sheet_rows(knowledge_xlsx, "扎带卡扣知识")
    connector_rows = load_sheet_rows(knowledge_xlsx, "连接器BOM映射")
    enrich_requirements_with_governance(req, tie_rows)

    selected_standard = select_wire_standard_size(req, wire_standard_rows)
    wire_candidates = rank_wire_candidates(req, wire_rows, selected_standard)
    wire_candidates = augment_wire_candidates_with_standard_sizes(req, wire_candidates, wire_standard_rows)
    selected_wire = wire_candidates[0] if wire_candidates else None
    protection_candidates = rank_protection_candidates(req, protection_rows, selected_wire)
    connector_candidates = rank_connector_candidates(req, connector_rows)
    connector_bom_items = expand_connector_bom_items(req, connector_rows)
    tie_candidates = rank_tie_candidates(req, tie_rows)

    return {
        "requirements": asdict(req),
        "selected": {
            "wire": trim_candidate(selected_wire),
            "protection": trim_candidate(protection_candidates[0] if protection_candidates else None),
            "connectors": [trim_candidate(x) for x in connector_candidates],
            "connector_bom_items": [trim_candidate(x) for x in connector_bom_items],
            "tie_clips": [trim_candidate(x) for x in tie_candidates],
            "drawing_materials": [asdict(x) for x in req.materials_from_drawing],
        },
        "candidates": {
            "wire": [trim_candidate(x) for x in wire_candidates[:10]],
            "protection": [trim_candidate(x) for x in protection_candidates[:10]],
            "connectors": [trim_candidate(x) for x in connector_candidates],
            "connector_bom_items": [trim_candidate(x) for x in connector_bom_items],
            "tie_clips": [trim_candidate(x) for x in tie_candidates],
            "drawing_materials": [asdict(x) for x in req.materials_from_drawing],
        },
    }


def trim_candidate(candidate: dict[str, Any] | None) -> dict[str, Any] | None:
    if not candidate:
        return None
    keep = {
        "知识类型",
        "物料编号",
        "物料规格",
        "选用等级_治理",
        "选用等级",
        "屏蔽",
        "颜色",
        "序号",
        "名称",
        "材料",
        "规格",
        "内径",
        "规格外径",
        "护套外径",
        "绝缘外径",
        "绝缘护套材料",
        "标准绝缘外径",
        "标准护套外径",
        "标准线径依据",
        "供应商连接器型号",
        "中航物料号",
        "物料名称",
        "物料牌号/零件号",
        "厂家",
        "_score",
        "_reasons",
        "_query_model",
        "_bom_role",
        "_parent_model",
    }
    return {k: normalize_json_value(v) for k, v in candidate.items() if k in keep and v not in (None, "")}


def normalize_json_value(value: Any) -> Any:
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, list):
        return [normalize_json_value(v) for v in value]
    if isinstance(value, dict):
        return {str(k): normalize_json_value(v) for k, v in value.items()}
    return str(value)


def make_output_workbook(template_xls: Path, output_xlsx: Path, req: DrawingRequirements, rec: dict[str, Any]) -> None:
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    template_xlsx = convert_template_to_xlsx(template_xls)
    shutil.copyfile(template_xlsx, output_xlsx)
    wb = load_workbook(output_xlsx)

    clear_unused_template_sheets(wb, used_sheet_names={"接线表", "连接器及其所用辅材"})
    fill_connection_sheet(wb, req, rec)
    fill_connector_sheet(wb, rec)
    wb.save(output_xlsx)


def convert_template_to_xlsx(template_xls: Path) -> Path:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice:
        raise RuntimeError("未找到 libreoffice/soffice，无法把 .xls 模板转换为 .xlsx。")
    if not template_xls.exists():
        raise FileNotFoundError(f"模板文件不存在: {template_xls}")
    tmpdir = Path(tempfile.mkdtemp(prefix="jonhon_template_"))
    lo_home = tmpdir / "lo-home"
    lo_run = tmpdir / "lo-run"
    lo_home.mkdir(parents=True, exist_ok=True)
    lo_run.mkdir(parents=True, exist_ok=True)
    os.chmod(lo_home, 0o700)
    os.chmod(lo_run, 0o700)
    env = os.environ.copy()
    env.update(
        {
            "HOME": str(lo_home),
            "XDG_CONFIG_HOME": str(lo_home / ".config"),
            "XDG_CACHE_HOME": str(lo_home / ".cache"),
            "XDG_RUNTIME_DIR": str(lo_run),
        }
    )
    result = subprocess.run(
        [libreoffice, "--headless", "--convert-to", "xlsx", "--outdir", str(tmpdir), str(template_xls)],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=60,
        env=env,
    )
    if result.returncode != 0:
        raise RuntimeError(
            "模板 .xls 转 .xlsx 失败，已停止生成，避免输出丢失模板格式的空白文件。\n"
            f"命令: {libreoffice} --headless --convert-to xlsx --outdir {tmpdir} {template_xls}\n"
            f"返回码: {result.returncode}\nstdout: {result.stdout}\nstderr: {result.stderr}"
        )
    files = list(tmpdir.glob("*.xlsx"))
    if not files:
        raise RuntimeError(f"模板转换命令成功返回，但未在 {tmpdir} 生成 .xlsx 文件。")
    return files[0]


def set_cell(ws, cell: str, value: Any) -> None:
    ws[cell] = value
    ws[cell].alignment = Alignment(vertical="center", wrap_text=True)


def write_template_cell(ws, row: int, col: int, value: Any, template_row: int) -> None:
    apply_template_cell_style(ws, row, col, template_row)
    ws.cell(row, col).value = value


def apply_template_cell_style(ws, row: int, col: int, template_row: int) -> None:
    source = ws.cell(template_row, col)
    target = ws.cell(row, col)
    if source.has_style:
        target._style = copy(source._style)
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.protection = copy(source.protection)
    target.number_format = source.number_format
    target.font = Font(
        name=source.font.name or "宋体",
        sz=source.font.sz or 10,
        bold=source.font.bold,
        italic=source.font.italic,
        vertAlign=source.font.vertAlign,
        underline=source.font.underline,
        strike=source.font.strike,
        color="FF000000",
    )
    target.alignment = Alignment(
        horizontal=source.alignment.horizontal or "center",
        vertical="center",
        text_rotation=source.alignment.text_rotation,
        wrap_text=True,
        shrink_to_fit=source.alignment.shrink_to_fit,
        indent=source.alignment.indent,
    )


def clear_unused_template_sheets(wb, used_sheet_names: set[str]) -> None:
    for ws in wb.worksheets:
        if ws.title in used_sheet_names:
            continue
        for row in ws.iter_rows():
            for cell in row:
                if cell.__class__.__name__ != "MergedCell":
                    cell.value = None


def fill_connection_sheet(wb, req: DrawingRequirements, rec: dict[str, Any]) -> None:
    ws = wb["接线表"] if "接线表" in wb.sheetnames else wb.create_sheet("接线表", 0)
    template_row = 9
    last_data_row = ensure_template_row_capacity(
        ws, start_row=template_row, template_end_row=41, required_count=len(req.connection_rows), min_col=1, max_col=12
    )
    unmerge_region(ws, min_row=9, max_row=last_data_row, min_col=1, max_col=12)
    clear_region(ws, min_row=9, max_row=last_data_row, min_col=1, max_col=12)
    selected_wire = rec["selected"].get("wire") or {}
    wire_label = build_wire_spec_label(selected_wire)
    connector_map = build_connector_material_map(rec)
    if ws.max_row < 8:
        ws.append(["序号", "来自件号", "来自孔位", "来自线号", "接至件号", "接至孔位", "接至线号", "导线规格", "颜色", "长度", "备注"])
    set_cell(ws, "A6", "自动生成接线表")
    start = 9
    for idx, row in enumerate(req.connection_rows, start=1):
        r = start + idx - 1
        write_template_cell(ws, r, 1, idx, template_row)
        write_template_cell(ws, r, 2, connector_map.get(normalize_model(row.from_connector), ""), template_row)
        write_template_cell(ws, r, 3, row.from_pin, template_row)
        write_template_cell(ws, r, 4, row.from_signal, template_row)
        write_template_cell(ws, r, 5, connector_map.get(normalize_model(row.to_connector), ""), template_row)
        write_template_cell(ws, r, 6, row.to_pin, template_row)
        write_template_cell(ws, r, 7, row.to_signal, template_row)
        write_template_cell(ws, r, 8, wire_label, template_row)
        write_template_cell(ws, r, 9, row.color or req.wire.color, template_row)
        write_template_cell(ws, r, 10, row.length, template_row)
        write_template_cell(ws, r, 11, row.remark, template_row)


def build_wire_spec_label(selected_wire: dict[str, Any]) -> str:
    if not selected_wire:
        return ""
    parts = [
        norm_text(selected_wire.get("物料规格")),
        norm_text(selected_wire.get("绝缘护套材料")),
        norm_text(selected_wire.get("护套外径") or selected_wire.get("标准护套外径")),
    ]
    return " / ".join([p for p in parts if p])


def build_connector_material_map(rec: dict[str, Any]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for item in rec["selected"].get("connectors") or []:
        supplier_model = norm_text(item.get("供应商连接器型号"))
        material_no = norm_text(item.get("中航物料号"))
        if supplier_model and material_no:
            mapping[normalize_model(supplier_model)] = material_no
    return mapping


def fill_connector_sheet(wb, rec: dict[str, Any]) -> None:
    ws = wb["连接器及其所用辅材"] if "连接器及其所用辅材" in wb.sheetnames else wb.create_sheet("连接器及其所用辅材")
    set_cell(ws, "A1", "自动生成连接器及其所用辅材统计表")
    materials = build_connector_sheet_materials(rec)
    template_row = 4
    last_data_row = ensure_template_row_capacity(
        ws, start_row=template_row, template_end_row=13, required_count=len(materials), min_col=1, max_col=8
    )
    unmerge_region(ws, min_row=4, max_row=last_data_row, min_col=1, max_col=8)
    clear_region(ws, min_row=4, max_row=last_data_row, min_col=1, max_col=8)
    for idx, item in enumerate(materials, start=1):
        row = 3 + idx
        write_template_cell(ws, row, 1, idx, template_row)
        if item["kind"] == "connector_bom":
            write_template_cell(ws, row, 2, item.get("source_model"), template_row)
            write_template_cell(ws, row, 3, item.get("material_no"), template_row)
            write_template_cell(ws, row, 4, item.get("material_no"), template_row)
            write_template_cell(ws, row, 5, 1, template_row)
            write_template_cell(ws, row, 6, item.get("name"), template_row)
        else:
            write_template_cell(ws, row, 6, item.get("name"), template_row)
            write_template_cell(ws, row, 7, item.get("material_no"), template_row)
            write_template_cell(ws, row, 8, item.get("quantity") or 1, template_row)


def build_connector_sheet_materials(rec: dict[str, Any]) -> list[dict[str, str]]:
    materials: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in rec["selected"].get("connector_bom_items") or rec["selected"].get("connectors") or []:
        material_no = norm_text(item.get("中航物料号"))
        if not material_no:
            continue
        add_sheet_material(
            materials,
            seen,
            {
                "kind": "connector_bom",
                "material_no": material_no,
                "name": norm_text(item.get("名称")),
                "source_model": norm_text(item.get("_parent_model") or item.get("供应商连接器型号")),
            },
        )
    wire = rec["selected"].get("wire")
    if wire:
        material_no = norm_text(wire.get("物料编号"))
        if material_no:
            add_sheet_material(
                materials,
                seen,
                {
                    "kind": "auxiliary",
                    "material_no": material_no,
                    "name": build_wire_spec_label(wire) or norm_text(wire.get("物料规格")),
                    "source_model": norm_text(wire.get("物料规格")),
                    "quantity": drawing_material_quantity(rec, "wire") or "1",
                },
            )
    for item in rec["selected"].get("tie_clips") or []:
        material_no = norm_text(item.get("物料编号"))
        if not material_no:
            continue
        add_sheet_material(
            materials,
            seen,
            {
                "kind": "auxiliary",
                "material_no": material_no,
                "name": norm_text(item.get("物料名称") or item.get("物料牌号/零件号")),
                "source_model": norm_text(item.get("_query_model")),
            },
        )
    protection = rec["selected"].get("protection")
    if protection:
        material_no = norm_text(protection.get("物料编号"))
        if material_no:
            add_sheet_material(
                materials,
                seen,
                {
                    "kind": "auxiliary",
                    "material_no": material_no,
                    "name": norm_text(protection.get("名称") or protection.get("规格")),
                    "source_model": norm_text(protection.get("规格")),
                },
            )
    for item in rec["selected"].get("drawing_materials") or []:
        if not should_write_unmatched_drawing_material(item):
            continue
        name = norm_text(item.get("name"))
        if not name:
            continue
        add_sheet_material(
            materials,
            seen,
            {
                "kind": "drawing_unmatched",
                "material_no": "",
                "name": name,
                "source_model": norm_text(item.get("code")),
                "quantity": norm_text(item.get("quantity")) or "1",
                "remark": norm_text(item.get("remark")),
            },
        )
    return materials


def should_write_unmatched_drawing_material(item: dict[str, Any]) -> bool:
    category = norm_text(item.get("category"))
    if category in {"connector", "wire", "protection", "tie_clip"}:
        return False
    return bool(norm_text(item.get("name")))


def drawing_material_quantity(rec: dict[str, Any], category: str) -> str:
    for item in rec["selected"].get("drawing_materials") or []:
        if item.get("category") == category and norm_text(item.get("quantity")):
            return norm_text(item.get("quantity"))
    return ""


def add_sheet_material(materials: list[dict[str, str]], seen: set[str], item: dict[str, str]) -> None:
    key = normalize_model(item.get("material_no") or item.get("name", ""))
    if key and key not in seen:
        seen.add(key)
        materials.append(item)


def fill_protection_sheet(wb, rec: dict[str, Any]) -> None:
    ws = wb["护套类辅材表"] if "护套类辅材表" in wb.sheetnames else wb.create_sheet("护套类辅材表")
    set_cell(ws, "A1", "自动生成护套类辅材统计表")
    item = rec["selected"].get("protection")
    last_data_row = ensure_template_row_capacity(
        ws, start_row=4, template_end_row=13, required_count=1 if item else 0, min_col=1, max_col=12
    )
    unmerge_region(ws, min_row=4, max_row=last_data_row, min_col=1, max_col=12)
    clear_region(ws, min_row=4, max_row=last_data_row, min_col=1, max_col=12)
    if item:
        ws.cell(4, 1).value = 1
        ws.cell(4, 4).value = item.get("名称") or "波纹管"
        ws.cell(4, 5).value = item.get("物料编号")
        ws.cell(4, 6).value = item.get("规格")


def fill_recommendation_sheet(wb, rec: dict[str, Any]) -> None:
    if "推荐依据" in wb.sheetnames:
        del wb["推荐依据"]
    ws = wb.create_sheet("推荐依据")
    ws.append(["类别", "排序", "物料号/型号", "名称/规格", "分数", "推荐理由"])
    rows: list[tuple[str, list[dict[str, Any] | None]]] = [
        ("导线电缆", rec["candidates"].get("wire", [])),
        ("防护辅材", rec["candidates"].get("protection", [])),
        ("连接器", rec["candidates"].get("connectors", [])),
        ("扎带卡扣", rec["candidates"].get("tie_clips", [])),
    ]
    for category, candidates in rows:
        for idx, candidate in enumerate([c for c in candidates if c], start=1):
            code = candidate.get("物料编号") or candidate.get("中航物料号") or candidate.get("供应商连接器型号")
            name = candidate.get("名称") or candidate.get("物料名称") or candidate.get("物料规格") or candidate.get("规格")
            ws.append(
                [
                    category,
                    idx,
                    code,
                    name,
                    candidate.get("_score"),
                    "；".join(candidate.get("_reasons", [])),
                ]
            )
    style_header(ws)
    autosize(ws)


def unmerge_region(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    ranges = list(ws.merged_cells.ranges)
    for merged in ranges:
        if (
            merged.max_row >= min_row
            and merged.min_row <= max_row
            and merged.max_col >= min_col
            and merged.min_col <= max_col
        ):
            safe_unmerge_range(ws, merged)


def safe_unmerge_range(ws, merged) -> None:
    try:
        ws.unmerge_cells(str(merged))
        return
    except KeyError:
        pass
    try:
        ws.merged_cells.ranges.remove(merged)
    except (KeyError, ValueError):
        pass
    for row in range(merged.min_row, merged.max_row + 1):
        for col in range(merged.min_col, merged.max_col + 1):
            cell = ws._cells.get((row, col))
            if cell is not None and cell.__class__.__name__ == "MergedCell":
                del ws._cells[(row, col)]
            ws.cell(row, col)


def ensure_template_row_capacity(
    ws,
    start_row: int,
    template_end_row: int,
    required_count: int,
    min_col: int,
    max_col: int,
) -> int:
    """Extend a template data region by copying the final template row style."""
    available_count = max(0, template_end_row - start_row + 1)
    required_count = max(0, required_count)
    extra_count = max(0, required_count - available_count)
    if extra_count:
        insert_at = template_end_row + 1
        ws.insert_rows(insert_at, amount=extra_count)
        unmerge_region(ws, min_row=insert_at, max_row=insert_at + extra_count - 1, min_col=min_col, max_col=max_col)
        for row in range(insert_at, insert_at + extra_count):
            copy_row_format(ws, source_row=template_end_row, target_row=row, min_col=min_col, max_col=max_col)
    return template_end_row + extra_count


def copy_row_format(ws, source_row: int, target_row: int, min_col: int, max_col: int) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(min_col, max_col + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.protection = copy(source.protection)


def clear_region(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row, col)
            if cell.__class__.__name__ != "MergedCell":
                cell.value = None


def style_header(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def autosize(ws) -> None:
    for col in range(1, ws.max_column + 1):
        width = 10
        for row in range(1, min(ws.max_row, 200) + 1):
            value = ws.cell(row, col).value
            if value is not None:
                width = max(width, min(60, len(str(value)) * 1.2 + 2))
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def write_json(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def env_truthy(value: str | None) -> bool:
    return norm_text(value).lower() in {"1", "true", "yes", "on", "y"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ocr-json", type=Path, default=DEFAULT_OCR_JSON)
    parser.add_argument("--docx", type=Path, default=None, help="直接解析 Word 正文和表格，不经过 OCR。")
    parser.add_argument("--knowledge-xlsx", type=Path, default=DEFAULT_KNOWLEDGE_XLSX)
    parser.add_argument("--template-xls", type=Path, default=DEFAULT_TEMPLATE_XLS)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--env-file", type=Path, default=DEFAULT_ENV_FILE)
    parser.add_argument("--llm-api-url", default=None)
    parser.add_argument("--llm-api-key", default=None)
    parser.add_argument("--llm-model", default=None)
    parser.add_argument("--use-llm-extract", action="store_true", help="启用 LLM 辅助理解抽取，规则抽取作为兜底。")
    parser.add_argument("--no-llm-extract", action="store_true", help="即使配置了 LLM，也强制只使用规则抽取。")
    args = parser.parse_args()

    load_env_file(args.env_file)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    if args.docx:
        ocr_text, tables = load_docx_inputs(args.docx)
        write_json(args.output_dir / "docx_tables.json", tables)
        source_type = "docx"
    else:
        blocks = load_ocr_blocks(args.ocr_json)
        ocr_text = blocks_to_ocr_text(blocks)
        tables = extract_html_tables(blocks)
        source_type = "ocr"
    (args.output_dir / "ocr_normalized.txt").write_text(ocr_text, encoding="utf-8")
    write_json(args.output_dir / "ocr_tables.json", tables)

    use_llm_extract = (
        args.use_llm_extract
        or bool(args.llm_api_url or args.llm_api_key or args.llm_model)
        or env_truthy(os.getenv("JONHON_USE_LLM_EXTRACT"))
    )
    if args.no_llm_extract:
        use_llm_extract = False
    extractor = LLMExtractor(args.llm_api_url, args.llm_api_key, args.llm_model) if use_llm_extract else None
    req = build_requirements_from_text(
        ocr_text,
        tables,
        use_llm_extract=use_llm_extract,
        extractor=extractor,
        source_type=source_type,
    )

    rec = build_recommendations(req, args.knowledge_xlsx)
    write_json(args.output_dir / "requirements.json", asdict(req))
    write_json(args.output_dir / "recommendations.json", rec)

    output_xlsx = args.output_dir / "线缆组件产品接线表_自动生成_nofallback.xlsx"
    make_output_workbook(args.template_xls, output_xlsx, req, rec)

    print("OCR归一化文本:", args.output_dir / "ocr_normalized.txt")
    print("抽取结果:", args.output_dir / "requirements.json")
    print("推荐结果:", args.output_dir / "recommendations.json")
    print("回填文件:", output_xlsx)
    selected = rec.get("selected", {})
    print("推荐导线:", selected.get("wire"))
    print("推荐防护辅材:", selected.get("protection"))
    print("推荐连接器数:", len(selected.get("connectors") or []))
    print("推荐扎带/卡扣数:", len(selected.get("tie_clips") or []))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
