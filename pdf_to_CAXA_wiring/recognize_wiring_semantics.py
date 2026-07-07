#!/usr/bin/env python3
"""调用阿里百炼视觉模型，将 PDF 页面图片识别为接线语义 JSON。
conda run -n quankepytorch python `
  .\pdf_to_caxa_wiring\recognize_wiring_semantics.py `
  ".\pdf_to_caxa_wiring\output\pages\manifest.json" `
  -o ".\pdf_to_caxa_wiring\output\semantic.json"
  """

from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import time
from pathlib import Path
from typing import Any

from openai import OpenAI


DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_MODEL = "qwen3.7-plus"
DEFAULT_WIRE_MATERIAL_NO = "__DEFAULT_WIRE__"

SYSTEM_PROMPT = """你是线缆组件工程图识别器。你的任务是读取工程图页面，提取用于生成
CAXA 关系示意图的事实。只识别图中明确可见或表格明确给出的内容，不推测 CAD 坐标、
块尺寸或插入点。无法确定的字段使用 null，并写入 validation.unresolved_items。
必须只输出一个合法 JSON 对象，不要输出 Markdown。"""

USER_PROMPT = r"""
请结合所有页面识别线缆接线关系，并严格输出以下 JSON 结构：
{
  "schema_version": "1.0",
  "source": {"pdf_file": "文件名", "page_numbers": [1]},
  "components": [
    {
      "instance_id": "J1",
      "material_no": "图纸或明细表中的物料号，不能确认则为 null",
      "label_raw": "原图标注",
      "type": "connector",
      "relative_position": "left|right|upper|lower|center|unknown",
      "ports": [{"port_id": "1", "label_raw": "原图标注或 null"}],
      "confidence": 0.0
    }
  ],
  "wires": [
    {
      "instance_id": "W1",
      "material_no": "物料号或 null",
      "label_raw": "原图标注或 null",
      "type": "wire",
      "from": {"component": "J1", "port": "1"},
      "to": {"component": "J2", "port": "1"},
      "direction": "left_to_right|right_to_left|top_to_bottom|bottom_to_top|unknown",
      "length_mm": null,
      "accessories": [
        {
          "instance_id": "A1",
          "material_no": "物料号或 null",
          "label_raw": "原图标注或 null",
          "type": "cable_tie|terminal|sleeve|label|clip|other|unknown",
          "order": 1,
          "position_ratio": 0.25,
          "side": "on_wire|above|below|left|right|unknown",
          "confidence": 0.0
        }
      ],
      "confidence": 0.0
    }
  ],
  "unassigned_items": [],
  "validation": {
    "manual_review_required": false,
    "warnings": [],
    "unresolved_items": []
  }
}

识别规则：
1. 图中每一个序号（类似1，2，3...）都是需要识别的对象，也就是"label_raw"的1到8都要识别到，序号“3”到“6”都是辅材（accessories）。
2. components 目前只放连接器；每个实际连接器建立一个实例，编号 J1、J2……。
3. 每条实际导线建立一个 wires 实例，编号 W1、W2……。from/to 必须引用已存在的连接器。
4. 以导线 from 到 to 为方向，对辅材沿线排序，order 从 1 连续递增。
5. position_ratio 是辅材沿 from→to 路径的归一化相对位置：起点为 0，终点为 1。
   它只描述相对位置，不代表真实长度；无法可靠判断时为 null，但仍尽量给出 order。
6. 同一物料出现多次时必须建立多个 instance_id，不得合并。
7. 优先读取图中的接线关系表、明细栏、序号引出线及局部视图，并相互交叉校验。
8. 物料号、端口号和长度必须来自图纸可见文字(重点关注图右下表格“代号”那一栏，物料号要从代号栏获取，如果代号为空，再从右边的名称栏获取)，不得根据外观编造。
9. 视觉交叉但没有明确连接符号的线，不得认定为连接。
10. 如果明细表序号能够映射到图中引出序号，应据此确定构件类型和物料号。
11. confidence 范围为 0 到 1。任何关键字段低于 0.8 时，manual_review_required 必须为 true。
12. 不属于连接器、导线或导线上辅材的明细项目放入 unassigned_items，不要强行挂接。
13. 输出必须是标准 JSON，禁止解释文字、注释和代码围栏。
"""

PROMPT_ADJUSTMENTS = r"""

项目补充规则（必须覆盖上面的通用规则）：
1. PDF 中导线通常没有单独物料号。只要能确认某对象是实际导线，就必须输出到 wires；如果图中没有导线物料号，wire.material_no 固定写成 "__DEFAULT_WIRE__"，不要因为缺少导线物料号而放入 unresolved_items。
2. accessories 只保留真实辅材：必须是导线上/导线附近的实物辅材，并且能从图纸或明细表中明确读到 material_no。
3. 如果一个疑似辅材没有 material_no，大概率是标识文字、序号、标签框、说明线或模型误识别；不要放入 accessories。必要时可放入 unassigned_items，但不要影响 manual_review_required。
4. 不要为没有物料号的疑似辅材编造 material_no。
"""

COVERING_AND_DIMENSION_PROMPT = r"""

新增识别要求：
1. 必须识别导线外部的线段状包覆件，例如波纹管、护套管、保护套管。不要把波纹管当作普通 accessories。
2. 对每条 wire 可新增 coverings 数组。coverings 中对象格式：
   {
     "instance_id": "C1",
     "material_no": "图纸/明细表明确给出的物料号；没有则为 null",
     "label_raw": "图中序号或原始标注，例如 2",
     "type": "corrugated_tube|protective_tube|sleeve_cover|other|unknown",
     "order": 1,
     "start_exposed_length_mm": 45,
     "end_exposed_length_mm": 45,
     "position_ratio_start": 0.1,
     "position_ratio_end": 0.9,
     "confidence": 0.0
   }
3. 波纹管是沿导线方向覆盖一段长度的覆盖层；如果图中显示两端导线露出，必须提取两端露出长度。
4. 必须识别图中的真实尺寸标注数字，并按导线路径从左到右/从 from 到 to 输出到 wire.dimensions 数组。dimensions 中对象格式：
   {
     "instance_id": "D1",
     "value_mm": 45,
     "label_raw": "45",
     "order": 1,
     "from_ref": "wire_start|covering_start|accessory:A1|covering_end|unknown",
     "to_ref": "covering_start|covering_end|accessory:A1|wire_end|unknown",
     "confidence": 0.0
   }
5. 尺寸标注的 value_mm 必须来自图纸可见数字，不得根据 CAD 示意图坐标计算或编造。示例中类似 45、52、15、126、105、100、45 这样的尺寸链要按顺序输出。
6. 如果同一连接器之间有两根导线且尺寸/波纹管共同作用在这组导线上，可只在第一条 wire 上完整输出 coverings 和 dimensions，第二条 wire 可留空数组。
7. dimensions 的 from_ref/to_ref 必须引用实际对象锚点，禁止只输出数字。可用锚点名包括：
   - "wire_start"：导线 from 端
   - "wire_end"：导线 to 端
   - "covering_start"：第一段波纹管起点
   - "covering_end"：第一段波纹管终点
   - "accessory:A1"、"accessory:A2"：对应辅材实例中心
8. 尺寸线必须表达“两个锚点之间”的真实距离。例如左端露出长度 45 应输出 wire_start -> covering_start；波纹管端到辅材距离应输出 covering_start -> accessory:A1；右端露出长度 45 应输出 covering_end -> wire_end。
"""


class SemanticRecognitionError(RuntimeError):
    pass


def image_to_data_url(path: Path) -> str:
    if not path.is_file():
        raise FileNotFoundError(f"页面图片不存在：{path}")
    suffix_to_mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}
    mime = suffix_to_mime.get(path.suffix.lower())
    if mime is None:
        raise ValueError(f"不支持的页面图片格式：{path.suffix}")
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def load_manifest(path: Path) -> tuple[dict[str, Any], list[tuple[int, Path]]]:
    if not path.is_file():
        raise FileNotFoundError(f"页面清单不存在：{path}")
    manifest = json.loads(path.read_text(encoding="utf-8"))
    page_entries = manifest.get("pages")
    if not isinstance(page_entries, list) or not page_entries:
        raise ValueError("manifest.json 中 pages 必须是非空数组。")

    pages: list[tuple[int, Path]] = []
    for index, entry in enumerate(page_entries):
        if not isinstance(entry, dict):
            raise ValueError(f"pages[{index}] 不是对象。")
        page_number = entry.get("page_number")
        image_file = entry.get("image_file")
        if not isinstance(page_number, int) or not isinstance(image_file, str):
            raise ValueError(f"pages[{index}] 缺少有效 page_number 或 image_file。")
        image_path = path.parent / image_file
        if not image_path.is_file():
            raise FileNotFoundError(f"第 {page_number} 页图片不存在：{image_path}")
        pages.append((page_number, image_path))
    return manifest, pages


def validate_semantic(data: dict[str, Any]) -> None:
    required = {"schema_version", "source", "components", "wires", "validation"}
    missing = required - data.keys()
    if missing:
        raise SemanticRecognitionError(f"模型 JSON 缺少字段：{sorted(missing)}")
    if not isinstance(data["components"], list) or not isinstance(data["wires"], list):
        raise SemanticRecognitionError("components 和 wires 必须是数组。")

    component_ids = [item.get("instance_id") for item in data["components"] if isinstance(item, dict)]
    if len(component_ids) != len(data["components"]) or any(not item for item in component_ids):
        raise SemanticRecognitionError("存在缺少 instance_id 的连接器。")
    if len(component_ids) != len(set(component_ids)):
        raise SemanticRecognitionError("连接器 instance_id 重复。")

    valid_components = set(component_ids)
    wire_ids: set[str] = set()
    accessory_ids: set[str] = set()
    for wire in data["wires"]:
        if not isinstance(wire, dict) or not wire.get("instance_id"):
            raise SemanticRecognitionError("存在无效导线或导线缺少 instance_id。")
        wire_id = wire["instance_id"]
        if wire_id in wire_ids:
            raise SemanticRecognitionError(f"导线 instance_id 重复：{wire_id}")
        wire_ids.add(wire_id)
        for endpoint in ("from", "to"):
            value = wire.get(endpoint)
            if not isinstance(value, dict) or value.get("component") not in valid_components:
                raise SemanticRecognitionError(f"导线 {wire_id} 的 {endpoint} 未引用有效连接器。")

        accessories = wire.get("accessories", [])
        if not isinstance(accessories, list):
            raise SemanticRecognitionError(f"导线 {wire_id} 的 accessories 不是数组。")
        orders: list[int] = []
        for item in accessories:
            if not isinstance(item, dict) or not item.get("instance_id"):
                raise SemanticRecognitionError(f"导线 {wire_id} 存在无效辅材。")
            accessory_id = item["instance_id"]
            if accessory_id in accessory_ids:
                raise SemanticRecognitionError(f"辅材 instance_id 重复：{accessory_id}")
            accessory_ids.add(accessory_id)
            order = item.get("order")
            if not isinstance(order, int) or order < 1:
                raise SemanticRecognitionError(f"辅材 {accessory_id} 的 order 无效。")
            orders.append(order)
            ratio = item.get("position_ratio")
            if ratio is not None and (not isinstance(ratio, (int, float)) or not 0 <= ratio <= 1):
                raise SemanticRecognitionError(f"辅材 {accessory_id} 的 position_ratio 超出 0 到 1。")
        if len(orders) != len(set(orders)):
            raise SemanticRecognitionError(f"导线 {wire_id} 的辅材 order 重复。")


def normalize_semantic(data: dict[str, Any]) -> None:
    """项目约定清洗：导线给默认图块物料号，忽略无物料号的疑似辅材。"""
    validation = data.setdefault("validation", {})
    if not isinstance(validation, dict):
        validation = {"manual_review_required": True, "warnings": [], "unresolved_items": []}
        data["validation"] = validation

    warnings = validation.setdefault("warnings", [])
    if not isinstance(warnings, list):
        warnings = []
        validation["warnings"] = warnings

    for wire in data.get("wires", []):
        if not isinstance(wire, dict):
            continue

        if not wire.get("material_no"):
            wire["material_no"] = DEFAULT_WIRE_MATERIAL_NO

        accessories = wire.get("accessories", [])
        if not isinstance(accessories, list):
            continue

        kept_accessories: list[dict[str, Any]] = []
        skipped_ids: list[str] = []
        for accessory in accessories:
            if not isinstance(accessory, dict):
                continue
            if accessory.get("material_no"):
                kept_accessories.append(accessory)
            else:
                skipped_ids.append(str(accessory.get("instance_id", "<unknown>")))

        if skipped_ids:
            warnings.append(
                f"已忽略导线 {wire.get('instance_id')} 上缺少 material_no 的疑似辅材："
                + ", ".join(skipped_ids)
            )
        wire["accessories"] = kept_accessories


def recognize(
    manifest_path: Path,
    output_path: Path,
    model: str,
    base_url: str,
    retries: int,
) -> dict[str, Any]:
    # api_key = os.getenv("DASHSCOPE_API_KEY")
    api_key = "sk-ee011544ce574c93acd4c3042bd15550"
    if not api_key:
        raise SemanticRecognitionError("未设置环境变量 DASHSCOPE_API_KEY。")

    manifest, pages = load_manifest(manifest_path)
    source_pdf = manifest.get("source_pdf", {}).get("file_name", "unknown.pdf")
    page_numbers = [number for number, _ in pages]
    content: list[dict[str, Any]] = [
        {"type": "text", "text": f"源 PDF：{source_pdf}；输入页码：{page_numbers}。\n\n{USER_PROMPT}\n\n{PROMPT_ADJUSTMENTS}\n\n{COVERING_AND_DIMENSION_PROMPT}"}
    ]
    for page_number, image_path in pages:
        content.append({"type": "text", "text": f"下面是 PDF 第 {page_number} 页："})
        content.append(
            {"type": "image_url", "image_url": {"url": image_to_data_url(image_path)}}
        )

    client = OpenAI(api_key=api_key, base_url=base_url, timeout=300, max_retries=0)
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": content},
                ],
                response_format={"type": "json_object"},
                temperature=0,
                max_completion_tokens=12000,
                extra_body={"enable_thinking": False, "vl_high_resolution_images": True},
            )
            raw_content = response.choices[0].message.content
            if not raw_content:
                raise SemanticRecognitionError("模型返回内容为空。")
            result = json.loads(raw_content)
            if not isinstance(result, dict):
                raise SemanticRecognitionError("模型输出 JSON 顶层不是对象。")
            result["source"] = {"pdf_file": source_pdf, "page_numbers": page_numbers}
            normalize_semantic(result)
            validate_semantic(result)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            return result
        except Exception as exc:
            last_error = exc
            if attempt < retries:
                seconds = 2 ** (attempt - 1)
                print(f"[警告] 第 {attempt} 次识别失败：{exc}；{seconds} 秒后重试。", file=sys.stderr)
                time.sleep(seconds)
    raise SemanticRecognitionError(f"百炼识别失败：{last_error}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="使用阿里百炼识别接线图语义。")
    parser.add_argument("manifest", type=Path, help="PDF 提取阶段生成的 manifest.json")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/semantic.json"))
    parser.add_argument("--model", default=os.getenv("DASHSCOPE_VL_MODEL", DEFAULT_MODEL))
    parser.add_argument("--base-url", default=os.getenv("DASHSCOPE_BASE_URL", DEFAULT_BASE_URL))
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument(
        "--check-only", action="store_true", help="只检查清单和页面，不调用百炼 API"
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        if args.check_only:
            manifest, pages = load_manifest(args.manifest)
            print(f"清单检查通过：{manifest.get('source_pdf', {}).get('file_name', 'unknown')}")
            print(f"待识别页面：{[number for number, _ in pages]}")
            return 0
        result = recognize(
            args.manifest, args.output, args.model, args.base_url, max(1, args.retries)
        )
        print(f"识别完成：{args.output.resolve()}")
        print(f"连接器：{len(result['components'])}，导线：{len(result['wires'])}")
        print(f"需要人工复核：{result['validation'].get('manual_review_required', True)}")
        return 0
    except Exception as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
