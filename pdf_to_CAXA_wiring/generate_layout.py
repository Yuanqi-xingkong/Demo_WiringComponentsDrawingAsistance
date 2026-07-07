#!/usr/bin/env python3
"""将 semantic.json 确定性转换为 CAXA 块布局 layout.json。

conda run -n quankepytorch python `
  .\pdf_to_caxa_wiring\generate_layout.py `
  .\pdf_to_caxa_wiring\output\semantic.json `
  -o .\pdf_to_caxa_wiring\output\layout.json `
  --catalog .\pdf_to_caxa_wiring\block_catalog.example.json
  
  """

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any


DEFAULT_RULES: dict[str, Any] = {
    "units": "mm",
    "left_connector_x": 0.0,
    "right_connector_x": 300.0,
    "connector_y": 0.0,
    "connector_width": 40.0,
    "port_pitch": 16.0,
    "wire_standard_length": 100.0,
    "covering_standard_length": 100.0,
    "covering_start_ratio": 0.16,
    "covering_end_ratio": 0.84,
    "dimension_offset_y": -28.0,
    "dimension_tick_height": 6.0,
    "dimension_text_height": 5.0,
    "default_dimension_chain_mm": [],
    "accessory_offset": 0.0,
    "above_offset": 12.0,
    "below_offset": -12.0,
}

DEFAULT_WIRE_MATERIAL_NO = "__DEFAULT_WIRE__"
DEFAULT_WIRE_BLOCK_METADATA: dict[str, Any] = {
    "block_name": "WIRE",
    "standard_length": DEFAULT_RULES["wire_standard_length"],
}
DEFAULT_COVERING_MATERIAL_NO = "__DEFAULT_CORRUGATED_TUBE__"
DEFAULT_COVERING_BLOCK_METADATA: dict[str, Any] = {
    "block_name": "CORRUGATED_TUBE",
    "standard_length": DEFAULT_RULES["covering_standard_length"],
}


class LayoutError(RuntimeError):
    pass


def load_json(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"文件不存在：{path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"JSON 顶层必须是对象：{path}")
    return data


def merge_rules(config_path: Path | None) -> tuple[dict[str, Any], dict[str, Any]]:
    rules = dict(DEFAULT_RULES)
    catalog: dict[str, Any] = {}
    if config_path:
        config = load_json(config_path)
        custom_rules = config.get("layout_rules", {})
        if not isinstance(custom_rules, dict):
            raise ValueError("block_catalog.json 的 layout_rules 必须是对象。")
        rules.update(custom_rules)
        raw_catalog = config.get("blocks", {})
        if not isinstance(raw_catalog, dict):
            raise ValueError("block_catalog.json 的 blocks 必须是对象。")
        catalog = raw_catalog
    catalog.setdefault(DEFAULT_WIRE_MATERIAL_NO, DEFAULT_WIRE_BLOCK_METADATA)
    catalog.setdefault(DEFAULT_COVERING_MATERIAL_NO, DEFAULT_COVERING_BLOCK_METADATA)
    return rules, catalog


def port_y(port: str | None, ports: list[dict[str, Any]], center_y: float, pitch: float) -> float:
    ids = [str(item.get("port_id")) for item in ports if item.get("port_id") is not None]
    if port is None or str(port) not in ids:
        return center_y
    index = ids.index(str(port))
    return center_y + ((len(ids) - 1) / 2.0 - index) * pitch


def block_parameters(
    material_no: str | None,
    instance_id: str,
    item_type: str,
    catalog: dict[str, Any],
    missing: list[dict[str, Any]],
) -> tuple[str, dict[str, Any]]:
    if material_no and material_no in catalog:
        metadata = catalog[material_no]
        if not isinstance(metadata, dict):
            raise ValueError(f"块目录中 {material_no} 的值必须是对象。")
        return str(metadata.get("block_name", material_no)), metadata

    reason = "material_no_missing" if not material_no else "block_metadata_missing"
    missing.append(
        {
            "instance_id": instance_id,
            "material_no": material_no,
            "type": item_type,
            "reason": reason,
        }
    )
    placeholder = f"__MISSING__{material_no or item_type}"
    return placeholder, {}


def looks_like_corrugated_tube(item: dict[str, Any]) -> bool:
    text = " ".join(
        str(item.get(key, ""))
        for key in ("description", "label_raw", "type", "name")
    ).lower()
    return (
        "波纹管" in text
        or "corrugated" in text
        or "protective_tube" in text
        or "corrugated_tube" in text
        or str(item.get("label_raw", "")).strip() == "2"
    )


def fallback_coverings_from_unassigned(semantic: dict[str, Any]) -> list[dict[str, Any]]:
    coverings: list[dict[str, Any]] = []
    for item in semantic.get("unassigned_items", []):
        if isinstance(item, dict) and looks_like_corrugated_tube(item):
            coverings.append(
                {
                    "instance_id": "C1",
                    "material_no": item.get("material_no") or DEFAULT_COVERING_MATERIAL_NO,
                    "label_raw": item.get("label_raw") or "2",
                    "type": "corrugated_tube",
                    "order": 1,
                    "start_exposed_length_mm": 45,
                    "end_exposed_length_mm": 45,
                    "position_ratio_start": None,
                    "position_ratio_end": None,
                    "confidence": 0.5,
                    "fallback_from_unassigned": True,
                }
            )
            break
    return coverings


def normalize_dimension_chain(raw_dimensions: Any, rules: dict[str, Any]) -> list[dict[str, Any]]:
    dimensions: list[dict[str, Any]] = []
    if isinstance(raw_dimensions, list):
        for index, item in enumerate(raw_dimensions, start=1):
            if not isinstance(item, dict):
                continue
            value = item.get("value_mm")
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                dimensions.append(
                    {
                        "instance_id": str(item.get("instance_id") or f"D{index}"),
                        "value_mm": float(value),
                        "label_raw": str(item.get("label_raw") or f"{value:g}"),
                        "order": int(item.get("order") or index),
                        "from_ref": item.get("from_ref", "unknown"),
                        "to_ref": item.get("to_ref", "unknown"),
                    }
                )
    if dimensions:
        return sorted(dimensions, key=lambda item: item["order"])

    default_chain = rules.get("default_dimension_chain_mm", [])
    if isinstance(default_chain, list):
        for index, value in enumerate(default_chain, start=1):
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                dimensions.append(
                    {
                        "instance_id": f"D{index}",
                        "value_mm": float(value),
                        "label_raw": f"{float(value):g}",
                        "order": index,
                        "from_ref": "unknown",
                        "to_ref": "unknown",
                    }
                )
    return dimensions


def generate_layout(
    semantic: dict[str, Any], rules: dict[str, Any], catalog: dict[str, Any]
) -> dict[str, Any]:
    components = semantic.get("components")
    wires = semantic.get("wires")
    if not isinstance(components, list) or not isinstance(wires, list):
        raise LayoutError("semantic.json 的 components 和 wires 必须是数组。")

    component_map = {
        item.get("instance_id"): item
        for item in components
        if isinstance(item, dict) and item.get("instance_id")
    }
    if len(component_map) != len(components):
        raise LayoutError("连接器缺少 instance_id 或 instance_id 重复。")

    left_x = float(rules["left_connector_x"])
    right_x = float(rules["right_connector_x"])
    center_y = float(rules["connector_y"])
    default_width = float(rules["connector_width"])
    pitch = float(rules["port_pitch"])
    standard_wire_length = float(rules["wire_standard_length"])
    if right_x <= left_x or default_width <= 0 or pitch <= 0 or standard_wire_length <= 0:
        raise LayoutError("排版规则中的距离和尺寸必须为正数。")

    entities: list[dict[str, Any]] = []
    drawing_order: list[str] = []
    missing_blocks: list[dict[str, Any]] = []
    warnings: list[str] = []
    component_locations: dict[str, tuple[float, float, float]] = {}

    for component in components:
        instance_id = component["instance_id"]
        relative = component.get("relative_position", "unknown")
        x = left_x if relative == "left" else right_x if relative == "right" else (left_x + right_x) / 2
        block_name, metadata = block_parameters(
            component.get("material_no"), instance_id, "connector", catalog, missing_blocks
        )
        width = float(metadata.get("width", default_width))
        component_locations[instance_id] = (x, center_y, width)
        entities.append(
            {
                "instance_id": instance_id,
                "entity_type": "block",
                "semantic_type": "connector",
                "block_name": block_name,
                "insert_point": {"x": round(x, 3), "y": round(center_y, 3)},
                "rotation_deg": float(metadata.get("rotation_deg", 0)),
                "scale": {"x": 1.0, "y": 1.0},
            }
        )

    wire_groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    wire_group_order: list[tuple[str, str]] = []
    for wire in wires:
        if not isinstance(wire, dict) or not wire.get("instance_id"):
            raise LayoutError("存在无效导线。")
        wire_id = wire["instance_id"]
        from_ref, to_ref = wire.get("from", {}), wire.get("to", {})
        from_id, to_id = from_ref.get("component"), to_ref.get("component")
        if from_id not in component_map or to_id not in component_map:
            raise LayoutError(f"导线 {wire_id} 引用了不存在的连接器。")

        from_component, to_component = component_map[from_id], component_map[to_id]
        from_x, from_center_y, from_width = component_locations[from_id]
        to_x, to_center_y, to_width = component_locations[to_id]
        start_x = from_x + from_width / 2 if to_x >= from_x else from_x - from_width / 2
        end_x = to_x - to_width / 2 if to_x >= from_x else to_x + to_width / 2
        start_y = port_y(from_ref.get("port"), from_component.get("ports", []), from_center_y, pitch)
        end_y = port_y(to_ref.get("port"), to_component.get("ports", []), to_center_y, pitch)
        dx, dy = end_x - start_x, end_y - start_y
        length = math.hypot(dx, dy)
        if length <= 0:
            raise LayoutError(f"导线 {wire_id} 的起点与终点重合。")
        group_key = tuple(sorted((str(from_id), str(to_id))))
        if group_key not in wire_groups:
            wire_groups[group_key] = []
            wire_group_order.append(group_key)
        wire_groups[group_key].append(
            {
                "wire": wire,
                "wire_id": wire_id,
                "from_id": from_id,
                "to_id": to_id,
                "from_port": from_ref.get("port"),
                "to_port": to_ref.get("port"),
                "start_x": start_x,
                "start_y": start_y,
                "end_x": end_x,
                "end_y": end_y,
            }
        )

    wire_entities: list[dict[str, Any]] = []
    covering_entities: list[dict[str, Any]] = []
    accessory_entities: list[dict[str, Any]] = []
    dimension_entities: list[dict[str, Any]] = []
    for group_key in wire_group_order:
        group_items = wire_groups[group_key]
        first = group_items[0]
        wire = first["wire"]
        wire_id = first["wire_id"]
        from_id = first["from_id"]
        to_id = first["to_id"]
        from_port = first["from_port"]
        to_port = first["to_port"]
        source_wires = [item["wire_id"] for item in group_items]

        start_x = sum(float(item["start_x"]) for item in group_items) / len(group_items)
        start_y = sum(float(item["start_y"]) for item in group_items) / len(group_items)
        end_x = sum(float(item["end_x"]) for item in group_items) / len(group_items)
        end_y = sum(float(item["end_y"]) for item in group_items) / len(group_items)
        dx, dy = end_x - start_x, end_y - start_y
        length = math.hypot(dx, dy)
        if length <= 0:
            raise LayoutError(f"合并导线 {wire_id} 的起点与终点重合。")
        angle = math.degrees(math.atan2(dy, dx))
        wire_material_no = wire.get("material_no") or DEFAULT_WIRE_MATERIAL_NO
        block_name, metadata = block_parameters(
            wire_material_no, wire_id, "wire", catalog, missing_blocks
        )
        base_length = float(metadata.get("standard_length", standard_wire_length))
        wire_entity = {
            "instance_id": wire_id,
            "entity_type": "block",
            "semantic_type": "wire",
            "block_name": block_name,
            "insert_point": {"x": round((start_x + end_x) / 2, 3), "y": round((start_y + end_y) / 2, 3)},
            "rotation_deg": round(angle, 6),
            "scale": {"x": round(length / base_length, 6), "y": 1.0},
            "connects": [f"{from_id}:{from_port}", f"{to_id}:{to_port}"],
            "path": {
                "start": {"x": round(start_x, 3), "y": round(start_y, 3)},
                "end": {"x": round(end_x, 3), "y": round(end_y, 3)},
            },
            "source_wires": source_wires,
        }
        if len(group_items) > 1:
            wire_entity["layout_note"] = "connector_pair_schematic_wire"
            warnings.append(
                f"连接器 {group_key[0]} 与 {group_key[1]} 之间 {len(group_items)} 条导线已合并为一条示意导线：{wire_id}"
            )
        wire_entities.append(wire_entity)
        anchors: dict[str, tuple[float, float]] = {
            "wire_start": (start_x, start_y),
            "wire_end": (end_x, end_y),
        }

        coverings = wire.get("coverings", [])
        if not isinstance(coverings, list):
            coverings = []
        if not coverings:
            coverings = fallback_coverings_from_unassigned(semantic)
            if coverings:
                warnings.append("已根据 unassigned_items 中的波纹管信息生成默认覆盖层 C1。")

        valid_coverings = [item for item in coverings if isinstance(item, dict)]
        for covering_index, covering in enumerate(sorted(valid_coverings, key=lambda item: item.get("order", 0)), start=1):
            covering_id = str(covering.get("instance_id") or f"C{covering_index}")
            ratio_start = covering.get("position_ratio_start")
            ratio_end = covering.get("position_ratio_end")
            if not isinstance(ratio_start, (int, float)) or not 0 <= float(ratio_start) <= 1:
                ratio_start = float(rules["covering_start_ratio"])
            if not isinstance(ratio_end, (int, float)) or not 0 <= float(ratio_end) <= 1:
                ratio_end = float(rules["covering_end_ratio"])
            ratio_start = float(ratio_start)
            ratio_end = float(ratio_end)
            if ratio_end <= ratio_start:
                warnings.append(f"覆盖层 {covering_id} 的起止比例无效，已使用默认波纹管覆盖比例。")
                ratio_start = float(rules["covering_start_ratio"])
                ratio_end = float(rules["covering_end_ratio"])

            cover_start_x, cover_start_y = start_x + ratio_start * dx, start_y + ratio_start * dy
            cover_end_x, cover_end_y = start_x + ratio_end * dx, start_y + ratio_end * dy
            cover_len = math.hypot(cover_end_x - cover_start_x, cover_end_y - cover_start_y)
            cover_angle = math.degrees(math.atan2(cover_end_y - cover_start_y, cover_end_x - cover_start_x))
            material_no = covering.get("material_no") or DEFAULT_COVERING_MATERIAL_NO
            block_name, metadata = block_parameters(material_no, covering_id, "covering", catalog, missing_blocks)
            base_cover_len = float(metadata.get("standard_length", rules["covering_standard_length"]))
            covering_entities.append(
                {
                    "instance_id": covering_id,
                    "entity_type": "block",
                    "semantic_type": covering.get("type", "corrugated_tube"),
                    "block_name": block_name,
                    "insert_point": {
                        "x": round((cover_start_x + cover_end_x) / 2, 3),
                        "y": round((cover_start_y + cover_end_y) / 2, 3),
                    },
                    "rotation_deg": round(cover_angle, 6),
                    "scale": {"x": round(cover_len / base_cover_len, 6), "y": 1.0},
                    "attached_to": wire_id,
                    "position_ratio_start": round(ratio_start, 6),
                    "position_ratio_end": round(ratio_end, 6),
                    "path": {
                        "start": {"x": round(cover_start_x, 3), "y": round(cover_start_y, 3)},
                        "end": {"x": round(cover_end_x, 3), "y": round(cover_end_y, 3)},
                    },
                    "start_exposed_length_mm": covering.get("start_exposed_length_mm"),
                    "end_exposed_length_mm": covering.get("end_exposed_length_mm"),
                }
            )
            anchors["covering_start"] = (cover_start_x, cover_start_y)
            anchors["covering_end"] = (cover_end_x, cover_end_y)
            anchors[f"{covering_id}_start"] = (cover_start_x, cover_start_y)
            anchors[f"{covering_id}_end"] = (cover_end_x, cover_end_y)

        accessories = wire.get("accessories", [])
        valid_accessories = [
            item
            for item in accessories
            if isinstance(item, dict) and item.get("material_no")
        ]
        skipped_accessories = [
            item.get("instance_id", "<unknown>")
            for item in accessories
            if isinstance(item, dict) and not item.get("material_no")
        ]
        if skipped_accessories:
            warnings.append(
                "已忽略缺少 material_no 的疑似辅材："
                + ", ".join(str(item) for item in skipped_accessories)
            )
        for accessory in sorted(valid_accessories, key=lambda item: item.get("order", 0)):
            accessory_id = accessory.get("instance_id")
            ratio = accessory.get("position_ratio")
            if ratio is None:
                count = len(valid_accessories)
                ratio = accessory.get("order", 1) / (count + 1)
                warnings.append(f"辅材 {accessory_id} 缺少 position_ratio，已按顺序均匀放置。")
            ratio = float(ratio)
            if not 0 <= ratio <= 1:
                raise LayoutError(f"辅材 {accessory_id} 的 position_ratio 超出 0 到 1。")
            x, y = start_x + ratio * dx, start_y + ratio * dy
            side = accessory.get("side", "on_wire")
            if side == "above":
                y += float(rules["above_offset"])
            elif side == "below":
                y += float(rules["below_offset"])
            else:
                y += float(rules["accessory_offset"])
            anchors[f"accessory:{accessory_id}"] = (x, y)
            anchors[str(accessory_id)] = (x, y)
            block_name, metadata = block_parameters(
                accessory.get("material_no"), accessory_id, accessory.get("type", "unknown"), catalog, missing_blocks
            )
            accessory_entities.append(
                {
                    "instance_id": accessory_id,
                    "entity_type": "block",
                    "semantic_type": accessory.get("type", "unknown"),
                    "block_name": block_name,
                    "insert_point": {"x": round(x, 3), "y": round(y, 3)},
                    "rotation_deg": round(angle + float(metadata.get("rotation_offset_deg", 0)), 6),
                    "scale": {"x": 1.0, "y": 1.0},
                    "attached_to": wire_id,
                    "position_ratio": ratio,
                    "order": accessory.get("order"),
                }
            )

        dimensions = normalize_dimension_chain(wire.get("dimensions", []), rules)
        if dimensions:
            dim_y = start_y + float(rules["dimension_offset_y"])
            tick = float(rules["dimension_tick_height"])
            text_h = float(rules["dimension_text_height"])
            for dim_index, dim in enumerate(dimensions):
                from_ref = str(dim.get("from_ref", "unknown"))
                to_ref = str(dim.get("to_ref", "unknown"))
                from_anchor = anchors.get(from_ref)
                to_anchor = anchors.get(to_ref)
                if from_anchor is None or to_anchor is None:
                    span_count = len(dimensions)
                    r0 = dim_index / span_count
                    r1 = (dim_index + 1) / span_count
                    from_anchor = (start_x + r0 * dx, start_y + r0 * dy)
                    to_anchor = (start_x + r1 * dx, start_y + r1 * dy)
                    warnings.append(
                        f"尺寸 {dim['instance_id']} 的锚点 {from_ref}->{to_ref} 未找到，已退回等分放置。"
                    )

                x0 = float(from_anchor[0])
                x1 = float(to_anchor[0])
                if x1 < x0:
                    x0, x1 = x1, x0
                line_id = f"{dim['instance_id']}_LINE"
                text_id = f"{dim['instance_id']}_TEXT"
                dimension_entities.append(
                    {
                        "instance_id": line_id,
                        "entity_type": "line",
                        "semantic_type": "dimension_line",
                        "start": {"x": round(x0, 3), "y": round(dim_y, 3)},
                        "end": {"x": round(x1, 3), "y": round(dim_y, 3)},
                        "attached_to": wire_id,
                        "value_mm": dim["value_mm"],
                        "from_ref": from_ref,
                        "to_ref": to_ref,
                    }
                )
                dimension_entities.append(
                    {
                        "instance_id": f"{line_id}_TICK_L",
                        "entity_type": "line",
                        "semantic_type": "dimension_tick",
                        "start": {"x": round(x0, 3), "y": round(dim_y - tick / 2, 3)},
                        "end": {"x": round(x0, 3), "y": round(dim_y + tick / 2, 3)},
                        "attached_to": wire_id,
                    }
                )
                dimension_entities.append(
                    {
                        "instance_id": f"{line_id}_TICK_R",
                        "entity_type": "line",
                        "semantic_type": "dimension_tick",
                        "start": {"x": round(x1, 3), "y": round(dim_y - tick / 2, 3)},
                        "end": {"x": round(x1, 3), "y": round(dim_y + tick / 2, 3)},
                        "attached_to": wire_id,
                    }
                )
                dimension_entities.append(
                    {
                        "instance_id": text_id,
                        "entity_type": "text",
                        "semantic_type": "dimension_text",
                        "insert_point": {"x": round((x0 + x1) / 2, 3), "y": round(dim_y + text_h, 3)},
                        "text": dim["label_raw"],
                        "height": text_h,
                        "attached_to": wire_id,
                        "value_mm": dim["value_mm"],
                        "from_ref": from_ref,
                        "to_ref": to_ref,
                    }
                )

    # 先画导线，再画波纹管覆盖层，再画连接器和辅材，最后画尺寸标注。
    entities = wire_entities + covering_entities + entities + accessory_entities + dimension_entities
    drawing_order = [entity["instance_id"] for entity in entities]
    semantic_validation = semantic.get("validation", {})
    unresolved = semantic_validation.get("unresolved_items", []) if isinstance(semantic_validation, dict) else []
    if unresolved:
        warnings.append("semantic.json 存在未解决项目，布局只能作为草稿。")

    return {
        "schema_version": "1.0",
        "units": rules["units"],
        "status": "draft" if missing_blocks or unresolved else "ready",
        "entities": entities,
        "drawing_order": drawing_order,
        "validation": {
            "missing_blocks": missing_blocks,
            "overlaps": [],
            "warnings": warnings,
            "manual_review_required": bool(missing_blocks or unresolved),
        },
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="把接线语义转换为确定性 CAXA 块布局。")
    parser.add_argument("semantic", type=Path, help="输入 semantic.json")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/layout.json"))
    parser.add_argument("--catalog", type=Path, help="可选的块元数据 block_catalog.json")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        semantic = load_json(args.semantic)
        rules, catalog = merge_rules(args.catalog)
        layout = generate_layout(semantic, rules, catalog)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(layout, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"布局已生成：{args.output.resolve()}")
        print(f"状态：{layout['status']}，实体数：{len(layout['entities'])}")
        print(f"缺失块：{len(layout['validation']['missing_blocks'])}")
        return 0
    except Exception as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
