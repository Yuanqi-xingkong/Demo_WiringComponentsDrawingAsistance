#!/usr/bin/env python3
"""由 layout.json 通过固定模板生成 VS2010/ObjectCRX C++ 命令源码。

conda run -n quankepytorch python `
  .\pdf_to_caxa_wiring\generate_caxa_cpp.py `
  .\pdf_to_caxa_wiring\output\layout.json `
  -o .\pdf_to_caxa_wiring\output\AutoDrawWiring.cpp
 
conda run -n quankepytorch python `
  .\pdf_to_caxa_wiring\generate_caxa_cpp.py `
  .\pdf_to_caxa_wiring\output\layout.json `
  -o .\pdf_to_caxa_wiring\output\AutoDrawWiring.cpp `
  --allow-draft
 """

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


class CppGenerationError(RuntimeError):
    pass


def load_layout(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"layout.json 不存在：{path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise CppGenerationError("layout.json 顶层必须是对象。")
    return data


def cpp_string(value: str) -> str:
    """转义为 _T("...") 中可安全使用的 C++ 字符串。"""
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    escaped = escaped.replace("\r", "\\r").replace("\n", "\\n")
    return escaped


def cpp_number(value: Any, field: str) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise CppGenerationError(f"{field} 必须是数值。")
    number = float(value)
    if not math.isfinite(number):
        raise CppGenerationError(f"{field} 必须是有限数值。")
    if number == 0:
        number = 0.0
    return f"{number:.9g}"


def validate_layout(layout: dict[str, Any], allow_draft: bool) -> list[dict[str, Any]]:
    if layout.get("units") != "mm":
        raise CppGenerationError("当前模板只接受 units=mm。")
    if layout.get("status") != "ready" and not allow_draft:
        raise CppGenerationError(
            "layout.json 不是 ready 状态。补齐块目录后重建布局，"
            "或仅为联调使用 --allow-draft。"
        )
    entities = layout.get("entities")
    if not isinstance(entities, list) or not entities:
        raise CppGenerationError("entities 必须是非空数组。")

    by_id: dict[str, dict[str, Any]] = {}
    for index, entity in enumerate(entities):
        if not isinstance(entity, dict):
            raise CppGenerationError(f"entities[{index}] 不是对象。")
        instance_id = entity.get("instance_id")
        if not isinstance(instance_id, str) or not instance_id:
            raise CppGenerationError(f"entities[{index}] 缺少 instance_id。")
        if instance_id in by_id:
            raise CppGenerationError(f"instance_id 重复：{instance_id}")
        entity_type = entity.get("entity_type")
        if entity_type == "block":
            block_name = entity.get("block_name")
            if not isinstance(block_name, str) or not block_name:
                raise CppGenerationError(f"实体 {instance_id} 缺少 block_name。")
            if block_name.startswith("__MISSING__") and not allow_draft:
                raise CppGenerationError(f"实体 {instance_id} 使用缺失块占位名。")
            point = entity.get("insert_point")
            scale = entity.get("scale")
            if not isinstance(point, dict) or not isinstance(scale, dict):
                raise CppGenerationError(f"实体 {instance_id} 缺少 insert_point 或 scale。")
            cpp_number(point.get("x"), f"{instance_id}.insert_point.x")
            cpp_number(point.get("y"), f"{instance_id}.insert_point.y")
            sx = float(scale.get("x"))
            sy = float(scale.get("y"))
            if sx <= 0 or sy <= 0:
                raise CppGenerationError(f"实体 {instance_id} 的缩放比例必须大于0。")
            cpp_number(entity.get("rotation_deg", 0), f"{instance_id}.rotation_deg")
        elif entity_type == "line":
            start = entity.get("start")
            end = entity.get("end")
            if not isinstance(start, dict) or not isinstance(end, dict):
                raise CppGenerationError(f"线实体 {instance_id} 缺少 start 或 end。")
            cpp_number(start.get("x"), f"{instance_id}.start.x")
            cpp_number(start.get("y"), f"{instance_id}.start.y")
            cpp_number(end.get("x"), f"{instance_id}.end.x")
            cpp_number(end.get("y"), f"{instance_id}.end.y")
        elif entity_type == "text":
            point = entity.get("insert_point")
            if not isinstance(point, dict):
                raise CppGenerationError(f"文字实体 {instance_id} 缺少 insert_point。")
            cpp_number(point.get("x"), f"{instance_id}.insert_point.x")
            cpp_number(point.get("y"), f"{instance_id}.insert_point.y")
            cpp_number(entity.get("height", 5.0), f"{instance_id}.height")
            if not isinstance(entity.get("text"), str):
                raise CppGenerationError(f"文字实体 {instance_id} 缺少 text。")
        else:
            raise CppGenerationError(f"实体 {instance_id} 的 entity_type 不支持：{entity_type}")
        by_id[instance_id] = entity

    order = layout.get("drawing_order")
    if not isinstance(order, list) or set(order) != set(by_id) or len(order) != len(by_id):
        raise CppGenerationError("drawing_order 必须且只能包含每个 entity 的 instance_id 一次。")
    return [by_id[instance_id] for instance_id in order]


CPP_TEMPLATE = r'''// Generated from layout.json. Do not edit entity calls manually.
// Target: CAXA ObjectCRX, Visual Studio 2010.
#include "StdAfx.h"
#include "dbents.h"
#include "dbsymtb.h"
#include <math.h>

static bool InsertConfiguredBlock(
    const TCHAR* blockName,
    double x,
    double y,
    double rotationDeg,
    double scaleX,
    double scaleY)
{
    CRxDbDatabase* pDb = crxdbHostApplicationServices()->workingDatabase();
    if (pDb == NULL)
    {
        crxutPrintf(_T("\nAutoWiring: no working database."));
        return false;
    }

    CRxDbBlockTable* pBlockTable = NULL;
    CDraft::ErrorStatus es = pDb->getBlockTable(pBlockTable, CRxDb::kForRead);
    if (es != CDraft::eOk || pBlockTable == NULL)
    {
        crxutPrintf(_T("\nAutoWiring: cannot open block table."));
        return false;
    }

    if (!pBlockTable->has(blockName))
    {
        crxutPrintf(_T("\nAutoWiring: block not found: %s"), blockName);
        pBlockTable->close();
        return false;
    }

    CRxDbObjectId blockDefinitionId;
    es = pBlockTable->getAt(blockName, blockDefinitionId);
    if (es != CDraft::eOk)
    {
        pBlockTable->close();
        return false;
    }

    CRxDbBlockTableRecord* pModelSpace = NULL;
    es = pBlockTable->getAt(CRXDB_MODEL_SPACE, pModelSpace, CRxDb::kForWrite);
    pBlockTable->close();
    pBlockTable = NULL;
    if (es != CDraft::eOk || pModelSpace == NULL)
    {
        crxutPrintf(_T("\nAutoWiring: cannot open model space."));
        return false;
    }

    CRxDbBlockReference* pBlockReference = new CRxDbBlockReference();
    if (pBlockReference == NULL)
    {
        pModelSpace->close();
        return false;
    }
    pBlockReference->setPosition(CRxGePoint3d(x, y, 0.0));
    pBlockReference->setBlockTableRecord(blockDefinitionId);

    // 当前 hello2/CAXA ObjectCRX 环境中 setRotation(0) / setScaleFactors
    // 会在运行时返回错误，因此改用矩阵 transformBy 进行必要缩放。
    if (rotationDeg != 0.0)
    {
        crxutPrintf(_T("\nAutoWiring: rotation is not supported yet: %s"), blockName);
    }
    if (scaleX != 1.0 || scaleY != 1.0)
    {
        CRxGeMatrix3d scaleMat;
        scaleMat.setToScaling(scaleX, CRxGePoint3d(x, y, 0.0));
        es = pBlockReference->transformBy(scaleMat);
        if (es != CDraft::eOk)
        {
            delete pBlockReference;
            pModelSpace->close();
            crxutPrintf(_T("\nAutoWiring: failed to transform-scale block: %s"), blockName);
            return false;
        }
    }

    CRxDbObjectId blockReferenceId;
    es = pModelSpace->appendAcDbEntity(blockReferenceId, pBlockReference);
    if (es != CDraft::eOk)
    {
        delete pBlockReference;
        pModelSpace->close();
        crxutPrintf(_T("\nAutoWiring: failed to append block: %s"), blockName);
        return false;
    }

    pBlockReference->close();
    pModelSpace->close();
    return true;
}

static bool AddConfiguredLine(double x1, double y1, double x2, double y2)
{
    CRxDbDatabase* pDb = crxdbHostApplicationServices()->workingDatabase();
    if (pDb == NULL)
        return false;

    CRxDbBlockTable* pBlockTable = NULL;
    CDraft::ErrorStatus es = pDb->getBlockTable(pBlockTable, CRxDb::kForRead);
    if (es != CDraft::eOk || pBlockTable == NULL)
        return false;

    CRxDbBlockTableRecord* pModelSpace = NULL;
    es = pBlockTable->getAt(CRXDB_MODEL_SPACE, pModelSpace, CRxDb::kForWrite);
    pBlockTable->close();
    if (es != CDraft::eOk || pModelSpace == NULL)
        return false;

    CRxDbLine* pLine = new CRxDbLine(
        CRxGePoint3d(x1, y1, 0.0),
        CRxGePoint3d(x2, y2, 0.0));
    if (pLine == NULL)
    {
        pModelSpace->close();
        return false;
    }

    CRxDbObjectId lineId;
    es = pModelSpace->appendAcDbEntity(lineId, pLine);
    if (es != CDraft::eOk)
    {
        delete pLine;
        pModelSpace->close();
        return false;
    }

    pLine->close();
    pModelSpace->close();
    return true;
}

static bool AddConfiguredText(const TCHAR* text, double x, double y, double height)
{
    CRxDbDatabase* pDb = crxdbHostApplicationServices()->workingDatabase();
    if (pDb == NULL)
        return false;

    CRxDbBlockTable* pBlockTable = NULL;
    CDraft::ErrorStatus es = pDb->getBlockTable(pBlockTable, CRxDb::kForRead);
    if (es != CDraft::eOk || pBlockTable == NULL)
        return false;

    CRxDbBlockTableRecord* pModelSpace = NULL;
    es = pBlockTable->getAt(CRXDB_MODEL_SPACE, pModelSpace, CRxDb::kForWrite);
    pBlockTable->close();
    if (es != CDraft::eOk || pModelSpace == NULL)
        return false;

    CRxDbText* pText = new CRxDbText();
    if (pText == NULL)
    {
        pModelSpace->close();
        return false;
    }
    pText->setPosition(CRxGePoint3d(x, y, 0.0));
    pText->setHeight(height);
    pText->setTextString(text);

    CRxDbObjectId textId;
    es = pModelSpace->appendAcDbEntity(textId, pText);
    if (es != CDraft::eOk)
    {
        delete pText;
        pModelSpace->close();
        return false;
    }

    pText->close();
    pModelSpace->close();
    return true;
}

void cmdAutoDrawWiring()
{
    crxutPrintf(_T("\nAutoWiring: start."));
    int successCount = 0;
    int failureCount = 0;

{insert_calls}

    crxutPrintf(
        _T("\nAutoWiring: finished. success=%d, failed=%d."),
        successCount,
        failureCount);
}

/*
Register this command in On_kInitAppMsg with the same style used by the
current hello2 ObjectCRX project:

    crxedRegCmds->addCommand(
        _T("HelloApp"),
        _T("GAutoDrawWiring"),
        _T("AutoDrawWiring"),
        ACRX_CMD_MODAL,
        &cmdAutoDrawWiring
    );

Do not use crxutAddCommand / CRxCmd::kModal in this project version.
*/
'''


def generate_cpp(layout: dict[str, Any], allow_draft: bool) -> str:
    ordered_entities = validate_layout(layout, allow_draft)
    calls: list[str] = []
    for entity in ordered_entities:
        instance_id = entity["instance_id"]
        if not re.fullmatch(r"[A-Za-z0-9_.:-]+", instance_id):
            raise CppGenerationError(f"instance_id 含不支持的字符：{instance_id}")
        entity_type = entity.get("entity_type")
        if entity_type == "block":
            point = entity["insert_point"]
            scale = entity["scale"]
            args = (
                f'_T("{cpp_string(entity["block_name"])}"), '
                f'{cpp_number(point["x"], instance_id + ".x")}, '
                f'{cpp_number(point["y"], instance_id + ".y")}, '
                f'{cpp_number(entity.get("rotation_deg", 0), instance_id + ".rotation")}, '
                f'{cpp_number(scale["x"], instance_id + ".scale.x")}, '
                f'{cpp_number(scale["y"], instance_id + ".scale.y")}'
            )
            call = f"InsertConfiguredBlock({args})"
        elif entity_type == "line":
            start = entity["start"]
            end = entity["end"]
            call = (
                "AddConfiguredLine("
                f'{cpp_number(start["x"], instance_id + ".start.x")}, '
                f'{cpp_number(start["y"], instance_id + ".start.y")}, '
                f'{cpp_number(end["x"], instance_id + ".end.x")}, '
                f'{cpp_number(end["y"], instance_id + ".end.y")})'
            )
        elif entity_type == "text":
            point = entity["insert_point"]
            call = (
                "AddConfiguredText("
                f'_T("{cpp_string(entity["text"])}"), '
                f'{cpp_number(point["x"], instance_id + ".x")}, '
                f'{cpp_number(point["y"], instance_id + ".y")}, '
                f'{cpp_number(entity.get("height", 5.0), instance_id + ".height")})'
            )
        else:
            raise CppGenerationError(f"实体 {instance_id} 的 entity_type 不支持：{entity_type}")
        calls.extend(
            [
                f"    // {instance_id}: {entity.get('semantic_type', entity_type)}",
                f"    if ({call})",
                "        ++successCount;",
                "    else",
                "        ++failureCount;",
                "",
            ]
        )
    return CPP_TEMPLATE.replace("{insert_calls}", "\n".join(calls).rstrip())


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="由 layout.json 生成固定模板的 ObjectCRX C++ 源码。")
    parser.add_argument("layout", type=Path, help="输入 layout.json")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/AutoDrawWiring.cpp"))
    parser.add_argument(
        "--allow-draft",
        action="store_true",
        help="允许占位块，仅用于检查生成代码，禁止作为正式图纸",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        layout = load_layout(args.layout)
        code = generate_cpp(layout, args.allow_draft)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(code, encoding="utf-8-sig", newline="\n")
        print(f"C++ 已生成：{args.output.resolve()}")
        print("命令函数：cmdAutoDrawWiring")
        if args.allow_draft:
            print("警告：当前使用 --allow-draft，缺失块会在 CAXA 运行时失败。")
        return 0
    except Exception as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
