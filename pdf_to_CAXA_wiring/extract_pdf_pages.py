#!/usr/bin/env python3
"""将 PDF 图纸渲染为适合多模态模型识别的 PNG，并生成页面清单。
cd "D:\神州数码\业务\中航光电"

conda run -n quankepytorch python `
  .\pdf_to_caxa_wiring\extract_pdf_pages.py `
  ".\llm2img\用户提供图纸1.pdf" `
  -o ".\pdf_to_caxa_wiring\output\pages" `
  --dpi 200
  """

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - 只在依赖未安装时执行
    fitz = None  # type: ignore[assignment]


DEFAULT_DPI = 200


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_page_range(value: str | None, page_count: int) -> list[int]:
    """把用户输入的 1-based 页码范围转换为 0-based 页码列表。"""
    if not value:
        return list(range(page_count))

    selected: set[int] = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start, end = int(start_text), int(end_text)
            if start > end:
                raise ValueError(f"页码范围起点不能大于终点：{part}")
            selected.update(range(start, end + 1))
        else:
            selected.add(int(part))

    if not selected:
        raise ValueError("没有选择任何页面。")
    invalid = sorted(page for page in selected if page < 1 or page > page_count)
    if invalid:
        raise ValueError(f"页码超出范围 1-{page_count}：{invalid}")
    return [page - 1 for page in sorted(selected)]


def render_pdf(
    pdf_path: Path,
    output_dir: Path,
    dpi: int = DEFAULT_DPI,
    pages: str | None = None,
) -> Path:
    if fitz is None:
        raise RuntimeError("缺少 PyMuPDF，请先执行：python -m pip install PyMuPDF")
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF 文件不存在：{pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        raise ValueError(f"输入文件不是 PDF：{pdf_path}")
    if not 72 <= dpi <= 600:
        raise ValueError("DPI 必须在 72 到 600 之间。")

    output_dir.mkdir(parents=True, exist_ok=True)
    document = fitz.open(pdf_path)
    try:
        if document.needs_pass:
            raise ValueError("PDF 已加密，需要先解除密码保护。")

        page_indexes = parse_page_range(pages, document.page_count)
        page_records: list[dict[str, Any]] = []
        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)

        for page_index in page_indexes:
            page = document.load_page(page_index)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False, colorspace=fitz.csRGB)
            image_name = f"page_{page_index + 1:04d}.png"
            image_path = output_dir / image_name
            pixmap.save(image_path)

            page_records.append(
                {
                    "page_number": page_index + 1,
                    "image_file": image_name,
                    "width_px": pixmap.width,
                    "height_px": pixmap.height,
                    "dpi": dpi,
                    "sha256": sha256_file(image_path),
                }
            )
            print(f"已提取第 {page_index + 1} 页：{image_path}")

        manifest = {
            "schema_version": "1.0",
            "created_at": datetime.now(timezone.utc).isoformat(),
            "source_pdf": {
                "file_name": pdf_path.name,
                "absolute_path": str(pdf_path.resolve()),
                "sha256": sha256_file(pdf_path),
                "total_pages": document.page_count,
            },
            "rendering": {"format": "png", "color_space": "RGB", "dpi": dpi},
            "pages": page_records,
        }
        manifest_path = output_dir / "manifest.json"
        manifest_path.write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return manifest_path
    finally:
        document.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="把 PDF 图纸逐页渲染为 PNG，供阿里百炼多模态模型识别。"
    )
    parser.add_argument("pdf", type=Path, help="输入 PDF 文件")
    parser.add_argument(
        "-o", "--output-dir", type=Path, default=Path("output/pages"), help="输出目录"
    )
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI, help="渲染 DPI，默认 200")
    parser.add_argument(
        "--pages", help="可选页码，例如 1、1-3 或 1,3,5-7；默认提取全部页面"
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        manifest_path = render_pdf(args.pdf, args.output_dir, args.dpi, args.pages)
        print(f"提取完成，页面清单：{manifest_path.resolve()}")
        return 0
    except Exception as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
