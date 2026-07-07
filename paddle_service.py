"""OCR service call example.

This helper is optional. The web backend calls the OCR endpoint directly in
`jonhon_web_app.py`; this file is only for manually testing an OCR endpoint.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path

import requests


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_pdf", type=Path, help="PDF drawing to send to OCR service.")
    parser.add_argument(
        "--ocr-url",
        default=os.getenv("JONHON_OCR_URL", "http://127.0.0.1:8030/ocr"),
        help="OCR service endpoint.",
    )
    parser.add_argument("--output", type=Path, default=Path("ocr_result.json"))
    parser.add_argument("--timeout", type=int, default=int(os.getenv("JONHON_OCR_TIMEOUT", "300")))
    args = parser.parse_args()

    with args.input_pdf.open("rb") as fh:
        response = requests.post(args.ocr_url, files={"file": fh}, timeout=args.timeout)
    response.raise_for_status()

    try:
        payload = response.json()
        args.output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except ValueError:
        args.output.write_text(response.text, encoding="utf-8")

    print(f"OCR result saved to: {args.output}")


if __name__ == "__main__":
    main()
