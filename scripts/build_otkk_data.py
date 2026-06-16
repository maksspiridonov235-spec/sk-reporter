#!/usr/bin/env python3
"""Собрать sk_reporter/otkk{N}_data.py дословно из .doc (textutil, без правок)."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sk_reporter.otkk_verbatim import extract_six_rows_from_doc


def _py_str(s: str) -> str:
    return json.dumps(s, ensure_ascii=False)


def write_data_module(parsed: dict, out_path: Path, func_name: str, row6_var: str) -> None:
    row6 = parsed["rows"][5]["value"]
    lines = [
        f'"""ОТКК: шесть пунктов карты — дословно из {parsed["file"]}."""',
        "",
        "from __future__ import annotations",
        "",
        "from typing import Any",
        "",
        "",
        f"{row6_var} = {_py_str(row6)}",
        "",
        "",
        f"def {func_name}() -> dict[str, Any]:",
        "    rows = [",
    ]
    for row in parsed["rows"][:-1]:
        lines.append(f"        {{'label': {_py_str(row['label'])}, 'value': {_py_str(row['value'])}}},")
    lines.append("        {")
    lines.append(f"            'label': {_py_str(parsed['rows'][5]['label'])},")
    lines.append(f"            'value': {row6_var},")
    lines.append("        },")
    lines.append("    ]")
    lines.append("    return {")
    lines.append(f"        'id': {_py_str(parsed['id'])},")
    lines.append(f"        'code': {_py_str(parsed['code'])},")
    lines.append(f"        'title': {_py_str(parsed['title'])},")
    lines.append(f"        'file': {_py_str(parsed['file'])},")
    lines.append("        'rows': rows,")
    lines.append("        'signature': None,")
    lines.append(f"        'plain_text': {_py_str(parsed.get('plain_text') or '')},")
    lines.append("    }")
    lines.append("")
    out_path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("doc", type=Path, help="Путь к .doc/.docx")
    parser.add_argument("--out", type=Path, help="Выходной модуль (по умолчанию otkk{N}_data.py)")
    args = parser.parse_args()

    parsed = extract_six_rows_from_doc(args.doc)
    n = parsed["id"].split("-")[-1]
    out = args.out or ROOT / "sk_reporter" / f"otkk{n}_data.py"
    write_data_module(parsed, out, f"otkk{n}_parsed", f"_OTKK{n.upper()}_ROW6")
    print(f"OK {parsed['id']} -> {out}")
    for i, row in enumerate(parsed["rows"], 1):
        print(f"  {i}. {row['label'][:50]} — {len(row['value'])} симв.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
