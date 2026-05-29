#!/usr/bin/env python3
"""
NN33_가공본.xlsx (JLPT N3 AI 가공 엑셀) → jlpt_n3_expanded.json

엑셀 구조 (시트 `n3`, 2행이 헤더):
  expression, reading, meaning, part_of_speech,
  japanese_example, korean_example, words (영문),
  alt1_expression / alt1_furigana … alt8_*,
  tip1_expression / tip1_furigana / tip1_meaning / tip1_pos … tip2_*

사용 예:
  pip install -r requirements.txt
  python scripts/xlsx_to_json.py ^
    -i "data/NN33_가공본.xlsx" ^
    -o jlpt_n3_expanded.json
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ImportError:
    print("pandas가 필요합니다: pip install -r requirements.txt", file=sys.stderr)
    raise SystemExit(1)


def is_empty(val: Any) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    return str(val).strip() == ""


def cell_str(val: Any) -> str:
    if is_empty(val):
        return ""
    return str(val).strip()


def build_alt_spellings(row: pd.Series, max_alt: int = 8) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for i in range(1, max_alt + 1):
        kj = cell_str(row.get(f"alt{i}_expression"))
        fg = cell_str(row.get(f"alt{i}_furigana"))
        if kj and fg:
            out.append({"kanji": kj, "furigana": fg})
    return out


def build_tips(row: pd.Series, max_tip: int = 2) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for i in range(1, max_tip + 1):
        kj = cell_str(row.get(f"tip{i}_expression"))
        if not kj:
            continue
        out.append(
            {
                "kanji": kj,
                "furigana": cell_str(row.get(f"tip{i}_furigana")),
                "meaning": cell_str(row.get(f"tip{i}_meaning")),
                "pos": cell_str(row.get(f"tip{i}_pos")) or "—",
            }
        )
    return out


def row_to_record(row: pd.Series, idx: int, level: str) -> dict[str, Any] | None:
    kanji = cell_str(row.get("expression"))
    if not kanji:
        return None
    return {
        "id": idx,
        "level": level,
        "kanji": kanji,
        "furigana": cell_str(row.get("reading")),
        "meaning": cell_str(row.get("meaning")),
        "pos": cell_str(row.get("part_of_speech")) or "—",
        "meaning_en": cell_str(row.get("words")),
        "example_jp": cell_str(row.get("japanese_example")),
        "example_kr": cell_str(row.get("korean_example")),
        "alt_spellings": build_alt_spellings(row),
        "tips": build_tips(row),
    }


def load_sheet(path: Path, sheet: str, header_row: int) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, header=header_row, engine="openpyxl")
    # 컬럼명 공백 제거
    df.columns = [str(c).strip() for c in df.columns]
    return df


def validate_records(records: list[dict[str, Any]], strict: bool) -> list[str]:
    errors: list[str] = []
    for rec in records:
        rid = rec.get("id")
        for key in ("kanji", "furigana", "meaning", "meaning_en"):
            if not rec.get(key):
                errors.append(f"id {rid}: '{key}' 비어 있음")
        if strict and not rec.get("example_jp"):
            errors.append(f"id {rid}: example_jp 없음")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(
        description="NN33_가공본.xlsx → jlpt_n3_expanded.json (JLPT N3)"
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=Path("data/NN33_가공본.xlsx"),
        help="입력 엑셀 (기본: data/NN33_가공본.xlsx)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("jlpt_n3_expanded.json"),
        help="출력 JSON (기본: 프로젝트 루트)",
    )
    parser.add_argument("--sheet", default="n3", help="시트 이름 (기본: n3)")
    parser.add_argument(
        "--header-row",
        type=int,
        default=1,
        help="헤더가 있는 행 번호, 0부터 (기본: 1 = 엑셀 2행)",
    )
    parser.add_argument("--level", default="N3", help="JLPT 레벨 태그")
    parser.add_argument("--indent", type=int, default=4, help="JSON 들여쓰기")
    parser.add_argument("--strict", action="store_true", help="필수 필드 누락 시 종료")
    parser.add_argument("--dry-run", action="store_true", help="파일 저장 없이 검증만")
    args = parser.parse_args()

    if not args.input.is_file():
        print(f"엑셀 파일 없음: {args.input}", file=sys.stderr)
        print(
            "  원본 예: jlpt 웹 사이트 자료/N1~N5엑섹파일/NN33_가공본.xlsx\n"
            "  → data/ 폴더에 복사한 뒤 다시 실행하세요.",
            file=sys.stderr,
        )
        return 1

    df = load_sheet(args.input, args.sheet, args.header_row)
    records: list[dict[str, Any]] = []
    skipped = 0

    for _, row in df.iterrows():
        rec = row_to_record(row, len(records) + 1, args.level)
        if rec is None:
            skipped += 1
            continue
        records.append(rec)

    errors = validate_records(records, args.strict)
    if errors:
        for e in errors[:25]:
            print(e, file=sys.stderr)
        if len(errors) > 25:
            print(f"... 외 {len(errors) - 25}건", file=sys.stderr)
        if args.strict:
            return 1

    print(f"시트 '{args.sheet}': {len(records)}개 변환 (스킵 {skipped}행)")
    print(f"입력: {args.input.resolve()}")

    if args.dry_run:
        return 0

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8") as f:
        if args.indent:
            json.dump(records, f, ensure_ascii=False, indent=args.indent)
        else:
            json.dump(records, f, ensure_ascii=False, separators=(",", ":"))
        f.write("\n")

    print(f"저장: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
