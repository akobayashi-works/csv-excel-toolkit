"""csv-excel-toolkit のコマンドラインインターフェース.

サブコマンド:
  merge    複数CSVの結合（列の並び違い・列の有無を吸収）
  dedupe   重複行の除去（キー列指定可）
  split    指定列の値ごとにファイルを分割
  summary  グループ別集計（合計・件数）をCSV出力
  to-excel CSV → Excel(.xlsx) 変換（openpyxl が必要）
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from .io_utils import format_number, parse_number, read_csv, write_csv


# ---------------------------------------------------------------- merge
def cmd_merge(args) -> int:
    merged_fields: list[str] = []
    merged_rows: list[dict] = []
    for path in args.inputs:
        fields, rows = read_csv(path, args.encoding)
        for name in fields:
            if name not in merged_fields:
                merged_fields.append(name)
        merged_rows.extend(rows)
    write_csv(args.output, merged_fields, merged_rows, args.out_encoding)
    print(f"merge: {len(args.inputs)} ファイル / {len(merged_rows)} 行 "
          f"（列: {', '.join(merged_fields)}） -> {args.output}")
    return 0


# ---------------------------------------------------------------- dedupe
def cmd_dedupe(args) -> int:
    fields, rows = read_csv(args.input, args.encoding)
    keys = [k.strip() for k in args.key.split(",")] if args.key else fields
    missing = [k for k in keys if k not in fields]
    if missing:
        print(f"エラー: キー列が見つかりません: {', '.join(missing)}", file=sys.stderr)
        return 2
    seen = set()
    unique = []
    for row in rows:
        marker = tuple(row.get(k, "") for k in keys)
        if marker in seen:
            continue
        seen.add(marker)
        unique.append(row)
    write_csv(args.output, fields, unique, args.out_encoding)
    print(f"dedupe: {len(rows)} 行 -> {len(unique)} 行 "
          f"（{len(rows) - len(unique)} 行の重複を除去） -> {args.output}")
    return 0


# ---------------------------------------------------------------- split
def _safe_filename(value: str) -> str:
    name = re.sub(r'[\\/:*?"<>|\s]+', "_", str(value).strip())
    return name or "blank"


def cmd_split(args) -> int:
    fields, rows = read_csv(args.input, args.encoding)
    if args.column not in fields:
        print(f"エラー: 列が見つかりません: {args.column}", file=sys.stderr)
        return 2
    groups: dict[str, list[dict]] = {}
    for row in rows:
        groups.setdefault(row.get(args.column, ""), []).append(row)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    for value, group_rows in groups.items():
        out_path = outdir / f"{_safe_filename(value)}.csv"
        write_csv(out_path, fields, group_rows, args.out_encoding)
        print(f"split: {value!r} -> {out_path}（{len(group_rows)} 行）")
    print(f"split: 合計 {len(groups)} ファイル / {len(rows)} 行")
    return 0


# ---------------------------------------------------------------- summary
def cmd_summary(args) -> int:
    fields, rows = read_csv(args.input, args.encoding)
    group_cols = [c.strip() for c in args.group_by.split(",")]
    for col in group_cols + ([args.sum] if args.sum else []):
        if col not in fields:
            print(f"エラー: 列が見つかりません: {col}", file=sys.stderr)
            return 2
    groups: dict[tuple, dict] = {}
    order: list[tuple] = []
    skipped = 0
    for row in rows:
        key = tuple(row.get(c, "") for c in group_cols)
        if key not in groups:
            groups[key] = {"count": 0, "total": 0}
            order.append(key)
        groups[key]["count"] += 1
        if args.sum:
            num = parse_number(row.get(args.sum, ""))
            if num is None:
                skipped += 1
            else:
                groups[key]["total"] += num
    out_fields = group_cols + ["件数"]
    if args.sum:
        out_fields.append(f"{args.sum}_合計")
    out_rows = []
    for key in order:
        out = dict(zip(group_cols, key))
        out["件数"] = groups[key]["count"]
        if args.sum:
            out[f"{args.sum}_合計"] = format_number(groups[key]["total"])
        out_rows.append(out)
    write_csv(args.output, out_fields, out_rows, args.out_encoding)
    msg = f"summary: {len(rows)} 行 -> {len(out_rows)} グループ -> {args.output}"
    if skipped:
        msg += f"（数値として読めなかった {args.sum} を {skipped} 件スキップ）"
    print(msg)
    return 0


# ---------------------------------------------------------------- to-excel
def cmd_to_excel(args) -> int:
    try:
        from openpyxl import Workbook
    except ImportError:
        print("エラー: to-excel には openpyxl が必要です。\n"
              "インストール: pip install openpyxl\n"
              "または:       pip install csv-excel-toolkit[excel]",
              file=sys.stderr)
        return 2
    fields, rows = read_csv(args.input, args.encoding)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(fields)
    for row in rows:
        ws.append([row.get(name, "") for name in fields])
    Path(args.output).parent.mkdir(parents=True, exist_ok=True) \
        if str(Path(args.output).parent) not in ("", ".") else None
    wb.save(args.output)
    print(f"to-excel: {len(rows)} 行 -> {args.output}")
    return 0


# ---------------------------------------------------------------- parser
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="csv-excel-toolkit",
        description="CSV/Excel の日常処理（結合・重複除去・分割・集計・変換）を行うCLIツール",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    def common(p, input_help="入力CSVファイル"):
        p.add_argument("--encoding", default=None,
                       help="入力の文字コード（省略時は UTF-8 / Shift_JIS を自動判定）")
        p.add_argument("--out-encoding", default="utf-8",
                       help="出力の文字コード（既定: utf-8）")

    p = sub.add_parser("merge", help="複数CSVを結合する")
    p.add_argument("inputs", nargs="+", help="入力CSVファイル（2つ以上）")
    p.add_argument("-o", "--output", required=True, help="出力CSVファイル")
    common(p)
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser("dedupe", help="重複行を除去する")
    p.add_argument("input", help="入力CSVファイル")
    p.add_argument("-o", "--output", required=True, help="出力CSVファイル")
    p.add_argument("--key", default=None,
                   help="重複判定に使う列（カンマ区切り。省略時は全列）")
    common(p)
    p.set_defaults(func=cmd_dedupe)

    p = sub.add_parser("split", help="指定列の値ごとにファイルを分割する")
    p.add_argument("input", help="入力CSVファイル")
    p.add_argument("--column", required=True, help="分割の基準にする列名")
    p.add_argument("--outdir", required=True, help="出力ディレクトリ")
    common(p)
    p.set_defaults(func=cmd_split)

    p = sub.add_parser("summary", help="グループ別に集計する")
    p.add_argument("input", help="入力CSVファイル")
    p.add_argument("--group-by", required=True,
                   help="グループ化する列（カンマ区切りで複数可）")
    p.add_argument("--sum", default=None, help="合計する数値列（任意）")
    p.add_argument("-o", "--output", required=True, help="出力CSVファイル")
    common(p)
    p.set_defaults(func=cmd_summary)

    p = sub.add_parser("to-excel", help="CSVをExcel(.xlsx)に変換する")
    p.add_argument("input", help="入力CSVファイル")
    p.add_argument("-o", "--output", required=True, help="出力 .xlsx ファイル")
    p.add_argument("--encoding", default=None,
                   help="入力の文字コード（省略時は自動判定）")
    p.set_defaults(func=cmd_to_excel)

    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except FileNotFoundError as e:
        print(f"エラー: ファイルが見つかりません: {e.filename}", file=sys.stderr)
        return 2
    except ValueError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
