"""CSV の読み書きと文字コード判定を行うユーティリティ.

読み取り時は UTF-8（BOM 付き含む）を優先し、失敗したら Shift_JIS
（cp932）として解釈する。--encoding で明示指定も可能。
"""

from __future__ import annotations

import csv
import io
from pathlib import Path

# 判定に使う候補（先に一致したものを採用）
_CANDIDATES = ("utf-8-sig", "cp932")


def decode_bytes(data: bytes, encoding: str | None = None) -> str:
    """バイト列をテキストに変換する。encoding が None なら自動判定。"""
    if encoding:
        return data.decode(encoding)
    for enc in _CANDIDATES:
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    # 最後の手段: 置換文字で読み進める
    return data.decode("utf-8", errors="replace")


def read_csv(path: str | Path, encoding: str | None = None):
    """CSV を読み、(fieldnames, rows) を返す。rows は dict のリスト。"""
    data = Path(path).read_bytes()
    text = decode_bytes(data, encoding)
    reader = csv.DictReader(io.StringIO(text))
    if reader.fieldnames is None:
        raise ValueError(f"{path}: ヘッダー行がありません")
    fieldnames = [name.strip() for name in reader.fieldnames]
    rows = []
    for raw in reader:
        row = {}
        for key, value in raw.items():
            if key is None:
                continue
            row[key.strip()] = value if value is not None else ""
        rows.append(row)
    return fieldnames, rows


def write_csv(path: str | Path, fieldnames, rows, encoding: str = "utf-8") -> None:
    """rows（dict のリスト）を CSV として書き出す。"""
    path = Path(path)
    if path.parent and str(path.parent) not in ("", "."):
        path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding=encoding, newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({name: row.get(name, "") for name in fieldnames})


def parse_number(value: str):
    """数値文字列を int / float に変換する。変換できなければ None。"""
    if value is None:
        return None
    text = str(value).strip().replace(",", "").replace("¥", "").replace("円", "")
    if text == "":
        return None
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        return None


def format_number(value) -> str:
    """集計結果を表示用文字列にする（整数は整数のまま）。"""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)
