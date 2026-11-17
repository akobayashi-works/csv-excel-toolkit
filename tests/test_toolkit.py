"""csv-excel-toolkit のテスト。各サブコマンドと文字コード判定を検証する。"""

import csv

import pytest

from csv_excel_toolkit.cli import main
from csv_excel_toolkit.io_utils import read_csv


def read_rows(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def write_text(path, text, encoding="utf-8"):
    path.write_bytes(text.encode(encoding))
    return str(path)


# ------------------------------------------------------------------ merge
def test_merge_absorbs_column_order_and_extra_columns(tmp_path):
    a = write_text(tmp_path / "a.csv",
                   "ID,名前,金額\n1,りんご,100\n2,みかん,80\n")
    b = write_text(tmp_path / "b.csv",
                   "金額,ID,名前,担当\n120,3,ぶどう,佐藤\n")
    out = tmp_path / "out.csv"
    assert main(["merge", a, b, "-o", str(out)]) == 0
    rows = read_rows(out)
    assert len(rows) == 3
    # 列の並びが違っても値が正しく対応する
    assert rows[2]["名前"] == "ぶどう" and rows[2]["金額"] == "120"
    # 片方にしかない列は空文字で補完される
    assert rows[0]["担当"] == "" and rows[2]["担当"] == "佐藤"


# ------------------------------------------------------------------ dedupe
def test_dedupe_by_key_keeps_first(tmp_path):
    src = write_text(tmp_path / "in.csv",
                     "ID,名前\n1,りんご\n2,みかん\n1,りんご（重複）\n")
    out = tmp_path / "out.csv"
    assert main(["dedupe", src, "-o", str(out), "--key", "ID"]) == 0
    rows = read_rows(out)
    assert [r["ID"] for r in rows] == ["1", "2"]
    assert rows[0]["名前"] == "りんご"  # 先勝ち


def test_dedupe_full_row(tmp_path):
    src = write_text(tmp_path / "in.csv", "A,B\nx,1\nx,1\nx,2\n")
    out = tmp_path / "out.csv"
    assert main(["dedupe", src, "-o", str(out)]) == 0
    assert len(read_rows(out)) == 2


def test_dedupe_unknown_key_is_error(tmp_path):
    src = write_text(tmp_path / "in.csv", "A,B\nx,1\n")
    out = tmp_path / "out.csv"
    assert main(["dedupe", src, "-o", str(out), "--key", "ZZZ"]) == 2


# ------------------------------------------------------------------ split
def test_split_by_category(tmp_path):
    src = write_text(tmp_path / "in.csv",
                     "カテゴリ,金額\n文具,100\n食品,200\n文具,300\n")
    outdir = tmp_path / "out"
    assert main(["split", src, "--column", "カテゴリ",
                 "--outdir", str(outdir)]) == 0
    assert len(read_rows(outdir / "文具.csv")) == 2
    assert len(read_rows(outdir / "食品.csv")) == 1


# ------------------------------------------------------------------ summary
def test_summary_group_sum_and_count(tmp_path):
    src = write_text(
        tmp_path / "in.csv",
        "カテゴリ,金額\n文具,100\n食品,200\n文具,300\n食品,50\n")
    out = tmp_path / "out.csv"
    assert main(["summary", src, "--group-by", "カテゴリ",
                 "--sum", "金額", "-o", str(out)]) == 0
    rows = {r["カテゴリ"]: r for r in read_rows(out)}
    assert rows["文具"]["件数"] == "2" and rows["文具"]["金額_合計"] == "400"
    assert rows["食品"]["件数"] == "2" and rows["食品"]["金額_合計"] == "250"


def test_summary_handles_comma_numbers(tmp_path):
    src = write_text(tmp_path / "in.csv", '区分,金額\nA,"1,200"\nA,300\n')
    out = tmp_path / "out.csv"
    assert main(["summary", src, "--group-by", "区分",
                 "--sum", "金額", "-o", str(out)]) == 0
    rows = read_rows(out)
    assert rows[0]["金額_合計"] == "1500"


# ------------------------------------------------------------------ encoding
def test_read_shift_jis_autodetect(tmp_path):
    src = write_text(tmp_path / "sjis.csv", "名前,金額\nりんご,100\n",
                     encoding="cp932")
    fields, rows = read_csv(src)
    assert fields == ["名前", "金額"]
    assert rows[0]["名前"] == "りんご"


def test_read_utf8_bom(tmp_path):
    src = tmp_path / "bom.csv"
    src.write_bytes("﻿名前,金額\nみかん,80\n".encode("utf-8"))
    fields, rows = read_csv(src)
    assert fields[0] == "名前"
    assert rows[0]["名前"] == "みかん"


# ------------------------------------------------------------------ to-excel
def test_to_excel(tmp_path):
    pytest.importorskip("openpyxl")
    from openpyxl import load_workbook

    src = write_text(tmp_path / "in.csv", "名前,金額\nりんご,100\n")
    out = tmp_path / "out.xlsx"
    assert main(["to-excel", src, "-o", str(out)]) == 0
    ws = load_workbook(out).active
    values = [[c.value for c in row] for row in ws.iter_rows()]
    assert values == [["名前", "金額"], ["りんご", "100"]]


def test_missing_input_file_is_error(tmp_path):
    out = tmp_path / "out.csv"
    assert main(["dedupe", str(tmp_path / "none.csv"), "-o", str(out)]) == 2
