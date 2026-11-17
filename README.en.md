# csv-excel-toolkit

A Python CLI toolkit for the everyday CSV / Excel chores common in small-business office work: merging monthly sales files, removing duplicate rows, summarizing by category, splitting files by a column value, and converting CSV to Excel — all as reproducible one-line commands instead of manual spreadsheet work.

> **Note**: This repository is a demo sample project, not a deliverable from a client engagement. It is published to illustrate how business automation work is approached and what the output looks like.

[日本語 README](README.md)

## Highlights

- **Standard library first**: the core runs on the Python standard library only (openpyxl is used solely for Excel export)
- **Automatic encoding detection**: reads UTF-8 (with or without BOM) and Shift_JIS transparently; output defaults to UTF-8 and can be changed with `--out-encoding`
- **Tolerant merging**: files with different column orders, or columns present in only some files, are merged correctly

## Installation

```bash
pip install .
```

With Excel export support:

```bash
pip install ".[excel]"
```

For development (running the tests):

```bash
pip install -e ".[dev,excel]"
python -m pytest
```

## Usage

The `examples/` directory contains three monthly sales CSV files (January: UTF-8 / February: different column order plus an extra column / March: Shift_JIS with a duplicate row). The examples below run as-is.

### 1. merge — combine multiple CSV files

```bash
csv-excel-toolkit merge examples/sales_2026_01.csv examples/sales_2026_02.csv examples/sales_2026_03.csv -o merged.csv
```

```
merge: 3 files / 13 rows (columns: order id, date, category, product, amount, staff) -> merged.csv
```

### 2. dedupe — remove duplicate rows

```bash
csv-excel-toolkit dedupe merged.csv -o deduped.csv --key 注文ID
```

```
dedupe: 13 rows -> 12 rows (1 duplicate removed) -> deduped.csv
```

Without `--key`, rows are compared across all columns.

### 3. summary — group-by aggregation

```bash
csv-excel-toolkit summary deduped.csv --group-by カテゴリ --sum 金額 -o summary.csv
```

Output (summary.csv):

```csv
カテゴリ,件数,金額_合計
文具,5,2340
食品,4,5600
日用品,3,3030
```

Thousands separators such as `1,200` are handled automatically.

### 4. split — split a file by column values

```bash
csv-excel-toolkit split deduped.csv --column カテゴリ --outdir by_category
```

Creates one CSV per category (e.g. `by_category/文具.csv`).

### 5. to-excel — convert CSV to Excel (.xlsx)

```bash
csv-excel-toolkit to-excel deduped.csv -o deduped.xlsx
```

If openpyxl is not installed, the command exits with a clear message explaining how to install it.

## Commands

| Subcommand | Description | Key options |
|---|---|---|
| `merge` | Merge multiple CSVs, absorbing column order/presence differences | `-o` output |
| `dedupe` | Remove duplicate rows (first occurrence wins) | `--key` columns (comma-separated; default: all) |
| `split` | Split a file into one file per value of a column | `--column`, `--outdir` |
| `summary` | Per-group row counts and sums as CSV | `--group-by`, `--sum` |
| `to-excel` | Convert CSV to Excel (.xlsx) | `-o` output (requires openpyxl) |

Shared options: `--encoding` (input encoding; auto-detected when omitted), `--out-encoding` (output encoding; default UTF-8).

## Tests

```bash
python -m pytest
```

11 tests cover each subcommand (happy and error paths) plus Shift_JIS / UTF-8 BOM reading.

## License

MIT License — see [LICENSE](LICENSE).

## Author

Atsushi Kobayashi — software engineer working freelance on business automation, small-scale development, and maintenance of existing systems.
