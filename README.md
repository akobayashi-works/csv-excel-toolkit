# csv-excel-toolkit

中小企業の事務作業でよくある CSV / Excel 処理をまとめた、Python 製の CLI ツールです。
月次の売上ファイルの結合、重複行の除去、カテゴリ別の集計やファイル分割、Excel への変換といった「Excel で手作業になりがちな処理」を、コマンド一発で再現可能にします。

> **位置づけ**: 本リポジトリはデモ用のサンプルプロジェクトであり、実案件の納品物ではありません。業務自動化の進め方・成果物イメージを示すために公開しています。

[English README](README.en.md)

## 特長

- **標準ライブラリ中心**: 本体は Python 標準ライブラリのみで動作（Excel 変換だけ openpyxl を使用）
- **文字コード自動判定**: UTF-8（BOM 付き含む）と Shift_JIS を自動で見分けて読み込み。出力は既定で UTF-8、`--out-encoding` で変更可能
- **列の違いを吸収**: 結合時に、ファイルごとに列の並びが違ったり、一部ファイルにだけ列があったりしても正しく処理

## インストール

```bash
pip install .
```

Excel 変換（to-excel）も使う場合:

```bash
pip install ".[excel]"
```

開発用（テスト実行）:

```bash
pip install -e ".[dev,excel]"
python -m pytest
```

## 使い方

`examples/` に売上 CSV が3ファイル入っています（1月: UTF-8 / 2月: 列順が違い「担当者」列あり / 3月: Shift_JIS で重複行あり）。以下の実行例はそのまま動作します。

### 1. merge — 複数 CSV の結合

```bash
csv-excel-toolkit merge examples/sales_2026_01.csv examples/sales_2026_02.csv examples/sales_2026_03.csv -o merged.csv
```

```
merge: 3 ファイル / 13 行 （列: 注文ID, 日付, カテゴリ, 商品名, 金額, 担当者） -> merged.csv
```

列の並び違いや、2月分にだけある「担当者」列も自動で揃えられます（無い部分は空欄）。

### 2. dedupe — 重複行の除去

```bash
csv-excel-toolkit dedupe merged.csv -o deduped.csv --key 注文ID
```

```
dedupe: 13 行 -> 12 行 （1 行の重複を除去） -> deduped.csv
```

`--key` を省略すると行全体で重複判定します。

### 3. summary — グループ別集計

```bash
csv-excel-toolkit summary deduped.csv --group-by カテゴリ --sum 金額 -o summary.csv
```

出力（summary.csv）:

```csv
カテゴリ,件数,金額_合計
文具,5,2340
食品,4,5600
日用品,3,3030
```

金額列の「1,200」のようなカンマ区切り数値もそのまま集計できます。

### 4. split — 値ごとにファイル分割

```bash
csv-excel-toolkit split deduped.csv --column カテゴリ --outdir by_category
```

```
split: '文具' -> by_category/文具.csv（5 行）
split: '食品' -> by_category/食品.csv（4 行）
split: '日用品' -> by_category/日用品.csv（3 行）
```

### 5. to-excel — CSV から Excel へ変換

```bash
csv-excel-toolkit to-excel deduped.csv -o deduped.xlsx
```

openpyxl が無い環境では、インストール方法を示すエラーを表示して終了します。

## 機能一覧

| サブコマンド | 内容 | 主なオプション |
|---|---|---|
| `merge` | 複数 CSV の結合（列の並び違い・列の有無を吸収） | `-o` 出力先 |
| `dedupe` | 重複行の除去（先勝ち） | `--key` 判定列（カンマ区切り、省略時は全列） |
| `split` | 指定列の値ごとにファイルを分割 | `--column` 基準列、`--outdir` 出力先 |
| `summary` | グループ別の件数・合計を CSV 出力 | `--group-by` 集計軸、`--sum` 合計する数値列 |
| `to-excel` | CSV → Excel(.xlsx) 変換 | `-o` 出力先（要 openpyxl） |

全サブコマンド共通: `--encoding`（入力の文字コード指定。省略時は自動判定）、`--out-encoding`（出力の文字コード。既定 UTF-8）。

## テスト

```bash
python -m pytest
```

各サブコマンドの正常系・異常系、Shift_JIS / BOM 付き UTF-8 の読み取りなど 11 件のテストを用意しています。

## ライセンス

MIT License（[LICENSE](LICENSE) 参照）

## 作者

Atsushi Kobayashi（小林篤）— ソフトウェアエンジニア。業務自動化・小規模開発・既存システム改修をフリーランスで請け負っています。
