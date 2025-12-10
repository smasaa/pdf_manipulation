# PDF Manipulation Utilities

PDF操作を効率化するためのPythonユーティリティです。PyMuPDFライブラリを使用した、簡単で実用的なPDF処理機能を提供します。


## 概要

このプログラムは `pdf_manipulation.py` で、以下のようなPDF操作機能を提供します：

- **2in1レイアウト変換**：2ページを1ページに横並びで配置
- **ページ削除**：指定したページをPDFから削除
- **PDF マージ**：複数のPDFファイルを1つのファイルに統合
- **PDF分割（ページ単位）**：PDFを1ページずつ個別のファイルに分割
- **PDF分割（ページ数指定）**：指定した数ごとにPDFを分割


---

## インストール

### 必要なライブラリ

```bash
pip install pymupdf
```

### 推奨環境

- Python 3.13.11以上
- PyMuPDF 1.26.6以上

---

## 使用方法

### CLIからの実行

#### 1. 2in1レイアウト変換

2ページを1ページに横並びで配置します。

```bash
python pdf_manipulation.py 2in1 -i input.pdf -o output.pdf
```

**パラメータ：**
- `-i, --input`：入力PDFファイルのパス（必須）
- `-o, --output`：出力PDFファイルのパス（デフォルト: out.pdf）

**例：**
```bash
python pdf_manipulation.py 2in1 -i sample.pdf -o sample_2in1.pdf
```

---

#### 2. ページ削除

指定したページをPDFから削除します。

```bash
python pdf_manipulation.py delpages -i input.pdf -p "1,3,5" -o output.pdf
```

**パラメータ：**
- `-i, --input`：入力PDFファイルのパス（必須）
- `-p, --del_pages`：削除するページ番号（1ベース）（必須）
- `-o, --output`：出力PDFファイルのパス（デフォルト: out.pdf）

**ページ指定方法：**
- カンマ区切り：`"1,3,5"` → 1ページ、3ページ、5ページを削除
- スペース区切り：`"1 3 5"`
- 範囲指定：`"2-4"` → 2ページから4ページまで削除
- 複合指定：`"1,3-5,7"` → 1ページ、3～5ページ、7ページを削除

**例：**
```bash
python pdf_manipulation.py delpages -i sample.pdf -p "1-3,5" -o sample_deleted.pdf
```

---

#### 3. PDF マージ

複数のPDFファイルを1つのファイルに統合します。

```bash
python pdf_manipulation.py merge -i file1.pdf file2.pdf file3.pdf -o merged.pdf
```

**パラメータ：**
- `-i, --input`：マージするPDFファイルのパス（複数指定可能、必須）
- `-o, --output`：出力PDFファイルのパス（デフォルト: out.pdf）

**例：**
```bash
python pdf_manipulation.py merge -i chapter1.pdf chapter2.pdf chapter3.pdf -o complete.pdf
```

---

#### 4. PDF分割（ページ単位）

PDFを1ページずつ個別のファイルに分割します。

```bash
python pdf_manipulation.py split -i input.pdf -o output_dir/
```

**パラメータ：**
- `-i, --input`：入力PDFファイルのパス（必須）
- `-o, --output`：出力ディレクトリのパス（デフォルト: 入力ファイルと同じディレクトリ）

**出力ファイル名形式：**
`{元のファイル名}_p{ページ番号}.pdf`

**例：**
```bash
python pdf_manipulation.py split -i sample.pdf -o ./split_pages/
```

結果：
- `sample_p1.pdf`
- `sample_p2.pdf`
- `sample_p3.pdf`
- ...

---

#### 5. PDF分割（ページ数指定）

PDFを指定したページ数ごとに分割します。

```bash
python pdf_manipulation.py split_by_pages -i input.pdf -n 10 -o output_dir/
```

**パラメータ：**
- `-i, --input`：入力PDFファイルのパス（必須）
- `-n, --npages`：1つのドキュメントあたりのページ数（デフォルト: 10）
- `-o, --output`：出力ディレクトリのパス（デフォルト: 入力ファイルと同じディレクトリ）

**出力ファイル名形式：**
`{元のファイル名}_{連番}.pdf`

**例：**
```bash
python pdf_manipulation.py split_by_pages -i sample.pdf -n 5 -o ./split_docs/
```

結果（100ページのPDFの場合）：
- `sample_1.pdf`（1～5ページ）
- `sample_2.pdf`（6～10ページ）
- ...
- `sample_20.pdf`（96～100ページ）

---

### Pythonスクリプトからの使用

各機能は独立した関数として提供されており、Pythonスクリプトから直接呼び出すこともできます。

```python
from pdf_manipulation import merge_pdf_2in1, del_pdf_pages, merge_pdf, split_pdf, split_pdf_by_pages

# 2in1レイアウト変換
merge_pdf_2in1(pdf_path='input.pdf', output_path='output.pdf')

# ページ削除
del_pdf_pages(pdf_path='input.pdf', del_pages=[1, 3, 5], output_path='output.pdf')

# PDF マージ
merge_pdf(pdf_list=['file1.pdf', 'file2.pdf'], output_path='merged.pdf')

# ページ単位の分割
split_pdf(pdf_path='input.pdf', output_dir='./split/')

# ページ数指定の分割
split_pdf_by_pages(pdf_path='input.pdf', pages_per_doc=10, output_dir='./split/')
```

---

## 主要関数の説明

### `merge_pdf_2in1()`

PDFを2in1レイアウト（2ページを1ページに左右並べ）に変換します。

**パラメータ：**
- `pdf` (Document, optional)：PDFオブジェクト
- `pdf_path` (str, optional)：PDFファイルのパス
- `output_path` (str)：出力ファイルのパス（デフォルト: 'out.pdf'）

**戻り値：**
- ファイルを保存して None を返す

---

### `del_pdf_pages()`

指定したページをPDFから削除します。

**パラメータ：**
- `pdf` (Document, optional)：PDFオブジェクト
- `pdf_path` (str, optional)：PDFファイルのパス
- `del_pages` (int or list)：削除するページ番号（1ベース）
- `output_path` (str)：出力ファイルのパス（デフォルト: 'out.pdf'）


---

### `merge_pdf()`

複数のPDFを1つのPDFに統合します。

**パラメータ：**
- `pdf` (Document, optional)：ベースとなるPDFオブジェクト（Noneの場合は新規作成）
- `pdf_list` (list)：統合するPDFファイルのパスリスト
- `output_path` (str)：出力ファイルのパス（デフォルト: 'out.pdf'）

---

### `split_pdf()`

PDFを1ページずつ個別のファイルに分割します。

**パラメータ：**
- `pdf_path` (str)：分割対象のPDFファイルのパス（必須）
- `output_dir` (str, optional)：出力ディレクトリのパス（Noneの場合は入力ファイルと同じディレクトリ）

---

### `split_pdf_by_pages()`

PDFを指定したページ数ごとに分割します。

**パラメータ：**
- `pdf_path` (str)：分割対象のPDFファイルのパス（必須）
- `pages_per_doc` (int)：1つのドキュメントあたりのページ数（デフォルト: 10）
- `output_dir` (str, optional)：出力ディレクトリのパス（Noneの場合は入力ファイルと同じディレクトリ）

---


## 動作確認環境

- Python 3.13.11
- PyMuPDF 1.26.6
- Windows 10/11

---

## ライセンス

このプロジェクトのライセンス情報はプロジェクトのルートディレクトリを参照してください。

---

## 関連リンク

- [PyMuPDF公式ドキュメント](https://pymupdf.readthedocs.io/)
- [PyMuPDFインストール](https://pymupdf.readthedocs.io/en/latest/installation.html)
