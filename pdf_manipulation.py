# %%
import copy
import fitz  # PyMuPDF
from fitz import Document
from pathlib import Path
import argparse

# モジュール概要:
# このモジュールは PyMuPDF (fitz) を使った簡単なPDF操作ユーティリティを提供します。
# 主な機能:
#  - 2in1 レイアウト変換 (左右に2ページを並べる)
#  - ページ削除
#  - 複数PDFのマージ
#  - ページ単位での分割、指定ページ数ごとの分割
# CLI も提供しているため、コマンドラインから各機能を実行できます。

# %%
def check_pdf_args(pdf: Document | None = None, pdf_path: str | Path | None = None) -> Document:
    """
    PDFオブジェクトまたはPDFパスから、PDFオブジェクトを取得する
    
    args:
        pdf: PDFオブジェクト
        pdf_path: PDFファイルのパス
    
    return:
        PDFオブジェクト
    
    raises:
        ValueError: pdf と pdf_path の両方が指定されているか、両方が None の場合
    """
    # 引数チェック:
    # - pdf が None で pdf_path が指定されていればファイルを開く
    # - pdf が与えられていて pdf_path が None ならそのまま使用
    # - それ以外（両方 None や両方指定）はエラー
    if (pdf is None) and (pdf_path is not None):
        pdf = fitz.open(pdf_path)  # 元のPDFを開く
    elif (pdf is not None) and (pdf_path is None):
        pass
    else:
        raise ValueError("pdf or pdf_path must be specified")
    return pdf


def merge_pdf_2in1(
    pdf: Document | None = None,
    pdf_path: str | Path | None = None,
    output_path: str | Path = 'out.pdf',
) -> None:
    """
    PDFを2 in 1 レイアウト(2ページを1ページに横並び)に変換する
    pdf, pdf_pathのどちらかは必須

    args:
        pdf (fitz.Document, optional): 元のPDFオブジェクト
        pdf_path (str, optional): 元のPDFファイルのパス
        output_path (str): 変換後のPDFの保存先パス (デフォルト: 'out.pdf')
    
    return:
        None: 変換後のPDFは指定パスに保存される
    """

    pdf = check_pdf_args(pdf, pdf_path)
    # 新しいPDFを作成し、元PDFのページを2ページずつ横に並べる
    # 幅は元ページ幅の2倍、高さは元ページの高さを使用
    # show_pdf_page の Rect は左上(0,0)基準で描画領域を指定する
    new_pdf = fitz.open()  # 新しいPDFを作成
    for i in range(0, len(pdf), 2):  # ページを2つずつ処理
        new_page = new_pdf.new_page(width=pdf[0].rect.width * 2, height=pdf[0].rect.height)  # 新しいページを作成
        if i < len(pdf):  # 左側のページを描画
            page1 = pdf[i]
            new_page.show_pdf_page(fitz.Rect(0, 0, pdf[0].rect.width, pdf[0].rect.height), pdf, i)
        if i + 1 < len(pdf):  # 右側のページを描画
            page2 = pdf[i + 1]
            new_page.show_pdf_page(fitz.Rect(pdf[0].rect.width, 0, pdf[0].rect.width * 2, pdf[0].rect.height), pdf, i + 1)
    # --- 
    new_pdf.save(output_path)  # 新しいPDFを保存
    new_pdf.close()
    print('  Saved: {}'.format(output_path))



def del_pdf_pages(
    pdf: Document | None = None,
    pdf_path: str | Path | None = None,
    del_pages: int | list[int] | None = None,
    output_path: str | Path = 'out.pdf',
) -> None:
    """
    PDFから指定したページを削除する
    
    args:
        pdf (fitz.Document, optional): PDFオブジェクト
        pdf_path (str, optional): PDFファイルのパス
        del_pages (int or list): 削除するページ番号(1ベース)。整数または整数リストで指定
        output_path (str): 出力PDFの保存先パス (デフォルト: 'out.pdf')
    
    return:
        None: 指定パスに保存される
    """
    doc = check_pdf_args(pdf, pdf_path)
    # del_pages は 1 ベースで指定する（ユーザが直感的に扱えるように）
    # list を受け取った場合は降順で削除する（インデックスずれ防止）
    if type(del_pages) == int:
        doc.delete_page(del_pages-1)  # ページを削除
    elif type(del_pages) == list:
        del_pages.sort(reverse=True)
        for p in del_pages:
            doc.delete_page(p-1)
    else:
        print("Error: del_pages must be int or list")
        return

    doc.save(output_path)
    doc.close()
    print('  Saved: {}'.format(output_path))


def merge_pdf(
    pdf: Document | None = None,
    pdf_list: list[str | Path] | None = None,
    output_path: str | Path = 'out.pdf',
) -> None:
    """
    複数のPDFを1つのPDFにマージする
    
    args:
        pdf (fitz.Document, optional): ベースとなるPDFオブジェクト。Noneの場合は新規作成
        pdf_list (list): マージするPDFファイルのパスリスト
        output_path (str): 出力PDFの保存先パス (デフォルト: 'out.pdf')
    
    return:
        None: マージ後のPDFは指定パスに保存される
    """
    # pdf が指定されていればそのコピーをベースにマージ、なければ新規作成
    # pdf_list の順序で順に挿入される
    if pdf is None:
        merged_pdf = fitz.open()
    else:
        merged_pdf = copy(pdf)
    for pdf_file in pdf_list:
        with fitz.open(pdf_file) as doc:
            merged_pdf.insert_pdf(doc)
    # ---
    merged_pdf.save(output_path)
    merged_pdf.close()
    print('  Saved: {}'.format(output_path))


def split_pdf(pdf_path: str | Path, output_dir: str | Path | None = None) -> None:
    """
    PDFを1ページずつ分割し、個別のPDFファイルに保存する
    
    args:
        pdf_path (str): 分割対象のPDFファイルのパス
        output_dir (str, optional): 分割後のPDFファイルの保存先ディレクトリ。Noneの場合は元のPDFと同じディレクトリに保存
    
    return:
        None: 分割されたPDFは個別ファイルとして保存される
    """
    # 出力先ディレクトリの決定:
    # - output_dir が None の場合は入力ファイルと同じディレクトリ
    # 出力ファイル名は "{元名}_p{ページ番号}.pdf"
    fpath = Path(pdf_path)
    fstem = fpath.stem
    if output_dir is None:
        fparent = fpath.parent
    else:
        fparent = Path(output_dir)
        if not fparent.exists():
            fparent.mkdir(parents=True, exist_ok=True)
    # ---
    doc = fitz.open(pdf_path)
    npage = len(doc)
    for page_num in range(npage):
    # 新しいPDFドキュメントを作成
        new_doc = fitz.open()
        
        # ページを抽出して追加
        new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
        
        # ファイル名を指定して保存
        output_filename = Path(fparent) / f"{fstem}_p{page_num + 1}.pdf"
        new_doc.save(output_filename)
        new_doc.close()
        print(f'  {page_num + 1}/{npage}  Saved: {output_filename}')

    doc.close()


def split_pdf_by_pages(pdf_path: str | Path, pages_per_doc: int = 10, output_dir: str | Path | None = None) -> None:
    """
    PDFを指定ページ数ごとに分割し、複数のPDFファイルに保存する
    
    args:
        pdf_path (str): 分割対象のPDFファイルのパス
        pages_per_doc (int): 1つのドキュメントあたりのページ数 (デフォルト: 10)
        output_dir (str, optional): 分割後のPDFファイルの保存先ディレクトリ。Noneの場合は元のPDFと同じディレクトリに保存
    
    return:
        None: 分割されたPDFは複数ファイルとして保存される
    """
    # 指定した pages_per_doc ごとに区切って PDF を作成する。
    # 出力ファイル名は "{元名}_split{連番}.pdf" とする。
    fpath = Path(pdf_path)
    fstem = fpath.stem
    if output_dir is None:
        fparent = fpath.parent
    else:
        fparent = Path(output_dir)
        if not fparent.exists():
            fparent.mkdir(parents=True, exist_ok=True)
    
    doc = fitz.open(pdf_path)
    npage = len(doc)
    doc_num = 1
    
    for start_page in range(0, npage, pages_per_doc):
        end_page = min(start_page + pages_per_doc - 1, npage - 1)
        
        # 新しいPDFドキュメントを作成
        new_doc = fitz.open()
        
        # ページを抽出して追加
        new_doc.insert_pdf(doc, from_page=start_page, to_page=end_page)
        
        # ファイル名を指定して保存
        output_filename = Path(fparent) / f"{fstem}_{doc_num}.pdf"
        new_doc.save(output_filename)
        new_doc.close()
        print(f'  Document {doc_num} (pages {start_page + 1}-{end_page + 1}): Saved: {output_filename}')
        
        doc_num += 1
    
    doc.close()


def _parse_pages_list(s: str):
    """カンマ区切りまたはスペース区切りでページリスト/範囲を解釈する単純ユーティリティ
    例:
      "1,3,5" -> [1,3,5]
      "2-4"   -> [2,3,4]
    戻り値は 1 ベースの整数リスト（del_pdf_pages に渡す前提）。
    """
    # 空文字列は None を返す
    # 文字列中のカンマをスペースに置換して分割、範囲指定を '-' で解釈
    s = s.strip()
    if not s:
        return None
    parts = [p.strip() for p in s.replace(',', ' ').split()]
    nums = []
    for p in parts:
        if '-' in p:
            a, b = p.split('-', 1)
            nums.extend(range(int(a), int(b) + 1))
        else:
            nums.append(int(p))
    return nums


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF manipulation utilities")
    sub = parser.add_subparsers(dest="cmd", required=True)

    # 2in1
    p2 = sub.add_parser("2in1", help="Make 2-in-1 PDF")
    p2.add_argument("-i", "--input", required=True, help="input PDF path")
    p2.add_argument("-o", "--output", required=False, default="out.pdf", help="output PDF path")

    # merge
    pm = sub.add_parser("merge", help="Merge multiple PDFs")
    pm.add_argument("-i", "--input", required=True, nargs="+", help="PDF files to merge (provide one or more paths)")
    pm.add_argument("-o", "--output", required=False, default="out.pdf", help="output PDF path")

    # delete pages
    pd = sub.add_parser("delpages", help="Delete pages from PDF")
    pd.add_argument("-i", "--input", required=True, help="input PDF path")
    pd.add_argument("-p", "--del_pages", required=True,
                    help="pages to delete. single int or comma/space separated list (1-based). Ranges like 2-4 supported.")
    pd.add_argument("-o", "--output", required=False, default="out.pdf", help="output PDF path")

    # split (one page each)
    ps = sub.add_parser("split", help="Split PDF into single-page PDFs")
    ps.add_argument("-i", "--input", required=True, help="input PDF path")
    ps.add_argument("-o", "--output", required=False, help="output directory (default: same as input)")

    # split by pages
    sp = sub.add_parser("split_by_pages", help="Split PDF every N pages")
    sp.add_argument("-i", "--input", required=True, help="input PDF path")
    sp.add_argument("-n", "--npages", type=int, required=False, default=10, help="pages per output PDF")
    sp.add_argument("-o", "--output", required=False, help="output directory (default: same as input)")

    args = parser.parse_args()

    # 各サブコマンドごとに該当関数を呼び出す
    # ここで Path に変換する箇所は入出力パスを統一して扱うため
    if args.cmd == "2in1":
        merge_pdf_2in1(pdf=None, pdf_path=args.input, output_path=Path(args.output))
    elif args.cmd == "merge":
        merge_pdf(pdf=None, pdf_list=args.input, output_path=Path(args.output))
    elif args.cmd == "delpages":
        del_pages_raw = args.del_pages
        parsed = _parse_pages_list(del_pages_raw)
        del_pdf_pages(pdf=None, pdf_path=args.input, del_pages=(parsed if len(parsed) > 1 else parsed[0]), output_path=Path(args.output))
    elif args.cmd == "split":
        split_pdf(pdf_path=args.input, output_dir=(args.output if args.output is None else Path(args.output)))
    elif args.cmd == "split_by_pages":
        split_pdf_by_pages(pdf_path=args.input, pages_per_doc=args.npages, output_dir=(args.output if args.output is None else Path(args.output)))
    else:
        parser.print_help()

