#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高性能版: 文書画像抽出システム (PyMuPDF版)
targetディレクトリを再帰的にクロールし、.docxと.pdfファイルから全ての画像を抽出してresult.xlsxに出力

【使用状況】
バックオフィス内部使用向け - 最高性能重視

【ライセンス情報】
使用ライブラリ:
- python-docx: Apache 2.0 License
- PyMuPDF: AGPL v3 License (内部使用)
- Pillow: HPND License
- openpyxl: MIT License

※ PyMuPDF (AGPL v3) は内部使用目的のため、外部配布しない限り法的問題なし
"""

from pathlib import Path
from typing import List, Dict, Any, Optional
import sys
import io
import time
from PIL import Image
from docx import Document
import fitz  # PyMuPDF - 最高性能PDF処理ライブラリ
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ===== ファイルクロール機能 =====
def crawl_files(target_dir: Path, extensions: tuple = ('.docx', '.pdf')) -> List[Path]:
    """targetディレクトリを再帰的にクロールし、指定された拡張子のファイルを取得"""
    if not target_dir.exists():
        raise FileNotFoundError(f"ディレクトリが見つかりません: {target_dir}")
    
    files = []
    for extension in extensions:
        # 再帰的に検索
        files.extend(target_dir.rglob(f"*{extension}"))
    
    # パス順でソート
    return sorted(files)

# ===== 画像抽出機能（.docx）=====
def extract_images_from_docx(docx_path: Path) -> List[Dict[str, Any]]:
    """.docxファイルから画像を抽出"""
    if not docx_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {docx_path}")
    
    try:
        images = []
        doc = Document(docx_path)
        
        # 文書内の画像関係を取得
        image_index = 0
        for rel in doc.part.rels:
            relationship = doc.part.rels[rel]
            if "image" in relationship.target_ref:
                try:
                    # 画像データを取得
                    image_data = relationship.target_part.blob
                    
                    # PIL Imageとして読み込み
                    image = Image.open(io.BytesIO(image_data))
                    
                    images.append({
                        'file_path': docx_path,
                        'page_number': 1,  # Wordは単一ページとして扱う
                        'image_index': image_index,
                        'data': image_data,
                        'format': image.format or 'Unknown',
                        'size': image.size,
                        'mode': image.mode
                    })
                    
                    image_index += 1
                    print(f"    画像 {image_index}: {image.format} {image.size} {image.mode}")
                    
                except Exception as e:
                    print(f"    警告: 画像の読み込みに失敗 - {e}")
                    continue
        
        return images
        
    except Exception as e:
        print(f"エラー: .docxファイルの処理に失敗 - {e}")
        return []

# ===== 画像抽出機能（.pdf）- PyMuPDF版 =====
def extract_images_from_pdf(pdf_path: Path) -> List[Dict[str, Any]]:
    """.pdfファイルから画像を抽出 (PyMuPDF使用 - 最高性能)"""
    if not pdf_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {pdf_path}")
    
    try:
        images = []
        image_index = 0
        
        # PyMuPDFでPDFを開く
        pdf_doc = fitz.open(pdf_path)
        
        # 全ページをループして画像を抽出
        for page_num in range(len(pdf_doc)):
            page = pdf_doc[page_num]
            print(f"    ページ {page_num + 1}/{len(pdf_doc)} を処理中...")
            
            # ページ内の画像リストを取得
            image_list = page.get_images(full=True)
            
            for img_index, img in enumerate(image_list):
                try:
                    # 画像参照情報を取得
                    xref = img[0]  # 画像のxref番号
                    
                    # 画像データを抽出
                    base_image = pdf_doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # PIL Imageとして確認
                    pil_image = Image.open(io.BytesIO(image_bytes))
                    
                    images.append({
                        'file_path': pdf_path,
                        'page_number': page_num + 1,
                        'image_index': image_index,
                        'data': image_bytes,
                        'format': image_ext.upper(),
                        'size': pil_image.size,
                        'mode': pil_image.mode
                    })
                    
                    image_index += 1
                    print(f"      画像 {image_index}: {image_ext.upper()} {pil_image.size} {pil_image.mode}")
                    
                except Exception as e:
                    print(f"      警告: 画像 {img_index} の抽出に失敗 - {e}")
                    continue
        
        pdf_doc.close()
        return images
        
    except Exception as e:
        print(f"エラー: .pdfファイルの処理に失敗 - {e}")
        return []

# ===== 画像リサイズ機能 =====
def resize_image_for_excel(image_bytes: bytes, target_width: int = 100, target_height: int = 100) -> Optional[io.BytesIO]:
    """画像をExcel用にリサイズ（バイト→バイト）"""
    try:
        # バイトデータからPIL画像を作成
        image_buffer = io.BytesIO(image_bytes)
        with Image.open(image_buffer) as img:
            # RGBAまたはRGB形式に変換
            if img.mode not in ('RGB', 'RGBA'):
                img = img.convert('RGB')
            
            # アスペクト比を保持してリサイズ
            img.thumbnail((target_width, target_height), Image.Resampling.LANCZOS)
            
            # 透明な背景で中央に配置（100x100pxの画像を作成）
            new_img = Image.new('RGB', (target_width, target_height), (255, 255, 255))  # 白背景
            
            # 中央に配置
            x = (target_width - img.width) // 2
            y = (target_height - img.height) // 2
            new_img.paste(img, (x, y))
            
            # バイトストリームに保存
            output_buffer = io.BytesIO()
            new_img.save(output_buffer, format='PNG')
            output_buffer.seek(0)
            return output_buffer
            
    except Exception as e:
        print(f"画像リサイズエラー: {e}")
        return None

# ===== Excel出力機能 =====
def export_to_excel(file_list: List[Path], all_images: List[Dict], output_path: Path):
    """ファイルリストと画像をExcelに出力"""
    wb = Workbook()
    ws = wb.active
    
    # ヘッダー設定（A列のみ）
    ws['A1'] = 'ファイルパス'
    ws['A1'].font = Font(bold=True)
    
    # セルのサイズを100x100pxに設定
    cell_size_px = 100
    
    # 行の高さを設定 (ピクセルを ポイント に変換: 1pt ≈ 1.33px)
    row_height_pt = cell_size_px / 1.33
    
    # ファイルリストを処理
    for row_idx, file_path in enumerate(file_list, start=2):  # 2行目から開始
        # A列にファイルの絶対パス
        ws[f'A{row_idx}'] = str(file_path.absolute())
        
        # 行の高さを設定
        ws.row_dimensions[row_idx].height = row_height_pt
        
        # そのファイルに対応する画像を取得
        file_images = [img for img in all_images if img.get('file_path') == file_path]
        
        # 画像を水平方向に配置
        for img_idx, image_data in enumerate(file_images):
            col_idx = img_idx + 2  # B列から開始（A列はファイルパス）
            col_letter = get_column_letter(col_idx)
            
            # 列幅を設定 (ピクセルをExcel単位に変換)
            ws.column_dimensions[col_letter].width = cell_size_px / 7  # 約14.3
            
            # 画像をリサイズしてExcelに挿入
            resized_image_buffer = resize_image_for_excel(image_data['data'])
            if resized_image_buffer:
                try:
                    excel_image = ExcelImage(resized_image_buffer)
                    excel_image.width = cell_size_px
                    excel_image.height = cell_size_px
                    
                    # セルに画像を配置
                    ws.add_image(excel_image, f'{col_letter}{row_idx}')
                    
                except Exception as e:
                    print(f"Excel画像挿入エラー: {e}")
    
    # Excelファイルを保存
    wb.save(output_path)
    
    # ファイルサイズを取得
    file_size = output_path.stat().st_size / 1024  # KB
    print(f"✅ Excel出力完了: {output_path} ({file_size:.1f} KB)")

# ===== メイン処理 =====
def main():
    """メイン処理関数"""
    print("🔍 文書画像抽出システム (PyMuPDF高性能版)")
    print("=" * 50)
    
    target_dir = Path("target")
    
    if not target_dir.exists():
        print(f"❌ '{target_dir}' ディレクトリが見つかりません")
        return
    
    try:
        start_time = time.time()
        
        # ステップ1: ファイルクロール
        print("📂 ファイルクロール中...")
        files = crawl_files(target_dir)
        
        if not files:
            print("⚠️  対象ファイルが見つかりませんでした")
            return
        
        print(f"📊 見つかったファイル: {len(files)}個")
        for file_path in files:
            print(f"  - {file_path}")
        
        # ステップ2: 画像抽出
        print()
        print("🖼️  画像抽出中...")
        all_images = []
        
        for file_path in files:
            print(f"📄 処理中: {file_path.name}")
            
            if file_path.suffix.lower() == '.docx':
                images = extract_images_from_docx(file_path)
            elif file_path.suffix.lower() == '.pdf':
                images = extract_images_from_pdf(file_path)
            else:
                print(f"  ⚠️ 未対応の形式: {file_path.suffix}")
                continue
            
            all_images.extend(images)
            print(f"  📊 抽出数: {len(images)}枚")
        
        # ステップ3: Excel出力
        print()
        print("📊 Excel出力中...")
        output_path = Path("result.xlsx")
        export_to_excel(files, all_images, output_path)
        
        # 結果表示
        end_time = time.time()
        processing_time = end_time - start_time
        
        print()
        print("🎉 処理完了！")
        print(f"📈 処理結果:")
        print(f"  - 処理ファイル数: {len(files)}個")
        print(f"  - 抽出画像総数: {len(all_images)}枚")
        print(f"  - 処理時間: {processing_time:.2f}秒")
        print(f"  - 出力ファイル: {output_path}")
        
    except Exception as e:
        print(f"❌ エラーが発生しました: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 