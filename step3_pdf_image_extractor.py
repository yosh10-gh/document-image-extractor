#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ステップ3: .pdfファイルからの画像抽出機能
PDFファイルから埋め込まれた画像を抽出する
"""

from pathlib import Path
from typing import List, Dict, Any, Optional
import sys
import io
from PIL import Image
import fitz  # PyMuPDF

def extract_images_from_pdf(pdf_path: Path) -> List[Dict[str, Any]]:
    """
    .pdfファイルから画像を抽出
    
    Args:
        pdf_path (Path): .pdfファイルのパス
        
    Returns:
        List[Dict[str, Any]]: 抽出した画像情報のリスト
            各辞書には以下のキーが含まれる:
            - 'image': PIL.Image オブジェクト
            - 'format': 画像フォーマット（例: 'JPEG', 'PNG'）
            - 'size': (幅, 高さ) のタプル
            - 'index': PDF内での画像のインデックス
            - 'page': ページ番号
    """
    if not pdf_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {pdf_path}")
    
    try:
        # PDFドキュメントを開く
        pdf_doc = fitz.open(pdf_path)
        images = []
        image_index = 0
        
        # 各ページを処理
        for page_num in range(len(pdf_doc)):
            page = pdf_doc[page_num]
            
            # ページ内の画像を取得
            image_list = page.get_images(full=True)
            
            for img_index, img in enumerate(image_list):
                try:
                    # 画像のXREF（参照番号）を取得
                    xref = img[0]
                    
                    # 画像データを抽出
                    base_image = pdf_doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # PIL Imageオブジェクトに変換
                    image_stream = io.BytesIO(image_bytes)
                    pil_image = Image.open(image_stream)
                    
                    # RGBモードに変換（必要に応じて）
                    if pil_image.mode not in ('RGB', 'RGBA'):
                        if pil_image.mode == 'CMYK':
                            pil_image = pil_image.convert('RGB')
                        elif pil_image.mode in ('P', 'L'):
                            pil_image = pil_image.convert('RGB')
                    
                    # 画像情報を収集
                    image_info = {
                        'image': pil_image.copy(),
                        'format': image_ext.upper(),
                        'size': pil_image.size,
                        'index': image_index,
                        'page': page_num + 1,  # 1ベースのページ番号
                        'xref': xref,
                        'original_mode': base_image.get('colorspace', 'Unknown')
                    }
                    
                    images.append(image_info)
                    image_index += 1
                    
                    # ストリームを閉じる
                    image_stream.close()
                    
                except Exception as e:
                    print(f"警告: ページ {page_num + 1} の画像 {img_index} の処理中にエラー: {e}")
                    continue
        
        # PDFドキュメントを閉じる
        pdf_doc.close()
        
    except Exception as e:
        raise Exception(f"PDF ファイルの処理中にエラーが発生しました: {e}")
    
    return images

def resize_image_to_100px(image: Image.Image) -> Image.Image:
    """
    画像を100x100pxにリサイズ（アスペクト比維持）
    
    Args:
        image (Image.Image): 元の画像
        
    Returns:
        Image.Image: リサイズされた画像
    """
    # アスペクト比を維持して100x100内に収まるようにリサイズ
    image.thumbnail((100, 100), Image.Resampling.LANCZOS)
    
    # 新しい100x100の白背景画像を作成
    resized = Image.new('RGB', (100, 100), (255, 255, 255))
    
    # 元画像を中央に配置
    offset = ((100 - image.size[0]) // 2, (100 - image.size[1]) // 2)
    
    # 透明度がある場合の処理
    if image.mode == 'RGBA':
        resized.paste(image, offset, image)
    else:
        resized.paste(image, offset)
    
    return resized

def save_test_images(images: List[Dict[str, Any]], output_dir: Path, source_file: str) -> None:
    """
    テスト用に画像をファイルに保存
    
    Args:
        images (List[Dict[str, Any]]): 画像情報のリスト
        output_dir (Path): 出力ディレクトリ
        source_file (str): 元ファイル名
    """
    output_dir.mkdir(exist_ok=True)
    
    for i, img_info in enumerate(images):
        try:
            # 元画像を保存
            original_path = output_dir / f"{source_file}_original_{i:02d}.png"
            img_info['image'].save(original_path, 'PNG')
            
            # リサイズ画像を保存
            resized_image = resize_image_to_100px(img_info['image'])
            resized_path = output_dir / f"{source_file}_resized_{i:02d}.png"
            resized_image.save(resized_path, 'PNG')
            
            print(f"保存完了: {original_path.name} → {resized_path.name}")
            
        except Exception as e:
            print(f"画像保存エラー {i}: {e}")

def test_single_pdf(pdf_path: Path) -> None:
    """
    単一の.pdfファイルをテスト
    
    Args:
        pdf_path (Path): テスト対象の.pdfファイル
    """
    print(f"\n=== {pdf_path.name} のテスト ===")
    
    try:
        # 画像抽出
        images = extract_images_from_pdf(pdf_path)
        
        if not images:
            print("画像は見つかりませんでした。")
            return
        
        print(f"抽出された画像数: {len(images)}")
        
        # 画像情報を表示
        for i, img_info in enumerate(images):
            print(f"  画像 {i+1}: {img_info['format']} "
                  f"({img_info['size'][0]}x{img_info['size'][1]}px) "
                  f"[ページ {img_info['page']}]")
        
        # テスト用に画像を保存
        test_output_dir = Path("test_images")
        safe_filename = pdf_path.stem.replace(" ", "_").replace("(", "").replace(")", "")
        save_test_images(images, test_output_dir, safe_filename)
        
    except Exception as e:
        print(f"エラー: {e}")

def main():
    """メイン関数 - テスト実行"""
    print("=== .pdf画像抽出機能テスト ===")
    
    # step1からファイルリストを取得
    from step1_file_crawler import find_files
    
    try:
        files = find_files("target")
        pdf_files = [f for f in files if f.suffix.lower() == '.pdf']
        
        if not pdf_files:
            print("テスト対象の.pdfファイルが見つかりません。")
            return
        
        print(f"テスト対象の.pdfファイル数: {len(pdf_files)}")
        
        # 各.pdfファイルをテスト（最初の2つのみでテスト）
        for pdf_file in pdf_files[:2]:  # テスト時間短縮のため最初の2つのみ
            test_single_pdf(pdf_file)
        
        if len(pdf_files) > 2:
            print(f"\n注意: テスト時間短縮のため、最初の2つのPDFファイルのみテストしました。")
            print(f"残り {len(pdf_files) - 2} ファイルは統合テストで処理されます。")
        
        print(f"\n=== テスト完了 ===")
        print("抽出した画像は 'test_images' ディレクトリに保存されました。")
        
    except Exception as e:
        print(f"テスト実行エラー: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 