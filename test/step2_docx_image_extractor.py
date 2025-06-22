#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ステップ2: .docxファイルからの画像抽出機能
Word文書から埋め込まれた画像を抽出する
"""

from pathlib import Path
from typing import List, Dict, Any, Optional
import sys
import io
from PIL import Image
from docx import Document
from docx.document import Document as DocumentType

def extract_images_from_docx(docx_path: Path) -> List[Dict[str, Any]]:
    """
    .docxファイルから画像を抽出
    
    Args:
        docx_path (Path): .docxファイルのパス
        
    Returns:
        List[Dict[str, Any]]: 抽出した画像情報のリスト
            各辞書には以下のキーが含まれる:
            - 'image': PIL.Image オブジェクト
            - 'format': 画像フォーマット（例: 'JPEG', 'PNG'）
            - 'size': (幅, 高さ) のタプル
            - 'index': 文書内での画像のインデックス
    """
    if not docx_path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {docx_path}")
    
    try:
        # Word文書を開く
        doc = Document(docx_path)
        images = []
        
        # 文書内のリレーション（関連ファイル）を取得
        rels = doc.part.rels
        
        image_index = 0
        for rel_id, rel in rels.items():
            # 画像ファイルかどうかチェック
            if "image" in rel.target_part.content_type:
                try:
                    # 画像データを取得
                    image_data = rel.target_part.blob
                    
                    # PIL Imageオブジェクトに変換
                    image_stream = io.BytesIO(image_data)
                    pil_image = Image.open(image_stream)
                    
                    # 画像情報を収集
                    image_info = {
                        'image': pil_image.copy(),  # コピーして安全に保持
                        'format': pil_image.format or 'UNKNOWN',
                        'size': pil_image.size,
                        'index': image_index,
                        'rel_id': rel_id,
                        'content_type': rel.target_part.content_type
                    }
                    
                    images.append(image_info)
                    image_index += 1
                    
                    # ストリームを閉じる
                    image_stream.close()
                    
                except Exception as e:
                    print(f"警告: 画像 {rel_id} の処理中にエラー: {e}")
                    continue
                    
    except Exception as e:
        raise Exception(f"DOCX ファイルの処理中にエラーが発生しました: {e}")
    
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

def test_single_docx(docx_path: Path) -> None:
    """
    単一の.docxファイルをテスト
    
    Args:
        docx_path (Path): テスト対象の.docxファイル
    """
    print(f"\n=== {docx_path.name} のテスト ===")
    
    try:
        # 画像抽出
        images = extract_images_from_docx(docx_path)
        
        if not images:
            print("画像は見つかりませんでした。")
            return
        
        print(f"抽出された画像数: {len(images)}")
        
        # 画像情報を表示
        for i, img_info in enumerate(images):
            print(f"  画像 {i+1}: {img_info['format']} "
                  f"({img_info['size'][0]}x{img_info['size'][1]}px)")
        
        # テスト用に画像を保存
        test_output_dir = Path("test_images")
        safe_filename = docx_path.stem.replace(" ", "_").replace("(", "").replace(")", "")
        save_test_images(images, test_output_dir, safe_filename)
        
    except Exception as e:
        print(f"エラー: {e}")

def main():
    """メイン関数 - テスト実行"""
    print("=== .docx画像抽出機能テスト ===")
    
    # step1からファイルリストを取得
    from step1_file_crawler import find_files
    
    try:
        files = find_files("target")
        docx_files = [f for f in files if f.suffix.lower() == '.docx']
        
        if not docx_files:
            print("テスト対象の.docxファイルが見つかりません。")
            return
        
        print(f"テスト対象の.docxファイル数: {len(docx_files)}")
        
        # 各.docxファイルをテスト
        for docx_file in docx_files:
            test_single_docx(docx_file)
        
        print(f"\n=== テスト完了 ===")
        print("抽出した画像は 'test_images' ディレクトリに保存されました。")
        
    except Exception as e:
        print(f"テスト実行エラー: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 