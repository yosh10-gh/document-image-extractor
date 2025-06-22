#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ステップ1: ファイルクロール機能
targetディレクトリを再帰的に検索し、.docxと.pdfファイルを収集する
"""

from pathlib import Path
from typing import List, Tuple
import sys

def find_files(target_dir: str, extensions: Tuple[str, ...] = ('.docx', '.pdf')) -> List[Path]:
    """
    指定ディレクトリを再帰的に検索し、指定拡張子のファイルを収集
    
    Args:
        target_dir (str): 検索対象のディレクトリパス
        extensions (Tuple[str, ...]): 検索する拡張子のタプル
        
    Returns:
        List[Path]: 見つかったファイルのPathオブジェクトのリスト
    """
    target_path = Path(target_dir)
    
    # ディレクトリが存在するかチェック
    if not target_path.exists():
        raise FileNotFoundError(f"指定されたディレクトリが見つかりません: {target_dir}")
    
    if not target_path.is_dir():
        raise NotADirectoryError(f"指定されたパスはディレクトリではありません: {target_dir}")
    
    found_files = []
    
    # 各拡張子について再帰的に検索
    for extension in extensions:
        # **でサブディレクトリも含めて検索
        pattern = f"**/*{extension}"
        files = list(target_path.rglob(pattern))
        found_files.extend(files)
    
    # 重複を除去し、パスでソート
    found_files = sorted(list(set(found_files)))
    
    return found_files

def display_files(files: List[Path]) -> None:
    """
    見つかったファイルの一覧を表示
    
    Args:
        files (List[Path]): ファイルパスのリスト
    """
    if not files:
        print("対象のファイルは見つかりませんでした。")
        return
    
    print(f"見つかったファイル数: {len(files)}")
    print("-" * 50)
    
    for i, file_path in enumerate(files, 1):
        # ファイルサイズを取得
        try:
            file_size = file_path.stat().st_size
            size_mb = file_size / (1024 * 1024)  # MB単位
            print(f"{i:3d}. {file_path}")
            print(f"     サイズ: {size_mb:.2f} MB ({file_size:,} bytes)")
            print()
        except Exception as e:
            print(f"{i:3d}. {file_path}")
            print(f"     エラー: ファイル情報取得失敗 - {e}")
            print()

def main():
    """メイン関数 - テスト実行"""
    target_directory = "target"
    
    try:
        print("=== ファイルクロール機能テスト ===")
        print(f"検索対象ディレクトリ: {target_directory}")
        print(f"検索対象拡張子: .docx, .pdf")
        print()
        
        # ファイル検索実行
        files = find_files(target_directory)
        
        # 結果表示
        display_files(files)
        
        # 統計情報
        docx_files = [f for f in files if f.suffix.lower() == '.docx']
        pdf_files = [f for f in files if f.suffix.lower() == '.pdf']
        
        print("=== 統計情報 ===")
        print(f"総ファイル数: {len(files)}")
        print(f".docx ファイル: {len(docx_files)}")
        print(f".pdf ファイル: {len(pdf_files)}")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 