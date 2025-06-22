# 文書画像抽出システム - 完全マニュアル

このシステムは、Wordファイル（.docx）とPDFファイル（.pdf）から画像を自動抽出し、Excelファイルに一覧表示するPythonプログラムです。

## 📋 目次
1. [システム概要](#システム概要)
2. [動作環境](#動作環境)
3. [初回セットアップ（新しいPCの場合）](#初回セットアップ新しいpcの場合)
4. [既存環境からの移行](#既存環境からの移行)
5. [プログラムの使用方法](#プログラムの使用方法)
6. [出力結果の説明](#出力結果の説明)
7. [トラブルシューティング](#トラブルシューティング)
8. [技術仕様・ライセンス](#技術仕様ライセンス)

## 🎯 システム概要

### 機能
- `target`フォルダ内のすべてのフォルダを再帰的に検索
- .docxファイルと.pdfファイルから画像を自動抽出
- 抽出した画像を100×100pxにリサイズ
- Excelファイル（`result.xlsx`）に一覧表示
  - A列：ファイルの絶対パス
  - B列以降：各ファイルから抽出した全ての画像

### 対応ファイル形式
- **入力**：.docx、.pdf
- **画像形式**：PNG、JPEG、BMP、WMF、EMF など
- **出力**：Excel形式（.xlsx）

### 性能
- **高速処理**：0.4秒で195枚の画像を抽出
- **高精度**：埋め込み画像を直接抽出（ページレンダリングではない）
- **大容量対応**：数百ファイル・数千画像の一括処理が可能

## 💻 動作環境

### 必要なソフトウェア
- **Windows 10/11**
- **Python 3.12.10** （pyenv-winで管理）
- **PowerShell** （Windows標準搭載）

### 必要なPythonライブラリ
```
python-docx==1.1.2    # Word文書処理 (Apache 2.0)
PyMuPDF==1.24.12      # PDF処理 (AGPL v3 - 内部使用)
Pillow==11.0.0        # 画像処理 (HPND)
openpyxl==3.1.5       # Excel処理 (MIT)
```

**注意**：PyMuPDF (AGPL v3) は**内部使用**に限定されます。外部配布には注意が必要です。

## 🚀 初回セットアップ（新しいPCの場合）

### ステップ1: pyenv-winのインストール

1. **PowerShellを管理者権限で開く**
   ```
   スタートメニュー → 「PowerShell」で検索 → 右クリック → 「管理者として実行」
   ```

2. **Gitが利用可能か確認**
   ```powershell
   git --version
   ```
   エラーが出る場合は[Git for Windows](https://git-scm.com/download/win)をインストール

3. **pyenv-winをダウンロード**
   ```powershell
   git clone https://github.com/pyenv-win/pyenv-win.git %USERPROFILE%\.pyenv
   ```

4. **環境変数を設定**
   ```powershell
   [System.Environment]::SetEnvironmentVariable('PYENV',$env:USERPROFILE + "\.pyenv\pyenv-win\","User")
   [System.Environment]::SetEnvironmentVariable('PYENV_ROOT',$env:USERPROFILE + "\.pyenv\pyenv-win\","User")
   [System.Environment]::SetEnvironmentVariable('PYENV_HOME',$env:USERPROFILE + "\.pyenv\pyenv-win\","User")
   [System.Environment]::SetEnvironmentVariable('path', $env:USERPROFILE + "\.pyenv\pyenv-win\bin;" + $env:USERPROFILE + "\.pyenv\pyenv-win\shims;" + $env:path,"User")
   ```

5. **PowerShellを再起動して確認**
   ```powershell
   # 新しいPowerShellウィンドウで実行
   pyenv --version
   ```
   ✅ バージョン番号が表示されればOK

### ステップ2: Python 3.12.10のインストール

1. **Python 3.12.10をインストール**
   ```powershell
   pyenv install 3.12.10
   ```
   ⚠️ **注意：この処理には5-10分かかります**

2. **インストール確認**
   ```powershell
   pyenv versions
   ```
   ✅ `3.12.10`が表示されればOK

### ステップ3: プロジェクトセットアップ

1. **プロジェクトフォルダを作成・移動**
   ```powershell
   # 例：デスクトップにプロジェクトフォルダを作成
   mkdir "C:\Users\$env:USERNAME\Desktop\view_word_pdf_img"
   cd "C:\Users\$env:USERNAME\Desktop\view_word_pdf_img"
   ```

2. **このプロジェクト用のPythonバージョンを設定**
   ```powershell
   pyenv local 3.12.10
   ```

3. **仮想環境を作成・有効化**
   ```powershell
   pyenv exec python -m venv venv
   venv\Scripts\activate
   ```
   ✅ プロンプトに`(venv)`が表示されればOK

4. **必要なライブラリをインストール**
   ```powershell
   pip install python-docx==1.1.2 PyMuPDF==1.24.12 Pillow==11.0.0 openpyxl==3.1.5
   ```

5. **requirements.txtを作成**
   ```powershell
   pip freeze > requirements.txt
   ```

## 📦 既存環境からの移行

### 移行元PCでの作業

1. **プロジェクトフォルダ全体をコピー**
   - 📁 プロジェクトフォルダ全体をUSBやクラウドドライブにコピー
   - ⚠️ `venv`フォルダは除外してください（新しいPCで再作成）

### 移行先PCでの作業

1. **ステップ1〜2を完了**（上記参照）

2. **プロジェクトファイルを配置**
   ```powershell
   # コピーしたファイルを配置後、プロジェクトフォルダに移動
   cd "パス\to\your\project"
   ```

3. **Python環境をセットアップ**
   ```powershell
   pyenv local 3.12.10
   pyenv exec python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

## 🎮 プログラムの使用方法

### ファイル構成
```
プロジェクトフォルダ/
├── main.py                 # メインプログラム
├── requirements.txt        # 必要なライブラリ一覧
├── target/                 # 処理対象ファイルを配置するフォルダ
│   ├── sample.pdf
│   ├── document.docx
│   └── subfolder/          # サブフォルダも自動で検索
│       └── more_files.pdf
└── venv/                   # 仮想環境（自動作成）
```

### 実行手順

1. **仮想環境を有効化**
   ```powershell
   cd "プロジェクトフォルダのパス"
   venv\Scripts\activate
   ```

2. **処理したいファイルを配置**
   - `target`フォルダに.docxや.pdfファイルを配置
   - サブフォルダ内のファイルも自動的に処理されます

3. **プログラムを実行**
   ```powershell
   python main.py
   ```

4. **実行例**
   ```
   🔍 文書画像抽出システム (PyMuPDF高性能版)
   ==================================================
   📂 ファイルクロール中...
   📊 見つかったファイル: 9個
     - target\samples\A\A_sample_B (1).docx
     - target\samples\A\A_sample_B (1).pdf
     - target\samples\A\A_sample_B (2).pdf
     - target\samples\sample.pdf
     - target\samples\sample2020.docx
     - target\samples\sample_B.pdf

   🖼️  画像抽出中...
   📄 処理中: A_sample_B (1).docx
       画像 1: PNG (400, 398) RGBA
       画像 2: PNG (360, 400) RGBA
     📊 抽出数: 2枚
   📄 処理中: A_sample_B (1).pdf
       ページ 1/5 を処理中...
         画像 1: JPEG (347, 213) RGB
         画像 2: JPEG (417, 189) RGB
       ページ 2/5 を処理中...
       ページ 3/5 を処理中...
         画像 3: JPEG (286, 301) RGB
         画像 4: JPEG (297, 257) RGB
       ページ 4/5 を処理中...
         画像 5: JPEG (325, 314) RGB
       ページ 5/5 を処理中...
         画像 6: JPEG (951, 269) RGB
         画像 7: JPEG (281, 270) RGB
         画像 8: JPEG (228, 231) RGB
         画像 9: JPEG (298, 286) RGB
         画像 10: JPEG (300, 290) RGB
         画像 11: JPEG (256, 235) RGB
         画像 12: JPEG (319, 191) RGB
     📊 抽出数: 12枚

   📊 Excel出力中...
   ✅ Excel出力完了: result.xlsx (793.6 KB)

   🎉 処理完了！
   📈 処理結果:
     - 処理ファイル数: 9個
     - 抽出画像総数: 195枚
     - 処理時間: 0.44秒
     - 出力ファイル: result.xlsx
   ```

5. **結果を確認**
   - `result.xlsx`ファイルを開く
   - A列：ファイルパス
   - B列以降：抽出された画像（100×100px）

## 📊 出力結果の説明

### Excelファイルの構成
| 列 | 内容 | 説明 |
|---|---|---|
| A | ファイルパス | 処理したファイルの絶対パス |
| B以降 | 画像 | 各ファイルから抽出した画像（100×100px） |

### プログラムの動作
- **画像表示**: 制限なし（各ファイルの全画像を表示）
- **出力ファイル名**: `result.xlsx`

### 画像処理の詳細
- **サイズ**: 自動的に100×100pxにリサイズ
- **アスペクト比**: 維持（白背景で中央配置）
- **形式**: PNG形式で統一
- **エラー処理**: 読み込めない画像はスキップ（警告表示）

## 🛠️ トラブルシューティング

### Q1: 「pyenvコマンドが認識されない」
```powershell
# 解決方法1: 環境変数を再設定
[System.Environment]::SetEnvironmentVariable('path', $env:USERPROFILE + "\.pyenv\pyenv-win\bin;" + $env:USERPROFILE + "\.pyenv\pyenv-win\shims;" + $env:path,"User")

# 解決方法2: PowerShellを管理者権限で再起動
```

### Q2: 「Permission denied: result.xlsx」
```
原因：Excelファイルが開かれている
解決方法：
1. result.xlsxを閉じる
2. または、プログラムを再実行（別ファイル名で保存される）
```

### Q3: 「仮想環境を有効化できない」
```powershell
# PowerShellの実行ポリシーを変更（管理者権限）
Set-ExecutionPolicy RemoteSigned
```

### Q4: 「画像抽出でエラーが多発する」
```
原因：ファイルが破損している、または対応していない形式
対処法：
1. ファイルを他のソフトで開けるか確認
2. 別のファイルで動作テストする
3. ファイル形式を確認（.docx, .pdfのみ対応）
```

### Q5: 「メモリ不足エラー」
```
原因：大量の画像や巨大なファイルを処理
解決方法：
1. targetフォルダ内のファイル数を減らす
2. PCを再起動してメモリを解放
3. 大きなPDFファイルを分割する
```

### Q6: 「日本語ファイル名でエラー」
```
原因：文字エンコーディングの問題
解決方法：
1. ファイル名を英数字に変更
2. フォルダパスに日本語が含まれていないか確認
```

## 🔧 技術仕様・ライセンス

### 使用ライブラリと役割
| ライブラリ | 機能 | ライセンス | 商業利用 |
|---|---|---|---|
| `python-docx` | Word文書の画像抽出 | Apache 2.0 | ✅ 可能 |
| `PyMuPDF` | PDF文書の画像抽出 | AGPL v3 | ⚠️ 内部使用のみ |
| `Pillow` | 画像処理・リサイズ | HPND | ✅ 可能 |
| `openpyxl` | Excel読み書き | MIT | ✅ 可能 |

### ライセンスに関する重要な注意
- **PyMuPDF (AGPL v3)**：内部使用・バックオフィス作業では問題なし
- **外部配布時**：ソースコード公開義務が発生する可能性があります
- **商用製品組み込み**：法的確認が必要です

### システム制限
- **最大ファイルサイズ**: 制限なし（メモリ依存）
- **対応画像形式**: PNG、JPEG、BMP、WMF、EMF等
- **最大画像数**: 制限なし（Excel列数上限まで）
- **処理速度**: 約440画像/秒（環境依存）

### ファイル構造の詳細
```
main.py の構成:
├── crawl_files()              # ファイル検索機能
├── extract_images_from_docx() # Word画像抽出
├── extract_images_from_pdf()  # PDF画像抽出
├── resize_image_for_excel()   # 画像リサイズ
├── export_to_excel()          # Excel出力
└── main()                     # メイン処理
```

## 📞 サポート

### よくある質問
1. **Q: MacやLinuxでも動きますか？**
   A: Windows専用です。他のOSでは別途調整が必要です。

2. **Q: 商用利用は可能ですか？**
   A: 内部使用に限定すれば問題ありません。外部配布にはライセンス確認が必要です。

3. **Q: 大量のファイルを処理できますか？**
   A: メモリ使用量に注意し、必要に応じてファイルを分割してください。

### エラーログの確認
- プログラム実行時のエラーメッセージを確認
- 警告メッセージは処理継続可能（一部画像が読み込めない場合）
- 致命的エラーは処理停止（環境設定や権限の問題）

---

**🎉 セットアップ完了後は、targetフォルダにファイルを配置して`python main.py`を実行するだけで使用できます！** 