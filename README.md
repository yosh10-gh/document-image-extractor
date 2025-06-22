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
8. [技術仕様](#技術仕様)

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

## 💻 動作環境

### 必要なソフトウェア
- **Windows 10/11**
- **Python 3.12.10** （pyenv-winで管理）
- **PowerShell** （Windows標準搭載）

### 必要なPythonライブラリ
```
python-docx==1.1.2    # Word文書処理
PyMuPDF==1.24.12      # PDF処理
Pillow==11.0.0        # 画像処理
openpyxl==3.1.5       # Excel処理
```

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
   ============================================================
            文書画像抽出システム
     .docx/.pdfファイルから全画像を抽出してExcel出力
   ============================================================

   🔍 ステップ1: ファイル検索中...
      対象ディレクトリ: target
   ✅ 見つかったファイル:
      📄 .docx ファイル: 3 個
      📄 .pdf ファイル: 6 個
      📄 総ファイル数: 9 個

   🖼️  ステップ2: 画像抽出中...
      処理中 (1/9): sample.docx
        → 画像 2 枚を抽出
      ...

   ✅ 画像抽出完了: 総 195 枚

   📊 ステップ3: Excel出力中...
      出力ファイル: result.xlsx

   🎉 処理完了!
      処理時間: 0.27 秒
      処理ファイル数: 9
      抽出画像数: 195
      出力ファイル: result.xlsx
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
- **エラー処理**: 読み込めない画像は「エラー」と表示

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

## 🔧 技術仕様

### 使用ライブラリと役割
| ライブラリ | 機能 | 公式サイト |
|---|---|---|
| `python-docx` | Word文書の画像抽出 | https://python-docx.readthedocs.io/ |
| `PyMuPDF` | PDF文書の画像抽出 | https://pymupdf.readthedocs.io/ |
| `Pillow` | 画像処理・リサイズ | https://pillow.readthedocs.io/ |
| `openpyxl` | Excel読み書き | https://openpyxl.readthedocs.io/ |

### システム制限
- **最大ファイルサイズ**: 制限なし（メモリ依存）
- **対応画像形式**: PNG、JPEG、BMP、WMF、EMF等
- **最大画像数**: 制限なし（Excel列数上限まで）
- **処理速度**: 約700画像/秒（環境依存）

### ファイル構造の詳細
```
main.py の構成:
├── find_files()              # ファイル検索機能
├── extract_images_from_docx() # Word画像抽出
├── extract_images_from_pdf()  # PDF画像抽出
├── resize_image_for_excel()   # 画像リサイズ
├── create_excel_with_images() # Excel出力
└── main()                     # メイン処理
```

## 📞 サポート

### よくある質問
1. **Q: MacやLinuxでも動きますか？**
   A: Windows専用です。他のOSでは別途調整が必要です。

2. **Q: 商用利用は可能ですか？**
   A: 使用ライブラリのライセンスを確認してください。

3. **Q: 大量のファイルを処理できますか？**
   A: メモリ使用量に注意し、必要に応じてファイルを分割してください。

### エラーログの確認
- プログラム実行時のエラーメッセージを確認
- 警告メッセージは処理継続可能（一部画像が読み込めない場合）
- 致命的エラーは処理停止（環境設定や権限の問題）

---

**🎉 セットアップ完了後は、targetフォルダにファイルを配置して`python main.py`を実行するだけで使用できます！** 