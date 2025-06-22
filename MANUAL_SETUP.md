# 会社PC向け手動セットアップガイド

プロキシやネットワーク制限がある会社のパソコンでも、手動ダウンロードで確実にセットアップできる方法です。

## 📋 目次
1. [事前準備](#事前準備)
2. [Python環境の手動セットアップ](#python環境の手動セットアップ)
3. [プロジェクトファイルの準備](#プロジェクトファイルの準備)
4. [必要ライブラリの手動インストール](#必要ライブラリの手動インストール)
5. [動作確認](#動作確認)
6. [トラブルシューティング](#トラブルシューティング)

## 🚀 事前準備

### 必要なファイルを事前にダウンロード
以下のファイルを自宅PCなどでダウンロードし、USBメモリ等で会社PCに持ち込んでください：

#### 1. Pythonインストーラー
- **Python 3.12.10** (Windows x86-64)
- ダウンロード先: [https://www.python.org/downloads/release/python-31210/](https://www.python.org/downloads/release/python-31210/)
- ファイル名: `python-3.12.10-amd64.exe`

#### 2. プロジェクトファイル
- **GitHub ZIP**: [https://github.com/yosh10-gh/document-image-extractor/archive/refs/heads/main.zip](https://github.com/yosh10-gh/document-image-extractor/archive/refs/heads/main.zip)
- または「Code」→「Download ZIP」からダウンロード

#### 3. 必要ライブラリ（wheel形式）
以下を [https://pypi.org/](https://pypi.org/) から手動ダウンロード：

**python-docx 1.1.2:**
- `python_docx-1.1.2-py3-none-any.whl`
- ダウンロード: [https://pypi.org/project/python-docx/1.1.2/#files](https://pypi.org/project/python-docx/1.1.2/#files)

**PyMuPDF 1.24.12:**
- `PyMuPDF-1.24.12-cp312-cp312-win_amd64.whl`
- ダウンロード: [https://pypi.org/project/PyMuPDF/1.24.12/#files](https://pypi.org/project/PyMuPDF/1.24.12/#files)

**Pillow 11.0.0:**
- `pillow-11.0.0-cp312-cp312-win_amd64.whl`
- ダウンロード: [https://pypi.org/project/pillow/11.0.0/#files](https://pypi.org/project/pillow/11.0.0/#files)

**openpyxl 3.1.5:**
- `openpyxl-3.1.5-py2.py3-none-any.whl`
- ダウンロード: [https://pypi.org/project/openpyxl/3.1.5/#files](https://pypi.org/project/openpyxl/3.1.5/#files)

**依存関係ライブラリ:**
- `lxml-5.4.0-cp312-cp312-win_amd64.whl` (python-docxの依存)
- `et_xmlfile-2.0.0-py3-none-any.whl` (openpyxlの依存)

## 💻 Python環境の手動セットアップ

### ステップ1: Pythonのインストール

1. **インストーラーを実行**
   ```
   python-3.12.10-amd64.exe を右クリック → 「管理者として実行」
   ```

2. **重要な設定**
   - ✅ **「Add Python 3.12 to PATH」にチェック**
   - ✅ **「Install for all users」にチェック**
   - 「Install Now」をクリック

3. **インストール確認**
   - PowerShellまたはコマンドプロンプトを開く
   - `python --version` を実行
   - `Python 3.12.10` と表示されればOK

### ステップ2: プロジェクトフォルダの準備

1. **ZIPファイルを展開**
   ```
   document-image-extractor-main.zip を右クリック → 「すべて展開」
   ```

2. **フォルダを移動**
   ```
   展開されたフォルダを以下の場所に移動:
   C:\document-image-extractor\
   ```

3. **フォルダ構成確認**
   ```
   C:\document-image-extractor\
   ├── main.py
   ├── requirements.txt
   ├── README.md
   ├── target\
   └── その他のファイル
   ```

## 📦 必要ライブラリの手動インストール

### ステップ1: 仮想環境の作成

1. **プロジェクトフォルダに移動**
   ```cmd
   cd C:\document-image-extractor
   ```

2. **仮想環境を作成**
   ```cmd
   python -m venv venv
   ```

3. **仮想環境を有効化**
   ```cmd
   venv\Scripts\activate
   ```
   ✅ プロンプトに `(venv)` が表示されればOK

### ステップ2: wheelファイルから手動インストール

**事前にダウンロードしたwheelファイルをプロジェクトフォルダに配置してから実行:**

```cmd
# 基本ライブラリのインストール
pip install --no-index --find-links . lxml-5.4.0-cp312-cp312-win_amd64.whl
pip install --no-index --find-links . et_xmlfile-2.0.0-py3-none-any.whl

# メインライブラリのインストール
pip install --no-index --find-links . python_docx-1.1.2-py3-none-any.whl
pip install --no-index --find-links . PyMuPDF-1.24.12-cp312-cp312-win_amd64.whl
pip install --no-index --find-links . pillow-11.0.0-cp312-cp312-win_amd64.whl
pip install --no-index --find-links . openpyxl-3.1.5-py2.py3-none-any.whl
```

### ステップ3: インストール確認

```cmd
pip list
```

以下のパッケージが表示されればOK:
```
lxml            5.4.0
et-xmlfile      2.0.0
python-docx     1.1.2
PyMuPDF         1.24.12
Pillow          11.0.0
openpyxl        3.1.5
```

## 🎮 動作確認

### テスト実行

1. **サンプルファイルでテスト**
   ```cmd
   cd C:\document-image-extractor
   venv\Scripts\activate
   python main.py
   ```

2. **正常終了の確認**
   ```
   🎉 処理完了!
      処理時間: X.XX 秒
      処理ファイル数: X
      抽出画像数: X
      出力ファイル: result.xlsx
   ```

3. **結果ファイルの確認**
   - `result.xlsx` ファイルが作成される
   - Excelで開いて画像が表示されることを確認

## 🛠️ トラブルシューティング

### Q1: 「python コマンドが認識されない」
```cmd
# 解決方法1: フルパスで実行
C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\python.exe --version

# 解決方法2: 環境変数を手動設定
システムプロパティ → 環境変数 → PATH に追加:
C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\
C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\Scripts\
```

### Q2: 「wheelファイルのインストールに失敗」
```cmd
# 依存関係の問題の場合、順番を変えて実行
pip install --no-index --find-links . --force-reinstall lxml-5.4.0-cp312-cp312-win_amd64.whl

# または個別にインストール
pip install --no-deps --find-links . python_docx-1.1.2-py3-none-any.whl
```

### Q3: 「管理者権限が必要」エラー
```cmd
# PowerShellを管理者権限で開いて実行
右クリック → 「管理者として実行」
```

### Q4: 「プロキシエラー」が出る場合
```cmd
# プロキシ設定を無効化してローカルインストール
pip install --no-index --find-links . --trusted-host pypi.org --trusted-host pypi.python.org パッケージ名
```

## 📁 ファイル配置例

```
C:\document-image-extractor\
├── main.py                                          # メインプログラム
├── requirements.txt                                 # 依存関係一覧
├── target\                                         # 処理対象ファイル配置フォルダ
├── venv\                                          # 仮想環境（作成される）
├── result.xlsx                                    # 結果ファイル（生成される）
│
└── wheels\                                        # ダウンロードしたwheelファイル
    ├── python_docx-1.1.2-py3-none-any.whl
    ├── PyMuPDF-1.24.12-cp312-cp312-win_amd64.whl
    ├── pillow-11.0.0-cp312-cp312-win_amd64.whl
    ├── openpyxl-3.1.5-py2.py3-none-any.whl
    ├── lxml-5.4.0-cp312-cp312-win_amd64.whl
    └── et_xmlfile-2.0.0-py3-none-any.whl
```

## 🎯 実際の使用方法

### 日常的な使用手順

1. **プロジェクトフォルダに移動**
   ```cmd
   cd C:\document-image-extractor
   ```

2. **仮想環境を有効化**
   ```cmd
   venv\Scripts\activate
   ```

3. **処理したいファイルを配置**
   - `target\` フォルダに .docx や .pdf ファイルをコピー

4. **プログラム実行**
   ```cmd
   python main.py
   ```

5. **結果確認**
   - `result.xlsx` ファイルを開いて画像を確認

## 💡 ヒント

### 効率的な運用方法
- **バッチファイル作成**: 毎回のコマンド入力を省略
- **デスクトップショートカット**: ワンクリックでフォルダアクセス
- **定期実行**: タスクスケジューラで自動化

### バッチファイル例
`run_extractor.bat` として保存:
```batch
@echo off
cd /d C:\document-image-extractor
call venv\Scripts\activate
python main.py
pause
```

---

**これで会社のPCでも確実に動作します！** 🚀

何か問題が発生した場合は、具体的なエラーメッセージを確認して対処してください。 