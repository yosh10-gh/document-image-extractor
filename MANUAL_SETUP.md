# 会社PC向け手動セットアップガイド

プロキシやネットワーク制限がある会社のパソコンでも、手動ダウンロードで確実にセットアップできる方法です。

## 📋 目次
1. [⚠️ 会社環境での使用前の重要な注意事項](#会社環境での使用前の重要な注意事項)
2. [事前準備](#事前準備)
3. [Python環境の手動セットアップ](#python環境の手動セットアップ)
4. [プロジェクトファイルの準備](#プロジェクトファイルの準備)
5. [必要ライブラリの手動インストール](#必要ライブラリの手動インストール)
6. [動作確認](#動作確認)
7. [トラブルシューティング](#トラブルシューティング)

## ⚠️ 会社環境での使用前の重要な注意事項

### セキュリティとコンプライアンスの確認

**使用前に必ず確認してください:**

1. **IT部門への相談**
   - 社外ツールの使用許可
   - Python環境の構築許可
   - 業務データの処理に関する承認

2. **会社ポリシーの確認**
   - ソフトウェアインストール規則
   - 社外ライブラリの使用制限
   - データ処理・保存に関する規定

3. **セキュリティ設定の維持**
   - プロキシ設定は**絶対に変更しない**
   - ファイアウォール設定はそのまま
   - ウイルス対策ソフトは有効のまま

4. **データの取り扱い**
   - 機密文書の処理は事前承認を得る
   - 結果ファイルの保存場所を確認
   - 個人情報を含むファイルは慎重に処理

**推奨アプローチ:**
- まず**テスト用ファイル**で動作確認
- **非機密データ**での試用
- **IT部門承認後**に本格運用

## 🚀 事前準備

### 必要なファイルを事前にダウンロード
以下のファイルを自宅PCなどでダウンロードし、USBメモリ等で会社PCに持ち込んでください：

#### 1. Pythonインストーラー
- **Python 3.12.10** (Windows x86-64)
- ダウンロード先: [https://www.python.org/downloads/release/python-31210/](https://www.python.org/downloads/release/python-31210/)
- ファイル名: `python-3.12.10-amd64.exe`
- ⚠️ **重要**: 「Windows installer (64-bit)」を選択してください

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
- ダウンロード: [https://pypi.org/project/lxml/5.4.0/#files](https://pypi.org/project/lxml/5.4.0/#files)
- `et_xmlfile-2.0.0-py3-none-any.whl` (openpyxlの依存)
- ダウンロード: [https://pypi.org/project/et-xmlfile/2.0.0/#files](https://pypi.org/project/et-xmlfile/2.0.0/#files)

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

⚠️ **wheelファイルの配置場所**: 
```
C:\document-image-extractor\
├── main.py
├── python_docx-1.1.2-py3-none-any.whl        ← ここに配置
├── PyMuPDF-1.24.12-cp312-cp312-win_amd64.whl ← ここに配置
├── (その他のwheelファイルも同様)
└── venv\
```

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
⚠️ **重要**: 会社のプロキシ設定は**セキュリティ上重要**です。設定変更は避けてください。

**推奨対処法（安全）:**
```cmd
# 完全ローカルインストール（インターネット接続不要）
pip install --no-index --find-links . パッケージ名

# または、--no-depsオプションで依存関係チェックを無効化
pip install --no-deps --find-links . パッケージ名
```

**プロキシ問題の根本的解決策:**
1. **IT部門に相談**: 業務目的での使用許可を求める
2. **完全オフライン**: wheelファイルのみでインストール
3. **代替手段**: 会社承認済みのPython環境があるか確認

❌ **やってはいけないこと:**
- プロキシ設定の変更・無効化
- セキュリティソフトの停止
- 管理者権限でのネットワーク設定変更

## 📁 最終的なファイル配置例

```
C:\document-image-extractor\
├── main.py                                          # メインプログラム
├── requirements.txt                                 # 依存関係一覧
├── README.md                                        # 使用方法
├── MANUAL_SETUP.md                                  # このファイル
├── LICENSE                                          # ライセンス
├── CONTRIBUTING.md                                  # 貢献ガイド
├── target\                                         # 処理対象ファイル配置フォルダ
│   └── samples\                                    # サンプルファイル
├── venv\                                          # 仮想環境（作成される）
├── result.xlsx                                    # 結果ファイル（生成される）
│
├── python_docx-1.1.2-py3-none-any.whl            # ダウンロードしたwheelファイル
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

## 💼 IT部門向け情報

### セキュリティ観点での本ツールの特徴

**安全性:**
- **外部通信なし**: インストール後はオフラインで動作
- **データ送信なし**: ローカルファイル処理のみ
- **権限要求最小**: 標準ユーザー権限で動作
- **オープンソース**: コードが公開されており検証可能

**使用ライブラリ（すべて著名で安全）:**
- python-docx: Microsoft Word文書処理
- PyMuPDF: PDF処理（Mozilla財団関連）
- Pillow: 画像処理（Python公式推奨）
- openpyxl: Excel処理

**ネットワーク要件:**
- インストール時のみインターネット接続必要
- 運用時は完全オフライン動作
- プロキシ設定の変更不要

### 推奨セットアップ方法

1. **承認済み環境**: 既存のPython環境がある場合はそれを使用
2. **サンドボックス**: 仮想環境での隔離実行
3. **監査ログ**: 処理ファイルと結果ファイルの記録

---

**⚠️ 重要: 会社のセキュリティポリシーを最優先に、IT部門の承認を得てから使用してください。** 🔒

何か問題が発生した場合は、具体的なエラーメッセージを確認して対処してください。 