# 貢献ガイド - Contributing Guide

このプロジェクトへの貢献を歓迎します！このガイドでは、プロジェクトに貢献する方法を説明します。

## 🚀 クイックスタート

### 1. リポジトリをフォーク・クローン
```bash
git clone https://github.com/YOUR_USERNAME/document-image-extractor.git
cd document-image-extractor
```

### 2. 開発環境をセットアップ
詳細な手順は[README.md](README.md)を参照してください。

```bash
# Python 3.12.10をインストール（pyenv-win使用）
pyenv install 3.12.10
pyenv local 3.12.10

# 仮想環境を作成・有効化
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # macOS/Linux

# 依存関係をインストール
pip install -r requirements.txt
```

### 3. 動作テスト
```bash
# 基本テスト
python main.py

# 改良版テスト  
python main_improved.py
```

## 📋 貢献の種類

### バグ報告
- **Issue**を作成して詳細を報告
- 再現手順、期待する動作、実際の動作を記載
- エラーメッセージやスクリーンショットを添付

### 機能要求
- **Feature Request**のIssueを作成
- 用途、必要性、実装案を詳しく説明

### コード貢献
1. **Issue**を確認（または新規作成）
2. **ブランチ**を作成: `git checkout -b feature/your-feature`
3. **コード**を実装
4. **テスト**を実行
5. **Pull Request**を作成

## 🔧 開発ガイドライン

### コードスタイル
- **PEP 8**に従う
- **関数・クラス**には適切なdocstringを追加
- **変数名**は分かりやすく日本語コメント推奨

### ファイル構成
```
project/
├── main.py              # 基本版メインプログラム
├── main_improved.py     # 改良版メインプログラム  
├── step*.py            # 開発段階ファイル（学習用）
├── requirements.txt    # 依存関係
├── README.md          # メインドキュメント
├── CONTRIBUTING.md    # このファイル
└── target/           # テスト用ファイル配置
```

### 新機能追加時のチェックリスト
- [ ] 既存機能に影響しないか確認
- [ ] エラーハンドリングを適切に実装
- [ ] 日本語での分かりやすいエラーメッセージ
- [ ] README.mdの更新（必要に応じて）
- [ ] テストファイルでの動作確認

## 🧪 テスト方法

### 基本テスト
```bash
# サンプルファイルでテスト
python main_improved.py
```

### エラーテスト
- 存在しないファイル
- 破損したファイル
- 権限のないファイル
- 大容量ファイル

### パフォーマンステスト
- 大量ファイル処理
- メモリ使用量確認
- 処理時間測定

## 📚 実装可能な改良アイデア

### 新機能
- [ ] PowerPoint (.pptx) 対応
- [ ] 画像形式選択機能 (PNG/JPEG)
- [ ] バッチ処理モード
- [ ] GUI版の作成
- [ ] 画像サイズカスタマイズ
- [ ] 並列処理による高速化

### 改良案
- [ ] プログレスバー追加
- [ ] ログ出力機能
- [ ] 設定ファイル対応
- [ ] 多言語対応
- [ ] クラウドストレージ対応

### コード品質
- [ ] 単体テスト追加
- [ ] 型ヒント強化
- [ ] ドキュメント自動生成
- [ ] CI/CD設定

## 🐛 既知の問題

### 制限事項
- Windows専用（macOS/Linux未対応）
- WMF/EMF形式で一部警告
- 巨大ファイルでメモリ不足の可能性

### 改善予定
- 他OS対応
- メモリ効率の改善
- エラーメッセージの詳細化

## 💡 Pull Request ガイドライン

### 前準備
1. **最新のmain**をpull
2. **新しいブランチ**を作成
3. **関連Issue**があることを確認

### PR作成時
- **明確なタイトル**をつける
- **変更内容**を詳しく説明
- **テスト結果**を記載
- **スクリーンショット**を添付（UI変更時）

### レビュー対応
- **迅速な対応**を心がける
- **建設的な議論**を歓迎
- **コードの説明**を追加

## 📧 連絡先

質問や提案がある場合：
- **GitHub Issues**: バグ報告・機能要求
- **GitHub Discussions**: 一般的な議論
- **Pull Request**: コード貢献

## 🎉 貢献者

このプロジェクトに貢献してくださったすべての方に感謝します！

---

**Happy Coding! 🚀** 