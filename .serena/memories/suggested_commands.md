# 推奨コマンド

## 開発・実行コマンド

### アプリケーション実行
```bash
# ローカル開発実行
streamlit run app.py

# 本番環境実行（Heroku）
streamlit run app.py --server.port $PORT --server.address 0.0.0.0

# Dev Container環境実行
streamlit run app.py --server.enableCORS false --server.enableXsrfProtection false
```

### パッケージ管理
```bash
# 依存関係インストール
pip install -r requirements.txt

# 新しい依存関係追加後
pip freeze > requirements.txt
```

### Git操作
```bash
# 現在の状態確認
git status

# 変更をステージング
git add .

# コミット
git commit -m "コミットメッセージ"

# プッシュ
git push origin main
```

## ファイル操作
```bash
# ディレクトリ構造確認
ls -la

# ファイル内容確認
cat app.py

# ファイル検索
find . -name "*.py"

# パターン検索
grep -r "関数名" .
```

## 開発・テスト
このプロジェクトには以下がありません：
- テストフレームワーク
- リンター設定
- フォーマッター設定
- 自動化されたCI/CD

## システム情報
- **OS**: Linux
- **Python**: 3.11 (Dev Container)
- **ポート**: 8501 (Streamlit)