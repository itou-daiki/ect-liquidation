# タスク完了時のガイドライン

## タスク完了時にすべきこと

### 1. コード品質チェック
このプロジェクトには自動化されたlint/format/testツールが設定されていないため、手動での確認が必要：

```bash
# Python構文チェック
python -m py_compile app.py

# アプリケーション起動テスト
streamlit run app.py
```

### 2. 機能テスト
- CSVファイルのアップロード機能
- エンコーディング自動検出
- データ表示・プレビュー
- Excel生成・ダウンロード機能
- 各設定項目の動作確認

### 3. 必須確認事項
- Streamlitアプリが正常に起動するか
- CSVアップロード後にデータが正しく表示されるか
- Excel生成ボタンが正常に動作するか
- 生成されたExcelファイルが正しい形式か

### 4. デプロイ前確認
```bash
# requirements.txtの更新確認
pip freeze > requirements.txt

# Procfileの動作確認
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

### 5. Git操作
```bash
# 変更内容の確認
git status
git diff

# コミット前の最終確認
streamlit run app.py  # 動作確認

# コミット
git add .
git commit -m "機能追加/修正内容の説明"
```

## 注意事項
- このプロジェクトは日本語UIのため、エラーメッセージも日本語で表示
- テンプレートファイルが必要なため、templates/ディレクトリの存在確認
- Herokuデプロイ時はProcfileの設定が重要