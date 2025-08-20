# コードスタイルと規約

## 命名規約
- **関数名**: snake_case (例: `detect_encoding`, `load_csv_data`)
- **変数名**: snake_case (例: `uploaded_file`, `one_way_fee`)
- **定数**: 大文字のSNAKE_CASE使用無し（リテラル値を直接使用）

## ドキュメント
- **関数ドキュメント**: 日本語のdocstring使用
  ```python
  def detect_encoding(uploaded_file):
      """アップロードされたファイルのエンコーディングを検出"""
  ```

## コードスタイル
- **インデント**: 4スペース
- **エンコーディング**: UTF-8
- **改行**: LF (Unix style)
- **import順序**: 標準ライブラリ → サードパーティ → ローカル

## Streamlit固有の規約
- **ページ設定**: `st.set_page_config()` で日本語タイトルとアイコン設定
- **UI構成**: サイドバーで設定、メインエリアでデータ表示
- **エラーハンドリング**: `st.error()`, `st.warning()`, `st.success()` 使用
- **日本語UI**: すべてのUI要素が日本語

## データ処理
- **CSVエンコーディング**: chardetによる自動検出
- **エラーハンドリング**: try-except文で適切なエラーメッセージ表示
- **データ表示**: `st.dataframe()`, `st.metric()` 使用

## ファイル命名
- **出力ファイル**: 日本語ファイル名使用 (例: `高速道路利用実績簿_2025年4月.xlsx`)