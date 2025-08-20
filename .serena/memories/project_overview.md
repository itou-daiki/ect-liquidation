# 高速道路利用実績簿生成システム - プロジェクト概要

## プロジェクトの目的
CSVファイルから高速道路利用実績簿を自動生成するWebアプリケーション。ETCカード利用明細のCSVファイルをアップロードし、Excel形式の利用実績簿を生成・ダウンロードできる。

## 技術スタック
- **フレームワーク**: Streamlit (バージョン 1.45.1)
- **言語**: Python
- **主要ライブラリ**:
  - pandas 2.2.3 (データ処理)
  - chardet 5.2.0 (エンコーディング自動検出)
  - openpyxl 3.1.5 (Excelファイル操作)
  - requests 2.32.3 (HTTP通信)

## 主な機能
1. CSVファイルのアップロード（自動エンコーディング検出）
2. 年月の自動抽出
3. 高速道路区間の選択
4. 片道料金と月間認定額の設定
5. Excel形式の利用実績簿の生成とダウンロード

## デプロイメント
- Heroku対応 (Procfile使用)
- ポート設定: $PORT環境変数使用
- Dev Container対応 (.devcontainer/devcontainer.json)

## 設定ファイル
- `.streamlit/config.toml`: Streamlitアプリの設定（テーマ、アップロードサイズ制限など）
- `Procfile`: Herokuデプロイ用設定
- `.devcontainer/devcontainer.json`: VS Code Dev Container設定