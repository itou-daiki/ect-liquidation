# コードベース構造

## ディレクトリ構造
```
/workspaces/etc-statement-generator/
├── app.py                    # メインアプリケーションファイル
├── requirements.txt          # Python依存関係
├── Procfile                 # Herokuデプロイ設定
├── README.md                # プロジェクト説明
├── templates/               # Excelテンプレートディレクトリ
│   └── 2025_04_高速道路等利用実績簿（テンプレート）.xlsx
├── __pycache__/            # Python キャッシュディレクトリ
├── .streamlit/             # Streamlit設定
│   └── config.toml
├── .devcontainer/          # VS Code Dev Container設定
│   └── devcontainer.json
└── .serena/                # Serena MCP設定
    └── project.yml
```

## app.py の関数構造
1. `detect_encoding(uploaded_file)` - ファイルエンコーディング検出
2. `load_csv_data(uploaded_file, encoding)` - CSVデータ読み込み
3. `extract_year_month(df)` - データから年月抽出
4. `get_highway_sections()` - 高速道路区間リスト取得
5. `calculate_usage_amount_by_date_formula(df, one_way_fee)` - 利用料金計算
6. `calculate_daily_usage_from_csv(df, one_way_fee)` - 日別利用実績計算
7. `generate_expense_report_from_template(...)` - Excelレポート生成
8. `main()` - メインアプリケーション関数

## 主要な設定
- Streamlitポート: 8501 (Dev Container)
- 最大アップロードサイズ: 200MB
- デフォルト片道料金: 2,680円
- デフォルト月間認定額: 112,560円