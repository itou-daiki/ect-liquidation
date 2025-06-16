# 高速道路利用実績簿生成システム

CSVファイルから高速道路利用実績簿を自動生成するWebアプリケーションです。

## 機能

- CSVファイルのアップロード（自動エンコーディング検出）
- 年月の自動抽出
- 高速道路区間の選択
- 片道料金と月間認定額の設定
- Excel形式の利用実績簿の生成とダウンロード

## 使用方法

1. **CSVファイルをアップロード**
   - ETCカード利用明細のCSVファイルを選択してアップロード

2. **設定を調整**
   - 出発地点・到着地点を選択
   - 片道料金を入力（デフォルト: 2,680円）
   - 月間特別料金等加算額（認定額）を入力（デフォルト: 112,560円）

3. **実績簿を生成**
   - 「利用実績簿を生成」ボタンをクリック
   - Excelファイルをダウンロード

## セットアップ

### 必要な環境
- Python 3.8以上

### インストール
```bash
pip install -r requirements.txt
```

### 実行
```bash
streamlit run app.py
```

## デプロイ

### Streamlit Cloud
1. GitHubリポジトリにコードをプッシュ
2. [Streamlit Cloud](https://streamlit.io/cloud)にアクセス
3. リポジトリを選択してデプロイ

### Heroku
1. `Procfile`を作成:
```
web: streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

2. Herokuアプリを作成してデプロイ:
```bash
heroku create your-app-name
git push heroku main
```

## ファイル構成

```
etc-statement-generator/
├── app.py                 # メインアプリケーション
├── requirements.txt       # 依存関係
├── README.md             # このファイル
├── アップロードするファイルの例/
│   └── 202506170826-cleaned.csv  # サンプルCSVファイル
└── 生成するファイルの例/
    └── (生成されたExcelファイルの例)
```

## 対応するCSVフォーマット

以下の列が含まれたCSVファイルに対応しています：

- 利用年月日（自）
- 時分（自）
- 利用年月日（至）
- 時分（至）
- 利用ＩＣ（自）
- 利用ＩＣ（至）
- 通行料金
- 備考
- その他のETCカード利用明細に含まれる標準的な列

## 注意事項

- CSVファイルはShift_JISエンコーディングに対応
- 生成されるExcelファイルは日本語フォント（MS Gothic）を使用
- 高速道路区間リストは九州地方の主要ICを含む