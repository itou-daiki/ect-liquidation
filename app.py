import streamlit as st
import pandas as pd
import chardet
from datetime import datetime
import io
import requests
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import calendar
import shutil

# アプリケーションの設定
st.set_page_config(
    page_title="高速道路利用実績簿生成",
    page_icon="🛣️",
    layout="wide"
)

def detect_encoding(uploaded_file):
    """アップロードされたファイルのエンコーディングを検出"""
    raw_data = uploaded_file.read()
    uploaded_file.seek(0)  # ファイルポインタをリセット
    encoding = chardet.detect(raw_data)
    return encoding['encoding']

def load_csv_data(uploaded_file, encoding):
    """CSVファイルを読み込む"""
    try:
        df = pd.read_csv(uploaded_file, encoding=encoding)
        return df
    except Exception as e:
        st.error(f"CSVファイルの読み込みに失敗しました: {e}")
        return None

def extract_year_month(df):
    """データから年月を抽出"""
    if '利用年月日（自）' in df.columns:
        # 日付文字列から年月を抽出
        dates = df['利用年月日（自）'].dropna()
        sample_date = dates.iloc[0]
        
        # YY/MM/DD形式を解析
        if '/' in sample_date:
            parts = sample_date.split('/')
            if len(parts) >= 2:
                year = int(parts[0])
                month = int(parts[1])
                # 2桁年を4桁年に変換
                if year < 50:  # 25年以下は2025年以降と仮定
                    year += 2000
                elif year < 100:  # 50-99年は1950-1999年と仮定
                    year += 1900
                return year, month
    return None, None

def get_highway_sections():
    """高速道路区間のリストを取得"""
    # 九州地方の主要IC・SA・PA
    sections = [
        "大分米良",
        "日田",
        "福岡",
        "北九州",
        "熊本",
        "鹿児島",
        "宮崎",
        "佐賀",
        "長崎",
        "別府",
        "大分",
        "中津",
        "玖珠",
        "天瀬高塚",
        "杷木",
        "筑紫野",
        "太宰府",
        "春日",
        "古賀",
        "宗像",
        "若宮",
        "飯塚",
        "八幡",
        "小倉",
        "門司",
        "下関",
        "美祢",
        "山口",
        "防府",
        "徳山",
        "岩国"
    ]
    return sorted(sections)

def match_csv_to_template(df, year, month, highway_from, highway_to):
    """CSVデータをテンプレートの入力可能箇所にマッチング"""
    from datetime import datetime
    import calendar
    
    # 月の日数を取得
    last_day = calendar.monthrange(year, month)[1]
    
    # 入力データの初期化
    template_data = {
        'header_info': {
            'organization': '',  # C3
            'position': '',      # K3  
            'name': ''           # N3
        },
        'date_month': {
            'year': year - 2018,  # B5 (令和年)
            'month': month        # D5
        },
        'highway_info': {
            'from': highway_from,  # M5
            'to': highway_to,      # P5
            'one_way_fee': 2680    # M6
        },
        'daily_data': {}  # 日別の利用データ
    }
    
    # 日別データの初期化
    for day in range(1, last_day + 1):
        template_data['daily_data'][day] = {
            'outbound_confirmed': None,  # D列（往路利用確認）
            'outbound_amount': None,     # E列（往路利用金額）
            'return_confirmed': None,    # G列（復路利用確認）
            'return_amount': None        # H列（復路利用金額）
        }
    
    # CSVデータを解析して日別データにマッピング
    for index, row in df.iterrows():
        date_str = row['利用年月日（自）']
        
        try:
            if '/' in date_str:
                parts = date_str.split('/')
                if len(parts) >= 3:
                    year_part = int(parts[0])
                    month_part = int(parts[1])
                    day_part = int(parts[2])
                    
                    # 年を正規化
                    if year_part < 50:
                        year_part += 2000
                    elif year_part < 100:
                        year_part += 1900
                    
                    # 指定された年月と一致するかチェック
                    if year_part == year and month_part == month and 1 <= day_part <= last_day:
                        day = day_part
                        
                        # 往復判定
                        from_ic = str(row['利用ＩＣ（自）'])
                        to_ic = str(row['利用ＩＣ（至）'])
                        amount = row['通行料金']
                        
                        if highway_from in from_ic and highway_to in to_ic:
                            # 往路
                            template_data['daily_data'][day]['outbound_confirmed'] = '○'
                            template_data['daily_data'][day]['outbound_amount'] = amount
                        elif highway_to in from_ic and highway_from in to_ic:
                            # 復路
                            template_data['daily_data'][day]['return_confirmed'] = '○'
                            template_data['daily_data'][day]['return_amount'] = amount
        except (ValueError, IndexError):
            continue
    
    return template_data

def generate_expense_report_from_template(df, year, month, highway_from, highway_to, one_way_fee, monthly_allowance, organization="", position="", name=""):
    """テンプレートファイルをベースに利用実績簿を生成（水色箇所のみ入力）"""
    
    # テンプレートファイルをコピー
    template_path = '/workspaces/etc-statement-generator/生成するファイルの例/2025_04_高速道路等利用実績簿（テンプレート）.xlsx'
    wb = load_workbook(template_path)
    ws = wb.active
    
    # CSVデータをテンプレート形式にマッチング
    template_data = match_csv_to_template(df, year, month, highway_from, highway_to)
    
    # 水色箇所（入力可能箇所）のみに値を設定
    
    # ヘッダー情報
    if organization:
        ws['C3'] = organization
    if position:
        ws['K3'] = position  
    if name:
        ws['N3'] = name
    
    # 日付情報
    ws['B5'] = year - 2018  # 令和年
    ws['D5'] = month
    
    # 高速道路情報
    ws['M5'] = highway_from
    ws['P5'] = highway_to
    ws['M6'] = one_way_fee
    
    # 日別利用データを入力
    # 前半15日（13-27行）
    for row in range(13, 28):
        day = row - 12  # 1日から15日
        if day in template_data['daily_data']:
            data = template_data['daily_data'][day]
            
            # 往路データ
            if data['outbound_confirmed']:
                ws[f'D{row}'] = data['outbound_confirmed']
            if data['outbound_amount']:
                ws[f'E{row}'] = data['outbound_amount']
            
            # 復路データ
            if data['return_confirmed']:
                ws[f'G{row}'] = data['return_confirmed']
            if data['return_amount']:
                ws[f'H{row}'] = data['return_amount']
    
    # 後半（16-31日）の日別利用データを入力
    for row in range(13, 28):
        day = (row - 12) + 15  # 16日から31日（月によって調整）
        last_day = calendar.monthrange(year, month)[1]
        
        if day <= last_day and day in template_data['daily_data']:
            data = template_data['daily_data'][day]
            
            # 往路データ（右側）
            if data['outbound_confirmed']:
                ws[f'L{row}'] = data['outbound_confirmed']
            if data['outbound_amount']:
                ws[f'M{row}'] = data['outbound_amount']
            
            # 復路データ（右側）
            if data['return_confirmed']:
                ws[f'O{row}'] = data['return_confirmed']
            if data['return_amount']:
                ws[f'P{row}'] = data['return_amount']
    
    # 28日目の右側（L28, M28, O28, P28）も処理
    if 28 <= calendar.monthrange(year, month)[1]:
        day = 28
        if day in template_data['daily_data']:
            data = template_data['daily_data'][day]
            
            if data['outbound_confirmed']:
                ws['L28'] = data['outbound_confirmed']
            if data['outbound_amount']:
                ws['M28'] = data['outbound_amount']
            
            if data['return_confirmed']:
                ws['O28'] = data['return_confirmed']
            if data['return_amount']:
                ws['P28'] = data['return_amount']
    
    return wb

def main():
    st.title("🛣️ 高速道路利用実績簿生成システム")
    st.markdown("---")
    
    # サイドバーで設定
    st.sidebar.header("設定")
    
    # 基本情報設定
    st.sidebar.header("基本情報")
    organization = st.sidebar.text_input("所属", value="")
    position = st.sidebar.text_input("職", value="")
    name = st.sidebar.text_input("氏名", value="")
    
    st.sidebar.header("利用区間設定")
    # 高速道路区間選択
    highway_sections = get_highway_sections()
    highway_from = st.sidebar.selectbox("出発地点", highway_sections, index=0)  # 大分米良がデフォルト
    highway_to = st.sidebar.selectbox("到着地点", highway_sections, index=1)    # 日田がデフォルト
    
    st.sidebar.header("料金設定")
    # 片道料金設定
    one_way_fee = st.sidebar.number_input("片道料金（円）", min_value=0, value=2680, step=10)
    
    # 月間特別料金等加算額設定
    monthly_allowance = st.sidebar.number_input("月間特別料金等加算額（認定額）（円）", min_value=0, value=112560, step=100)
    
    # メインエリア
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("CSVファイルのアップロード")
        uploaded_file = st.file_uploader("ETCカード利用明細CSVファイルをアップロードしてください", type=['csv'])
        
        if uploaded_file is not None:
            # エンコーディング検出
            encoding = detect_encoding(uploaded_file)
            st.info(f"検出されたエンコーディング: {encoding}")
            
            # CSVデータ読み込み
            df = load_csv_data(uploaded_file, encoding)
            
            if df is not None:
                # 年月を抽出
                year, month = extract_year_month(df)
                
                if year and month:
                    st.success(f"データ期間: {year}年{month}月")
                    
                    # データプレビュー
                    st.subheader("データプレビュー")
                    st.dataframe(df.head(10))
                    
                    # 統計情報
                    total_records = len(df)
                    total_fee = df['通行料金'].sum()
                    
                    col1_stat, col2_stat, col3_stat = st.columns(3)
                    with col1_stat:
                        st.metric("総利用回数", f"{total_records}回")
                    with col2_stat:
                        st.metric("総利用料金", f"¥{total_fee:,}")
                    with col3_stat:
                        expected_trips = monthly_allowance // one_way_fee
                        st.metric("想定利用回数", f"{expected_trips}回")
                    
                    # 実績簿生成ボタン
                    if st.button("利用実績簿を生成", type="primary"):
                        try:
                            wb = generate_expense_report_from_template(df, year, month, highway_from, highway_to, one_way_fee, monthly_allowance, organization, position, name)
                            
                            # Excelファイルをバイナリデータに変換
                            excel_buffer = io.BytesIO()
                            wb.save(excel_buffer)
                            excel_buffer.seek(0)
                            
                            # ダウンロードボタン
                            st.download_button(
                                label="📥 Excelファイルをダウンロード",
                                data=excel_buffer.getvalue(),
                                file_name=f"高速道路利用実績簿_{year}年{month}月.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            st.success("利用実績簿が正常に生成されました！")
                            
                        except Exception as e:
                            st.error(f"実績簿の生成に失敗しました: {e}")
                else:
                    st.error("データから年月を抽出できませんでした。")
    
    with col2:
        st.header("使用方法")
        st.markdown("""
        1. **CSVファイルをアップロード**
           - ETCカード利用明細のCSVファイルを選択
        
        2. **設定を確認**
           - 出発地点・到着地点を選択
           - 片道料金を入力
           - 月間認定額を入力
        
        3. **実績簿を生成**
           - 「利用実績簿を生成」ボタンをクリック
           - Excelファイルをダウンロード
        """)
        
        st.markdown("---")
        st.subheader("現在の設定")
        if organization:
            st.write(f"**所属:** {organization}")
        if position:
            st.write(f"**職:** {position}")
        if name:
            st.write(f"**氏名:** {name}")
        st.write(f"**利用区間:** {highway_from} ⇔ {highway_to}")
        st.write(f"**片道料金:** ¥{one_way_fee:,}")
        st.write(f"**月間認定額:** ¥{monthly_allowance:,}")
        
        st.markdown("---")
        st.subheader("📋 テンプレート準拠")
        st.markdown("""
        **正式テンプレート使用**
        - 公式フォーマット完全準拠
        - 水色箇所のみデータ入力
        - 数式・レイアウト保持
        - Excel機能完全再現
        """)

if __name__ == "__main__":
    main()