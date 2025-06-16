import streamlit as st
import pandas as pd
import chardet
from datetime import datetime
import io
import requests
import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    """高速道路区間のリストを取得（実際のAPIの代わりにダミーデータ）"""
    # 実際のAPIを使用する場合はここを修正
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
        "福岡",
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

def generate_expense_report(df, year, month, highway_from, highway_to, one_way_fee, monthly_allowance):
    """利用実績簿を生成"""
    # Excelワークブックを作成
    wb = Workbook()
    ws = wb.active
    ws.title = f"{year}年{month}月利用実績簿"
    
    # スタイル設定
    title_font = Font(name='MS Gothic', size=16, bold=True)
    subtitle_font = Font(name='MS Gothic', size=12, bold=True)
    header_font = Font(name='MS Gothic', size=11, bold=True)
    normal_font = Font(name='MS Gothic', size=10)
    thick_border = Border(
        left=Side(style='thick'),
        right=Side(style='thick'),
        top=Side(style='thick'),
        bottom=Side(style='thick')
    )
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # タイトル行
    ws.merge_cells('A1:I1')
    ws['A1'] = f"高速道路等利用実績簿"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 25
    
    # 年月行
    ws.merge_cells('A2:I2')
    ws['A2'] = f"（{year}年{month}月分）"
    ws['A2'].font = subtitle_font
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 20
    
    # 空行
    ws.row_dimensions[3].height = 10
    
    # 基本情報
    ws['A4'] = "利用区間"
    ws['A4'].font = header_font
    ws.merge_cells('B4:D4')
    ws['B4'] = f"{highway_from} ⇔ {highway_to}"
    ws['B4'].font = normal_font
    ws['B4'].alignment = Alignment(horizontal='left')
    
    ws['E4'] = "片道料金"
    ws['E4'].font = header_font
    ws.merge_cells('F4:G4')
    ws['F4'] = f"¥{one_way_fee:,}"
    ws['F4'].font = normal_font
    ws['F4'].alignment = Alignment(horizontal='right')
    
    ws['A5'] = "月間特別料金等加算額（認定額）"
    ws['A5'].font = header_font
    ws.merge_cells('B5:D5')
    ws['B5'] = f"¥{monthly_allowance:,}"
    ws['B5'].font = normal_font
    ws['B5'].alignment = Alignment(horizontal='right')
    
    # 空行
    ws.row_dimensions[6].height = 10
    
    # ヘッダー行
    headers = ['利用日', '出発IC', '到着IC', '出発時刻', '到着時刻', '通行料金', '往復区分', '備考', '認定回数']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col, value=header)
        cell.font = header_font
        cell.border = thick_border
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[7].height = 18
    
    # データ行を追加
    row = 8
    total_fee = 0
    certified_count = 0
    
    for index, data_row in df.iterrows():
        # 利用日（YY/MM/DD → YYYY/MM/DD）
        date_str = data_row['利用年月日（自）']
        if '/' in date_str:
            parts = date_str.split('/')
            if len(parts) >= 3:
                year_part = int(parts[0])
                if year_part < 50:
                    year_part += 2000
                elif year_part < 100:
                    year_part += 1900
                formatted_date = f"{year_part}/{parts[1]}/{parts[2]}"
            else:
                formatted_date = date_str
        else:
            formatted_date = date_str
            
        ws.cell(row=row, column=1, value=formatted_date)
        ws.cell(row=row, column=2, value=data_row['利用ＩＣ（自）'])
        ws.cell(row=row, column=3, value=data_row['利用ＩＣ（至）'])
        ws.cell(row=row, column=4, value=data_row['時分（自）'])
        ws.cell(row=row, column=5, value=data_row['時分（至）'])
        ws.cell(row=row, column=6, value=data_row['通行料金'])
        ws.cell(row=row, column=6).number_format = '¥#,##0'
        
        # 往復判定（より詳細な判定）
        from_ic = str(data_row['利用ＩＣ（自）'])
        to_ic = str(data_row['利用ＩＣ（至）'])
        
        if highway_from in from_ic and highway_to in to_ic:
            direction = "往路"
            certified_count += 1
        elif highway_to in from_ic and highway_from in to_ic:
            direction = "復路"
            certified_count += 1
        else:
            direction = "対象外"
            
        ws.cell(row=row, column=7, value=direction)
        ws.cell(row=row, column=8, value=data_row['備考'])
        
        # 認定回数（往復の場合のみカウント）
        if direction in ["往路", "復路"]:
            ws.cell(row=row, column=9, value=1)
        else:
            ws.cell(row=row, column=9, value=0)
        
        total_fee += data_row['通行料金']
        
        # セルにボーダーを適用
        for col in range(1, 10):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.font = normal_font
            cell.alignment = Alignment(horizontal='center' if col in [1, 4, 5, 7, 9] else 'left')
        
        row += 1
    
    # 合計行
    ws.cell(row=row, column=5, value="合計")
    ws.cell(row=row, column=5).font = header_font
    ws.cell(row=row, column=6, value=total_fee)
    ws.cell(row=row, column=6).number_format = '¥#,##0'
    ws.cell(row=row, column=6).font = header_font
    ws.cell(row=row, column=9, value=certified_count)
    ws.cell(row=row, column=9).font = header_font
    
    for col in range(1, 10):
        ws.cell(row=row, column=col).border = thick_border
    
    # 承認欄
    row += 2
    ws.cell(row=row, column=1, value="承認者")
    ws.cell(row=row, column=1).font = header_font
    ws.merge_cells(f'B{row}:D{row}')
    ws.cell(row=row, column=2, value="印")
    ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
    ws.cell(row=row, column=2).border = thin_border
    
    ws.cell(row=row, column=6, value="申請者")
    ws.cell(row=row, column=6).font = header_font
    ws.merge_cells(f'G{row}:I{row}')
    ws.cell(row=row, column=7, value="印")
    ws.cell(row=row, column=7).alignment = Alignment(horizontal='center')
    ws.cell(row=row, column=7).border = thin_border
    
    # 列幅を調整
    column_widths = [12, 18, 18, 10, 10, 12, 12, 25, 8]
    for col, width in enumerate(column_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    
    return wb

def main():
    st.title("🛣️ 高速道路利用実績簿生成システム")
    st.markdown("---")
    
    # サイドバーで設定
    st.sidebar.header("設定")
    
    # 高速道路区間選択
    highway_sections = get_highway_sections()
    highway_from = st.sidebar.selectbox("出発地点", highway_sections, index=0)  # 大分米良がデフォルト
    highway_to = st.sidebar.selectbox("到着地点", highway_sections, index=1)    # 日田がデフォルト
    
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
                            wb = generate_expense_report(df, year, month, highway_from, highway_to, one_way_fee, monthly_allowance)
                            
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
        st.write(f"**利用区間:** {highway_from} ⇔ {highway_to}")
        st.write(f"**片道料金:** ¥{one_way_fee:,}")
        st.write(f"**月間認定額:** ¥{monthly_allowance:,}")

if __name__ == "__main__":
    main()