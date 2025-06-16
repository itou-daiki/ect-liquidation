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

def generate_expense_report(df, year, month, highway_from, highway_to, one_way_fee, monthly_allowance, organization="", position="", name=""):
    """高速道路等利用実績簿を参考ファイルと完全に同じ形式で生成"""
    from datetime import datetime, timedelta
    import calendar
    
    wb = Workbook()
    ws = wb.active
    ws.title = "利用実績簿"
    
    # フォント設定（参考ファイルと同じ）
    ms_mincho = Font(name='ＭＳ 明朝')
    ms_gothic = Font(name='ＭＳ ゴシック')
    ms_p_mincho = Font(name='ＭＳ Ｐ明朝')
    
    # 境界線設定
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # タイトル行（A1:S1をマージ）
    ws.merge_cells('A1:S1')
    ws['A1'] = '高速道路等利用実績簿'
    ws['A1'].font = Font(name='ＭＳ 明朝', size=16)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24
    
    # 空行
    ws.row_dimensions[2].height = 17.25
    
    # 所属・職・氏名行
    ws.merge_cells('A3:B3')
    ws['A3'] = '所　　属'
    ws['A3'].font = Font(name='ＭＳ 明朝', size=11)
    ws['A3'].alignment = Alignment(vertical='center')
    ws['A3'].border = thin_border
    
    ws.merge_cells('C3:H3')
    ws['C3'] = organization if organization else ''
    ws['C3'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['C3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['C3'].border = thin_border
    
    ws['J3'] = '職'
    ws['J3'].font = Font(name='ＭＳ 明朝', size=11)
    ws['J3'].alignment = Alignment(vertical='center')
    ws['J3'].border = thin_border
    
    ws.merge_cells('K3:L3')
    ws['K3'] = position if position else ''
    ws['K3'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['K3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['K3'].border = thin_border
    
    ws['M3'] = '氏名'
    ws['M3'].font = Font(name='ＭＳ 明朝', size=11)
    ws['M3'].alignment = Alignment(vertical='center')
    ws['M3'].border = thin_border
    
    ws.merge_cells('N3:Q3')
    ws['N3'] = name if name else ''
    ws['N3'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['N3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['N3'].border = thin_border
    
    ws.row_dimensions[3].height = 24
    ws.row_dimensions[4].height = 8.25
    
    # 年月行
    ws['A5'] = '令和'
    ws['A5'].font = Font(name='ＭＳ 明朝', size=11)
    ws['A5'].alignment = Alignment(vertical='center')
    
    # 令和年換算（西暦年 - 2018）
    reiwa_year = year - 2018
    ws['B5'] = reiwa_year
    ws['B5'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['B5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['B5'].border = thin_border
    
    ws['C5'] = '年'
    ws['C5'].font = Font(name='ＭＳ 明朝', size=10)
    ws['C5'].alignment = Alignment(vertical='center')
    
    ws['D5'] = month
    ws['D5'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['D5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['D5'].border = thin_border
    
    ws['E5'] = '月分'
    ws['E5'].font = Font(name='ＭＳ 明朝', size=10)
    ws['E5'].alignment = Alignment(vertical='center')
    
    # 高速道路利用区間
    ws.merge_cells('J5:L5')
    ws['J5'] = '高速道路利用区間'
    ws['J5'].font = Font(name='ＭＳ 明朝', size=9)
    ws['J5'].alignment = Alignment(vertical='center')
    ws['J5'].border = thin_border
    
    ws.merge_cells('M5:N5')
    ws['M5'] = highway_from
    ws['M5'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['M5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['M5'].border = thin_border
    
    ws.merge_cells('P5:Q5')
    ws['P5'] = highway_to
    ws['P5'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['P5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['P5'].border = thin_border
    
    ws.row_dimensions[5].height = 24
    
    # 利用区間の片道料金行
    ws.merge_cells('J6:L6')
    ws['J6'] = '利用区間の片道料金\n（割引前）'
    ws['J6'].font = Font(name='ＭＳ 明朝', size=9)
    ws['J6'].alignment = Alignment(vertical='center')
    ws['J6'].border = thin_border
    
    ws.merge_cells('M6:P6')
    ws['M6'] = one_way_fee
    ws['M6'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['M6'].alignment = Alignment(vertical='center')
    ws['M6'].border = thin_border
    
    ws.row_dimensions[6].height = 24
    
    # １ヶ月の特別料金等加算額行
    ws.merge_cells('J7:L7')
    ws['J7'] = '１ヶ月の特別料金等加算額（認定額）'
    ws['J7'].font = Font(name='ＭＳ 明朝', size=8)
    ws['J7'].alignment = Alignment(vertical='center')
    ws['J7'].border = thin_border
    
    ws.merge_cells('M7:P7')
    ws['M7'] = f'=M6*42'  # 参考ファイルと同じ数式
    ws['M7'].font = Font(name='ＭＳ ゴシック', size=12)
    ws['M7'].alignment = Alignment(vertical='center')
    ws['M7'].border = thin_border
    
    ws.row_dimensions[7].height = 24
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 24
    ws.row_dimensions[10].height = 24
    
    # ヘッダー行の設定（11-12行目）
    ws.row_dimensions[11].height = 30
    ws.row_dimensions[12].height = 24.75
    
    # 左側のカラム（前半15日）
    ws.merge_cells('B11:B12')
    ws['B11'] = '日'
    ws['B11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['B11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['B11'].border = thin_border
    
    ws.merge_cells('C11:C12')
    ws['C11'] = '曜日'
    ws['C11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['C11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['C11'].border = thin_border
    
    ws.merge_cells('D11:F11')
    ws['D11'] = '往　路'
    ws['D11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['D11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['D11'].border = thin_border
    
    ws['D12'] = '利用確認'
    ws['D12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['D12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['D12'].border = thin_border
    
    ws['E12'] = '利用金額'
    ws['E12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['E12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['E12'].border = thin_border
    
    ws['F12'] = '確認'
    ws['F12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['F12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['F12'].border = thin_border
    
    ws.merge_cells('G11:I11')
    ws['G11'] = '復　路'
    ws['G11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['G11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['G11'].border = thin_border
    
    ws['G12'] = '利用確認'
    ws['G12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['G12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['G12'].border = thin_border
    
    ws['H12'] = '利用金額'
    ws['H12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['H12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['H12'].border = thin_border
    
    ws['I12'] = '確認'
    ws['I12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['I12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['I12'].border = thin_border
    
    # 右側のカラム（後半15日）
    ws.merge_cells('J11:J12')
    ws['J11'] = '日'
    ws['J11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['J11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['J11'].border = thin_border
    
    ws.merge_cells('K11:K12')
    ws['K11'] = '曜日'
    ws['K11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['K11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['K11'].border = thin_border
    
    ws.merge_cells('L11:N11')
    ws['L11'] = '往　路'
    ws['L11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['L11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['L11'].border = thin_border
    
    ws['L12'] = '利用確認'
    ws['L12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['L12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['L12'].border = thin_border
    
    ws['M12'] = '利用金額'
    ws['M12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['M12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['M12'].border = thin_border
    
    ws['N12'] = '確認'
    ws['N12'].font = Font(name='ＭＳ 明朝', size=10)
    ws['N12'].alignment = Alignment(horizontal='center', vertical='center')
    ws['N12'].border = thin_border
    
    ws.merge_cells('O11:Q11')
    ws['O11'] = '復　路'
    ws['O11'].font = Font(name='ＭＳ 明朝', size=10)
    ws['O11'].alignment = Alignment(horizontal='center', vertical='center')
    ws['O11'].border = thin_border
    
    # 月の初日を計算
    first_day = datetime(year, month, 1)
    last_day = datetime(year, month, calendar.monthrange(year, month)[1])
    
    # 入力欄のセル設定（E56, E57に月の初日と最終日を設定）
    ws['E56'] = first_day
    ws['E57'] = last_day
    
    # CSVデータとグリッドのマッチング
    grid_data = match_csv_to_grid(df, year, month, highway_from, highway_to, one_way_fee)
    
    # 日付と曜日の数式を設定（13-27行目は前半15日、28-42行目は後半15日相当）
    for row in range(13, 28):  # 前半15日
        ws.row_dimensions[row].height = 21
        
        # 日付数式（参考ファイルと同じ）
        if row == 13:
            ws[f'B{row}'] = '=$E$56'
        else:
            ws[f'B{row}'] = f'=IF(B{row-1}=$E$57,"-",IF(B{row-1}="-","-",B{row-1}+1))'
        ws[f'B{row}'].font = Font(name='ＭＳ Ｐ明朝', size=11)
        ws[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center')
        ws[f'B{row}'].border = thin_border
        
        # 曜日数式
        ws[f'C{row}'] = f'=IF(B{row}="-","-",TEXT(WEEKDAY(B{row}),"aaa"))'
        ws[f'C{row}'].font = Font(name='ＭＳ Ｐ明朝', size=11)
        ws[f'C{row}'].alignment = Alignment(horizontal='center', vertical='center')
        ws[f'C{row}'].border = thin_border
        
        day = row - 12  # 1日から開始
        
        # CSVデータから往復データを取得して設定
        if day in grid_data:
            # 往路データ
            ws[f'D{row}'] = grid_data[day]['outbound_confirmed']
            if grid_data[day]['outbound_amount']:
                ws[f'E{row}'] = grid_data[day]['outbound_amount']
            
            # 復路データ  
            ws[f'G{row}'] = grid_data[day]['return_confirmed']
            if grid_data[day]['return_amount']:
                ws[f'H{row}'] = grid_data[day]['return_amount']
        
        # セルのスタイル設定
        for col in ['D', 'E', 'F', 'G', 'H', 'I']:
            cell = ws[f'{col}{row}']
            cell.font = Font(name='ＭＳ ゴシック', size=9)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            if col in ['E', 'H']:  # 利用金額の列
                cell.number_format = '0'
            if col in ['F', 'I']:  # 確認の列
                if col == 'F':
                    cell.value = f'=IF(E{row}>$M$6,"×","")'
                else:
                    cell.value = f'=IF(H{row}>$M$6,"×","")'
        
        # 後半15日（右側）の設定
        right_day = day + 15
        if right_day <= calendar.monthrange(year, month)[1]:
            # 後半の日付数式
            if row == 13:
                ws[f'J{row}'] = f'=IF(B27=$E$57,"-",IF(B27="-","-",B27+1))'
            else:
                ws[f'J{row}'] = f'=IF(J{row-1}=$E$57,"-",IF(J{row-1}="-","-",J{row-1}+1))'
            
            ws[f'J{row}'].font = Font(name='ＭＳ Ｐ明朝', size=11)
            ws[f'J{row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'J{row}'].border = thin_border
            
            # 後半の曜日数式
            ws[f'K{row}'] = f'=IF(J{row}="-","-",TEXT(WEEKDAY(J{row}),"aaa"))'
            ws[f'K{row}'].font = Font(name='ＭＳ Ｐ明朝', size=11)
            ws[f'K{row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'K{row}'].border = thin_border
            
            # 後半のCSVデータを取得して設定
            if right_day in grid_data:
                # 往路データ
                ws[f'L{row}'] = grid_data[right_day]['outbound_confirmed']
                if grid_data[right_day]['outbound_amount']:
                    ws[f'M{row}'] = grid_data[right_day]['outbound_amount']
                
                # 復路データ
                ws[f'O{row}'] = grid_data[right_day]['return_confirmed']
                if grid_data[right_day]['return_amount']:
                    ws[f'P{row}'] = grid_data[right_day]['return_amount']
            
            # 後半のセルスタイル設定
            for col in ['L', 'M', 'N', 'O', 'P', 'Q']:
                cell = ws[f'{col}{row}']
                cell.font = Font(name='ＭＳ ゴシック', size=9)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
                if col in ['M', 'P']:  # 利用金額の列
                    cell.number_format = '0'
                if col in ['N', 'Q']:  # 確認の列
                    if col == 'N':
                        cell.value = f'=IF(M{row}>$M$6,"×","")'
                    else:
                        cell.value = f'=IF(P{row}>$M$6,"×","")'
    
    # 列幅設定（参考ファイルと同じ）
    column_widths = {
        'A': 6.77734375, 'B': 6.109375, 'C': 6.109375, 'D': 6.77734375,
        'E': 7.88671875, 'F': 5.77734375, 'G': 6.77734375, 'H': 7.88671875,
        'I': 5.88671875, 'J': 6.109375, 'K': 13.0, 'L': 6.77734375,
        'M': 7.88671875, 'N': 6.109375, 'O': 6.77734375, 'P': 7.88671875,
        'Q': 5.77734375, 'R': 6.77734375, 'S': 13.0
    }
    
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
    
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
                            wb = generate_expense_report(df, year, month, highway_from, highway_to, one_way_fee, monthly_allowance, organization, position, name)
                            
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
        st.subheader("📋 新機能")
        st.markdown("""
        **完全準拠の公式フォーマット**
        - 参考ファイルと同一レイアウト
        - 自動日付・曜日計算
        - CSV データの自動マッチング
        - Excel数式の完全再現
        """)

if __name__ == "__main__":
    main()