import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

def set_page():
    """ページの基本設定"""
    st.set_page_config(page_title="勤怠時間", layout="centered")
    st.title("勤怠時間")

def upload_files():
    """Streamlitにファイルアップローダーを設置し、勤怠関連ファイルのアップロード
       対応ファイル形式: CSV ('当期勤怠', '前期勤怠'), Excel ('総労働時間', '時間外労働')"""
    touki = st.file_uploader('当期勤怠ファイルの読み込み', type='csv', key='touki_csv')
    zenki = st.file_uploader('前期勤怠ファイルの読み込み', type='csv', key='zenki_csv')
    worktime = st.file_uploader('総労働時間ファイルの読み込み', type='xlsx', key='worktime_xlsx')
    overtime = st.file_uploader('時間外労働ファイルの読み込み', type='xlsx', key='overtime_xlsx')
    return touki, zenki, worktime, overtime

def load_csv_data(uploaded_file, usecols):
    """共通関数
        アップロードされたCSVファイルから指定された列を読み込み、データフレームを返却
       ファイルがNoneの場合や読み込みに失敗した場合はNoneを返却"""
    try:
        if uploaded_file is not None:
            df = pd.read_csv(uploaded_file, encoding='shift_jis', header=5, usecols=usecols)
            df = df[df.iloc[:, 0] != '合計']  # 合計行を除外
            return df
    except Exception as e:
        st.error(f'ファイル読み込みエラー: {e}')
        return None

def load_excel_data(uploaded_file):
    """共通関数
        アップロードされたExcelファイルから最初のシートを読み込み、workbookとsheetを返却
       ファイルがNoneの場合や読み込みに失敗した場合はNoneを返却"""
    try:
        if uploaded_file is not None:
            memory = BytesIO(uploaded_file.read())
            workbook = load_workbook(memory)
            sheet = workbook.worksheets[0] # 最初のシートを取得
            return workbook, sheet
    except Exception as e:
        st.error(f'ファイル読み込みエラー: {e}')
        return None, None

def update_excel_sheet(account_period, workbook, sheet, data, selected_month):
    """Excelファイルのシートを更新する共通関数
        指定された月に対応する列を更新し、Excelシートにデータを反映"""

    headers = [cell.value for cell in sheet[2]] # ヘッダー情報を取得
    month_col = headers.index(selected_month) + 1     # 更新する列番号を特定
    for i, row in enumerate(sheet.iter_rows(min_row=3, min_col=1, max_col=sheet.max_column), start=3):
        employee_cd = row[0].value
        if employee_cd in data:
            # 会社員CDと担当者CDが一致する場合
            # 前期の場合は同一行、当期の場合は一行下
            row_index = i if account_period == "前期" else i + 1
            sheet.cell(row=row_index, column=month_col).value = data[employee_cd]

def download_updated_file(workbook, file_name):
    """ダウンロードボタンを設置する共通関数
        更新したExcelファイルをダウンロード用に準備"""
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    st.download_button(label=f'更新された{file_name}ファイルをダウンロード',
                       data=output,
                       file_name=f'更新された{file_name}.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# クリアボタンの処理
def handle_clear():
    """クリアボタンの動作でファイルの準備状態をリセット"""
    st.session_state.file_ready = False
    st.experimental_rerun()
    
def main():
    set_page()
    selected_month = st.selectbox('月を選択してください', [f'{i}月' for i in range(1, 13)])
    touki, zenki, worktime, overtime = upload_files()

    if 'file_ready' not in st.session_state:
        st.session_state.file_ready = False
        
    if st.button('実行') and not st.session_state.file_ready:
        if not touki and not zenki:
            st.error('当期勤怠ファイル、前期勤怠ファイルを選択してください。')
            return
        if not worktime and not overtime:
            st.error('総労働時間ファイル、時間外労働ファイルを選択してください。')
            return

        st.session_state.file_ready = True
        st.experimental_rerun()

    if st.session_state.file_ready:
        # データの読み込みとシートの更新
        process_worksheets(touki, zenki, worktime, overtime, selected_month)
        st.button('クリア', on_click=handle_clear)

def process_worksheets(touki, zenki, worktime, overtime, selected_month):
    """ファイルからデータを読み込み、Excelシートを更新後、ダウンロードします。"""
    df_touki = load_csv_data(touki, [0, 11, 12])  # '社員CD', '残業','総労働時間'列
    df_zenki = load_csv_data(zenki, [0, 11, 12])  # '社員CD', '残業','総労働時間'列
    workbook_worktime, sheet_worktime = load_excel_data(worktime)
    workbook_overtime, sheet_overtime = load_excel_data(overtime)

    # 総労働時間の更新とダウンロード
    if workbook_worktime:
        update_and_download(workbook_worktime, sheet_worktime, df_touki, df_zenki, "総労働時間", selected_month)
    # 時間外労働の更新とダウンロード
    if workbook_overtime:
        update_and_download(workbook_overtime, sheet_overtime, df_touki, df_zenki, "残業", selected_month)

def update_and_download(workbook, sheet, df_touki, df_zenki, file_label, month):
    """データフレームからデータを抽出し、シートを更新してファイルをダウンロードします。"""
    if df_touki is not None:
        touki_data = df_touki.set_index('社員CD')[file_label].to_dict()
        update_excel_sheet("当期", workbook, sheet, touki_data, month)
    if df_zenki is not None:
        zenki_data = df_zenki.set_index('社員CD')[file_label].to_dict()
        update_excel_sheet("前期", workbook, sheet, zenki_data, month)
    download_updated_file(workbook, file_label)

if __name__ == "__main__":
    main()