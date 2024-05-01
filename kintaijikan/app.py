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
    # ラジオボタンで使用する選択肢のリスト
    options = ['総労働時間', '残業']
    selected_option = st.radio("設定するファイル種別を選択してください",options, index=0)

    output_file = st.file_uploader(f'{selected_option}ファイルの読み込み', type='xlsx', key='output_xlsx')

    return touki, zenki, selected_option, output_file

def load_csv_data(uploaded_file, usecols):
    """共通関数
        アップロードされたCSVファイルから指定された列を読み込み、データフレームを返却
       ファイルがNoneの場合や読み込みに失敗した場合はNoneを返却"""
    try:
        if uploaded_file is not None:
            df = pd.read_csv(uploaded_file, encoding='shift_jis', header=5, usecols=usecols)
            df = df[df.iloc[:, 0] != '合計']  # 合計行を除外
            return df
        else:
            return None
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
        else :
            return None, None
    except Exception as e:
        st.error(f'ファイル読み込みエラー: {e}')
        return None, None

def update_excel_sheet(account_period, selected_option, sheet, df, selected_month):
    """Excelファイルのシートを更新する共通関数
        指定された月に対応する列を更新し、Excelシートにデータを反映"""
    data = df.set_index('社員CD')[selected_option].to_dict()
        
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

def main():
    set_page()
    selected_month = st.selectbox('月を選択してください', [f'{i}月' for i in range(1, 13)])
    touki, zenki, selected_option, output_file = upload_files()

    if st.button('実行'):

        # 入力チェック
        if not touki and not zenki:
            st.error('当期勤怠ファイル、前期勤怠ファイルを選択してください。')
            return
        if not output_file:
            st.error(f'{selected_option}ファイルを選択してください。')
            return
        
        process_files(touki, zenki, selected_option, output_file, selected_month)


def process_files(touki, zenki,  selected_option, output_file, selected_month):
    """ファイルからデータを読み込み、適切に処理し、ダウンロードを提供します。"""
    file_name = None
    if selected_option == "総労働時間":
        # 社員CD, 総労働時間列を読み込む
        df_touki = load_csv_data(touki, [0, 12])
        df_zenki = load_csv_data(zenki, [0, 12])
    elif selected_option == "残業":
        # 社員CD, 残業時間列を読み込む
        df_touki = load_csv_data(touki, [0, 11])
        df_zenki = load_csv_data(zenki, [0, 11])

    workbook, sheet = load_excel_data(output_file)
    update_and_download(workbook, sheet, df_touki, df_zenki, selected_option, selected_month)



def update_and_download(workbook, sheet, df_touki, df_zenki, selected_option, month):
    """データフレームからデータを抽出し、シートを更新してファイルをダウンロードします。"""
    if df_touki is not None:
        update_excel_sheet("当期", selected_option, sheet, df_touki, month)
    if df_zenki is not None:
        update_excel_sheet("前期", selected_option, sheet, df_zenki, month)

    download_updated_file(workbook, selected_option)

if __name__ == "__main__":
    main()