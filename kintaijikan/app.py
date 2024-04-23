import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def process_execution(self):
    
    return

def main():
    # Streamlitアプリの基本設定、ログイン前は中央配置
    st.set_page_config(page_title="勤怠時間", layout="centered")
    st.title("勤怠時間")
    #st.header("Upload CSV for Calculation")
    
    # 月の選択肢を設定
    st.text('出力する設定月')
    months = ['１月', '２月', '３月', '４月', '５月', '６月', '７月', '８月', '９月', '１０月', '１１月', '１２月']
    selected_month = st.selectbox('月を選択してください', months)
    
#     st.text('当期勤怠ファイルの読み込み')
#     touki_uploaded_file = st.file_uploader('file of Excel', type='xlsx', key='touki_xlsx')
#     if not touki_uploaded_file:
#         st.info("Excelファイルを選択してください")
# #        st.stop()
        
#     st.text('前期勤怠ファイルの読み込み')
#     zenki_uploaded_file = st.file_uploader('file of Excel', type='xlsx', key='zenki_xlsx')
#     if not zenki_uploaded_file:
#         st.info("Excelファイルを選択してください")
# #        st.stop()

#     st.text('総労働時間ファイルの読み込み')
#     worktime_uploaded_file = st.file_uploader('file of Excel', type='xlsx', key='worktime_xlsx')
#     if not worktime_uploaded_file:
#         st.info("Excelファイルを選択してください")
# #        st.stop()

#     st.text('時間外労働ファイルの読み込み')
#     overtime_uploaded_file = st.file_uploader('file of Excel', type='xlsx', key='overtime_lsx')
    
#     if not overtime_uploaded_file:
#         st.info("Excelファイルを選択してください")
# #        st.stop()
    # Excelファイルアップローダー
    #touki_uploaded_file = st.file_uploader('当期勤怠ファイルの読み込み', type='xlsx', key='touki_xlsx')
    #zenki_uploaded_file = st.file_uploader('前期勤怠ファイルの読み込み', type='xlsx', key='zenki_xlsx')
    touki_uploaded_file = st.file_uploader('当期勤怠ファイルの読み込み', type='csv', key='touki_csv')
    zenki_uploaded_file = st.file_uploader('前期勤怠ファイルの読み込み', type='csv', key='zenki_csv')
    worktime_uploaded_file = st.file_uploader('総労働時間ファイルの読み込み', type='xlsx', key='worktime_xlsx')
    overtime_uploaded_file = st.file_uploader('時間外労働ファイルの読み込み', type='xlsx', key='overtime_xlsx')


    if st.button('実行'):
        if not touki_uploaded_file and not zenki_uploaded_file :
            # エラーメッセージ
            st.error('当期勤怠ファイル、前期勤怠ファイルを選択してください。')
            return
        
        if not worktime_uploaded_file and not overtime_uploaded_file :
            # エラーメッセージ
            st.error('総労働時間ファイル、時間外労働ファイルを選択してください。')
            return
            
        #process_execution()
        # ファイルAの内容を読み込み、7行目までをスキップしてデータを取得
        # # 当期ファイル(Excel)
        # # 6行目と7行目を見出しとして使用
        # df_touki = pd.read_excel(touki_uploaded_file, header=[5, 6], skiprows=5)
        # # 空欄行と'合計'行を除外し、必要な列のみ取得
        # df_touki = df_touki[df_touki.iloc[:, 0] != '合計']
        # df_touki = df_touki.dropna(subset=[df_touki.columns[0]])  # 1列目に空欄がある行を削除
        # df_touki = df_touki[[df_touki.columns[0], df_touki.columns[12]]]  # 1列目(社員CD)、13列目(総労働時間)のみ取得
        
        # 当期ファイル(CSV)読み込み
        if touki_uploaded_file is not None:
            df_touki = pd.read_csv(touki_uploaded_file,encoding='shift_jis', header=5, usecols=[0, 12])  # 1列目と13列目のみ取得
            df_touki = df_touki[df_touki.iloc[:, 0] != '合計']  # 1列目に「合計」が含まれる行を除外

        
        # 総労働時間ファイル
        # 2行目と3行目を見出しとして使用
        # df_worktime = pd.read_excel(worktime_uploaded_file, sheet_name='22-23期比較', header=[1], skiprows=2)
        # # 必要な列のインデックスを作成（Pythonでは0からカウントが始まるため、1を引く）
        # columns_needed = [0] + list(range(2, 15))  # 0は1列目、2は3列目、14は15列目に対応
        # # 選択した列のみを抽出
        # df_worktime = df_worktime.iloc[:, columns_needed]
        # # 空欄行と'999999'行を除外し、必要な列のみ取得
        # df_worktime = df_worktime[df_worktime.iloc[:, 0] != '999999'] 
        # df_worktime = df_worktime.dropna(subset=[df_worktime.columns[0]])  # 1列目に空欄がある行を削除
        
        
        # ファイルを読み込む（ヘッダが2行目のみ、デフォルトで最初のシートを読み込む）
        # 選択する列は1列目、3列目から15列目まで
        columns_needed = [0] + list(range(2, 15))
        # 1列目を文字列として読み込む
        dtype_spec = {0: str}
        df_worktime = pd.read_excel(worktime_uploaded_file,sheet_name=0, header=1, usecols=columns_needed, dtype=dtype_spec)
        # 1列目の担当者CDの行結合対応（結合されている場合、最初の行のみ値が入っている、残りの行は欠損値(NaN)）
        # 1列目の担当者CDが NaN の場合、直前の行の値で埋める
        df_worktime.iloc[:, 0] = df_worktime.iloc[:, 0].fillna(method='ffill')
        # 1列目が '999999' や空欄ではない行のみを取得
        df_worktime = df_worktime[(df_worktime.iloc[:, 0] != '999999') & (df_worktime.iloc[:, 0].notna())]


        # ファイルBの列見出しを探索して、選択した月の列インデックスを取得
        month_index = df_worktime.columns.get_loc(selected_month)

        # 当期ファイルを元に総労働時間ファイルを更新
        for index, row in df_touki.iterrows():
            employee_cd = row['社員CD']  # 社員CD
            total_hours = row['総労働時間']  # 総労働時間

            # 担当者CDが一致し、期が'当期'である行を見つける
            for _, worktime_row in df_worktime.iterrows():
                if worktime_row['担当者CD'] == employee_cd and worktime_row['期'] == '当期':
                    df_worktime.loc[_, selected_month] = total_hours  # 指定された月の列に時間を更新

        # 更新されたデータフレームをExcelファイルに変換し、ダウンロード
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_worktime.to_excel(writer, index=False)
            # writer.save()  # この行は不要で、削除する

        output.seek(0)

        # ダウンロード用のファイル名を定義
        file_name = '23期月別個人別時間外労働比較3月.xlsx'

        # ダウンロードボタンをStreamlitに表示
        st.download_button(
            label="更新されたファイルをダウンロード",
            data=output,
            file_name=file_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == "__main__":
    main()