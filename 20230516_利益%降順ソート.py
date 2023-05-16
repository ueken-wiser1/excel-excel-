import pandas as pd
from openpyxl import load_workbook
import datetime

t = datetime.datetime.now().time()
file = 'C:/Users/touko/OneDrive/株価分析/excel/株式データ/test/20230516/20230216_OSCI.xlsx'
# Excelファイルを読み込みます
df = pd.read_excel(file)

# 並べ替えたい列の列番号（0から始まるインデックス）を指定します
column_index = 25  # 例えば、3列目を指定する場合

# 列番号を使用してデータフレームを並べ替えます
sorted_df = df.sort_values(by=df.columns[column_index], ascending=False)
sorted_df.to_excel(file, index=False)
# 並べ替えた結果を表示します
print(sorted_df)
wb_watch_list = load_workbook(file)
ws_watch_list = wb_watch_list.active
last_row_watch = ws_watch_list.max_row

for i in range(2, last_row_watch+1):
    ws_watch_list.cell(i, 26).value = (last_row_watch-i)*100/last_row_watch

wb_watch_list.save(file)
wb_watch_list.close()

print(t)
t = datetime.datetime.now().time()
print(t)

#稼働終了アナウンス
import winsound
winsound.Beep(500,100)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(500,100)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
#------------お約束終了---末尾