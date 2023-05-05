#対象ファイル：注目銘柄
#最終列だけを見ていくコード
#最終列の左側に数字がなければ、その行はスキップ
#最終列の左側に数字があった場合、その列の証券コードを読み込み
#当日の日次データで同じ証券コードをサーチ
#あったら、その行の終値を最終列に書込
#書き込んだ終値が左側の数字より大きければオレンジ塗り、小さければ青塗り
#最終列の数字とその行の一番左側にある数字の差を銘柄名の右側に書込
#→注目してからその銘柄は上がっているのか下がっているのか見極め
#その数字は注目した時の終値の何%かを書込
#注目してからその数字が2%を超えるまで何日かかったか、日付の差分の書込
#excel表をpowerpointに貼り付けて画像化
#更新単位は一日だから、注目銘柄リストは一日単位で作っていく
#そのため、注目銘柄リストのファイルは一日単位で作っていって、フォルダ内の各ファイルに対して触っていくという形にしよう
#注目銘柄リストは毎日オレが手動で作る必要がある
#計算シートを触ってテーブルにしてスコアで降順にする作業が入るため
#完全自動にできなくはないが、現状はそこを自動化しても旨みは強くない

#フォルダ内の注目銘柄リストexcelを開く
#読み取るべき数値は、証券コードと証券コード書かれた行、最終列と最終列最初の行に書かれた日付
#1. 注目銘柄リストフォルダ内のファイルをリスト化して、ファイルを一つずつ開いていく：注目リストファイル
#2. 注目リストファイルの最終行を取得
#3. 注目リストファイルの最終列を取得
#4. 注目リストファイルの最終列1行目に書かれた日付を取得
#5. 注目リストファイルの取得した日付をファイル名に合うよう変換→これはファイル作成時にそのように対応する？
#6. 注目リストファイルの銘柄行の注目完了フラグが立っているか確認する
#7. フラグが立っている場合、その行はスキップする
#8. 注目リストファイルの証券コード列の最初の行の証券コードを取得
#9. 取得した日付に対応するファイルを完了フォルダからサーチし開く：日次データファイル
#10. 日次データファイルに対して、取得した証券コードに対応する行を検索
#11. 日次データファイルの対応する行の終値を取得
#12. 注目リストファイルの対応する行の最終列に取得した終値を記載
#13. 注目リストファイルの記載した終値と注目開始時の終値を比較
#14. 比較割合が2%を越えたら、注目完了フラグを立てる
#15. 注目完了フラグが立った場合、その行は赤いハッチングをかける
#16. 注目リストファイルの注目完了フラグが立ったら、利益(最新記載終値-注目開始終値)を記載
#17. 注目リストファイルの最新日付の列がn列目であった場合、利損(最新記載終値-注目開始終値)を記載し、注目完了フラグを立てる←強制終了
#18. 強制終了した行の利益率が2%未満の場合、その行は青いハッチングをかける
#19. 注目リストファイルの次の銘柄に移る
#20. 注目リストファイルの最終行まで終わったら、次のファイルに移る
'''
folder_path: 注目銘柄リストがあるフォルダのパス
excel_file: 注目銘柄リストのExcelファイル名
stock_list_workbook: 注目銘柄リストのExcelワークブックオブジェクト
stock_list_worksheet: 注目銘柄リストのExcelワークシートオブジェクト
last_row: 最終行のインデックス
last_column: 最終列のインデックス
date_in_cell: 最終列1行目に書かれた日付
file_date: 取得した日付をファイル名に合うように変換したもの
completed_flag: 注目完了フラグ
security_code: 証券コード列の最初の行の証券コード
target_file: 取得した日付に対応するファイル
target_workbook: 対応するファイルのワークブックオブジェクト
target_worksheet: 対応するファイルのワークシートオブジェクト
target_row: 取得した証券コードに対応する行のインデックス
closing_price: 対応する行の終値
attention_start_closing_price: 注目開始時の終値
comparison_ratio: 比較割合（最新記載終値と注目開始終値の比率）
profit: 利益（最新記載終値 - 注目開始終値）
current_stock_row: 現在処理中の銘柄行のインデックス
n_columns: 最新日付の列がn列目である場合の列インデックス
force_close_profit_loss: 強制終了時の利損（最新記載終値 - 注目開始終値）
'''


import os
import glob
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def find_security_code_row(ws, code):
    for row2 in range(2, ws.max_row + 1):
        #print(code)
        if int(ws.cell(row=row2, column=2).value) == int(code):
            return row2
    return None

watch_list_folder = 'C:/Users/touko/OneDrive/株価分析/excel/株式データ/test/20230501/'
completed_folder = 'C:/Users/touko/OneDrive/株価分析/excel/株式データ/完了/'
n_columns = 28

plus_fill = PatternFill(patternType='solid', fgColor='ee82ee') #前日より値上がり
minus_fill = PatternFill(patternType='solid', fgColor='00bfff') #前日より値下がり
attained_fill = PatternFill(patternType='solid', fgColor='adff2f') #+2%目標を達成したら
unattained_fill = PatternFill(patternType='solid', fgColor='696969') #+2%目標未達で、5日目を迎えたら
for watch_list_file in glob.glob(os.path.join(watch_list_folder, '*.xlsx')):
    wb_watch_list = load_workbook(watch_list_file)
    ws_watch_list = wb_watch_list.active

    last_row = ws_watch_list.max_row
    #print('最終行は'+str(last_row))
    last_col = ws_watch_list.max_column
    #print('最終列は'+str(last_col))
    file_date = ws_watch_list.cell(row=1, column=last_col).value.strftime('%Y%m%d')
    #print('対象日付は'+str(file_date))

    for row in range(2, last_row + 1):
        completed_flag = ws_watch_list.cell(row=row, column=1).value
        if completed_flag:
            print(str(security_code)+'には完了フラグが立っています。')
            continue

        security_code = ws_watch_list.cell(row=row, column=4).value
        print(security_code)
        daily_data_file = os.path.join(completed_folder, f'{file_date}_allkabu1.xlsx')
        print(daily_data_file)
        print(watch_list_file)

        if os.path.exists(daily_data_file):
            wb_daily_data = load_workbook(daily_data_file)
            ws_daily_data = wb_daily_data.active

            security_code_row = find_security_code_row(ws_daily_data, security_code)
            print(security_code_row)

            if security_code_row:
                closing_price = ws_daily_data.cell(row=security_code_row, column=4).value
                ws_watch_list.cell(row=row, column=last_col).value = closing_price
                print(ws_watch_list.cell(row=row, column=last_col).value)
                start_closing_price = ws_watch_list.cell(row=row, column=6).value
                diff=ws_watch_list.cell(row=row, column=last_col).value - ws_watch_list.cell(row=row, column=last_col-1).value
                if diff > 0:
                    ws_watch_list.cell(row=row, column=last_col).fill = plus_fill
                elif diff < 0:
                    ws_watch_list.cell(row=row,column=last_col).fill = minus_fill

                if closing_price is not None and start_closing_price is not None:
                    ratio = ((closing_price - start_closing_price) / start_closing_price)*100
                    ws_watch_list.cell(row=row, column=8).value = ratio
                    ws_watch_list.cell(row=row, column=7).value = closing_price - start_closing_price
                    if ratio > 2:
                        ws_watch_list.cell(row=row, column=2).value = True
                        for cell in ws_watch_list.iter_rows(min_row=row, max_row=row,values_only=True):
                            for c in cell:
                                c.fill = attained_fill
                        profit = ws_watch_list.cell(row=row,column=last_col).value - ws_watch_list.cell(row=row,column=15).value
                        print(str(ws_watch_list.cell(row=row,column=13).value)+"_"+str(ws_watch_list.cell(row=row,column=14).value)) + "は2%目標を達成しました。利益は"+str(profit)+"円です。"

                    elif last_col == n_columns:
                        ws_watch_list.cell(row=row, column=2).value = False
                        for cell in ws_watch_list.iter_rows(min_row=row, max_row=row,values_only=True):
                            for c in cell:
                                c.fill = unattained_fill
                        loss = ws_watch_list.cell(row=row,column=last_col).value - ws_watch_list.cell(row=row,column=15).value
                        print(str(ws_watch_list.cell(row=row,column=13).value)+"_"+str(ws_watch_list.cell(row=row,column=14).value)) + "は目標を達成できず、終了しました。損失は"+str(loss)+"円です。"
            wb_daily_data.close()

    wb_watch_list.save(watch_list_file)
    print(watch_list_file+'を保存しました')
    wb_watch_list.close()
