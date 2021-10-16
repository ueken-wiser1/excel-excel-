
#融資貸株データのNoneをなくす関数
def none_to_result(input_sheet, row_input, column_input, result):
    if input_sheet.cell(row=row_input,column=column_input).value is None:
        input_sheet.cell(row=row_input,column=column_input).value = result
    else:
        pass

#信用残等前日比の関数
def shinyokeisan(inte, input_sheet, row_input, column_input, ):
    if inte == 2:
        input_sheet.cell(row=row_input,column=column_input).value = 0
    else:
        input_sheet.cell(row=row_input,column=column_input).value = input_sheet.cell(row=row_input,column=column_input).value - input_sheet.cell(row=row_input-1,column=column_input).value

#回転日数の関数
#もし参照セルの値の内、分母に来る値が全て0なら回転日数は0を返す
#これは一個しかないから、今はいらない

#配列を作る関数
#株価等の配列を作る
def array_making(array_result, inte01, input_sheet, lastrow, row_inte, column_inte):
    array_result = []
    for inte01 in range(2, lastrow):
        array_result.append(input_sheet.cell(row=row_inte,column=column_inte))

#移動平均線関係の数値を計算する関数
#配列の作成を付けること
def sum_partial(array_result, input_sheet, row_inte, column_inte, inte, inte01, inte02):
#array_result：配列の名称　input_sheet：使用するシート
#row_inte：参照する行　column_inte：参照する列　inte：計算範囲　inte01：配列を取り込む時の便宜文字
#inte02：移動平均を計算する時の便宜文字
    array_result = []
    lastrow = input_sheet.max_row
    for inte01 in range(2, lastrow):
        array_result.append(input_sheet.cell(row=row_inte,column=column_inte))
    if inte+1 < inte01:
        pass
    else:
        for inte02 in range(inte):
            input_sheet.cell(row=inte,column=column_inte).value = sum(array_result[inte02-inte:inte])
