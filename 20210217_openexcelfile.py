"準備：サードパーティのモジュールを事前にインストールしておくこと"
import os
"openpyxlモジュールをインポート"
import openpyxl
import requests

"Excel文書を読み込む"
"文書の場所はカレントディレクトリ"
wb = openpyxl.load_workbook('kabu.xlsx')

"シートを取得する"
wb.get_sheet_names()
sheet =wb.get_sheet_by_name('2月配当')
sheet
type(sheet)
print(sheet.title)

a = sheet['A4'].value

print(a)



kabu = requests.get('http://kabutan.jp/stock/?code=2292')

print(len(kabu))

print(a)

print()

type(wb)

print()