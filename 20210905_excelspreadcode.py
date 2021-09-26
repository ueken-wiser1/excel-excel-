import os
from re import L
import openpyxl
import pandas
import requests
import bs4
import time
import datetime
import glob
import re
import sys
import winsound

#どんな動きをさせるのか
#1.     フォルダ-市場を規定する
#2.     フォルダ-銘柄を規定する
#3.     excel-市場を開く
#4.     excel-市場のn行目をコピーする
#5.     excel-市場のn行目に書かれた証券コードを読み込む
#6.     証券コードと同じ数字をファイル名に含むexcel-銘柄をフォルダ-Bから選択して開く
#7.     excel-銘柄の一番最後の行にコピーをペーストする。
#8.     excel-銘柄を閉じる
#9.     3に戻る
#10.    excel-市場の全てのデータをコピーしたら、excel-市場を閉じる
#11.    2に戻る
#12.    フォルダ-Aの全てのファイルを走査したら、プログラムを終了する

#プログラム
#1.     フォルダ-市場を規定する
#       pyファイルの直下にmarketフォルダを作成
#2.     フォルダ-銘柄を規定する
#       pyファイルの直下にcompanyフォルダを作成
#3.     excel-市場を開く

dir01 = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/"
dir02 = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/株式/"
#print(dir01)
file_list = glob.glob(dir01 + '*.xlsx')
stock_list = glob.glob(dir02 + '*.xlsx')
#print(file_list)
#print(stock_list)
#print(file_list)
#print(os.path.split(file_list[1]))
#stock_list=glob.glob(dir01 + '/*.xlsx')
'''
name_list=[]
for i in stock_list:
        file = os.path.basename(i)
        name = os.path.split(file)
        name_list.append(name)
'''
t = datetime.datetime.now().time()
#print(i)
#print(name_list)

for l in file_list:
#フォルダ-市場内のexcelを開く
    wb_market = openpyxl.load_workbook(l) 
    print(wb_market)
    sheet01 = wb_market.worksheets[0]
#4.     excel-市場のn行目をコピーする
#    j = 1
#    print(j)
    for j in range(2, sheet01.max_row + 1):
#5.     excel-市場のn行目に書かれた証券コードを読み込む
        stock_code = sheet01.cell(row=j, column=2).value
#        print(stock_code)
        #6.     証券コードと同じ数字をファイル名に含むexcel-銘柄をフォルダ-Bから選択して開く
#       ストックコードと同じコードをファイル名に含むファイルを検索する
#       検索結果はリスト形式。リストの一番目を開く形にする
#                company_book = stock_code.find in stock_list
        book_search_list = glob.glob(dir02 + str(stock_code) + '*.xlsx')
#                print(book_search_list)
        company_book = book_search_list[0]
        print(company_book)
        wb_company = openpyxl.load_workbook(company_book) #フォルダ-銘柄のexcelを開く
        sheet02 = wb_company.worksheets[0]
        last_row = sheet02.max_row
        last_column = sheet02.max_column
#        print(last_row)
        for k in range(1, sheet02.max_column):
#                file_path = os.path(dir02 + stock_code + '*')
#                print(file_path)
                row_copy = sheet01.cell(row=j, column=k).value
#                print(row_copy)

#                print(last_row)
#7.     excel-銘柄の一番最後の行にコピーをペーストする。

                sheet02.cell(row = last_row+1, column=k, value=row_copy)
#                wb_company.save(company_book)
                k += 1
        wb_company.save(company_book)
        j += 1

#8.     excel-銘柄を閉じる
        wb_company.close()
#10.    excel-市場の全てのデータをコピーしたら、excel-市場を閉じる
    wb_market.close()

#12.    フォルダ-Aの全てのファイルを走査したら、プログラムを終了する
print(t)
t = datetime.datetime.now().time()
print(t)

winsound.Beep(500,50)  #ビープ音（500Hzの音を50msec流す）