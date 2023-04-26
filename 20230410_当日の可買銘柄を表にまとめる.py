import imp
import os
from re import L
from turtle import color
import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import numpy as np
import bottleneck as bn
import requests
import bs4
import time
import datetime
import glob
import re
import sys
import winsound

#開始時間取得
t = datetime.datetime.now()
d = datetime.datetime.now()
d1 = d.strftime('%Y%m%d')
#開始時間取得

#対象フォルダ指定
dirdaily = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/"
dirstorage = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/test/20230413/"

file_list = glob.glob(dirdaily + '*.xlsx')

for l in file_list:
    #日付データを順番に開く
    wb_daily = openpyxl.load_workbook(l)
    wb_sim = openpyxl.Workbook()
    sheetsim = wb_sim.worksheets[0]
    print(l)
    sheetdaily = wb_daily.worksheets[0]
    lastrow = sheetdaily.max_row+1
    lastcolumn = sheetdaily.max_column
    daycode = sheetdaily.cell(2,1).value
    daycode_format = daycode.strftime('%Y%m%d')

    sheetsim.cell(1,1).value = daycode
    sheetsim.cell(2,1).value = "証券コード"
    sheetsim.cell(2,2).value = "会社名"
    sheetsim.cell(2,3).value = "株価"
    sheetsim.cell(2,4).value = "前日比"
    sheetsim.cell(2,5).value = "5日線比率"
    sheetsim.cell(2,6).value = "25日線比率"
    sheetsim.cell(2,7).value = "75日線比率"
    sheetsim.cell(2,8).value = "利回り"
    sheetsim.cell(2,9).value = "PER"
    sheetsim.cell(2,10).value = "PBR"
    sheetsim.cell(2,11).value = "RSIスコア"
    sheetsim.cell(2,12).value = "ボリンジャーバンドスコア"
    sheetsim.cell(2,13).value = "MACDスコア"
    sheetsim.cell(2,14).value = "テクニカルスコア"
    k=3
    for i in range(2, lastrow):
        if sheetdaily.cell(i,4).value is None:
            pass
        else:
            if sheetdaily.cell(i,199).value is not None and sheetdaily.cell(i,200).value is not None and sheetdaily.cell(i,201).value is not None:
                if sheetdaily.cell(i,4).value<1000 and sheetdaily.cell(i,229).value == 1 and sheetdaily.cell(i,231).value == 1:
                    sheetsim.cell(k,1).value = sheetdaily.cell(i,2).value #証券コード
                    sheetsim.cell(k,2).value = sheetdaily.cell(i,3).value #会社名
                    sheetsim.cell(k,3).value = sheetdaily.cell(i,4).value #株価
                    sheetsim.cell(k,4).value = sheetdaily.cell(i,5).value #前日比
                    sheetsim.cell(k,5).value = 100-100*(sheetdaily.cell(i,4).value/sheetdaily.cell(i,199).value)
                    sheetsim.cell(k,6).value = 100-100*(sheetdaily.cell(i,4).value/sheetdaily.cell(i,200).value)
                    sheetsim.cell(k,7).value = 100-100*(sheetdaily.cell(i,4).value/sheetdaily.cell(i,201).value)
                    sheetsim.cell(k,8).value = sheetdaily.cell(i,61).value
                    sheetsim.cell(k,9).value = sheetdaily.cell(i,17).value #PER
                    sheetsim.cell(k,10).value = sheetdaily.cell(i,18).value #PBR
                    sheetsim.cell(k,11).value = sheetdaily.cell(i,324).value #RSIスコア
                    sheetsim.cell(k,12).value = sheetdaily.cell(i,325).value #ボリンジャーバンドスコア
                    sheetsim.cell(k,13).value = sheetdaily.cell(i,326).value #MACDスコア
                    sheetsim.cell(k,14).value = sheetdaily.cell(i,327).value #テクニカルスコア
                    print(sheetdaily.cell(i,2).value + '_' + sheetdaily.cell(i,3).value)
                    k += 1
            else:
                pass

    wb_sim.save(dirstorage+daycode_format+'_'+'OSCI.xlsx')
#終了時間取得-経過時間
print(t)
t1 = datetime.datetime.now()
print(t1)
dt = t1-t
print(dt)
#終了時間取得-経過時間

#稼働終了アナウンス
winsound.Beep(500,100)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(500,100)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
#------------お約束終了---末尾