import os
from re import L
import openpyxl
import time
import datetime
import sys
import winsound
import glob
import xlrd
import numpy as np
import pandas as pd

t = datetime.datetime.now().time()
d = datetime.date.today()

scorebook_dir = 'C:/Users/touko/OneDrive/株価分析/excel/株式データ/スコアブック/'
today_scorebook = scorebook_dir+str(d)+'score.xlsx'
df = pd.read_excel(today_scorebook, sheet_name='allscore', engine="openpyxl")

'''
ソートキー：売り時/買い時スコア　標準化出来高変化率　標準化約定回数変化率　時価総額2　PBR
'''

sortkey = "売り時/買い時スコア"
ascend = True

df_s = df.sort_values(sortkey, ascending=ascend)
print(df_s.iat[0, 2])
print(df_s.iat[1, 2])
print(df_s.iat[2, 2])
print(df_s.iat[3, 2])
print(df_s.iat[4, 2])
print(df_s.iat[5, 2])
print(df_s.iat[6, 2])
print(df_s.iat[7, 2])

sortkey = "PBR"
ascend = True

df_s = df.sort_values(sortkey, ascending=ascend)
print(df_s.iat[0, 2])
print(df_s.iat[1, 2])
print(df_s.iat[2, 2])
print(df_s.iat[3, 2])
print(df_s.iat[4, 2])
print(df_s.iat[5, 2])
print(df_s.iat[6, 2])
print(df_s.iat[7, 2])

winsound.Beep(500,50)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(1000,50)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(500,50)  #ビープ音（500Hzの音を50msec流す）
winsound.Beep(500,50)  #ビープ音（500Hzの音を50msec流す）