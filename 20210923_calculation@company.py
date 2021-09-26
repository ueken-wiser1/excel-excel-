import os
from re import L
import openpyxl
import requests
import bs4
import time
import datetime
import sys
import winsound
import glob

#どんな動きをさせるのか
#excelを開く
#参照excelの各シートに記載された証券コードを読み込む
#シート名は"株式", "マーケット", "為替", "投信"
#読み込んだ証券コードにより株探URLを作成し、銘柄ページへ移動
#必要な情報を取得する
t = datetime.datetime.now().time()
#excelを開く
#対象：信用規制
dir_stock = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/株式/"
#ダウンロードしたexcel(以下databook)を開く
stock_list = glob.glob(dir_stock + '*.xlsx')

for l in stock_list:
    stockbook = openpyxl.load_workbook(l)
#    print(str(stockbook))
    sheet01 = stockbook.worksheets[0]
    lastrow_stockbook = sheet01.max_row + 1
    for j in range(2,lastrow_stockbook+1):
        print(l)
#売買代金2=VWAP(T)*出来高(K)
        sheet01.cell(row=j,column=25).value = str('=T')+str(j)+str('*K')+str(j)
#信用売残前週比：(V)
        sheet01.cell(row=j,column=26).value = str('=V')+str(j)+str('*V')+str(j-1)
#信用買残前週比：(W)
        sheet01.cell(row=j,column=27).value = str('=W')+str(j)+str('*W')+str(j-1)
#出来高前日比：(K)
        sheet01.cell(row=j,column=28).value = str('=K')+str(j)+str('*K')+str(j-1)
#約定回数前日比：(H)
        sheet01.cell(row=j,column=29).value = str('=H')+str(j)+str('*H')+str(j-1)
#平均約定金額平均約定金額=出来高(K)/約定回数(H)*VWAP(T)
        sheet01.cell(row=j,column=33).value = str('=K')+str(j)+str('/H')+str(j)+str('*T')+str(j)
#回転日数=((融資残(DI)+貸株残(DL))*2)/(融資新規(DG)+融資返済(DH)+貸株新規(DJ)+貸株返済(DK)
        sheet01.cell(row=j,column=118).value = str('=(DI')+str(j)+str('+DL')+str(j)+str(')*2/(DG')+str(j)+str('+DH')+str(j)+str('+DJ')+str(j)+str('+DK')+str(j)+str(')')
#現在株価との差=株価(D)-みんかぶ目標株価(EU)
        sheet01.cell(row=j,column=152).value = str('=D')+str(j)+str('-EU')+str(j)
#移動平均線数値　5日＝当日〜4日前の株価総和/5
        if j < 6 :
                pass
        else:
                sheet01.cell(row=j,column=201).value = str('=SUM(D')+str(j-5)+str(':D')+str(j)+str(')/5')
#移動平均線数値　25日＝当日〜24日前の株価総和/25
        if j < 26 :
                pass
        else:
                sheet01.cell(row=j,column=202).value = str('=SUM(D')+str(j-25)+str(':D')+str(j)+str(')/25')
#移動平均線数値　75日＝当日〜74日前の株価総和/75
        if j < 76 :
                pass
        else:
                sheet01.cell(row=j,column=203).value = str('=SUM(D')+str(j-75)+str(':D')+str(j)+str(')/75')
#移動平均乖離率    5日＝（株価ー移動平均5日）/移動平均5日
        sheet01.cell(row=j,column=204).value = str('=(D')+str(j)+str('-GS')+str(j)+str(')/GS')+str(j)
#移動平均乖離率    25日＝（株価ー移動平均25日）/移動平均25日
        sheet01.cell(row=j,column=205).value = str('=(D')+str(j)+str('-GT')+str(j)+str(')/GT')+str(j)
#移動平均乖離率    75日＝（株価ー移動平均75日）/移動平均75日
        sheet01.cell(row=j,column=206).value = str('=(D')+str(j)+str('-GU')+str(j)+str(')/GU')+str(j)
#標準化出来高＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',AVERAGE(K:K),STDEV.P(K:K))')
#標準化約定回数＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=502).value = str('=STANDARDIZE(H')+str(j)+str(',AVERAGE(H:H),STDEV.P(H:H))')
#標準化回転日数＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=503).value = str('=STANDARDIZE(DN')+str(j)+str(',AVERAGE(DN:DN),STDEV.P(DN:DN))')
#標準化移動平均乖離率5日＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=504).value = str('=STANDARDIZE(GV')+str(j)+str(',AVERAGE(GV:GV),STDEV.P(GV:GV))')
#標準化移動平均乖離率25日＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=505).value = str('=STANDARDIZE(GW')+str(j)+str(',AVERAGE(GW:GW),STDEV.P(GW:GW))')
#標準化移動平均乖離率75日＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=506).value = str('=STANDARDIZE(GX')+str(j)+str(',AVERAGE(GX:GX),STDEV.P(GX:GX))')
#標準化信用買残＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=507).value = str('=STANDARDIZE(Y')+str(j)+str(',AVERAGE(Y:Y),STDEV.P(Y:Y))')
#標準化融資新規＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=508).value = str('=STANDARDIZE(DG')+str(j)+str(',AVERAGE(DG:DG),STDEV.P(DG:DG))')
#標準化融資返済＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=509).value = str('=STANDARDIZE(DH')+str(j)+str(',AVERAGE(DH:DH),STDEV.P(DH:DH))')
#標準化信用売残＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=510).value = str('=STANDARDIZE(X')+str(j)+str(',AVERAGE(X:X),STDEV.P(X:X))')
#標準化貸株新規＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=511).value = str('=STANDARDIZE(DJ')+str(j)+str(',AVERAGE(DJ:DJ),STDEV.P(DJ:DJ))')
#標準化貸株返済＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=512).value = str('=STANDARDIZE(DK')+str(j)+str(',AVERAGE(DK:DK),STDEV.P(DK:DK))')
#標準化貸株超過＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=513).value = str('=STANDARDIZE(DQ')+str(j)+str(',AVERAGE(DQ:DQ),STDEV.P(DQ:DQ))')
#標準化出来高変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=514).value = str('=STANDARDIZE(AB')+str(j)+str(',AVERAGE(AB:AB),STDEV.P(AB:AB))')
#標準化約定回数変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=515).value = str('=STANDARDIZE(AC')+str(j)+str(',AVERAGE(AC:AC),STDEV.P(AC:AC))')
#標準化信用買残変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=516).value = str('=STANDARDIZE(AA')+str(j)+str(',AVERAGE(AA:AA),STDEV.P(AA:AA))')
#標準化信用売残変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=517).value = str('=STANDARDIZE(Z')+str(j)+str(',AVERAGE(Z:Z),STDEV.P(Z:Z))')
#標準化平均約定金額＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
        sheet01.cell(row=j,column=518).value = str('=STANDARDIZE(AG')+str(j)+str(',AVERAGE(AG:AG),STDEV.P(AG:AG))')
#売り時/買い時スコア＝12*標準化出来高(SG)+15*標準化約定回数(SH)+6*標準化平均約定金額(SX)+3/標準化回転日数(SI)+9*標準化移動平均乖離率5日(SJ)+6*標準化信用買残(SM)+10*信用規制(CW)+8*増担規制(CX)+4*標準化融資新規(SN)+2/標準化融資返済(SO)+4/標準化信用売残(SP)+5*空売り規制(DB)+3/標準化貸株新規(SQ)+2*標準化貸株返済(SR)+1/標準化貸株超過(SS)+(20*IR)
        sheet01.cell(row=j,column=1000).value = str('=12*SG')+str(j)+str('+15*SH')+str(j)+str('+6*SX')+str(j)+str('+3/SI')+str(j)+str('+9*SJ')+str(j)+str('+6*SM')+str(j)+str('+10*CW')+str(j)+str('+8*CX')+str(j)+str('+4*SN')+str(j)+str('+2/SO')+str(j)+str('+4/SP')+str(j)+str('+5*DB')+str(j)+str('+3/SQ')+str(j)+str('+2*SR')+str(j)+str('+1/SS')+str(j)

        stockbook.save(l)
print(t)
t = datetime.datetime.now().time()
print(t)


winsound.Beep(1000,1000)  #ビープ音（800Hzの音を1000msec流す）
