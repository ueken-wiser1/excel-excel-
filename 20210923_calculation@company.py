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
        lastrow_stockbook = sheet01.max_row

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
#移動平均乖離率　5日＝（株価ー移動平均5日）/移動平均5日
                if j < 6 :
                        pass
                else:
                        sheet01.cell(row=j,column=204).value = str('=(D')+str(j)+str('-GS')+str(j)+str(')/GS')+str(j)
#移動平均乖離率　25日＝（株価ー移動平均25日）/移動平均25日
                if j < 26 :
                        pass
                else:
                        sheet01.cell(row=j,column=205).value = str('=(D')+str(j)+str('-GT')+str(j)+str(')/GT')+str(j)
#移動平均乖離率　75日＝（株価ー移動平均75日）/移動平均75日
                if j < 76 :
                        pass
                else:
                        sheet01.cell(row=j,column=206).value = str('=(D')+str(j)+str('-GU')+str(j)+str(')/GU')+str(j)


#平均出来高＝平均値（average(対象列)）
                sheet01.cell(row=j,column=251).value = str('=AVERAGE(K:K)')
#平均約定回数＝平均値（average(対象列)）
                sheet01.cell(row=j,column=252).value = str('=AVERAGE(H:H)')
#平均回転日数＝平均値（average(対象列)）
                sheet01.cell(row=j,column=253).value = str('=AVERAGE(DN:DN)')
#平均移動平均乖離率5日＝平均値（average(対象列)）
                sheet01.cell(row=j,column=254).value = str('=AVERAGE(GV:GV)')
#平均移動平均乖離率25日＝平均値（average(対象列)）
                sheet01.cell(row=j,column=255).value = str('=AVERAGE(GW:GW)')
#平均移動平均乖離率75日＝平均値（average(対象列)）
                sheet01.cell(row=j,column=256).value = str('=AVERAGE(GX:GX)')
#平均信用買残＝平均値（average(対象列)）
                sheet01.cell(row=j,column=257).value = str('=AVERAGE(W:W)')
#平均融資新規＝平均値（average(対象列)）
                sheet01.cell(row=j,column=258).value = str('=AVERAGE(DG:DG)')
#平均融資返済＝平均値（average(対象列)）
                sheet01.cell(row=j,column=259).value = str('=AVERAGE(DH:DH)')
#平均信用売残＝平均値（average(対象列)）
                sheet01.cell(row=j,column=260).value = str('=AVERAGE(V:V)')
#平均貸株新規＝平均値（average(対象列)）
                sheet01.cell(row=j,column=261).value = str('=AVERAGE(DJ:DJ)')
#平均貸株返済＝平均値（average(対象列)）
                sheet01.cell(row=j,column=262).value = str('=AVERAGE(DK:DK)')
#平均貸株超過＝平均値（average(対象列)）
                sheet01.cell(row=j,column=263).value = str('=AVERAGE(DL:DL)')
#平均出来高変化率＝平均値（average(対象列)）
                sheet01.cell(row=j,column=264).value = str('=AVERAGE(AB:AB)')
#平均約定回数変化率＝平均値（average(対象列)）
                sheet01.cell(row=j,column=265).value = str('=AVERAGE(AC:AC)')
#平均信用買残変化率＝平均値（average(対象列)）
                sheet01.cell(row=j,column=266).value = str('=AVERAGE(AA:AA)')
#平均信用売残変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=267).value = str('=AVERAGE(Z:Z)')
#平均平均約定金額＝平均値（average(対象列)）
                sheet01.cell(row=j,column=268).value = str('=AVERAGE(AG:AG)')
#標準偏差出来高＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=269).value = str('=STDEV.P(K:K)')
#標準偏差約定回数＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=270).value = str('=STDEV.P(H:H)')
#標準偏差回転日数＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=271).value = str('=STDEV.P(DN:DN)')
#標準偏差移動平均乖離率5日＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=272).value = str('=STDEV.P(GV:GV)')
#標準偏差移動平均乖離率25日＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=273).value = str('=STDEV.P(GW:GW)')
#標準偏差移動平均乖離率75日＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=274).value = str('=STDEV.P(GX:GX)')
#標準偏差信用買残＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=275).value = str('=STDEV.P(W:W)')
#標準偏差融資新規＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=276).value = str('=STDEV.P(DG:DG))')
#標準偏差融資返済＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=277).value = str('=STDEV.P(DH:DH)')
#標準偏差信用売残＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=278).value = str('=STDEV.P(V:V)')
#標準偏差貸株新規＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=279).value = str('=STDEV.P(DJ:DJ)')
#標準偏差貸株返済＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=280).value = str('=STDEV.P(DK:DK))')
#標準偏差貸株超過＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=281).value = str('=STDEV.P(DL:DL))')
#標準偏差出来高変化率＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=282).value = str('=STDEV.P(AB:AB)')
#標準偏差約定回数変化率＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=283).value = str('=STDEV.P(AC:AC)')
#標準偏差信用買残変化率＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=284).value = str('=STDEV.P(AA:AA)')
#標準偏差信用売残変化率＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=285).value = str('=STDEV.P(Z:Z))')
#標準偏差平均約定金額＝標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=286).value = str('=STDEV.P(AG:AG))')

#標準化出来高＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IQ')+str(j)+str(',JI')+str(j)+str(')')
#標準化約定回数＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IR')+str(j)+str(',JJ')+str(j)+str(')')
#標準化回転日数＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IS')+str(j)+str(',JK')+str(j)+str(')')
#標準化移動平均乖離率5日＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IT')+str(j)+str(',JL')+str(j)+str(')')
#標準化移動平均乖離率25日＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IU')+str(j)+str(',JM')+str(j)+str(')')
#標準化移動平均乖離率75日＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IV')+str(j)+str(',JN')+str(j)+str(')')
#標準化信用買残＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IW')+str(j)+str(',JO')+str(j)+str(')')
#標準化融資新規＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IX')+str(j)+str(',JP')+str(j)+str(')')
#標準化融資返済＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IY')+str(j)+str(',JQ')+str(j)+str(')')
#標準化信用売残＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',IZ')+str(j)+str(',JR')+str(j)+str(')')
#標準化貸株新規＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JA')+str(j)+str(',JS')+str(j)+str(')')
#標準化貸株返済＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JB')+str(j)+str(',JT')+str(j)+str(')')
#標準化貸株超過＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JC')+str(j)+str(',JU')+str(j)+str(')')
#標準化出来高変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JD')+str(j)+str(',JV')+str(j)+str(')')
#標準化約定回数変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JE')+str(j)+str(',JW')+str(j)+str(')')
#標準化信用買残変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JF')+str(j)+str(',JX')+str(j)+str(')')
#標準化信用売残変化率＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JG')+str(j)+str(',JY')+str(j)+str(')')
#標準化平均約定金額＝standardize（当日の値，平均値（average(対象列)），標準偏差(stdevp(対象列)）
                sheet01.cell(row=j,column=501).value = str('=STANDARDIZE(K')+str(j)+str(',JH')+str(j)+str(',JZ')+str(j)+str(')')
#売り時/買い時スコア＝12*標準化出来高(SG)+15*標準化約定回数(SH)+6*標準化平均約定金額(SX)+3/標準化回転日数(SI)+9*標準化移動平均乖離率5日(SJ)+6*標準化信用買残(SM)+10*信用規制(CW)+8*増担規制(CX)+4*標準化融資新規(SN)+2/標準化融資返済(SO)+4/標準化信用売残(SP)+5*空売り規制(DB)+3/標準化貸株新規(SQ)+2*標準化貸株返済(SR)+1/標準化貸株超過(SS)+(20*IR)
                sheet01.cell(row=j,column=1000).value = str('=12*SG')+str(j)+str('+15*SH')+str(j)+str('+6*SX')+str(j)+str('+3/(1+SI')+str(j)+str(')+9*SJ')+str(j)+str('+6*SM')+str(j)+str('+10*CW')+str(j)+str('+8*CX')+str(j)+str('+4*SN')+str(j)+str('+2/(1+SO')+str(j)+str(')+4/(1+SP')+str(j)+str(')+5*DB')+str(j)+str('+3/(1+SQ')+str(j)+str(')+2*SR')+str(j)+str('+1/(1+SS')+str(j)+str(')+20*AY')+str(j)

        stockbook.save(l)
        stockbook.close()
print(t)
t = datetime.datetime.now().time()
print(t)


winsound.Beep(500,50)  #ビープ音（500Hzの音を50msec流す）
