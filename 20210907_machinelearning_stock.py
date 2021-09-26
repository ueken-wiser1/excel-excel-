import datetime
import pandas as pd
import matplotlib as plt
from sklearn.linear_model import LinearRegression

def readtempdata(filename): #気象データを読み込み
    df = pd.read_csv(filename,encoding="shift_jis",skiprows=6)
    df.columns = ["日付", "気温", "品質情報", "均質番号"]
    df['日付'] = pd.to_datetime(df['日付']) #曜日の列を追加 月曜日=0 日曜日=6
    df["曜日"] = df["日付"].dt.dayofweek
    df["月"] = df["日付"].dt.month
    df["日"] = df["日付"].dt.day
    return df

def readsalesdata(filename): #売上データを読み込み
    df = pd.read_excel(filename)
    df['日付'] = pd.to_datetime(df['日付'])
    return df

#気象データを読み込み
dfTemp = readtempdata("気象データ2017.csv") #売上データを読み込み
dfshop = readsalesdata("洋菓子店売上リスト2017.xlsx")

#気象データと売上データを統合
df = dfTemp.copy()
df = df.merge(dfshop,how="inner",on="日付")
x = df["月", "日", "曜日", "気温"] #与えるデータ
y = df["売上金額"] #求めるデータ

#重回帰分析
model = LinearRegression()
model.fit(x,y) #訓練の開始

#気象データを読み込み pg36_02
df = readtempdata("気象データ2018.csv")
x = df["月", "日", "曜日", "気温"] #与えるデータ
df["予測値"] = model.predict(x) #予測値の追加
with pd.ExcelWriter("洋菓子店2018年売上予測.xlsx",date_format='YYYY/MM/DD',datetime_format='YYYY/MM/DD') as writer:
    df.to_excel(writer,index=False) #シートを書き出し
