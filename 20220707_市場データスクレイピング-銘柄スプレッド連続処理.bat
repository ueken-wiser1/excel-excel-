python C:\Users\touko\program\github\excel-excel-\excel-excel-/20211108_marketscraping.py

echo off
 
cscript "C:\Users\touko\program\github\excel-excel-\excel-excel-\allkabu1タグ削除.vbs"

python C:\Users\touko\program\github\excel-excel-\excel-excel-/20211112_getnewsinfo_onlytoday.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20210905_stocknewssearch.py

python C:\Users\touko\program\github\excel-excel-\excel-excel-/20211109_getcsvDL_karauri.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-\preprocess-marketdata/20211108_regulation.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-\preprocess-marketdata/20211109_karauri.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-\preprocess-marketdata/20211109_mashitan.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-\preprocess-marketdata/20211109_shinakashi.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-\preprocess-marketdata/20211109_yushikashikabu.py

python C:\Users\touko\program\github\excel-excel-\excel-excel-/20230412_銘柄スプレッド.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20230420_日次データ計算.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20221026_日付データ回帰.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20230115_5日線超えた銘柄リスト.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20230120_5日線下回った銘柄リスト.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20230410_当日の可買銘柄を表にまとめる.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20230417_売り時通知.py
python C:\Users\touko\program\github\excel-excel-\excel-excel-/20221026_5日線超えた銘柄リスト.py
pause