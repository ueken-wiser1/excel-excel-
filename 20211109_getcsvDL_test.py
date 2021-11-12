import os
import openpyxl
import requests
import bs4
import time
import datetime
import re
import urllib.request
import urllib
from urllib.parse import urljoin
import urllib3

#------------お約束開始---冒頭
#稼働時間計測開始
import datetime
t = datetime.datetime.now().time()
#------------お約束終了---冒頭

#------------プログラム本文---ここから

today = datetime.date.today()
d = today.strftime('%Y%m%d')
dir = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/スクレイピング生データ/"
wb = openpyxl.load_workbook(dir + 'stockcodelist.xlsx')
name = wb.get_sheet_names

#"株式"
sheet = wb.get_sheet_by_name('株式')
t = datetime.datetime.now().time()

shinazan_URL = 'https://www.taisyaku.jp/search/result/index/1/'
res03 = requests.get(shinazan_URL)
print(shinazan_URL)
res03.raise_for_status()
soup03 = bs4.BeautifulSoup(res03.text, 'html.parser')
atags = soup03.find_all("a")
dtags_con_shina = soup03.select("a")[22]
dtags_con_kashi = soup03.select("a")[28]
url_shina = dtags_con_shina.get("href")
url_kashi = dtags_con_kashi.get("href")

savepoint03 = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/ダウンロードデータ/05.品貸料率/"+str(d)+"_shinakashi.xlsx"
savepoint04 = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/ダウンロードデータ/06.銘柄別融資・貸株残高一覧表/"+str(d)+"_kashikabu.xlsx"

csv_shina = urllib.request.urlopen(url_shina).read()
with open(savepoint03, mode="wb") as f:
    f.write(csv_shina)
csv_kashi = urllib.request.urlopen(url_kashi).read()
with open(savepoint04, mode="wb") as f:
    f.write(csv_kashi)
    
# shinazan_URL = 'https://www.taisyaku.jp/search/result/index/1/'
# res03 = requests.get(shinazan_URL)
# print(shinazan_URL)
# res03.raise_for_status()
# soup03 = bs4.BeautifulSoup(res03.text, 'html.parser')
# atags = soup03.find_all("a")
# print(len(atags))
# for i in range(1,len(atags)+1):
#     dtags_con = soup03.select("a")[i]
#     url = dtags_con.get("href")
#     print(i,dtags_con,url)




# savepoint02 = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/ダウンロードデータ/04.貸借取引銘柄別増担保金徴収措置一覧/"+str(d)+"mashitan.xlsx"
# xlsx = urllib.request.urlopen(url).read()

# with open(savepoint02, mode="wb") as f:
#     f.write(xlsx)
# karauri_file_base = 'https://www.jpx.co.jp/'
# karauri_file_url = karauri_file_base + url

# savepoint = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/ダウンロードデータ/03.空売り価格規制トリガー抵触銘柄一覧/"+str(d)+".csv"
# csv = urllib.request.urlopen(karauri_file_url).read()

# with open(savepoint, mode="wb") as f:
#     f.write(csv)



#------------プログラム本文---ここまで

#------------お約束開始---末尾
#稼働時間表示
print(t)
t = datetime.datetime.now().time()
print(t)

#稼働終了アナウンス
import winsound
# winsound.Beep(500,100)  #ビープ音（500Hzの音を50msec流す）
# winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
# winsound.Beep(500,100)  #ビープ音（500Hzの音を50msec流す）
# winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
# winsound.Beep(750,50)  #ビープ音（500Hzの音を50msec流す）
#------------お約束終了---末尾