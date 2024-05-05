import requests
import bs4
import os
import openpyxl
import time
import datetime
import urllib


#どんな動きをさせるのか
#株探URLに読み込んだ銘柄コードを組み合わせて銘柄ページに移動
#行【銘柄コード】に数値があれば、値をURLに組み込む
#       print(code_row)
kabutan_URL = 'https://kabutan.jp/stock/?code=1605'
kabutankessan_URL = 'https://kabutan.jp/stock/?code=1605'

#銘柄ページから情報を読み込む
res2 = requests.get(kabutankessan_URL)
print(kabutan_URL)
res2.raise_for_status()
soup2 = bs4.BeautifulSoup(res2.text, 'html.parser')

#td_elem = soup.select('span')[1]
#print(td_elem)

#d_today = datetime.date.today()
#print(d_today)

#a = str(d_today) in str(td_elem)
#print(a)
#欲しい情報：<span class="market"><a href="/themes/?industry=25&market=2">
#tdtags=soup.find_all('td')       #全aタグ取得
#print('tdタグ数：', len(tdtags))  #aタグ数取得
tags = soup2.find_all('li')       #全aタグ取得
#print('tdタグ数：', len(tdtags))  #aタグ数取得a53 td25
num_tags = int(len(tags))
print(num_tags)

for i in range(num_tags):
    #print(i)
#    print(d_today)
    td_elem = soup2.select('li')[i]
    tag_content = td_elem.get_text()
    theme = "theme"
    if theme in str(td_elem):
        print(tag_content)
        #print(str(td_elem))
#    td_elem2 = soup.select('td')[i+2].get_text()
#    print(td_elem)

#    if str(d_today) in str(td_elem):
#        print(td_elem2)





