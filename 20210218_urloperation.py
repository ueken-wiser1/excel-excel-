
# coding: utf-8

import os
from openpyxl import Workbook
import requests
import bs4
import pandas as pd

import requests
import time
import datetime

import winsound


taisho_URL= 'http://kabutan.jp/stock/kabuka/?code=1301'

res = requests.get(taisho_URL)
#print(taisho_URL)
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, 'html.parser')



tr_tags = soup.find_all('tr')
num_tr_tags = int(len(tr_tags))
#print(num_tr_tags)

td_tags = soup.find_all('td')
num_td_tags = int(len(td_tags))
#print(num_td_tags)


a_tags = soup.find_all('a')
num_a_tags = int(len(a_tags))
print(num_a_tags)

#rows = soup.find_all('tr')

#print(rows)
'''
for row_index, row in enumerate(rows):
    cells = row.find_all(['td', 'th'])

    for cell_index, cell in enumerate(cells):
        if row_index>4:
            t = datetime.datetime.now().time()
            date = datetime.datetime.now().date()
            print(f"Row Index: {row_index}, Cell Index: {cell_index}, Text: {cell.text}")
            #print(cell_index)
            #print(date)
'''
'''
for i in range(1,num_tr_tags):
    print(i)
    td_elem = soup.select('tr')[i].get_text
    print(td_elem)
'''

            

for i in range(1,50):
    td_elem = soup.select('span')[i]
    print(i)
    print(td_elem)
#株価td-32，値下がり幅span-14，値下がり率span-15，出来高td-35

#    print(td_tags)
#winsound.Beep(1000,1000)  #ビープ音（800Hzの音を1000msec流す）