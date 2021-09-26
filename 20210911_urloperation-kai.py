
# coding: utf-8

import os
import openpyxl
import requests
import bs4
import time
import datetime
import sys
import sys
import codecs
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys as keys
import winsound


taisho_URL= 'https://www.jpx.co.jp/markets/equities/margin-reg/index.html'

res = requests.get(taisho_URL)
print(taisho_URL)
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, 'html.parser')


tr_tags = soup.find_all('tr')
num_tr_tags = int(len(tr_tags))
print(num_tr_tags)

td_tags = soup.find_all('td')
num_td_tags = int(len(td_tags))
print(num_td_tags)

'''
a_tags = soup.find_all('a')
num_a_tags = int(len(a_tags))
print(num_a_tags)
'''
for i in range(1,num_td_tags):
    print(i)
    td_elem = soup.select('td')[i].get_text
    print(td_elem)

for i in range(1,num_td_tags+1):
    
    if i < num_td_tags:
        if i % 5 == 1:
            print(i)
            td_elem = soup.select('td')[i].get_text
            print(td_elem)
    else:
            print(i)

for i in reversed(range(1,num_td_tags+1)):
    
    if i < num_td_tags:
        if i % 4 == 3:
            print(i)
            td_elem = soup.select('td')[i].get_text
            print(td_elem)
    else:
            print(i)
            
'''
for i in range(1,num_td_tags):
    td_elem = soup.select('td')[i]
    print(i)
    print(td_elem)
'''
#    print(td_tags)
#winsound.Beep(1000,1000)  #ビープ音（800Hzの音を1000msec流す）