
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
import glob

dir01 = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/"
dir02 = "C:/Users/touko/OneDrive/株価分析/excel/株式データ/株式"
#print(dir01)
file_list = glob.glob(dir02 + '/*.xlsx')
#print(file_list)
stock_code = 1301
ff = [s for s in file_list if str(stock_code) in s]
print(ff)
f = str(ff)
print(f)
#wb_company = openpyxl.load_workbook(str(ff))
#print(os.path.split(file_list[1]))
#stock_list=glob.glob(dir01 + '/*.xlsx')
name_list=[]

for i in file_list:
        file = os.path.basename(i)
#        print(file)
        stock_code = 1301
        
        name = os.path.split(file)
        #print(name)
        name_list.append(name)
#print(i)
#print(name_list)


