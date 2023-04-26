import glob
import os
import openpyxl
import requests
import time
import datetime
import sys
import winsound

#フォルダ内画像をexcelリストにする
#フォルダを指定
#excelを開く
#フォルダ内を走査
#通し番号、ファイル名、フルパスをexcelに転記する

dir_list = "C:/Users/touko/OneDrive/自動化用/02.ラーメン画像/"
dir_pic = dir_list + "ラーメン画像/"
wb =openpyxl.load_workbook(dir_list + "ラーメン画像テーブル.xlsx")
files = glob.glob(dir_pic + "*.JPG")
for file in files:
#    print(file)
    
