import sys
import sys
from selenium import webdriver
from selenium.webdriver.common.keys import Keys as keys
import os

# 自動ログイン関数を宣言
#
#
def AutoLogin():
  # 起動するブラウザを宣言します 
  browser = webdriver.Chrome('C:/Users/touko/program/chromedriver.exe') 
  # ログイン対象のWebページURLを宣言します 
  url = "https://account.kabutan.jp/login" 
  # 対象URLをブラウザで表示します。 
  browser.get(url)
  # ログインIdとパスワードの入力領域を取得します。 
  login_id = browser.find_element_by_xpath("//input[@id='session_email']") 
  login_pw = browser.find_element_by_xpath("//input[@id='session_password']")
  # ログインIDとパスワードを入力します。
  userid = "toukouikitai@hotmail.com" 
  userpw = "s4b4egqekabutan"
  login_id.send_keys(userid) 
  login_pw.send_keys(userpw)
  # ログインボタンをクリックします。 

  login_btn = browser.find_element_by_xpath(".//input[@type='submit']")
  login_btn.click()

# AutoLogin関数を実行します。
#
ret = AutoLogin()