#a   -*- coding: utf-8 -*-
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import openpyxl

# 出品情報を格納するリスト
exhibition_data = []
options = webdriver.ChromeOptions()
# ヘッドレスモード
# options.add_argument('--headless')
# webdriverのパスを指定
service = Service(executable_path="chromedriver-mac-arm64/chromedriver")
driver = webdriver.Chrome(service=service)
# 検索先のURL
driver.get('https://hiroba.dqx.jp/sc/search/')
time.sleep(1)
# 検索窓入力
s = driver.find_element(By.XPATH,'//*[@id="sqexid"]')
s.send_keys('user')
s = driver.find_element(By.XPATH,'//*[@id="password"]')
s.send_keys('pass')
# ログインボタンクリック
driver.find_element(By.XPATH,'//*[@id="login-button"]').click()
driver.find_element(By.XPATH,'//*[@id="welcome_box"]/div[2]/a').click()
driver.find_element(By.XPATH,'//*[@id="contentArea"]/div/div[2]/form/table/tbody/tr[2]/td[3]/a').click()
# 検索
time.sleep(2)

#すべての文字で検索
hiragana = [chr(i) for i in range(12353, 12436)]
for h in hiragana:
    # 検索するアイテム名を格納
    search_word = h
    # 検索フォームに入力
    s = driver.find_element(By.XPATH,'//*[@id="searchword"]').clear()
    s = driver.find_element(By.XPATH,'//*[@id="searchword"]')
    s.send_keys(search_word)
    driver.find_element(By.XPATH,'//*[@id="searchBoxArea"]/form/p[2]/input').click()
    time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="searchTabItem"]').click()
    time.sleep(1)
    while(True):
        # 出品データの取得
        elements = driver.find_elements(By.TAG_NAME,'tr')
        # リストに格納
        for elem in elements:
            exhibition_data.append(elem.text.split())
        if len(driver.find_elements(By.XPATH,'//*[@class="next"]')) > 0 :
            driver.find_element(By.XPATH,'//*[@class="next"]').click()
        else:
            break
    sleep(5)
# 無駄な要素を削除
exhibition_data = [i for i in exhibition_data if len(i) > 1 and i[0] != "アイテム名"]

df_item = pd.DataFrame(exhibition_data)
df_item = df_item.drop_duplicates().reset_index(drop=True)

is_dougu = df_item.iloc[:, 4] == "装備可能な職業はありません"
df_dougu = df_item[is_dougu].reset_index(drop=True)
df_soubi = df_item[~ is_dougu].reset_index(drop=True)


workbook = openpyxl.Workbook()
sheet = workbook.active
# Excelのシートへ出力
for i in range(len(exhibition_data)):
    for j in range(len(exhibition_data[i])):
        sheet.cell(row=i + 1, column=j + 1).value = exhibition_data[i][j]
# リストの中身を一旦リセット
del exhibition_data[:]
# Excelファイルを保存
workbook.save(search_word + '.xlsx')
# ワークブックを閉じる
workbook.close()
# ドライバーを閉じる
driver.quit()
