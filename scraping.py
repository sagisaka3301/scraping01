# ブラウザ操作等
from selenium import webdriver
import glob
import time
from bs4 import BeautifulSoup
import sys
from selenium.common.exceptions import NoSuchElementException
# Excel操作関連
import openpyxl
import os


# 現在のパス
now_path = "xxxx"
# 作成したいExcelのファイル名を指定
efile_name = "xxxx.xlsx"
efile_path = now_path + efile_name
# sheetの名前を指定
esheet_name1 = "xxxx"
# 対象サイトのリンク
web_link = "https://xxxx"
# 対象のhtmlタグ
tag = 'xxxx'
# 対象のクラス
class_name = "xxxx"
# 「もっとみる」等ボタンがある際、対象ボタンのクラス名
view_more = 'xxxx'

# Excelファイルの作成
def creExcel():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = esheet_name1
    wb.save(efile_path)


exist_conf = os.path.exists(efile_path)

if exist_conf == False:
    creExcel()
    print('Excelファイル「' + efile_name + '」を作成しました。')
    print(glob.glob("*.xlsx"))
else:
    print(efile_name + 'は既に存在しているファイル名です。')

chrome_d = 'chromedriver.exe'

if os.path.exists(now_path + chrome_d) == True:
    try:
        # ブラウザの起動処理
        browser = webdriver.Chrome(executable_path = now_path + chrome_d)
        browser.implicitly_wait(3)
        print('ブラウザを起動しました。')
    except:
        print('Chromedriverのバージョンが使用中のChromeバージョンと異なっている可能性があります。')
        sys.exit()
else:
    print('Chromedriverがインストールされていません。')
    pass   

browser.get(web_link)
time.sleep(3)

# 「もっと見る」等のボタンがあるときに、自動で押す処理
# end_numは、もっと見るボタンを押す回数。
# ボタンがある限り押すときは、while True:にする。
num = 0
end_num = 10
if view_more != '':
    while num < end_num:
        try:
            next_btn = browser.find_element_by_class_name(view_more)
            next_btn.click()
            num += 1
            time.sleep(1)
        except NoSuchElementException:
            print('終了しました。')
            break

html = browser.page_source
soup = BeautifulSoup(html, "html.parser")
tag_list = soup.find_all(tag, class_=class_name)
text_list = []
for i in range(0, len(tag_list)):
    text_list.append(tag_list[i].text)
    print(text_list[i])

# Excelにリストを書き込む
wb_load = openpyxl.load_workbook(efile_path)
wb_sheet = wb_load[esheet_name1]
for j in range (1, len(text_list) + 1):
    wb_sheet['A' + str(j)].value = text_list[j - 1]
wb_load.save(efile_path)
browser.quit()
