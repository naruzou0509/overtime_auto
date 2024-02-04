# module
import openpyxl as op
from selenium import webdriver
from selenium.webdriver.common.by import By

# cellの値取得関数
def wb_value(user,cell_place):
    wb = op.load_workbook(r"C:\Users\{}\Documents\overtime_list_filter.xlsm".format(user), data_only=True) # book読み込み
    ws = wb["List"] # シートの「List」を指定
    cell = ws[cell_place].value # 対象月の取得
    wb.close() # bookを閉じる
    return cell

# 転記関数
def input_post(user,cell_place,driver,i):
    cell_start = wb_value(user,cell_place[0]) # 始業の値
    cell_end = wb_value(user,cell_place[1]) # 終業の値
    cell_cause = wb_value(user,cell_place[2]) # 理由の値

    if not cell_cause : # 変数が空のとき
        return # 処理終了

    ele_start = driver.find_elements(By.NAME, "PeriodStart") # 始業の場所指定
    ele_start[i].send_keys(cell_start) # 始業の場所に値入力
    ele_end = driver.find_elements(By.NAME, "PeriodEnd") # 終業の場所指定
    ele_end[i].send_keys(cell_end) # 終業の場所に値入力
    ele_cause = driver.find_elements(By.NAME, "Cause") # 理由の場所指定
    ele_cause[i].send_keys(cell_cause) # 理由の場所に値入力


