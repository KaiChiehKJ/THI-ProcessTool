import pandas as pd
import random
from datetime import datetime 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def googlemap_crawler(placelist, searchcolumns = 'POIName', starturl = "https://www.google.com.tw/maps/@22.3912397,120.2980826,10z?hl=zh-TW&entry=ttu"):
    df = placelist.copy()

    # 打開瀏覽器
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get(starturl)
    time.sleep(random.uniform(5, 10))

    name_list = []
    googlename_list = []
    star_list = []
    comments_list = []
    googlemapurl  = []

    # 怕google擋爬蟲，所以每爬一次batch_size筆（50）資料就會等stop_bug秒（60)
    batch_size = random.randint(50, 100)
    stop_bug = random.uniform(5, 20)
    for i in range(0, len(df), batch_size):
        batch_df = df.iloc[i:i+batch_size]

        for name in list(batch_df[searchcolumns]):
            # 開始在輸入框輸入字樣
            search_box = driver.find_element(By.ID, "searchboxinput")
            search_box.clear()
            search_box.send_keys(name)  # 帶入查詢字樣
            search_box.send_keys(u'\ue007')  # 按下enter
            time.sleep(random.uniform(0, 3))

            # 如果有不只一個選項要執行，就點選第一個
            try:
                # 使用更通用的 XPath，只查找 class 為 'hfpxzc' 的第一個元素
                first_result = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@class='hfpxzc']"))
                )
                first_result.click()
                time.sleep(random.uniform(3, 5))
            except Exception as e:
                pass


            name_list.append(name)
            try:
                element = driver.find_element(By.XPATH, googlestarxpath)
                if element.text:
                    star_list.append(element.text)
                else:
                    star_list.append(None)
            except Exception as e:
                star_list.append(None)

            try:
                element = driver.find_element(By.XPATH, googlecommentxpath)
                if element.text:
                    comments_list.append(element.text)
                else:
                    comments_list.append(None)
            except Exception as e:
                comments_list.append(None)

            try:
                element = driver.find_element(By.XPATH, googlenamexpath)
                if element.text:
                    googlename_list.append(element.text)
                else:
                    googlename_list.append(None)
            except Exception as e:
                googlename_list.append(None)

            
            try:
                current_url = driver.current_url
                googlemapurl.append(current_url)
            except:
                googlemapurl.append(None)
            

            
            driver.get(starturl)
            time.sleep(0.5)
        time.sleep(stop_bug)
    # 關閉瀏覽器
    driver.quit()

    # ===== 把爬蟲資料轉成表格 =====
    results_df = pd.DataFrame({
        searchcolumns: name_list,
        'GoogleName': googlename_list,
        'POIStar': star_list,
        'POIComment': comments_list,
        'googlemapurl':googlemapurl
    })

    # 去掉 'POIComments' 中的括號、千分為逗號刪掉
    results_df['POIComment'] = results_df['POIComment'].str.replace(r'[\(\),]', '', regex=True)
    results_df['POIStar'] = pd.to_numeric(results_df['POIStar'], errors='coerce').fillna(0).astype(float)
    results_df['POIComment'] = pd.to_numeric(results_df['POIComment'], errors='coerce').fillna(0).astype(int)

    outputdf = pd.merge(df, results_df, on = searchcolumns)

    return outputdf
