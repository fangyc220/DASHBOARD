import pandas as pd
from datetime import datetime
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time

WEEKLY_WORK_LIST = '每週射出完工單.xlsx'
GET_ITEMS = ['製令單號', '機種代號', '完工日期', '完工數量', '批    號']
CHECK_LOGIN_URL = 'http://erpweb.pyramids.com.tw/index1.htm'
URL_TEST = 'http://erpweb1.pyramids.com.tw/Bart_prj/TMS/BadRecord/BadRecordRe.aspx?Com=C1&TypeNo=MX052407260005&StartDate=20240815'
URL1 = 'http://erpweb1.pyramids.com.tw/Bart_prj/TMS/BadRecord/BadRecordRe.aspx?Com=C1&TypeNo='
URL2 = '&StartDate='

"""

製令單號            完工日期      品號              批號           完工數量     不良原因    不良數量
MD052408310030    20240810  20649-102     MDTMS240901001-00      100        氣泡        3
MD052408310030    20240810  20649-102     MDTMS240901005-00      120        氣泡        3


"""


def main():
    # working_data = input_data(WEEKLY_WORK_LIST)
    # logging.basicConfig(level=logging.DEBUG)
    driver = webdriver.Chrome()
    driver.get(CHECK_LOGIN_URL)
    driver.implicitly_wait(3)

    username = 'PY310'
    password = '1'

    username_input = driver.find_element(By.NAME, 'ID')
    username_input.send_keys(username)
    password_input = driver.find_element(By.NAME, 'PWD')
    password_input.send_keys(password)
    login_button = driver.find_element(By.NAME, 'Button')
    login_button.click()
    driver.implicitly_wait(2)

    driver.get('http://erpweb1.pyramids.com.tw/Bart_prj/TMS/BadRecord/BadRecord.aspx?user_id=PY310&user_pw=1&Page1=1&Rnd=+')
    dropdown_ele = driver.find_element(By.NAME, 'DropDownList3')
    select = Select(dropdown_ele)
    select.select_by_value('-1')
    soup = BeautifulSoup(driver.page_source, features='html.parser')
    tags = soup.find_all('tr', style=lambda value: value and 'background' in value)
    for tag in tags:
        tokens = tag.text.strip().split()
        print(tokens)
        if len(tokens) > 3:
            print('製令單號 --> ', tokens[0], '完工日期 --> ', tokens[3][0:10])

    time.sleep(3)


def get_defect_num(data=None):
    def check_login_required():
        response = requests.get(CHECK_LOGIN_URL)
        soup = BeautifulSoup(response.text, 'html.parser')
        login_info = soup.find_all('input', {'name': 'ID'})
        if login_info:
            return True
        return False

    if check_login_required():
        print('==' * 30)
        print('=== 需要登入帳號密碼 ===')
        print('==' * 30)
        username = str(input('輸入工號: '))
        password = str(input('輸入密碼: '))
        session = requests.Session()
        login_data = {
            'ID': username,
            'PWD': password
        }
        # response = session.post(CHECK_LOGIN_URL, data=login_data)
        response = requests.get(CHECK_LOGIN_URL, params=login_data)
        response = requests.get(URL_TEST)
        soup = BeautifulSoup(response.text, 'html.parser')
        print(soup.text)
        # print(response.text)
        return session if response.ok else None
    # else:
    #     for i in range(len(data)):
    #         url = URL1 + data.iloc[i]['製令單號'] + URL2 + data.iloc[i]['完工日期']
    #         response = requests.get(url)
    #         soup = BeautifulSoup(response.text, 'html.parser')
    #         """
    #         TODO: 開始爬蟲抓取資料
    #         data = soup.find_all('div', class='content')
    #         """


def input_data(path):
    data = pd.read_excel(path)
    cols = data.iloc[2]
    data = data.iloc[3:len(data) - 1].reset_index(drop=True)
    cols = {data.columns[i]: cols[i] for i in range(len(cols))}
    data = data.rename(columns=cols)
    data = data[GET_ITEMS]

    # 轉變時間格式 2024/8/6 --> 20240806
    for i in range(len(data)):
        date_obj = datetime.strptime(str(data.iloc[i]['完工日期'])[0:10], "%Y-%m-%d")
        new_date = date_obj.strftime('%Y%m%d')
        data.iloc[i]['完工日期'] = new_date
    return data


"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import requests

# 1. 使用 Selenium 登錄網站
# 設置 WebDriver（這裡使用 ChromeDriver）
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# 打開登錄頁面
driver.get("http://example.com/login")

# 找到用戶名和密碼輸入框並填寫
username_input = driver.find_element(By.NAME, "username")
password_input = driver.find_element(By.NAME, "password")

username_input.send_keys("your_username")
password_input.send_keys("your_password")

# 找到登錄按鈕並點擊
login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
login_button.click()

# 等待頁面加載
driver.implicitly_wait(10)

# 2. 獲取登錄後的 cookies
cookies = driver.get_cookies()

# 關閉瀏覽器
driver.quit()

# 3. 使用 requests 和這些 cookies 來抓取數據
# 轉換 cookies 格式以便使用在 requests 中
session_cookies = {cookie['name']: cookie['value'] for cookie in cookies}

# 使用 session 來保持會話狀態
session = requests.Session()
session.cookies.update(session_cookies)

# 發送 GET 請求到目標頁面
target_url = "http://example.com/target-page"
response = session.get(target_url)

# 檢查響應狀態並處理頁面數據
if response.status_code == 200:
    print(response.text)  # 這裡可以使用 BeautifulSoup 進行進一步的數據處理
else:
    print(f"抓取失敗，狀態碼: {response.status_code}")



"""




if __name__ == '__main__':
    main()