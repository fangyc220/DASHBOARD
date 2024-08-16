import pandas as pd
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import logging

WEEKLY_WORK_LIST = '每週射出完工單.xlsx'
GET_ITEMS = ['製令單號', '機種代號', '完工日期', '完工數量', '批    號']
CHECK_LOGIN_URL = 'http://erpweb.pyramids.com.tw/index1.htm'

URL1 = 'http://erpweb1.pyramids.com.tw/Bart_prj/TMS/BadRecord/BadRecordRe.aspx?Com=C1&TypeNo='
URL2 = '&StartDate='

"""

製令單號            完工日期      品號              批號           完工數量     不良原因    不良數量
MD052408310030    20240810  20649-102     MDTMS240901001-00      100        氣泡        3
MD052408310030    20240810  20649-102     MDTMS240901005-00      120        氣泡        3


"""


def main():
    # working_data = input_data(WEEKLY_WORK_LIST)
    logging.basicConfig(level=logging.DEBUG)
    get_defect_num()


def get_defect_num(data=None):
    def check_login_required():
        response = requests.get(CHECK_LOGIN_URL)
        soup = BeautifulSoup(response.text, 'html.parser')
        # login_info = soup.find_all('div', {'id': 'login-in'})
        print(soup.text)
        return True

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
        response = session.post(CHECK_LOGIN_URL, data=login_data)
        soup = BeautifulSoup(response.text, 'html.parser')
        print(soup.text)
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


if __name__ == '__main__':
    main()