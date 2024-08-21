import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

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
    defect_dict = {}
    for tag in tags:
        tokens = tag.text.strip().split()
        if len(tokens) > 3:
            date = trans_date_foam(tokens[3][0:10])
            defect_url = URL1 + tokens[0] + URL2 + date
            driver.get(defect_url)
            soup = BeautifulSoup(driver.page_source, features='html.parser')
            defect_tags = soup.find_all('tr', style=lambda value: value and 'background' in value and len(value) < 30)
            [get_defect_dict(defect_dict, defect_tag) if defect_tag.text.strip() not in defect_dict else None for defect_tag in defect_tags]

    print(defect_dict)

    """
    {
    'MD05240515001628384 REV.A 2024-08-13起蒼2': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13', '不良項目': '起蒼', '不良數': '2'}, 
    'MD05240521002123297 REV.A 2024-08-01撞傷7': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-01', '不良項目': '撞傷', '不良數': '7'}, 
    'MD05240521002123297 REV.A 2024-08-02撞傷8': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-02', '不良項目': '撞傷', '不良數': '8'}, 
    'MD05240722002028478 REV.02 2024-08-01刮傷110': {'製令單號': 'MD052407220020', '品號': '28478 REV.02', '完工日期': '2024-08-01', '不良項目': '刮傷', '不良數': '110'}, 
    'MX052407260005303-1716 REV.2 2024-08-15其他27': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2', '完工日期': '2024-08-15', '不良項目': '其他', '不良數': '27'}
    }
    
    """


    weekly_work_excel_clean()
    insert_defect_num_to_excel(defect_dict)


def weekly_work_excel_clean():
    weekly_data = pd.read_excel(WEEKLY_WORK_LIST)
    columns = weekly_data.iloc[2]
    weekly_data = weekly_data.iloc[3:]
    weekly_data = weekly_data.rename(columns=columns)
    weekly_data.reset_index(inplace=True, drop=True)
    weekly_data['INDEX_KEY'] = [str(weekly_data['製令單號'].iloc[i]) + str(weekly_data['完工日期'].iloc[i]) for i in range(len(weekly_data))]
    # print(weekly_data)

def insert_defect_num_to_excel(dic):
    pass


def get_defect_dict(defect_dict, defect_tag):
    """
    defect_dice[MD05240515001628384 REV.A 2024-08-13起蒼2] = {
        '製令單號': 'MD052405150016',
        '品號': '28384 REV.A',
        '完工日期': '2024-08-13',
        '不良項目': '起蒼',
        '不良數': '2'
    }
    """
    defect_dict[defect_tag.text.strip()] = {}
    defect_lst = defect_tag.text.strip().split(' ')
    defect_dict[defect_tag.text.strip()]['製令單號'] = defect_lst[0][:14]
    defect_dict[defect_tag.text.strip()]['品號'] = defect_lst[0][14:] + ' ' + defect_lst[1]
    defect_dict[defect_tag.text.strip()]['完工日期'] = defect_lst[2][:10]
    defect_dict[defect_tag.text.strip()]['不良項目'] = get_defect_item_and_num(defect_lst[2][10:])[0]
    defect_dict[defect_tag.text.strip()]['不良數'] = get_defect_item_and_num(defect_lst[2][10:])[1]


def get_defect_item_and_num(defect_str):
    defect_item, defect_num = '', ''
    for ch in defect_str:
        if ch.isdigit():
            defect_num += ch
        else:
            defect_item += ch
    return defect_item, defect_num


def trans_date_foam(date):
    date_obj = datetime.strptime(date, '%Y-%m-%d')
    new_date = date_obj.strftime('%Y%m%d')
    return new_date


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
