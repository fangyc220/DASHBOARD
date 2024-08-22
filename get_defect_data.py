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
    # driver = webdriver.Chrome()
    # driver.get(CHECK_LOGIN_URL)
    # driver.implicitly_wait(3)
    #
    # username = 'PY310'
    # password = '1'
    #
    # username_input = driver.find_element(By.NAME, 'ID')
    # username_input.send_keys(username)
    # password_input = driver.find_element(By.NAME, 'PWD')
    # password_input.send_keys(password)
    # login_button = driver.find_element(By.NAME, 'Button')
    # login_button.click()
    # driver.implicitly_wait(2)
    #
    # driver.get('http://erpweb1.pyramids.com.tw/Bart_prj/TMS/BadRecord/BadRecord.aspx?user_id=PY310&user_pw=1&Page1=1&Rnd=+')
    # dropdown_ele = driver.find_element(By.NAME, 'DropDownList3')
    # select = Select(dropdown_ele)
    # select.select_by_value('-1')
    # soup = BeautifulSoup(driver.page_source, features='html.parser')
    # tags = soup.find_all('tr', style=lambda value: value and 'background' in value)
    # defect_dict = {}
    # for tag in tags:
    #     tokens = tag.text.strip().split()
    #     if len(tokens) > 3:
    #         date = trans_date_foam(tokens[3][0:10])
    #         defect_url = URL1 + tokens[0] + URL2 + date
    #         driver.get(defect_url)
    #         soup = BeautifulSoup(driver.page_source, features='html.parser')
    #         defect_tags = soup.find_all('tr', style=lambda value: value and 'background' in value and len(value) < 30)
    #         [get_defect_dict(defect_dict, defect_tag) if defect_tag.text.strip() not in defect_dict else None for defect_tag in defect_tags]
    #
    # print(defect_dict)

    """
    {
    'MD05240515001628384 REV.A 2024-08-13起蒼2': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13', '不良項目': '起蒼', '不良數': '2'},
    'MD05240521002123297 REV.A 2024-08-01撞傷7': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-01', '不良項目': '撞傷', '不良數': '7'},
    'MD05240521002123297 REV.A 2024-08-02撞傷8': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-02', '不良項目': '撞傷', '不良數': '8'},
    'MD05240722002028478 REV.02 2024-08-01刮傷110': {'製令單號': 'MD052407220020', '品號': '28478 REV.02', '完工日期': '2024-08-01', '不良項目': '刮傷', '不良數': '110'},
    'MX052407260005303-1716 REV.2 2024-08-15其他27': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2', '完工日期': '2024-08-15', '不良項目': '其他', '不良數': '27'}
    }

    """

    defect_dict = {
    'MD05240515001628384 REV.A 2024-08-13起蒼2': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13', '不良項目': '起蒼', '不良數': '2'},
    'MD05240521002123297 REV.A 2024-08-01撞傷7': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-01', '不良項目': '撞傷', '不良數': '7'},
    'MD05240521002123297 REV.A 2024-08-02撞傷8': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-02', '不良項目': '撞傷', '不良數': '8'},
    'MD05240722002028478 REV.02 2024-08-01刮傷110': {'製令單號': 'MD052407220020', '品號': '28478 REV.02', '完工日期': '2024-08-01', '不良項目': '刮傷', '不良數': '110'},
    'MX052407260005303-1716 REV.2 2024-08-15其他27': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2', '完工日期': '2024-08-15', '不良項目': '其他', '不良數': '27'},
    'MD05240515001628384 REV.A 2024-08-13黑點7': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13', '不良項目': '黑點', '不良數': '7'},
    'MD05240515001628384 REV.A 2024-08-13混色不均19': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13', '不良項目': '混色不均', '不良數': '19'},

    }

    weekly_data = weekly_work_excel_clean(defect_dict)
    print(weekly_data.loc[126])
    print(weekly_data.loc[134])
    print(weekly_data.loc[113])



def weekly_work_excel_clean(defect_dict):
    need_columns = ['製令單號', '機種代號', '完工日期', '完工數量', '批    號']
    weekly_data = pd.read_excel(WEEKLY_WORK_LIST)
    columns = weekly_data.iloc[2]
    weekly_data = weekly_data.iloc[3:]
    weekly_data = weekly_data.rename(columns=columns)
    weekly_data = weekly_data[need_columns]
    working_num = weekly_data.groupby(['製令單號', '完工日期'])['完工數量'].sum().reset_index()  # 依據製令單號、完工日期分類，並且只有sum完工數量的內容
    working_num = working_num.merge(weekly_data[['製令單號', '機種代號']], on='製令單號', how='left')  # 依照製令單號，把品號補回來
    working_num.drop_duplicates(subset=['製令單號', '完工日期'], inplace=True)  # 在merge時會有重覆的狀況，因此移除相同的製令單號、完工日期的rows
    working_num.reset_index(drop=True, inplace=True)
    """
                       製令單號                 完工日期      完工數量                     機種代號
    0    MD052310310276  2024-08-06 00:00:00      1100        15171-103 REV.D_T
    1    MD052310310276  2024-08-07 00:00:00      1050        15171-103 REV.D_T
    2    MD052310310276  2024-08-08 00:00:00  77883179        15171-103 REV.D_T
    3    MD052310310276  2024-08-09 00:00:00      1800        15171-103 REV.D_T
    4    MD052310310277  2024-08-05 00:00:00      1700  15171-103 REV.D_T-PRINT
    """

    working_num['不良項目'] = pd.NA
    working_num['不良數'] = 0

    insert_defect_num_to_excel(defect_dict, working_num)
    return working_num


def insert_defect_num_to_excel(defect_dict, working_num_date):
    for key in defect_dict:
        number = defect_dict[key]['製令單號']
        date = pd.to_datetime(defect_dict[key]['完工日期'])
        get_index = working_num_date.loc[(working_num_date['製令單號'] == number) & (working_num_date['完工日期'] == date)].index
        if len(get_index) > 0 and working_num_date.loc[get_index.tolist()[0], '不良數'] == 0:
            working_num_date.loc[get_index.tolist()[0], '不良項目'] = defect_dict[key]['不良項目']
            working_num_date.loc[get_index.tolist()[0], '不良數'] = defect_dict[key]['不良數']
        elif len(get_index) > 0 and working_num_date.loc[get_index.tolist()[0], '不良數'] != 0:
            new_index = len(working_num_date)
            working_num_date.loc[new_index] = working_num_date.loc[get_index.tolist()[0]]
            working_num_date.loc[new_index, '不良項目'] = defect_dict[key]['不良項目']
            working_num_date.loc[new_index, '不良數'] = defect_dict[key]['不良數']

        all_index_for_add_num = working_num_date.loc[(working_num_date['製令單號'] == number) & (working_num_date['完工日期'] == date)].index
        for index in all_index_for_add_num.tolist():
            working_num_date.loc[index, '完工數量'] += int(defect_dict[key]['不良數'])

    working_num_date['不良率 (%)'] = (working_num_date['不良數'].astype('int') / working_num_date['完工數量']) * 100


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
