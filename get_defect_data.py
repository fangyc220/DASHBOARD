import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

WEEKLY_WORK_LIST = '每週射出完工單.xlsx'
CHECK_LOGIN_URL = 'http://erpweb.pyramids.com.tw/index1.htm'
URL1 = 'http://erpweb1.pyramids.com.tw/Bart_prj/TMS/BadRecord/BadRecordRe.aspx?Com=C1&TypeNo='
URL2 = '&StartDate='


def get_defect_data_from_web():
    """
    {
    'MD05240515001628384 REV.A 2024-08-13起蒼2': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13', '不良項目': '起蒼', '不良數': '2'},
    'MD05240521002123297 REV.A 2024-08-01撞傷7': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-01', '不良項目': '撞傷', '不良數': '7'},
    'MD05240521002123297 REV.A 2024-08-02撞傷8': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-02', '不良項目': '撞傷', '不良數': '8'},
    'MD05240722002028478 REV.02 2024-08-01刮傷110': {'製令單號': 'MD052407220020', '品號': '28478 REV.02', '完工日期': '2024-08-01', '不良項目': '刮傷', '不良數': '110'},
    'MX052407260005303-1716 REV.2 2024-08-15其他27': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2', '完工日期': '2024-08-15', '不良項目': '其他', '不良數': '27'}
    }

    """
    # username = str(input('輸入工號 : '))
    # password = str(input('輸入密碼 : '))

    username = 'PY310'
    password = '2'

    data = loading_data()
    search_month = get_month_from_excel(data)

    driver = webdriver.Chrome()
    driver.get(CHECK_LOGIN_URL)
    driver.implicitly_wait(3)

    username_input = driver.find_element(By.NAME, 'ID')
    username_input.send_keys(username)
    password_input = driver.find_element(By.NAME, 'PWD')
    password_input.send_keys(password)
    login_button = driver.find_element(By.NAME, 'Button')
    login_button.click()
    driver.implicitly_wait(2)

    defect_dict = {}
    for month in search_month:
        print(f'開始抓取{month}月不良資料')
        driver.get('http://erpweb1.pyramids.com.tw/Bart_prj/TMS/BadRecord/BadRecord.aspx?user_id=' + username + '&user_pw=' + password + '&Page1=1&Rnd=+')
        dropdown_month = driver.find_element(By.NAME, 'DropDownList2')
        select_month = Select(dropdown_month)
        select_month.select_by_value(month)

        dropdown_ele = driver.find_element(By.NAME, 'DropDownList3')
        select = Select(dropdown_ele)
        select.select_by_value('-1')

        soup = BeautifulSoup(driver.page_source, features='html.parser')
        tags = soup.find_all('tr', style=lambda value: value and 'background' in value)
        for tag in tags:
            tokens = tag.text.strip().split()
            if len(tokens) > 3:
                date = trans_date_foam(tokens[3][0:10])
                defect_url = URL1 + tokens[0] + URL2 + date
                driver.get(defect_url)
                soup = BeautifulSoup(driver.page_source, features='html.parser')
                defect_tags = soup.find_all('tr', style=lambda value: value and 'background' in value and len(value) < 30)
                [get_defect_dict(defect_dict, defect_tag) if defect_tag.text.strip() not in defect_dict else None for defect_tag in defect_tags]
        print(f'抓取完畢{month}月不良資料')
    return defect_dict, data


def weekly_work_excel_clean(defect_dict, weekly_data):
    """
    :param defect_dict: dict
    :param weekly_data: Dataframe
    :return: working_num, total_working_num

    working_num = (細分每日的不良狀況, 比較細項一點)
                   製令單號        完工日期    完工數量  ...  不良數 不良率 (%)               all_info
    134  MX052407260005  2024-08-15  3603.0  ...  177    4.91   303-1716 REV.2, 混色不均
    135  MD052408060004  2024-08-06  5770.0  ...   50    0.87  17878 REV.C-CLEAN, 縮水
    126  MX052407260005  2024-08-15  3603.0  ...   27    0.75     303-1716 REV.2, 其他
    124  MD052408060004  2024-08-06  5770.0  ...   22    0.38  17878 REV.C-CLEAN, 縮水
    133  MX052407260005  2024-08-15  3603.0  ...   11    0.31     303-1716 REV.2, 黑點
    74   MD052405150016  2024-08-13  5568.0  ...    2    0.04        28384 REV.A, 起蒼

    total_working_num = (把相同的製令、相同的不良項目相加)
             製令單號               機種代號  不良項目  不良數    完工數量   不良率 (%)
    0  MD052405150016        28384 REV.A    起蒼    2  5566.0   0.03592
    1  MD052408060004  17878 REV.C-CLEAN    縮水   72  5798.0  1.226576
    2  MX052407260005     303-1716 REV.2    其他   27  3388.0   0.79063
    3  MX052407260005     303-1716 REV.2  混色不均  177  3388.0  4.964937
    4  MX052407260005     303-1716 REV.2    黑點   11  3388.0  0.323625

    """
    # need_columns = ['製令單號', '機種代號', '完工日期', '完工數量', '批    號']
    # weekly_data = pd.read_excel(WEEKLY_WORK_LIST)
    # columns = weekly_data.iloc[2]
    # weekly_data = weekly_data.iloc[3:]
    # weekly_data = weekly_data.rename(columns=columns)
    # weekly_data = weekly_data[need_columns]
    working_num = weekly_data.groupby(['製令單號', '完工日期'])['完工數量'].sum().reset_index()  # 依據製令單號、完工日期分類，並且只有sum完工數量的內容
    working_num = working_num.merge(weekly_data[['製令單號', '機種代號']], on='製令單號', how='left')  # 依照製令單號，把品號補回來
    working_num.drop_duplicates(subset=['製令單號', '完工日期'], inplace=True)  # 在merge時會有重覆的狀況，因此移除相同的製令單號、完工日期的rows
    working_num.reset_index(drop=True, inplace=True)

    working_num['不良項目'] = pd.NA
    working_num['不良數'] = 0

    insert_defect_num_to_excel(defect_dict, working_num)
    working_num.sort_values(by='不良率 (%)', ascending=False, inplace=True)
    working_num['all_info'] = working_num['機種代號'] + ', ' + working_num['不良項目']

    total_columns = ['製令單號', '機種代號', '不良項目', '不良數']
    total_working_num = working_num[total_columns]
    production_num = weekly_data.groupby(['製令單號'])['完工數量'].sum().reset_index()

    total_working_num['不良數'] = pd.to_numeric(total_working_num['不良數'], errors='coerce').astype(int)
    total_working_num = total_working_num.groupby(['製令單號', '機種代號', '不良項目'])['不良數'].sum().reset_index()
    total_working_num = total_working_num.merge(production_num[['製令單號', '完工數量']], on='製令單號', how='left')
    total_working_num['不良率 (%)'] = (total_working_num['不良數'] / (total_working_num['不良數'] + total_working_num['完工數量']) * 100)
    total_working_num.sort_values(by='不良率 (%)', ascending=False, inplace=True)
    total_working_num['不良率 (%)'] = total_working_num['不良率 (%)'].apply(lambda x: f"{x:.2f}")
    total_working_num['繪圖x軸使用'] = total_working_num['機種代號'] + ' / ' + total_working_num['製令單號']
    return working_num, total_working_num


def insert_defect_num_to_excel(defect_dict, working_num_date):
    for key in defect_dict:
        number = defect_dict[key]['製令單號']
        date = pd.to_datetime(defect_dict[key]['完工日期'])
        defect_item = defect_dict[key]['不良項目']
        get_index = working_num_date.loc[(working_num_date['製令單號'] == number) & (working_num_date['完工日期'] == date)].index
        if len(get_index) > 0 and working_num_date.loc[get_index.tolist()[0], '不良數'] == 0:
            working_num_date.loc[get_index.tolist()[0], '不良項目'] = defect_dict[key]['不良項目']
            working_num_date.loc[get_index.tolist()[0], '不良數'] = int(defect_dict[key]['不良數'])
        elif len(get_index) > 0 and working_num_date.loc[get_index.tolist()[0], '不良數'] != 0:
            get_index_second_check = working_num_date.loc[(working_num_date['製令單號'] == number) & (working_num_date['完工日期'] == date) & (working_num_date['不良項目'] == defect_item)].index
            if len(get_index_second_check) > 0:
                working_num_date.loc[get_index_second_check.tolist()[0], '不良數'] += int(defect_dict[key]['不良數'])
            elif len(get_index_second_check) == 0:
                new_index = len(working_num_date)
                working_num_date.loc[new_index] = working_num_date.loc[get_index.tolist()[0]]
                working_num_date.loc[new_index, '不良項目'] = defect_dict[key]['不良項目']
                working_num_date.loc[new_index, '不良數'] = int(defect_dict[key]['不良數'])

        all_index_for_add_num = working_num_date.loc[(working_num_date['製令單號'] == number) & (working_num_date['完工日期'] == date)].index
        for index in all_index_for_add_num.tolist():
            working_num_date.loc[index, '完工數量'] += int(defect_dict[key]['不良數'])

    working_num_date['不良率 (%)'] = ((working_num_date['不良數'].astype('int') / working_num_date['完工數量']) * 100).astype(float)
    working_num_date['不良率 (%)'] = working_num_date['不良率 (%)'].apply(lambda x: f"{x:.2f}")
    for i in range(len(working_num_date)):
        if isinstance(working_num_date.loc[i, '完工日期'], datetime):
            working_num_date.loc[i, '完工日期'] = working_num_date.loc[i, '完工日期'].strftime('%Y-%m-%d')


def loading_data():
    need_columns = ['製令單號', '機種代號', '完工日期', '完工數量', '批    號']
    data = pd.read_excel(WEEKLY_WORK_LIST)
    columns = data.iloc[2]
    data = data.iloc[3:]
    data = data.rename(columns=columns)
    data = data[need_columns]
    return data


def get_month_from_excel(data):
    all_month = []
    for date in data['完工日期']:
        if type(date) is datetime and '0' in str(date)[5:7]:
            all_month.append(str(date)[6]) if str(date)[6] not in all_month else None
        elif type(date) is datetime:
            all_month.append(str(date)[5:7])
    return all_month


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