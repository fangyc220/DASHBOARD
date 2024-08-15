import pandas as pd
import datetime import datetime
WEEKLY_WORK_LIST = '每週射出完工單.xlsx'
GET_ITEMS = ['製令單號', '機種代號', '完工日期', '完工數量', '批    號']


def main():
    working_data = input_data(WEEKLY_WORK_LIST)


def input_data(path):
    data = pd.read_excel(path)
    cols = data.iloc[2]
    data = data.iloc[3:len(data) - 1].reset_index(drop=True)
    cols = {data.columns[i]: cols[i] for i in range(len(cols))}
    data = data.rename(columns=cols)
    data = data[GET_ITEMS]

    # 轉變時間格式
    data_obj = datetime.strptime("2024/8/5", "")
    return data


if __name__ == '__main__':
    main()
