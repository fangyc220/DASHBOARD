import pandas as pd

DATA_PATH = 'data.xlsx'
ALL_DATA = 'all_data.xlsx'


def main():
    raw_data = pd.read_excel(DATA_PATH)
    p_data = process_data(raw_data)
    p_data.to_excel("output.xlsx")


def get_all_production_num():
    all_data = pd.read_excel(ALL_DATA)
    all_data.drop([0, 1], inplace=True)
    col = all_data.loc[2]
    all_data.rename(columns=col, inplace=True)
    all_data.drop(2, inplace=True)
    all_data = all_data[['製令單號', '機種代號', '完工數量']]
    all_data = all_data.groupby('製令單號').agg({
        '完工數量': 'sum',
        '機種代號': 'first'
    })
    return all_data


def process_data(raw_data=pd.read_excel(DATA_PATH)):
    data = {}
    all_production_num = get_all_production_num()
    for i in range(len(raw_data)):
        if raw_data.iloc[i]['製令單號'] not in data:
            try:
                data[raw_data.iloc[i]['製令單號']] = {}
                data[raw_data.iloc[i]['製令單號']]['總不良數'] = raw_data.iloc[i]['不良總數']
                data[raw_data.iloc[i]['製令單號']]['總生產數'] = all_production_num.loc[raw_data.iloc[i]['製令單號']]['完工數量']
                data[raw_data.iloc[i]['製令單號']]['品號'] = all_production_num.loc[raw_data.iloc[i]['製令單號']]['機種代號']
            except:
                print(f'Not find {raw_data.iloc[i]["製令單號"]} data.')
        else:
            data[raw_data.iloc[i]['製令單號']]['總不良數'] += raw_data.iloc[i]['不良總數']
    data = pd.DataFrame(data).T
    data = data.reset_index()
    data.rename(columns={'index': '製令單號'}, inplace=True)
    data.dropna(inplace=True)
    data['總不良率'] = pd.to_numeric((data['總不良數'] / (data['總生產數'] + data['總不良數'])) * 100, errors='coerce').round(2)
    data = data[['品號', '總生產數', '總不良數', '總不良率', '製令單號']]
    return data


if __name__ == '__main__':
    main()
