import dash
from dash import dcc, html, dash_table
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
import plotly.express as px
import get_defect_data
import pandas as pd

# 數據



# defect_data = {
#     'MD05240515001628384 REV.A 2024-08-13起蒼2': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13',
#                                                 '不良項目': '起蒼', '不良數': '2'},
#     'MD05240521002123297 REV.A 2024-08-01撞傷7': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-01',
#                                                 '不良項目': '撞傷', '不良數': '7'},
#     'MD05240521002123297 REV.A 2024-08-02撞傷8': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-02',
#                                                 '不良項目': '撞傷', '不良數': '8'},
#     'MD05240722002028478 REV.02 2024-08-01刮傷110': {'製令單號': 'MD052407220020', '品號': '28478 REV.02', '完工日期': '2024-08-01',
#                                                    '不良項目': '刮傷', '不良數': '110'},
#     'MX052407260005303-1716 REV.2 2024-08-15其他27': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2',
#                                                     '完工日期': '2024-08-15', '不良項目': '其他', '不良數': '27'},
#     'MX052407260005303-1716 REV.2 2024-08-15黑點11': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2',
#                                                     '完工日期': '2024-08-15', '不良項目': '黑點', '不良數': '11'},
#     'MX052407260005303-1716 REV.2 2024-08-15混色不均177': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2',
#                                                        '完工日期': '2024-08-15', '不良項目': '混色不均', '不良數': '177'},
#
# }
#


columns = ['機種代號', '完工日期', '不良項目', '不良數', '不良率 (%)']


app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    html.H1("TMS每日不良統計", className="mt-4 mb-4", id='WebTable'),
    dbc.Button('載入資料', id='button_for_loading_data', color='success', n_clicks=None),
    # dbc.DropdownMenu(id='Dropdown_for_display', label='顯示週期', children=[dbc.DropdownMenuItem('每日不良趨勢'), dbc.DropdownMenuItem('一週不良統計')]),
    dcc.Dropdown(id='Dropdown_for_display', options=[{'label': '每日不良趨勢', 'value': 'daily'}, {'label': '一週不良統計', 'value': 'weekly'}], value='daily'),
    dbc.Card([
        dbc.CardHeader("每日不良統計"),
        dbc.CardBody([
            dcc.Graph(
                id='不良趨勢圖',
                style={
                 'whiteSpace': 'normal',
                 'height': '500px',
                 'width': '1800px'
                        },
                )
        ]),
    ], style={'margin-bottom': '20px'}),

    dbc.Card([
        dbc.CardHeader('不良明細表'),
        dbc.CardBody([
            dash_table.DataTable(
                    page_size=10,
                    style_cell={
                         'textAlign': 'center',
                         'width': '10px',
                         'minWidth': '3px',
                         'maxWidth': '8px'
                                },
                    style_table={
                         'whiteSpace': 'normal',
                         'height': '350px',
                         'width': '1550px',
                         'overflow': 'auto',
                         'margin-left': '70px'
                                },
                    id='不良明細表',
                    style_data_conditional=[
                        {
                            'if': {
                                'filter_query': '{不良率 (%)} > 2',
                                'column_id': columns,
                            },
                            'color': 'red',
                            'fontWeight': 'bold'
                        }
                    ]
            )
        ])
    ])
], fluid=True)


@app.callback(
    Output(component_id='不良明細表', component_property='data'),
    Output(component_id='不良趨勢圖', component_property='figure'),
    Input(component_id='button_for_loading_data', component_property='n_clicks'),
    Input(component_id='Dropdown_for_display', component_property='value')
)
def loading_data(n_clicks, user_choose):
    print(n_clicks)
    print(user_choose)
    if n_clicks is not None:
        defect_data = {
        'MD05240515001628384 REV.A 2024-08-13起蒼2': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13', '不良項目': '起蒼', '不良數': '2'},
        'MD05240521002123297 REV.A 2024-08-01撞傷7': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-01', '不良項目': '撞傷', '不良數': '7'},
        'MD05240521002123297 REV.A 2024-08-02撞傷8': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-02', '不良項目': '撞傷', '不良數': '8'},
        'MD05240722002028478 REV.02 2024-08-01刮傷110': {'製令單號': 'MD052407220020', '品號': '28478 REV.02', '完工日期': '2024-08-01', '不良項目': '刮傷', '不良數': '110'},
        'MX052407260005303-1716 REV.2 2024-08-15其他27': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2', '完工日期': '2024-08-15', '不良項目': '其他', '不良數': '27'},
        'MX052407260005303-1716 REV.2 2024-08-15黑點11': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2', '完工日期': '2024-08-15', '不良項目': '黑點', '不良數': '11'},
        'MX052407260005303-1716 REV.2 2024-08-15混色不均177': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2', '完工日期': '2024-08-15', '不良項目': '混色不均', '不良數': '177'},
        'MD05240806000417878 REV.C-CLEAN 2024-08-06縮水22': {'製令單號': 'MD052408060004', '品號': '17878 REV.C-CLEAN','完工日期': '2024-08-06', '不良項目': '縮水', '不良數': '22'},
        'MD05240806000417878 REV.C-CLEAN 2024-08-06縮水50': {'製令單號': 'MD052408060004', '品號': '17878 REV.C-CLEAN','完工日期': '2024-08-06', '不良項目': '縮水', '不良數': '50'}
        }
        df = get_defect_data.weekly_work_excel_clean(defect_data)[0]
        fig = px.bar(df, x='完工日期', y='不良率 (%)', color='all_info', barmode='group',
                     labels={'不良率 (%)': '不良率 (%)', '完工日期': '完工日期', 'all_info': '品號 / 不良項目'},
                     title='每日不良趨勢')
        return df[columns].to_dict('records'), fig


if __name__ == '__main__':
    app.run_server(debug=True)
