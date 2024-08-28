import dash
from dash import dcc, html, dash_table
import dash_bootstrap_components as dbc
import plotly.express as px
import get_defect_data

# 數據
defect_data = {
    'MD05240515001628384 REV.A 2024-08-13起蒼2': {'製令單號': 'MD052405150016', '品號': '28384 REV.A', '完工日期': '2024-08-13',
                                                '不良項目': '起蒼', '不良數': '2'},
    'MD05240521002123297 REV.A 2024-08-01撞傷7': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-01',
                                                '不良項目': '撞傷', '不良數': '7'},
    'MD05240521002123297 REV.A 2024-08-02撞傷8': {'製令單號': 'MD052405210021', '品號': '23297 REV.A', '完工日期': '2024-08-02',
                                                '不良項目': '撞傷', '不良數': '8'},
    'MD05240722002028478 REV.02 2024-08-01刮傷110': {'製令單號': 'MD052407220020', '品號': '28478 REV.02', '完工日期': '2024-08-01',
                                                   '不良項目': '刮傷', '不良數': '110'},
    'MX052407260005303-1716 REV.2 2024-08-15其他27': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2',
                                                    '完工日期': '2024-08-15', '不良項目': '其他', '不良數': '27'},
    'MX052407260005303-1716 REV.2 2024-08-15黑點11': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2',
                                                    '完工日期': '2024-08-15', '不良項目': '黑點', '不良數': '11'},
    'MX052407260005303-1716 REV.2 2024-08-15混色不均177': {'製令單號': 'MX052407260005', '品號': '303-1716 REV.2',
                                                       '完工日期': '2024-08-15', '不良項目': '混色不均', '不良數': '177'},

}
df = get_defect_data.weekly_work_excel_clean(defect_data)
columns = ['機種代號', '完工日期', '不良項目', '不良數', '不良率 (%)']

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    html.H1("TMS每日不良統計", className="mt-4 mb-4"),
    dbc.Card([
        dbc.CardHeader("每日不良統計"),
        dbc.CardBody([
            dcc.Graph(
                id='不良趨勢圖',
                figure=px.bar(df, x='完工日期', y='不良率 (%)', color='all_info', barmode='group',
                              labels={'不良率 (%)': '不良率 (%)', '完工日期': '完工日期', 'all_info': '品號 / 不良項目'},
                              title='每日不良趨勢'
                              )
                )
        ]),
    ], style={'margin-bottom': '20px'}),

    dbc.Card([
        dbc.CardHeader('不良明細表'),
        dbc.CardBody([
            dash_table.DataTable(
                    data=df[columns].to_dict('records'),
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


if __name__ == '__main__':
    app.run_server(debug=True)
