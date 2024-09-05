import dash
from dash import dcc, html, dash_table
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
import plotly.express as px
import get_defect_data
import pandas as pd

columns = ['機種代號', '完工日期', '不良項目', '不良數', '不良率 (%)']
weekly_columns = ['製令單號', '機種代號', '不良項目', '不良數', '完工數量', '不良率 (%)']
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    html.H1("TMS不良統計表", className="mt-4 mb-4", id='WebTable'),
    dbc.Button('載入資料', id='button_for_loading_data', color='success', n_clicks=None),
    dcc.Store(id='Daily_data'),
    dcc.Store(id='Weekly_data'),
    dcc.Dropdown(id='Dropdown_for_display', options=[{'label': '每日不良趨勢', 'value': 'daily'}, {'label': '一週不良統計', 'value': 'weekly'}], value='daily'),
    dbc.Card([
        dbc.CardHeader("TMS不良率統計"),
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
                    style_data_conditional=[]
            )
        ])
    ])
], fluid=True)


@app.callback(
    Output(component_id='Daily_data', component_property='data'),
    Output(component_id='Weekly_data', component_property='data'),
    Input(component_id='button_for_loading_data', component_property='n_clicks'),
    prevent_initial_call=True
)
def loading_data(n_clicks):
    defect_data, data = get_defect_data.get_defect_data_from_web()
    daily, weekly = get_defect_data.weekly_work_excel_clean(defect_data, data)
    return daily.to_json(orient='split'), weekly.to_json(orient='split')


@app.callback(
    Output(component_id='不良明細表', component_property='data'),
    Output(component_id='不良明細表', component_property='style_data_conditional'),
    Output(component_id='不良趨勢圖', component_property='figure'),
    Input(component_id='Dropdown_for_display', component_property='value'),
    Input(component_id='Daily_data', component_property='data'),
    Input(component_id='Weekly_data', component_property='data')
)
def display_fig(user_choose, daily_data, weekly_data):
    if user_choose == 'daily':
        df = pd.read_json(daily_data, orient='split')
        fig = px.bar(df, x='完工日期', y='不良率 (%)', color='all_info', barmode='group',
                     labels={'不良率 (%)': '不良率 (%)', '完工日期': '完工日期', 'all_info': '品號 / 不良項目'})
        fig.update_layout(legend=dict(font=dict(size=18)))  # 你可以調整這個數值來改變字體大小
        style = [{'if': {'filter_query': '{不良率 (%)} > 2', 'column_id': columns}, 'color': 'red', 'fontWeight': 'bold'}]
        return df[columns].to_dict('records'), style, fig
    else:
        df = pd.read_json(weekly_data, orient='split')
        fig = px.bar(df, x='機種代號', y='不良率 (%)', color='不良項目', barmode='group',
                     labels={'不良率 (%)': '不良率 (%)', '機種代號': '品號'})
        fig.update_layout(legend=dict(font=dict(size=18)))  # 你可以調整這個數值來改變字體大小
        style = [{'if': {'filter_query': '{不良率 (%)} > 2', 'column_id': weekly_columns}, 'color': 'red', 'fontWeight': 'bold'}]
        return df[weekly_columns].to_dict('records'), style, fig


if __name__ == '__main__':
    app.run_server(debug=True)
