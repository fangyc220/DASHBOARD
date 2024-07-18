from dash import Dash, html, dash_table, dcc
import plotly.express as px
import pandas as pd
import data_process

df = data_process.process_data()
app = Dash()

app.layout = html.Div([
    html.Div(children='My First App with Data.'),
    dash_table.DataTable(data=df.to_dict('records'), page_size=10),
    dcc.Graph(figure=px.histogram(df, x='品號', y='總不良率'))
])

if __name__ == '__main__':
    app.run(debug=True)
