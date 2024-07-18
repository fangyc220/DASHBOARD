from dash import Dash, html, dash_table
import pandas as pd

df = pd.read_excel('data.xlsx')
app = Dash()

app.layout = html.Div([
    html.Div(children='My First App with Data.'),
    dash_table.DataTable(data=df.to_dict('records'), page_size=10)
])

if __name__ == '__main__':
    app.run(debug=True)
