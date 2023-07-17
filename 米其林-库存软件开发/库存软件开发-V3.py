import os
import pandas as pd
import warnings
import plotly.graph_objects as go
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output
from datetime import datetime

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

new_file = r'C:\GPT\米其林库存监测表—' + datetime.now().strftime('%Y%m%d') + '.xlsx'
df = pd.read_excel(new_file, sheet_name='原始数据')
df['时间'] = pd.to_datetime(df['时间'])

app = Dash(__name__)

app.layout = html.Div([
    dcc.DatePickerRange(
        id='my-date-picker-range',
        min_date_allowed=df['时间'].min(),
        max_date_allowed=df['时间'].max(),
        initial_visible_month=df['时间'].min(),
        start_date=df['时间'].min(),
        end_date=df['时间'].max()
    ),
    html.Div(id='output-container-date-picker-range'),
    dcc.Graph(id='bar-chart'),
    dash_table.DataTable(
        id='table',
        columns=[{"name": i, "id": i} for i in df.columns],
        data=df.to_dict('records'),
    )
])

@app.callback(
    Output('output-container-date-picker-range', 'children'),
    Output('table', 'data'),
    Output('bar-chart', 'figure'),
    [Input('my-date-picker-range', 'start_date'),
     Input('my-date-picker-range', 'end_date')])
def update_output(start_date, end_date):
    dff = df.loc[(df['时间'] >= pd.to_datetime(start_date)) & (df['时间'] <= pd.to_datetime(end_date))]
    data = dff.to_dict('records')

    summary = dff.groupby('花纹').agg({'全国现货库存': 'sum', '全国采购在途数量': 'sum', '全国昨日出库商品件数': 'sum'})

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=summary.index,
        y=summary['全国现货库存'],
        name='全国现货库存',
    ))
    fig.add_trace(go.Bar(
        x=summary.index,
        y=summary['全国采购在途数量'],
        name='全国采购在途数量',
    ))
    fig.add_trace(go.Bar(
        x=summary.index,
        y=summary['全国昨日出库商品件数'],
        name='全国昨日出库商品件数',
    ))
    fig.update_layout(barmode='group')

    return f"Start Date: {start_date} End Date: {end_date}", data, fig


if __name__ == '__main__':
    app.run_server(debug=True)
