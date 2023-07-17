import os
import pandas as pd
import warnings
import plotly.graph_objects as go
from dash import Dash, dcc, html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output, State
from datetime import datetime

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

new_file = r'C:\GPT\米其林库存监测表—' + datetime.now().strftime('%Y%m%d') + '.xlsx'
df = pd.read_excel(new_file, sheet_name='原始数据')
df['时间'] = pd.to_datetime(df['时间'])

# Convert SKU and CAI to string for search
df['SKU'] = df['SKU'].astype(str)
df['CAI'] = df['CAI'].astype(str)

# Create a new dataframe with all dates
all_dates = pd.date_range(start=df['时间'].min(), end=df['时间'].max())
df_all_dates = pd.DataFrame(all_dates, columns=['时间'])

app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            html.Div([
                dbc.Label('Start Date: '),
                dcc.DatePickerSingle(
                    id='start_date',
                    min_date_allowed=df['时间'].min(),
                    max_date_allowed=df['时间'].max(),
                    initial_visible_month=df['时间'].min(),
                    date=df['时间'].min()
                ),
                dbc.Label('End Date: '),
                dcc.DatePickerSingle(
                    id='end_date',
                    min_date_allowed=df['时间'].min(),
                    max_date_allowed=df['时间'].max(),
                    initial_visible_month=df['时间'].min(),
                    date=df['时间'].max()
                ),
                dbc.Label('Flower Patterns: '),
                dcc.Dropdown(
                    id='flower-pattern',
                    options=[{'label': 'ALL', 'value': 'ALL'}] + [{'label': i, 'value': i} for i in
                                                                  df['花纹'].unique()],
                    value=['ALL'],
                    multi=True,
                    placeholder="Select Flower Patterns",
                ),
                dbc.Label('Sizes: '),
                dcc.Dropdown(
                    id='size',
                    options=[{'label': 'ALL', 'value': 'ALL'}] + [{'label': i, 'value': i} for i in
                                                                  df['尺寸'].unique()],
                    value=['ALL'],
                    multi=True,
                    placeholder="Select Sizes",
                ),
                dbc.Label('DIMs: '),
                dcc.Dropdown(
                    id='dim',
                    options=[{'label': 'ALL', 'value': 'ALL'}] + [{'label': i, 'value': i} for i in df['DIM'].unique()],
                    value=['ALL'],
                    multi=True,
                    placeholder="Select DIMs",
                ),
                dbc.Label('Search SKU or CAI: '),
                dcc.Input(
                    id='search_input',
                    type='text',
                    placeholder='Enter SKU or CAI',
                )
            ]),
            html.Div(id='output-container-date-picker-range'),
        ], width=3),
        dbc.Col([
            dcc.Graph(id='bar-chart'),
            dcc.Graph(id='city-bar-chart'),
        ], width=9),
    ]),
], fluid=True)


@app.callback(
    [Output('output-container-date-picker-range', 'children'),
     Output('bar-chart', 'figure'),
     Output('city-bar-chart', 'figure')],
    [Input('start_date', 'date'),
     Input('end_date', 'date'),
     Input('flower-pattern', 'value'),
     Input('size', 'value'),
     Input('dim', 'value'),
     Input('search_input', 'value')])
def update_output(start_date, end_date, flower_patterns, sizes, dims, search_input):
    dff = pd.merge(df_all_dates, df, how='left', on='时间')
    dff.fillna(0, inplace=True)  # Fill NA values with 0
    dff = dff.loc[(dff['时间'] >= pd.to_datetime(start_date)) & (dff['时间'] <= pd.to_datetime(end_date))]

    if 'ALL' not in flower_patterns:
        dff = dff[dff['花纹'].isin(flower_patterns)]
    if 'ALL' not in sizes:
        dff = dff[dff['尺寸'].isin(sizes)]
    if 'ALL' not in dims:
        dff = dff[dff['DIM'].isin(dims)]

    if search_input:
        search_input = str(search_input)  # Convert search input to string
        dff = dff[(dff['SKU'].str.contains(search_input)) | (dff['CAI'].str.contains(search_input))]

    summary = dff.groupby('花纹').agg({'全国现货库存': 'sum', '全国采购在途数量': 'sum', '全国昨日出库商品件数': 'sum'})
    summary.loc['Total'] = summary.sum()  # Add a total row

    city_columns = ['北京现货库存', '上海现货库存', '广州现货库存', '成都现货库存', '武汉现货库存', '沈阳现货库存',
                    '西安现货库存', '德州现货库存']
    city_summary = dff.groupby('花纹')[city_columns].sum()
    city_summary.loc['Total'] = city_summary.sum()  # Add a total row

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=summary.index,
        y=summary['全国现货库存'],
        name='全国现货库存',
        text=summary['全国现货库存'],
        textposition='auto',
    ))
    fig.add_trace(go.Bar(
        x=summary.index,
        y=summary['全国采购在途数量'],
        name='全国采购在途数量',
        text=summary['全国采购在途数量'],
        textposition='auto',
    ))
    fig.add_trace(go.Bar(
        x=summary.index,
        y=summary['全国昨日出库商品件数'],
        name='全国昨日出库商品件数',
        text=summary['全国昨日出库商品件数'],
        textposition='auto',
    ))
    fig.update_layout(barmode='group')

    city_fig = go.Figure()
    for city in city_columns:
        city_fig.add_trace(go.Bar(
            x=city_summary.index,
            y=city_summary[city],
            name=city,
            text=city_summary[city],
            textposition='auto',
        ))
    city_fig.update_layout(barmode='group')

    return f"Start Date: {start_date} End Date: {end_date}", fig, city_fig


if __name__ == '__main__':
    app.run_server(debug=True)
