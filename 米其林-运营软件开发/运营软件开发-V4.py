import dash
from dash import dcc
from dash import html
import pandas as pd
import plotly.graph_objects as go
from dash.dependencies import Input, Output, State
from dash.exceptions import PreventUpdate

# 读取数据
df = pd.read_excel(r'C:\GPT\米其林销售表-by day—' + pd.to_datetime('today').strftime('%Y%m%d') + '.xlsx')

# 确保 '时间' 列都是 datetime 对象
df['时间'] = pd.to_datetime(df['时间'])

# 创建 Dash 应用
app = dash.Dash(__name__)

# 获取筛选项的唯一值
patterns = df['花纹'].unique().tolist()
sizes = df['尺寸'].unique().tolist()
dims = df['DIM'].unique().tolist()

# 定义布局
app.layout = html.Div([
    html.Div([
        html.Div([
            html.Div([
                html.Label('开始日期'),
                dcc.DatePickerSingle(
                    id='start-date-picker',
                    min_date_allowed=df['时间'].min(),
                    max_date_allowed=df['时间'].max(),
                    initial_visible_month=df['时间'].min(),
                    date=df['时间'].min()
                ),
            ], style={'width': '48%', 'display': 'inline-block'}),
            html.Div([
                html.Label('结束日期'),
                dcc.DatePickerSingle(
                    id='end-date-picker',
                    min_date_allowed=df['时间'].min(),
                    max_date_allowed=df['时间'].max(),
                    initial_visible_month=df['时间'].max(),
                    date=df['时间'].max()
                ),
            ], style={'width': '48%', 'float': 'right', 'display': 'inline-block'})
        ], style={'borderBottom': 'thin lightgrey solid', 'backgroundColor': 'rgb(250, 250, 250)', 'padding': '10px 5px'}),
        html.Div([
            html.Div([
                html.Label('花纹'),
                dcc.Dropdown(
                    id='pattern-filter',
                    options=[{'label': 'All', 'value': 'All'}] + [{'label': i, 'value': i} for i in patterns],
                    value='All',
                    multi=True
                ),
            ], style={'width': '30%', 'display': 'inline-block'}),
            html.Div([
                html.Label('尺寸'),
                dcc.Dropdown(
                    id='size-filter',
                    options=[{'label': 'All', 'value': 'All'}] + [{'label': i, 'value': i} for i in sizes],
                    value='All',
                    multi=True
                ),
            ], style={'width': '30%', 'display': 'inline-block'}),
            html.Div([
                html.Label('DIM'),
                dcc.Dropdown(
                    id='dim-filter',
                    options=[{'label': 'All', 'value': 'All'}] + [{'label': i, 'value': i} for i in dims],
                    value='All',
                    multi=True
                ),
            ], style={'width': '30%', 'float': 'right', 'display': 'inline-block'})
        ], style={'padding': '10px 5px'}),
        html.Div([
            html.Label('SKU / CAI 搜索'),
            dcc.Input(
                id='search-box',
                type='text',
                placeholder='输入 SKU 或 CAI'
            ),
        ], style={'padding': '10px 5px'}),
    ], style={'width': '48%', 'display': 'inline-block', 'vertical-align': 'top'}),
    html.Div([
        dcc.Graph(id='graph-output')
    ], style={'width': '100%', 'padding-top': '20px'}),
    html.Div([
        dcc.Graph(id='combo-graph-output')
    ], style={'width': '100%', 'padding-top': '20px'}),
])

@app.callback(
    Output('graph-output', 'figure'),
    Output('combo-graph-output', 'figure'),
    [Input('start-date-picker', 'date'),
     Input('end-date-picker', 'date'),
     Input('pattern-filter', 'value'),
     Input('size-filter', 'value'),
     Input('dim-filter', 'value'),
     Input('search-box', 'value')]
)
def update_output(start_date, end_date, selected_patterns, selected_sizes, selected_dims, search_value):
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    mask = (df['时间'] >= start_date) & (df['时间'] <= end_date)
    filtered_df = df.loc[mask]

    if 'All' not in selected_patterns:
        filtered_df = filtered_df[filtered_df['花纹'].isin(selected_patterns)]
    if 'All' not in selected_sizes:
        filtered_df = filtered_df[filtered_df['尺寸'].isin(selected_sizes)]
    if 'All' not in selected_dims:
        filtered_df = filtered_df[filtered_df['DIM'].isin(selected_dims)]

    if search_value:
        filtered_df = filtered_df[(filtered_df['SKU'].astype(str) == search_value) | (filtered_df['CAI'].astype(str) == search_value)]

    grouped_df = filtered_df.groupby('花纹', as_index=False)['成交商品件数'].sum()

    total = grouped_df['成交商品件数'].sum()

    total_row = pd.DataFrame({'花纹': ['Total'], '成交商品件数': [total]})
    grouped_df = pd.concat([grouped_df, total_row], ignore_index=True)

    fig = go.Figure([go.Bar(x=grouped_df["花纹"], y=grouped_df["成交商品件数"], text=grouped_df["成交商品件数"], textposition='auto')])

    fig.update_layout(title_text='Sales by Pattern')

    # 复合图
    grouped_df_by_date = filtered_df.groupby('时间', as_index=False).agg({'访客数': 'sum', '成交人数': 'sum'})
    grouped_df_by_date['转化率'] = grouped_df_by_date['成交人数'] / grouped_df_by_date['访客数']

    fig_combo = go.Figure(data=[
        go.Bar(name='访客数', x=grouped_df_by_date["时间"], y=grouped_df_by_date["访客数"], text=grouped_df_by_date["访客数"], textposition='auto'),
        go.Scatter(name='成交转化率', x=grouped_df_by_date["时间"], y=grouped_df_by_date["转化率"], yaxis='y2', text=[f"{i:.2%}" for i in grouped_df_by_date["转化率"]], textposition='top center')
    ])

    fig_combo.update_layout(
        title_text='访客数 & 成交转化率',
        yaxis=dict(
            title='访客数',
        ),
        yaxis2=dict(
            title='成交转化率',
            overlaying='y',
            side='right'
        ),
    )

    return fig, fig_combo

if __name__ == '__main__':
    app.run_server(debug=True)