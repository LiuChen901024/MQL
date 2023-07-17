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

# 在每个 SKU 的级别上计算销售利润
df['销售利润'] = df['成交金额'] - df['成本'] * df['成交商品件数']

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
        ], style={'borderBottom': 'thin lightgrey solid', 'backgroundColor': 'rgb(250, 250, 250)',
                  'padding': '10px 5px'}),
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
        html.H3('各花纹的销售数量'),
        dcc.Graph(id='graph-output')
    ], style={'width': '100%', 'padding-top': '20px'}),
    html.Div([
        html.H3('访客数和转化率'),
        dcc.Graph(id='combo-graph-output')
    ], style={'width': '100%', 'padding-top': '20px'}),
    html.Div([
        html.H3('成交商品件数和件单价'),
        dcc.Graph(id='second-combo-graph-output')
    ], style={'width': '100%', 'padding-top': '20px'}),
    html.Div([
        html.H3('各花纹的销售利润'),
        dcc.Graph(id='profit-graph-output')
    ], style={'width': '100%', 'padding-top': '20px'}),
])



def get_total(grouped_df):
    total = grouped_df['成交商品件数'].sum()
    total_df = pd.DataFrame([['总计', total]], columns=['花纹', '成交商品件数'])
    return pd.concat([grouped_df, total_df], ignore_index=True)


def update_output(start_date, end_date, selected_patterns, selected_sizes, selected_dims, search_value):
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    if start_date > end_date:
        raise ValueError("结束日期不能早于开始日期")

    mask = (df['时间'] >= start_date) & (df['时间'] <= end_date)
    filtered_df = df.loc[mask]

    if 'All' not in selected_patterns:
        filtered_df = filtered_df[filtered_df['花纹'].isin(selected_patterns)]
    if 'All' not in selected_sizes:
        filtered_df = filtered_df[filtered_df['尺寸'].isin(selected_sizes)]
    if 'All' not in selected_dims:
        filtered_df = filtered_df[filtered_df['DIM'].isin(selected_dims)]
    if not selected_patterns:
        selected_patterns = ['All']
    if not selected_sizes:
        selected_sizes = ['All']
    if not selected_dims:
        selected_dims = ['All']

    if search_value:
        filtered_df = filtered_df[
            (filtered_df['SKU'].astype(str) == search_value) | (filtered_df['CAI'].astype(str) == search_value)]

    if filtered_df.empty:
        raise ValueError("没有找到匹配的数据")

    return filtered_df


def create_bar_figure(x, y):
    return go.Figure(data=[
        go.Bar(x=x, y=y, text=y, textposition='auto')
    ])


def create_combo_figure(x, y1, y2, y1_name, y2_name):
    return go.Figure(data=[
        go.Bar(name=y1_name, x=x, y=y1, text=y1, textposition='auto'),
        go.Scatter(name=y2_name, x=x, y=y2, yaxis='y2', mode='lines+markers', text=y2, textposition="top center"),
    ]).update_layout(
        yaxis=dict(
            title=y1_name,
            titlefont=dict(color="#1f77b4"),
            tickfont=dict(color="#1f77b4")
        ),
        yaxis2=dict(
            title=y2_name,
            titlefont=dict(color="#ff7f0e"),
            tickfont=dict(color="#ff7f0e"),
            anchor="free",
            overlaying="y",
            side="right",
            position=1
        )
    )

@app.callback(
    Output('graph-output', 'figure'),
    Output('combo-graph-output', 'figure'),
    Output('second-combo-graph-output', 'figure'),
    Output('profit-graph-output', 'figure'),
    [Input('start-date-picker', 'date'),
     Input('end-date-picker', 'date'),
     Input('pattern-filter', 'value'),
     Input('size-filter', 'value'),
     Input('dim-filter', 'value'),
     Input('search-box', 'value')]
)
def update_figures(start_date, end_date, selected_patterns, selected_sizes, selected_dims, search_value):
    try:
        filtered_df = update_output(start_date, end_date, selected_patterns, selected_sizes, selected_dims, search_value)
    except ValueError as e:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update

    grouped_df = filtered_df.groupby('花纹', as_index=False)['成交商品件数'].sum()
    total_grouped_df = get_total(grouped_df)

    fig = go.Figure(data=[
        go.Bar(x=total_grouped_df['花纹'], y=total_grouped_df['成交商品件数'], text=total_grouped_df['成交商品件数'], textposition='auto')
    ])

    grouped_df_by_date = filtered_df.groupby('时间', as_index=False).agg({'访客数': 'sum', '成交人数': 'sum'})
    grouped_df_by_date['转化率'] = grouped_df_by_date['成交人数'] / grouped_df_by_date['访客数']

    fig_combo = go.Figure(data=[
        go.Bar(name='访客数', x=grouped_df_by_date["时间"], y=grouped_df_by_date["访客数"],
               text=grouped_df_by_date["访客数"], textposition='auto'),
        go.Scatter(name='转化率', x=grouped_df_by_date["时间"], y=grouped_df_by_date["转化率"], yaxis='y2',
                   mode='lines+markers', text=grouped_df_by_date["转化率"].apply(lambda x: '{:.2%}'.format(x)),
                   textposition="top center"),
    ])

    fig_combo.update_layout(
        yaxis=dict(
            title='访客数',
            titlefont=dict(
                color="#1f77b4"
            ),
            tickfont=dict(
                color="#1f77b4"
            )
        ),
        yaxis2=dict(
            title='转化率',
            titlefont=dict(
                color="#ff7f0e"
            ),
            tickfont=dict(
                color="#ff7f0e"
            ),
            anchor="free",
            overlaying="y",
            side="right",
            position=1
        )
    )

    grouped_df_by_date_2 = filtered_df.groupby('时间', as_index=False).agg({'成交商品件数': 'sum', '成交金额': 'sum'})
    grouped_df_by_date_2['件单价'] = grouped_df_by_date_2['成交金额'] / grouped_df_by_date_2['成交商品件数']

    fig_combo_2 = go.Figure(data=[
        go.Bar(name='成交商品件数', x=grouped_df_by_date_2["时间"], y=grouped_df_by_date_2["成交商品件数"],
               text=grouped_df_by_date_2["成交商品件数"], textposition='auto'),
        go.Scatter(name='件单价', x=grouped_df_by_date_2["时间"], y=grouped_df_by_date_2["件单价"], yaxis='y2',
                   mode='lines+markers', text=grouped_df_by_date_2["件单价"].apply(lambda x: '{:.0f}'.format(x)),
                   textposition="top center"),
    ])

    fig_combo_2.update_layout(
        yaxis=dict(
            title='成交商品件数',
            titlefont=dict(
                color="#1f77b4"
            ),
            tickfont=dict(
                color="#1f77b4"
            )
        ),
        yaxis2=dict(
            title='件单价',
            titlefont=dict(
                color="#ff7f0e"
            ),
            tickfont=dict(
                color="#ff7f0e"
            ),
            anchor="free",
            overlaying="y",
            side="right",
            position=1
        )
    )

    # 计算每个花纹的销售利润
    grouped_df_profit = filtered_df.groupby('花纹', as_index=False)['销售利润'].sum()
    fig_profit = go.Figure(data=[
        go.Bar(name='销售利润', x=grouped_df_profit['花纹'], y=grouped_df_profit['销售利润'],
               text=grouped_df_profit['销售利润'], textposition='auto')
    ])

    # 返回图像
    return fig, fig_combo, fig_combo_2, fig_profit


if __name__ == '__main__':
    app.run_server(debug=True)