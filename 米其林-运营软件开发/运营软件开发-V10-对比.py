import dash
from dash import dcc
from dash import html
import pandas as pd
import plotly.graph_objects as go
from dash.dependencies import Input, Output
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
            html.Label('开始日期1'),
            dcc.DatePickerSingle(
                id='start-date-picker1',
                min_date_allowed=df['时间'].min(),
                max_date_allowed=df['时间'].max(),
                initial_visible_month=df['时间'].min(),
                date=df['时间'].min()
            ),
            html.Label('结束日期1'),
            dcc.DatePickerSingle(
                id='end-date-picker1',
                min_date_allowed=df['时间'].min(),
                max_date_allowed=df['时间'].max(),
                initial_visible_month=df['时间'].max(),
                date=df['时间'].max()
            ),
            html.Label('开始日期2'),
            dcc.DatePickerSingle(
                id='start-date-picker2',
                min_date_allowed=df['时间'].min(),
                max_date_allowed=df['时间'].max(),
                initial_visible_month=df['时间'].min(),
                date=df['时间'].min()
            ),
            html.Label('结束日期2'),
            dcc.DatePickerSingle(
                id='end-date-picker2',
                min_date_allowed=df['时间'].min(),
                max_date_allowed=df['时间'].max(),
                initial_visible_month=df['时间'].max(),
                date=df['时间'].max()
            )
        ], style={'borderBottom': 'thin lightgrey solid', 'backgroundColor': 'rgb(250, 250, 250)',
                  'padding': '10px 5px'}),

    ]),
    html.Div([
        html.H3('各花纹的销售数量'),
        dcc.Graph(id='graph-output')
    ], style={'width': '100%', 'padding-top': '20px'}),
])

def filter_data(start_date, end_date):
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    if start_date > end_date:
        raise ValueError("结束日期不能早于开始日期")

    mask = (df['时间'] >= start_date) & (df['时间'] <= end_date)
    filtered_df = df.loc[mask]

    if filtered_df.empty:
        raise ValueError("没有找到匹配的数据")

    return filtered_df

@app.callback(
    Output('graph-output', 'figure'),
    [Input('start-date-picker1', 'date'),
     Input('end-date-picker1', 'date'),
     Input('start-date-picker2', 'date'),
     Input('end-date-picker2', 'date')]
)
def update_output(start_date1, end_date1, start_date2, end_date2):
    try:
        filtered_df1 = filter_data(start_date1, end_date1)
        filtered_df2 = filter_data(start_date2, end_date2)
    except ValueError as e:
        raise PreventUpdate

    grouped_df1 = filtered_df1.groupby('花纹', as_index=False)['成交商品件数'].sum()
    grouped_df2 = filtered_df2.groupby('花纹', as_index=False)['成交商品件数'].sum()

    fig = go.Figure(data=[
        go.Bar(name='时间范围1', x=grouped_df1['花纹'], y=grouped_df1['成交商品件数'], text=grouped_df1['成交商品件数'], textposition='auto'),
        go.Bar(name='时间范围2', x=grouped_df2['花纹'], y=grouped_df2['成交商品件数'], text=grouped_df2['成交商品件数'], textposition='auto')
    ])

    fig.update_layout(barmode='group')

    return fig

if __name__ == '__main__':
    app.run_server(debug=True)
