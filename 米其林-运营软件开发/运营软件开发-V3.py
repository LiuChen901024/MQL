import dash
from dash import dcc
from dash import html
import pandas as pd
import plotly.graph_objects as go

# 读取数据
df = pd.read_excel(r'C:\GPT\米其林销售表-by day—' + pd.to_datetime('today').strftime('%Y%m%d') + '.xlsx')

# 确保 '时间' 列都是 datetime 对象
df['时间'] = pd.to_datetime(df['时间'])

# 创建 Dash 应用
app = dash.Dash(__name__)

# 定义布局
app.layout = html.Div([
    html.Div([
        dcc.DatePickerSingle(
            id='start-date-picker',
            min_date_allowed=df['时间'].min(),
            max_date_allowed=df['时间'].max(),
            initial_visible_month=df['时间'].min(),
            date=df['时间'].min()
        ),
        dcc.DatePickerSingle(
            id='end-date-picker',
            min_date_allowed=df['时间'].min(),
            max_date_allowed=df['时间'].max(),
            initial_visible_month=df['时间'].max(),
            date=df['时间'].max()
        )
    ]),
    dcc.Graph(id='graph-output')
])

# 定义回调函数，根据选定的日期范围更新图形
@app.callback(
    dash.dependencies.Output('graph-output', 'figure'),
    [dash.dependencies.Input('start-date-picker', 'date'),
     dash.dependencies.Input('end-date-picker', 'date')]
)
def update_output(start_date, end_date):
    # 将输入的日期字符串转换为 datetime 对象
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)

    # 筛选日期范围
    mask = (df['时间'] >= start_date) & (df['时间'] <= end_date)
    filtered_df = df.loc[mask]

    # 对 '花纹' 列进行汇总
    grouped_df = filtered_df.groupby('花纹', as_index=False)['成交商品件数'].sum()

    # 计算总和
    total = grouped_df['成交商品件数'].sum()

    # 添加总和行
    total_row = pd.DataFrame({'花纹': ['Total'], '成交商品件数': [total]})
    grouped_df = pd.concat([grouped_df, total_row], ignore_index=True)

    # 创建柱状图
    fig = go.Figure([go.Bar(x=grouped_df["花纹"], y=grouped_df["成交商品件数"], text=grouped_df["成交商品件数"], textposition='auto')])

    fig.update_layout(title_text='Sales by Pattern')

    return fig

# 启动服务器
if __name__ == '__main__':
    app.run_server(debug=True)
