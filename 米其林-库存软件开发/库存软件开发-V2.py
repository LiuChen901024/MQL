import os
import pandas as pd
import warnings
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output
from datetime import datetime

# 忽略openpyxl的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 读取主表
template_file = r'C:\GPT\米其林库存监测表.xlsx'
with pd.ExcelFile(template_file) as xls:
    template_df = pd.read_excel(xls, sheet_name='原始数据')
    dictionary_df = pd.read_excel(xls, sheet_name='字典')

# 遍历指定目录下的所有Excel文件
folder = r'C:\GPT\米其林库存'
for filename in os.listdir(folder):
    if filename.endswith('.xlsx'):
        # 读取数据
        df = pd.read_excel(os.path.join(folder, filename))
        # 将数据添加到主表中
        template_df = pd.concat([template_df, df], ignore_index=True)

# 去重
template_df.drop_duplicates(subset=['时间', 'SKU'], keep='first', inplace=True)

# 筛选 "是否影分身" 列，删除选项 “是” 的行数据
template_df = template_df.query("`是否影分身` != '是'")

# 筛选 "上下柜状态" 列，删除选项 “下柜” 和 “可上柜” 的行数据
template_df = template_df.query("`上下柜状态` != '下柜' and `上下柜状态` != '可上柜'")

# 根据SKU从字典表中匹配数据，并替换主表中的数据
template_df = template_df.merge(dictionary_df[['SKU', 'CAI', '花纹', '尺寸', 'DIM']], on='SKU', how='left')

# 将新数据插入到原有列之间
column_order = template_df.columns.tolist()
for col in ['CAI', '花纹', '尺寸', 'DIM']:
    column_order.insert(column_order.index('品牌'), column_order.pop(column_order.index(col)))

template_df = template_df[column_order]

# 生成新文件名，包含当前日期
new_file = r'C:\GPT\米其林库存监测表—' + datetime.now().strftime('%Y%m%d') + '.xlsx'

# 使用 ExcelWriter 保存多个表到同一个 Excel 文件中
with pd.ExcelWriter(new_file) as writer:
    template_df.to_excel(writer, index=False, sheet_name='原始数据')
    dictionary_df.to_excel(writer, index=False, sheet_name='字典')

app = Dash(__name__)

df = pd.read_excel(new_file, sheet_name='原始数据')
df['时间'] = pd.to_datetime(df['时间'])

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
    dash_table.DataTable(
        id='table',
        columns=[{"name": i, "id": i} for i in df.columns],
        data=df.to_dict('records'),
    )
])

@app.callback(
    Output('output-container-date-picker-range', 'children'),
    Output('table', 'data'),
    [Input('my-date-picker-range', 'start_date'),
     Input('my-date-picker-range', 'end_date')])
def update_output(start_date, end_date):
    dff = df.loc[(df['时间'] >= start_date) & (df['时间'] <= end_date)]
    data = dff.to_dict('records')
    return f"Start Date: {start_date} End Date: {end_date}", data


if __name__ == '__main__':
    app.run_server(debug=True)
