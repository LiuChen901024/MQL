import os
import pandas as pd
import numpy as np
import warnings
from datetime import datetime

# 忽略openpyxl的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# 读取主表
template_file = r'C:\GPT\米其林销售表-by day.xlsx'
with pd.ExcelFile(template_file) as xls:
    template_df = pd.read_excel(xls, sheet_name='原始数据')
    dictionary_df = pd.read_excel(xls, sheet_name='字典')
    cost_df = pd.read_excel(xls, sheet_name='采购成本')  # 读取 "采购成本" 表

# 遍历指定目录下的所有Excel文件
folder = r'C:\GPT\米其林销售-by day'
for filename in os.listdir(folder):
    if filename.endswith('.xlsx'):
        # 读取数据
        df = pd.read_excel(os.path.join(folder, filename))
        # 将数据添加到主表中
        template_df = pd.concat([template_df, df], ignore_index=True)

# 去重
template_df.drop_duplicates(subset=['时间', 'SKU'], keep='first', inplace=True)

# 筛选 "成交商品件数" 列，删除选项 “0” 的行数据
template_df = template_df.query("`成交商品件数` != 0")

# 根据SKU从字典表中匹配数据，并替换主表中的数据
template_df = template_df.merge(dictionary_df[['SKU', 'CAI', '花纹', '尺寸', 'DIM']], on='SKU', how='left')

# 根据SKU将采购成本数据添加到主表中
template_df = template_df.merge(cost_df[['SKU', '7月成本']].rename(columns={'7月成本': '成本'}), on='SKU', how='left')

# 找出 "成本" 列中为 NaN 或 0 的行，并用 "成交金额/成交商品件数" 代替
template_df.loc[(template_df['成本'].isnull()) | (template_df['成本'] == 0), '成本'] = template_df['成交金额'] / template_df['成交商品件数']

# 将新数据插入到原有列之间
column_order = template_df.columns.tolist()
for col in ['CAI', '花纹', '尺寸', 'DIM', '成本']:  # '成本' 也需要插入
    column_order.insert(column_order.index('品牌'), column_order.pop(column_order.index(col)))

template_df = template_df[column_order]

# 生成新文件名，包含当前日期
new_file = r'C:\GPT\米其林销售表-by day—' + datetime.now().strftime('%Y%m%d') + '.xlsx'

# 使用 ExcelWriter 保存多个表到同一个 Excel 文件中
with pd.ExcelWriter(new_file) as writer:
    template_df.to_excel(writer, index=False, sheet_name='原始数据')
    dictionary_df.to_excel(writer, index=False, sheet_name='字典')
