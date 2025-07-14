import pandas as pd  
from datetime import datetime  
import numpy as np  # 用于创建NaN值
from enum import Enum

# # 定义枚举类型  
# class TransactionType(Enum):  
#     房租水电 = '房租水电'  
#     生活开支 = '生活开支'
#     日用百货 = '日用百货'
#     交通出行 = '交通出行'
#     数码电器 = '数码电器'
#     文化休闲 = '文化休闲'
#     投资理财 = '投资理财'  
#     鞋帽服饰 = '鞋帽服饰'  
#     餐饮零食 = '餐饮零食'
#     消费还款 = '消费还款'
#     其他 = '其他'
# class Expense(Enum):  
#     收入 = '收入'  
#     支出 = '支出'
#     不计收支 = '不计收支'
#     待定 = '待定'
# class TransactionStatus(Enum):  
#     交易成功 = '交易成功'  
#     交易失败 = '交易失败'
#     待定 = '待定'

# 交易分类映射
transaction_type_mapping = {  
    '电脑办公': '数码电器',  
    '白条':'消费还款', 
    '其他':'其他',  
}
default_transaction_type = '其他'  # 或者选择其他默认值

# 交易方式映射
expense_mapping = {  
    '收入': '收入',  
    '支出': '支出',  
    '不计收支': '不计收支', 
    '待定': '待定'
}
default_expense = '待定'  # 或者选择其他默认值

# 交易状态映射
transaction_status_mapping = {  
    '交易成功': '交易成功',  
    '交易失败': '交易失败',
    '待定': '待定'  
}
default_transaction_status = '待定'  # 或者选择其他默认值
 
file_path = 'bill_20240513233706431_55.csv'  
df = pd.read_csv(file_path, skiprows=22, names=[  
    '交易时间', '交易分类', '商户名称', '交易说明', '收/支', '金额', '收/付款方式', '交易状态', '交易订单号', '商家订单号', '备注'  
])  
  
# 修改交易时间列，将其解析为"年/月/日 时:分:秒"格式  注意：原始CSV中的时间可能是带引号的字符串，需要去除引号  
df['交易时间'] = df['交易时间'].str.replace(r'\t', '', regex=True)
# df['交易时间'] = df['交易时间'].str.strip('"').apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S').strftime('%Y/%m/%d %H:%M:%S'))  
df['交易时间'] = pd.to_datetime(df['交易时间']).dt.strftime('%Y/%m/%d %H:%M:%S')  

# 应用映射并处理不在映射中的情况  
df['交易分类'] = df['交易分类'].map(transaction_type_mapping).fillna(default_transaction_type)
df['收/支'] = df['收/支'].map(expense_mapping).fillna(default_expense)
df['交易状态'] = df['交易状态'].map(transaction_status_mapping).fillna(default_transaction_status) 

# 创建新DataFrame ,对于不存在的数据，我们将其初始化为NaN  
columns = [  
    '交易日期', '交易类型', '分类细化', '交易对方', '商品',  
    '收/支', '金额', '交易方式', '交易状态', '交易订单号', '备注'  
]  
df_new = pd.DataFrame(columns=columns)
df_new['交易日期'] = df['交易时间']
df_new['交易类型'] = df['交易分类']
df_new['分类细化'] = np.nan
df_new['交易对方'] = df['商户名称']
df_new['商品'] = df['交易说明']
df_new['收/支'] = df['收/支']
df_new['金额'] = df['金额']
df_new['交易方式'] = df['收/付款方式']
df_new['交易状态'] = df['交易状态']
df_new['交易订单号'] = df['交易订单号']
df_new['备注'] = np.nan
print(df_new)

# 如果需要，你可以将None替换为字符串'未指定'或其他默认值  
# df_new.replace(None, '未指定', inplace=True)

# 将DataFrame保存到本地Excel文件  
file_path = 'E:\\生活\\账单\\scripts\\transactions.xlsx'  
df_new.to_excel(file_path, index=False, engine='openpyxl')  
  
print(f'Excel文件已保存到：{file_path}')