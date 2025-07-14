# 数据处理模块
import pandas as pd
from config import BILL_COLUMNS

def handlefunc(init_data):
    """交易类型修正"""
    for index, row in init_data.iterrows():
        if row['交易类型'] == '餐饮美食':
            row['交易类型'] = '餐饮零食'
        elif row['交易类型'] == '商户消费' and row['交易对方'] == 'LAWSON':
            row['交易类型'] = '餐饮零食'
        else:
            row['交易状态'] = '其他'
            print(f"第{index}行数据有误，请检查！")
    return init_data

def handle_total_data(data):
    """清洗数据"""
    # 删除无效列和行
    if '交易单号' in data.columns:
        data = data.drop(columns=['交易单号'])
    data = data[data['收/支'] != '/'].reset_index(drop=True)
    conditions_to_drop = ['等待确认收货', '提现已到账', '已全额退款', '已退款','充值完成', '退款成功', '还款成功', '交易关闭']
    data = data[~data['交易状态'].isin(conditions_to_drop)].reset_index(drop=True)
    # 更新交易状态
    data.loc[data['交易状态'].isin(['对方已收钱', '已存入零钱', '支付成功', '已转账', '已收钱']), '交易状态'] = '交易成功'
    return data