# 数据读取模块
import pandas as pd
import os
import glob
from src.config import BILL_COLUMNS, EXPENSE_MAPPING, DEFAULT_EXPENSE, TRANSACTION_STATUS_MAPPING, DEFAULT_TRANSACTION_STATUS, FORMAT_STR

class DataReader:
    def __init__(self):
        return
    def read_data(self,path):
        """读取数据"""
        if path.endswith('.csv'):
            if '京东' in path:
                return self.read_data_jd(path)
            elif '微信' in path:
                return self.read_data_wx(path)
            elif '支付宝' in path:
                return self.read_data_zfb(path)
    
    def read_data_wx(self,path):
        """
        读取微信账单
        该函数读取微信账单数据，进行必要的数据清洗和格式化，以便后续处理和分析
        参数:
        - path: 微信账单文件的路径，用于指定读取文件的位置
        返回值:
        - d_wx: 清洗和格式化后的微信账单数据框
        注意：
        - 该函数会自动跳过前16行，这些行通常包含不重要的信息，如账单的标题和描述。
        - 该函数会自动将交易时间列转换为日期时间格式，以便于后续处理。
        - 该函数会自动将收/支列转换为枚举类型，并使用默认值填充缺失值。
        """
        # 读取CSV文件，跳过前16行，这些行通常包含不必要的信息
        d_wx = pd.read_csv(path, skiprows=16, encoding='utf-8')
        # 选择数据框中的特定列，这些列包含所需的信息
        d_wx = d_wx.iloc[:, [0,1,2,3,4,5,6,7,8,10]]
        # 移除数据中的空格，以确保数据一致性
        d_wx = self.strip_in_data(d_wx)
        # 将交易时间列转换为日期时间格式，以便于后续处理
        d_wx.iloc[:, 0] = d_wx.iloc[:, 0].astype('datetime64[ns]')
        # 重命名列，使其更具可读性和一致性
        d_wx.rename(columns={'当前状态': '交易状态','支付方式': '交易方式', '金额(元)': '金额'}, inplace=True)
        # 格式化交易时间，确保时间格式的一致性
        d_wx['交易时间'] = pd.to_datetime(d_wx['交易时间']).dt.strftime(FORMAT_STR)
        # 插入来源列，标识数据来自微信
        d_wx.insert(1, '来源', "微信")
        # 返回清洗和格式化后的数据框
        return d_wx

    def read_data_zfb(self,path):
        """
        读取支付宝账单
        
        本函数通过 pandas 读取支付宝账单数据，进行数据预处理并返回处理后的 DataFrame 对象
        
        参数:
        path: str - 支付宝账单文件的路径
        
        返回:
        DataFrame - 处理后的账单数据，格式按照要求统一
        """
        # 读取支付宝账单数据，跳过25行无用信息，(skiprows从0开始计数，0-24索引不会读取，从25索引行即实际第26行开始读取)使用 gbk 编码
        d_zfb = pd.read_csv(path, skiprows=24, encoding='gbk')

        # 移除数据中的空格，避免后续处理时出现错误
        d_zfb = self.strip_in_data(d_zfb)

        # 选择特定的列进行后续处理
        d_zfb = d_zfb.iloc[:, [0,1,2,4,8,7,5,6]]
        d_zfb.insert(0, None, '支付宝')
        d_zfb.insert(3, '类型细化', '待定')
        print("数据基本信息：")
        d_zfb.info()
        
        # # 将交易时间转换为 datetime 类型，便于时间处理
        # d_zfb.iloc[:, 0] = d_zfb.iloc[:, 0].astype('datetime64[ns]')
        
        # # 将交易金额转换为 float 类型，便于后续的数值计算
        # d_zfb.iloc[:, 5] = d_zfb.iloc[:, 5].astype('float64')
        
        # # 重命名列名，使得数据更加直观易懂
        # d_zfb.rename(columns={'交易分类': '交易类型','商品说明': '商品', '收/付款方式': '交易方式','交易订单号':'交易单号'}, inplace=True)
        
        # # 格式化交易时间，统一时间格式
        # d_zfb['交易时间'] = pd.to_datetime(d_zfb['交易时间']).dt.strftime(FORMAT_STR)
        
        # # 插入来源列，标记数据来源于支付宝
        # d_zfb.insert(1, '来源', "支付宝")
        
        # # 返回处理后的账单数据
        # return d_zfb

    def read_data_jd(file_path):
        """读取京东账单"""
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        processed = [lines[21]] + [line.replace('\t', '') for line in lines[22:]]
        df = pd.read_csv(StringIO(''.join(processed)), index_col=False)
        df = df.iloc[:, [0,1,2,3,4,5,6,7, 8, 9,10]]
        df['交易时间'] = pd.to_datetime(df['交易时间'], format=FORMAT_STR, errors='coerce').dt.strftime(FORMAT_STR)
        df['收/支'] = df['收/支'].map(EXPENSE_MAPPING).fillna(DEFAULT_EXPENSE)
        df['交易状态'] = df['交易状态'].map(TRANSACTION_STATUS_MAPPING).fillna(DEFAULT_TRANSACTION_STATUS)
        df_new = pd.DataFrame(columns=BILL_COLUMNS)
        df_new['交易时间'] = df['交易时间']
        df_new['交易类型'] = df['交易分类']
        df_new['交易对方'] = df['商户名称']
        df_new['商品'] = df['交易说明']
        df_new['收/支'] = df['收/支']
        df_new['金额'] = df['金额']
        df_new['交易方式'] = df['收/付款方式']
        df_new['交易状态'] = df['交易状态']
        df_new['交易单号'] = df['交易订单号']
        df_new.insert(1, '来源', "京东")
        return df_new

    def strip_in_data(self,data):
        """清理数据"""
        data = data.rename(columns={col: col.strip() for col in data.columns})
        if data.iloc[:, 6].dtype == 'object':
            data.iloc[:, 6] = data.iloc[:, 6].str.strip().str.replace('¥', '').astype('float64')
        return data