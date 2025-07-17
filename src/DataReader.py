# 数据读取模块
import pandas as pd
import re
import os
import glob
from src.config import BILL_COLUMNS, EXPENSE_MAPPING, DEFAULT_EXPENSE, TRANSACTION_STATUS_MAPPING, DEFAULT_TRANSACTION_STATUS, FORMAT_STR, TRANSACTION_STATUS_MAP

class DataReader:
    def __init__(self):
        pass

    def read_dir_data(self, path):
        """
        遍历指定路径下的所有CSV文件，根据文件名调用不同的处理函数，
        并将处理后的数据合并为一个DataFrame返回。
        """
        dfs = []

        # 遍历指定路径及其子目录下的所有文件
        for root, dirs, files in os.walk(path):
            for file in files:
                if file.endswith('.csv'):
                    file_path = os.path.join(root, file)

                    # 根据文件名中的关键词调用对应的处理函数
                    if 'bill' in file:
                        df = self.read_data_jd(file_path)
                    elif '微信' in file:
                        df = self.read_data_wx(file_path)
                    elif '支付宝' in file:
                        df = self.read_data_zfb(file_path)
                    else:
                        continue  # 忽略不匹配任何关键词的文件

                    if not df.empty:
                        dfs.append(df)

        # 合并所有处理后的DataFrame
        if dfs:
            return pd.concat(dfs, ignore_index=True)
        else:
            return pd.DataFrame()  # 返回空DataFrame，避免None引发后续错误

    def read_sigle_data(self,path):
        """读取数据"""
        if path.endswith('.csv'):
            if 'bill' in path:
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
        d_wx = d_wx.iloc[:, [0,1,2,3,7,6,4,5]]
        d_wx.insert(0, '来源', '微信')
        d_wx.insert(3, '类型细化', '待定')
        # 移除数据中的空格，避免后续处理时出现错误
        d_wx = self.strip_in_data(d_wx)
        # 重命名列名
        d_wx.columns = BILL_COLUMNS
        # 清洗数据
        d_wx = self.clean_data(d_wx)

        # 统一交易状态
        d_wx['交易状态'] = d_wx['交易状态'].apply(self.unify_transaction_status)
        # print("数据基本信息：")
        # d_wx.info()
        # print(d_wx.loc[0])
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
        d_zfb.insert(0, '来源', '支付宝')
        d_zfb.insert(3, '类型细化', '待定')
        # 重命名列名
        d_zfb.columns = BILL_COLUMNS
        # 清洗数据
        d_zfb = self.clean_data(d_zfb)

        # 统一交易状态
        d_zfb['交易状态'] = d_zfb['交易状态'].apply(self.unify_transaction_status)
        # print("数据基本信息：")
        # d_zfb.info()
        # print(d_zfb.loc[0])
        
        # 返回处理后的账单数据
        return d_zfb

    def read_data_jd(self,path):
        """读取京东账单"""
        # 读取CSV文件，跳过前21行，这些行通常包含不必要的信息
        d_jd = pd.read_csv(path, skiprows=21, encoding='utf-8')
        # 选择数据框中的特定列，这些列包含所需的信息
        d_jd = d_jd.iloc[:, [0,1,2,3,7,6,4,5]]
        d_jd.insert(0, '来源', '京东')
        d_jd.insert(3, '类型细化', '待定')
        # 移除数据中的空格，避免后续处理时出现错误
        d_jd = self.strip_in_data(d_jd)
        # 重命名列名
        d_jd.columns = BILL_COLUMNS
        # 清洗数据
        d_jd = self.clean_data(d_jd)

        # 统一交易状态
        d_jd['交易状态'] = d_jd['交易状态'].apply(self.unify_transaction_status)
        # print("数据基本信息：")
        # d_jd.info()
        # print(d_jd.loc[0])
        return d_jd

    def strip_in_data(self, data):
        """
        去除DataFrame中列名与数据中的首尾空格。
        
        参数:
        data: DataFrame - 需要处理的数据
        
        返回:
        DataFrame - 处理后的数据
        """
        # 移除列名中的前后空格
        data.columns = data.columns.str.strip()
        # 移除每个单元格中的前后空格
        data = data.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        return data
    
    def clean_data(self, df):
        """
        进一步清洗数据，包括标准化日期格式和金额字段
        
        参数:
        df: DataFrame - 需要处理的数据
        
        返回:
        DataFrame - 处理后的数据
        """
        # 清洗交易时间列
        df['交易时间'] = pd.to_datetime(df['交易时间'], errors='coerce').dt.strftime(FORMAT_STR)
        
        # 清洗金额列，移除非数字字符并转换为float
        def clean_amount(amount):
            # 确保输入是字符串
            if pd.isna(amount) or not isinstance(amount, (str, int, float)):
                return amount  # 或者返回一个默认值，取决于你的需求
            amount_str = str(amount)
            # 移除所有非数字及小数点字符
            cleaned_str = re.sub(r'[^\d.]', '', amount_str)
            try:
                return float(cleaned_str)
            except ValueError:
                # 如果无法转换为float，则返回NaN
                return None
        
        df['金额'] = df['金额'].apply(clean_amount)
        
        return df
    
    def unify_transaction_status(self,raw_status):
        """
        统一交易状态分类
        :param raw_status: 原始交易状态（如 "支付成功", "交易关闭"）
        :return: 统一后的分类（"交易成功" 或 "交易失败"）
        """
        return TRANSACTION_STATUS_MAP.get(raw_status, "交易失败")