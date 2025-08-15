import os
import glob
from src.DataReader import DataReader
from src.DataProcessor import DataProcessor
from src.config import INIT_BILL_PATH, OUTPUT_DIR, SAVE_PATH
import sys
import argparse

def main():
    # 创建参数解析器
    parser = argparse.ArgumentParser(description='账单数据处理脚本')
    
    # 添加命令行选项
    parser.add_argument('-d', action='store_true', 
                        help='从配置的默认路径读取预处理的CSV文件，写入Excel并生成图表')
    parser.add_argument('-t', action='store_true', 
                        help='使用配置中的训练数据训练模型并保存')
    parser.add_argument('-n', action='store_true', 
                        help='将人工审核的账单数据添加到训练数据集')
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 根据参数执行相应操作
    if args.d:
        print("执行CSV转Excel并生成图表...")
        # 实际应用中调用对应的函数
        # process_csv_to_excel()
    
    if args.t:
        print("训练模型并保存...")
        # 实际应用中调用对应的函数
        # train_model()
    
    if args.n:
        print("添加人工审核数据到训练集...")
        # 实际应用中调用对应的函数
        # add_manual_data()
    # 读取数据
    zfb_files = r'E:\生活\账单\scripts\data\支付宝交易明细(20250401-20250430).csv'
    wx_files = r'E:\myselfProgram\billHanler\BillAssistant\data\微信支付账单(20250401-20250430)——【解压密码可在微信支付公众号查看】.csv'
    jd_files = r'E:\myselfProgram\billHanler\BillAssistant\data\bill_20240303111234011_712.csv'
    data_dir = r'E:\生活\账单\scripts\data'
    data_reader = DataReader()

    # 如果执行脚本输入参数，例如 reprocess 即：python main.py --reprocess，则重新载入修改后的账单数据，重新生成账单图片（主要修改带人工确认选项）
    df_total = data_reader.read_dir_data(data_dir)
    # print("数据基本信息：")
    # df_total.info()
    # print(df_total.loc[0])
    # data_processor = DataProcessor()
    # data_processor.process_data(df_total)

if __name__ == "__main__":
    main()