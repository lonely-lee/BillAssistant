import os
import glob
from src.DataReader import DataReader
from src.DataProcessor import DataProcessor
from src.config import INIT_BILL_PATH, OUTPUT_DIR, SAVE_PATH

def main():
    
    # 读取数据
    zfb_files = r'E:\myselfProgram\billHanler\BillAssistant\data\支付宝交易明细(20250401-20250430).csv'
    wx_files = r'E:\myselfProgram\billHanler\BillAssistant\data\微信支付账单(20250401-20250430)——【解压密码可在微信支付公众号查看】.csv'
    jd_files = r'E:\myselfProgram\billHanler\BillAssistant\data\bill_20240303111234011_712.csv'
    data_dir = r'E:\myselfProgram\billHanler\BillAssistant\data'
    data_reader = DataReader()
    df_total = data_reader.read_dir_data(data_dir)
    print("数据基本信息：")
    df_total.info()
    print(df_total.loc[0])
    data_processor = DataProcessor()
    data_processor.process_data(df_total)

if __name__ == "__main__":
    main()