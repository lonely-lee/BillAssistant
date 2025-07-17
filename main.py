import os
import glob
from src.DataReader import DataReader
from src.DataProcessor import DataProcessor
# from src.processor import handlefunc, handle_total_data
# from src.exporter import save_to_excel
from src.config import INIT_BILL_PATH, OUTPUT_DIR, SAVE_PATH

def main():
    print("running")
    # # 创建输出目录
    # if not os.path.exists(OUTPUT_DIR):
    #     os.makedirs(OUTPUT_DIR)
    #     print(f"文件夹 {OUTPUT_DIR} 创建成功")
    
    # 获取文件路径
    # wx_files = glob.glob(os.path.join(INIT_BILL_PATH, '微信支付账单*.csv'))
    # zfb_files = glob.glob(os.path.join(INIT_BILL_PATH, 'alipay*.csv'))
    # jd_files = glob.glob(os.path.join(INIT_BILL_PATH, '京东交易流水*.csv'))
    
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
    # data_wx = read_data_wx(wx_files[0])
    # data_zfb = read_data_zfb(zfb_files[0])
    # data_jd = read_data_jd(jd_files[0])
    
    # # 合并数据
    # init_data_merge = pd.concat([data_wx, data_zfb, data_jd], axis=0)
    
    # # 处理数据
    # handled_data_merge, drop_1, drop_2 = handle_total_data(init_data_merge)
    
    # # 导出到Excel
    # save_to_excel(handled_data_merge, drop_1, drop_2)
    # print("\n运行成功！write successfully!")

if __name__ == "__main__":
    main()