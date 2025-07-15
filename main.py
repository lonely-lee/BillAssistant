import os
import glob
from src.DataReader import DataReader
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
    data_reader = DataReader()
    data_reader.read_data(wx_files)
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