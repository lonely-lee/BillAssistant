import os
import glob
from src.DataReader import DataReader
from src.DataProcessor import DataProcessor
from src.config import INIT_BILL_PATH, OUTPUT_DIR, SAVE_PATH
import sys
import argparse
import shutil

def check_and_clear_directory(directory):
    """
    检查指定目录是否存在：
    - 若不存在，创建目录
    - 若存在，清空目录下所有文件和子目录
    """
    # 检查目录是否存在
    if not os.path.exists(directory):
        # 创建目录（包括多级目录）
        os.makedirs(directory, exist_ok=True)
        print(f"目录不存在，已创建：{directory}")
    else:
        # 遍历目录下所有内容并删除
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            try:
                # 若为文件或链接，直接删除
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.unlink(item_path)
                    print(f"删除文件：{item_path}")
                # 若为子目录，递归删除
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                    print(f"删除子目录及内容：{item_path}")
            except Exception as e:
                print(f"删除 {item_path} 失败：{e}")
        print(f"目录已清空：{directory}")
def main():
    # 创建数据读取器
    data_reader = DataReader()
    # 创建数据处理器
    data_processor = DataProcessor()
    # 创建空数据
    df_total = []

    # 创建参数解析器
    parser = argparse.ArgumentParser(description='账单数据处理脚本')
    # 添加命令行选项
    parser.add_argument('-r', action='store_true', 
                        help='从配置的默认路径读取初始账单CSV文件，并作预处理')
    parser.add_argument('-d', action='store_true', 
                        help='从配置的默认路径读取预处理的CSV文件，写入Excel并生成图表')
    parser.add_argument('-t', action='store_true', 
                        help='使用配置中的训练数据训练模型并保存')
    parser.add_argument('-n', action='store_true', 
                        help='将人工审核的账单数据添加到训练数据集')
    # 解析命令行参数
    args = parser.parse_args()
    # 根据参数执行相应操作
    if args.r:
        print("读取初始账单并做预处理...")
        # 检查并清空输出目录
        check_and_clear_directory(OUTPUT_DIR)
        data_reader.read_dir_data()

    if args.d:
        print("将处理后的账单数据生成Excel和图表...")
        data_processor.process_data()
    
    if args.t:
        print("训练模型并保存...")
        data_reader.tringger_ml_training()
    
    if args.n:
        print("添加人工审核数据到训练集...")
        data_reader.add_new_train_data()
    # # 读取数据
    # data_dir = r'E:\生活\账单\scripts\data'
    # data_reader = DataReader()

    # # 如果执行脚本输入参数，例如 reprocess 即：python main.py --reprocess，则重新载入修改后的账单数据，重新生成账单图片（主要修改带人工确认选项）
    # df_total = data_reader.read_dir_data(data_dir)

if __name__ == "__main__":
    main()