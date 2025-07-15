import os

# 配置文件
# 列定义
"""
账单列定义
来源：支付宝、微信、京东
交易时间：
交易类型：交通出行、餐饮美食、投资理财、生活开支、日用百货、文化休闲、鞋帽服饰、餐饮零食、消费还款、其他
类型细化：预留后续精细化
交易对方：商家名称
商品：商品名称或说明
交易状态：交易成功、交易失败、待定
交易方式：花呗、信用卡、现金、其他
收/支：收入、支出、不计收支、待定
金额：

支付宝映射：
    0：新增来源支付宝
    1：0
    2：1
    3：待定
    4：2
    5：4
    6：8
    7：7
    8：5
    9：6
微信映射：
    0：新增来源微信
    1：0
    2：1
    3：待定
    4：2
    5：3
    6：7
    7：6
    8：4
    9：5
"""
BILL_COLUMNS = [
    '来源','交易时间', '交易类型', '类型细化', '交易对方', 
    '商品', '交易状态', '交易方式','收/支', '金额'
]


# 映射表
EXPENSE_MAPPING = {
    '收入': '收入',  
    '支出': '支出',  
    '不计收支': '不计收支', 
    '待定': '待定'
}

TRANSACTION_STATUS_MAPPING = {
    '交易成功': '交易成功',  
    '交易失败': '交易失败',
    '待定': '待定'  
}

# 默认值
DEFAULT_EXPENSE = '待定'
DEFAULT_TRANSACTION_STATUS = '待定'

# 文件路径配置
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
INIT_BILL_PATH = os.path.join(ROOT_DIR, "初始账单", "25年01月份")
OUTPUT_DIR = os.path.join(ROOT_DIR, "账单", "25年01月份整理")
SAVE_PATH = os.path.join(OUTPUT_DIR, "25_01_total_bill.xlsx")

# 格式定义
FORMAT_STR = '%Y-%m-%d %H:%M:%S'