import pandas as pd

# 读取CSV文件
csv_file = input("请输入CSV文件路径：")
df = pd.read_csv(csv_file)

# 考虑收/支类型处理金额
df['调整金额'] = df.apply(lambda x: x['金额'] if x['收/支'] == '收入' else -x['金额'] if x['收/支'] == '支出' else 0, axis=1)


# 创建Excel写入对象
with pd.ExcelWriter('bill.xlsx', engine='openpyxl') as writer:
    # 保存所有数据到total表，包含调整后的金额
    df.to_excel(writer, sheet_name='total', index=False)
    
    # 筛选金额最大的10笔交易（区分收支）
    top_10_income = df[df['收/支'] == '收入'].nlargest(10, '金额')
    top_10_expense = df[df['收/支'] == '支出'].nlargest(10, '金额')
    
    # 创建分析工作表
    ansly_sheet = writer.sheets['ansly'] = writer.book.create_sheet('ansly')
    
    # 写入收入金额最大的10笔交易
    ansly_sheet.append(['收入金额最大的10笔交易'])
    ansly_sheet.append(top_10_income.columns.tolist())
    for row in top_10_income.values.tolist():
        ansly_sheet.append(row)
    
    # 空行分隔
    ansly_sheet.append([])
    
    # 写入支出金额最大的10笔交易
    ansly_sheet.append(['支出金额最大的10笔交易'])
    ansly_sheet.append(top_10_expense.columns.tolist())
    for row in top_10_expense.values.tolist():
        ansly_sheet.append(row)
    
    # 计算总收支情况
    total_income = df[df['收/支'] == '收入']['金额'].sum()
    total_expense = df[df['收/支'] == '支出']['金额'].sum()
    net_expense = total_expense - total_income  # 净支出 = 总支出 - 总收入
    
    # 写入总收支统计
    ansly_sheet.append([])
    ansly_sheet.append(['总收支统计'])
    ansly_sheet.append(['总收入', total_income])
    ansly_sheet.append(['总支出', total_expense])
    ansly_sheet.append(['净支出', net_expense])
    
    # 计算各交易类型收支情况
    type_stats = df.groupby(['交易类型', '收/支'])['金额'].sum().unstack(fill_value=0)
    
    # 确保包含所有收支类型列
    if '收入' not in type_stats.columns:
        type_stats['收入'] = 0
    if '支出' not in type_stats.columns:
        type_stats['支出'] = 0
    
    # 计算净支出
    type_stats['净支出'] = type_stats['支出'] - type_stats['收入']
    
    # 写入交易类型统计标题
    ansly_sheet.append([])
    ansly_sheet.append(['各交易类型收支统计'])
    ansly_sheet.append(['交易类型', '总收入', '总支出', '净支出'])
    
    # 写入交易类型统计数据
    type_data_start_row = ansly_sheet.max_row + 1
    for index, row in type_stats.iterrows():
        ansly_sheet.append([index, row['收入'], row['支出'], row['净支出']])
    
    # 创建柱状图 - 交易类型支出分布
    from openpyxl.chart import BarChart, Reference
    
    chart = BarChart()
    chart.type = 'col'
    chart.title = '各交易类型支出分布'
    chart.y_axis.title = '金额'
    chart.x_axis.title = '交易类型'
    
    # 设置数据范围 - 支出
    data = Reference(ansly_sheet, min_col=3, min_row=type_data_start_row, 
                     max_col=3, max_row=type_data_start_row + len(type_stats) - 1)
    categories = Reference(ansly_sheet, min_col=1, min_row=type_data_start_row, 
                           max_row=type_data_start_row + len(type_stats) - 1)
    
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(categories)
    
    # 设置图表样式
    chart.shape = 4
    ansly_sheet.add_chart(chart, f'F{type_data_start_row - 2}')  # 放置图表
    
    # 创建柱状图 - 交易类型收入分布
    chart_income = BarChart()
    chart_income.type = 'col'
    chart_income.title = '各交易类型收入分布'
    chart_income.y_axis.title = '金额'
    chart_income.x_axis.title = '交易类型'
    
    # 设置数据范围 - 收入
    data_income = Reference(ansly_sheet, min_col=2, min_row=type_data_start_row, 
                            max_col=2, max_row=type_data_start_row + len(type_stats) - 1)
    
    chart_income.add_data(data_income, titles_from_data=False)
    chart_income.set_categories(categories)
    
    # 设置图表样式
    chart_income.shape = 4
    ansly_sheet.add_chart(chart_income, f'F{type_data_start_row + len(type_stats) + 1}')  # 放置图表    