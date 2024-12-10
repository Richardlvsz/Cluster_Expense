import pandas as pd
import numpy as np

# 读取Excel文件中的Sheet1和SAL
file_path = '20241125.xlsx'
sheet1_df = pd.read_excel(file_path, sheet_name='Sheet1')
sal_df = pd.read_excel(file_path, sheet_name='Sal')

# 将空值填充为0
sal_df = sal_df.fillna(0)
sheet1_df = sheet1_df.fillna(0)

# 创建一个字典来存储每个员工每个月的工资
salary_dict = {}
for _, row in sal_df.iterrows():
    emp_id = row['emp_id']
    # 遍历除emp_id外的所有列（月份）
    for month in sal_df.columns:
        if month != 'emp_id':
            salary = row[month]
            # 只在工资不为0时添加到字典
            if salary != 0:
                if emp_id not in salary_dict:
                    salary_dict[emp_id] = {}
                salary_dict[emp_id][int(month)] = salary

# 遍历Sheet1，根据emp_id和month列填充工资
for index, row in sheet1_df.iterrows():
    emp_id = row['emp_id']
    month = int(row['month'])
    if emp_id in salary_dict and month in salary_dict[emp_id]:
        sheet1_df.at[index, 'salary'] = salary_dict[emp_id][month]

# 检查未匹配的emp_id和月份，并添加新行
new_rows = []
for emp_id, months in salary_dict.items():
    for month, salary in months.items():
        # 只在工资不为0时添加新行
        if salary != 0 and not ((sheet1_df['emp_id'] == emp_id) & (sheet1_df['month'] == month)).any():
            # 创建新行
            new_row = {'emp_id': emp_id, 'month': month, 'salary': salary}
            # 将其他字段设为0
            for col in sheet1_df.columns:
                if col not in new_row:
                    new_row[col] = 0
            new_rows.append(new_row)

# 使用 pd.concat 添加新行到 DataFrame
if new_rows:
    new_rows_df = pd.DataFrame(new_rows)
    sheet1_df = pd.concat([sheet1_df, new_rows_df], ignore_index=True)

# 删除工资为0的行
sheet1_df = sheet1_df[sheet1_df['salary'] != 0]

# 将合并后的内容写回到Excel文件中的Sheet1
with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
    sheet1_df.to_excel(writer, sheet_name='Sheet1', index=False)



