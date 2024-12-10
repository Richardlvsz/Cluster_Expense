import pandas as pd
import numpy as np
import os

# 生成示例数据
np.random.seed(42)  # 设置随机种子以确保可重复性

# 生成50个员工ID
emp_ids = [f'EMP{str(i).zfill(3)}' for i in range(1, 51)]

# 创建数据列表
data = []
for year in [2023, 2024]:
    for month in range(1, 13):
        for emp_id in emp_ids:
            # 基本工资在8000-15000之间
            base_salary = np.random.uniform(8000, 15000)
            
            # 各项保险基于工资的百分比
            row = {
                'emp_id': emp_id,
                'year': year,
                'month': month,
                'SAL': round(base_salary, 2),
                'HF': round(base_salary * 0.12, 2),  # 住房公积金
                'PEN': round(base_salary * 0.08, 2),  # 养老保险
                'UEM': round(base_salary * 0.02, 2),  # 失业保险
                'MED1': round(base_salary * 0.02, 2),  # 医疗保险1
                'MED2': round(base_salary * 0.01, 2),  # 医疗保险2
                'INJ': round(base_salary * 0.005, 2)  # 工伤保险
            }
            data.append(row)

# 创建DataFrame
df = pd.DataFrame(data)

# 确保uploads文件夹存在
if not os.path.exists('uploads'):
    os.makedirs('uploads')

# 保存为Excel文件
output_file = 'uploads/sample_expense_data.xlsx'
df.to_excel(output_file, index=False)
print(f"示例数据已生成并保存到: {output_file}")
print(f"总共生成了 {len(data)} 条记录")
print("\n数据预览:")
print(df.head())
