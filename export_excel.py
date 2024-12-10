import sqlite3
import pandas as pd

# 连接到 SQLite 数据库
db_path = "cluster_expense.db"  # 请替换为实际的数据库路径
conn = sqlite3.connect(db_path)

# 从数据库中读取 expense 表
query = "SELECT * FROM expense"
df = pd.read_sql_query(query, conn)

# 关闭数据库连接
conn.close()

# 指定目标 Excel 文件路径
target_file = r"D:\OneDrive - Marriott International\Attachments\20241121\Cluster_Expense\Export20241207.xlsx"

# 将数据写入 Excel 文件
df.to_excel(target_file, sheet_name='expense', index=False)

print("数据已成功从 SQLite 导出到 Export20241207.xlsx")