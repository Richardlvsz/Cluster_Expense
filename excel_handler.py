import pandas as pd
import sqlite3
from datetime import datetime
from config import SOCIAL_INSURANCE_CONFIG
import logging

class ExcelHandler:
    def __init__(self, db_path='employee.db'):
        self.db_path = db_path
        # 设置日志记录
        logging.basicConfig(
            filename='excel_handler.log',
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)

    def get_db_connection(self):
        """创建数据库连接"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def process_excel(self, file_path, month):
        """处理Excel文件并导入数据库"""
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            print("Excel open suceassed")
            # 验证必需的列是否存在
            required_columns = ['姓名', '工资']
            if not all(col in df.columns for col in required_columns):
                raise ValueError("Excel文件必须包含 '姓名' 和 '工资' 列")

            conn = self.get_db_connection()
            cursor = conn.cursor()

            # 处理每一行数据
            for _, row in df.iterrows():
                # 1. 插入或更新员工信息
                cursor.execute('''
                    INSERT OR REPLACE INTO employees (name, salary, department)
                    VALUES (?, ?, ?)
                ''', (row['姓名'], row['工资'], row.get('部门', None)))
                
                employee_id = cursor.lastrowid

                # 2. 计算并插入各项保险记录
                for insurance_type, config in SOCIAL_INSURANCE_CONFIG.items():
                    if insurance_type == 'SALARY':  # 跳过工资配置项
                        continue
                        
                    # 计算缴费基数和金额
                    base_amount = min(max(row['工资'], config['min_base']), config['max_base'])
                    amount = base_amount * config['rate']

                    # 插入保险记录
                    cursor.execute('''
                        INSERT INTO insurance_records 
                        (employee_id, month, insurance_type, base_amount, amount)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (employee_id, month, config['code'], base_amount, amount))

            conn.commit()
            return True, "数据导入成功"

        except Exception as e:
            if 'conn' in locals():
                conn.rollback()
            return False, f"导入失败: {str(e)}"
            
        finally:
            if 'conn' in locals():
                conn.close()

    def get_monthly_summary(self, month):
        """获取指定月份的汇总数据"""
        conn = self.get_db_connection()
        try:
            cursor = conn.cursor()
            
            # 获取该月份的所有记录
            query = '''
                SELECT 
                    e.name,
                    e.salary,
                    e.department,
                    ir.insurance_type,
                    ir.base_amount,
                    ir.amount
                FROM employees e
                JOIN insurance_records ir ON e.id = ir.employee_id
                WHERE ir.month = ?
                ORDER BY e.name, ir.insurance_type
            '''
            
            cursor.execute(query, (month,))
            records = cursor.fetchall()
            
            # 将结果转换为更易于使用的格式
            summary = {}
            for record in records:
                if record['name'] not in summary:
                    summary[record['name']] = {
                        'salary': record['salary'],
                        'department': record['department'],
                        'insurances': {}
                    }
                summary[record['name']]['insurances'][record['insurance_type']] = {
                    'base': record['base_amount'],
                    'amount': record['amount']
                }
                
            return summary
            
        finally:
            conn.close() 

    def import_expenses(self, file_path):
        """导入社保公积金费用Excel文件到employee_expenses表"""
        self.logger.info(f"开始导入文件: {file_path}")
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            self.logger.debug(f"读取到的Excel数据: {df.head()}")

            # 验证必需的列是否存在
            required_columns = ['emp_id', 'year', 'month']
            expense_columns = ['HF', 'PEN', 'UEM', 'MED1', 'MED2', 'INJ', 'UF']
            
            # 验证所有必需列
            missing_columns = [col for col in required_columns + expense_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Excel文件缺少以下列: {', '.join(missing_columns)}")

            conn = self.get_db_connection()
            cursor = conn.cursor()

            # 获取当前时间和文件名
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            file_name = file_path.split('/')[-1]

            # 处理每一行数据
            for _, row in df.iterrows():
                # 验证必需字段的值
                if pd.isna(row['emp_id']) or pd.isna(row['year']) or pd.isna(row['month']):
                    self.logger.warning(f"跳过无效行: {row}")
                    continue  # 跳过无效行

                # 确保 emp_id 为5位字符
                emp_id = str(row['emp_id']).zfill(5)

                # 对每个费用类型创建记录
                for expense_type in expense_columns:
                    amount = row[expense_type]
                    if pd.notna(amount):  # 只插入非空值
                        cursor.execute('''
                            INSERT INTO employee_expenses 
                            (emp_id, year, month, trans_type, amount, create_time, remarks)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            emp_id,               # 保留为5位字符
                            int(row['year']),     # 确保year为整数
                            int(row['month']),    # 确保month为整数
                            expense_type,
                            float(amount),        # 确保amount为浮点数
                            current_time,
                            file_name
                        ))
                        self.logger.debug(f"插入记录: emp_id={emp_id}, year={row['year']}, month={row['month']}, type={expense_type}, amount={amount}")

            conn.commit()
            self.logger.info("费用数据导入成功")
            return True, "费用数据导入成功"

        except Exception as e:
            self.logger.error(f"导入失败: {str(e)}", exc_info=True)
            if 'conn' in locals():
                conn.rollback()
            return False, f"费用导入失败: {str(e)}"
            
        finally:
            if 'conn' in locals():
                conn.close()
                self.logger.info("数据库连接已关闭")

    def import_cluster_expense(self, file_path, db_path='Cluster_Expense.db'):
        """导入Excel数据到Cluster_Expense.db的expense表
        
        字段说明：
        - emp_id: 员工ID（5位字符）
        - SAL: 工资
        - HF (Housing Fund): 住房公积金
        - PEN (Pension): 养老保险
        - UEM (Unemployment): 失业保险
        - MED1, MED2 (Medical Insurance): 医疗保险
        - INJ (Injury Insurance): 工伤保险
        - UF (Unit Fund): 工会费
        """
        self.logger.info(f"开始导入文件到Cluster_Expense: {file_path}")
        try:
            # 读取Excel文件的Sheet1
            df = pd.read_excel(file_path, sheet_name='Sheet1')
            self.logger.debug(f"读取到的Excel数据: {df.head()}")

            # 确保emp_id为5位字符
            df['emp_id'] = df['emp_id'].astype(str).str.zfill(5)

            # 连接到Cluster_Expense数据库
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            # 创建expense表（如果不存在）
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS expense (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    emp_id TEXT NOT NULL,           -- 员工ID（5位字符）
                    year INTEGER NOT NULL,          -- 年份
                    month INTEGER NOT NULL,         -- 月份
                    SAL REAL NOT NULL,             -- 工资
                    HF REAL NOT NULL,              -- Housing Fund（住房公积金）
                    PEN REAL NOT NULL,             -- Pension（养老保险）
                    UEM REAL NOT NULL,             -- Unemployment（失业保险）
                    MED1 REAL NOT NULL,            -- Medical Insurance 1（医疗保险1）
                    MED2 REAL NOT NULL,            -- Medical Insurance 2（医疗保险2）
                    INJ REAL NOT NULL,             -- Injury Insurance（工伤保险）
                    UF REAL NOT NULL,              -- Unit Fund（工会费）
                    create_time TEXT NOT NULL,      -- 记录创建时间
                    UNIQUE(emp_id, year, month)     -- 确保同一员工同一月份不会重复记录
                )
            ''')

            # 获取当前时间
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # 将DataFrame数据插入到expense表
            for _, row in df.iterrows():
                cursor.execute('''
                    INSERT OR REPLACE INTO expense 
                    (emp_id, year, month, SAL, HF, PEN, UEM, MED1, MED2, INJ, UF, create_time)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    row['emp_id'],
                    int(row['year']),
                    int(row['month']),
                    float(row['SAL']),
                    float(row['HF']),
                    float(row['PEN']),
                    float(row['UEM']),
                    float(row['MED1']),
                    float(row['MED2']),
                    float(row['INJ']),
                    float(row['UF']),
                    current_time
                ))

            conn.commit()
            self.logger.info("数据导入Cluster_Expense.db成功")
            return True, "数据导入成功"

        except Exception as e:
            self.logger.error(f"导入失败: {str(e)}", exc_info=True)
            if 'conn' in locals():
                conn.rollback()
            return False, f"导入失败: {str(e)}"
            
        finally:
            if 'conn' in locals():
                conn.close()
                self.logger.info("数据库连接已关闭")