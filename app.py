from flask import Flask, render_template, request, flash, redirect, jsonify, url_for
import pandas as pd
from werkzeug.utils import secure_filename
import os
import sqlite3
from config import SOCIAL_INSURANCE_CONFIG  # 导入配置
from excel_handler import ExcelHandler
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # 用于flash消息
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DATABASE'] = 'Cluster_Expense.db'  # 更新数据库路径

# 确保上传文件夹存在
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def get_db_connection():
    """创建数据库连接"""
    conn = sqlite3.connect(app.config['DATABASE'])
    conn.row_factory = sqlite3.Row  # 这样可以通过列名访问数据
    return conn

def init_db():
    """初始化数据库表"""
    conn = get_db_connection()
    c = conn.cursor()
    
    # 创建expense表
    c.execute('''
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
            UF REAL NOT NULL              -- Union Fee（工会经费）
        )
    ''')
    
    conn.commit()
    conn.close()

# 在应用启动时初始化数据库
init_db()

def get_monthly_summary():
    """获取按月份汇总的费用数据"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        query = """
        SELECT 
            year,
            month,
            ROUND(SUM(SAL), 2) as total_salary,
            ROUND(SUM(PEN), 2) as total_pension,
            ROUND(SUM(MED1 + MED2), 2) as total_medical,
            ROUND(SUM(INJ), 2) as total_injury,
            ROUND(SUM(UEM), 2) as total_unemployment,
            ROUND(SUM(HF), 2) as total_hf,
            ROUND(SUM(UF), 2) as total_union_fee,
            ROUND(SUM(HF + PEN + UEM + MED1 + MED2 + INJ), 2) as total_insurance,
            ROUND(SUM(SAL + HF + PEN + UEM + MED1 + MED2 + INJ + UF), 2) as grand_total
        FROM expense
        GROUP BY year, month
        ORDER BY year DESC, month DESC;
        """
        
        cursor.execute(query)
        results = cursor.fetchall()
        
        # Convert results to list of dictionaries for easier handling
        formatted_results = []
        for row in results:
            formatted_results.append({
                'year': row[0],
                'month': row[1],
                'total_salary': float(row[2]) if row[2] is not None else 0.0,
                'total_pension': float(row[3]) if row[3] is not None else 0.0,
                'total_medical': float(row[4]) if row[4] is not None else 0.0,
                'total_injury': float(row[5]) if row[5] is not None else 0.0,
                'total_unemployment': float(row[6]) if row[6] is not None else 0.0,
                'total_hf': float(row[7]) if row[7] is not None else 0.0,
                'total_union_fee': float(row[8]) if row[8] is not None else 0.0,
                'total_insurance': float(row[9]) if row[9] is not None else 0.0,
                'grand_total': float(row[10]) if row[10] is not None else 0.0
            })
            
        return formatted_results
    except Exception as e:
        print(f"Error in get_monthly_summary: {str(e)}")
        return []
    finally:
        conn.close()

def check_db_content():
    """Check if there's any data in the expense table"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM expense")
    count = cursor.fetchone()[0]
    cursor.execute("SELECT * FROM expense LIMIT 1")
    sample = cursor.fetchone()
    conn.close()
    return count, sample

def get_prev_month(year, month):
    if month == 1:
        return year - 1, 12
    return year, month - 1

@app.route('/')
def index():
    print("Rendering index.html")  # Debugging statement
    
    try:
        # Check database content
        row_count, sample_row = check_db_content()
        print(f"Database has {row_count} rows")
        if sample_row:
            print("Sample row:", dict(sample_row))
        
        # Get monthly summary data from database
        monthly_data = get_monthly_summary()
        print("Monthly data:", monthly_data)  # Debug print
        
        # Initialize lists for labels and data
        trend_labels = []
        total_amount_data = []
        total_salary_data = []
        total_insurance_data = []
        
        # Process the monthly data
        for row in monthly_data:
            print("Processing row:", row)  # Debug print
            # Create label in format "YYYY-MM"
            label = f"{row['year']}-{row['month']:02d}"
            trend_labels.append(label)
            
            # Extract values from row 
            total_salary = row['total_salary']
            total_insurance = row['total_insurance']
            
            # Calculate total amount (salary + insurance)
            total_amount = total_salary + total_insurance
            
            total_amount_data.append(total_amount)
            total_salary_data.append(total_salary)
            total_insurance_data.append(total_insurance)
        
        print("Labels:", trend_labels)  # Debug print
        print("Total amount data:", total_amount_data)  # Debug print
        print("Total salary data:", total_salary_data)  # Debug print
        print("Total insurance data:", total_insurance_data)  # Debug print
        
        # Structure the data as expected by the template
        trend_values = {
            'total_amount': total_amount_data,
            'total_salary': total_salary_data,
            'total_insurance': total_insurance_data
        }
        
        return render_template('index.html', 
                             monthly_data=monthly_data,
                             trend_labels=trend_labels, 
                             trend_values=trend_values)
    
    except Exception as e:
        print(f"Error in index route: {str(e)}")  # Debug print
        # Return empty data in case of error
        return render_template('index.html', 
                             monthly_data=[],
                             trend_labels=[], 
                             trend_values={'total_amount': [], 
                                         'total_salary': [], 
                                         'total_insurance': []})

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('没有选择文件')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('没有选择文件')
            return redirect(request.url)
        
        if file and file.filename.endswith('.xlsx'):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            try:
                # 读取Excel文件
                df = pd.read_excel(filepath)
                
                # 获取数据库连接
                conn = get_db_connection()
                cursor = conn.cursor()
                
                # 清空现有数据
                cursor.execute("DELETE FROM expense")
                
                # 处理每一行数据
                for index, row in df.iterrows():
                    # 假设Excel表格的列名与数据库字段对应
                    cursor.execute("""
                        INSERT INTO expense (
                            emp_id, year, month, 
                            SAL, HF, PEN, UEM, 
                            MED1, MED2, INJ, UF
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        str(row['员工ID']),
                        int(row['年份']),
                        int(row['月份']),
                        float(row['工资总额']),
                        float(row['住房公积金']),
                        float(row['养老保险']),
                        float(row['失业保险']),
                        float(row['医疗保险1']),
                        float(row['医疗保险2']),
                        float(row['工伤保险']),
                        float(row['工会经费'])
                    ))
                
                conn.commit()
                conn.close()
                
                flash('文件上传成功并已处理数据')
            except Exception as e:
                flash(f'处理文件时出错: {str(e)}')
            finally:
                # 删除上传的文件
                if os.path.exists(filepath):
                    os.remove(filepath)
            
            return redirect(url_for('index'))
        else:
            flash('只允许上传.xlsx格式的文件')
            return redirect(request.url)
    
    # 获取月度汇总数据
    monthly_data = get_monthly_summary()
    
    # 获取趋势数据
    trend_data = get_trend_data()
    
    return render_template('index.html', 
                         monthly_data=monthly_data,
                         trend_labels=trend_data['labels'],
                         trend_values=trend_data['values'])

@app.route('/monthly_detail/<int:year>/<int:month>')
def monthly_detail_page(year, month):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Get previous month
        prev_year, prev_month = get_prev_month(year, month)
        
        # Get employee details for both current and previous month
        query = """
        WITH current_month AS (
            SELECT 
                emp_id,
                SAL as salary,
                PEN as pension,
                MED1 + MED2 as medical,
                INJ as injury,
                UEM as unemployment,
                HF as housing_fund,
                UF as union_fee,
                SAL + PEN + MED1 + MED2 + INJ + UEM + HF + UF as total
            FROM expense
            WHERE year = ? AND month = ?
        ),
        prev_month AS (
            SELECT 
                emp_id,
                SAL as salary,
                PEN as pension,
                MED1 + MED2 as medical,
                INJ as injury,
                UEM as unemployment,
                HF as housing_fund,
                UF as union_fee,
                SAL + PEN + MED1 + MED2 + INJ + UEM + HF + UF as total
            FROM expense
            WHERE year = ? AND month = ?
        )
        SELECT 
            c.emp_id,
            c.salary as curr_salary,
            p.salary as prev_salary,
            c.pension as curr_pension,
            p.pension as prev_pension,
            c.medical as curr_medical,
            p.medical as prev_medical,
            c.injury as curr_injury,
            p.injury as prev_injury,
            c.unemployment as curr_unemployment,
            p.unemployment as prev_unemployment,
            c.housing_fund as curr_housing_fund,
            p.housing_fund as prev_housing_fund,
            c.union_fee as curr_union_fee,
            p.union_fee as prev_union_fee,
            c.total as curr_total,
            p.total as prev_total
        FROM current_month c
        LEFT JOIN prev_month p ON c.emp_id = p.emp_id
        ORDER BY c.emp_id
        """
        
        cursor.execute(query, (year, month, prev_year, prev_month))
        results = cursor.fetchall()
        
        # Convert results to list of dictionaries with comparison data
        employee_data = []
        for row in results:
            emp_data = {
                'emp_id': row[0],
                'salary': {
                    'current': float(row[1]) if row[1] is not None else 0.0,
                    'previous': float(row[2]) if row[2] is not None else 0.0
                },
                'pension': {
                    'current': float(row[3]) if row[3] is not None else 0.0,
                    'previous': float(row[4]) if row[4] is not None else 0.0
                },
                'medical': {
                    'current': float(row[5]) if row[5] is not None else 0.0,
                    'previous': float(row[6]) if row[6] is not None else 0.0
                },
                'injury': {
                    'current': float(row[7]) if row[7] is not None else 0.0,
                    'previous': float(row[8]) if row[8] is not None else 0.0
                },
                'unemployment': {
                    'current': float(row[9]) if row[9] is not None else 0.0,
                    'previous': float(row[10]) if row[10] is not None else 0.0
                },
                'housing_fund': {
                    'current': float(row[11]) if row[11] is not None else 0.0,
                    'previous': float(row[12]) if row[12] is not None else 0.0
                },
                'union_fee': {
                    'current': float(row[13]) if row[13] is not None else 0.0,
                    'previous': float(row[14]) if row[14] is not None else 0.0
                },
                'total': {
                    'current': float(row[15]) if row[15] is not None else 0.0,
                    'previous': float(row[16]) if row[16] is not None else 0.0
                }
            }
            
            # Calculate changes
            for key in ['salary', 'pension', 'medical', 'injury', 'unemployment', 'housing_fund', 'union_fee', 'total']:
                current = emp_data[key]['current']
                previous = emp_data[key]['previous']
                emp_data[key]['change'] = current - previous
                emp_data[key]['change_rate'] = (current - previous) / previous * 100 if previous != 0 else 0
            
            employee_data.append(emp_data)
        
        # Calculate totals
        totals = {
            'current': {},
            'previous': {},
            'change': {},
            'change_rate': {}
        }
        
        for key in ['salary', 'pension', 'medical', 'injury', 'unemployment', 'housing_fund', 'union_fee', 'total']:
            curr_sum = sum(emp[key]['current'] for emp in employee_data)
            prev_sum = sum(emp[key]['previous'] for emp in employee_data)
            totals['current'][key] = curr_sum
            totals['previous'][key] = prev_sum
            totals['change'][key] = curr_sum - prev_sum
            totals['change_rate'][key] = (curr_sum - prev_sum) / prev_sum * 100 if prev_sum != 0 else 0
        
        return render_template('monthly_detail.html',
                             year=year,
                             month=month,
                             prev_year=prev_year,
                             prev_month=prev_month,
                             employee_data=employee_data,
                             totals=totals)
    
    except Exception as e:
        print(f"Error in monthly_detail_page: {str(e)}")
        return render_template('monthly_detail.html',
                             year=year,
                             month=month,
                             prev_year=prev_year,
                             prev_month=prev_month,
                             employee_data=[],
                             totals={})
    finally:
        conn.close()

@app.route('/api/employee_comparison/<int:year>/<int:month>/<string:expense_type>')
def get_employee_comparison_by_type(year, month, expense_type):
    """获取特定费用类型的员工对比数据"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 计算上个月的年份和月份
    prev_month = month - 1
    prev_year = year
    if prev_month == 0:
        prev_month = 12
        prev_year = year - 1
    
    # 映射费用类型到数据库字段
    field_map = {
        'salary': 'SAL',
        'housing_fund': 'HF',
        'pension': 'PEN',
        'unemployment': 'UEM',
        'medical': 'MED1 + MED2',
        'injury': 'INJ',
        'total': 'SAL + HF + PEN + UEM + MED1 + MED2 + INJ'
    }
    
    if expense_type not in field_map:
        return {'error': '无效的费用类型'}, 400
    
    field = field_map[expense_type]
    
    # 分别查询当前月和上月的数据
    current_query = f"""
    SELECT 
        emp_id,
        {field} as amount
    FROM expense
    WHERE year = ? AND month = ?
    """
    
    prev_query = f"""
    SELECT 
        emp_id,
        {field} as amount
    FROM expense
    WHERE year = ? AND month = ?
    """
    
    current_data = {row['emp_id']: float(row['amount']) for row in cursor.execute(current_query, (year, month)).fetchall()}
    prev_data = {row['emp_id']: float(row['amount']) for row in cursor.execute(prev_query, (prev_year, prev_month)).fetchall()}
    
    # 合并所有员工ID
    all_emp_ids = set(list(current_data.keys()) + list(prev_data.keys()))
    
    # 计算每个员工的费用变化
    comparison_data = []
    for emp_id in sorted(all_emp_ids):
        curr = current_data.get(emp_id, 0.0)
        prev = prev_data.get(emp_id, 0.0)
        change = curr - prev
        change_rate = ((curr - prev) / prev * 100) if prev != 0 else 100 if curr > 0 else 0
        
        comparison_data.append({
            'emp_id': emp_id,
            'current': curr,
            'previous': prev,
            'change': change,
            'change_rate': change_rate
        })
    
    conn.close()
    return jsonify({'data': comparison_data})

def monthly_detail(year, month):
    """月度详情页面"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 获取指定月份的费用汇总数据
    query = """
    SELECT 
        year,
        month,
        SUM(SAL) as total_salary,
        SUM(PEN) as pension,
        SUM(MED1 + MED2) as medical,
        SUM(INJ) as injury,
        SUM(UEM) as unemployment,
        SUM(HF) as housing_fund,
        SUM(UF) as union_fee,
        COUNT(*) * 2 as union_fee  -- 假设工会经费为每人2元
    FROM expense
    WHERE year = ? AND month = ?
    GROUP BY year, month
    """
    
    cursor.execute(query, (year, month))
    summary = cursor.fetchone()
    conn.close()
    
    if summary:
        data = {
            'year_month': f"{year}-{month:02d}",
            'total_salary': summary['total_salary'],
            'pension': summary['pension'],
            'medical': summary['medical'],
            'injury': summary['injury'],
            'unemployment': summary['unemployment'],
            'housing_fund': summary['housing_fund'],
            'union_fee': summary['union_fee']
        }
    else:
        data = {
            'year_month': f"{year}-{month:02d}",
            'total_salary': 0,
            'pension': 0,
            'medical': 0,
            'injury': 0,
            'unemployment': 0,
            'housing_fund': 0,
            'union_fee': 0
        }
    
    return render_template('monthly_detail.html',
                         year=year,
                         month=month,
                         data=data)

@app.route('/preview_excel', methods=['POST'])
def preview_excel():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    
    if file and file.filename.endswith('.xlsx'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            df = pd.read_excel(filepath)
            # Convert emp_id to 5-digit string with leading zeros
            df['emp_id'] = df['emp_id'].astype(str).str.zfill(5)
            # Convert DataFrame to list of dictionaries for template rendering
            records = df.to_dict('records')
            return render_template('preview_excel.html', records=records, columns=df.columns.tolist())
        except Exception as e:
            flash(f'Error reading Excel file: {str(e)}')
            return redirect(url_for('index'))
    else:
        flash('Please upload an Excel file')
        return redirect(url_for('index'))

@app.route('/import_selected', methods=['POST'])
def import_selected():
    try:
        selected_records = request.json.get('selected_records', [])
        conn = get_db_connection()
        cursor = conn.cursor()
        
        for record in selected_records:
            cursor.execute('''
                INSERT INTO expense (emp_id, year, month, SAL, HF, PEN, UEM, MED1, MED2, INJ, UF)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                record['emp_id'],
                record['year'],
                record['month'],
                record['SAL'],
                record['HF'],
                record['PEN'],
                record['UEM'],
                record['MED1'],
                record['MED2'],
                record['INJ'],
                record['UF']
            ))
        
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '选中的记录已成功导入'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'导入失败: {str(e)}'})

if __name__ == '__main__':
    app.run(debug=True) 