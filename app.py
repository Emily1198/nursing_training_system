import os
from flask import Flask, request, redirect, url_for, session, render_template_string, send_file, flash
import psycopg2
from psycopg2.extras import RealDictCursor
import pandas as pd
import io

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'nursing_secret_key_12345')

# 取得 Supabase 資料庫連線字串
DATABASE_URL = os.environ.get('DATABASE_URL')

def get_db_connection():
    if not DATABASE_URL:
        raise ValueError("未設定 DATABASE_URL 環境變數，請檢查 Render 環境變數設定。")
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    return conn

# 初始化資料庫表單结构
def init_db():
    conn = get_db_connection()
    cur = conn.cursor()
    
    # 建立員工表
    cur.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            id SERIAL PRIMARY KEY,
            emp_id VARCHAR(50) UNIQUE NOT NULL,
            name VARCHAR(100) NOT NULL,
            position VARCHAR(50) NOT NULL
        );
    ''')
    
    # 建立課程表
    cur.execute('''
        CREATE TABLE IF NOT EXISTS courses (
            id SERIAL PRIMARY KEY,
            course_name VARCHAR(150) NOT NULL,
            course_type VARCHAR(50) NOT NULL,
            hours NUMERIC(4,1) NOT NULL,
            course_date DATE NOT NULL,
            is_nursing_credit BOOLEAN DEFAULT FALSE,
            is_longterm_credit BOOLEAN DEFAULT FALSE
        );
    ''')
    
    # 建立完訓紀錄表
    cur.execute('''
        CREATE TABLE IF NOT EXISTS records (
            id SERIAL PRIMARY KEY,
            emp_id VARCHAR(50) NOT NULL,
            course_id INT NOT NULL,
            completion_date DATE NOT NULL,
            UNIQUE(emp_id, course_id)
        );
    ''')
    
    # 建立動態選單表 (職稱、課程類型)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS categories (
            id SERIAL PRIMARY KEY,
            cat_type VARCHAR(20) NOT NULL,
            cat_value VARCHAR(50) NOT NULL,
            UNIQUE(cat_type, cat_value)
        );
    ''')
    
    conn.commit()
    cur.close()
    conn.close()

# HTML 樣式範本 (含系統標題)
HTML_HEADER = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>家園教育訓練統計管理系統</title>
    <style>
        body { font-family: "Microsoft JhengHei", Arial, sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; }
        .container { max-width: 1000px; margin: auto; background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1, h2 { color: #2c3e50; text-align: center; }
        .nav { display: flex; justify-content: center; gap: 15px; margin-bottom: 25px; background: #eef2f5; padding: 10px; border-radius: 5px; }
        .nav a { text-decoration: none; color: #34495e; font-weight: bold; padding: 8px 16px; border-radius: 4px; }
        .nav a:hover { background-color: #3498db; color: white; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: center; }
        th { background-color: #3498db; color: white; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input, select { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }
        .btn { background-color: #2ecc71; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        .btn:hover { background-color: #27ae60; }
        .btn-export { background-color: #e67e22; }
        .btn-export:hover { background-color: #d35400; }
        .login-box { max-width: 400px; margin: 80px auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.15); }
    </style>
</head>
<body>
<div class="container">
"""

HTML_FOOTER = """
</div>
</body>
</html>
"""

# 登入頁面
LOGIN_TEMPLATE = HTML_HEADER + """
<div class="login-box">
    <!-- 系統標題 -->
    <h2>家園教育訓練<br>統計管理系統</h2>
    <p style="text-align: center; color: #7f8c8d; font-size: 0.9em; margin-bottom: 25px;">Educational Training Statistical Management System</p>
    
    {% if error %}
    <p style="color: red; text-align: center;">{{ error }}</p>
    {% endif %}
    
    <form method="POST">
        <div class="form-group">
            <label>管理員密碼：</label>
            <input type="password" name="password" placeholder="請輸入密碼" required>
        </div>
        <button type="submit" class="btn" style="width: 100%;">登入系統</button>
    </form>
</div>
""" + HTML_FOOTER

# 主選單與導覽頁面
NAV_BAR = """
<h1>家園教育訓練統計管理系統</h1>
<div class="nav">
    <a href="{{ url_for('dashboard') }}">儀表板</a>
    <a href="{{ url_for('employees') }}">員工管理</a>
    <a href="{{ url_for('courses') }}">課程管理</a>
    <a href="{{ url_for('records') }}">完訓登錄</a>
    <a href="{{ url_for('reports') }}">統計報表匯出</a>
    <a href="{{ url_for('logout') }}" style="color: #e74c3c;">登出</a>
</div>
"""

@app.route('/', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        pwd = request.form.get('password')
        if pwd == os.environ.get('ADMIN_PASSWORD', 'admin123'): # 預設密碼 admin123
            session['logged_in'] = True
            return redirect(url_for('dashboard'))
        else:
            error = "密碼錯誤，請重新輸入。"
    return render_template_string(LOGIN_TEMPLATE, error=error)

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/dashboard')
def dashboard():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    try:
        init_db()
    except Exception as e:
        return f"資料庫連接初始化失敗：{str(e)}"
        
    return render_template_string(HTML_HEADER + NAV_BAR + """
    <h2>歡迎使用統計管理系統</h2>
    <p style="text-align:center;">請點選上方導覽列開始進行員工、課程管理或完訓統計報表匯出。</p>
    """ + HTML_FOOTER)

# 員工管理
@app.route('/employees', methods=['GET', 'POST'])
def employees():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    if request.method == 'POST':
        emp_id = request.form.get('emp_id')
        name = request.form.get('name')
        position = request.form.get('position')
        
        cur.execute("INSERT INTO employees (emp_id, name, position) VALUES (%s, %s, %s) ON CONFLICT (emp_id) DO NOTHING",
                    (emp_id, name, position))
        conn.commit()
        return redirect(url_for('employees'))
        
    cur.execute("SELECT * FROM employees ORDER BY emp_id")
    emp_list = cur.fetchall()
    cur.close()
    conn.close()
    
    return render_template_string(HTML_HEADER + NAV_BAR + """
    <h2>員工資料管理</h2>
    <form method="POST" style="margin-bottom: 20px;">
        <div style="display: flex; gap: 10px;">
            <input type="text" name="emp_id" placeholder="員工編號" required>
            <input type="text" name="name" placeholder="員工姓名" required>
            <input type="text" name="position" placeholder="職稱 (如：護理師、照服員)" required>
            <button type="submit" class="btn">新增員工</button>
        </div>
    </form>
    <table>
        <tr><th>員工編號</th><th>姓名</th><th>職稱</th></tr>
        {% for emp in emp_list %}
        <tr><td>{{ emp.emp_id }}</td><td>{{ emp.name }}</td><td>{{ emp.position }}</td></tr>
        {% endfor %}
    </table>
    """ + HTML_FOOTER, emp_list=emp_list)

# 課程管理
@app.route('/courses', methods=['GET', 'POST'])
def courses():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    if request.method == 'POST':
        c_name = request.form.get('course_name')
        c_type = request.form.get('course_type')
        hours = request.form.get('hours')
        c_date = request.form.get('course_date')
        is_nursing = True if request.form.get('is_nursing') else False
        is_longterm = True if request.form.get('is_longterm') else False
        
        cur.execute("""
            INSERT INTO courses (course_name, course_type, hours, course_date, is_nursing_credit, is_longterm_credit)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (c_name, c_type, hours, c_date, is_nursing, is_longterm))
        conn.commit()
        return redirect(url_for('courses'))
        
    cur.execute("SELECT * FROM courses ORDER BY course_date DESC")
    course_list = cur.fetchall()
    cur.close()
    conn.close()
    
    return render_template_string(HTML_HEADER + NAV_BAR + """
    <h2>教育訓練課程管理</h2>
    <form method="POST" style="margin-bottom: 20px;">
        <div class="form-group"><input type="text" name="course_name" placeholder="課程名稱" required></div>
        <div class="form-group"><input type="text" name="course_type" placeholder="課程類別 (如：消防安全、感染管制)" required></div>
        <div class="form-group"><input type="number" step="0.5" name="hours" placeholder="時數" required></div>
        <div class="form-group"><input type="date" name="course_date" required></div>
        <div class="form-group">
            <label><input type="checkbox" name="is_nursing" value="1"> 採計護理人員積分</label>
            <label><input type="checkbox" name="is_longterm" value="1"> 採計長照人員積分</label>
        </div>
        <button type="submit" class="btn">新增課程</button>
    </form>
    <table>
        <tr><th>日期</th><th>課程名稱</th><th>類別</th><th>時數</th><th>護理積分</th><th>長照積分</th></tr>
        {% for c in course_list %}
        <tr>
            <td>{{ c.course_date }}</td>
            <td>{{ c.course_name }}</td>
            <td>{{ c.course_type }}</td>
            <td>{{ c.hours }}</td>
            <td>{{ '✓' if c.is_nursing_credit else '-' }}</td>
            <td>{{ '✓' if c.is_longterm_credit else '-' }}</td>
        </tr>
        {% endfor %}
    </table>
    """ + HTML_FOOTER, course_list=course_list)

# 完訓登錄
@app.route('/records', methods=['GET', 'POST'])
def records():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    conn = get_db_connection()
    cur = conn.cursor()
    
    if request.method == 'POST':
        emp_id = request.form.get('emp_id')
        course_id = request.form.get('course_id')
        comp_date = request.form.get('completion_date')
        
        cur.execute("""
            INSERT INTO records (emp_id, course_id, completion_date)
            VALUES (%s, %s, %s) ON CONFLICT (emp_id, course_id) DO NOTHING
        """, (emp_id, course_id, comp_date))
        conn.commit()
        return redirect(url_for('records'))
        
    cur.execute("SELECT * FROM employees ORDER BY name")
    emps = cur.fetchall()
    cur.execute("SELECT * FROM courses ORDER BY course_date DESC")
    courses = cur.fetchall()
    
    cur.execute("""
        SELECT r.id, e.emp_id, e.name, c.course_name, r.completion_date
        FROM records r
        JOIN employees e ON r.emp_id = e.emp_id
        JOIN courses c ON r.course_id = c.id
        ORDER BY r.completion_date DESC
    """)
    records_list = cur.fetchall()
    
    cur.close()
    conn.close()
    
    return render_template_string(HTML_HEADER + NAV_BAR + """
    <h2>完訓紀錄登錄</h2>
    <form method="POST" style="margin-bottom: 20px;">
        <div class="form-group">
            <select name="emp_id" required>
                <option value="">-- 選擇員工 --</option>
                {% for e in emps %}
                <option value="{{ e.emp_id }}">{{ e.emp_id }} - {{ e.name }} ({{ e.position }})</option>
                {% endfor %}
            </select>
        </div>
        <div class="form-group">
            <select name="course_id" required>
                <option value="">-- 選擇課程 --</option>
                {% for c in courses %}
                <option value="{{ c.id }}">{{ c.course_date }} - {{ c.course_name }} ({{ c.hours }}小時)</option>
                {% endfor %}
            </select>
        </div>
        <div class="form-group">
            <input type="date" name="completion_date" required>
        </div>
        <button type="submit" class="btn">新增完訓紀錄</button>
    </form>
    <table>
        <tr><th>完訓日期</th><th>工號</th><th>姓名</th><th>完訓課程</th></tr>
        {% for r in records_list %}
        <tr>
            <td>{{ r.completion_date }}</td>
            <td>{{ r.emp_id }}</td>
            <td>{{ r.name }}</td>
            <td>{{ r.course_name }}</td>
        </tr>
        {% endfor %}
    </table>
    """ + HTML_FOOTER, emps=emps, courses=courses, records_list=records_list)

# 統計報表與 Excel 匯出
@app.route('/reports')
def reports():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    return render_template_string(HTML_HEADER + NAV_BAR + """
    <h2>統計報表匯出</h2>
    <div style="display: flex; gap: 20px; justify-content: center; margin-top: 30px;">
        <a href="{{ url_for('export_matrix') }}"><button class="btn btn-export">匯出「評鑑專用交叉矩陣表」(.xlsx)</button></a>
        <a href="{{ url_for('export_summary') }}"><button class="btn btn-export">匯出「多頁簽統計報表」(.xlsx)</button></a>
    </div>
    """ + HTML_FOOTER)

# 匯出評鑑矩陣 Excel
@app.route('/export/matrix')
def export_matrix():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    conn = get_db_connection()
    query = """
        SELECT e.emp_id AS "員工編號", e.name AS "姓名", e.position AS "職稱",
               c.course_name AS "課程名稱", c.hours AS "時數"
        FROM records r
        JOIN employees e ON r.emp_id = e.emp_id
        JOIN courses c ON r.course_id = c.id
    """
    df = pd.read_sql(query, conn)
    conn.close()
    
    if df.empty:
        pivot_df = pd.DataFrame(columns=['員工編號', '姓名', '職稱'])
    else:
        pivot_df = df.pivot_table(index=['員工編號', '姓名', '職稱'], 
                                  columns='課程名稱', 
                                  values='時數', 
                                  aggfunc='sum', 
                                  fill_value=0).reset_index()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pivot_df.to_excel(writer, sheet_name='評鑑交叉矩陣表', index=False)
        
    output.seek(0)
    return send_file(output, download_name="Evaluation_Matrix_Report.xlsx", as_attachment=True)

# 匯出多頁簽統計 Excel
@app.route('/export/summary')
def export_summary():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    conn = get_db_connection()
    emp_df = pd.read_sql("SELECT emp_id AS 員工編號, name AS 姓名, position AS 職稱 FROM employees", conn)
    course_df = pd.read_sql("SELECT course_name AS 課程名稱, course_type AS 類別, hours AS 時數, course_date AS 日期 FROM courses", conn)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        emp_df.to_excel(writer, sheet_name='員工名冊', index=False)
        course_df.to_excel(writer, sheet_name='課程清單', index=False)
        
    conn.close()
    output.seek(0)
    return send_file(output, download_name="Statistical_Summary_Report.xlsx", as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
