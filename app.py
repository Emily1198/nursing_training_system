import os
import psycopg2
import psycopg2.extras
from flask import Flask, render_template_string, request, redirect, url_for, flash, session
from functools import wraps

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "nursing_secret_key_secure_2026")

# 1. 資料庫連線函數
def get_db_connection():
    db_url = os.getenv("DATABASE_URL")
    if not db_url:
        raise ValueError("未設定 DATABASE_URL 環境變數，請確認 Render 的 Environment Variables。")
    conn = psycopg2.connect(db_url)
    return conn

# 2. 自動檢查與修復資料表結構
def ensure_schema():
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        
        # 建立使用者表
        c.execute('''CREATE TABLE IF NOT EXISTS users (
                        id SERIAL PRIMARY KEY,
                        username VARCHAR(100) UNIQUE NOT NULL,
                        password VARCHAR(100) NOT NULL
                    )''')
        c.execute("INSERT INTO users (username, password) VALUES ('admin', 'admin123') ON CONFLICT (username) DO NOTHING")
        
        # 建立員工表
        c.execute('''CREATE TABLE IF NOT EXISTS employees (
                        id SERIAL PRIMARY KEY,
                        emp_no VARCHAR(50) UNIQUE,
                        name VARCHAR(100) NOT NULL,
                        position VARCHAR(100) NOT NULL,
                        status VARCHAR(20) DEFAULT '在職'
                    )''')
        
        # 建立課程表
        c.execute('''CREATE TABLE IF NOT EXISTS courses (
                        id SERIAL PRIMARY KEY,
                        course_date VARCHAR(20),
                        title VARCHAR(200),
                        category VARCHAR(100),
                        hours REAL,
                        source VARCHAR(100),
                        method VARCHAR(100),
                        has_nursing_points INTEGER DEFAULT 0,
                        has_ltc_points INTEGER DEFAULT 0
                    )''')
                    
        # 檢查並自動補充缺少欄位
        c.execute("""
            SELECT column_name 
            FROM information_schema.columns 
            WHERE table_name = 'courses'
        """)
        existing_cols = [col[0] for col in c.fetchall()]
        if 'has_nursing_points' not in existing_cols:
            c.execute("ALTER TABLE courses ADD COLUMN has_nursing_points INTEGER DEFAULT 0")
        if 'has_ltc_points' not in existing_cols:
            c.execute("ALTER TABLE courses ADD COLUMN has_ltc_points INTEGER DEFAULT 0")

        # 建立完訓紀錄表
        c.execute('''CREATE TABLE IF NOT EXISTS training_records (
                        id SERIAL PRIMARY KEY,
                        course_id INTEGER,
                        employee_id INTEGER,
                        completion_date VARCHAR(20),
                        UNIQUE(course_id, employee_id)
                    )''')

        # 建立下拉選單選項表
        c.execute('''CREATE TABLE IF NOT EXISTS dropdown_options (
                        id SERIAL PRIMARY KEY,
                        category_type VARCHAR(50) NOT NULL,
                        option_name VARCHAR(100) NOT NULL,
                        UNIQUE(category_type, option_name)
                    )''')
        
        default_options = [
            ('position', '院長/護理師'), ('position', '主任/護理師'), ('position', '護理師'),
            ('position', '照服員'), ('position', '社工師'), ('position', '營養師'), ('position', '其他'),
            ('course_category', '傳染病與群聚事件'), ('course_category', '感染管制'),
            ('course_category', '消防安全'), ('course_category', '緊急應變'), ('course_category', '性別平等'), ('course_category', '其他'),
            ('source', '長照平台'), ('source', '內訓'), ('source', '外訓'),
            ('method', '線上'), ('method', '實體')
        ]
        
        for cat_type, opt_name in default_options:
            c.execute(
                "INSERT INTO dropdown_options (category_type, option_name) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                (cat_type, opt_name)
            )

        conn.commit()
        c.close()
    except Exception as e:
        print(f"❌ 結構自動修復失敗: {e}")
    finally:
        if conn:
            conn.close()

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            flash("🔒 請先登入系統！")
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

LOGIN_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>護理機構教育訓練管理系統 - 登入</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light d-flex align-items-center justify-content-center" style="min-height: 100vh;">
<div class="text-center" style="width: 100%; max-width: 400px;">
    <div class="mb-4">
        <h2 class="fw-bold text-primary mb-1">🏥 護理訓練管理系統</h2>
        <p class="text-muted small mb-0">教育訓練與完訓統計管理平台</p>
    </div>
    <div class="card shadow-lg text-start">
        <div class="card-header bg-dark text-white text-center py-3">
            <h5 class="mb-0 fw-bold">🔒 管理員登入</h5>
        </div>
        <div class="card-body p-4">
            {% with messages = get_flashed_messages() %}
              {% if messages %}
                {% for message in messages %}
                  <div class="alert alert-danger py-2 small">{{ message|safe }}</div>
                {% endfor %}
              {% endif %}
            {% endwith %}
            <form action="/login" method="post">
                <div class="mb-3">
                    <label class="form-label fw-bold">帳號</label>
                    <input type="text" name="username" class="form-control" placeholder="admin" required>
                </div>
                <div class="mb-3">
                    <label class="form-label fw-bold">密碼</label>
                    <input type="password" name="password" class="form-control" placeholder="admin123" required>
                </div>
                <button type="submit" class="btn btn-primary w-100 fw-bold py-2 mt-2">登入系統</button>
            </form>
        </div>
    </div>
</div>
</body>
</html>
'''

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>教育訓練統計管理系統</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light pb-5">
<nav class="navbar navbar-expand-lg navbar-dark bg-dark mb-4 shadow">
    <div class="container-fluid">
        <span class="navbar-brand mb-0 h1">🏥 教育訓練統計管理系統</span>
        <div class="d-flex align-items-center">
            <span class="text-light me-3 small">👤 使用者：<strong>{{ session['username'] }}</strong></span>
            <a href="/logout" class="btn btn-sm btn-outline-light">🚪 登出</a>
        </div>
    </div>
</nav>

<div class="container">
    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for message in messages %}
          <div class="alert alert-warning alert-dismissible fade show shadow-sm" role="alert">
            {{ message|safe }}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
          </div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <div class="row">
        <!-- 1. 新增員工 -->
        <div class="col-md-5 mb-4">
            <div class="card shadow-sm h-100">
                <div class="card-header bg-primary text-white font-weight-bold">👤 新增員工</div>
                <div class="card-body">
                    <form action="/add_employee" method="post">
                        <div class="mb-2"><input type="text" name="emp_no" class="form-control" placeholder="工號 (如: EMP001)" required></div>
                        <div class="mb-2"><input type="text" name="name" class="form-control" placeholder="姓名" required></div>
                        <div class="mb-2">
                            <select name="position" class="form-select" required>
                                <option value="">選擇職位...</option>
                                {% for opt in opts.position %}
                                <option value="{{ opt }}">{{ opt }}</option>
                                {% endfor %}
                            </select>
                        </div>
                        <button type="submit" class="btn btn-primary w-100 mt-2">儲存員工資料</button>
                    </form>
                </div>
            </div>
        </div>

        <!-- 2. 建立課程 -->
        <div class="col-md-7 mb-4">
            <div class="card shadow-sm h-100">
                <div class="card-header bg-success text-white font-weight-bold">📚 建立新課程</div>
                <div class="card-body">
                    <form action="/add_course" method="post">
                        <div class="row g-2 mb-2">
                            <div class="col-md-4"><label class="form-label mb-0 small fw-bold">課程日期</label><input type="date" name="course_date" class="form-control" required></div>
                            <div class="col-md-5"><label class="form-label mb-0 small fw-bold">課程名稱</label><input type="text" name="title" class="form-control" placeholder="如: 傳染病防治" required></div>
                            <div class="col-md-3">
                                <label class="form-label mb-0 small fw-bold">課程類別</label>
                                <select name="category" class="form-select" required>
                                    {% for opt in opts.course_category %}
                                    <option value="{{ opt }}">{{ opt }}</option>
                                    {% endfor %}
                                </select>
                            </div>
                        </div>
                        <div class="row g-2 mb-2">
                            <div class="col-md-4"><label class="form-label mb-0 small fw-bold">時數</label><input type="number" step="0.5" name="hours" class="form-control" placeholder="如 1.0" required></div>
                            <div class="col-md-4">
                                <label class="form-label mb-0 small fw-bold">訓練來源</label>
                                <select name="source" class="form-select" required>
                                    {% for opt in opts.source %}
                                    <option value="{{ opt }}">{{ opt }}</option>
                                    {% endfor %}
                                </select>
                            </div>
                            <div class="col-md-4">
                                <label class="form-label mb-0 small fw-bold">上課方式</label>
                                <select name="method" class="form-select" required>
                                    {% for opt in opts.method %}
                                    <option value="{{ opt }}">{{ opt }}</option>
                                    {% endfor %}
                                </select>
                            </div>
                        </div>
                        <div class="mb-3 bg-light p-2 rounded border">
                            <span class="small fw-bold me-3">🎖️ 積分屬性設定：</span>
                            <div class="form-check form-check-inline">
                                <input class="form-check-input" type="checkbox" name="has_nursing_points" id="chk_nursing" value="1">
                                <label class="form-check-input-label small fw-bold text-primary" for="chk_nursing">護理積分</label>
                            </div>
                            <div class="form-check form-check-inline">
                                <input class="form-check-input" type="checkbox" name="has_ltc_points" id="chk_ltc" value="1">
                                <label class="form-check-input-label small fw-bold text-success" for="chk_ltc">長照積分</label>
                            </div>
                        </div>
                        <button type="submit" class="btn btn-success w-100 fw-bold">建立課程</button>
                    </form>
                </div>
            </div>
        </div>
    </div>

    <!-- 3. 資料檢視區 -->
    <div class="row">
        <div class="col-md-5 mb-4">
            <div class="card shadow-sm h-100">
                <div class="card-header bg-secondary text-white fw-bold">👥 目前員工名冊 (共 {{ employees|length }} 人)</div>
                <div class="card-body p-0" style="max-height: 300px; overflow-y: auto;">
                    <table class="table table-striped table-hover mb-0 small align-middle">
                        <thead class="table-light sticky-top">
                            <tr><th>工號</th><th>姓名</th><th>職位</th><th>狀態</th><th>操作</th></tr>
                        </thead>
                        <tbody>
                            {% for e in employees %}
                            <tr>
                                <form action="/update_employee/{{ e[0] }}" method="post">
                                    <td>{{ e[1] }}</td>
                                    <td class="fw-bold">{{ e[2] }}</td>
                                    <td>{{ e[3] }}</td>
                                    <td>
                                        <select name="status" class="form-select form-select-sm p-1">
                                            <option value="在職" {% if e[4] == '在職' %}selected{% endif %}>在職</option>
                                            <option value="離職" {% if e[4] == '離職' %}selected{% endif %}>離職</option>
                                        </select>
                                    </td>
                                    <td><button type="submit" class="btn btn-outline-primary btn-sm py-0 px-1">更新</button></td>
                                </form>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="col-md-7 mb-4">
            <div class="card shadow-sm h-100">
                <div class="card-header bg-secondary text-white fw-bold">📖 目前課程清單 (共 {{ courses|length }} 門)</div>
                <div class="card-body p-0" style="max-height: 300px; overflow-y: auto;">
                    <table class="table table-striped table-hover mb-0 small align-middle">
                        <thead class="table-light sticky-top">
                            <tr><th>日期</th><th>課程名稱</th><th>類別</th><th>時數</th><th>積分屬性</th><th>完訓</th></tr>
                        </thead>
                        <tbody>
                            {% for c in courses %}
                            <tr>
                                <td>{{ c[1] }}</td>
                                <td class="fw-bold">{{ c[2] }}</td>
                                <td>{{ c[3] }}</td>
                                <td>{{ c[4] }}</td>
                                <td>
                                    {% if c[7] %}<span class="badge bg-primary">護理</span>{% endif %}
                                    {% if c[8] %}<span class="badge bg-success">長照</span>{% endif %}
                                    {% if not c[7] and not c[8] %}<span class="text-muted small">無</span>{% endif %}
                                </td>
                                <td><span class="badge bg-info text-dark">{{ c[9] }} 人</span></td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
'''

def get_dropdown_options():
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT category_type, option_name FROM dropdown_options")
        rows = c.fetchall()
        c.close()
        opts = {'position': [], 'course_category': [], 'source': [], 'method': []}
        for cat, name in rows:
            if cat in opts:
                opts[cat].append(name)
        return opts
    except Exception as e:
        print(f"取得選單失敗: {e}")
        return {'position': [], 'course_category': [], 'source': [], 'method': []}
    finally:
        if conn:
            conn.close()

@app.route('/login', methods=['GET', 'POST'])
def login():
    ensure_schema() # 登入頁面確保資料表已就緒
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password'].strip()
        conn = None
        try:
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("SELECT id, username FROM users WHERE username = %s AND password = %s", (username, password))
            user = c.fetchone()
            c.close()
            if user:
                session['logged_in'] = True
                session['user_id'] = user[0]
                session['username'] = user[1]
                return redirect(url_for('index'))
            else:
                flash("❌ 帳號或密碼錯誤！")
        except Exception as e:
            flash(f"❌ 資料庫連線失敗: {e}")
        finally:
            if conn:
                conn.close()
    return render_template_string(LOGIN_TEMPLATE)

@app.route('/logout')
def logout():
    session.clear()
    flash("👋 已成功登出系統。")
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    ensure_schema()  # 自動檢查修復資料表
    employees = []
    courses = []
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT id, emp_no, name, position, status FROM employees ORDER BY emp_no ASC")
        employees = c.fetchall()
        
        c.execute('''
            SELECT c.id, c.course_date, c.title, c.category, c.hours, c.source, c.method,
                   COALESCE(c.has_nursing_points, 0), COALESCE(c.has_ltc_points, 0),
                   COUNT(r.id) as attendee_count
            FROM courses c
            LEFT JOIN training_records r ON c.id = r.course_id
            GROUP BY c.id, c.course_date, c.title, c.category, c.hours, c.source, c.method, c.has_nursing_points, c.has_ltc_points
            ORDER BY c.course_date DESC
        ''')
        courses = c.fetchall()
        c.close()
    except Exception as e:
        flash(f"⚠️ 資料載入異常: {e}")
    finally:
        if conn:
            conn.close()
            
    opts = get_dropdown_options()
    return render_template_string(HTML_TEMPLATE, employees=employees, courses=courses, opts=opts)

@app.route('/add_employee', methods=['POST'])
@login_required
def add_employee():
    emp_no = request.form['emp_no'].strip()
    name = request.form['name'].strip()
    position = request.form['position']
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("INSERT INTO employees (emp_no, name, position) VALUES (%s, %s, %s)", (emp_no, name, position))
        conn.commit()
        c.close()
        flash(f"✅ 已新增員工：{name} ({position})")
    except Exception as e:
        flash("❌ 新增失敗：工號可能已存在！")
    finally:
        if conn:
            conn.close()
    return redirect(url_for('index'))

@app.route('/add_course', methods=['POST'])
@login_required
def add_course():
    course_date = request.form['course_date']
    title = request.form['title'].strip()
    category = request.form['category']
    hours = float(request.form['hours'])
    source = request.form['source']
    method = request.form['method']
    has_nursing_points = 1 if request.form.get('has_nursing_points') else 0
    has_ltc_points = 1 if request.form.get('has_ltc_points') else 0

    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute('''
            INSERT INTO courses (course_date, title, category, hours, source, method, has_nursing_points, has_ltc_points)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        ''', (course_date, title, category, hours, source, method, has_nursing_points, has_ltc_points))
        conn.commit()
        c.close()
        flash(f"✅ 已成功建立課程：{title}")
    except Exception as e:
        flash(f"❌ 建立課程失敗: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for('index'))

@app.route('/update_employee/<int:emp_id>', methods=['POST'])
@login_required
def update_employee(emp_id):
    status = request.form['status']
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("UPDATE employees SET status = %s WHERE id = %s", (status, emp_id))
        conn.commit()
        c.close()
        flash("✅ 員工狀態更新成功！")
    except Exception as e:
        flash(f"❌ 更新失敗: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
