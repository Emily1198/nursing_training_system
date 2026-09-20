import os
import io
import psycopg2
import psycopg2.extras
import pandas as pd
from datetime import datetime
from functools import wraps
from flask import Flask, render_template_string, request, redirect, url_for, flash, session, send_file

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "nursing_secret_key_secure_2026")

# 1. 資料庫連線
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
        
        c.execute('''CREATE TABLE IF NOT EXISTS users (
                        id SERIAL PRIMARY KEY,
                        username VARCHAR(100) UNIQUE NOT NULL,
                        password VARCHAR(100) NOT NULL
                    )''')
        c.execute("INSERT INTO users (username, password) VALUES ('admin', 'admin123') ON CONFLICT (username) DO NOTHING")
        
        c.execute('''CREATE TABLE IF NOT EXISTS employees (
                        id SERIAL PRIMARY KEY,
                        emp_no VARCHAR(50) UNIQUE,
                        name VARCHAR(100) NOT NULL,
                        position VARCHAR(100) NOT NULL,
                        status VARCHAR(20) DEFAULT '在職',
                        arrival_date VARCHAR(20)
                    )''')
        
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

        c.execute('''CREATE TABLE IF NOT EXISTS training_records (
                        id SERIAL PRIMARY KEY,
                        course_id INTEGER REFERENCES courses(id) ON DELETE CASCADE,
                        employee_id INTEGER REFERENCES employees(id) ON DELETE CASCADE,
                        completion_date VARCHAR(20),
                        UNIQUE(course_id, employee_id)
                    )''')

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
            c.execute("INSERT INTO dropdown_options (category_type, option_name) VALUES (%s, %s) ON CONFLICT DO NOTHING", (cat_type, opt_name))

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

def get_dropdown_options():
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT category_type, option_name FROM dropdown_options ORDER BY id ASC")
        rows = c.fetchall()
        c.close()
        opts = {'position': [], 'course_category': [], 'source': [], 'method': []}
        for cat, name in rows:
            if cat in opts:
                opts[cat].append(name)
        return opts
    except Exception as e:
        return {'position': [], 'course_category': [], 'source': [], 'method': []}
    finally:
        if conn:
            conn.close()

# ================= HTML 範本 =================
LOGIN_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>家園教育訓練統計管理系統 - 登入</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light d-flex align-items-center justify-content-center" style="min-height: 100vh;">
<div class="text-center" style="width: 100%; max-width: 400px;">
    <h2 class="fw-bold text-primary mb-3">🏥 家園教育訓練統計管理系統</h2>
    <div class="card shadow">
        <div class="card-header bg-dark text-white fw-bold">🔒 管理員登入</div>
        <div class="card-body p-4">
            {% with messages = get_flashed_messages() %}
              {% if messages %}
                {% for message in messages %}
                  <div class="alert alert-danger py-2 small">{{ message|safe }}</div>
                {% endfor %}
              {% endif %}
            {% endwith %}
            <form action="/login" method="post">
                <div class="mb-3 text-start"><label class="form-label fw-bold">帳號</label><input type="text" name="username" class="form-control" required></div>
                <div class="mb-3 text-start"><label class="form-label fw-bold">密碼</label><input type="password" name="password" class="form-control" required></div>
                <button type="submit" class="btn btn-primary w-100 fw-bold">登入系統</button>
            </form>
        </div>
    </div>
</div>
</body>
</html>
'''

MAIN_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>家園教育訓練統計管理系統</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light pb-5">
<nav class="navbar navbar-expand-lg navbar-dark bg-dark mb-3 shadow">
    <div class="container-fluid">
        <span class="navbar-brand mb-0 h1">🏥 家園教育訓練統計管理系統</span>
        <div class="d-flex align-items-center">
            <span class="text-light me-3 small">👤 使用者：<strong>{{ session['username'] }}</strong></span>
            <button class="btn btn-sm btn-outline-info me-2" data-bs-toggle="modal" data-bs-target="#accountModal">⚙️ 帳號密碼</button>
            <button class="btn btn-sm btn-outline-warning me-2" data-bs-toggle="modal" data-bs-target="#optionsModal">⚙️ 選項管理</button>
            <a href="/logout" class="btn btn-sm btn-outline-light">🚪 登出</a>
        </div>
    </div>
</nav>

<div class="container-fluid px-4">
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

    <div class="row g-3 mb-4">
        <!-- 1. 新增員工 -->
        <div class="col-md-3">
            <div class="card shadow-sm h-100">
                <div class="card-header bg-primary text-white fw-bold">👤 1. 新增員工</div>
                <div class="card-body">
                    <form action="/add_employee" method="post">
                        <div class="mb-2"><input type="text" name="emp_no" class="form-control form-control-sm" placeholder="工號 (如: EMP001)" required></div>
                        <div class="mb-2"><input type="text" name="name" class="form-control form-control-sm" placeholder="姓名" required></div>
                        <div class="mb-2">
                            <select name="position" class="form-select form-select-sm" required>
                                <option value="">選擇職位...</option>
                                {% for opt in opts.position %}<option value="{{ opt }}">{{ opt }}</option>{% endfor %}
                            </select>
                        </div>
                        <div class="mb-2"><label class="form-label mb-0 small">到職日</label><input type="date" name="arrival_date" class="form-control form-control-sm"></div>
                        <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">儲存員工資料</button>
                    </form>
                </div>
            </div>
        </div>

        <!-- 2. 建立新課程 -->
        <div class="col-md-5">
            <div class="card shadow-sm h-100">
                <div class="card-header bg-success text-white fw-bold">📚 2. 建立新課程</div>
                <div class="card-body">
                    <form action="/add_course" method="post" id="courseForm">
                        <input type="hidden" name="force_save" id="force_save" value="0">
                        <div class="row g-2 mb-2">
                            <div class="col-md-4"><label class="form-label mb-0 small fw-bold">課程預設日期</label><input type="date" name="course_date" class="form-control form-control-sm" required></div>
                            <div class="col-md-5"><label class="form-label mb-0 small fw-bold">課程名稱</label><input type="text" name="title" class="form-control form-control-sm" placeholder="課程名稱" required></div>
                            <div class="col-md-3">
                                <label class="form-label mb-0 small fw-bold">課程類別</label>
                                <select name="category" class="form-select form-select-sm" required>
                                    {% for opt in opts.course_category %}<option value="{{ opt }}">{{ opt }}</option>{% endfor %}
                                </select>
                            </div>
                        </div>
                        <div class="row g-2 mb-2">
                            <div class="col-md-4"><label class="form-label mb-0 small fw-bold">時數</label><input type="number" step="0.5" name="hours" class="form-control form-control-sm" placeholder="1.0" required></div>
                            <div class="col-md-4">
                                <label class="form-label mb-0 small fw-bold">訓練來源</label>
                                <select name="source" class="form-select form-select-sm" required>
                                    {% for opt in opts.source %}<option value="{{ opt }}">{{ opt }}</option>{% endfor %}
                                </select>
                            </div>
                            <div class="col-md-4">
                                <label class="form-label mb-0 small fw-bold">上課方式</label>
                                <select name="method" class="form-select form-select-sm" required>
                                    {% for opt in opts.method %}<option value="{{ opt }}">{{ opt }}</option>{% endfor %}
                                </select>
                            </div>
                        </div>
                        <div class="mb-2 bg-light p-2 rounded border">
                            <span class="small fw-bold me-2">🎖️ 積分：</span>
                            <div class="form-check form-check-inline"><input class="form-check-input" type="checkbox" name="has_nursing_points" id="cn" value="1"><label class="form-check-label small text-primary fw-bold" for="cn">護理積分</label></div>
                            <div class="form-check form-check-inline"><input class="form-check-input" type="checkbox" name="has_ltc_points" id="cl" value="1"><label class="form-check-label small text-success fw-bold" for="cl">長照積分</label></div>
                        </div>
                        <button type="submit" class="btn btn-success btn-sm w-100 fw-bold">建立課程</button>
                    </form>
                </div>
            </div>
        </div>

        <!-- 3. 登記完訓紀錄 (手動/Excel) -->
        <div class="col-md-4">
            <div class="card shadow-sm h-100">
                <div class="card-header bg-warning text-dark fw-bold">📝 3. 登記完訓紀錄</div>
                <div class="card-body">
                    <ul class="nav nav-tabs nav-justified mb-2" id="recordTab" role="tablist">
                        <li class="nav-item"><button class="nav-item nav-link active py-1 small fw-bold" id="manual-tab" data-bs-toggle="tab" data-bs-target="#manual" type="button">✍️ 手動登記</button></li>
                        <li class="nav-item"><button class="nav-item nav-link py-1 small fw-bold" id="excel-tab" data-bs-toggle="tab" data-bs-target="#excel" type="button">📤 Excel 上傳</button></li>
                    </ul>
                    <div class="tab-content" id="recordTabContent">
                        <div class="tab-pane fade show active" id="manual">
                            <form action="/add_record_manual" method="post">
                                <div class="mb-2">
                                    <select name="course_id" class="form-select form-select-sm" required>
                                        <option value="">選擇課程...</option>
                                        {% for c in courses %}<option value="{{ c[0] }}">{{ c[1] }} - {{ c[2] }}</option>{% endfor %}
                                    </select>
                                </div>
                                <div class="mb-2">
                                    <select name="employee_id" class="form-select form-select-sm" required>
                                        <option value="">選擇員工...</option>
                                        {% for e in employees %}<option value="{{ e[0] }}">{{ e[1] }} {{ e[2] }} ({{ e[3] }})</option>{% endfor %}
                                    </select>
                                </div>
                                <div class="mb-2"><label class="form-label mb-0 small">完訓日期 (線上不同完訓日可在此輸入)</label><input type="date" name="completion_date" class="form-control form-control-sm"></div>
                                <button type="submit" class="btn btn-warning btn-sm w-100 fw-bold">登記完訓</button>
                            </form>
                        </div>
                        <div class="tab-pane fade" id="excel">
                            <form action="/add_record_excel" method="post" enctype="multipart/form-data">
                                <div class="mb-2">
                                    <select name="course_id" class="form-select form-select-sm" required>
                                        <option value="">選擇欲登記之課程...</option>
                                        {% for c in courses %}<option value="{{ c[0] }}">{{ c[1] }} - {{ c[2] }}</option>{% endfor %}
                                    </select>
                                </div>
                                <div class="mb-2">
                                    <input type="file" name="file" class="form-control form-control-sm" accept=".xlsx, .xls" required>
                                    <div class="form-text small">Excel 內需有「姓名」欄位即可</div>
                                </div>
                                <button type="submit" class="btn btn-warning btn-sm w-100 fw-bold">解析並匯入完訓</button>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- 報表匯出與篩選區 -->
    <div class="card shadow-sm mb-4 border-info">
        <div class="card-header bg-info text-white fw-bold">📊 匯出評鑑與統計報表 (支援條件篩選)</div>
        <div class="card-body py-2">
            <form action="/export_report" method="get" class="row g-2 align-items-center">
                <div class="col-md-2">
                    <select name="report_type" class="form-select form-select-sm fw-bold" required>
                        <option value="evaluation">🏆 評鑑專用交叉總表</option>
                        <option value="summary">📈 簡易統計報表 (三分頁)</option>
                    </select>
                </div>
                <div class="col-md-2"><input type="date" name="start_date" class="form-control form-control-sm" placeholder="開始日期"></div>
                <div class="col-md-2"><input type="date" name="end_date" class="form-control form-control-sm" placeholder="結束日期"></div>
                <div class="col-md-2">
                    <select name="position" class="form-select form-select-sm">
                        <option value="">所有職位...</option>
                        {% for opt in opts.position %}<option value="{{ opt }}">{{ opt }}</option>{% endfor %}
                    </select>
                </div>
                <div class="col-md-2">
                    <select name="category" class="form-select form-select-sm">
                        <option value="">所有課程類別...</option>
                        {% for opt in opts.course_category %}<option value="{{ opt }}">{{ opt }}</option>{% endfor %}
                    </select>
                </div>
                <div class="col-md-2"><button type="submit" class="btn btn-dark btn-sm w-100 fw-bold">📥 產生並匯出 Excel</button></div>
            </form>
        </div>
    </div>

    <!-- 資料檢視與修正區 -->
    <div class="row">
        <!-- 員工清單 -->
        <div class="col-md-5 mb-4">
            <div class="card shadow-sm">
                <div class="card-header bg-secondary text-white fw-bold">👥 員工名冊 (共 {{ employees|length }} 人)</div>
                <div class="card-body p-0" style="max-height: 400px; overflow-y: auto;">
                    <table class="table table-striped table-hover mb-0 align-middle small">
                        <thead class="table-light sticky-top">
                            <tr><th>工號</th><th>姓名</th><th>職位</th><th>到職日</th><th>狀態</th><th>操作</th></tr>
                        </thead>
                        <tbody>
                            {% for e in employees %}
                            <tr>
                                <form action="/update_employee/{{ e[0] }}" method="post" onsubmit="return confirm('確定要更新此員工資料？');">
                                    <td>{{ e[1] }}</td>
                                    <td><input type="text" name="name" value="{{ e[2] }}" class="form-control form-control-sm p-1" style="width:70px;" required></td>
                                    <td>
                                        <select name="position" class="form-select form-select-sm p-1">
                                            {% for opt in opts.position %}
                                            <option value="{{ opt }}" {% if e[3] == opt %}selected{% endif %}>{{ opt }}</option>
                                            {% endfor %}
                                        </select>
                                    </td>
                                    <td><input type="text" name="arrival_date" value="{{ e[5] or '' }}" class="form-control form-control-sm p-1" style="width:85px;"></td>
                                    <td>
                                        <select name="status" class="form-select form-select-sm p-1">
                                            <option value="在職" {% if e[4] == '在職' %}selected{% endif %}>在職</option>
                                            <option value="離職" {% if e[4] == '離職' %}selected{% endif %}>離職</option>
                                            <option value="留職停薪" {% if e[4] == '留職停薪' %}selected{% endif %}>留職停薪</option>
                                        </select>
                                    </td>
                                    <td>
                                        <button type="submit" class="btn btn-primary btn-sm py-0 px-1">儲存</button>
                                        <a href="/delete_employee/{{ e[0] }}" class="btn btn-danger btn-sm py-0 px-1" onclick="return confirm('⚠️ 警告：確定要刪除員工「{{ e[2] }}」？這會一併刪除其所有完訓紀錄！');">刪除</a>
                                    </td>
                                </form>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- 課程清單與完訓查詢 -->
        <div class="col-md-7 mb-4">
            <div class="card shadow-sm">
                <div class="card-header bg-secondary text-white fw-bold">📖 課程清冊 (共 {{ courses|length }} 門)</div>
                <div class="card-body p-0" style="max-height: 400px; overflow-y: auto;">
                    <table class="table table-striped table-hover mb-0 align-middle small">
                        <thead class="table-light sticky-top">
                            <tr><th>預設日期</th><th>課程名稱</th><th>類別</th><th>時數</th><th>完訓</th><th>操作</th></tr>
                        </thead>
                        <tbody>
                            {% for c in courses %}
                            <tr>
                                <form action="/update_course/{{ c[0] }}" method="post" onsubmit="return confirm('確定要修正此課程內容？');">
                                    <td><input type="date" name="course_date" value="{{ c[1] }}" class="form-control form-control-sm p-1" required></td>
                                    <td><input type="text" name="title" value="{{ c[2] }}" class="form-control form-control-sm p-1" required></td>
                                    <td>
                                        <select name="category" class="form-select form-select-sm p-1">
                                            {% for opt in opts.course_category %}
                                            <option value="{{ opt }}" {% if c[3] == opt %}selected{% endif %}>{{ opt }}</option>
                                            {% endfor %}
                                        </select>
                                    </td>
                                    <td><input type="number" step="0.5" name="hours" value="{{ c[4] }}" class="form-control form-control-sm p-1" style="width:55px;" required></td>
                                    <td>
                                        <a href="/course_attendees/{{ c[0] }}" class="badge bg-info text-dark text-decoration-none">完訓 {{ c[9] }} 人 🔍</a>
                                    </td>
                                    <td>
                                        <button type="submit" class="btn btn-primary btn-sm py-0 px-1">儲存</button>
                                        <a href="/delete_course/{{ c[0] }}" class="btn btn-danger btn-sm py-0 px-1" onclick="return confirm('⚠️ 警告：確定要刪除課程「{{ c[2] }}」？這會一併刪除所有人此課的完訓紀錄！');">刪除</a>
                                    </td>
                                </form>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Modal: 選項管理 -->
<div class="modal fade" id="optionsModal" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header bg-dark text-white"><h5 class="modal-title">⚙️ 下拉選單選項目維護</h5><button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">
        <form action="/add_dropdown_option" method="post" class="mb-3">
            <div class="input-group">
                <select name="category_type" class="form-select" required>
                    <option value="position">員工職位</option>
                    <option value="course_category">課程類別</option>
                    <option value="source">訓練來源</option>
                    <option value="method">上課方式</option>
                </select>
                <input type="text" name="option_name" class="form-control" placeholder="新增選項名稱" required>
                <button class="btn btn-success" type="submit">新增</button>
            </div>
        </form>
      </div>
    </div>
  </div>
</div>

<!-- Modal: 帳號密碼修改 -->
<div class="modal fade" id="accountModal" tabindex="-1">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header bg-dark text-white"><h5 class="modal-title">⚙️ 帳號與密碼管理</h5><button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button></div>
      <div class="modal-body">
        <form action="/update_account" method="post">
            <div class="mb-2"><label class="form-label small fw-bold">帳號</label><input type="text" name="username" value="{{ session['username'] }}" class="form-control" required></div>
            <div class="mb-3"><label class="form-label small fw-bold">新密碼</label><input type="password" name="password" class="form-control" placeholder="若不修改請留空"></div>
            <button class="btn btn-primary w-100" type="submit" onclick="return confirm('確定要更新帳號密碼？');">儲存修改</button>
        </form>
      </div>
    </div>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
'''

# 完訓人員清單與修正頁面
COURSE_ATTENDEES_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>課程完訓人員管理</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light p-4">
<div class="container">
    <div class="d-flex justify-content-between align-items-center mb-3">
        <h4>📖 課程完訓名單：<span class="text-primary">{{ course[2] }}</span> ({{ course[1] }})</h4>
        <a href="/" class="btn btn-secondary btn-sm">⬅️ 返回主頁面</a>
    </div>
    {% with messages = get_flashed_messages() %}
      {% if messages %}{% for message in messages %}<div class="alert alert-warning py-2 small">{{ message|safe }}</div>{% endfor %}{% endif %}
    {% endwith %}
    <div class="card shadow-sm">
        <div class="card-body p-0">
            <table class="table table-striped mb-0 align-middle">
                <thead class="table-dark">
                    <tr><th>工號</th><th>姓名</th><th>職位</th><th>完訓日期</th><th>操作</th></tr>
                </thead>
                <tbody>
                    {% for r in records %}
                    <tr>
                        <form action="/update_record/{{ r[0] }}" method="post" onsubmit="return confirm('確定要更新完訓日期？');">
                            <td>{{ r[1] }}</td>
                            <td class="fw-bold">{{ r[2] }}</td>
                            <td>{{ r[3] }}</td>
                            <td><input type="date" name="completion_date" value="{{ r[4] }}" class="form-control form-control-sm" style="width:160px;" required></td>
                            <td>
                                <button type="submit" class="btn btn-primary btn-sm">儲存</button>
                                <a href="/delete_record/{{ r[0] }}" class="btn btn-danger btn-sm" onclick="return confirm('確定要移除此完訓紀錄？');">移除完訓</a>
                            </td>
                        </form>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
</div>
</body>
</html>
'''

# ================= 路由處理邏輯 =================

@app.route('/login', methods=['GET', 'POST'])
def login():
    ensure_schema()
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
            if conn: conn.close()
    return render_template_string(LOGIN_TEMPLATE)

@app.route('/logout')
def logout():
    session.clear()
    flash("👋 已成功登出系統。")
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    ensure_schema()
    employees = []
    courses = []
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT id, emp_no, name, position, status, arrival_date FROM employees ORDER BY emp_no ASC")
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
        if conn: conn.close()
            
    opts = get_dropdown_options()
    return render_template_string(MAIN_TEMPLATE, employees=employees, courses=courses, opts=opts)

# 需求 1 & 8: 新增員工與刪除確認
@app.route('/add_employee', methods=['POST'])
@login_required
def add_employee():
    emp_no = request.form['emp_no'].strip()
    name = request.form['name'].strip()
    position = request.form['position']
    arrival_date = request.form.get('arrival_date', '').strip()
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("INSERT INTO employees (emp_no, name, position, arrival_date) VALUES (%s, %s, %s, %s)", (emp_no, name, position, arrival_date))
        conn.commit()
        c.close()
        flash(f"✅ 已新增員工：{name} ({position})")
    except Exception as e:
        flash("❌ 新增失敗：工號可能已存在！")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

@app.route('/update_employee/<int:emp_id>', methods=['POST'])
@login_required
def update_employee(emp_id):
    name = request.form['name'].strip()
    position = request.form['position']
    arrival_date = request.form['arrival_date'].strip()
    status = request.form['status']
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("UPDATE employees SET name=%s, position=%s, arrival_date=%s, status=%s WHERE id=%s", (name, position, arrival_date, status, emp_id))
        conn.commit()
        c.close()
        flash("✅ 員工資料更新成功！")
    except Exception as e:
        flash(f"❌ 更新失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

@app.route('/delete_employee/<int:emp_id>')
@login_required
def delete_employee(emp_id):
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM employees WHERE id = %s", (emp_id,))
        conn.commit()
        c.close()
        flash("🗑️ 已成功刪除員工及其完訓紀錄。")
    except Exception as e:
        flash(f"❌ 刪除失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

# 需求 6: 課程重複提示與建置
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
    force_save = request.form.get('force_save', '0')

    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        
        if force_save != '1':
            c.execute("SELECT id FROM courses WHERE title=%s AND course_date=%s AND hours=%s", (title, course_date, hours))
            dup = c.fetchone()
            if dup:
                flash(f"⚠️ 提示：發現一模一樣的課程「{title} ({course_date})」，若仍要重複新增，請重新點擊建立！")
                c.close()
                return redirect(url_for('index'))

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
        if conn: conn.close()
    return redirect(url_for('index'))

@app.route('/update_course/<int:course_id>', methods=['POST'])
@login_required
def update_course(course_id):
    course_date = request.form['course_date']
    title = request.form['title'].strip()
    category = request.form['category']
    hours = float(request.form['hours'])
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("UPDATE courses SET course_date=%s, title=%s, category=%s, hours=%s WHERE id=%s", (course_date, title, category, hours, course_id))
        conn.commit()
        c.close()
        flash("✅ 課程資料修訂成功！")
    except Exception as e:
        flash(f"❌ 修正失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

@app.route('/delete_course/<int:course_id>')
@login_required
def delete_course(course_id):
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM courses WHERE id = %s", (course_id,))
        conn.commit()
        c.close()
        flash("🗑️ 已成功刪除課程及其所有完訓紀錄。")
    except Exception as e:
        flash(f"❌ 刪除失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

# 需求 1, 3, 5: 完訓紀錄登記與缺失員工提醒
@app.route('/add_record_manual', methods=['POST'])
@login_required
def add_record_manual():
    course_id = request.form['course_id']
    employee_id = request.form['employee_id']
    completion_date = request.form.get('completion_date', '').strip()

    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        if not completion_date:
            c.execute("SELECT course_date FROM courses WHERE id=%s", (course_id,))
            completion_date = c.fetchone()[0]

        c.execute("INSERT INTO training_records (course_id, employee_id, completion_date) VALUES (%s, %s, %s) ON CONFLICT DO NOTHING", (course_id, employee_id, completion_date))
        conn.commit()
        c.close()
        flash("✅ 完訓紀錄登記成功！")
    except Exception as e:
        flash(f"❌ 登記失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

@app.route('/add_record_excel', methods=['POST'])
@login_required
def add_record_excel():
    course_id = request.form['course_id']
    file = request.files.get('file')
    if not file:
        flash("❌ 請選擇 Excel 檔案！")
        return redirect(url_for('index'))

    conn = None
    try:
        df = pd.read_excel(file)
        name_col = None
        for col in df.columns:
            if '姓名' in str(col):
                name_col = col; break
        if not name_col:
            flash("❌ Excel 內找不到「姓名」欄位！")
            return redirect(url_for('index'))

        names = df[name_col].dropna().astype(str).str.strip().tolist()

        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT course_date FROM courses WHERE id=%s", (course_id,))
        default_date = c.fetchone()[0]

        c.execute("SELECT id, name FROM employees")
        emp_map = {row[1]: row[0] for row in c.fetchall()}

        missing_names = []
        success_cnt = 0
        for name in names:
            if name in emp_map:
                emp_id = emp_map[name]
                c.execute("INSERT INTO training_records (course_id, employee_id, completion_date) VALUES (%s, %s, %s) ON CONFLICT DO NOTHING", (course_id, emp_id, default_date))
                success_cnt += 1
            else:
                missing_names.append(name)
        
        conn.commit()
        c.close()

        msg = f"✅ Excel 解析完成，成功登記 {success_cnt} 人！"
        if missing_names:
            msg += f"<br>⚠️ <strong>系統未找到以下員工，請先去新增員工：</strong> {', '.join(missing_names)}"
        flash(msg)
    except Exception as e:
        flash(f"❌ Excel 匯入失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

# 需求 12 & 13: 查看單獨課程完訓名單與修正
@app.route('/course_attendees/<int:course_id>')
@login_required
def course_attendees(course_id):
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT id, course_date, title FROM courses WHERE id=%s", (course_id,))
        course = c.fetchone()
        
        c.execute('''
            SELECT r.id, e.emp_no, e.name, e.position, r.completion_date
            FROM training_records r
            JOIN employees e ON r.employee_id = e.id
            WHERE r.course_id = %s
            ORDER BY e.emp_no ASC
        ''', (course_id,))
        records = c.fetchall()
        c.close()
        return render_template_string(COURSE_ATTENDEES_TEMPLATE, course=course, records=records)
    except Exception as e:
        flash(f"❌ 載入失敗: {e}")
        return redirect(url_for('index'))
    finally:
        if conn: conn.close()

@app.route('/update_record/<int:record_id>', methods=['POST'])
@login_required
def update_record(record_id):
    completion_date = request.form['completion_date']
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("UPDATE training_records SET completion_date=%s WHERE id=%s", (completion_date, record_id))
        conn.commit()
        c.close()
        flash("✅ 完訓日期修正成功！")
    except Exception as e:
        flash(f"❌ 修正失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(request.referrer or url_for('index'))

@app.route('/delete_record/<int:record_id>')
@login_required
def delete_record(record_id):
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM training_records WHERE id=%s", (record_id,))
        conn.commit()
        c.close()
        flash("🗑️ 已成功移除完訓紀錄。")
    except Exception as e:
        flash(f"❌ 移除失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(request.referrer or url_for('index'))

# 需求 10 & 14: 帳號密碼修改與下拉選單項目新增
@app.route('/update_account', methods=['POST'])
@login_required
def update_account():
    username = request.form['username'].strip()
    password = request.form['password'].strip()
    user_id = session['user_id']

    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        if password:
            c.execute("UPDATE users SET username=%s, password=%s WHERE id=%s", (username, password, user_id))
        else:
            c.execute("UPDATE users SET username=%s WHERE id=%s", (username, user_id))
        conn.commit()
        c.close()
        session['username'] = username
        flash("✅ 帳號密碼修改成功！")
    except Exception as e:
        flash(f"❌ 修改失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

@app.route('/add_dropdown_option', methods=['POST'])
@login_required
def add_dropdown_option():
    cat_type = request.form['category_type']
    opt_name = request.form['option_name'].strip()
    conn = None
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("INSERT INTO dropdown_options (category_type, option_name) VALUES (%s, %s) ON CONFLICT DO NOTHING", (cat_type, opt_name))
        conn.commit()
        c.close()
        flash(f"✅ 已成功新增選項：{opt_name}")
    except Exception as e:
        flash(f"❌ 新增失敗: {e}")
    finally:
        if conn: conn.close()
    return redirect(url_for('index'))

# 需求 7: 評鑑專用動態交叉總表與三分頁報表匯出
def to_roc_date(date_str):
    if not date_str: return "X"
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return f"{dt.year - 1911}.{dt.month:02d}.{dt.day:02d}"
    except:
        return date_str

@app.route('/export_report')
@login_required
def export_report():
    report_type = request.args.get('report_type', 'evaluation')
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    pos_filter = request.args.get('position', '').strip()
    cat_filter = request.args.get('category', '').strip()

    conn = get_db_connection()
    c = conn.cursor()

    emp_sql = "SELECT id, emp_no, name, position, arrival_date FROM employees WHERE 1=1"
    emp_params = []
    if pos_filter:
        emp_sql += " AND position = %s"
        emp_params.append(pos_filter)
    emp_sql += " ORDER BY emp_no ASC"
    c.execute(emp_sql, emp_params)
    employees = c.fetchall()

    crs_sql = "SELECT id, course_date, title, category, hours, source, method, has_nursing_points, has_ltc_points FROM courses WHERE 1=1"
    crs_params = []
    if start_date:
        crs_sql += " AND course_date >= %s"; crs_params.append(start_date)
    if end_date:
        crs_sql += " AND course_date <= %s"; crs_params.append(end_date)
    if cat_filter:
        crs_sql += " AND category = %s"; crs_params.append(cat_filter)
    crs_sql += " ORDER BY course_date ASC, id ASC"
    c.execute(crs_sql, crs_params)
    courses = c.fetchall()

    c.execute("SELECT course_id, employee_id, completion_date FROM training_records")
    rec_rows = c.fetchall()
    rec_map = {(r[0], r[1]): r[2] for r in rec_rows}
    c.close()
    conn.close()

    output = io.BytesIO()
    wb = openpyxl.Workbook()

    thin = Side(border_style="thin", color="000000")
    double = Side(border_style="double", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    font_bold = Font(name="標楷體", size=10, bold=True)
    font_normal = Font(name="標楷體", size=10)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    if report_type == 'evaluation':
        ws = wb.active
        ws.title = "評鑑專用總表"

        title_text = "教育課程清冊總表"
        if cat_filter: title_text = f"{cat_filter} 教育課程清冊總表"
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(courses) + 4)
        ws.cell(row=1, column=1, value=title_text).font = Font(name="標楷體", size=16, bold=True)
        ws.cell(row=1, column=1).alignment = align_center

        headers_r2 = ["序號", "職稱", "姓名", "課程名稱"]
        headers_r3 = ["", "", "", "方式/時數"]
        headers_r4 = ["", "", "", "到職日"]

        for crs in courses:
            headers_r2.append(crs[2])
            headers_r3.append(f"{crs[5]}/{crs[4]}Hr")
            headers_r4.append("上課日期")

        headers_r2.append("總時數")
        headers_r3.append("")
        headers_r4.append("")

        ws.append(headers_r2)
        ws.append(headers_r3)
        ws.append(headers_r4)

        ws.merge_cells("A2:A4"); ws.merge_cells("B2:B4"); ws.merge_cells("C2:C4")

        row_idx = 5
        for idx, emp in enumerate(employees, start=1):
            arr_date_roc = to_roc_date(emp[4])
            row_data = [idx, emp[3], emp[2], arr_date_roc]
            total_hrs = 0.0

            for crs in courses:
                key = (crs[0], emp[0])
                if key in rec_map:
                    comp_date = rec_map[key]
                    row_data.append(to_roc_date(comp_date))
                    total_hrs += float(crs[4])
                else:
                    row_data.append("X")

            row_data.append(total_hrs)
            ws.append(row_data)

            # 設定儲存格樣式
            for col_i in range(1, len(row_data) + 1):
                cell = ws.cell(row=row_idx, column=col_i)
                cell.font = font_normal
                cell.alignment = align_center
                cell.border = border_all
            row_idx += 1

        for r in range(2, 5):
            for c_i in range(1, len(headers_r2) + 1):
                cell = ws.cell(row=r, column=c_i)
                cell.font = font_bold
                cell.alignment = align_center
                cell.fill = fill_header
                cell.border = border_all

    else: # 簡易統計報表三分頁
        ws1 = wb.active; ws1.title = "人員完訓時數總表"
        ws1.append(["工號", "姓名", "職位", "到職日", "完訓總時數"])
        for emp in employees:
            t_hrs = sum([crs[4] for crs in courses if (crs[0], emp[0]) in rec_map])
            ws1.append([emp[1], emp[2], emp[3], emp[4], t_hrs])

        ws2 = wb.create_sheet(title="課程開課統計表")
        ws2.append(["課程日期", "課程名稱", "類別", "時數", "完訓人數"])
        for crs in courses:
            cnt = sum([1 for emp in employees if (crs[0], emp[0]) in rec_map])
            ws2.append([crs[1], crs[2], crs[3], crs[4], cnt])

        ws3 = wb.create_sheet(title="完訓明細紀錄")
        ws3.append(["工號", "姓名", "職位", "課程名稱", "完訓日期", "時數"])
        for crs in courses:
            for emp in employees:
                if (crs[0], emp[0]) in rec_map:
                    ws3.append([emp[1], emp[2], emp[3], crs[2], rec_map[(crs[0], emp[0])], crs[4]])

    wb.save(output)
    output.seek(0)
    filename = f"Training_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    app.run(debug=True)
