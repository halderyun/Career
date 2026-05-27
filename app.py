import os
import sqlite3
from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
import pytz
from datetime import datetime
import io
from openpyxl import Workbook

app = Flask(__name__)
CORS(app)

# 使用 /tmp 目录下的 SQLite 文件（Render 可写）
DB_PATH = '/tmp/career.db'

def init_db():
    """初始化数据库表并插入企业数据（如果为空）"""
    with sqlite3.connect(DB_PATH) as conn:
        # 创建企业表
        conn.execute('''
            CREATE TABLE IF NOT EXISTS companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                capacity INTEGER DEFAULT 5,
                description TEXT DEFAULT '',
                category TEXT DEFAULT ''
            )
        ''')
        # 创建报名表（联合唯一约束防止同姓名+班级重复）
        conn.execute('''
            CREATE TABLE IF NOT EXISTS registrations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_name TEXT NOT NULL,
                class_name TEXT NOT NULL,
                company_id INTEGER,
                ip_address TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(company_id) REFERENCES companies(id),
                UNIQUE(student_name, class_name)
            )
        ''')
        # 检查是否已有企业数据，若无则插入初始数据
        cur = conn.execute("SELECT COUNT(*) FROM companies")
        if cur.fetchone()[0] == 0:
            companies_data = [
                ('常高新金隆控股（集团）有限公司', 4, '常州高新区金融投资平台，服务实体经济发展。', '金融财会'),
                ('常州常高新新能源产业投资有限公司', 4, '聚焦新能源领域投资，推动绿色能源转型。', '科技类'),
                ('工商银行', 10, '中国大型国有商业银行，提供全方位金融服务。', '金融财会'),
                ('贺尔碧格压缩机技术（中国）有限公司', 10, '全球领先的压缩机技术解决方案提供商。', '工程制造业'),
                ('梅特勒-托利多（常州）测量技术有限公司', 10, '精密仪器和称重解决方案的领导者。', '工程制造业'),
                ('常州千红生化制药股份有限公司', 10, '生物医药企业，专注于酶制剂和生化药品。', '科技类'),
                ('常州三新供电服务有限公司', 10, '电力供应与服务企业。', '科技类'),
                ('小松（常州）工程机械有限公司',10, '世界知名工程机械制造商。', '工程制造业'),
                ('太平洋财产保险股份有限公司常州分公司', 5, '财产保险服务，提供风险保障。', '金融财会'),
                ('天纳克（常州）减振系统有限公司', 10, '汽车减振系统及排气系统制造商。', '工程制造业'),
                ('威乐（常州）水泵有限公司', 5, '全球领先的水泵及水处理解决方案供应商。', '工程制造业')
            ]
            for comp in companies_data:
                conn.execute('INSERT INTO companies (name, capacity, description, category) VALUES (?, ?, ?, ?)', comp)
            conn.commit()
    print("数据库初始化完成（表已存在或已创建）")

# 在应用启动时立即初始化数据库（关键：必须放在这里，而不是 if __name__ 内部）
init_db()

# ---------- 配置 ----------
CHINA_TZ = pytz.timezone('Asia/Shanghai')
START_TIME = datetime(2026, 6, 1, 8, 0, 0)   # 请修改为实际开始时间
END_TIME   = datetime(2026, 6, 7, 20, 0, 0)   # 请修改为实际结束时间
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

def is_registration_open():
    now = datetime.now(CHINA_TZ)
    start_local = CHINA_TZ.localize(START_TIME) if START_TIME.tzinfo is None else START_TIME
    end_local = CHINA_TZ.localize(END_TIME) if END_TIME.tzinfo is None else END_TIME
    return start_local <= now <= end_local

def get_client_ip():
    if request.headers.get('X-Forwarded-For'):
        return request.headers.get('X-Forwarded-For').split(',')[0]
    return request.remote_addr

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/check_time')
def check_time():
    return jsonify({
        'open': is_registration_open(),
        'start_time': START_TIME.strftime('%Y-%m-%d %H:%M'),
        'end_time': END_TIME.strftime('%Y-%m-%d %H:%M')
    })

@app.route('/api/companies')
def get_companies():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute('''
        SELECT c.id, c.name, c.capacity, c.description, c.category,
               COUNT(r.id) as registered
        FROM companies c
        LEFT JOIN registrations r ON c.id = r.company_id
        GROUP BY c.id
        ORDER BY c.category, c.id
    ''')
    rows = cur.fetchall()
    companies = [{
        'id': row['id'],
        'name': row['name'],
        'capacity': row['capacity'],
        'description': row['description'],
        'category': row['category'],
        'registered': row['registered'],
        'is_full': row['registered'] >= row['capacity']
    } for row in rows]
    conn.close()
    return jsonify({'success': True, 'companies': companies})

@app.route('/api/register', methods=['POST'])
def register():
    if not is_registration_open():
        return jsonify({'success': False, 'error': '当前不在报名时间段内'}), 403
    data = request.get_json()
    student_name = data.get('student_name', '').strip()
    class_name = data.get('class_name', '').strip()
    company_id = data.get('company_id')
    if not student_name or not class_name or not company_id:
        return jsonify({'success': False, 'error': '请填写姓名和班级'}), 400
    ip = get_client_ip()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        # 检查企业容量
        cur.execute("SELECT capacity FROM companies WHERE id = ?", (company_id,))
        row = cur.fetchone()
        if not row:
            return jsonify({'success': False, 'error': '企业不存在'}), 404
        capacity = row[0]
        cur.execute("SELECT COUNT(*) FROM registrations WHERE company_id = ?", (company_id,))
        current_cnt = cur.fetchone()[0]
        if current_cnt >= capacity:
            return jsonify({'success': False, 'error': '该企业名额已满'}), 409
        # 插入报名（唯一约束会自动阻止重复）
        cur.execute('''
            INSERT INTO registrations (student_name, class_name, company_id, ip_address)
            VALUES (?, ?, ?, ?)
        ''', (student_name, class_name, company_id, ip))
        conn.commit()
        return jsonify({'success': True, 'message': '报名成功'})
    except sqlite3.IntegrityError as e:
        if 'UNIQUE constraint failed' in str(e):
            return jsonify({'success': False, 'error': '您已经报名过了，每人仅限一次'}), 409
        else:
            conn.rollback()
            return jsonify({'success': False, 'error': '数据库错误，请稍后重试'}), 500
    except Exception as e:
        conn.rollback()
        print(f"报名错误: {e}")
        return jsonify({'success': False, 'error': '服务器错误，请稍后重试'}), 500
    finally:
        conn.close()

@app.route('/admin/export')
def export_excel():
    password = request.args.get('password', '')
    if password != ADMIN_PASSWORD:
        return jsonify({'error': '密码错误'}), 401
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute('''
        SELECT r.id, r.student_name, r.class_name, c.name as company_name,
               r.ip_address, r.created_at
        FROM registrations r
        JOIN companies c ON r.company_id = c.id
        ORDER BY r.created_at DESC
    ''')
    rows = cur.fetchall()
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "报名记录"
    ws.append(['ID', '学生姓名', '班级', '意向企业', 'IP地址', '报名时间'])
    for row in rows:
        ws.append([row['id'], row['student_name'], row['class_name'], row['company_name'], row['ip_address'], row['created_at']])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='报名统计.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)