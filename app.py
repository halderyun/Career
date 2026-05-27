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

DB_PATH = '/tmp/career.db'  # Render 可写目录

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                capacity INTEGER DEFAULT 5,
                description TEXT DEFAULT '',
                category TEXT DEFAULT ''
            )
        ''')
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
        cur = conn.execute("SELECT COUNT(*) FROM companies")
        if cur.fetchone()[0] == 0:
            companies = [
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
            for comp in companies:
                conn.execute('INSERT INTO companies (name, capacity, description, category) VALUES (?,?,?,?)', comp)
            conn.commit()
    print("数据库初始化完成")

init_db()

CHINA_TZ = pytz.timezone('Asia/Shanghai')
START_TIME = datetime(2026, 6, 1, 8, 0, 0)
END_TIME   = datetime(2026, 6, 7, 20, 0, 0)
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

def is_open():
    now = datetime.now(CHINA_TZ)
    start = CHINA_TZ.localize(START_TIME) if START_TIME.tzinfo is None else START_TIME
    end = CHINA_TZ.localize(END_TIME) if END_TIME.tzinfo is None else END_TIME
    return start <= now <= end

def get_ip():
    if request.headers.get('X-Forwarded-For'):
        return request.headers.get('X-Forwarded-For').split(',')[0]
    return request.remote_addr

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/check_time')
def check_time():
    return jsonify({'open': is_open(), 'start_time': START_TIME.strftime('%Y-%m-%d %H:%M'), 'end_time': END_TIME.strftime('%Y-%m-%d %H:%M')})

@app.route('/api/companies')
def companies():
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
    result = []
    for row in rows:
        result.append({
            'id': row['id'],
            'name': row['name'],
            'capacity': row['capacity'],
            'description': row['description'],
            'category': row['category'],
            'registered': row['registered'],
            'is_full': row['registered'] >= row['capacity']
        })
    conn.close()
    return jsonify({'success': True, 'companies': result})

@app.route('/api/register', methods=['POST'])
def register():
    if not is_open():
        return jsonify({'success': False, 'error': '报名未开放'}), 403
    data = request.get_json()
    name = data.get('student_name', '').strip()
    cls = data.get('class_name', '').strip()
    cid = data.get('company_id')
    if not name or not cls or not cid:
        return jsonify({'success': False, 'error': '请填写完整'}), 400
    ip = get_ip()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        cur.execute("SELECT capacity FROM companies WHERE id=?", (cid,))
        cap = cur.fetchone()
        if not cap:
            return jsonify({'success': False, 'error': '企业不存在'}), 404
        cur.execute("SELECT COUNT(*) FROM registrations WHERE company_id=?", (cid,))
        cnt = cur.fetchone()[0]
        if cnt >= cap[0]:
            return jsonify({'success': False, 'error': '名额已满'}), 409
        cur.execute('INSERT INTO registrations (student_name, class_name, company_id, ip_address) VALUES (?,?,?,?)',
                    (name, cls, cid, ip))
        conn.commit()
        return jsonify({'success': True, 'message': '报名成功'})
    except sqlite3.IntegrityError:
        return jsonify({'success': False, 'error': '您已经报名过了'}), 409
    except Exception as e:
        conn.rollback()
        print(e)
        return jsonify({'success': False, 'error': '服务器错误'}), 500
    finally:
        conn.close()

@app.route('/admin/export')
def export():
    pwd = request.args.get('password')
    if pwd != ADMIN_PASSWORD:
        return '密码错误', 401
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
    ws.append(['ID', '姓名', '班级', '企业', 'IP', '时间'])
    for row in rows:
        ws.append([row['id'], row['student_name'], row['class_name'], row['company_name'], row['ip_address'], row['created_at']])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='报名统计.xlsx')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)