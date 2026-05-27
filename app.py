import os
import psycopg2
import psycopg2.extras
from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
import pytz
from datetime import datetime
import io
from openpyxl import Workbook

app = Flask(__name__)
CORS(app)

# 数据库配置
DATABASE_URL = os.environ.get('DATABASE_URL')
if not DATABASE_URL:
    raise Exception("DATABASE_URL environment variable not set")
# 修复 postgres:// 为 postgresql://
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

def get_db_connection():
    return psycopg2.connect(DATABASE_URL)

def init_db():
    with get_db_connection() as conn:
        with conn.cursor() as cur:
            cur.execute('''
                CREATE TABLE IF NOT EXISTS companies (
                    id SERIAL PRIMARY KEY,
                    name TEXT UNIQUE NOT NULL,
                    capacity INTEGER DEFAULT 5,
                    description TEXT DEFAULT '',
                    category TEXT DEFAULT ''
                )
            ''')
            cur.execute('''
                CREATE TABLE IF NOT EXISTS registrations (
                    id SERIAL PRIMARY KEY,
                    student_name TEXT NOT NULL,
                    class_name TEXT NOT NULL,
                    company_id INTEGER REFERENCES companies(id),
                    ip_address TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(student_name, class_name)
                )
            ''')
            # 初始化企业数据
            cur.execute("SELECT COUNT(*) FROM companies")
            if cur.fetchone()[0] == 0:
                companies_data = [
                    ('常高新金隆控股（集团）有限公司', 5, '常州高新区金融投资平台', '金融投资'),
                    ('常州常高新新能源产业投资有限公司', 5, '聚焦新能源领域投资', '新能源'),
                    ('工商银行', 5, '大型国有商业银行', '金融'),
                    ('贺尔碧格压缩机技术（中国）有限公司', 5, '压缩机技术解决方案', '装备制造'),
                    ('梅特勒-托利多（常州）测量技术有限公司', 5, '精密仪器领导者', '仪器仪表'),
                    ('常州千红生化制药股份有限公司', 5, '生物医药', '生物医药'),
                    ('常州三新供电服务有限公司', 5, '电力供应服务', '电力能源'),
                    ('小松（常州）工程机械有限公司', 5, '工程机械制造商', '工程机械'),
                    ('太平洋财产保险股份有限公司常州分公司', 5, '财产保险', '保险'),
                    ('天纳克（常州）减振系统有限公司', 5, '汽车减振系统', '汽车零部件'),
                    ('威乐（常州）水泵有限公司', 5, '水泵及水处理', '泵业制造')
                ]
                for comp in companies_data:
                    cur.execute('INSERT INTO companies (name, capacity, description, category) VALUES (%s, %s, %s, %s)', comp)
            conn.commit()

# 初始化数据库（在应用启动时执行）
init_db()

# ---------- 配置 ----------
CHINA_TZ = pytz.timezone('Asia/Shanghai')
START_TIME = datetime(2026, 6, 1, 8, 0, 0)
END_TIME   = datetime(2026, 6, 7, 20, 0, 0)
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
    with get_db_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute('''
                SELECT c.id, c.name, c.capacity, c.description, c.category,
                       COUNT(r.id) as registered
                FROM companies c
                LEFT JOIN registrations r ON c.id = r.company_id
                GROUP BY c.id
                ORDER BY c.category, c.id
            ''')
            rows = cur.fetchall()
            companies = []
            for row in rows:
                companies.append({
                    'id': row['id'],
                    'name': row['name'],
                    'capacity': row['capacity'],
                    'description': row['description'],
                    'category': row['category'],
                    'registered': row['registered'],
                    'is_full': row['registered'] >= row['capacity']
                })
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
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cur:
                # 检查企业容量
                cur.execute("SELECT capacity FROM companies WHERE id = %s", (company_id,))
                row = cur.fetchone()
                if not row:
                    return jsonify({'success': False, 'error': '企业不存在'}), 404
                capacity = row[0]
                cur.execute("SELECT COUNT(*) FROM registrations WHERE company_id = %s", (company_id,))
                current_cnt = cur.fetchone()[0]
                if current_cnt >= capacity:
                    return jsonify({'success': False, 'error': '该企业名额已满'}), 409
                # 插入报名
                cur.execute('''
                    INSERT INTO registrations (student_name, class_name, company_id, ip_address)
                    VALUES (%s, %s, %s, %s)
                ''', (student_name, class_name, company_id, ip))
                conn.commit()
                return jsonify({'success': True, 'message': '报名成功'})
    except psycopg2.IntegrityError as e:
        if 'unique constraint' in str(e).lower():
            return jsonify({'success': False, 'error': '您已经报名过了，每人仅限一次'}), 409
        else:
            return jsonify({'success': False, 'error': '数据库错误，请稍后重试'}), 500
    except Exception as e:
        print(f"报名错误: {e}")
        return jsonify({'success': False, 'error': '服务器错误，请稍后重试'}), 500

@app.route('/admin/export')
def export_excel():
    password = request.args.get('password', '')
    if password != ADMIN_PASSWORD:
        return jsonify({'error': '密码错误'}), 401
    with get_db_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute('''
                SELECT r.id, r.student_name, r.class_name, c.name as company_name,
                       r.ip_address, r.created_at
                FROM registrations r
                JOIN companies c ON r.company_id = c.id
                ORDER BY r.created_at DESC
            ''')
            rows = cur.fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "报名记录"
    ws.append(['ID', '学生姓名', '班级', '意向企业', 'IP地址', '报名时间'])
    for row in rows:
        ws.append([row['id'], row['student_name'], row['class_name'], row['company_name'], row['ip_address'], row['created_at'].strftime('%Y-%m-%d %H:%M:%S')])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='报名统计.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)