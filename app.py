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
                ('常高新金隆控股（集团）有限公司', 4, '地址：新北区秀水河路智富中心。常高新是常州高新区金融投资平台，服务实体经济发展，主要从事以自有资金从事投资活动、融资咨询服务等业务。', '金融财会'),
                ('常州常高新新能源产业投资有限公司', 4, '地址：新北区秀水河路智富中心。常高新新能源主要聚焦新能源领域投资，推动绿色能源转型。', '科创技术'),
                ('贺尔碧格传动技术（常州）有限公司', 10, '地址：新北区创业东路16号粤海工业园7A。贺尔碧格传动技术作为贺尔碧格集团的一大事业部，是汽车行业公认的换档装置和离合器方面的专家，主要生产汽车自动变速箱及其零部件。', '工程制造'),
                ('梅特勒-托利多（常州）测量技术有限公司', 10, '地址：新北区太湖西湖111号。梅特勒托利多是目前精密仪器和称重解决方案的领导者。', '工程制造'),
                ('常州千红生化制药股份有限公司', 15, '地址：新北区云河路518号。生物医药企业，专注于酶制剂和生化药品，主要生产冻干粉针剂（含抗肿瘤药）、小容量注射剂、片剂、胶囊剂、原料药。', '科创技术'),
                ('常州三新供电服务有限公司', 6, '地址：天宁区青洋北路157号。国网江苏省电力有限公司所属企业，主要从事乡镇供电所的生产经营业务，具体为农村配网线路和设备的巡视、检修和故障处理，电力业扩报装、装表接电、用电检查、电费回收和客户服务等工作。', '科创技术'),
                ('小松（常州）工程机械有限公司',10, '地址：新北区黄河西路389号。世界知名工程机械制造商，常州公司年产能为1万台，产品广泛用于矿山开采、交通建设等；拥有年产2.2万吨高效环保铸铁生产线，主要生产工程机械用高强度缸体及油压件。', '工程制造'),
                ('中信银行常州支行', 12, '地址：天宁区吊桥路1-41号。中信银行是中国改革开放中最早成立的新兴商业银行之一， 是中国最早参与国内外金融市场融资的商业银行，也是一家主营金融业务的股份有限公司。', '金融财会'),
                ('天纳克（常州）减振系统有限公司', 10, '地址：新北区浏阳河路36号。天纳克常州是美国天纳克集团全资子公司，年产机动车减振器、独立悬挂、减振器芯总成和专用减振器1800万只，客户涵盖如理想、大众、通用、长安福特、奔驰、宝马、吉利汽车等40多家国内外知名主机厂。', '工程制造'),
                ('威乐（常州）水泵有限公司', 5, '地址：新北区河海西路350号。威乐集团主要生产水泵，产品被广泛用于建筑领域、市政建设、水务、工业等领域。常州高新区项目集聚研发、生产、配送等多个功能，使用最先进生产设施，运用威乐德国总部新落成的超现代化智能工厂的运营方式与数字化标准流程。', '工程制造'),
                ('常州书香之家文化科技有限公司', 6, '地址：常州新华书店（和平北路店）。“阅读＋教育＋心理”的跨界成长空间，展示如何运用教育心理学激发阅读力与思辨力，是探索教育与文化传播行业的绝佳窗口。', '人文教育')
            ]
            for comp in companies:
                conn.execute('INSERT INTO companies (name, capacity, description, category) VALUES (?,?,?,?)', comp)
            conn.commit()
    print("数据库初始化完成")

init_db()

CHINA_TZ = pytz.timezone('Asia/Shanghai')
START_TIME = datetime(2026, 6, 2, 13, 46, 0)
END_TIME   = datetime(2026, 6, 3, 23, 59, 59)
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin5525')

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
