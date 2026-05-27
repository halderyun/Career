import sqlite3
import os
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, send_file
from flask_cors import CORS
import pytz
import io
from openpyxl import Workbook

app = Flask(__name__)
CORS(app)

# ================= 配置 =================
CHINA_TZ = pytz.timezone('Asia/Shanghai')
START_TIME = datetime(2026, 5, 27, 8, 0, 0)
END_TIME   = datetime(2026, 6, 7, 20, 0, 0)

def is_registration_open():
    now = datetime.now(CHINA_TZ)
    start_local = CHINA_TZ.localize(START_TIME) if START_TIME.tzinfo is None else START_TIME
    end_local = CHINA_TZ.localize(END_TIME) if END_TIME.tzinfo is None else END_TIME
    return start_local <= now <= end_local

def init_db():
    with sqlite3.connect('career.db') as conn:
        # 企业表
        conn.execute('''
            CREATE TABLE IF NOT EXISTS companies (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                capacity INTEGER NOT NULL DEFAULT 5,
                description TEXT NOT NULL DEFAULT '',
                category TEXT NOT NULL DEFAULT ''
            )
        ''')
        # 报名表
        conn.execute('''
            CREATE TABLE IF NOT EXISTS registrations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_name TEXT NOT NULL,
                class_name TEXT NOT NULL,
                company_id INTEGER NOT NULL,
                openid TEXT,
                phone_or_student_id TEXT,
                ip_address TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (company_id) REFERENCES companies (id)
            )
        ''')
        # 索引
        conn.execute("CREATE INDEX IF NOT EXISTS idx_openid ON registrations(openid)")
        # 初始化企业数据（如果表为空）
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
                conn.execute('''
                    INSERT INTO companies (name, capacity, description, category)
                    VALUES (?, ?, ?, ?)
                ''', comp)
            conn.commit()

# ================= API =================
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/check_time')
def check_time():
    return jsonify({
        'open': is_registration_open(),
        'start_time': START_TIME.strftime('%Y-%m-%d %H:%M'),
        'end_time': END_TIME.strftime('%Y-%m-%d %H:%M')
    })

@app.route('/api/companies')
def get_companies():
    conn = sqlite3.connect('career.db')
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
    unique_id = data.get('unique_id', '').strip()   # 手机号

    if not student_name or not class_name or not company_id:
        return jsonify({'success': False, 'error': '请填写完整信息'}), 400
    if not unique_id:
        return jsonify({'success': False, 'error': '请确认信息'}), 400

    conn = sqlite3.connect('career.db')
    cur = conn.cursor()

    # 检查唯一性（基于手机号）
    cur.execute("SELECT id FROM registrations WHERE student_name = ?", (student_name,))
    if cur.fetchone():
        conn.close()
        return jsonify({'success': False, 'error': '你已经报名过了，每人仅限一次'}), 409

    # 检查企业名额
    cur.execute("SELECT capacity FROM companies WHERE id = ?", (company_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'error': '企业不存在'}), 404
    capacity = row[0]
    cur.execute("SELECT COUNT(*) FROM registrations WHERE company_id = ?", (company_id,))
    current_cnt = cur.fetchone()[0]
    if current_cnt >= capacity:
        conn.close()
        return jsonify({'success': False, 'error': '该企业名额已满'}), 409

    # 插入报名记录
    ip = request.headers.get('X-Forwarded-For', request.remote_addr)
    cur.execute('''
        INSERT INTO registrations (student_name, class_name, company_id, phone_or_student_id, ip_address)
        VALUES (?, ?, ?, ?, ?)
    ''', (student_name, class_name, company_id, unique_id, ip))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '报名成功'})

# ================= 导出接口 =================
@app.route('/admin/export')
def export_excel():
    password = request.args.get('password', '')
    if password != 'admin123':   # 可修改环境变量
        return jsonify({'error': '密码错误'}), 401

    conn = sqlite3.connect('career.db')
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute('''
        SELECT r.id, r.student_name, r.class_name, c.name as company_name,
               r.phone_or_student_id, r.ip_address, r.created_at
        FROM registrations r
        JOIN companies c ON r.company_id = c.id
        ORDER BY r.created_at DESC
    ''')
    rows = cur.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "报名记录"
    if rows:
        headers = ['ID', '学生姓名', '班级', '意向企业', '学号/手机号', 'IP地址', '报名时间']
        ws.append(headers)
        for row in rows:
            ws.append([
                row['id'],
                row['student_name'],
                row['class_name'],
                row['company_name'],
                row['phone_or_student_id'],
                row['ip_address'],
                row['created_at']
            ])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='报名统计.xlsx')

# ================= 前端 HTML =================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
    <title>职业探索活动报名</title>
    <style>
        *{margin:0;padding:0;box-sizing:border-box}
        body{background:#f5f7fb;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto;padding:16px 12px 40px}
        .container{max-width:640px;margin:0 auto}
        .header{background:linear-gradient(135deg,#1e6f9f,#2c7da0);margin:-16px -12px 20px -12px;padding:28px 20px 24px;color:white;border-radius:0 0 28px 28px}
        .header h1{font-size:1.8rem;font-weight:600}
        .time-status{background:rgba(255,255,255,0.2);border-radius:40px;display:inline-block;padding:6px 14px;font-size:0.8rem;margin-top:12px}
        .group-section{margin-bottom:30px}
        .group-title{font-size:1.3rem;font-weight:600;color:#1e4a6e;background:#eef2fa;padding:8px 16px;border-radius:30px;margin:16px 0 12px;display:inline-block}
        .company-card{background:white;border-radius:24px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);border:1px solid #eef2f6}
        .company-card.active{border-color:#2c7da0}
        .card-header{padding:16px 18px;display:flex;justify-content:space-between;align-items:center;cursor:pointer}
        .company-name{font-size:1.2rem;font-weight:600;color:#0f3b5c}
        .stats{background:#eef2fa;padding:4px 12px;border-radius:30px;font-size:0.8rem}
        .stats.full{background:#fee2e2;color:#b91c1c}
        .intro-area,.form-area{padding:0 18px 12px;border-top:1px solid #edf2f7;display:none}
        .show{display:block}
        .intro-text{color:#4a5b6e;line-height:1.45;margin-top:12px}
        .form-group{margin-bottom:14px}
        label{font-size:0.85rem;font-weight:500;color:#2c5a7a;display:block;margin-bottom:6px}
        input{width:100%;padding:12px 14px;border:1px solid #cbdde9;border-radius:16px;font-size:0.95rem}
        input:focus{outline:none;border-color:#2c7da0}
        .submit-btn{background:#2c7da0;width:100%;border:none;padding:14px;border-radius:40px;color:white;font-weight:600;font-size:1rem;margin-top:8px;cursor:pointer}
        .submit-btn:disabled{background:#b9cddf;cursor:not-allowed}
        .message{margin-top:12px;padding:10px;border-radius:14px;font-size:0.85rem;display:none}
        .success{background:#e0f2e9;color:#276749}
        .error{background:#ffe6e5;color:#c53030}
        .footer{text-align:center;margin-top:30px;font-size:0.7rem;color:#8ba0b5}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>2026CBS职业探索活动</h1>
        <p>选择你想探索的领域，点击企业报名 · 名额有限 报满即止</p>
        <div class="time-status" id="timeStatus">加载中...</div>
    </div>
    <div id="companyList"></div>
    <div class="footer">每人仅限报名一次，请如实填写学号/手机号</div>
</div>
<script>
    const API_BASE = window.location.origin + '/api';
    let companiesData = [], timeOpen = false, expandedId = null, submitting = false;

    async function fetchTimeStatus() {
        try {
            const res = await fetch(API_BASE + '/check_time');
            const data = await res.json();
            timeOpen = data.open;
            document.getElementById('timeStatus').innerHTML = data.open ? `✅ 报名进行中 ${data.start_time} 至 ${data.end_time}` : `⛔ 报名未开始或已结束 开放时间：${data.start_time} 至 ${data.end_time}`;
        } catch(e) { console.error(e); }
    }

    async function fetchCompanies() {
        try {
            const res = await fetch(API_BASE + '/companies');
            const data = await res.json();
            if(data.success) { companiesData = data.companies; renderCompanies(); }
        } catch(e) { console.error(e); document.getElementById('companyList').innerHTML = '<div style="padding:40px;text-align:center">加载失败，请刷新</div>'; }
    }

    function renderCompanies() {
        const container = document.getElementById('companyList');
        if(!companiesData.length) { container.innerHTML = '<div style="padding:40px;text-align:center">暂无企业</div>'; return; }
        const groupMap = new Map();
        companiesData.forEach(c => { const cat = c.category || '其他'; if(!groupMap.has(cat)) groupMap.set(cat, []); groupMap.get(cat).push(c); });
        let html = '';
        for(let [cat, comps] of groupMap.entries()) {
            html += `<div class="group-section"><div class="group-title">${cat}</div>`;
            for(let comp of comps) {
                const isFull = comp.is_full, remaining = comp.capacity - comp.registered;
                const statsText = isFull ? '已满员' : `剩余 ${remaining}/${comp.capacity}`;
                const statsClass = isFull ? 'stats full' : 'stats';
                const expanded = (expandedId === comp.id);
                html += `
                    <div class="company-card ${expanded ? 'active' : ''}" data-id="${comp.id}">
                        <div class="card-header" onclick="toggleCard(${comp.id})">
                            <span class="company-name">🏢 ${comp.name}</span>
                            <span class="${statsClass}">${statsText}</span>
                        </div>
                        <div class="intro-area ${expanded ? 'show' : ''}" id="intro-${comp.id}">
                            <div class="intro-text">${escapeHtml(comp.description)}</div>
                        </div>
                        <div class="form-area ${expanded ? 'show' : ''}" id="form-${comp.id}">
                            <div class="form-group"><label>学生姓名</label><input type="text" id="name-${comp.id}" placeholder="中文姓名" autocomplete="off"></div>
                            <div class="form-group"><label>班级</label><input type="text" id="class-${comp.id}" placeholder="例：A/B/C/D" autocomplete="off"></div>
                            <div class="form-group"><label>手机号</label><input type="text" id="unique-${comp.id}" placeholder="请填写手机号" autocomplete="off"></div>
                            <button class="submit-btn" id="btn-${comp.id}" onclick="submitRegistration(${comp.id})" ${isFull || !timeOpen ? 'disabled' : ''}>${isFull ? '名额已满' : (timeOpen ? '立即报名' : '未开放')}</button>
                            <div id="msg-${comp.id}" class="message"></div>
                        </div>
                    </div>
                `;
            }
            html += `</div>`;
        }
        container.innerHTML = html;
    }

    window.toggleCard = function(id) {
        if(expandedId === id) {
            document.getElementById(`intro-${id}`).classList.remove('show');
            document.getElementById(`form-${id}`).classList.remove('show');
            document.querySelector(`.company-card[data-id="${id}"]`).classList.remove('active');
            expandedId = null;
        } else {
            if(expandedId !== null) {
                document.getElementById(`intro-${expandedId}`).classList.remove('show');
                document.getElementById(`form-${expandedId}`).classList.remove('show');
                document.querySelector(`.company-card[data-id="${expandedId}"]`).classList.remove('active');
            }
            document.getElementById(`intro-${id}`).classList.add('show');
            document.getElementById(`form-${id}`).classList.add('show');
            document.querySelector(`.company-card[data-id="${id}"]`).classList.add('active');
            expandedId = id;
        }
    };

    window.submitRegistration = async function(id) {
        if(submitting) return;
        const company = companiesData.find(c => c.id === id);
        if(!company || company.is_full) { showMsg(id, '该企业已满员', 'error'); return; }
        if(!timeOpen) { showMsg(id, '当前不在报名时间内', 'error'); return; }
        const name = document.getElementById(`name-${id}`).value.trim();
        const cls = document.getElementById(`class-${id}`).value.trim();
        const unique = document.getElementById(`unique-${id}`).value.trim();
        if(!name || !cls) { showMsg(id, '请填写姓名和班级', 'error'); return; }
        if(!unique) { showMsg(id, '请填写学号或手机号', 'error'); return; }

        submitting = true;
        const btn = document.getElementById(`btn-${id}`);
        btn.disabled = true;
        btn.innerText = '提交中...';
        try {
            const res = await fetch(API_BASE + '/register', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ student_name: name, class_name: cls, company_id: id, unique_id: unique })
            });
            const data = await res.json();
            if(data.success) {
                showMsg(id, '报名成功！', 'success');
                document.getElementById(`name-${id}`).value = '';
                document.getElementById(`class-${id}`).value = '';
                document.getElementById(`unique-${id}`).value = '';
                await fetchCompanies();
            } else {
                showMsg(id, data.error || '报名失败', 'error');
                if(data.error && data.error.includes('满')) await fetchCompanies();
            }
        } catch(e) { showMsg(id, '网络错误，请重试', 'error'); }
        finally {
            submitting = false;
            btn.disabled = false;
            if(company.is_full) btn.innerText = '名额已满';
            else if(timeOpen) btn.innerText = '立即报名';
            else btn.innerText = '未开放';
        }
    };

    function showMsg(companyId, msg, type) {
        const div = document.getElementById(`msg-${companyId}`);
        if(!div) return;
        div.textContent = msg;
        div.className = `message ${type}`;
        div.style.display = 'block';
        setTimeout(() => div.style.display = 'none', 3000);
    }

    function escapeHtml(str) {
        if(!str) return '';
        return str.replace(/[&<>]/g, function(m) { if(m==='&') return '&amp;'; if(m==='<') return '&lt;'; if(m==='>') return '&gt;'; return m; });
    }

    async function init() { await fetchTimeStatus(); await fetchCompanies(); }
    init();
</script>
</body>
</html>
"""

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)), debug=False)