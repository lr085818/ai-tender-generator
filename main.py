from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
import os
from dotenv import load_dotenv
import sqlite3
from datetime import datetime
import uuid
import json

# 加载环境变量
load_dotenv()

app = Flask(__name__)
CORS(app)

# 配置
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['GENERATED_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)

import routes
import users

# 注册蓝图
app.register_blueprint(routes.bp, url_prefix='/api/bidding')
app.register_blueprint(users.bp, url_prefix='/api/users')

# 全局数据库连接辅助
def get_db():
    """获取数据库连接（统一入口，所有模块共用）"""
    conn = sqlite3.connect('bidding.db', timeout=30, check_same_thread=False)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA synchronous=NORMAL')
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """初始化数据库表结构（只应调用一次）"""
    conn = get_db()
    cursor = conn.cursor()
    
    # 创建用户表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fingerprint_id TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 创建招投标文件表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bidding (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            original_filename TEXT NOT NULL,
            storage_path TEXT NOT NULL,
            document_key TEXT UNIQUE NOT NULL,
            status TEXT DEFAULT 'Uploaded',
            other_response_format TEXT,
            bid_document TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/')
def index():
    response = send_file('templates/index.html')
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.route('/api')
def api_index():
    return jsonify({
        'service': 'AI智能招投标文档生成系统-后端',
        'status': 'running',
        'api_endpoints': {
            '上传招标文件': 'POST /api/bidding/upload',
            '预分析招标文件': 'POST /api/bidding/pre-analysis_bid',
            '章节分析': 'POST /api/bidding/chapter-analysis_bid',
            '章节设计': 'POST /api/bidding/chapter-design',
            '生成投标文件': 'POST /api/bidding/generate-bid-document',
            'OnlyOffice保存回调': 'POST /api/bidding/save-callback',
            '用户识别': 'POST /api/users/identify',
        }
    })

@app.route('/api/bidding/list', methods=['GET'])
def list_bidding():
    user_id = request.args.get('userId')
    conn = get_db()
    cursor = conn.cursor()
    if user_id:
        cursor.execute('SELECT id, original_filename, status, created_at FROM bidding WHERE user_id = ? ORDER BY id DESC', (user_id,))
    else:
        cursor.execute('SELECT id, original_filename, status, created_at FROM bidding ORDER BY id DESC')
    rows = cursor.fetchall()
    conn.close()
    return jsonify({'biddingList': [dict(r) for r in rows]})

@app.route('/api/outputs/<path:filename>')
def uploaded_file(filename):
    print(f"Trying to serve file: {filename}")
    file_path = os.path.join(app.config['GENERATED_FOLDER'], filename)
    print(f"Full path: {file_path}, Exists: {os.path.exists(file_path)}")
    return send_from_directory(app.config['GENERATED_FOLDER'], filename)


if __name__ == '__main__':
    # 仅在主进程中初始化数据库（禁用 reloader 避免双进程竞争）
    init_db()
    port = int(os.environ.get('PORT', 3012))
    debug = os.environ.get('FLASK_DEBUG', 'true').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug, use_reloader=False,
            exclude_patterns=['outputs/*', 'uploads/*', '__pycache__/*']) 