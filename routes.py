import logging
from flask import Blueprint, request, jsonify, current_app
import os
import sqlite3
import json
import requests
from pathlib import Path
import mammoth
import re
from werkzeug.utils import secure_filename
import PyPDF2
from qwen_client import call_dashscope_api
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import shutil

# 创建蓝图
bp = Blueprint('bidding', __name__)

# 临时的内存存储，用于在 upload -> pre-analysis -> chapter-analysis 之间传递小量状态
temp_analysis_store = {}
_temp_store_lock = threading.Lock()

# 允许上传的文件类型
ALLOWED_EXTENSIONS = {'.docx', '.doc', '.pdf'}

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS

def get_db():
    """获取数据库连接"""
    conn = sqlite3.connect('bidding.db', timeout=30, check_same_thread=False)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA synchronous=NORMAL')
    conn.row_factory = sqlite3.Row
    return conn

def read_tender_file(bidding_id):
    """读取招标文件，失败时抛出异常"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM bidding WHERE id = ?', (bidding_id,))
    bidding = cursor.fetchone()
    conn.close()
    if not bidding:
        raise Exception('招标书不存在')
    file_path = Path(bidding['storage_path'])
    if not file_path.exists():
        raise Exception(f'文件不存在: {file_path}')

    # 先读文件头判断真实类型（magic bytes）
    with open(file_path, 'rb') as f:
        header = f.read(16)

    is_pdf_by_content = header.startswith(b'%PDF')
    is_pdf_by_ext = file_path.suffix.lower() == '.pdf'

    if is_pdf_by_content or is_pdf_by_ext:
        return _read_pdf(file_path)

    # 非 PDF，尝试 mammoth（DOCX/DOC）
    try:
        with open(bidding['storage_path'], 'rb') as f:
            result = mammoth.extract_raw_text(f)
        return result.value
    except Exception as e:
        error_msg = str(e).lower()
        # 如果 mammoth 报 zip 错误，尝试用 PDF 方式兜底
        if 'zip' in error_msg or 'file is not' in error_msg:
            return _read_pdf(file_path)
        raise

def _read_pdf(file_path):
    """读取PDF文件，失败时抛出异常"""
    text = ""
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
    if not text.strip():
        raise Exception("PDF文件内容为空或无法提取文本")
    return text
    
def save_bid_section(content, section_name, output_dir, tender_name):
    '''保存投标文件小节'''
    output_path = Path(output_dir) / tender_name / f"{section_name}.txt"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
    except Exception as e:
            logging.error(f"保存章节 {section_name} 时出错: {e}")

def merge_sections(output_dir, tender_name, sections):
        """合并所有章节内容为一个完整的文档"""
        sections_dir = Path(output_dir) / tender_name
        if not sections_dir.exists():
            logging.error(f"目录 {sections_dir} 不存在！")
            return
        
        # 获取所有章节文件
        section_files = list(sections_dir.glob("*.txt"))
        if not section_files:
            logging.error(f"在 {sections_dir} 目录下未找到章节文件！")
            return
        
        # 创建合并后的文档
        merged_content = ["# 投标文件\n\n"] + [
            f"## {section_name}\n\n{content}\n\n"
            for section_name in sections
            if (section_file := sections_dir / f"{section_name}.txt").exists()
            for content in [section_file.read_text(encoding='utf-8')]
        ]
        output_file = sections_dir / f"{tender_name}_完整投标文件.md"
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("\n".join(merged_content))
            logging.info(f"已生成完整投标文件：{output_file}")
            return output_file
        except Exception as e:
            logging.error(f"保存合并文件时出错: {e}")
            return None    

@bp.route('/upload', methods=['POST'])
def upload_bidding():
    """上传招标文件"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded.'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected.'}), 400

    # 文件类型校验
    if not allowed_file(file.filename):
        return jsonify({'error': '不支持的文件类型，仅允许 .docx、.doc、.pdf'}), 400

    user_id = request.form.get('userId')
    if not user_id:
        return jsonify({'error': 'User ID is required for upload.'}), 400

    try:
        import uuid
        original_filename = file.filename
        safe_filename = secure_filename(original_filename)
        unique_filename = f"{uuid.uuid4()}-{safe_filename}"
        file_path = os.path.join(current_app.config['UPLOAD_FOLDER'], unique_filename)

        # 保存文件到 uploads 目录
        file.save(file_path)

        # 校验文件大小（防止空文件）
        if os.path.getsize(file_path) == 0:
            os.remove(file_path)
            return jsonify({'error': '上传的文件为空'}), 400

        # 写入 DB
        document_key = str(uuid.uuid4())
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO bidding (user_id, original_filename, storage_path, document_key, status)
            VALUES (?, ?, ?, ?, ?)
        ''', (user_id, original_filename, file_path, document_key, 'Uploaded'))
        bidding_id = cursor.lastrowid
        conn.commit()
        conn.close()

        # 在内存临时存储中记录初始条目，analysisData 和 directoryStructure 先为空
        try:
            with _temp_store_lock:
                temp_analysis_store[bidding_id] = {
                    'biddingId': bidding_id,
                    'analysisData': None,
                    'directoryStructure': None,
                }
        except Exception:
            logging.exception('初始化 temp_analysis_store 失败')

        # 返回 minimal 信息（前端随后调用 /generate-bid-document）
        return jsonify({
            'message': 'File uploaded successfully.',
            'biddingId': bidding_id,
            'originalFilename': original_filename
        }), 201

    except Exception as e:
        logging.exception("upload_bidding failed")
        return jsonify({'error': f'Server error during file upload: {str(e)}'}), 500
    
@bp.route('/pre-analysis_bid', methods=['POST'])
def pre_analysis_bid():
    """预处理招标文件"""
    data = request.get_json()
    bidding_id = data.get('biddingId')
    if not bidding_id:
        return jsonify({'error': 'Missing biddingId'}), 400
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM bidding WHERE id = ?', (bidding_id,))
        bidding = cursor.fetchone()
        conn.close()
        if not bidding:
            return jsonify({'error': '招标书不存在'}), 404
        # 读取文件内容
        try:
            bid_content = read_tender_file(bidding_id)
        except FileNotFoundError as fe:
            logging.error(f"文件不存在: {fe}")
            return jsonify({'error': f'招标文件已丢失，请重新上传：{fe}'}), 404
        except Exception as fe:
            # read_tender_file 用 Exception 抛了"文件不存在"
            msg = str(fe)
            if '文件不存在' in msg or '不存在' in msg:
                logging.error(f"招标文件丢失: {msg}")
                return jsonify({'error': f'招标文件已丢失，请重新上传该项目'}), 404
            logging.exception('读取招标文件失败')
            return jsonify({'error': f'读取招标文件失败: {msg}'}), 500

        # 对过长的内容进行截断
        max_content_length = 20000
        if len(bid_content) > max_content_length:
            head_size = 15000
            tail_size = 5000
            bid_content = bid_content[:head_size] + \
                "\n\n...[中间内容省略]...\n\n" + \
                bid_content[-tail_size:]
            logging.info(f"预分析：招标文件内容过长，已截断为{len(bid_content)}字符")

        pre_analysis_prompt =  f'''
        你是一个资深的招投标文件分析师，请根据以下招标书内容，提炼出完整信息，并严格按照下面的JSON格式返回你的分析结果，不要有任何多余的解释，只返回以下json内容。
        {{
            "bidding_requirements":"...",
            "bidding_summary":"...",
            "bidding_meta":"..."
        }}
        "bidding_requirements": 必须包含的文件和材料。
        "bidding_summary":对招标书内容的总结，包括服务内容、服务期限与服务地点以及质量标准等。
        "bidding_meta":招标书中具体的实质性要求内容和评分标准。
        招标书内容如下:
        ---
        {bid_content}
        ---
        '''
        response = call_dashscope_api([
            {'role': 'user', 'content': pre_analysis_prompt}
        ], timeout=300)
        # print(f'[INFO] Pre-analysis response: {response}')
        try:
            http_data = response['choices'][0]['message']['content']
        except (KeyError, IndexError, TypeError):
            return jsonify({'error': 'API响应格式错误'}), 500

        # 解析JSON响应
        clean_json = re.sub(r'<think>.*?</think>', '', http_data, flags=re.DOTALL).strip()
        json_match = re.search(r'```json\n(.*?)\n```', clean_json, re.DOTALL)

        try:
            if json_match:
                analysis_result = json.loads(json_match.group(1), strict=False)
            else:
                clean_json = clean_json.replace('```json', '').replace('```', '').strip()
                analysis_result = json.loads(clean_json, strict=False)
            # 将 pre-analysis 的结果写入临时存储（如果存在对应的 bidding_id）
            try:
                with _temp_store_lock:
                    if bidding_id in temp_analysis_store:
                        temp_analysis_store[bidding_id]['analysisData'] = analysis_result
                    else:
                        temp_analysis_store[bidding_id] = {
                            'biddingId': bidding_id,
                            'analysisData': analysis_result,
                            'directoryStructure': None,
                        }
            except Exception:
                logging.exception('写入 temp_analysis_store.analysisData 失败')
        except Exception as e:
            print(f'[ERROR] JSON解析失败: {str(e)}')
            return jsonify({'error': 'API返回内容解析失败'}), 500
        return jsonify(analysis_result)

    except Exception as e:
        print(f'[ERROR] Pre-analysis failed for bidding {bidding_id}: {str(e)}')
        return jsonify({'error': '预分析失败，请稍后重试。'}), 500        

@bp.route('/chapter-analysis_bid', methods=['POST'])
def chapter_analysis_bid():
    """招标文件章节分析"""
    data = request.get_json()
    
    bidding_id = data.get('biddingId')
    if not bidding_id:
        return jsonify({'error': 'Missing biddingId'}), 400
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM bidding WHERE id = ?', (bidding_id,))
        bidding = cursor.fetchone()
        conn.close()
        if not bidding:
            return jsonify({'error': '招标书不存在'}), 404
        # 读取文件内容
        try:
            bid_content = read_tender_file(bidding_id)
        except Exception as fe:
            msg = str(fe)
            if '文件不存在' in msg or '不存在' in msg:
                return jsonify({'error': '招标文件已丢失，请重新上传该项目'}), 404
            return jsonify({'error': f'读取招标文件失败: {msg}'}), 500

        # 对过长的内容进行截断，保留前20000字符（包含目录和格式要求等关键信息）
        max_content_length = 20000
        if len(bid_content) > max_content_length:
            # 保留开头部分（通常包含目录和格式要求）和结尾部分
            head_size = 15000
            tail_size = 5000
            bid_content_truncated = bid_content[:head_size] + \
                "\n\n...[中间内容省略]...\n\n" + \
                bid_content[-tail_size:]
            logging.info(f"招标文件内容过长({len(bid_content)}字符)，已截断为{len(bid_content_truncated)}字符")
        else:
            bid_content_truncated = bid_content

        post_analysis_prompt = f'''
        你是一个资深的招投标文件结构分析师，请根据以下招标书内容，输出投标书其他响应文件的格式章节内容（除封面章节），并严格按照下面的JSON格式返回你的分析结果，不要有任何多余的解释。
        {{
            "chapter_format":"..."
        }}
        "chapter_format":招标书中的其他响应文件格式，如果有表格内容，请用Markdown表格的形式返回。
        招标书内容如下:
        ---
        {bid_content_truncated}
        ---
        '''
        
        # 使用较长超时并支持重试
        response = None
        last_error = None
        for attempt in range(2):
            try:
                response = call_dashscope_api([
                    {'role': 'user', 'content': post_analysis_prompt}
                ], timeout=300)
                break
            except requests.exceptions.Timeout:
                last_error = "API请求超时"
                logging.warning(f"章节分析第{attempt+1}次尝试超时，{'重试中...' if attempt == 0 else '放弃'}")
            except requests.exceptions.ConnectionError as ce:
                last_error = f"网络连接错误: {ce}"
                logging.warning(f"章节分析第{attempt+1}次连接错误: {ce}")
        
        if response is None:
            return jsonify({'error': f'章节分析失败: {last_error}，请稍后重试'}), 504

        try:
             http_data = response['choices'][0]['message']['content']
        except (KeyError, IndexError, TypeError):
             return jsonify({'error': 'API响应格式错误'}), 500


        #解析JSON响应
        clean_json = re.sub(r'<think>.*?</think>', '', http_data, flags=re.DOTALL).strip()
        json_match = re.search(r'```json\n(.*?)\n```', clean_json, re.DOTALL)

        try:
            if json_match:
                analysis_result = json.loads(json_match.group(1), strict=False)
            else:
                clean_json = clean_json.replace('```json', '').replace('```', '').strip()
                analysis_result = json.loads(clean_json, strict=False)
        except Exception as e:
            try:
                import ast
                clean_json = re.sub(r'<think>.*?</think>', '', http_data, flags=re.DOTALL).strip()
                clean_json = clean_json.replace('```json', '').replace('```', '').strip()
                analysis_result = ast.literal_eval(clean_json)
            except Exception as e2:
                print(f"[ERROR] JSON parse error: {e}, Fallback also failed: {e2}")
                return jsonify({'error': 'AI生成的章节格式无法解析'}), 500
            
        # 将结构存入 temp_analysis_store
        with _temp_store_lock:
            if bidding_id not in temp_analysis_store:
                temp_analysis_store[bidding_id] = {}
            temp_analysis_store[bidding_id]['directoryStructure'] = analysis_result
            
        return jsonify(analysis_result)

    except Exception as e:
        print(f'[ERROR] Chapter-analysis failed for bidding {bidding_id}: {str(e)}')
        return jsonify({'error': '章节提取分析失败，请稍后重试。'}), 500
    
@bp.route('/chapter-design', methods=['POST'])
def chapter_design():
    """投标文件章节设计"""
    data = request.get_json()
    bidding_id = data.get('biddingId') 
    logging.info(f"chapter-design called for bidding_id: {bidding_id}")
    
    # 获取分析结果和目录结构
    analysis_data = temp_analysis_store.get(bidding_id, {}).get('analysisData')
    directory_structure = temp_analysis_store.get(bidding_id, {}).get('directoryStructure')

    if not all([bidding_id, analysis_data, directory_structure]):
        return jsonify({'error': 'Incomplete analysis request'}), 400

    # 提取招标书信息
    bidding_requirements = analysis_data.get('bidding_requirements', '')
    bidding_summary = analysis_data.get('bidding_summary', '')
    bidding_meta = analysis_data.get('bidding_meta', '')

    # 精简directory_structure，避免prompt过长
    dir_str = json.dumps(directory_structure, ensure_ascii=False) if isinstance(directory_structure, dict) else str(directory_structure)
    if len(dir_str) > 8000:
        dir_str = dir_str[:8000] + "\n...[内容过长已截断]"
    
    # 精简其他字段
    if len(str(bidding_meta)) > 5000:
        bidding_meta = str(bidding_meta)[:5000] + "...[已截断]"
    if len(str(bidding_summary)) > 3000:
        bidding_summary = str(bidding_summary)[:3000] + "...[已截断]"

    # 构建提示词
    chapter_design_prompt = (
        f"你是一个资深的投标文件目录结构设计专家，请根据以下信息，整理出最终的标书章节结构：\n\n"
        f"必须包含的文件和材料：{bidding_requirements}\n"
        f"招标书内容总结：{bidding_summary}\n"
        f"招标书具体要求和评分标准：{bidding_meta}\n"
        f"投标书章节大纲：{dir_str}\n\n"
        "基于以上招标文件要求和行业经验，请补充章节的子节目录，确保投标文件完整且符合要求。\n"
        "要求：\n"
        "1、输出的投标书章节结构必须遵循目录结构，并包含所有必要的子章节。\n"
        "2、章节大纲中某一章如果是xxx表、xxx函、xxx清单、封面等，则该章下不需要再细分子节，返回原本的章节内容。\n"
        "3、输出必须是有效的JSON格式，格式如下：\n"
        '''{
  "chapters": [
    {
      "title": "",
      "type": "normal|table",
      "content": "",
      "sections": [
        {
          "title": "",
          "subsections": [
            {
              "title": "",
              "describe": ""
            }
          ]
        }
      ]
    }
  ]
}'''
        "\n字段说明：\n"
        "title：章节标题。\n"
        "type：章节类型，normal表示文本章节，table表示表格章节。\n"
        "content：table章节需填写原本章节内容。\n"
        "sections：二级标题。\n"
        "subsections：三级标题，最少5-7点三级标题。\n"
        "describe：三级标题内容的描述。\n"
    )

    try:
        # 调用 LLM API，支持重试
        response = None
        last_error = None
        for attempt in range(2):
            try:
                response = call_dashscope_api([
                    {'role': 'user', 'content': chapter_design_prompt}
                ], timeout=600)
                break
            except requests.exceptions.Timeout:
                last_error = "API请求超时"
                logging.warning(f"章节设计第{attempt+1}次尝试超时")
            except requests.exceptions.ConnectionError as ce:
                last_error = f"网络连接错误: {ce}"
                logging.warning(f"章节设计第{attempt+1}次连接错误: {ce}")
        
        if response is None:
            return jsonify({'error': f'章节设计失败: {last_error}，请稍后重试'}), 504

        logging.info(f'[INFO] chapter-design response received')

        # 获取返回内容
        try:
            http_data = response['choices'][0]['message']['content']
        except (KeyError, IndexError, TypeError):
            return jsonify({'error': 'API响应格式错误'}), 500

        # 解析 JSON 响应并清理控制字符
        clean_json = re.sub(r'<think>.*?</think>', '', http_data, flags=re.DOTALL).strip()
        json_match = re.search(r'```json\s*(.*?)\s*```', clean_json, re.DOTALL)
        json_text = json_match.group(1) if json_match else clean_json

        # 清理非法控制字符
        json_text = re.sub(r'[\x00-\x1F\x7F]', '', json_text)
        json_text = json_text.replace('\r', '').replace('\t', '').strip()
        json_text = json_text.rstrip(", \n")

        try:
            analysis_result = json.loads(json_text, strict=False)
        except json.JSONDecodeError as e:
            # 如果JSON解析失败，尝试修复常见的JSON错误
            try:
                import ast
                analysis_result = ast.literal_eval(json_text)
            except:
                print("------ JSON Parse Error ------")
                print(f"Error: {e}")
                print("Raw text snippet:")
                print(json_text[:2000])
                return jsonify({'error': f'JSON解析失败: {str(e)}'}), 500

        return jsonify(analysis_result)

    except Exception as e:
        print(f'[ERROR] Chapter-design failed for bidding {bidding_id}: {str(e)}')
        return jsonify({'error': '章节生成失败，请稍后重试。'}), 500

    



@bp.route('/generate-bid-document', methods=['POST'])
def generate_bid_document():
    """生成完整投标书文件（标准Word格式）"""
    from generate_route import generate_bid_document_new
    return generate_bid_document_new(get_db, read_tender_file, temp_analysis_store)


@bp.route('/document-preview/<int:bidding_id>', methods=['GET'])
def document_preview(bidding_id):
    """获取生成的投标文件内容，用于前端只读预览"""
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM bidding WHERE id=?', (bidding_id,))
    bidding = cur.fetchone()
    conn.close()

    if not bidding:
        return jsonify({'error': '项目不存在'}), 404

    # 优先读取生成的Word文件内容
    bid_document_path = ''
    try:
        bid_document_path = bidding['bid_document'] if bidding['bid_document'] else ''
    except (KeyError, IndexError):
        pass
    
    if bid_document_path and Path(bid_document_path).exists() and bid_document_path.endswith('.docx'):
        try:
            from docx import Document as DocxDocument
            doc = DocxDocument(bid_document_path)
            # 提取所有段落文本，保留结构
            lines = []
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    lines.append(text)
                else:
                    lines.append('')
            content = '\n'.join(lines)
            # 构造下载URL
            safe_name = Path(bid_document_path).name
            port = os.environ.get('PORT', '3012')
            download_url = f"http://localhost:{port}/api/outputs/{safe_name}"
            return jsonify({
                'content': content,
                'title': bidding['original_filename'],
                'fileUrl': download_url,
                'generatedAt': None
            })
        except Exception as e:
            logging.exception(f"读取Word文档失败: {e}")

    # 回退：尝试读取Markdown文件
    tender_name = Path(bidding['original_filename']).stem
    markdown_file = Path("outputs") / tender_name / f"{tender_name}_完整投标文件.md"
    if markdown_file.exists():
        try:
            content = markdown_file.read_text(encoding='utf-8')
            return jsonify({
                'content': content,
                'title': bidding['original_filename'],
                'generatedAt': None
            })
        except Exception as e:
            logging.exception(f"读取Markdown失败: {e}")

    # 再回退：尝试在outputs目录找响应文件
    response_file = Path("outputs") / f"{tender_name}_响应文件.docx"
    if response_file.exists():
        try:
            from docx import Document as DocxDocument
            doc = DocxDocument(str(response_file))
            lines = [para.text for para in doc.paragraphs]
            content = '\n'.join(lines)
            return jsonify({
                'content': content,
                'title': bidding['original_filename'],
                'generatedAt': None
            })
        except Exception as e:
            logging.exception(f"读取响应文件失败: {e}")

    return jsonify({'error': '文档尚未生成'}), 404