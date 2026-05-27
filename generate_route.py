"""
新的投标文件生成路由 - 使用标准Word格式生成器
"""
import os
import json
import logging
import shutil
from pathlib import Path
from flask import request, jsonify, current_app
from werkzeug.utils import secure_filename
from concurrent.futures import ThreadPoolExecutor

from qwen_client import call_dashscope_api
from bid_document_generator import generate_bid_document_word


def generate_bid_document_new(get_db, read_tender_file, temp_analysis_store):
    """生成完整投标书文件（标准Word格式）"""
    data = request.get_json()
    bidding_id = data.get('biddingId')
    chapter_design = data.get('chapterDesign')
    if not bidding_id:
        return jsonify({'error': 'Missing biddingId'}), 400
    if not chapter_design:
        return jsonify({'error': 'Missing chapterDesign'}), 400

    if isinstance(chapter_design, str):
        try:
            chapter_design = json.loads(chapter_design)
        except Exception as e:
            return jsonify({'error': 'chapterDesign JSON 解析失败'}), 400

    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM bidding WHERE id = ?', (bidding_id,))
        bidding = cursor.fetchone()
        conn.close()
        if not bidding:
            return jsonify({'error': '招标书不存在'}), 404

        tender_name = Path(bidding['original_filename']).stem

        # 从预分析存储中获取结构化分析数据
        analysis_data = temp_analysis_store.get(bidding_id, {}).get('analysisData', {}) or {}
        bidding_requirements = analysis_data.get('bidding_requirements', '')
        bidding_summary = analysis_data.get('bidding_summary', '')
        bidding_meta = analysis_data.get('bidding_meta', '')

        project_info = {
            'project_name': tender_name,
            'project_id': '{{项目编号/包号}}',
            'purchaser': '{{采购人}}',
            'agency': '{{采购代理机构}}',
            'supplier_name': '{{供应商名称}}',
            'date': '{{日期}}',
        }

        # 读取招标文件完整内容用于AI参考（保留更多内容）
        try:
            bid_content = read_tender_file(bidding_id)
            bid_content_ref = bid_content[:30000] if len(bid_content) > 30000 else bid_content
        except Exception:
            bid_content_ref = ""

        logging.info("开始AI生成技术方案内容...")

        def gen_implementation():
            prompt = f'''你是一个专业的投标书撰写专家。请根据以下招标文件信息，生成"项目管理及实施方案"章节内容。

项目名称：{tender_name}

=== 招标书内容 ===
{bid_content_ref[:15000]}

=== 预分析 — 招标要求 ===
{bidding_requirements[:2000]}

=== 预分析 — 招标评分标准 ===
{bidding_meta[:2000]}

要求：
1. 严格依据上述招标文件内容生成，实施方案必须匹配该项目的实际需求
2. 包含项目管理架构、实施方案、进度计划、项目达成效果
3. 进度计划和工期要求必须从招标文件中提取，不要使用默认值
4. 字数要求：1500-2500字
5. 直接输出正文内容，不要输出标题'''
            try:
                resp = call_dashscope_api([{'role': 'user', 'content': prompt}], timeout=180)
                return resp['choices'][0]['message']['content']
            except Exception as e:
                logging.error(f"生成实施方案失败: {e}")
                return None

        def gen_training():
            prompt = f'''你是一个专业的投标书撰写专家。请根据以下招标文件信息，生成"培训方案"章节内容。

项目名称：{tender_name}

=== 招标书内容 ===
{bid_content_ref[:12000]}

=== 预分析 — 招标要求 ===
{bidding_requirements[:2000]}

要求：
1. 严格依据上述招标文件内容生成，培训方案必须匹配该项目的实际需求
2. 包含培训目标、对象、时间地点、内容安排、师资、考核方式、资料准备
3. 培训天数、内容必须从招标文件中提取，不要使用默认值
4. 字数要求：2000-3000字
5. 直接输出正文内容，不要输出标题'''
            try:
                resp = call_dashscope_api([{'role': 'user', 'content': prompt}], timeout=180)
                return resp['choices'][0]['message']['content']
            except Exception as e:
                logging.error(f"生成培训方案失败: {e}")
                return None

        def gen_after_sales():
            prompt = f'''你是一个专业的投标书撰写专家。请根据以下招标文件信息，生成"售后服务方案"章节内容。

项目名称：{tender_name}

=== 招标书内容 ===
{bid_content_ref[:10000]}

=== 预分析 — 招标要求 ===
{bidding_requirements[:2000]}

=== 预分析 — 招标评分标准 ===
{bidding_meta[:2000]}

要求：
1. 严格依据上述招标文件内容生成，售后方案必须匹配该项目的实际需求
2. 包含保修期限、保修范围、响应机制、免责条款、其他承诺
3. 保修期限和响应要求必须从招标文件中提取，不要使用默认值
4. 字数要求：800-1200字
5. 直接输出正文内容，不要输出标题'''
            try:
                resp = call_dashscope_api([{'role': 'user', 'content': prompt}], timeout=180)
                return resp['choices'][0]['message']['content']
            except Exception as e:
                logging.error(f"生成售后方案失败: {e}")
                return None

        def gen_equipment_list():
            """从招标文件中提取设备/货物清单"""
            prompt = f'''你是一个专业的招投标分析师。请从以下招标文件内容中，提取出采购设备/货物/服务清单。
严格按照JSON数组格式输出，不要有任何多余解释。

项目名称：{tender_name}

=== 招标文件内容 ===
{bid_content_ref[:20000]}

=== 预分析 — 招标要求 ===
{bidding_requirements[:3000]}

=== 预分析 — 招标总结 ===
{bidding_summary[:2000]}

请提取清单，格式如下：
[
  {{"name": "设备名称1", "brand_model": "{{品牌/型号}}", "quantity": "数量", "unit": "单位", "unit_price": "{{单价}}", "total_price": "{{合计}}"}},
  {{"name": "设备名称2", ...}}
]

要求：
1. 设备名称必须来自招标文件原文
2. 数量、单位必须与招标文件一致
3. 如果招标文件中有品牌/型号要求，填入；否则填"{{品牌/型号}}"
4. 单价和合计填"{{单价}}"/"{{合计}}"占位符
5. 如果是服务类项目，提取服务项目清单
6. 只返回JSON数组，不要任何其他文字'''
            try:
                resp = call_dashscope_api([{'role': 'user', 'content': prompt}], timeout=180)
                content = resp['choices'][0]['message']['content']
                json_match = __import__('re').search(r'```json\s*(\[.*?\])\s*```', content, __import__('re').DOTALL)
                if json_match:
                    content = json_match.group(1)
                content = content.strip().replace('```json', '').replace('```', '').strip()
                equipment_list = json.loads(content)
                logging.info(f"成功从招标文件提取设备清单: {len(equipment_list)} 项")
                return equipment_list
            except Exception as e:
                logging.error(f"提取设备清单失败: {e}")
                return None

        with ThreadPoolExecutor(max_workers=4) as executor:
            future_impl = executor.submit(gen_implementation)
            future_train = executor.submit(gen_training)
            future_after = executor.submit(gen_after_sales)
            future_equip = executor.submit(gen_equipment_list)
            ai_implementation = future_impl.result(timeout=240)
            ai_training = future_train.result(timeout=240)
            ai_after_sales = future_after.result(timeout=240)
            equipment_list = future_equip.result(timeout=240)

        logging.info(f"AI内容生成完成，设备清单: {'已提取' if equipment_list else '未提取到，使用占位'}")

        output_dir = Path('outputs')
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(output_dir / f'{tender_name}_响应文件.docx')

        generated_docx_path = generate_bid_document_word(
            project_info=project_info,
            chapter_design=chapter_design,
            equipment_list=equipment_list,
            ai_implementation=ai_implementation,
            ai_training=ai_training,
            ai_after_sales=ai_after_sales,
            output_path=output_path
        )

        generated_docx_path = Path(generated_docx_path)
        gen_folder = Path(current_app.config.get('GENERATED_FOLDER', 'outputs'))
        gen_folder.mkdir(parents=True, exist_ok=True)

        safe_name = secure_filename(generated_docx_path.name)
        if not safe_name:
            safe_name = f"bid_{bidding_id}.docx"
        target = gen_folder / safe_name
        if generated_docx_path.resolve() != target.resolve():
            shutil.copy2(str(generated_docx_path), str(target))

        file_url = f"http://localhost:{os.environ.get('PORT', '3012')}/api/outputs/{safe_name}"

        conn = get_db()
        cur = conn.cursor()
        cur.execute('UPDATE bidding SET bid_document=?, status=? WHERE id=?',
                    (str(target), 'Generated', bidding['id']))
        conn.commit()
        conn.close()

        return jsonify({
            'message': '投标文件已生成（标准Word格式）',
            'fileUrl': file_url
        }), 201

    except Exception as e:
        logging.exception(f"生成投标书过程出错: {e}")
        return jsonify({'error': f'生成投标书失败: {str(e)}'}), 500
