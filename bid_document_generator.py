"""
投标文件Word文档生成器
生成标准格式的投标响应文件，包含封面、目录、各章节内容
按照招标文件格式要求：
- 标题一：一、二、三、（中文数字，黑体三号/16pt）
- 标题二：1. 2. 3.（黑体四号/14pt）
- 标题三：1.1 1.2 1.3（黑体小四/12pt）
- 标题四：1.1.1 1.1.2（楷体小四/12pt加粗）
- 正文：仿宋_GB2312 四号/12pt，首行缩进2字符
"""
import os
import logging
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ===== 中文数字转换 =====
CN_NUMBERS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
              '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十',
              '二十一', '二十二', '二十三', '二十四', '二十五', '二十六', '二十七', '二十八', '二十九', '三十']


def to_cn_number(n):
    """阿拉伯数字转中文"""
    if n < len(CN_NUMBERS):
        return CN_NUMBERS[n]
    return str(n)


# ===== 多级编号管理器 =====
class HeadingNumbering:
    """管理多级标题自动编号"""
    def __init__(self):
        self.counters = [0, 0, 0, 0]  # level1, level2, level3, level4
    
    def reset(self):
        self.counters = [0, 0, 0, 0]
    
    def next_level1(self, title):
        """标题一：一、标题"""
        self.counters[0] += 1
        self.counters[1] = 0
        self.counters[2] = 0
        self.counters[3] = 0
        return f"{to_cn_number(self.counters[0])}、{title}"
    
    def next_level2(self, title):
        """标题二：1. 标题"""
        self.counters[1] += 1
        self.counters[2] = 0
        self.counters[3] = 0
        return f"{self.counters[1]}. {title}"
    
    def next_level3(self, title):
        """标题三：1.1 标题"""
        self.counters[2] += 1
        self.counters[3] = 0
        return f"{self.counters[1]}.{self.counters[2]} {title}"
    
    def next_level4(self, title):
        """标题四：1.1.1 标题"""
        self.counters[3] += 1
        return f"{self.counters[1]}.{self.counters[2]}.{self.counters[3]} {title}"
    
    def format(self, level, title):
        """根据级别自动编号"""
        if level == 1:
            return self.next_level1(title)
        elif level == 2:
            return self.next_level2(title)
        elif level == 3:
            return self.next_level3(title)
        elif level == 4:
            return self.next_level4(title)
        return title


# ===== 样式工具函数 =====

def set_cell_border(cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            element = OxmlElement(f'w:{edge}')
            for key in ('sz', 'val', 'color', 'space'):
                if key in edge_data:
                    element.set(qn(f'w:{key}'), str(edge_data[key]))
            tcBorders.append(element)
    tcPr.append(tcBorders)


def add_page_break(doc):
    """添加分页符"""
    doc.add_page_break()


def set_run_font(run, font_name='仿宋_GB2312', font_size=12, bold=False):
    """设置文字格式"""
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)


def add_heading_styled(doc, text, level=1, font_name=None, font_size=None, 
                       numbered=False, numbering=None):
    """
    添加带样式的标题 — 按照标书排版规范
    level=1: 正文标题 — 微软雅黑 三号(16pt) 加粗
    level=2: 段落标题 — 宋体 小四号(12pt) 加粗
    level=3: 三级标题 — 宋体 小四号(12pt) 加粗
    level=4: 四级标题 — 宋体 小四号(12pt)
    """
    # 标书排版规范字体
    defaults = {
        1: ('微软雅黑', 16),   # 正文标题：三号加粗微软雅黑
        2: ('宋体', 12),       # 段落标题：小四号加粗宋体
        3: ('宋体', 12),       # 三级：小四号加粗宋体
        4: ('宋体', 12),       # 四级：小四号宋体
    }
    default_font, default_size = defaults.get(level, ('宋体', 12))
    font_name = font_name or default_font
    font_size = font_size or default_size
    
    # 自动编号
    if numbered and numbering:
        text = numbering.format(level, text)
    
    para = doc.add_paragraph()
    # 标题不缩进
    para.paragraph_format.first_line_indent = 0
    
    # level 1 居中
    if level == 1:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.space_before = Pt(12)
        para.space_after = Pt(12)
    else:
        para.space_before = Pt(6)
        para.space_after = Pt(4)
    
    run = para.add_run(text)
    is_bold = level <= 3  # 1-3级加粗，4级不加粗
    set_run_font(run, font_name=font_name, font_size=font_size, bold=is_bold)
    
    return para


def add_normal_text(doc, text, font_name='宋体', font_size=12, 
                    indent=True, bold=False, align=None):
    """添加正文段落 — 按照标书排版规范
    - 字体：宋体 小四号(12pt)
    - 首行缩进: 2字符
    - 行间距: 1.5倍
    """
    from docx.enum.text import WD_LINE_SPACING
    para = doc.add_paragraph()
    if indent:
        para.paragraph_format.first_line_indent = Pt(24)  # 2字符缩进（12pt×2）
    if align:
        para.alignment = align
    # 1.5倍行距
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = para.add_run(text)
    set_run_font(run, font_name=font_name, font_size=font_size, bold=bold)
    return para


def add_signature_block(doc, project_info, include_legal_rep=False):
    """添加签名盖章区域（右对齐）"""
    doc.add_paragraph()
    supplier = project_info.get('supplier_name', '{{供应商名称}}')
    date = project_info.get('date', '{{日期}}')
    
    add_normal_text(doc, f'供应商名称（加盖公章）：{supplier}', 
                    indent=False, align=WD_ALIGN_PARAGRAPH.RIGHT)
    if include_legal_rep:
        add_normal_text(doc, '法定代表人或授权代表（签字或签章）：______________', 
                        indent=False, align=WD_ALIGN_PARAGRAPH.RIGHT)
    add_normal_text(doc, f'日期：{date}', 
                    indent=False, align=WD_ALIGN_PARAGRAPH.RIGHT)


def add_table_with_data(doc, headers, rows, col_widths=None):
    """添加表格"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 设置表格边框
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    borders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        borders.append(border)
    tblPr.append(borders)
    
    # 表头
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                set_run_font(run, font_name='黑体', font_size=10, bold=True)
    # 数据行
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = str(cell_text)
            for para in cell.paragraphs:
                for run in para.runs:
                    set_run_font(run, font_name='仿宋_GB2312', font_size=10)
    # 设置列宽
    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(width)
    return table


# ===== 章节生成函数 =====

def generate_cover_page(doc, project_info):
    """生成封面 — 严格按照投标文件格式"""
    from docx.enum.text import WD_LINE_SPACING
    
    # 空行
    for _ in range(3):
        doc.add_paragraph()
    
    # 主标题：响 应 文 件（48pt加粗，3倍行距）
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    para.paragraph_format.line_spacing = 3.0
    run = para.add_run('响 应 文 件')
    set_run_font(run, font_name='方正小标宋简体', font_size=48, bold=True)
    
    # 副标题：（商务技术文件）（36pt加粗，3倍行距）
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    para.paragraph_format.line_spacing = 3.0
    run = para.add_run('（商务技术文件）')
    set_run_font(run, font_name='黑体', font_size=36, bold=True)
    
    for _ in range(4):
        doc.add_paragraph()
    
    # 项目信息（14pt加粗，两端对齐）
    info_lines = [
        f"项目名称：{project_info.get('project_name', '{{项目名称}}')}",
        '',
        f"项目编号/包号：{project_info.get('project_id', '{{项目编号/包号}}')}",
    ]
    for line in info_lines:
        if not line:
            doc.add_paragraph()
            continue
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = para.add_run(line)
        set_run_font(run, font_name='仿宋_GB2312', font_size=14, bold=True)
    
    for _ in range(5):
        doc.add_paragraph()
    
    # 供应商信息（14pt加粗）
    supplier_lines = [
        f"供应商名称：{project_info.get('supplier_name', '{{供应商名称}}')}",
        f"{project_info.get('date', '{{日期}}')}",
    ]
    for line in supplier_lines:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_before = Pt(8)
        para.paragraph_format.space_after = Pt(8)
        run = para.add_run(line)
        set_run_font(run, font_name='仿宋_GB2312', font_size=14, bold=True)
    
    add_page_break(doc)


# ===== 动态章节渲染引擎 =====

AI_CONTENT_KEYWORDS = {
    '项目管理': 'implementation',
    '实施方案': 'implementation',
    '培训方案': 'training',
    '售后服务': 'after_sales',
    '售后方案': 'after_sales',
}


def _detect_special_chapter(title, ai_contents):
    """检测章节是否有AI预生成内容，返回内容字符串或None"""
    if not ai_contents:
        return None
    for keyword, key in AI_CONTENT_KEYWORDS.items():
        if keyword in title:
            content = ai_contents.get(key)
            if content:
                return content
    return None


def generate_toc_dynamic(doc, chapter_design):
    """从 chapter_design 动态生成目录页"""
    add_heading_styled(doc, '目  录', level=1)

    for i, chapter in enumerate(chapter_design, 1):
        cn_num = to_cn_number(i)
        chapter_title = chapter.get('title', '') if isinstance(chapter, dict) else chapter
        para = doc.add_paragraph()
        run = para.add_run(f"{cn_num}、{chapter_title}")
        set_run_font(run, font_name='仿宋_GB2312', font_size=12)

        sections = chapter.get('sections', []) if isinstance(chapter, dict) else []
        for j, section in enumerate(sections, 1):
            section_title = section.get('title', '') if isinstance(section, dict) else section
            if not section_title:
                continue
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Cm(1)
            run = para.add_run(f"{j}. {section_title}")
            set_run_font(run, font_name='仿宋_GB2312', font_size=12)

    add_page_break(doc)


def _render_ai_content_block(doc, content):
    """渲染AI生成的文本内容（自动识别标题层级）"""
    if not content:
        add_normal_text(doc, '（内容待生成）')
        return
    lines = content.strip().split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith('# '):
            add_heading_styled(doc, line[2:], level=2)
        elif line.startswith('## '):
            add_heading_styled(doc, line[3:], level=3)
        elif line.startswith('### '):
            add_normal_text(doc, line[4:], bold=True)
        elif line.startswith('- ') or line.startswith('• '):
            add_normal_text(doc, line)
        elif line.startswith('|') and '|' in line[1:]:
            add_normal_text(doc, line, font_size=10)
        else:
            add_normal_text(doc, line)


def _render_equipment_table(doc, equipment_list, project_info, numbering):
    """渲染分项报价表（使用设备清单）"""
    project_id = project_info.get('project_id', '{{项目编号}}')
    project_name = project_info.get('project_name', '{{项目名称}}')

    add_normal_text(doc, f'项目编号/包号：{project_id}', indent=False)
    add_normal_text(doc, f'项目名称：{project_name}', indent=False)
    doc.add_paragraph()
    headers = ['序号', '设备/服务名称', '品牌/型号', '数量', '单位', '单价（元）', '合计（元）']
    rows = []
    if equipment_list:
        for i, equip in enumerate(equipment_list, 1):
            rows.append([
                str(i),
                equip.get('name', '{{名称}}'),
                equip.get('brand_model', '{{品牌/型号}}'),
                equip.get('quantity', '{{数量}}'),
                equip.get('unit', '{{单位}}'),
                equip.get('unit_price', '{{单价}}'),
                equip.get('total_price', '{{合计}}'),
            ])
    else:
        rows.append(['1', '详见招标文件采购清单', '详见招标文件要求', '详见招标文件', '项', '{{单价}}', '{{合计}}'])
    rows.append(['', '合计', '', '', '', '', '{{总价}}'])
    add_table_with_data(doc, headers, rows, col_widths=[1.2, 4, 3, 1.5, 1.2, 2.5, 2.5])
    doc.add_paragraph()
    supplier = project_info.get('supplier_name', '{{供应商名称}}')
    add_normal_text(doc, f'供应商名称（加盖公章）：{supplier}', indent=False)
    add_normal_text(doc, f'日期：{project_info.get("date", "{{日期}}")}', indent=False)
    add_page_break(doc)


def _render_markdown_table(doc, markdown_content):
    """将markdown表格渲染为Word表格"""
    lines = markdown_content.strip().split('\n')
    table_lines = [l for l in lines if l.startswith('|')]
    if not table_lines:
        add_normal_text(doc, markdown_content)
        return
    headers = [h.strip() for h in table_lines[0].split('|')[1:-1]]
    data_lines = [l for l in table_lines[2:] if not l.replace('|', '').replace('-', '').replace(':', '').strip()]
    rows = []
    for line in data_lines:
        cells = [c.strip() for c in line.split('|')[1:-1]]
        if cells:
            rows.append(cells)
    if headers:
        add_table_with_data(doc, headers, rows)
        doc.add_paragraph()


def _render_chapter_dynamic(doc, chapter, project_info, numbering,
                            equipment_list=None, ai_contents=None):
    """动态渲染单个章节"""
    title = chapter.get('title', '')
    chapter_type = chapter.get('type', 'normal')

    # AI预生成内容章节
    ai_content = _detect_special_chapter(title, ai_contents)
    if ai_content is not None:
        add_heading_styled(doc, title, level=1, numbered=True, numbering=numbering)
        _render_ai_content_block(doc, ai_content)
        add_page_break(doc)
        return

    add_heading_styled(doc, title, level=1, numbered=True, numbering=numbering)

    # 分项报价表特殊处理
    if '分项报价' in title:
        _render_equipment_table(doc, equipment_list, project_info, numbering)
        return

    # 报价一览表
    if '报价一览' in title:
        project_name = project_info.get('project_name', '{{项目名称}}')
        project_id = project_info.get('project_id', '{{项目编号}}')
        add_normal_text(doc, f'项目编号/包号：{project_id}', indent=False)
        add_normal_text(doc, f'项目名称：{project_name}', indent=False)
        doc.add_paragraph()
        headers = ['包号', '品目号', '标的名称', '报价（人民币元）']
        rows = [['1', '1-1', project_name, '{{总报价}}']]
        add_table_with_data(doc, headers, rows, col_widths=[2, 2, 8, 4])
        doc.add_paragraph()
        supplier = project_info.get('supplier_name', '{{供应商名称}}')
        add_normal_text(doc, f'供应商名称（加盖公章）：{supplier}', indent=False)
        add_normal_text(doc, '法定代表人或授权代表（签字或签章）：______________', indent=False)
        add_normal_text(doc, f'日期：{project_info.get("date", "{{日期}}")}', indent=False)
        add_page_break(doc)
        return

    # 产品彩页章节
    if '彩页' in title or '产品资料' in title:
        if equipment_list:
            for equip in equipment_list:
                add_heading_styled(doc, equip.get('name', '产品'), level=2, numbered=True, numbering=numbering)
                add_normal_text(doc, '（此处粘贴产品彩页资料并加盖公章）', align=WD_ALIGN_PARAGRAPH.CENTER)
                doc.add_paragraph()
        else:
            add_normal_text(doc, '（此处粘贴产品彩页资料并加盖公章）', align=WD_ALIGN_PARAGRAPH.CENTER)
        add_page_break(doc)
        return

    # 表格类型章节
    content = chapter.get('content', '')
    if chapter_type == 'table' or content:
        _render_markdown_table(doc, content)
        add_page_break(doc)
        return

    # 普通文本章节 — 渲染 sections/subsections
    sections = chapter.get('sections', [])
    for section in sections:
        section_title = section.get('title', '') if isinstance(section, dict) else section
        section_describe = section.get('describe', '') if isinstance(section, dict) else ''

        if section_title:
            add_heading_styled(doc, section_title, level=2, numbered=True, numbering=numbering)
        if section_describe:
            add_normal_text(doc, section_describe)

        subsections = section.get('subsections', []) if isinstance(section, dict) else []
        for sub in subsections:
            sub_title = sub.get('title', '') if isinstance(sub, dict) else sub
            sub_describe = sub.get('describe', '') if isinstance(sub, dict) else ''
            if sub_title:
                add_heading_styled(doc, sub_title, level=3, numbered=True, numbering=numbering)
            if sub_describe:
                add_normal_text(doc, sub_describe)

    # 需要签名区的章节
    signature_keywords = ['响应书', '授权委托', '资格声明', '偏离表', '报价表']
    if any(kw in title for kw in signature_keywords):
        supplier = project_info.get('supplier_name', '{{供应商名称}}')
        add_normal_text(doc, f'供应商名称（加盖公章）：{supplier}', indent=False)
        add_normal_text(doc, f'日期：{project_info.get("date", "{{日期}}")}', indent=False)

    add_page_break(doc)


# ===== 主生成函数 =====

def generate_bid_document_word(project_info, chapter_design=None,
                               equipment_list=None,
                               ai_implementation=None, ai_training=None, ai_after_sales=None,
                               output_path=None):
    """
    生成完整的投标响应文件Word文档（动态章节模式）

    Args:
        project_info: dict, 项目基本信息
        chapter_design: list, 章节设计（来自AI），为None时使用兼容默认模板
        equipment_list: list, 设备清单（来自AI提取）
        ai_implementation: str, AI生成的实施方案
        ai_training: str, AI生成的培训方案
        ai_after_sales: str, AI生成的售后方案
        output_path: str, 输出文件路径

    Returns:
        str: 生成的文件路径
    """
    doc = Document()

    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.6)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.header_distance = Cm(2.0)
        section.footer_distance = Cm(1.5)

    style = doc.styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(12)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    numbering = HeadingNumbering()

    ai_contents = {}
    if ai_implementation:
        ai_contents['implementation'] = ai_implementation
    if ai_training:
        ai_contents['training'] = ai_training
    if ai_after_sales:
        ai_contents['after_sales'] = ai_after_sales

    logging.info("开始生成投标文件Word文档...")

    generate_cover_page(doc, project_info)

    if chapter_design and isinstance(chapter_design, list) and len(chapter_design) > 0:
        # ===== 动态章节模式 =====
        generate_toc_dynamic(doc, chapter_design)
        for chapter in chapter_design:
            _render_chapter_dynamic(
                doc, chapter, project_info, numbering,
                equipment_list=equipment_list,
                ai_contents=ai_contents,
            )
    else:
        # ===== 兼容模式：默认18章模板 =====
        default_chapters = [
            {'title': '满足《中华人民共和国政府采购法》第二十二条规定', 'sections': [{'title': '营业执照', 'subsections': []}, {'title': '供应商资格声明书', 'subsections': []}]},
            {'title': '落实政府采购政策需满足的资格要求', 'sections': [{'title': '中小企业政策证明文件', 'subsections': []}]},
            {'title': '本项目的特定资格要求', 'sections': [{'title': '联合协议', 'subsections': []}, {'title': '其他特定资格要求', 'subsections': []}]},
            {'title': '磋商保证金凭证/交款单据复印件'},
            {'title': '响应书（实质性格式）'},
            {'title': '授权委托书（实质性格式）'},
            {'title': '报价一览表'},
            {'title': '分项报价表'},
            {'title': '合同条款偏离表（实质性格式）'},
            {'title': '采购需求偏离表（实质性格式）'},
            {'title': '本国产品标准证明文件'},
            {'title': '招标文件要求提供或投标人认为应附的其他材料'},
            {'title': '竞争性磋商文件要求提供或供应商认为应附的其他材料', 'sections': [{'title': '评审标准中所述业绩一览表', 'subsections': []}]},
            {'title': '项目管理及实施方案'},
            {'title': '培训方案'},
            {'title': '售后服务方案'},
            {'title': '供应商信息采集表'},
            {'title': '生产厂商提供的产品彩页资料'},
        ]
        generate_toc_dynamic(doc, default_chapters)
        for chapter in default_chapters:
            _render_chapter_dynamic(
                doc, chapter, project_info, numbering,
                equipment_list=equipment_list,
                ai_contents=ai_contents,
            )

    if not output_path:
        output_dir = Path('outputs')
        output_dir.mkdir(parents=True, exist_ok=True)
        safe_name = project_info.get('project_name', '投标文件').replace('/', '_').replace('\\', '_')
        output_path = str(output_dir / f'{safe_name}_响应文件.docx')

    doc.save(output_path)
    logging.info(f"投标文件已生成：{output_path}")
    return output_path


# ===== 测试入口 =====
if __name__ == '__main__':
    test_info = {
        'project_name': '2026年北京小汤山医院更新B区报告厅音响项目',
        'project_id': '0701-264106080026/01',
        'purchaser': '北京小汤山医院',
        'agency': '中技国际招标有限公司',
        'supplier_name': '{{供应商名称}}',
        'date': '2026年__月__日',
    }

    output = generate_bid_document_word(test_info)
    print(f'测试文件已生成: {output}')
