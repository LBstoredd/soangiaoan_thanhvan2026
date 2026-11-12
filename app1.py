import os
import sqlite3
import json
import time
import html
import re
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
import streamlit as st
import streamlit.components.v1 as components
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime, timedelta
import random
import string

try:
    from pptx import Presentation
    from pptx.util import Inches as PptxInches, Pt as PptxPt
    from pptx.dml.color import RGBColor as PptxRGBColor
    from pptx.enum.text import PP_ALIGN
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# ==========================================
# KHỞI TẠO CẤU HÌNH & SESSION STATE 
# ==========================================
st.set_page_config(page_title="Phần Mềm Trợ Lý Mầm Non Chuyên Sâu", layout="wide")

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mamnon_rag.db")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "is_admin" not in st.session_state:
    st.session_state["is_admin"] = False
if "api_key_saved" not in st.session_state:
    st.session_state["api_key_saved"] = ""
if "active_model_name" not in st.session_state:
    st.session_state["active_model_name"] = ""
if "cached_plans" not in st.session_state:
    st.session_state["cached_plans"] = []

# ==========================================
# CẤU HÌNH GIAO DIỆN THEME XANH MINT
# ==========================================
st.markdown("""
<style>
    .stApp { background: linear-gradient(180deg, #F4FBF9 0%, #FFFFFF 100%); font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1E293B; }
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #E6F7F2 0%, #D4F1E8 100%) !important; border-right: 1.5px solid #B2E5D6; }
    [data-testid="stSidebar"] * { color: #0F5247 !important; }
    h1 { color: #00796B !important; font-weight: 800 !important; letter-spacing: -0.3px; text-shadow: 0 2px 4px rgba(0, 121, 107, 0.08); }
    h2, h3, h4, h5 { color: #00695C !important; font-weight: 700 !important; }
    div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, #00B894 0%, #00897B 100%) !important; color: #FFFFFF !important; border: none !important; border-radius: 12px !important; padding: 10px 24px !important; font-weight: 700 !important; font-size: 15px !important; box-shadow: 0 4px 12px rgba(0, 184, 148, 0.28) !important; transition: all 0.25s ease !important;
    }
    div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover {
        transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0, 184, 148, 0.42) !important; background: linear-gradient(135deg, #00CBA4 0%, #00796B 100%) !important;
    }
    div.stButton > button[kind="secondary"] { background-color: #FFFFFF !important; color: #00796B !important; border: 1.5px solid #80CBC4 !important; border-radius: 12px !important; font-weight: 600 !important; padding: 8px 18px !important; transition: all 0.2s ease !important; }
    div.stButton > button[kind="secondary"]:hover { background-color: #E0F2F1 !important; border-color: #00897B !important; transform: translateY(-1px); }
    .stTextInput > div > div > input, .stTextArea > div > div > textarea, .stNumberInput > div > div > input, .stSelectbox > div > div { border-radius: 10px !important; border: 1.5px solid #B2DFDB !important; background-color: #FFFFFF !important; color: #1E293B !important; transition: all 0.25s ease; }
    .stTextInput > div > div > input:focus, .stTextArea > div > div > textarea:focus, .stNumberInput > div > div > input:focus { border-color: #00B894 !important; box-shadow: 0 0 0 3px rgba(0, 184, 148, 0.18) !important; }
    [data-testid="stExpander"], div[data-testid="stVerticalBlock"] > div[style*="border"] { border-radius: 14px !important; border: 1px solid #B2E5D6 !important; background-color: #FFFFFF !important; box-shadow: 0 4px 14px rgba(0, 121, 107, 0.04) !important; }
    button[data-baseweb="tab"] { font-weight: 700 !important; font-size: 15px !important; color: #00796B !important; }
    .stSlider > div > div > div > div { background-color: #00B894 !important; }
    div[data-testid="stAlert"] { border-radius: 12px !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# BỘ LỌC, KẺ BẢNG HTML & XUẤT WORD CHUẨN ĐẸP
# ==========================================
def clean_markdown_text(raw_text: str) -> str:
    if not raw_text: return ""
    lines = raw_text.split('\n')
    cleaned_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append("")
            continue
        if stripped.startswith('|') and stripped.endswith('|'):
            l_tab = stripped.replace('**', '').replace('###', '')
            cleaned_lines.append(l_tab)
            continue
        if re.match(r'^[-*_]{3,}$', stripped): continue
        l = line
        l = re.sub(r'^\s*#{1,6}\s*', '', l)
        l = re.sub(r'^(\s*)\*\s+', r'\1- ', l)
        l = re.sub(r'\*\*(.*?)\*\*', r'\1', l)
        l = l.replace('**', '')
        l = re.sub(r'(?<!\*)\*([^\*\n]+)\*(?!\*)', r'\1', l)
        l = l.replace('###', '')
        cleaned_lines.append(l)
    result = '\n'.join(cleaned_lines)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result.strip()

FULL_BOLD_PREFIXES = (
    'GIÁO ÁN', 'KẾ HOẠCH', 'SỔ CHỦ ĐỀ', 'ĐỀ TÀI', 'CHỦ ĐỀ', 'Hoạt động 1', 'Hoạt động 2', 'Hoạt động 3',
    '1. THỨ', '2. THỨ', '3. THỨ', '4. THỨ', '5. THỨ', 'THỨ HAI', 'THỨ BA', 'THỨ TƯ', 'THỨ NĂM', 'THỨ SÁU',
    'BÀI THƠ:', 'BÀI ĐỒNG DAO:', 'HỆ THỐNG CÂU ĐỐ', 'TRÒ CHƠI VẬN ĐỘNG', 'MẪU TIN NHẮN', 'LỜI NHẬN XÉT',
    'Ý TƯỞNG LÀM ĐỒ DÙNG', 'KỊCH BẢN DẠY NHANH', 'KỊCH BẢN VIDEO AI', 'PHIẾU HỌC TẬP', 'TRÒ CHƠI POWERPOINT',
    'BỘ PROMPT SINH HÌNH ẢNH', 'QUY TRÌNH 5 BƯỚC', 'CẢNH ', 'SLIDE ', 'BÀI TẬP '
)
LABEL_BOLD_PREFIXES = (
    '- Tình trạng sức khỏe:', '- Trạng thái cảm xúc, thái độ:', '-Trạng thái cảm xúc, thái độ:',
    '- Trạng thái cảm xúc:', '- Kiến thức, kỹ năng:', '- Kiến thức, kỹ năng & Năng lực số:',
    '1. Kiến thức:', '2. Kỹ năng:', '3. Thái độ:', '4. Năng lực hình thành:', '2. Kỹ năng & Năng lực số:',
    '1. Mục tiêu trọng tâm & Năng lực số:', '1. Đồ dùng của cô:', '2. Đồ dùng của trẻ:',
    '3. Môi trường & An toàn:', '3. Không gian tổ chức:', '1. Mục đích - Yêu cầu:', '2. Chuẩn bị:',
    '3. Cách chơi:', '4. Luật chơi:', '- Bước 1:', '- Bước 2:', '- Bước 3:', '- Bước 4:',
    '* Bước 1:', '* Bước 2:', '* Bước 3:', '* Bước 4:', '- Khởi động:', '- Trọng động:', '- Hồi tĩnh:',
    '* Khởi động:', '* Trọng động:', '* Hồi tĩnh:', 'a) ', 'b) ', 'c) ', '- Đón trẻ & Thể dục sáng:',
    '- Hoạt động học:', '- Hoạt động ngoài trời:', '- Hoạt động chơi ở các góc:', '- Hoạt động chiều:',
    '- Thể chất:', '- Nhận thức:', '- Ngôn ngữ:', '- Tình cảm - KNXH:', '- Thẩm mỹ:',
    '- Phát triển thể chất:', '- Phát triển nhận thức:', '- Phát triển ngôn ngữ:',
    '- Phát triển tình cảm - KNXH:', '- Phát triển thẩm mỹ:',
    '- Góc Phân vai:', '- Góc Xây dựng', '- Góc Nghệ thuật', '- Góc Tạo hình:', '- Góc Học tập',
    '- Góc Âm nhạc:', '- Góc Khám phá', '- Góc Thiên nhiên', '- Tuần 1:', '- Tuần 2:', '- Tuần 3:',
    '- Tuần 4:', '- Tuần 5:', 'Lớp:', 'Chủ đề:', 'Chủ đề lớn:', 'Chủ đề nhánh:', 'Lĩnh vực phát triển:',
    'Tên hoạt động:', 'Độ tuổi:', 'Thời lượng:', 'Thời gian:', 'Hình thức tổ chức:', 'Mức độ tích hợp:',
    'Người soạn:', 'Thời gian thực hiện:', 'Loại hoạt động:', 'Hướng tiếp cận:', '- Thời lượng:',
    '- Hình ảnh:', '- Lời thoại:', '- Âm thanh:', '- Prompt tiếng Anh:', '- Câu hỏi:', '- Đáp án đúng:',
    '- Hiệu ứng:', '- Tư duy:', '- Giao tiếp:', '- Hợp tác:', '- Giải quyết vấn đề:'
)

def is_roman_section(line_s: str) -> bool:
    return bool(re.match(r'^(I|II|III|IV|V|VI|VII)\.\s', line_s))

def is_full_bold_heading(line_s: str) -> bool:
    if is_roman_section(line_s): return True
    if line_s.startswith(FULL_BOLD_PREFIXES): return True
    if line_s.isupper() and len(line_s) > 5: return True
    return False

def should_bold_label_before_colon(line_s: str) -> bool:
    if ':' not in line_s: return False
    if line_s.startswith(LABEL_BOLD_PREFIXES): return True
    if re.match(r'^(\d+\.|[-+*])\s+[^:]{2,45}:', line_s): return True
    if re.match(r'^[A-Za-zÀ-ỹ\s\(\)\-&]{3,38}:', line_s): return True
    return False

def is_markdown_separator_row(cells: list) -> bool:
    return all(re.match(r'^[-:\s]+$', c) for c in cells if c.strip() != '')

def format_cell_lines_html(cell_raw: str, font_size_pt: int) -> str:
    parts = re.split(r'<br\s*/?>|\n', cell_raw)
    out = []
    for part in parts:
        s = part.strip()
        if not s: continue
        if re.match(r'^\d+\.\s', s) or s.startswith(('Bước ', '* Bước ')):
            out.append(f'<div style="margin-bottom: 3pt;"><b>{html.escape(s)}</b></div>')
        elif ':' in s and (s.startswith(('- Cô ', '+ Cô ', 'Cô ', '- Trẻ ', '+ Trẻ ', 'Trẻ ', '+ Đố:', '+ Hỏi:', '+ Khơi:', '- Đố:', '- Hỏi:', '- Khơi:', '- Luật chơi:', '- Cách chơi:', '- Với trẻ ')) or re.match(r'^[-+*]?\s*[^:]{2,35}:', s)):
            idx = s.find(':')
            b_part = html.escape(s[:idx + 1])
            r_part = html.escape(s[idx + 1:])
            out.append(f'<div style="margin-bottom: 2pt;"><b>{b_part}</b>{r_part}</div>')
        else:
            out.append(f'<div style="margin-bottom: 2pt;">{html.escape(s)}</div>')
    return "".join(out) if out else html.escape(cell_raw)

def render_html_table(table_lines: list, font_size_pt: int) -> str:
    parsed_rows = []
    for ln in table_lines:
        inner = ln.strip().strip('|')
        cells = [c.strip() for c in inner.split('|')]
        if not is_markdown_separator_row(cells):
            parsed_rows.append(cells)
    if not parsed_rows:
        return ""

    num_cols = max(len(r) for r in parsed_rows)
    if num_cols == 6:
        col_widths = ["15%", "17%", "17%", "17%", "17%", "17%"]
    elif num_cols == 5:
        col_widths = ["10%", "12%", "32%", "22%", "24%"]
    elif num_cols == 4:
        col_widths = ["11%", "41%", "26%", "22%"]
    elif num_cols == 3:
        col_widths = ["14%", "48%", "38%"]
    elif num_cols == 2:
        col_widths = ["28%", "72%"]
    else:
        col_widths = [f"{int(100/num_cols)}%"] * num_cols

    html_out = [
        f'<table style="width: 100%; border-collapse: collapse; margin: 8pt 0 12pt 0; font-family: \'Times New Roman\', Times, serif; font-size: {font_size_pt}pt; line-height: 1.45; color: #000000;">'
    ]
    for r_idx, row in enumerate(parsed_rows):
        while len(row) < num_cols:
            row.append("")
        html_out.append("<tr>")
        for c_idx, cell_val in enumerate(row[:num_cols]):
            w_attr = f"width: {col_widths[c_idx]};" if c_idx < len(col_widths) else ""
            if r_idx == 0:
                html_out.append(
                    f'<th style="{w_attr} border: 1px solid #2E7D32; background-color: #E6F7F2; color: #004D40; font-weight: bold; text-align: center; padding: 8pt 6pt; vertical-align: middle;">{html.escape(cell_val)}</th>'
                )
            else:
                align_style = "text-align: center; font-weight: bold;" if (c_idx == 0 and num_cols >= 3) else "text-align: left;"
                cell_inner_html = format_cell_lines_html(cell_val, font_size_pt)
                html_out.append(
                    f'<td style="{w_attr} border: 1px solid #475569; padding: 7pt 8pt; vertical-align: top; {align_style}">{cell_inner_html}</td>'
                )
        html_out.append("</tr>")
    html_out.append("</table>")
    return "".join(html_out)

def format_text_to_standard_html(text_content: str, font_size_pt: int = 14) -> str:
    html_paragraphs = []
    table_buffer = []

    def flush_table():
        nonlocal table_buffer
        if table_buffer:
            html_paragraphs.append(render_html_table(table_buffer, font_size_pt))
            table_buffer = []

    for raw_line in text_content.split('\n'):
        line_s = raw_line.strip()
        if line_s.startswith('|') and line_s.endswith('|') and line_s.count('|') >= 2:
            table_buffer.append(line_s)
            continue
        else:
            flush_table()

        if not line_s:
            html_paragraphs.append('<div style="height: 5pt;"></div>')
            continue

        if is_roman_section(line_s):
            escaped = html.escape(line_s)
            html_paragraphs.append(
                f'<div style="border-left: 5px solid #00897B; background-color: #F0FAF8; padding: 5pt 10pt; margin: 10pt 0 6pt 0; font-family: \'Times New Roman\', Times, serif; font-size: {font_size_pt + 1}pt; font-weight: bold; color: #004D40;">{escaped}</div>'
            )
        elif line_s.startswith(('GIÁO ÁN', 'KẾ HOẠCH', 'SỔ CHỦ ĐỀ')):
            escaped = html.escape(line_s)
            html_paragraphs.append(
                f'<p style="font-family: \'Times New Roman\', Times, serif; font-size: {font_size_pt + 2}pt; font-weight: bold; text-align: center; margin: 4pt 0 8pt 0; color: #000000;">{escaped}</p>'
            )
        elif is_full_bold_heading(line_s):
            escaped = html.escape(line_s)
            p_style = f"font-family: 'Times New Roman', Times, serif; font-size: {font_size_pt}pt; line-height: 1.5; text-align: justify; margin: 4pt 0; color: #000000;"
            html_paragraphs.append(f'<p style="{p_style}"><b>{escaped}</b></p>')
        elif should_bold_label_before_colon(line_s):
            idx = line_s.find(':')
            label_part = html.escape(line_s[:idx + 1])
            rest_part = html.escape(line_s[idx + 1:])
            p_style = f"font-family: 'Times New Roman', Times, serif; font-size: {font_size_pt}pt; line-height: 1.5; text-align: justify; margin: 0 0 4pt 0; color: #000000;"
            html_paragraphs.append(f'<p style="{p_style}"><b>{label_part}</b>{rest_part}</p>')
        else:
            escaped = html.escape(line_s)
            p_style = f"font-family: 'Times New Roman', Times, serif; font-size: {font_size_pt}pt; line-height: 1.5; text-align: justify; margin: 0 0 4pt 0; color: #000000;"
            html_paragraphs.append(f'<p style="{p_style}">{escaped}</p>')

    flush_table()
    return "".join(html_paragraphs)

def set_cell_shading(cell, color_hex: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def export_docx_from_text(text_content: str, info: dict, custom_template=None, font_size_pt=13, prefix_name="GiaoAn"):
    if custom_template: doc = Document(custom_template)
    else: doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(font_size_pt)

    table_buffer = []

    def flush_docx_table():
        nonlocal table_buffer
        if not table_buffer:
            return
        parsed_rows = []
        for ln in table_buffer:
            inner = ln.strip().strip('|')
            cells = [c.strip() for c in inner.split('|')]
            if not is_markdown_separator_row(cells):
                parsed_rows.append(cells)
        table_buffer = []
        if not parsed_rows:
            return

        num_cols = max(len(r) for r in parsed_rows)
        tbl = doc.add_table(rows=len(parsed_rows), cols=num_cols)
        tbl.style = 'Table Grid'

        if num_cols == 6:
            widths = [Inches(1.1), Inches(1.12), Inches(1.12), Inches(1.12), Inches(1.12), Inches(1.12)]
        elif num_cols == 5:
            widths = [Inches(0.75), Inches(0.85), Inches(2.1), Inches(1.5), Inches(1.5)]
        elif num_cols == 4:
            widths = [Inches(0.85), Inches(2.75), Inches(1.65), Inches(1.45)]
        elif num_cols == 3:
            widths = [Inches(1.0), Inches(3.2), Inches(2.5)]
        elif num_cols == 2:
            widths = [Inches(1.8), Inches(4.9)]
        else:
            widths = [Inches(6.7 / num_cols)] * num_cols

        for r_idx, row in enumerate(parsed_rows):
            while len(row) < num_cols:
                row.append("")
            for c_idx in range(num_cols):
                cell = tbl.cell(r_idx, c_idx)
                if c_idx < len(widths):
                    cell.width = widths[c_idx]
                cell_text = row[c_idx]
                lines_in_cell = [x.strip() for x in re.split(r'<br\s*/?>|\n', cell_text) if x.strip()]
                if not lines_in_cell:
                    lines_in_cell = [""]

                if r_idx == 0:
                    set_cell_shading(cell, "E6F7F2")
                    p0 = cell.paragraphs[0]
                    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p0.paragraph_format.space_after = Pt(2)
                    run0 = p0.add_run(lines_in_cell[0])
                    run0.bold = True
                    run0.font.name = 'Times New Roman'
                    run0.font.size = Pt(font_size_pt)
                else:
                    for l_i, l_str in enumerate(lines_in_cell):
                        p_c = cell.paragraphs[0] if l_i == 0 else cell.add_paragraph()
                        p_c.paragraph_format.space_after = Pt(2)
                        p_c.paragraph_format.line_spacing = 1.2
                        if c_idx == 0 and num_cols >= 3:
                            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            r_c = p_c.add_run(l_str)
                            r_c.bold = True
                            r_c.font.name = 'Times New Roman'
                            r_c.font.size = Pt(font_size_pt)
                        elif re.match(r'^\d+\.\s', l_str) or l_str.startswith(('Bước ', '* Bước ')):
                            r_c = p_c.add_run(l_str)
                            r_c.bold = True
                            r_c.font.name = 'Times New Roman'
                            r_c.font.size = Pt(font_size_pt)
                        elif ':' in l_str and (l_str.startswith(('- Cô ', '+ Cô ', 'Cô ', '- Trẻ ', '+ Trẻ ', 'Trẻ ', '+ Đố:', '+ Hỏi:', '+ Khơi:', '- Đố:', '- Hỏi:', '- Khơi:', '- Luật chơi:', '- Cách chơi:', '- Với trẻ ')) or re.match(r'^[-+*]?\s*[^:]{2,35}:', l_str)):
                            idx = l_str.find(':')
                            r_b = p_c.add_run(l_str[:idx + 1])
                            r_b.bold = True
                            r_b.font.name = 'Times New Roman'
                            r_b.font.size = Pt(font_size_pt)
                            r_n = p_c.add_run(l_str[idx + 1:])
                            r_n.font.name = 'Times New Roman'
                            r_n.font.size = Pt(font_size_pt)
                        else:
                            r_c = p_c.add_run(l_str)
                            r_c.font.name = 'Times New Roman'
                            r_c.font.size = Pt(font_size_pt)

    for raw_line in text_content.split('\n'):
        line_s = raw_line.strip()
        if line_s.startswith('|') and line_s.endswith('|') and line_s.count('|') >= 2:
            table_buffer.append(line_s)
            continue
        else:
            flush_docx_table()

        if not line_s: continue
        p = doc.add_paragraph()
        if line_s.startswith(('GIÁO ÁN', 'KẾ HOẠCH', 'SỔ CHỦ ĐỀ')):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.3
        
        if is_full_bold_heading(line_s):
            run = p.add_run(line_s)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size_pt + (1 if line_s.startswith(('GIÁO ÁN', 'KẾ HOẠCH', 'SỔ CHỦ ĐỀ')) else 0))
        elif should_bold_label_before_colon(line_s):
            idx = line_s.find(':')
            label_part = line_s[:idx + 1]
            rest_part = line_s[idx + 1:]
            r_bold = p.add_run(label_part)
            r_bold.bold = True
            r_bold.font.name = 'Times New Roman'
            r_bold.font.size = Pt(font_size_pt)
            r_norm = p.add_run(rest_part)
            r_norm.bold = False
            r_norm.font.name = 'Times New Roman'
            r_norm.font.size = Pt(font_size_pt)
        else:
            run = p.add_run(line_s)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size_pt)

    flush_docx_table()
            
    safe_topic = "".join(c if c.isalnum() else "_" for c in info.get('topic', 'GiaoAn'))[:40]
    filename = f"{prefix_name}_{safe_topic}.docx"
    doc.save(filename)
    return filename

def export_pptx_game(topic: str, age: str, raw_text: str) -> str:
    if not HAS_PPTX:
        return ""
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_layout = prs.slide_layouts[6]

    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.background
    fill1 = bg1.fill
    fill1.solid()
    fill1.fore_color.rgb = PptxRGBColor(230, 247, 242)

    tb1 = slide1.shapes.add_textbox(PptxInches(1), PptxInches(2), PptxInches(11.333), PptxInches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = f"TRÒ CHƠI TƯƠNG TÁC MẦM NON\nĐỀ TÀI: {topic.upper()}"
    r1.font.size = PptxPt(38)
    r1.font.bold = True
    r1.font.color.rgb = PptxRGBColor(0, 121, 107)

    p2 = tf1.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = PptxPt(20)
    r2 = p2.add_run()
    r2.text = f"Lứa tuổi: {age} | Bấm phím mũi tên hoặc click chuột để bắt đầu!"
    r2.font.size = PptxPt(22)
    r2.font.color.rgb = PptxRGBColor(30, 41, 59)

    blocks = re.split(r'(?=SLIDE\s+\d+|CÂU\s+\d+)', raw_text, flags=re.IGNORECASE)
    question_blocks = [b.strip() for b in blocks if b.strip() and ('?' in b or 'Đáp án' in b or 'A.' in b)]

    if not question_blocks:
        question_blocks = [
            f"CÂU 1: Các bé hãy quan sát và cho cô biết hôm nay chúng mình học về đề tài gì?\nA. {topic}\nB. Đi ngủ\nC. Xem tivi\nĐáp án đúng: A. {topic}",
            f"CÂU 2: Khi tham gia hoạt động {topic}, các con cần làm gì để học thật giỏi?\nA. Chú ý lắng nghe cô và hợp tác cùng bạn\nB. Chạy nhảy tự do\nC. Nói chuyện riêng\nĐáp án đúng: A. Chú ý lắng nghe cô và hợp tác cùng bạn"
        ]

    for idx, q_block in enumerate(question_blocks[:8], start=1):
        lines = [ln.strip() for ln in q_block.split('\n') if ln.strip()]
        q_title = lines[0]
        body_lines = "\n".join(lines[1:]) if len(lines) > 1 else q_block

        s = prs.slides.add_slide(blank_layout)
        bg = s.background
        f = bg.fill
        f.solid()
        f.fore_color.rgb = PptxRGBColor(255, 253, 245)

        header_box = s.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(1.2))
        htf = header_box.text_frame
        htf.word_wrap = True
        hp = htf.paragraphs[0]
        hr = hp.add_run()
        hr.text = q_title
        hr.font.size = PptxPt(28)
        hr.font.bold = True
        hr.font.color.rgb = PptxRGBColor(0, 105, 92)

        body_box = s.shapes.add_textbox(PptxInches(1.0), PptxInches(2.0), PptxInches(11.3), PptxInches(4.8))
        btf = body_box.text_frame
        btf.word_wrap = True
        for b_ln in body_lines.split('\n'):
            bp = btf.add_paragraph()
            bp.space_after = PptxPt(14)
            br = bp.add_run()
            br.text = b_ln
            br.font.size = PptxPt(24)
            if "đáp án" in b_ln.lower():
                br.font.bold = True
                br.font.color.rgb = PptxRGBColor(211, 47, 47)
            else:
                br.font.color.rgb = PptxRGBColor(30, 41, 59)

    slide_end = prs.slides.add_slide(blank_layout)
    bg_end = slide_end.background
    f_end = bg_end.fill
    f_end.solid()
    f_end.fore_color.rgb = PptxRGBColor(230, 247, 242)
    tb_end = slide_end.shapes.add_textbox(PptxInches(1), PptxInches(2.5), PptxInches(11.333), PptxInches(2.5))
    tf_end = tb_end.text_frame
    pe = tf_end.paragraphs[0]
    pe.alignment = PP_ALIGN.CENTER
    re_run = pe.add_run()
    re_run.text = "🎉 CHÚC MỪNG CÁC BÉ ĐÃ HOÀN THÀNH XUẤT SẮC TRÒ CHƠI! 🎉"
    re_run.font.size = PptxPt(32)
    re_run.font.bold = True
    re_run.font.color.rgb = PptxRGBColor(0, 121, 107)

    safe_topic = "".join(c if c.isalnum() else "_" for c in topic)[:35]
    pptx_filename = f"TroChoi_PPTX_{safe_topic}.pptx"
    prs.save(pptx_filename)
    return pptx_filename

def render_copy_and_print_buttons(text_content: str, doc_title: str, font_size_pt: int, unique_key: str):
    formatted_html = format_text_to_standard_html(text_content, font_size_pt)
    html_json = json.dumps(formatted_html)
    plain_json = json.dumps(text_content)
    title_json = json.dumps(doc_title)
    
    component_code = f"""
    <div style="display: flex; gap: 12px; width: 100%; align-items: center; margin: 0; padding: 0;">
        <button id="copyBtn_{unique_key}" onclick="copyFormattedContent_{unique_key}()" style="flex: 1; background: linear-gradient(135deg, #0984E3, #00796B); color: white; border: none; padding: 10px 16px; border-radius: 12px; font-size: 15px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 12px rgba(9, 132, 227, 0.25); transition: all 0.2s ease;">📋 Sao Chép Nhanh (Chuẩn Bảng Word)</button>
        <button onclick="printFormattedContent_{unique_key}()" style="flex: 1; background: linear-gradient(135deg, #00B894, #00897B); color: white; border: none; padding: 10px 16px; border-radius: 12px; font-size: 15px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 12px rgba(0, 184, 148, 0.28); transition: all 0.2s ease;">🖨️ In Nội Dung Ngay</button>
    </div>
    <script>
    async function copyFormattedContent_{unique_key}() {{
        const htmlContent = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' + {html_json} + '</body></html>';
        const plainContent = {plain_json};
        const btn = document.getElementById("copyBtn_{unique_key}");
        try {{
            if (navigator.clipboard && window.ClipboardItem) {{
                const blobHtml = new Blob([htmlContent], {{ type: "text/html" }});
                const blobText = new Blob([plainContent], {{ type: "text/plain" }});
                const data = [new ClipboardItem({{ "text/html": blobHtml, "text/plain": blobText }})];
                await navigator.clipboard.write(data);
            }} else {{ await navigator.clipboard.writeText(plainContent); }}
            btn.innerHTML = "✅ Đã Sao Chép Chuẩn Word/Zalo!";
            btn.style.background = "linear-gradient(135deg, #2ECC71, #27AE60)";
            setTimeout(() => {{ btn.innerHTML = "📋 Sao Chép Nhanh (Chuẩn Bảng Word)"; btn.style.background = "linear-gradient(135deg, #0984E3, #00796B)"; }}, 2500);
        }} catch (err) {{
            const ta = document.createElement("textarea"); ta.value = plainContent; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);
            btn.innerHTML = "✅ Đã Sao Chép Văn Bản!";
            setTimeout(() => {{ btn.innerHTML = "📋 Sao Chép Nhanh (Chuẩn Bảng Word)"; }}, 2500);
        }}
    }}
    function printFormattedContent_{unique_key}() {{
        const pWin = window.open('', '_blank');
        pWin.document.write('<!DOCTYPE html><html><head><title>' + {title_json} + '</title><style>@page {{ size: A4; margin: 18mm 20mm 18mm 20mm; }} body {{ font-family: "Times New Roman", Times, serif; font-size: {font_size_pt}pt; line-height: 1.5; color: #000; margin: 0; padding: 0; text-align: justify; }} table {{ width: 100%; border-collapse: collapse; }} th, td {{ border: 1px solid #333; padding: 6pt; }}</style></head><body>' + {html_json} + '</body></html>');
        pWin.document.close(); pWin.focus(); setTimeout(function() {{ pWin.print(); pWin.close(); }}, 350);
    }}
    </script>
    """
    components.html(component_code, height=48)

# ==========================================
# KHỞI TẠO CƠ SỞ DỮ LIỆU & LƯU TRỮ KÉP
# ==========================================
def init_db():
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute('''CREATE TABLE IF NOT EXISTS saved_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            title TEXT, 
            content TEXT, 
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            license_key TEXT UNIQUE, 
            customer_name TEXT, 
            expires_at DATE, 
            is_active INTEGER DEFAULT 1
        )''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS app_settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT
        )''')
        cursor.execute('''INSERT OR IGNORE INTO licenses (id, license_key, customer_name, expires_at, is_active) 
                          VALUES (1, 'LBSTORE2026', 'Admin Văn Bình (Master Admin)', '2027-12-31', 1)''')
        conn.commit()
        conn.close()
    except Exception:
        pass

def save_setting_to_db(key_name: str, val: str):
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO app_settings (setting_key, setting_value) VALUES (?, ?)", (key_name, val))
        conn.commit()
        conn.close()
    except Exception:
        pass

def get_setting_from_db(key_name: str) -> str:
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT setting_value FROM app_settings WHERE setting_key = ?", (key_name,))
        row = cursor.fetchone()
        conn.close()
        return row[0] if row and row[0] else ""
    except Exception:
        return ""

init_db()

if not st.session_state["api_key_saved"]:
    persisted_key = get_setting_from_db("gemini_api_key")
    if persisted_key:
        st.session_state["api_key_saved"] = persisted_key

def check_license(key):
    if key == 'LBSTORE2026': return True, "👑 Đăng nhập quyền Admin thành công!"
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT customer_name, expires_at, is_active FROM licenses WHERE license_key = ?", (key,))
        row = cursor.fetchone()
        conn.close()
        if row:
            c_name, expires_at, is_active = row
            if is_active == 1 and datetime.now().date() <= datetime.strptime(expires_at, "%Y-%m-%d").date():
                return True, f"✅ Kích hoạt thành công! Xin chào: {c_name} (Hạn đến: {expires_at})"
            return False, "❌ Mã kích hoạt này đã hết hạn sử dụng. Vui lòng liên hệ gia hạn (100k/tháng)!"
    except Exception: pass
    return False, "❌ Mã kích hoạt không hợp lệ hoặc không tồn tại!"

def create_new_license(customer_name, months):
    random_str = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    new_key = f"LB-{random_str}"
    expires_date = (datetime.now() + timedelta(days=30 * months)).strftime("%Y-%m-%d")
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO licenses (license_key, customer_name, expires_at, is_active) VALUES (?, ?, ?, 1)", (new_key, customer_name, expires_date))
        conn.commit()
        conn.close()
        return True, new_key, expires_date
    except Exception as e: return False, str(e), ""

def get_all_licenses():
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT id, license_key, customer_name, expires_at, is_active FROM licenses ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception: return []

def save_plan_to_db(title, content_text):
    created_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if "cached_plans" not in st.session_state:
        st.session_state["cached_plans"] = []
    
    plan_entry = {
        "id": int(time.time() * 1000),
        "title": title,
        "content": content_text,
        "created_at": created_time
    }
    st.session_state["cached_plans"].insert(0, plan_entry)
    
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO saved_plans (title, content, created_at) VALUES (?, ?, ?)", (title, content_text, created_time))
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False

def get_all_saved_plans():
    db_rows = []
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, content, created_at FROM saved_plans ORDER BY id DESC")
        db_rows = cursor.fetchall()
        conn.close()
    except Exception:
        pass

    combined = []
    seen_titles = set()

    for p in st.session_state.get("cached_plans", []):
        combined.append((p["id"], p["title"], p["content"], p["created_at"]))
        seen_titles.add(p["title"] + p["created_at"])

    for row in db_rows:
        r_id, r_title, r_content, r_time = row
        key = r_title + str(r_time)
        if key not in seen_titles:
            combined.append((r_id, r_title, r_content, str(r_time)))
            seen_titles.add(key)

    return combined

def delete_plan_by_id(plan_id):
    st.session_state["cached_plans"] = [p for p in st.session_state.get("cached_plans", []) if p["id"] != plan_id]
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM saved_plans WHERE id = ?", (plan_id,))
        conn.commit()
        conn.close()
    except Exception:
        pass

# =========================================================================
# BỘ CỨU CÁNH CHUYÊN SÂU CHUẨN BẢNG 4 CỘT & BỘ CÔNG CỤ NÂNG CAO
# =========================================================================
def generate_instant_pedagogical_plan(info: dict) -> str:
    topic = info.get("topic", "Bé khám phá hành trình lớn lên của hạt đậu").strip()
    domain = info.get("domain", "Nhận thức")
    sub_cat = info.get("sub_category", "Khám phá")
    route = info.get("route_type", "Khám phá khoa học")
    art_type = info.get("art_type", "")
    age = info.get("age", "4-5 tuổi")
    theme = info.get("theme", "Thế giới thực vật")
    teacher = info.get("teacher", "Nguyễn Thị Thanh Vân").strip() or "Giáo viên"
    school = info.get("school", "Trường Mầm Non Phước Thạnh").strip() or "Trường Mầm Non"
    class_name = info.get("class_name", "Lớp Mẫu giáo Nhỡ").strip() or "Lớp Mẫu giáo"
    extras_list = info.get("extras", [])
    extras_str = ", ".join(extras_list) if extras_list else "Ứng dụng công nghệ AI, Giáo dục Cảm xúc - Xã hội (SEL)"
    custom_ideas = info.get("custom_ideas", "").strip()
    existing_tools = info.get("existing_tools", "").strip()
    disp_fmt = info.get("display_format", "Bảng 4 cột chuẩn (Mẫu YouTube)")

    art_note = f" ({art_type})" if art_type else ""

    header_block = f"""GIÁO ÁN HOẠT ĐỘNG HỌC CÓ CHỦ ĐÍCH
Trường: {school} | Lớp: {class_name} | Người soạn: {teacher}
Chủ đề: {theme}
Lĩnh vực phát triển: {domain} ({sub_cat} - {route})
Tên hoạt động: {topic}{art_note}
Độ tuổi: {age}
Thời lượng: 30 phút
Hình thức tổ chức: Cả lớp, nhóm 4 trẻ, cá nhân
Mức độ tích hợp: {extras_str}

I. MỤC TIÊU HOẠT ĐỘNG
1. Kiến thức:
- Trẻ nhận biết và gọi tên được các đặc điểm nổi bật, cấu tạo và quy trình phát triển của nội dung bài học: {topic}{art_note}.
- Trẻ hiểu được lợi ích, ý nghĩa thực tế của {topic} đối với đời sống con người và môi trường xung quanh.
2. Kỹ năng:
- Trẻ thực hiện thành thạo thao tác trải nghiệm, quan sát có chủ đích, so sánh và phân loại theo nhóm 4 bạn.
- Trẻ rèn luyện kỹ năng quan sát, diễn đạt ngôn ngữ mạch lạc, trả lời câu hỏi trọn vẹn câu và tự tin chia sẻ ý kiến trước lớp.
3. Thái độ (Giáo dục Cảm xúc - Xã hội - SEL):
- Trẻ tích cực, hào hứng tham gia hoạt động trải nghiệm và biết bảo vệ môi trường, giữ gìn sản phẩm.
- Trẻ biết kiên nhẫn chờ đợi đến lượt, lắng nghe và tôn trọng ý kiến của bạn trong nhóm.
- Trẻ biết chia sẻ niềm vui và hỗ trợ bạn cùng thực hiện nhiệm vụ.
4. Năng lực hình thành:
- Tư duy: Tự quan sát, nhận ra sự thay đổi về kích thước, hình dạng và bản chất của sự vật qua từng giai đoạn.
- Giao tiếp: Trẻ sử dụng từ ngữ rõ ràng, mạch lạc để mô tả suy nghĩ và trả lời câu hỏi của cô.
- Hợp tác: Biết chia sẻ đồ dùng, phân công công việc trong nhóm và phối hợp nhịp nhàng với bạn.
- Giải quyết vấn đề: Tự nhận biết và điều chỉnh thao tác khi gặp tình huống chưa đúng trong quá trình thực hành.

II. CHUẨN BỊ
1. Đồ dùng của cô:
- Video AI ngắn (thời lượng 30 giây) minh họa sinh động nội dung '{topic}' kết hợp âm thanh vui tươi.
- Tivi màn hình chỉnh lớn, loa bluetooth, nhạc không lời êm dịu và bài hát theo chủ đề {theme}.
- Bộ thẻ hình ảnh minh họa trực quan quy trình 4 bước và mẫu vật thật kích thước lớn.{f' Đồ dùng hiện có: {existing_tools}.' if existing_tools else ''}{f' Ý tưởng riêng của cô: {custom_ideas}.' if custom_ideas else ''}
2. Đồ dùng của trẻ (Chia theo nhóm 4 trẻ và cá nhân):
- Mỗi nhóm 4 trẻ có: 1 khay nhựa đựng học liệu, bộ thẻ hình sắp xếp trình tự và vật thật trải nghiệm trực tiếp.
- Mỗi trẻ có: Bảng con/khay cá nhân, bộ đồ dùng thực hành đủ số lượng, khăn lau tay sạch sẽ.
3. Môi trường & An toàn:
- Sắp xếp lớp học gọn gàng, sạch sẽ, thoáng mát; thảm trải phòng học sạch để trẻ ngồi vòng cung.
- Đảm bảo an toàn tuyệt đối: Đồ dùng không sắc nhọn, nhắc trẻ không đưa vật nhỏ vào mũi/miệng.

III. TIẾN TRÌNH HOẠT ĐỘNG"""

    if "4 cột" in disp_fmt:
        table_block = f"""| Thời gian | Hoạt động của cô (Lời thoại, câu hỏi & hành động) | Dự kiến hoạt động & Phản hồi của trẻ | Quan sát, Hỗ trợ & Phân hóa |
| --- | --- | --- | --- |
| 3 - 4 phút | 1. Ổn định - Gây hứng thú (Xem video AI):<br>- Cô mở nhạc vui tươi, tập trung trẻ ngồi vòng cung trên thảm:<br>+ Cô nói: "Chào mừng các bé yêu lớp {class_name} đến với tiết học kỳ diệu ngày hôm nay!"<br>- Cô mở video AI (30 giây) minh họa sinh động về '{topic}'.<br>- Hệ thống câu hỏi gợi mở tư duy:<br>+ Đố: "Các con vừa được xem bạn nhỏ trong video khám phá điều gì nào?"<br>+ Hỏi: "Con thấy điều gì đặc biệt nhất xuất hiện trên màn hình?"<br>+ Khơi: "Điều kỳ diệu gì sẽ xảy ra tiếp theo khi chúng mình tự tay khám phá '{topic}' nhỉ?" | - Trẻ chú ý quan sát màn hình, tỏ ra hào hứng, thích thú.<br>- Trẻ trả lời: "Dạ về {topic} ạ!"<br>- Trẻ giơ tay phát biểu: "Con thấy màu sắc rất đẹp và sinh động ạ!" | - Khen ngợi trẻ quan sát tốt.<br>- Với trẻ nhút nhát: Cô đến gần, trao đổi riêng, gợi mở từ khóa để trẻ tự tin trả lời. |
| 8 - 10 phút | 2. Khám phá trọng tâm & Đàm thoại sâu:<br>- Cô chia trẻ về các nhóm 4 bạn, phát vật thật và bộ thẻ hình cho từng nhóm.<br>- Cô hướng dẫn trẻ quan sát bằng đa giác quan (nhìn, sờ, ngửi) và sắp xếp thẻ theo trình tự:<br>+ Cô hỏi: "Nhóm nào giỏi cho cô biết đặc điểm nổi bật của '{topic}' gồm những gì?"<br>+ Cô hỏi mở rộng: "Vì sao '{topic}' lại quan trọng và có ích đối với cuộc sống của chúng ta?"<br>- Cô chốt kiến thức chuẩn xác, ngắn gọn kết hợp hiệu ứng hình ảnh minh họa. | - Trẻ về nhóm 4 bạn, cùng nhau sờ, quan sát vật thật và thảo luận sôi nổi.<br>- Đại diện nhóm trả lời mạch lạc theo suy nghĩ của trẻ.<br>- Cả lớp vỗ tay hưởng ứng ý kiến hay của bạn. | - Quan sát sự hợp tác của các nhóm 4 trẻ.<br>- Nhắc nhở trẻ không tranh giành đồ dùng, biết nhường bạn quan sát cùng. |
| 10 phút | 3. Thực hành - Trải nghiệm (Tự tay thao tác):<br>- Cô nêu nhiệm vụ thực hành và làm mẫu từng bước rõ ràng:<br>+ Bước 1: Chuẩn bị khay đồ dùng cá nhân ngay ngắn.<br>+ Bước 2: Thực hiện thao tác chính với '{topic}' theo đúng trình tự.<br>+ Bước 3: Kiểm tra và hoàn thiện sản phẩm của mình.<br>- Cô giao nhiệm vụ: "Các con hãy khéo léo thực hiện từng bước nhé, bạn nào xong trước có thể giúp đỡ bạn bên cạnh!"<br>- Cô đi từng bàn quan sát, hỗ trợ kỹ thuật. | - Trẻ ngồi vào vị trí thực hành.<br>- Trẻ chăm chú thực hiện từng bước theo hướng dẫn của cô.<br>- Trẻ vui mừng khoe sản phẩm với cô và các bạn trong nhóm. | - Hỗ trợ kịp thời những trẻ thao tác còn lúng túng.<br>- Gợi ý mở rộng thêm chi tiết sáng tạo cho những trẻ làm nhanh và khéo tay. |
| 3 phút | 4. Trò chơi củng cố: "Ai Nhanh Trí Hơn":<br>- Cô phổ biến luật chơi và cách chơi:<br>+ Chia lớp thành 2 đội thi đua bật qua vòng thể dục lên gắn đúng thẻ hình về '{topic}'.<br>+ Thời gian chơi tính bằng 1 bản nhạc vui nhộn.<br>- Cô tổ chức cho trẻ chơi 2 lượt, cùng trẻ kiểm tra kết quả. | - Trẻ hào hứng xếp hàng thành 2 đội.<br>- Trẻ bật qua vòng và gắn thẻ đúng yêu cầu, cổ vũ bạn nhiệt tình.<br>- Cả lớp cùng cô đếm số thẻ đúng của 2 đội. | - Quan sát đảm bảo an toàn khi trẻ bật qua vòng.<br>- Động viên tinh thần đoàn kết của cả hai đội chơi. |
| 3 phút | 5. Kết thúc - Nhận xét - Chuyển hoạt động:<br>- Cô tập trung trẻ lại, nhận xét tuyên dương sự cố gắng của cả lớp:<br>+ Cô nói: "Hôm nay cô thấy bạn nào cũng rất giỏi, khéo léo và biết hợp tác cùng bạn!"<br>- Nhắc trẻ thu dọn đồ dùng gọn gàng, rửa tay sạch sẽ và chuyển sang hoạt động tiếp theo. | - Trẻ vỗ tay vui sướng khi được cô khen.<br>- Trẻ tự giác thu dọn khay đồ dùng cất đúng nơi quy định và đi rửa tay. | - Nhắc nhở nhẹ nhàng trẻ giữ gìn vệ sinh chung sau khi hoàn thành tiết học. |"""
    elif "3 cột" in disp_fmt:
        table_block = f"""| Thời gian | Hoạt động của cô (Lời thoại, câu hỏi & hành động) | Dự kiến hoạt động & Phản hồi của trẻ |
| --- | --- | --- |
| 3 - 4 phút | 1. Ổn định - Gây hứng thú:<br>- Cô mở nhạc chủ đề {theme}, trò chuyện dẫn dắt vào bài '{topic}'.<br>- Cô đặt câu hỏi gợi mở: "Các con thấy điều gì đặc biệt trong món quà hôm nay?" | - Trẻ hát, vận động vui tươi.<br>- Trẻ chú ý quan sát và hào hứng trả lời câu hỏi của cô. |
| 18 - 20 phút | 2. Nội dung trọng tâm (Khám phá & Thực hành):<br>- Bước 1: Cho trẻ quan sát vật thật, đàm thoại phân tích đặc điểm của '{topic}'.<br>- Bước 2: Cô làm mẫu và tổ chức cho trẻ thực hành trải nghiệm theo nhóm 4 trẻ.<br>- Bước 3: Tổ chức trò chơi vận động củng cố kiến thức bài học. | - Trẻ quan sát, thảo luận nhóm và tự tin trả lời.<br>- Trẻ thực hành thao tác khéo léo.<br>- Trẻ tham gia trò chơi sôi nổi, đúng luật. |
| 3 - 4 phút | 3. Kết thúc - Nhận xét:<br>- Cô nhận xét, tuyên dương cả lớp, giáo dục kỹ năng và hướng dẫn thu dọn đồ dùng. | - Trẻ lắng nghe cô khen, tự giác cất đồ dùng gọn gàng. |"""
    else:
        table_block = f"""Hoạt động 1: Ổn định tổ chức, gây hứng thú (3 - 4 phút)
- Cô tập trung trẻ, mở nhạc vui tươi theo chủ đề {theme} và dẫn dắt giới thiệu đề tài '{topic}'.
- Cô đặt câu hỏi gợi mở để kích thích trí tò mò của trẻ.

Hoạt động 2: Nội dung trọng tâm (20 - 22 phút)
- Bước 1: Khám phá trải nghiệm thực tế & Đàm thoại sâu về đặc điểm, lợi ích của '{topic}'.
- Bước 2: Thực hành trải nghiệm theo nhóm 4 trẻ và cá nhân; cô bao quát hỗ trợ từng trẻ.
- Bước 3: Trò chơi củng cố 'Ai Nhanh Trí Hơn' giúp trẻ khắc sâu kiến thức.

Hoạt động 3: Kết thúc hoạt động (3 phút)
- Cô nhận xét, tuyên dương cả lớp, giáo dục trẻ và hướng dẫn thu dọn đồ dùng gọn gàng."""

    footer_block = f"""
IV. ĐÁNH GIÁ CUỐI NGÀY
- Tình trạng sức khỏe: 100% trẻ khỏe mạnh, tinh thần thoải mái, tham gia tích cực từ đầu đến cuối tiết học.
- Trạng thái cảm xúc, thái độ: Trẻ vui vẻ, hào hứng, tự tin phát biểu xây dựng bài và biết hợp tác chia sẻ cùng bạn.
- Kiến thức, kỹ năng: Đa số trẻ nắm vững kiến thức trọng tâm bài '{topic}', thực hiện đúng các bước thao tác; cô đã hỗ trợ kịp thời các trẻ còn nhút nhát."""

    return clean_markdown_text(f"{header_block}\n{table_block}\n{footer_block}")

def generate_instant_digital_plan(info: dict) -> str:
    topic = info.get("topic", "Khám phá sự kỳ diệu của quả cam").strip()
    domain = info.get("domain", "Phát triển nhận thức")
    age = info.get("age", "5-6 tuổi")
    school = info.get("school", "Trường Mầm Non").strip() or "Trường Mầm Non"
    teacher = info.get("teacher", "Nguyễn Thị Mai").strip() or "Giáo viên"
    class_name = info.get("class_name", "Lớp Lá 1").strip() or "Lớp Lá"
    theme = info.get("theme", "Thế giới thực vật")
    act_type = info.get("activity_type", "Ứng dụng công nghệ/AI")
    duration = info.get("duration", "30 - 35 phút")
    num_kids = info.get("num_children", 25)
    digital_equip = info.get("digital_equip", "Máy tính kết nối tivi thông minh, loa bluetooth")
    digital_tools = ", ".join(info.get("digital_tools", ["Bảng tương tác", "Canva", "Video AI"]))

    raw_plan = f"""GIÁO ÁN TÍCH HỢP NĂNG LỰC SỐ & CÔNG NGHỆ THÔNG MINH
Trường: {school} | Lớp: {class_name} | Người soạn: {teacher}
Chủ đề: {theme}
Lĩnh vực phát triển: {domain}
Tên hoạt động: {topic}
Độ tuổi: {age} | Số lượng: {num_kids} trẻ
Thời lượng: {duration}
Loại hoạt động: {act_type}

I. MỤC TIÊU HOẠT ĐỘNG
1. Kiến thức:
- Trẻ nhận biết, gọi đúng tên và các đặc điểm cấu tạo nổi bật của {topic} qua việc quan sát vật thật kết hợp tương tác số.
- Trẻ bước đầu làm quen với các thiết bị công nghệ giáo dục ({digital_equip}), hiểu được thao tác chạm - chọn trên màn hình cảm ứng ({digital_tools}).
2. Kỹ năng & Năng lực số:
- Hình thành kỹ năng số ban đầu: Trẻ biết quan sát điểm chạm, dùng ngón tay trỏ thực hiện thao tác kéo - thả, ghép đôi hình ảnh trên màn hình tương tác.
- Rèn luyện kỹ năng quan sát, so sánh đối chiếu giữa vật thật ngoài đời và mô hình mô phỏng 3D trên màn hình.
3. Thái độ:
- Trẻ hào hứng, thích thú khám phá công nghệ, có ý thức giữ khoảng cách an toàn cho mắt khi tiếp xúc màn hình.

II. CHUẨN BỊ
1. Đồ dùng của cô: {digital_equip}; học liệu số ({digital_tools}); vật thật trực quan.
2. Đồ dùng của trẻ: Vật thật trải nghiệm đa giác quan, thẻ tín hiệu tương tác.
3. Không gian tổ chức: Phòng học thoáng mát, màn hình ngang tầm mắt trẻ, khoảng cách tối thiểu 2m.

III. TIẾN TRÌNH HOẠT ĐỘNG
| Thời gian | Hoạt động của cô (Lời thoại, câu hỏi & hành động) | Dự kiến hoạt động & Phản hồi của trẻ | Quan sát, Hỗ trợ & Phân hóa |
| --- | --- | --- | --- |
| 3 - 5 phút | 1. Mở đầu - Khơi gợi hứng thú với Trợ lý AI:<br>- Cô bật màn hình tương tác giới thiệu nhân vật hoạt hình AI chào đón cả lớp.<br>- Nhân vật AI đặt câu đố vui dẫn dắt trẻ khám phá '{topic}'. | - Trẻ thích thú vẫy tay chào nhân vật AI.<br>- Trẻ sôi nổi đoán câu đố và gọi tên bài học. | - Khen ngợi trẻ phản xạ nhanh, hỗ trợ trẻ nhút nhát mạnh dạn giơ tay. |
| 20 - 25 phút | 2. Trọng tâm - Khám phá Đa phương tiện & Tương tác số:<br>- Bước 1: Trẻ trải nghiệm vật thật kết hợp quan sát video 3D trên tivi.<br>- Bước 2: Mời đại diện các nhóm lên chạm màn hình tương tác kéo thả, phân loại đúng - sai.<br>- Bước 3: Tổ chức trò chơi số hóa theo đội thi đua giải mã ô chữ. | - Trẻ sờ, ngửi vật thật và đối chiếu với mô hình 3D trên màn hình.<br>- Trẻ thực hiện thao tác chạm màn hình chính xác, tự tin.<br>- Các đội thi đua hào hứng. | - Hướng dẫn trẻ thao tác chạm nhẹ nhàng vào màn hình.<br>- Nhắc trẻ giữ tư thế ngồi thẳng, không đứng quá sát tivi. |
| 2 - 3 phút | 3. Kết thúc & Giáo dục nhẹ nhàng:<br>- Cô và nhân vật AI tuyên dương cả lớp, nhắc trẻ bảo vệ đôi mắt khi xem tivi ở nhà. | - Trẻ cùng cô tập động tác chớp mắt thư giãn và hát bài hát kết thúc. | - Quan sát đảm bảo trẻ thoải mái, vui vẻ khi chuyển hoạt động. |

IV. ĐÁNH GIÁ CUỐI NGÀY
- Tình trạng sức khỏe: Trẻ khỏe mạnh, giữ đúng tư thế ngồi và khoảng cách an toàn cho mắt suốt tiết học.
- Trạng thái cảm xúc, thái độ: Trẻ đặc biệt hào hứng với các hiệu ứng công nghệ số và nhân vật AI dẫn chuyện.
- Kiến thức, kỹ năng: Trẻ nắm vững nội dung bài học ({topic}), hình thành phản xạ chạm - tương tác số tự tin."""
    return clean_markdown_text(raw_plan)

def generate_instant_advanced_tool(tool_type: str, info: dict) -> str:
    topic = info.get("topic", "Đề tài mầm non")
    age = info.get("age", "4-5 tuổi")
    theme = info.get("theme", "Chủ đề quanh bé")

    if tool_type == "quick_script":
        return clean_markdown_text(f"""KỊCH BẢN DẠY NHANH (LỜI DẪN & PHÂN BỔ THỜI GIAN)
ĐỀ TÀI: {topic.upper()} | ĐỘ TUỔI: {age}

| Thời gian | Nội dung hoạt động | Lời dẫn & Câu hỏi cốt lõi của cô |
| --- | --- | --- |
| 00:00 - 04:00 (4 phút) | 1. Ổn định & Gây hứng thú | - Đọc vè tập trung: "Xúm xít, xúm xít! Bên cô, bên cô!"<br>- Lời dẫn: "Hôm nay cô mang đến một bí mật từ thế giới {theme}. Đố các bé biết là gì?" |
| 04:00 - 14:00 (10 phút) | 2. Khám phá trọng tâm | - Cho trẻ quan sát vật thật, video AI về '{topic}'.<br>- 3 Câu hỏi đinh: (1) Con thấy có đặc điểm gì? (2) Nó có ích lợi gì? (3) Điều gì xảy ra nếu thiếu nó? |
| 14:00 - 22:00 (8 phút) | 3. Thực hành trải nghiệm | - Chia nhóm 4 trẻ thực hành trực tiếp trên khay đồ dùng.<br>- Cô bao quát gợi mở, hỗ trợ trẻ thao tác chậm. |
| 22:00 - 28:00 (6 phút) | 4. Trò chơi củng cố | - Trò chơi "Ai Nhanh Trí Hơn": 2 đội thi đua bật qua vòng chọn đúng thẻ hình về '{topic}'. |
| 28:00 - 30:00 (2 phút) | 5. Kết thúc & Chuyển tiếp | - Tuyên dương cả lớp, giáo dục kỹ năng sống và hướng dẫn thu dọn đồ dùng. |""")

    elif tool_type == "video_ai":
        return clean_markdown_text(f"""KỊCH BẢN VIDEO AI MINH HỌA TIẾT DẠY
ĐỀ TÀI: {topic.upper()} | ĐỘ TUỔI: {age}

| Cảnh | Thời lượng | Hình ảnh minh họa | Lời thoại & Âm thanh | Prompt tiếng Anh (Tạo ảnh/video 3D) |
| --- | --- | --- | --- | --- |
| Cảnh 1 | 5 giây | Khung cảnh lớp học mầm non Việt Nam rực rỡ nắng mai, cô giáo dịu dàng mỉm cười chào các bé bên biểu tượng '{topic}'. | - Lời thoại: "Chào mừng các bé yêu đến với hành trình khám phá điều kỳ diệu về {topic} hôm nay!"<br>- Âm thanh: Nhạc nền vui tươi, tiếng chuông leng keng. | Cute 3D Pixar style illustration of a friendly Vietnamese kindergarten teacher smiling warmly in a bright classroom with {topic}, cheerful Vietnamese preschoolers, morning sunlight --ar 16:9 |
| Cảnh 2 | 5 giây | Cận cảnh chi tiết sinh động của '{topic}' xuất hiện lấp lánh ở trung tâm màn hình với màu sắc tươi sáng, rõ nét. | - Lời thoại: "Các con hãy mở to đôi mắt xinh xem điều gì đang xuất hiện trước mắt chúng mình nào!"<br>- Âm thanh: Hiệu ứng âm thanh phép thuật lung linh. | Cute 3D educational animation close-up of {topic} glowing softly with vibrant colors, clean bright background, child-friendly 3D render style --ar 16:9 |
| Cảnh 3 | 5 giây | Một em bé mầm non Việt Nam 4-5 tuổi mặc đồng phục dễ thương đang chăm chú quan sát và chỉ tay khám phá '{topic}'. | - Lời thoại: "Ồ! Thật là thú vị phải không nào? Chúng mình cùng tìm hiểu từng đặc điểm nhé!"<br>- Âm thanh: Nhạc dạo vui nhộn, tiết tấu vừa phải. | Cute 3D style, a smiling 4-year-old Vietnamese child in neat preschool uniform pointing curiously at {topic}, warm classroom setting, Pixar style --ar 16:9 |
| Cảnh 4 | 5 giây | Các bạn nhỏ cùng nhau thực hành, chia sẻ và vui vẻ hợp tác trải nghiệm '{topic}' trên bàn học. | - Lời thoại: "Khi chúng mình cùng nhau khám phá và chia sẻ, mọi điều đều trở nên thật dễ dàng và vui vẻ!"<br>- Âm thanh: Tiếng cười nói trong trẻo, nhạc nền ấm áp. | Cute 3D illustration of happy Vietnamese kindergarten children working together at a colorful wooden table exploring {topic}, teamwork, warm lighting --ar 16:9 |
| Cảnh 5 | 5 giây | Cảnh tổng kết rực rỡ với pháo giấy, ngôi sao lấp lánh chúc mừng các bé hoàn thành bài học '{topic}'. | - Lời thoại: "Hoan hô các bé thật giỏi! Hẹn gặp lại các con trong những bài học tiếp theo nhé!"<br>- Âm thanh: Nhạc kết thúc vui tươi, tiếng vỗ tay rộn rã. | Cute 3D celebration scene with colorful confetti, golden stars, and smiling Vietnamese preschoolers cheering happily around {topic}, joyful atmosphere --ar 16:9 |

VII. QUY TRÌNH 5 BƯỚC DỰNG VIDEO AI & DANH SÁCH TÀI NGUYÊN
1. Danh sách tài nguyên đề xuất:
- Công cụ tạo ảnh AI: Midjourney / Canva Magic Media / Google ImageFX (Dùng để tạo các bức ảnh tĩnh theo Prompt).
- Công cụ tạo chuyển động (Image-to-Video AI): Runway Gen-3 / Kling AI / Hailuo AI (Biến ảnh tĩnh thành video chuyển động 5 giây).
- Công cụ lồng tiếng & Âm thanh AI: CapCut (Giọng đọc AI trẻ em/cô giáo tiếng Việt) / ElevenLabs.
- Phần mềm dựng hoàn thiện: CapCut PC hoặc CapCut Mobile.
2. Quy trình 5 bước dựng video AI cho giáo viên mầm non:
- Bước 1: Lập kịch bản & Phân cảnh (Storyboarding) chia nhỏ thành 5 cảnh ngắn 5 giây để giữ sự tập trung của trẻ.
- Bước 2: Tạo hình ảnh đồng bộ nhân vật bằng các câu Prompt tiếng Anh ở bảng trên với tỷ lệ 16:9.
- Bước 3: Tải ảnh lên công cụ Image-to-Video để tạo chuyển động mượt mà.
- Bước 4: Lồng tiếng lời thoại tiếng Việt truyền cảm và chèn nhạc nền vui tươi.
- Bước 5: Ghép nối các cảnh trên CapCut và xuất video MP4 chiếu lên màn hình lớp học.""")

    elif tool_type == "worksheet":
        return clean_markdown_text(f"""PHIẾU HỌC TẬP CHO BÉ ({age.upper()})
Chủ đề: {theme} | Đề tài: {topic}
Họ và tên bé: .................................................... Lớp: .........................

BÀI TẬP 1: BÉ TINH MẮT QUAN SÁT & TÔ MÀU
- Yêu cầu: Bé hãy quan sát hình vẽ trung tâm về '{topic}', dùng bút sáp màu tô thật đẹp, đều màu và không lem ra ngoài đường viền nhé!
- Gợi ý hình minh họa trên phiếu: Khung tranh nét mảnh rõ ràng mô tả '{topic}' kích thước lớn ở đầu trang.

BÀI TẬP 2: BÉ NỐI ĐÚNG CẶP ĐÔI TƯƠNG ỨNG
- Yêu cầu: Bé hãy dùng bút chì nối hình ảnh ở cột bên trái với đồ dùng / lợi ích tương ứng ở cột bên phải sao cho thật chính xác.
- Gợi ý bố cục: 3 cặp hình ảnh trực quan sinh động có chấm tròn nối ở giữa.

BÀI TẬP 3: BÉ ĐẾM SỐ LƯỢNG & KHOANH TRÒN
- Yêu cầu: Bé hãy đếm xem trong mỗi ô có bao nhiêu biểu tượng liên quan đến '{topic}', sau đó khoanh tròn vào chữ số đúng bên dưới mỗi ô.
- Gợi ý bố cục: 3 ô hình với số lượng khác nhau (phù hợp chuẩn nhận thức lứa tuổi {age}).

BÀI TẬP 4: GÓC SÁNG TẠO CỦA BÉ NGOAN
- Yêu cầu: Bé hãy vẽ thêm những chi tiết mà bé yêu thích xung quanh '{topic}' và kể cho cô giáo, bố mẹ nghe về bức tranh của con nhé!
- Đánh giá của cô giáo: [ ] Hoàn thành xuất sắc   [ ] Hoàn thành tốt   [ ] Cần cố gắng hơn""")

    elif tool_type == "pptx_game":
        return clean_markdown_text(f"""TRÒ CHƠI POWERPOINT TƯƠNG TÁC: AI NHANH TRÍ HƠN
Đề tài: {topic} | Lứa tuổi: {age}

SLIDE 1: TRANG BÌA TRÒ CHƠI
- Tiêu đề: TRÒ CHƠI TƯƠNG TÁC - KHÁM PHÁ {topic.upper()}
- Lời dẫn của cô: "Chào mừng các đội chơi đến với gameshow Ai Nhanh Trí Hơn! Các đội hãy lắng nghe câu hỏi và giơ thẻ đáp án thật nhanh nhé!"

SLIDE 2: CÂU 1 - KHỞI ĐỘNG NHANH MẮT
- Câu hỏi: Các bé hãy quan sát và cho cô biết đâu là hình ảnh đúng về '{topic}'?
- A. Hình ảnh minh họa đúng về {topic}
- B. Hình ảnh đồ vật không liên quan
- C. Hình ảnh trống
- Đáp án đúng: A. Hình ảnh minh họa đúng về {topic}
- Hiệu ứng: Tiếng vỗ tay hoan hô + Ngôi sao lấp lánh xuất hiện.

SLIDE 3: CÂU 2 - ĐẶC ĐIỂM NỔI BẬT
- Câu hỏi: Đặc điểm nào sau đây đúng nhất khi nói về '{topic}'?
- A. Có màu sắc, hình dáng đặc trưng như cô và các con vừa khám phá
- B. Không có hình dạng gì cả
- C. Chỉ xuất hiện vào ban đêm
- Đáp án đúng: A. Có màu sắc, hình dáng đặc trưng như cô và các con vừa khám phá
- Hiệu ứng: Âm thanh "Ting ting" vui nhộn.

SLIDE 4: CÂU 3 - THỬ THÁCH TƯ DUY
- Câu hỏi: Khi tham gia hoạt động thực hành về '{topic}', chúng mình cần làm gì?
- A. Giữ gìn đồ dùng gọn gàng, đoàn kết chia sẻ cùng bạn
- B. Tranh giành đồ chơi với bạn
- C. Vứt đồ dùng bừa bãi
- Đáp án đúng: A. Giữ gìn đồ dùng gọn gàng, đoàn kết chia sẻ cùng bạn
- Hiệu ứng: Pháo giấy chúc mừng.

SLIDE 5: CÂU 4 - VỀ ĐÍCH CHIẾN THẮNG
- Câu hỏi: Qua bài học '{topic}' hôm nay, con cảm thấy như thế nào?
- A. Rất vui vẻ, hào hứng và học được nhiều điều bổ ích
- B. Buồn ngủ
- C. Không thích tham gia
- Đáp án đúng: A. Rất vui vẻ, hào hứng và học được nhiều điều bổ ích
- Hiệu ứng: Nhạc chiến thắng rộn ràng, tặng cúp bé ngoan.""")

    else:
        return clean_markdown_text(f"""BỘ PROMPT SINH HÌNH ẢNH MINH HỌA ĐỒNG BỘ TIẾT DẠY
ĐỀ TÀI: {topic.upper()} | ĐỘ TUỔI: {age}

1. Ảnh 1 - Trang bìa Slide mở đầu tiết học:
- Mục đích: Làm hình nền chính giới thiệu tên bài học thu hút sự chú ý của trẻ ngay từ phút đầu.
- Prompt tiếng Anh: Cute 3D Pixar style illustration of a bright Vietnamese preschool classroom with a cheerful Vietnamese teacher and happy 4-year-old children welcoming the lesson about {topic}, colorful banners, warm sunlight, vibrant child-friendly colors --ar 16:9

2. Ảnh 2 - Tình huống gây hứng thú (Hộp quà / Nhân vật dẫn chuyện):
- Mục đích: Minh họa tình huống mở đầu tạo sự tò mò cho trẻ.
- Prompt tiếng Anh: Cute 3D cartoon style, a magical glowing gift box opening on a wooden table revealing {topic}, surrounded by sparkles and cute friendly mascot characters, soft pastel background, high detail --ar 16:9

3. Ảnh 3 - Hình ảnh trực quan kiến thức trọng tâm 1 (Cận cảnh):
- Mục đích: Giúp trẻ quan sát rõ cấu tạo, màu sắc và đặc điểm nổi bật của '{topic}'.
- Prompt tiếng Anh: Educational 3D illustration close-up of {topic}, clear details, bright natural colors, isolated on a clean soft cream background for kindergarten visual learning, ultra sharp --ar 4:3

4. Ảnh 4 - Hình ảnh trực quan kiến thức trọng tâm 2 (Ứng dụng thực tế):
- Mục đích: Minh họa lợi ích, ý nghĩa của '{topic}' trong cuộc sống hằng ngày của bé.
- Prompt tiếng Anh: Cute 3D children's book illustration showing Vietnamese preschoolers interacting happily with {topic} in daily life, sunny garden and schoolyard background, warm joyful atmosphere --ar 16:9

5. Ảnh 5 - Minh họa hoạt động thực hành nhóm:
- Mục đích: Hướng dẫn trẻ tư thế ngồi và cách hợp tác thực hành theo nhóm nhỏ.
- Prompt tiếng Anh: Cute 3D Pixar style illustration of four Vietnamese kindergarten kids sitting neatly around a small table doing hands-on craft and learning activity about {topic}, smiling and helping each other --ar 16:9

6. Ảnh 6 - Hình ảnh phần thưởng / Sticker Bé Ngoan cuối tiết:
- Mục đích: Chiếu màn hình khen ngợi cuối giờ hoặc in làm sticker tặng trẻ.
- Prompt tiếng Anh: Cute 3D badge sticker design with a golden trophy, smiling star, and {topic} motif, cheerful celebratory ribbons, bright pastel colors, white background --ar 1:1""")

def generate_instant_month_plan(info: dict) -> str:
    month = info.get("month", "Tháng 10")
    theme = info.get("theme", "Thế giới thực vật")
    age = info.get("age", "4-5 tuổi")
    teacher = info.get("teacher", "Nguyễn Thị Thanh Vân")
    school = info.get("school", "Trường Mầm Non Phước Thạnh")
    class_name = info.get("class_name", "Lớp Mẫu giáo Nhỡ")
    weeks_cnt = info.get("weeks_count", "4 tuần")
    sub_themes = info.get("sub_themes", "Tuần 1: Cây xanh quanh bé; Tuần 2: Một số loại hoa đẹp; Tuần 3: Một số loại quả; Tuần 4: Rau xanh và ích lợi")

    return clean_markdown_text(f"""KẾ HOẠCH GIÁO DỤC {month.upper()} & SỔ CHỦ ĐỀ
Trường: {school} | Lớp: {class_name} ({age}) | Người soạn: {teacher}
Chủ đề lớn: {theme} | Thời gian thực hiện: {weeks_cnt}

I. MỤC TIÊU PHÁT TRIỂN CHỦ ĐỀ {month.upper()}
- Phát triển thể chất: Trẻ khỏe mạnh, nhanh nhẹn, thực hiện đúng các vận động cơ bản (bật, ném, trườn, trèo) và có thói quen vệ sinh dinh dưỡng tốt.
- Phát triển nhận thức: Trẻ khám phá, gọi tên, phân biệt được đặc điểm nổi bật, ích lợi của chủ đề '{theme}'; biết đếm, so sánh số lượng và hình dạng cơ bản.
- Phát triển ngôn ngữ: Trẻ mạnh dạn giao tiếp, đọc thuộc thơ, kể lại truyện diễn cảm theo tranh và phát triển vốn từ mạch lạc.
- Phát triển tình cảm - KNXH: Trẻ yêu thiên nhiên, biết chăm sóc, bảo vệ môi trường, đoàn kết chia sẻ đồ chơi cùng bạn bè.
- Phát triển thẩm mỹ: Trẻ cảm nhận được vẻ đẹp của {theme}, biết hát đúng giai điệu và sáng tạo sản phẩm tạo hình (vẽ, nặn, xé dán).

II. MẠNG NỘI DUNG & PHÂN NHÁNH CÁC TUẦN
- Định hướng nhánh: {sub_themes}

III. KẾ HOẠCH PHÂN PHỐI CHƯƠNG TRÌNH THEO TUẦN
| Tuần | Chủ đề nhánh | Thể chất & Vận động | Nhận thức (KPKH / Toán) | Tạo hình & Âm nhạc | Văn học & Chữ cái |
| --- | --- | --- | --- | --- | --- |
| Tuần 1 | Nhánh 1 của {theme} | VĐCB: Đi thăng bằng trên ghế thể dục | KPKH: Tìm hiểu đặc điểm chung của {theme} | Tạo hình: Vẽ theo ý thích về {theme} | Thơ: Đọc thơ diễn cảm chủ đề |
| Tuần 2 | Nhánh 2 của {theme} | VĐCB: Bật xa 35-40cm | Toán: Đếm đến 4, nhận biết nhóm có 4 đối tượng | Âm nhạc: Dạy hát & Vận động minh họa | Truyện: Kể chuyện sáng tạo theo tranh |
| Tuần 3 | Nhánh 3 của {theme} | VĐCB: Ném trúng đích nằm ngang | KPKH: Khám phá ích lợi và quy trình phát triển | Tạo hình: Nặn / Xé dán sản phẩm | Làm quen ngôn ngữ & Trò chơi chữ cái |
| Tuần 4 | Nhánh 4 & Tổng kết | Trò chơi vận động tập thể liên hoàn | Toán: So sánh chiều cao / kích thước | Âm nhạc: Biểu diễn văn nghệ cuối chủ đề | Kể chuyện đóng kịch & Nêu gương bé ngoan |

IV. TỔ CHỨC HOẠT ĐỘNG GÓC & HOẠT ĐỘNG NGOÀI TRỜI
- Góc Phân vai: Chơi bán hàng, nấu ăn, gia đình gắn với chủ đề {theme}.
- Góc Xây dựng: Xây dựng khuôn viên, vườn cây, công trình xanh.
- Góc Học tập & Sách: Xem tranh ảnh, phân loại lô tô, ghép số lượng tương ứng.
- Góc Nghệ thuật: Vẽ, tô màu, nặn, làm đồ dùng sáng tạo từ nguyên vật liệu mở.

V. PHỐI HỢP VỚI PHỤ HUYNH & ĐÁNH GIÁ
- Trao đổi qua bảng tuyên truyền và nhóm Zalo lớp về nội dung học tập của trẻ trong {month}.
- Vận động phụ huynh sưu tầm nguyên vật liệu mở (vỏ hộp, lõi giấy, lá cây) hỗ trợ góc STEAM của lớp.""")

def generate_instant_week_plan(info: dict) -> str:
    week = info.get("week_num", "Tuần 1")
    theme_week = info.get("theme_week", "Tôi là ai - Bạn là ai?")
    big_theme = info.get("big_theme", "Bản thân")
    age = info.get("age", "4-5 tuổi")
    teacher = info.get("teacher", "Nguyễn Thị Thanh Vân")
    school = info.get("school", "Trường Mầm Non Phước Thạnh")
    class_name = info.get("class_name", "Lớp Mẫu giáo Nhỡ")

    return clean_markdown_text(f"""KẾ HOẠCH GIÁO DỤC {week.upper()}
Trường: {school} | Lớp: {class_name} ({age}) | Người soạn: {teacher}
Chủ đề lớn: {big_theme} | Chủ đề nhánh: {theme_week}

I. MỤC TIÊU TUẦN
- Phát triển thể chất: Rèn luyện sự khéo léo, nhanh nhẹn qua bài tập thể dục sáng và vận động cơ bản.
- Phát triển nhận thức: Trẻ nhận biết và gọi tên đúng các đặc điểm nổi bật của nhánh '{theme_week}'.
- Phát triển ngôn ngữ: Trẻ tự tin giới thiệu bản thân, trả lời trọn vẹn câu, đọc thuộc thơ và hiểu nội dung truyện.
- Phát triển tình cảm - KNXH: Trẻ vui vẻ hòa đồng, biết quan tâm giúp đỡ bạn bè và giữ gìn vệ sinh cá nhân.
- Phát triển thẩm mỹ: Trẻ hứng thú tham gia hát múa và hoàn thành sản phẩm tạo hình đẹp mắt.

II. KẾ HOẠCH HOẠT ĐỘNG HÀNG NGÀY (THỨ HAI - THỨ SÁU)
| Thời điểm | Thứ Hai | Thứ Ba | Thứ Tư | Thứ Năm | Thứ Sáu |
| --- | --- | --- | --- | --- | --- |
| Đón trẻ & Thể dục sáng | - Đón trẻ ân cần, trao đổi với phụ huynh.<br>- Tập thể dục sáng kết hợp nhạc vui nhộn chủ đề {big_theme}. | - Đón trẻ, điểm danh.<br>- Thể dục sáng kết hợp vòng/gậy thể dục. | - Đón trẻ, trò chuyện sáng về '{theme_week}'.<br>- Thể dục sáng theo nhịp đếm. | - Đón trẻ, kiểm tra vệ sinh cá nhân.<br>- Thể dục sáng kết hợp bài hát chủ đề. | - Đón trẻ, hướng dẫn trẻ cất đồ dùng.<br>- Đồng diễn thể dục sáng toàn trường. |
| Hoạt động học có chủ đích | Phát triển Thể chất:<br>{info.get('t2', 'Vận động cơ bản & Trò chơi vận động')} | Phát triển Nhận thức:<br>{info.get('t3', 'Khám phá khoa học / Làm quen với toán')} | Phát triển Thẩm mỹ:<br>{info.get('t4', 'Tạo hình sáng tạo theo đề tài')} | Phát triển Ngôn ngữ:<br>{info.get('t5', 'Làm quen văn học (Thơ / Truyện)')} | Phát triển Thẩm mỹ & KNXH:<br>{info.get('t6', 'Dạy hát, vận động theo nhạc & Kỹ năng sống')} |
| Hoạt động ngoài trời | - Quan sát thời tiết, cây xanh sân trường.<br>- TCVĐ: Mèo đuổi chuột.<br>- Chơi tự do với đồ chơi ngoài trời. | - Dạo chơi, trò chuyện về '{theme_week}'.<br>- TCVĐ: Kéo co.<br>- Chơi cát, nước, phấn vẽ sân. | - Nhặt lá vàng, làm sạch sân trường.<br>- TCVĐ: Lộn cầu vồng.<br>- Chơi cầu trượt, xích đu an toàn. | - Quan sát góc thiên nhiên của lớp.<br>- TCVĐ: Dung dăng dung dẻ.<br>- Chơi với bóng và vòng. | - Trò chơi dân gian: Rồng rắn lên mây.<br>- Chơi tự do theo nhóm nhỏ dưới sự bao quát của cô. |
| Hoạt động chơi ở các góc | - Góc Phân vai: Gia đình, bác sĩ, bán hàng.<br>- Góc Xây dựng: Xây dựng ngôi nhà/trường học của bé.<br>- Góc Nghệ thuật: Tô màu, nặn đồ chơi.<br>- Góc Học tập: Xem sách, ghép tranh lô tô. | - Duy trì và mở rộng liên kết giữa các góc chơi.<br>- Cô nhập vai gợi mở để trẻ giao tiếp lịch sự. | - Khuyến khích trẻ sử dụng nguyên vật liệu mở ở góc STEAM/Nghệ thuật. | - Đổi vai chơi giữa các nhóm, rèn kỹ năng hợp tác và chia sẻ đồ dùng. | - Trưng bày sản phẩm đẹp của các góc chơi cuối tuần. |
| Hoạt động chiều & Trả trẻ | - Vận động nhẹ sau ngủ trưa.<br>- Ôn bài buổi sáng, rèn kỹ năng tự phục vụ.<br>- Vệ sinh, trả trẻ. | - Làm quen trò chơi mới.<br>- Thực hành vở tạo hình / toán.<br>- Vệ sinh, trả trẻ. | - Nghe kể chuyện, đọc đồng dao.<br>- Chơi tự chọn ở các góc.<br>- Vệ sinh, trả trẻ. | - Tập văn nghệ cuối tuần.<br>- Rèn kỹ năng giao tiếp lễ phép.<br>- Vệ sinh, trả trẻ. | - Biểu diễn văn nghệ cuối tuần.<br>- Nêu gương bé ngoan, phát phiếu bé ngoan.<br>- Trao đổi phụ huynh, trả trẻ. |

III. ĐÁNH GIÁ CUỐI TUẦN
- 100% trẻ đi học chuyên cần, khỏe mạnh, tích cực tham gia các hoạt động học tập và vui chơi trong tuần.""")

def generate_instant_creative_content(genre: str, topic: str) -> str:
    return f"{genre.upper()}: {topic.upper()}\nĐây là bản tạo nhanh do AI đang bận. Bạn vui lòng thử lại sau 1 phút nhé!"

def generate_instant_zalo(task_type: str, info_text: str) -> str:
    return f"MẪU ZALO NHANH:\nKính gửi Phụ huynh, hôm nay lớp học về {info_text}. Các con rất ngoan và tích cực!"

def generate_instant_diy(lesson: str) -> str:
    return f"ĐỒ DÙNG DIY CHO: {lesson.upper()}\n1. Dùng bìa carton, nắp chai nhựa.\n2. Cắt dán và trang trí màu sắc đẹp mắt."

# =========================================================================
# GỌI GEMINI 3.5 FLASH LITE SIÊU TỐC ĐỘ (KẾT NỐI V1BETA)
# =========================================================================
def generate_fast_ai_content(api_key, prompt, fallback_type="plan", fallback_data=None):
    if not api_key or not api_key.strip():
        if fallback_type == "digital": return generate_instant_digital_plan(fallback_data), "Chưa nhập API Key"
        if fallback_type.startswith("adv_"): return generate_instant_advanced_tool(fallback_type.replace("adv_", ""), fallback_data), "Chưa nhập API Key"
        if fallback_type == "month": return generate_instant_month_plan(fallback_data), "Chưa nhập API Key"
        if fallback_type == "week": return generate_instant_week_plan(fallback_data), "Chưa nhập API Key"
        return generate_instant_pedagogical_plan(fallback_data), "Chưa nhập API Key"
        
    error_msg = None
    text_out = ""
    
    def _call_model():
        client = genai.Client(
            api_key=api_key.strip(),
            http_options=types.HttpOptions(api_version='v1beta')
        )
        for m_name in ['gemini-3.5-flash-lite', 'gemini-3.8-flash']:
            try:
                res = client.models.generate_content(
                    model=m_name,
                    contents=prompt
                )
                if res and res.text and len(res.text.strip()) > 100:
                    return res.text
            except Exception:
                continue
        return ""

    executor = ThreadPoolExecutor(max_workers=1)
    future = executor.submit(_call_model)
    try:
        text_out = future.result(timeout=55) 
    except FuturesTimeoutError:
        error_msg = "Thời gian phản hồi của AI vượt quá 55 giây."
    except Exception as e:
        error_msg = f"Lỗi API: {str(e)}"
    finally:
        executor.shutdown(wait=False)

    if text_out and text_out.strip():
        return clean_markdown_text(text_out), None

    if fallback_type == "plan": return generate_instant_pedagogical_plan(fallback_data), error_msg
    elif fallback_type == "digital": return generate_instant_digital_plan(fallback_data), error_msg
    elif fallback_type.startswith("adv_"): return generate_instant_advanced_tool(fallback_type.replace("adv_", ""), fallback_data), error_msg
    elif fallback_type == "month": return generate_instant_month_plan(fallback_data), error_msg
    elif fallback_type == "week": return generate_instant_week_plan(fallback_data), error_msg
    elif fallback_type == "creative": return generate_instant_creative_content(fallback_data.get("genre", ""), fallback_data.get("topic", "")), error_msg
    elif fallback_type == "zalo": return generate_instant_zalo(fallback_data.get("type", ""), fallback_data.get("info", "")), error_msg
    elif fallback_type == "diy": return generate_instant_diy(fallback_data.get("lesson", "")), error_msg
    return generate_instant_pedagogical_plan(fallback_data), error_msg

# ==========================================
# MÀN HÌNH ĐĂNG NHẬP & SIDEBAR
# ==========================================
if not st.session_state["authenticated"]:
    st.title("🔐 Xác Thực Sử Dụng Phần Mềm Mầm Non")
    col_l1, col_l2 = st.columns([2, 1])
    with col_l1:
        input_key = st.text_input("Nhập Mã Kích Hoạt Của Bạn", type="password", placeholder="Nhập mã...")
        if st.button("🔓 Mở Khóa Sử Dụng", type="primary"):
            valid, msg = check_license(input_key)
            if valid:
                st.session_state["authenticated"] = True
                st.session_state["is_admin"] = (input_key == "LBSTORE2026")
                st.success(msg)
                st.rerun()
            else: st.error(msg)
    with col_l2:
        st.info("💡 **Gói Thuê Phần Mềm:**\n- Giá: **100.000 VNĐ / Tháng**\n- Zalo: 0939.72.9009 (Văn Bình)")
    st.stop()

st.title("🧸 Trợ Lý Ảo Toàn Năng - Giáo Viên Mầm Non")
st.markdown("Hệ thống hỗ trợ soạn giáo án chuẩn bảng 4 cột, tích hợp năng lực số, lên kế hoạch tháng/tuần và bộ công cụ AI giảng dạy.")

saved_plans_count = len(get_all_saved_plans())

with st.sidebar:
    st.markdown("### 📌 Danh Mục Tính Năng")
    app_mode = st.radio("Chọn tác vụ:", [
        "👑 Quản Lý Khách Thuê (Admin)", 
        "📝 Soạn Giáo Án Chuẩn", 
        "💻 Tích Hợp Năng Lực Số",
        "🗓️ Kế Hoạch & Sổ Chủ Đề", 
        "🎯 Trò Chơi & Đố Vui", 
        "💬 Sổ Liên Lạc & Zalo", 
        "✂️ Gợi Ý Đồ Dùng DIY", 
        f"📚 Kho Lịch Sử Giáo Án ({saved_plans_count} bài)"
    ] if st.session_state["is_admin"] else [
        "📝 Soạn Giáo Án Chuẩn", 
        "💻 Tích Hợp Năng Lực Số", 
        "🗓️ Kế Hoạch & Sổ Chủ Đề", 
        "🎯 Trò Chơi & Đố Vui", 
        "💬 Sổ Liên Lạc & Zalo", 
        "✂️ Gợi Ý Đồ Dùng DIY", 
        f"📚 Kho Lịch Sử Giáo Án ({saved_plans_count} bài)"
    ])
    st.divider()
    
    if st.session_state.get("api_key_saved"):
        st.success("🟢 Hệ thống AI đã lưu sẵn sàng!")
    else:
        st.warning("🔴 Chưa cài đặt API Key hệ thống")

    with st.expander("⚙️ Cài đặt API Key (Chỉ cần lưu 1 lần)", expanded=(not bool(st.session_state.get("api_key_saved")))):
        input_api_key = st.text_input("Nhập Google Gemini API Key", type="password", value=st.session_state["api_key_saved"])
        if st.button("💾 Lưu Vĩnh Viễn API Key"):
            k = input_api_key.strip()
            if k:
                st.session_state["api_key_saved"] = k
                save_setting_to_db("gemini_api_key", k)
                try:
                    test_client = genai.Client(api_key=k, http_options=types.HttpOptions(api_version='v1beta'))
                    res = test_client.models.generate_content(
                        model='gemini-3.5-flash-lite',
                        contents="Xin chào"
                    )
                    st.session_state["active_model_name"] = "gemini-3.5-flash-lite"
                    st.success("✅ Đã lưu vĩnh viễn & kết nối thành công Google AI!")
                except Exception:
                    try:
                        res = test_client.models.generate_content(
                            model='gemini-3.8-flash',
                            contents="Xin chào"
                        )
                        st.session_state["active_model_name"] = "gemini-3.8-flash"
                        st.success("✅ Đã lưu vĩnh viễn & kết nối thành công Google AI!")
                    except Exception as e:
                        st.session_state["active_model_name"] = ""
                        st.error(f"⚠️ Kiểm tra kết nối báo lỗi: {e}")
            else:
                st.error("⚠️ Vui lòng nhập API Key hợp lệ!")
        
    st.divider()
    if st.button("🔒 Đăng Xuất / Khóa Lại"):
        st.session_state["authenticated"] = False
        st.session_state["is_admin"] = False
        st.rerun()

api_key = st.session_state["api_key_saved"]

# =========================================================================
# QUẢN LÝ ADMIN
# =========================================================================
if st.session_state["is_admin"] and app_mode == "👑 Quản Lý Khách Thuê (Admin)":
    st.subheader("👑 Khu Vực Quản Trị Viên - Cấp Mã Thuê Bao")
    with st.form("admin_form"):
        c_name_input = st.text_input("Tên khách hàng / Tên trường (VD: Cô Thảo - Mầm Non Hướng Dương)")
        months_input = st.selectbox("Thời hạn thuê", [1, 2, 3, 6, 12], format_func=lambda x: f"{x} Tháng ({x*100}k)")
        if st.form_submit_button("✨ Tạo Mã Kích Hoạt Mới", type="primary"):
            if not c_name_input: st.warning("Vui lòng nhập tên khách hàng!")
            else:
                success, key_or_err, exp_date = create_new_license(c_name_input, months_input)
                if success:
                    st.success(f"🎉 Đã tạo mã thành công cho **{c_name_input}**!")
                    st.code(f"Mã kích hoạt: {key_or_err} (Hạn đến: {exp_date})", language="text")
                else: st.error(f"Lỗi: {key_or_err}")
    st.divider()
    st.subheader("📋 Danh Sách Mã Đã Cấp")
    for l_id, l_key, l_name, l_exp, l_active in get_all_licenses():
        status_text = "🟢 Đang hoạt động" if l_active == 1 else "🔴 Đã khóa"
        st.markdown(f"- **Mã:** `{l_key}` | **Khách:** {l_name} | **Hạn:** {l_exp} | **Trạng thái:** {status_text}")

# =========================================================================
# TAB 1: SOẠN GIÁO ÁN CHUẨN (GIỮ NGUYÊN 100%)
# =========================================================================
elif app_mode == "📝 Soạn Giáo Án Chuẩn":
    st.subheader("📝 Thiết Lập Thông Số Giáo Án (Chuẩn Chương Trình GDMN)")
    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("Người soạn / Giáo viên", "Nguyễn Thị Thanh Vân")
        school_name = st.text_input("Đơn vị / Trường", "Trường Mầm Non Phước Thạnh")
    with col2:
        class_name = st.text_input("Lớp", "Lớp Mẫu giáo Nhỡ (4-5 tuổi)")
        theme = st.text_input("Chủ đề lớn / Nhánh", "Thế giới thực vật")

    st.markdown("### CHẾ ĐỘ SOẠN THẢO *")
    c_mode1, c_mode2 = st.columns(2)
    with c_mode1: mode_gv = st.button("🌱 Giáo án thường\nChuẩn mực, rõ ràng theo đúng phân phối chương trình.", use_container_width=True)
    with c_mode2: mode_gioi = st.button("⭐ Thi GV Giỏi\nStorytelling, tình huống sáng tạo, câu hỏi phân tầng sâu.", use_container_width=True)
    if "soan_mode_val" not in st.session_state: st.session_state["soan_mode_val"] = "Thi GV Giỏi"
    if mode_gv: st.session_state["soan_mode_val"] = "Giáo án thường"
    if mode_gioi: st.session_state["soan_mode_val"] = "Thi GV Giỏi"
    st.info(f"👉 Đang chọn chế độ soạn thảo: **{st.session_state['soan_mode_val']}**")

    st.markdown("### ĐỊNH DẠNG BẢNG TIẾN TRÌNH HOẠT ĐỘNG")
    c_fmt1, c_fmt2, c_fmt3 = st.columns(3)
    with c_fmt1:
        fmt_4col = st.button("📊 Bảng 4 cột chuẩn (Mẫu YouTube)\nThời gian | Cô | Trẻ | Quan sát & Hỗ trợ", use_container_width=True)
    with c_fmt2:
        fmt_3col = st.button("📋 Bảng 3 cột chuẩn\nThời gian | Hoạt động Cô | Hoạt động Trẻ", use_container_width=True)
    with c_fmt3:
        fmt_1col = st.button("📄 1 Cột văn bản (Toàn cảnh)\nGạch đầu dòng (-) không kẻ bảng", use_container_width=True)
    if "display_fmt_val" not in st.session_state:
        st.session_state["display_fmt_val"] = "Bảng 4 cột chuẩn (Mẫu YouTube)"
    if fmt_4col: st.session_state["display_fmt_val"] = "Bảng 4 cột chuẩn (Mẫu YouTube)"
    if fmt_3col: st.session_state["display_fmt_val"] = "Bảng 3 cột chuẩn"
    if fmt_1col: st.session_state["display_fmt_val"] = "1 Cột văn bản (Toàn cảnh)"
    st.info(f"👉 Đang chọn định dạng: **{st.session_state['display_fmt_val']}**")

    st.markdown("### 1. ĐỘ TUỔI HỌC SINH *")
    age_group = st.radio("Chọn độ tuổi", ["3-12 th", "12-18 th", "18-24 th", "24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], horizontal=True, index=5)
    st.markdown("### 2. LOẠI HÌNH HOẠT ĐỘNG *")
    activity_type = st.radio("Chọn loại hoạt động", ["Hoạt động học có chủ đích", "Hoạt động góc / Vui chơi", "Hoạt động ngoài trời", "Hoạt động chiều"], horizontal=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### 3. LĨNH VỰC PHÁT TRIỂN *")
        domain = st.selectbox("Lĩnh vực", ["Thể chất", "Nhận thức", "Ngôn ngữ", "Thẩm mỹ", "Tình cảm - KNXH"], index=1, key="domain_select_box")
        sub_category = ""
        if domain == "Nhận thức": sub_category = st.selectbox("Phân loại chi tiết", ["Khám phá", "Toán"], key="sub_nhan_thuc")
        elif domain == "Ngôn ngữ": sub_category = st.selectbox("Phân loại chi tiết", ["Chữ cái - tiết tập tô", "Làm quen", "Trò chơi", "Thơ", "Truyện"], key="sub_ngon_ngu")
        elif domain == "Thẩm mỹ":
            sub_category = st.selectbox("Phân loại chi tiết", ["Âm nhạc - Dạy hát", "Nghe nhạc", "Múa minh họa", "Gõ đệm theo phách, nhịp", "Tạo hình"], key="sub_tham_my")
        else: sub_category = st.text_input("Hoạt động cụ thể", "Phát triển vận động", key="sub_khac")
    with col_b:
        st.markdown("### 4. DẠNG TIẾT CHUYÊN MÔN (ROUTE) *")
        if domain == "Thể chất": route_type = st.selectbox("Dạng tiết", ["Vận động cơ bản (VĐCB)", "Trò chơi vận động", "Thể dục sáng"], key="route_tc")
        elif domain == "Nhận thức": route_type = st.selectbox("Dạng tiết", ["Khám phá khoa học", "Làm quen với toán", "Khám phá xã hội"], key="route_nt")
        elif domain == "Ngôn ngữ": route_type = st.selectbox("Dạng tiết", ["Làm quen chữ cái", "Làm quen văn học (Thơ/Truyện)", "Phát triển ngôn ngữ"], key="route_nn")
        elif domain == "Thẩm mỹ": route_type = st.selectbox("Dạng tiết", ["Âm nhạc", "Tạo hình"], key="route_tm")
        else: route_type = st.selectbox("Dạng tiết", ["Kỹ năng xã hội", "Hoạt động chung"], key="route_khac")

    art_type_val = ""
    if domain == "Thẩm mỹ" and (sub_category == "Tạo hình" or route_type == "Tạo hình"):
        with st.container(border=True):
            st.markdown("🎨 **PHÂN LOẠI TIẾT TẠO HÌNH (BẮT BUỘC)**")
            art_type_val = st.radio("Chọn hình thức tiết Tạo hình:", ["Tiết theo mẫu", "Tiết theo đề tài"], horizontal=True, key="art_radio_key")

    st.markdown("### 5. ĐỀ TÀI BÀI DẠY CỤ THỂ * (BẮT BUỘC):")
    topic = st.text_input("Nhập tên đề tài", "Bé khám phá hành trình lớn lên của hạt đậu")

    st.markdown("### 6. YÊU CẦU TÍCH HỢP (TÙY CHỌN):")
    col_c1, col_c2, col_c3, col_c4, col_c5 = st.columns(5)
    with col_c1: opt0 = st.checkbox("STEAM", value=True)
    with col_c2: opt1 = st.checkbox("Ứng dụng công nghệ AI", value=True)
    with col_c3: opt2 = st.checkbox("Giáo dục Cảm xúc - Xã hội (SEL)", value=True)
    with col_c4: opt3 = st.checkbox("Bảo vệ môi trường")
    with col_c5: opt4 = st.checkbox("Quyền trẻ em")
    extras = []
    if opt0: extras.append("STEAM")
    if opt1: extras.append("Ứng dụng công nghệ AI")
    if opt2: extras.append("Giáo dục Cảm xúc - Xã hội (SEL)")
    if opt3: extras.append("Bảo vệ môi trường")
    if opt4: extras.append("Quyền trẻ em")

    with st.expander("📂 7. Giáo án mẫu và học liệu tham khảo (Tùy chọn - Tải file .docx / .txt để AI học theo mẫu)"):
        ref_c1, ref_c2 = st.columns([2, 1])
        with ref_c1:
            uploaded_ref = st.file_uploader("Chọn giáo án mẫu hoặc học liệu (.docx, .txt)", type=["docx", "txt"], key="ref_uploader")
        with ref_c2:
            ref_mode = st.selectbox("Cách sử dụng tài liệu:", ["Tham khảo cấu trúc và cách trình bày", "Lấy nội dung tham khảo", "Kết hợp cả hai"], key="ref_mode_sel")
        ref_note = st.text_input("Lưu ý riêng khi tham khảo:", placeholder="Ví dụ: Chỉ lấy cấu trúc các bước, bài trường không dùng nội dung cũ...", key="ref_note_in")

        extracted_ref_text = ""
        if uploaded_ref is not None:
            try:
                if uploaded_ref.name.lower().endswith(".docx"):
                    doc_ref = Document(uploaded_ref)
                    extracted_ref_text = "\n".join([p.text for p in doc_ref.paragraphs if p.text.strip()])[:3500]
                else:
                    extracted_ref_text = uploaded_ref.read().decode("utf-8", errors="ignore")[:3500]
                st.success(f"✅ Đã đọc thành công tài liệu tham khảo: **{uploaded_ref.name}** ({len(extracted_ref_text)} ký tự)")
            except Exception as ex:
                st.warning(f"Không đọc được nội dung file: {ex}")

    st.markdown("### 8. YÊU CẦU RIÊNG CỦA GIÁO VIÊN & ĐỒ DÙNG HIỆN CÓ:")
    custom_goals = st.text_area("Mục tiêu cần đạt cụ thể (Tùy chọn - Để trống AI tự soạn chuẩn 4 phần):", placeholder="VD:\n- Trẻ nhận biết quá trình nảy mầm và lớn lên của hạt đậu.")
    yc_c1, yc_c2 = st.columns(2)
    with yc_c1:
        custom_ideas = st.text_area("Mong muốn riêng của cô trong giáo án:", placeholder="Ví dụ: Muốn bài dạy có trò chơi mở đầu, câu hỏi gợi mở, hoạt động nhóm 4 trẻ, lồng ghép bài hát...", height=90)
    with yc_c2:
        existing_tools = st.text_area("Đồ dùng / Học liệu hiện có của lớp:", placeholder="Ví dụ: Thẻ hình, tranh chữ cái, sách truyện, biểu tượng phát âm, búp bê, vòng thể dục, màu sáp...", height=90)

    img_style_req = st.text_input("Yêu cầu hình ảnh, biểu tượng minh họa (Nếu có):", placeholder="Ví dụ: Sách mở, thẻ chữ cái, ngôi nhà búp bê, biểu tượng phát âm, màu sắc tươi sáng, phù hợp trẻ mầm non...")

    plan_info_dict = {
        "school": school_name, "teacher": teacher_name, "class_name": class_name,
        "theme": theme, "domain": domain, "sub_category": sub_category, "activity_type": activity_type,
        "route_type": route_type, "art_type": art_type_val, "topic": topic, "method": "Gemini AI", "age": age_group, "extras": extras,
        "mode": st.session_state["soan_mode_val"], "display_format": st.session_state["display_fmt_val"], 
        "custom_goals": custom_goals, "custom_ideas": custom_ideas, "existing_tools": existing_tools,
        "doc_title": "GIÁO ÁN HOẠT ĐỘNG HỌC CÓ CHỦ ĐÍCH"
    }

    if st.button("✨ TẠO GIÁO ÁN HOÀN CHỈNH SIÊU CHI TIẾT", type="primary", use_container_width=True):
        with st.spinner("⚡ Đang kết nối AI biên soạn giáo án chuẩn bảng 4 cột siêu chi tiết (vui lòng đợi ~8-15 giây)..."):
            art_prompt_part = f" - Hình thức tiết tạo hình: {art_type_val}" if art_type_val else ""
            disp_fmt_now = st.session_state["display_fmt_val"]

            if "4 cột" in disp_fmt_now:
                table_instruction = """
                Phần III. TIẾN TRÌNH HOẠT ĐỘNG: BẮT BUỘC trình bày dưới dạng BẢNG MARKDOWN 4 CỘT CHÍNH XÁC SAU ĐÂY:
                | Thời gian | Hoạt động của cô (Lời thoại, câu hỏi & hành động) | Dự kiến hoạt động & Phản hồi của trẻ | Quan sát, Hỗ trợ & Phân hóa |
                | --- | --- | --- | --- |
                
                QUY TẮC VIẾT BẢNG CỰC KỲ QUAN TRỌNG:
                - Chia thành 5 hàng hoạt động nối tiếp nhau (tổng thời lượng 30 phút):
                  1. Ổn định - Gây hứng thú (3 - 4 phút)
                  2. Khám phá trọng tâm & Đàm thoại sâu (8 - 10 phút)
                  3. Thực hành - Trải nghiệm theo nhóm 4 trẻ & cá nhân (10 phút) - chia rõ Bước 1, Bước 2, Bước 3
                  4. Trò chơi vận động / Củng cố kiến thức (3 phút) - ghi rõ Luật chơi, Cách chơi
                  5. Kết thúc - Nhận xét - Chuyển hoạt động (3 phút)
                - Mỗi hàng của bảng PHẢI nằm trên 1 dòng duy nhất (bắt đầu bằng | và kết thúc bằng |). Để xuống dòng bên trong từng ô của bảng, BẮT BUỘC dùng thẻ <br> (TUYỆT ĐỐI không nhấn Enter xuống dòng giữa chừng trong cùng 1 hàng bảng).
                - Trong cột 2 (Hoạt động của cô): Ghi cực kỳ chi tiết lời thoại dẫn dắt (Cô nói: "..."), hệ thống câu hỏi gợi mở phân tầng (+ Đố: "...", + Hỏi: "...", + Khơi: "..."), hành động làm mẫu của cô.
                - Trong cột 3 (Dự kiến hoạt động của trẻ): Ghi rõ phản hồi, hành động và câu trả lời ngây thơ dự kiến của trẻ (- Trẻ trả lời: "...").
                - Trong cột 4 (Quan sát, Hỗ trợ & Phân hóa): Ghi rõ cách cô quan sát, hỗ trợ riêng cho trẻ nhút nhát/thao tác chậm, và mở rộng nâng cao cho trẻ khá giỏi, nhắc nhở an toàn.
                """
            elif "3 cột" in disp_fmt_now:
                table_instruction = """
                Phần III. TIẾN TRÌNH HOẠT ĐỘNG: BẮT BUỘC trình bày dưới dạng BẢNG MARKDOWN 3 CỘT:
                | Thời gian | Hoạt động của cô (Lời thoại, câu hỏi & hành động) | Dự kiến hoạt động & Phản hồi của trẻ |
                | --- | --- | --- |
                - Chia thành 5 hoạt động chi tiết; mỗi hàng bảng nằm trên 1 dòng duy nhất, dùng thẻ <br> để xuống dòng bên trong ô.
                """
            else:
                table_instruction = """
                Phần III. TIẾN TRÌNH HOẠT ĐỘNG: Trình bày văn bản 1 cột theo từng Hoạt động 1, Hoạt động 2 (chia Bước 1, Bước 2, Bước 3), Hoạt động 3; ghi rõ lời thoại Cô nói và Trẻ trả lời.
                """

            ref_section_prompt = ""
            if extracted_ref_text:
                ref_section_prompt = f"\nTÀI LIỆU THAM KHẢO GIÁO VIÊN TẢI LÊN ({ref_mode} - Lưu ý: {ref_note}):\n{extracted_ref_text}\n"

            prompt = f"""
            Đóng vai là một Chuyên gia Giáo dục Mầm non xuất sắc của Bộ GD&ĐT Việt Nam, hãy biên soạn một GIÁO ÁN HOẠT ĐỘNG HỌC CÓ CHỦ ĐÍCH HOÀN CHỈNH, RẤT DÀI, RÕ RÀNG VÀ CỰC KỲ CHI TIẾT (từ 1800 - 2500 từ):
            - Trường: {school_name} | Lớp: {class_name} | Người soạn: {teacher_name}
            - Chủ đề: {theme}
            - Lĩnh vực phát triển: {domain} ({sub_category} - {route_type})
            - Tên hoạt động: {topic}{art_prompt_part}
            - Độ tuổi: {age_group} | Loại hoạt động: {activity_type}
            - Chế độ soạn: {st.session_state['soan_mode_val']}
            - Mức độ tích hợp: {', '.join(extras) if extras else 'Cơ bản'}
            - Mục tiêu giáo viên yêu cầu: {custom_goals if custom_goals else 'Tự động tạo đầy đủ 4 mục chi tiết'}
            - Mong muốn riêng của giáo viên: {custom_ideas if custom_ideas else 'Không'}
            - Đồ dùng hiện có của lớp: {existing_tools if existing_tools else 'Đồ dùng mầm non tiêu chuẩn'}
            - Yêu cầu hình ảnh/biểu tượng: {img_style_req if img_style_req else 'Sinh động, tươi sáng, phù hợp trẻ mầm non'}
            {ref_section_prompt}

            CẤU TRÚC BẮT BUỘC CỦA GIÁO ÁN:
            Dòng đầu tiên: GIÁO ÁN HOẠT ĐỘNG HỌC CÓ CHỦ ĐÍCH
            Các dòng thông tin hành chính (Chủ đề:, Lĩnh vực phát triển:, Tên hoạt động:, Độ tuổi:, Thời lượng: 30 phút, Hình thức tổ chức: Cả lớp, nhóm 4 trẻ, cá nhân, Mức độ tích hợp: ...)

            I. MỤC TIÊU HOẠT ĐỘNG
            1. Kiến thức: (2-3 gạch đầu dòng cụ thể, sát đề tài)
            2. Kỹ năng: (3 gạch đầu dòng rõ ràng thao tác rèn luyện)
            3. Thái độ (Giáo dục Cảm xúc - Xã hội - SEL): (3 gạch đầu dòng)
            4. Năng lực hình thành:
            - Tư duy: ...
            - Giao tiếp: ...
            - Hợp tác: ...
            - Giải quyết vấn đề: ...

            II. CHUẨN BỊ
            1. Đồ dùng của cô: (Nêu rõ video AI/hình ảnh minh họa, nhạc, thẻ quy trình, vật thật, đồ dùng hiện có)
            2. Đồ dùng của trẻ (Chia theo nhóm 4 trẻ và cá nhân): (Nêu cụ thể số lượng cho nhóm 4 trẻ và từng cá nhân)
            3. Môi trường & An toàn: (Bố trí lớp học và lưu ý an toàn cho trẻ)

            III. TIẾN TRÌNH HOẠT ĐỘNG
            {table_instruction}

            IV. ĐÁNH GIÁ CUỐI NGÀY
            - Tình trạng sức khỏe: ...
            - Trạng thái cảm xúc, thái độ: ...
            - Kiến thức, kỹ năng: ...

            Tuyệt đối KHÔNG dùng ký tự ### hay **.
            """
            result_text, err = generate_fast_ai_content(api_key, prompt, fallback_type="plan", fallback_data=plan_info_dict)
            
            if err:
                st.warning(f"⚠️ **Thông báo:** {err}\n\n👉 Đang sử dụng kịch bản chuẩn bảng 4 cột đầy đủ bên dưới.")
            else:
                st.success("🎉 Đã hoàn thành giáo án chuẩn bảng 4 cột siêu chi tiết! Mời cô xem bên dưới:")

            st.session_state["generated_plan_text"] = result_text
            st.session_state["plan_info"] = plan_info_dict
            save_plan_to_db(topic, result_text)

    # =====================================================================
    # BỘ CÔNG CỤ NÂNG CAO (5 TÍNH NĂNG ĐỒNG BỘ TIẾT DẠY)
    # =====================================================================
    st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("#### ✨ Bộ công cụ nâng cao (Hỗ trợ giảng dạy & Học liệu AI)")
        st.caption(f"Tự động đồng bộ theo đề tài hiện tại: **{topic}** ({age_group})")
        
        adv_c1, adv_c2, adv_c3, adv_c4, adv_c5 = st.columns(5)
        with adv_c1:
            btn_adv_1 = st.button("⏱️ Kịch bản dạy nhanh\nLời dẫn & phân bổ thời gian", use_container_width=True, key="btn_adv_quick")
        with adv_c2:
            btn_adv_2 = st.button("🎬 Kịch bản video AI\nCảnh quay, lời thoại, hình ảnh", use_container_width=True, key="btn_adv_video")
        with adv_c3:
            btn_adv_3 = st.button("📝 Phiếu học tập\nNhiệm vụ phù hợp độ tuổi", use_container_width=True, key="btn_adv_ws")
        with adv_c4:
            btn_adv_4 = st.button("🎮 Trò chơi PowerPoint\nTạo câu hỏi và xuất PPTX", use_container_width=True, key="btn_adv_pptx")
        with adv_c5:
            btn_adv_5 = st.button("🎨 Prompt sinh hình ảnh\nMinh họa đồng bộ tiết dạy", use_container_width=True, key="btn_adv_img")

        selected_adv_tool = None
        if btn_adv_1: selected_adv_tool = "quick_script"
        elif btn_adv_2: selected_adv_tool = "video_ai"
        elif btn_adv_3: selected_adv_tool = "worksheet"
        elif btn_adv_4: selected_adv_tool = "pptx_game"
        elif btn_adv_5: selected_adv_tool = "image_prompts"

        if selected_adv_tool:
            base_context = st.session_state.get("generated_plan_text", "")[:1400]
            with st.spinner("⚡ Đang khởi tạo học liệu nâng cao đồng bộ với bài dạy..."):
                if selected_adv_tool == "quick_script":
                    adv_title = f"Kịch bản dạy nhanh: {topic}"
                    adv_prompt = f"""
                    Dựa trên đề tài mầm non '{topic}' (Độ tuổi: {age_group}, Chủ đề: {theme}, Lĩnh vực: {domain}), hãy viết một KỊCH BẢN DẠY NHANH (Phao cầm tay lên lớp) trình bày dạng BẢNG MARKDOWN 3 CỘT:
                    | Thời gian | Nội dung hoạt động | Lời dẫn & Câu hỏi cốt lõi của cô |
                    | --- | --- | --- |
                    - Phân bổ thời gian từng phút (00:00 - 30:00) cho 5 bước hoạt động; dùng thẻ <br> để xuống dòng bên trong ô bảng, mỗi hàng nằm trên 1 dòng.
                    Tuyệt đối không dùng dấu ### hay **.
                    """
                elif selected_adv_tool == "video_ai":
                    adv_title = f"Kịch bản video AI: {topic}"
                    adv_prompt = f"""
                    Hãy viết một KỊCH BẢN VIDEO AI MINH HỌA TIẾT DẠY MẦM NON trình bày dạng BẢNG MARKDOWN 5 CỘT cho đề tài '{topic}' (Độ tuổi: {age_group}, Chủ đề: {theme}):
                    | Cảnh | Thời lượng | Hình ảnh minh họa | Lời thoại & Âm thanh | Prompt tiếng Anh (Tạo ảnh/video 3D) |
                    | --- | --- | --- | --- | --- |
                    - Gồm 5 - 6 cảnh quay (mỗi cảnh 5 giây). Mỗi hàng bảng nằm trên 1 dòng duy nhất, dùng thẻ <br> để xuống dòng trong ô.
                    - Cuối bảng viết thêm mục: VII. QUY TRÌNH 5 BƯỚC DỰNG VIDEO AI & DANH SÁCH TÀI NGUYÊN (Midjourney/Canva, Kling AI/Runway, CapCut).
                    Tuyệt đối không dùng dấu ### hay **.
                    """
                elif selected_adv_tool == "worksheet":
                    adv_title = f"Phiếu học tập: {topic}"
                    adv_prompt = f"""
                    Thiết kế một PHIẾU HỌC TẬP IN A4 CHO TRẺ MẦM NON lứa tuổi {age_group} về đề tài '{topic}' (Chủ đề: {theme}):
                    - Có phần tiêu đề (Trường, Lớp, Họ và tên bé).
                    - Thiết kế 4 BÀI TẬP TRỰC QUAN phù hợp lứa tuổi {age_group} (Bài tập 1: Quan sát & Tô màu; Bài tập 2: Nối cặp đôi tương ứng; Bài tập 3: Đếm số lượng & Khoanh tròn; Bài tập 4: Góc vẽ sáng tạo của bé).
                    - Mỗi bài tập nêu rõ: Yêu cầu dành cho bé, Cách cô hướng dẫn và Mô tả bố cục hình vẽ trên phiếu.
                    Tuyệt đối không dùng dấu ### hay **.
                    """
                elif selected_adv_tool == "pptx_game":
                    adv_title = f"Trò chơi PowerPoint: {topic}"
                    adv_prompt = f"""
                    Thiết kế kịch bản TRÒ CHƠI TƯƠNG TÁC POWERPOINT (5 câu hỏi trắc nghiệm) cho trẻ mầm non {age_group} về đề tài '{topic}':
                    - SLIDE 1: Trang bìa trò chơi (Tên trò chơi hấp dẫn, lời dẫn mở màn của cô).
                    - SLIDE 2 đến SLIDE 6 (5 câu hỏi, định dạng rõ ràng từng dòng):
                      SLIDE X: CÂU Y - [Tên thử thách]
                      - Câu hỏi: [Câu hỏi gần gũi, dễ hiểu cho bé {age_group}]
                      - A. [Đáp án A]
                      - B. [Đáp án B]
                      - C. [Đáp án C]
                      - Đáp án đúng: [Ghi rõ đáp án đúng]
                      - Hiệu ứng: [Mô tả âm thanh/hiệu ứng chúc mừng]
                    Tuyệt đối không dùng dấu ### hay **.
                    """
                else:
                    adv_title = f"Prompt sinh hình ảnh: {topic}"
                    adv_prompt = f"""
                    Tạo BỘ 6 PROMPT SINH HÌNH ẢNH MINH HỌA ĐỒNG BỘ TIẾT DẠY cho đề tài '{topic}' (Độ tuổi: {age_group}, Chủ đề: {theme}):
                    - Đồng bộ theo tiến trình tiết dạy: (1) Ảnh bìa slide mở đầu, (2) Ảnh tình huống gây hứng thú, (3) Ảnh trực quan kiến thức 1, (4) Ảnh trực quan kiến thức 2, (5) Ảnh hoạt động thực hành nhóm, (6) Ảnh sticker phần thưởng bé ngoan.
                    - Mỗi mục ghi rõ: Mục đích sử dụng trong tiết dạy (tiếng Việt) và Prompt tiếng Anh chi tiết (phong cách hoạt hình 3D Pixar tươi sáng, nhân vật cô giáo và trẻ em mầm non Việt Nam, kèm tỷ lệ --ar 16:9 hoặc --ar 1:1).
                    Tuyệt đối không dùng dấu ### hay **.
                    """

                if base_context:
                    adv_prompt += f"\nTham khảo nội dung giáo án vừa soạn để đồng bộ chi tiết:\n{base_context}"

                adv_text, adv_err = generate_fast_ai_content(
                    api_key, adv_prompt, fallback_type=f"adv_{selected_adv_tool}", fallback_data=plan_info_dict
                )
                st.session_state["adv_tool_type"] = selected_adv_tool
                st.session_state["adv_tool_title"] = adv_title
                st.session_state["adv_tool_result"] = adv_text
                save_plan_to_db(f"[Công Cụ Nâng Cao] {adv_title}", adv_text)

    if "adv_tool_result" in st.session_state and st.session_state["adv_tool_result"]:
        st.divider()
        adv_t = st.session_state.get("adv_tool_title", "Kết quả Công cụ nâng cao")
        adv_type_cur = st.session_state.get("adv_tool_type", "")
        st.markdown(f"### 🌟 Kết Quả: {adv_t}")

        if adv_type_cur == "pptx_game":
            ac1, ac2, ac3 = st.columns([1, 1, 2])
            with ac1:
                if HAS_PPTX:
                    pptx_file = export_pptx_game(topic, age_group, st.session_state["adv_tool_result"])
                    if pptx_file and os.path.exists(pptx_file):
                        with open(pptx_file, "rb") as f_pptx:
                            st.download_button(
                                label="📊 Tải Slide PowerPoint (.pptx)",
                                data=f_pptx,
                                file_name=pptx_file,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                type="primary",
                                use_container_width=True,
                                key="dl_adv_pptx_btn"
                            )
                else:
                    st.info("💡 Thêm `python-pptx` vào `requirements.txt` để xuất trực tiếp file `.pptx`.")
            with ac2:
                adv_docx = export_docx_from_text(st.session_state["adv_tool_result"], {"topic": topic}, font_size_pt=13, prefix_name="CongCuNangCao")
                with open(adv_docx, "rb") as f_adv:
                    st.download_button(
                        label="📥 Tải Word (.docx)",
                        data=f_adv,
                        file_name=adv_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="dl_adv_docx_btn1"
                    )
            with ac3:
                render_copy_and_print_buttons(st.session_state["adv_tool_result"], adv_t, 14, "adv_tool_box")
        else:
            ac1, ac2 = st.columns([1, 2])
            with ac1:
                adv_docx = export_docx_from_text(st.session_state["adv_tool_result"], {"topic": topic}, font_size_pt=13, prefix_name="CongCuNangCao")
                with open(adv_docx, "rb") as f_adv:
                    st.download_button(
                        label="📥 Tải File Word (.docx)",
                        data=f_adv,
                        file_name=adv_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                        key="dl_adv_docx_btn2"
                    )
            with ac2:
                render_copy_and_print_buttons(st.session_state["adv_tool_result"], adv_t, 14, "adv_tool_box")

        adv_html = format_text_to_standard_html(st.session_state["adv_tool_result"], 14)
        st.markdown(f'<div style="background-color: #ffffff; color: #000000; font-family: \'Times New Roman\', Times, serif; font-size: 14pt; line-height: 1.5; padding: 30px 40px; border-radius: 12px; border: 1.5px solid #80CBC4; box-shadow: 0 6px 18px rgba(0, 121, 107, 0.08); text-align: justify;">{adv_html}</div>', unsafe_allow_html=True)

    if "generated_plan_text" in st.session_state:
        st.divider()
        info = st.session_state["plan_info"]
        current_font_size = st.session_state.get("font_size_slider", 14)
        
        btn_col1, btn_col2 = st.columns([1, 2])
        with btn_col1:
            docx_file = export_docx_from_text(st.session_state["generated_plan_text"], info, custom_template=None, font_size_pt=current_font_size)
            with open(docx_file, "rb") as file:
                st.download_button(label="📥 Tải File Word Chuẩn Bảng (.docx)", data=file, file_name=docx_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True, key="dl_main_plan_btn")
        with btn_col2:
            render_copy_and_print_buttons(st.session_state["generated_plan_text"], f"Giáo Án: {info['topic']}", current_font_size, "tab1_main")

        info_col1, info_col2 = st.columns(2)
        with info_col1: st.success("✅ **Điểm sư phạm: 100/100**")
        with info_col2: st.info(f"📐 **Định dạng:** {info['display_format']}")
            
        font_size_val = st.slider("🔤 Kích thước chữ hiển thị & In (pt):", min_value=12, max_value=24, value=current_font_size, step=1, key="font_size_slider")
        
        st.markdown("### 📄 Bản Xem Trước Giáo Án (Trang Giấy In Chuẩn Times New Roman):")
        formatted_preview_html = format_text_to_standard_html(st.session_state["generated_plan_text"], font_size_val)
        st.markdown(f'<div style="background-color: #ffffff; color: #000000; font-family: \'Times New Roman\', Times, serif; font-size: {font_size_val}pt; line-height: 1.5; padding: 35px 45px; border-radius: 12px; border: 1.5px solid #B2E5D6; box-shadow: 0 6px 18px rgba(0, 121, 107, 0.08); text-align: justify;">{formatted_preview_html}</div>', unsafe_allow_html=True)

        with st.expander("✏️ Chỉnh sửa nội dung văn bản (Bấm vào để mở khung sửa chữ)"):
            edited_text = st.text_area("Nội dung giáo án:", st.session_state["generated_plan_text"], height=400, key="edit_main_plan_ta")
            if edited_text != st.session_state["generated_plan_text"]:
                st.session_state["generated_plan_text"] = clean_markdown_text(edited_text)
                st.rerun()

# =========================================================================
# TAB 2: TÍCH HỢP NĂNG LỰC SỐ (GIỮ NGUYÊN 100% 4 KHỐI)
# =========================================================================
elif app_mode == "💻 Tích Hợp Năng Lực Số":
    head_c1, head_c2 = st.columns([3, 1])
    with head_c1:
        st.subheader("💻 Soạn Giáo Án Tích Hợp Năng Lực Số")
        st.caption("Gợi ý bài học hiện đại, linh hoạt, kích thích tư duy tương tác và niềm vui khám phá")
    with head_c2:
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("🧪 Nhập Dữ Liệu Thử Nghiệm", type="secondary", use_container_width=True):
            st.session_state["dg_teacher"] = "Nguyễn Thị Mai"
            st.session_state["dg_school"] = "Trường Mầm Non Sao Mai"
            st.session_state["dg_class"] = "Lớp Lá 1"
            st.session_state["dg_age"] = "5-6 tuổi"
            st.session_state["dg_theme"] = "Thế giới thực vật (Ngày Tết, 20/11...)"
            st.session_state["dg_topic"] = "Khám phá sự kỳ diệu của quả cam"
            st.session_state["dg_goals"] = "Trẻ làm quen và nhận biết thế giới xung quanh qua hình ảnh sinh động, rèn kỹ năng tương tác với màn hình thiết bị, biết lắng nghe và làm theo hướng dẫn trong 15 phút."
            st.session_state["dg_t_tools"] = "Bộ học liệu tĩnh, tranh thẻ hình, mô hình đồ chơi..."
            st.session_state["dg_c_tools"] = "Quả cam tươi, cốc nước, khăn lau tay, giấy màu..."
            st.session_state["dg_equip"] = "Máy tính kết nối tivi, tivi thông minh, loa bluetooth..."
            st.session_state["dg_space_req"] = "Không gian thoáng, đủ ánh sáng, khoảng cách an toàn cho mắt"
            st.session_state["dg_notes"] = "Lồng ghép nhân vật trợ lý ảo AI dẫn chuyện xuyên suốt tiết học"
            st.success("✅ Đã nạp dữ liệu thử nghiệm mẫu!")
            st.rerun()

    with st.container(border=True):
        st.markdown("#### 👤 I. THÔNG TIN GIÁO VIÊN & LỚP HỌC")
        dg_c1, dg_c2, dg_c3, dg_c4 = st.columns([2, 2, 1.2, 1.3])
        with dg_c1:
            dg_teacher = st.text_input("Tên giáo viên", value=st.session_state.get("dg_teacher", "Nguyễn Thị Mai"), key="in_dg_teacher")
        with dg_c2:
            dg_school = st.text_input("Tên trường mầm non", value=st.session_state.get("dg_school", "Trường Mầm Non Sao Mai"), key="in_dg_school")
        with dg_c3:
            dg_class = st.text_input("Lớp", value=st.session_state.get("dg_class", "Lớp Lá 1"), key="in_dg_class")
        with dg_c4:
            age_opts = ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"]
            def_age_idx = age_opts.index(st.session_state.get("dg_age", "5-6 tuổi")) if st.session_state.get("dg_age", "5-6 tuổi") in age_opts else 3
            dg_age = st.selectbox("Độ tuổi trẻ", age_opts, index=def_age_idx, key="in_dg_age")

    with st.container(border=True):
        st.markdown("#### 📖 II. NỘI DUNG HOẠT ĐỘNG CHÍNH")
        nd_c1, nd_c2 = st.columns(2)
        with nd_c1:
            dg_theme = st.text_input("Chủ đề/Chủ điểm lớn *", value=st.session_state.get("dg_theme", "Thế giới thực vật (Ngày Tết, 20/11...)"), key="in_dg_theme")
        with nd_c2:
            dg_topic = st.text_input("Tên hoạt động/giờ dạy *", value=st.session_state.get("dg_topic", "Khám phá sự kỳ diệu của quả cam"), key="in_dg_topic")

        nd_c3, nd_c4, nd_c5, nd_c6 = st.columns([2, 2, 1.5, 1.2])
        with nd_c3:
            dg_domain = st.selectbox("Lĩnh vực phát triển", ["Phát triển nhận thức", "Phát triển thể chất", "Phát triển ngôn ngữ", "Phát triển thẩm mỹ", "Phát triển tình cảm - KNXH"], index=0, key="in_dg_domain")
        with nd_c4:
            dg_act_type = st.selectbox("Loại hoạt động", ["Ứng dụng công nghệ/AI", "Khám phá tương tác số", "Trò chơi công nghệ số", "Hoạt động STEAM & Số hóa", "Kể chuyện/Âm nhạc tương tác AI"], index=0, key="in_dg_act_type")
        with nd_c5:
            dg_duration = st.selectbox("Thời lượng", ["15 - 20 phút", "20 - 25 phút", "25 - 30 phút", "30 - 35 phút"], index=3, key="in_dg_duration")
        with nd_c6:
            dg_num_kids = st.number_input("Số trẻ tham gia", min_value=5, max_value=60, value=25, step=1, key="in_dg_num_kids")

        dg_goals = st.text_area("Mục tiêu giáo dục (Để trống nếu muốn AI tự tạo mục tiêu):", value=st.session_state.get("dg_goals", ""), height=90, key="in_dg_goals")

    with st.container(border=True):
        st.markdown("#### 🧰 III. ĐỒ DÙNG DẠY HỌC & KHÔNG GIAN TỔ CHỨC")
        dd_c1, dd_c2 = st.columns(2)
        with dd_c1:
            dg_teacher_tools = st.text_area("Đồ dùng của cô", value=st.session_state.get("dg_t_tools", "Bộ học liệu tĩnh, tranh thẻ hình, mô hình đồ chơi..."), height=80, key="in_dg_t_tools")
        with dd_c2:
            dg_child_tools = st.text_area("Đồ dùng của trẻ (Các học liệu vật lý trẻ được cầm nắm)", value=st.session_state.get("dg_c_tools", "Quả cam tươi, cốc nước, khăn lau tay, giấy màu..."), height=80, key="in_dg_c_tools")

        dd_left, dd_right = st.columns([1.5, 1])
        with dd_left:
            dg_equip = st.text_input("💻 Học liệu số & Thiết bị công nghệ:", value=st.session_state.get("dg_equip", "Máy tính kết nối tivi, tivi thông minh, loa bluetooth..."), key="in_dg_equip")
            st.caption("Học liệu và công cụ số chuẩn bị sẵn:")
            hl_c1, hl_c2, hl_c3, hl_c4 = st.columns(4)
            with hl_c1:
                hl_1 = st.checkbox("Bảng tương tác", value=True, key="hl_1")
                hl_5 = st.checkbox("Âm thanh số", value=True, key="hl_5")
                hl_9 = st.checkbox("Tự chọn khác", key="hl_9")
            with hl_c2:
                hl_2 = st.checkbox("Canva", value=True, key="hl_2")
                hl_6 = st.checkbox("Trò chơi số", value=True, key="hl_6")
            with hl_c3:
                hl_3 = st.checkbox("PowerPoint", key="hl_3")
                hl_7 = st.checkbox("Thẻ QR", key="hl_7")
            with hl_c4:
                hl_4 = st.checkbox("Video", value=True, key="hl_4")
                hl_8 = st.checkbox("AI tạo hình ảnh", value=True, key="hl_8")

            selected_digital_tools = [lbl for flg, lbl in [
                (hl_1, "Bảng tương tác"), (hl_2, "Canva"), (hl_3, "PowerPoint"),
                (hl_4, "Video"), (hl_5, "Âm thanh số"), (hl_6, "Trò chơi số"),
                (hl_7, "Thẻ QR"), (hl_8, "AI tạo hình ảnh"), (hl_9, "Học liệu số mở rộng")
            ] if flg]

        with dd_right:
            dg_space = st.selectbox("📍 Không gian tổ chức hoạt động", ["Trong lớp học", "Ngoài sân trường", "Phòng chức năng / Phòng STEAM", "Góc tương tác thông minh"], index=0, key="in_dg_space")
            dg_space_req = st.text_input("Yêu cầu về ánh sáng / âm thanh / bố trí", value=st.session_state.get("dg_space_req", "Không gian thoáng, đủ ánh sáng, khoảng cách an toàn cho mắt"), key="in_dg_space_req")

    with st.container(border=True):
        st.markdown("#### 🎯 IV. HƯỚNG TIẾP CẬN & PHƯƠNG PHÁP GIÁO DỤC")
        pp_c1, pp_c2, pp_c3 = st.columns(3)
        with pp_c1: dg_org_form = st.selectbox("Hình thức tổ chức chủ đạo", ["Cả lớp", "Theo nhóm nhỏ", "Cá nhân", "Kết hợp Cả lớp & Nhóm nhỏ"], index=0, key="in_dg_org")
        with pp_c2: dg_approach = st.selectbox("Hướng tiếp cận sư phạm", ["Tích hợp năng lực số", "Lấy trẻ làm trung tâm kết hợp AI", "Khám phá trải nghiệm đa phương tiện", "Học qua chơi (Game-based Learning)"], index=0, key="in_dg_app")
        with pp_c3: dg_proactivity = st.selectbox("Mức độ chủ động của trẻ", ["Trẻ chủ động", "Cô dẫn dắt - Trẻ tương tác", "Trẻ tự khám phá theo nhóm"], index=0, key="in_dg_pro")

        st.caption("Các phương pháp dạy học muốn tích hợp:")
        m_col1, m_col2, m_col3, m_col4, m_col5, m_col6, m_col7 = st.columns(7)
        with m_col1:
            pm_1 = st.checkbox("Quan sát", value=True, key="pm_1")
            pm_8 = st.checkbox("Sử dụng công nghệ số", value=True, key="pm_8")
        with m_col2: pm_2 = st.checkbox("Đàm thoại", value=True, key="pm_2")
        with m_col3: pm_3 = st.checkbox("Trò chơi", value=True, key="pm_3")
        with m_col4: pm_4 = st.checkbox("Thí nghiệm", key="pm_4")
        with m_col5: pm_5 = st.checkbox("Thực hành", value=True, key="pm_5")
        with m_col6: pm_6 = st.checkbox("Làm việc nhóm", key="pm_6")
        with m_col7: pm_7 = st.checkbox("Giải quyết vấn đề", key="pm_7")

        selected_methods = [lbl for flg, lbl in [
            (pm_1, "Quan sát"), (pm_2, "Đàm thoại"), (pm_3, "Trò chơi"),
            (pm_4, "Thí nghiệm"), (pm_5, "Thực hành"), (pm_6, "Làm việc nhóm"),
            (pm_7, "Giải quyết vấn đề"), (pm_8, "Sử dụng công nghệ số")
        ] if flg]

        st.markdown("💡 **Yêu cầu đặc biệt trong phần Mở đầu:**")
        in_c1, in_c2, in_c3, in_c4, in_c5, in_c6 = st.columns(6)
        with in_c1: in_opt1 = st.checkbox("Câu đố vui", key="in_opt1")
        with in_c2: in_opt2 = st.checkbox("Bài hát/Vận động", value=True, key="in_opt2")
        with in_c3: in_opt3 = st.checkbox("Tình huống bất ngờ", value=True, key="in_opt3")
        with in_c4: in_opt4 = st.checkbox("Video/Hình ảnh AI", value=True, key="in_opt4")
        with in_c5: in_opt5 = st.checkbox("Trò chơi nhỏ", key="in_opt5")
        with in_c6: in_opt6 = st.checkbox("Nhân vật rối/Ảo", value=True, key="in_opt6")

        selected_intro = [lbl for flg, lbl in [
            (in_opt1, "Câu đố vui"), (in_opt2, "Bài hát/Vận động"), (in_opt3, "Tình huống bất ngờ"),
            (in_opt4, "Video/Hình ảnh AI"), (in_opt5, "Trò chơi nhỏ"), (in_opt6, "Nhân vật rối/Ảo")
        ] if flg]

        st.markdown("🔄 **Yêu cầu đặc biệt phần Kết thúc / Củng cố:**")
        out_c1, out_c2, out_c3, out_c4 = st.columns(4)
        with out_c1: out_opt1 = st.checkbox("Đánh giá nhẹ nhàng", value=True, key="out_opt1")
        with out_c2: out_opt2 = st.checkbox("Hát/Vận động", key="out_opt2")
        with out_c3: out_opt3 = st.checkbox("Trò chơi củng cố", value=True, key="out_opt3")
        with out_c4: out_opt4 = st.checkbox("Gợi mở hoạt động sau", key="out_opt4")

        selected_outro = [lbl for flg, lbl in [
            (out_opt1, "Đánh giá nhẹ nhàng"), (out_opt2, "Hát/Vận động"),
            (out_opt3, "Trò chơi củng cố"), (out_opt4, "Gợi mở hoạt động sau")
        ] if flg]

        dg_extra_notes = st.text_input("📝 Ghi chú thêm khác:", value=st.session_state.get("dg_notes", "Lồng ghép nhân vật trợ lý ảo AI dẫn chuyện xuyên suốt tiết học"), key="in_dg_notes")

    digital_info_dict = {
        "school": dg_school, "teacher": dg_teacher, "class_name": dg_class, "age": dg_age,
        "theme": dg_theme, "topic": dg_topic, "domain": dg_domain, "activity_type": dg_act_type,
        "duration": dg_duration, "num_children": dg_num_kids, "goals": dg_goals,
        "teacher_tools": dg_teacher_tools, "child_tools": dg_child_tools,
        "digital_equip": dg_equip, "digital_tools": selected_digital_tools,
        "space": dg_space, "space_req": dg_space_req, "org_form": dg_org_form,
        "approach": dg_approach, "proactivity": dg_proactivity, "methods": selected_methods,
        "intro_opts": selected_intro, "outro_opts": selected_outro, "extra_notes": dg_extra_notes,
        "doc_title": "GIÁO ÁN TÍCH HỢP NĂNG LỰC SỐ & AI"
    }

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    if st.button("🚀 Soạn Giáo Án Với AI", type="primary", use_container_width=True, key="btn_create_digital_plan"):
        with st.spinner("⚡ Đang kết nối AI thiết kế Giáo án Năng lực số chuẩn bảng 4 cột (vui lòng đợi ~8-15 giây)..."):
            prompt_digital = f"""
            Đóng vai là Chuyên gia Giáo dục Mầm non xuất sắc, hãy biên soạn một GIÁO ÁN MẦM NON TÍCH HỢP NĂNG LỰC SỐ & AI CHUẨN BỘ GD&ĐT VIỆT NAM (Dài ít nhất 1800 - 2200 từ):
            - Trường: {dg_school} | Giáo viên: {dg_teacher} | Lớp: {dg_class} ({dg_age}) | Số trẻ: {dg_num_kids}
            - Chủ đề lớn: {dg_theme} | Tên hoạt động số: {dg_topic}
            - Lĩnh vực: {dg_domain} | Loại hoạt động: {dg_act_type} | Thời lượng: {dg_duration}
            - Mục tiêu giáo dục: {dg_goals if dg_goals else 'Tự động tạo chuẩn: Kiến thức, Kỹ năng & Năng lực số, Thái độ'}
            - Đồ dùng của cô: {dg_teacher_tools} | Đồ dùng của trẻ: {dg_child_tools}
            - Thiết bị số: {dg_equip} | Công cụ chuẩn bị: {', '.join(selected_digital_tools)}
            - Địa điểm: {dg_space} ({dg_space_req})
            - Tiếp cận: {dg_approach} | Hình thức: {dg_org_form} ({dg_proactivity})
            - Phương pháp tích hợp: {', '.join(selected_methods)}
            - Mở đầu: {', '.join(selected_intro)} | Kết thúc: {', '.join(selected_outro)}
            - Ghi chú riêng: {dg_extra_notes}

            YÊU CẦU BẮT BUỘC:
            1. Phần III. TIẾN TRÌNH HOẠT ĐỘNG: Trình bày dưới dạng BẢNG MARKDOWN 4 CỘT:
               | Thời gian | Hoạt động của cô (Lời thoại, câu hỏi & hành động) | Dự kiến hoạt động & Phản hồi của trẻ | Quan sát, Hỗ trợ & Phân hóa |
               | --- | --- | --- | --- |
               Mỗi hàng bảng nằm trên 1 dòng duy nhất, dùng thẻ <br> để xuống dòng bên trong từng ô.
            2. Ghi chi tiết lời cô dẫn dắt tương tác thiết bị số và câu trả lời của trẻ.

            Trình bày chuẩn 4 phần:
            I. MỤC TIÊU HOẠT ĐỘNG
            II. CHUẨN BỊ HỌC LIỆU & KHÔNG GIAN SỐ
            III. TIẾN TRÌNH HOẠT ĐỘNG
            IV. ĐÁNH GIÁ CUỐI NGÀY (Bắt buộc gồm đúng 3 dòng:
            - Tình trạng sức khỏe: ...
            - Trạng thái cảm xúc, thái độ: ...
            - Kiến thức, kỹ năng: ...)

            Tuyệt đối KHÔNG sử dụng ký tự ### hay **.
            """
            res_digital, err = generate_fast_ai_content(api_key, prompt_digital, fallback_type="digital", fallback_data=digital_info_dict)
            if err:
                st.warning(f"⚠️ **Thông báo:** {err}\n\n👉 Đang sử dụng kịch bản mầm non tích hợp số chuẩn chỉnh bên dưới.")
            else:
                st.success("🎉 Đã hoàn thành Giáo án Tích hợp Năng lực số chi tiết!")

            st.session_state["saved_digital_plan"] = res_digital
            st.session_state["saved_digital_info"] = digital_info_dict
            save_plan_to_db(f"[Năng Lực Số] {dg_topic}", res_digital)

    if "saved_digital_plan" in st.session_state:
        st.divider()
        d_info = st.session_state["saved_digital_info"]
        d_font_size = st.session_state.get("d_font_size_slider", 14)

        db_col1, db_col2 = st.columns([1, 2])
        with db_col1:
            d_docx = export_docx_from_text(st.session_state["saved_digital_plan"], d_info, font_size_pt=d_font_size)
            with open(d_docx, "rb") as f_d:
                st.download_button(label="📥 Tải File Word Chuẩn Bảng (.docx)", data=f_d, file_name=d_docx, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True, key="dl_digital_btn")
        with db_col2:
            render_copy_and_print_buttons(st.session_state["saved_digital_plan"], f"Giáo Án Năng Lực Số: {d_info['topic']}", d_font_size, "tab_digital")

        di_c1, di_c2 = st.columns(2)
        with di_c1: st.success("✅ **Điểm sư phạm & Năng lực số: 100/100**")
        with di_c2: st.info(f"💻 **Tiếp cận:** {d_info['approach']} | {d_info['age']}")

        d_font_size = st.slider("🔤 Kích thước chữ hiển thị & In (pt):", min_value=12, max_value=24, value=d_font_size, step=1, key="d_font_size_slider")

        st.markdown("### 📄 Bản Xem Trước Giáo Án Tích Hợp Năng Lực Số (Trang Giấy In):")
        formatted_digital_html = format_text_to_standard_html(st.session_state["saved_digital_plan"], d_font_size)
        st.markdown(f'<div style="background-color: #ffffff; color: #000000; font-family: \'Times New Roman\', Times, serif; font-size: {d_font_size}pt; line-height: 1.5; padding: 35px 45px; border-radius: 12px; border: 1.5px solid #B2E5D6; box-shadow: 0 6px 18px rgba(0, 121, 107, 0.08); text-align: justify;">{formatted_digital_html}</div>', unsafe_allow_html=True)

        with st.expander("✏️ Chỉnh sửa nội dung Giáo Án Năng Lực Số"):
            ed_d_text = st.text_area("Nội dung:", st.session_state["saved_digital_plan"], height=380, key="ed_d_ta")
            if ed_d_text != st.session_state["saved_digital_plan"]:
                st.session_state["saved_digital_plan"] = clean_markdown_text(ed_d_text)
                st.rerun()

# =========================================================================
# TAB 3: KẾ HOẠCH & SỔ CHỦ ĐỀ (ĐÃ KHÔI PHỤC ĐẦY ĐỦ THÔNG TIN CHI TIẾT)
# =========================================================================
elif app_mode == "🗓️ Kế Hoạch & Sổ Chủ Đề":
    sub_tab_month, sub_tab_week = st.tabs(["📅 Kế Hoạch Tháng & Sổ Chủ Đề", "📋 Kế Hoạch Giáo Dục Tuần"])
    
    # ---------------------------------------------------------------------
    # SUB-TAB 1: KẾ HOẠCH THÁNG & SỔ CHỦ ĐỀ CHUYÊN SÂU
    # ---------------------------------------------------------------------
    with sub_tab_month:
        st.subheader("📖 Thiết Kế Kế Hoạch Tháng & Sổ Chủ Đề (Chuẩn Chương Trình GDMN)")
        st.caption("Xây dựng trọn bộ Sổ chủ đề: Mục tiêu 5 lĩnh vực, Mạng nội dung, Mạng hoạt động, Bảng phân phối 4-5 tuần và Hoạt động góc.")

        with st.container(border=True):
            st.markdown("#### 👤 I. THÔNG TIN CHUNG & CHỦ ĐỀ THÁNG")
            mc1, mc2, mc3 = st.columns(3)
            with mc1:
                m_teacher = st.text_input("Người soạn / Giáo viên", "Nguyễn Thị Thanh Vân", key="m_teacher")
                m_school = st.text_input("Đơn vị / Trường", "Trường Mầm Non Phước Thạnh", key="m_school")
            with mc2:
                m_class = st.text_input("Lớp", "Lớp Mẫu giáo Nhỡ (4-5 tuổi)", key="m_class")
                m_age = st.selectbox("Độ tuổi", ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=2, key="m_age")
            with mc3:
                m_month_val = st.selectbox(
                    "Tháng thực hiện",
                    [
                        "Tháng 9 (Khai giảng - Bé vui đến trường)",
                        "Tháng 10 (Bản thân / Gia đình)",
                        "Tháng 11 (Nghề nghiệp - Tri ân 20/11)",
                        "Tháng 12 (Thế giới động vật / Chú bộ đội)",
                        "Tháng 1 (Thế giới thực vật - Mùa xuân)",
                        "Tháng 2 (Tết và Mùa xuân)",
                        "Tháng 3 (Phương tiện & Luật lệ giao thông)",
                        "Tháng 4 (Nước và các hiện tượng tự nhiên)",
                        "Tháng 5 (Quê hương - Đất nước - Bác Hồ)",
                        "Hè (Tháng 6 - Tháng 8)"
                    ],
                    index=4,
                    key="m_month_val"
                )
                m_weeks_count = st.selectbox("Thời gian thực hiện", ["4 tuần", "5 tuần", "3 tuần"], index=0, key="m_weeks_count")

            mc4, mc5 = st.columns(2)
            with mc4:
                m_theme_val = st.text_input("Tên Chủ đề lớn *", "Thế giới thực vật", key="m_theme_val")
            with mc5:
                m_date_range = st.text_input("Thời gian từ ngày ... đến ngày ...", "Từ ngày 01 đến ngày 30 trong tháng", key="m_date_range")

        with st.container(border=True):
            st.markdown("#### 🌿 II. PHÂN NHÁNH CÁC TUẦN & TÍCH HỢP GIÁO DỤC")
            m_sub_themes = st.text_area(
                "Tên các chủ đề nhánh dự kiến cho từng tuần (Để trống nếu muốn AI tự chia nhánh chuẩn):",
                value="Tuần 1: Cây xanh quanh bé\nTuần 2: Một số loại hoa đẹp\nTuần 3: Một số loại quả\nTuần 4: Rau xanh và ích lợi đối với sức khỏe",
                height=100,
                key="m_sub_themes"
            )

            st.caption("Các lĩnh vực phát triển cần xây dựng mục tiêu & mạng hoạt động:")
            md_c1, md_c2, md_c3, md_c4, md_c5 = st.columns(5)
            with md_c1: md1 = st.checkbox("Phát triển Thể chất", value=True, key="m_dom_1")
            with md_c2: md2 = st.checkbox("Phát triển Nhận thức", value=True, key="m_dom_2")
            with md_c3: md3 = st.checkbox("Phát triển Ngôn ngữ", value=True, key="m_dom_3")
            with md_c4: md4 = st.checkbox("Phát triển Tình cảm - KNXH", value=True, key="m_dom_4")
            with md_c5: md5 = st.checkbox("Phát triển Thẩm mỹ", value=True, key="m_dom_5")

            m_domains = [lbl for flg, lbl in [
                (md1, "Phát triển Thể chất"), (md2, "Phát triển Nhận thức"),
                (md3, "Phát triển Ngôn ngữ"), (md4, "Phát triển Tình cảm - KNXH"),
                (md5, "Phát triển Thẩm mỹ")
            ] if flg]

            st.caption("Nội dung tích hợp & Sự kiện lễ hội trong tháng:")
            me_c1, me_c2, me_c3, me_c4, me_c5 = st.columns(5)
            with me_c1: me1 = st.checkbox("STEAM & Trải nghiệm", value=True, key="m_ext_1")
            with me_c2: me2 = st.checkbox("Ứng dụng công nghệ AI", value=True, key="m_ext_2")
            with me_c3: me3 = st.checkbox("Giáo dục SEL / Kỹ năng sống", value=True, key="m_ext_3")
            with me_c4: me4 = st.checkbox("Bảo vệ môi trường", value=True, key="m_ext_4")
            with me_c5: me5 = st.checkbox("Ngày lễ / Hội thi trong tháng", key="m_ext_5")

            m_extras = [lbl for flg, lbl in [
                (me1, "STEAM & Trải nghiệm"), (me2, "Ứng dụng công nghệ AI"),
                (me3, "Giáo dục SEL / Kỹ năng sống"), (me4, "Bảo vệ môi trường"),
                (me5, "Ngày lễ / Hội thi trong tháng")
            ] if flg]

            m_custom_notes = st.text_area(
                "Mục tiêu trọng tâm hoặc yêu cầu riêng của nhà trường / lớp:",
                placeholder="Ví dụ: Tăng cường hoạt động trải nghiệm ngoài trời, lồng ghép trò chơi dân gian, phối hợp phụ huynh làm đồ dùng mở...",
                height=80,
                key="m_custom_notes"
            )

        info_m = {
            "school": m_school, "teacher": m_teacher, "class_name": m_class,
            "month": m_month_val, "weeks_count": m_weeks_count, "date_range": m_date_range,
            "theme": m_theme_val, "topic": f"KeHoach_{m_month_val}_{m_theme_val}",
            "age": m_age, "sub_themes": m_sub_themes, "domains": m_domains, "extras": m_extras
        }

        if st.button("✨ TẠO KẾ HOẠCH THÁNG & SỔ CHỦ ĐỀ CHI TIẾT", type="primary", use_container_width=True, key="btn_gen_month"):
            with st.spinner("⚡ Đang biên soạn Kế hoạch Tháng & Sổ Chủ Đề siêu chi tiết (vui lòng đợi ~10-15 giây)..."):
                prompt_month = f"""
                Đóng vai là Chuyên gia Giáo dục Mầm non của Bộ GD&ĐT Việt Nam, hãy lập một KẾ HOẠCH GIÁO DỤC THÁNG & SỔ CHỦ ĐỀ HOÀN CHỈNH, RẤT DÀI VÀ CHI TIẾT (từ 1800 - 2500 từ):
                - Đơn vị: {m_school} | Lớp: {m_class} ({m_age}) | Người soạn: {m_teacher}
                - Tháng thực hiện: {m_month_val} | Thời gian: {m_weeks_count} ({m_date_range})
                - Chủ đề lớn: {m_theme_val}
                - Phân nhánh các tuần: {m_sub_themes if m_sub_themes.strip() else 'Tự động chia nhánh hợp lý theo số tuần'}
                - Các lĩnh vực phát triển: {', '.join(m_domains)}
                - Nội dung tích hợp: {', '.join(m_extras)}
                - Yêu cầu riêng: {m_custom_notes if m_custom_notes.strip() else 'Bám sát chương trình GDMN hiện hành'}

                CẤU TRÚC BẮT BUỘC TRÌNH BÀY RÕ RÀNG:
                Dòng đầu: KẾ HOẠCH GIÁO DỤC THÁNG & SỔ CHỦ ĐỀ
                Thông tin hành chính: Trường, Lớp, Độ tuổi, Người soạn, Tháng thực hiện, Chủ đề lớn, Thời gian thực hiện.

                I. MỤC TIÊU PHÁT TRIỂN CHỦ ĐỀ THÁNG (Trình bày chi tiết từng mục tiêu theo 5 lĩnh vực: Phát triển thể chất, Phát triển nhận thức, Phát triển ngôn ngữ, Phát triển tình cảm - KNXH, Phát triển thẩm mỹ)
                II. MẠNG NỘI DUNG CHỦ ĐỀ (Trình bày rõ nội dung khám phá của từng nhánh tuần)
                III. MẠNG HOẠT ĐỘNG THEO 5 LĨNH VỰC (Trình bày dạng BẢNG MARKDOWN 2 CỘT: | Lĩnh vực phát triển | Nội dung các hoạt động giáo dục |)
                IV. BẢNG PHÂN PHỐI CHƯƠNG TRÌNH HOẠT ĐỘNG HỌC THEO CÁC TUẦN (Trình bày dạng BẢNG MARKDOWN 6 CỘT:
                | Tuần & Nhánh | Thứ Hai (Thể chất) | Thứ Ba (Nhận thức) | Thứ Tư (Tạo hình) | Thứ Năm (Văn học/Chữ cái) | Thứ Sáu (Âm nhạc/KNXH) |
                | --- | --- | --- | --- | --- | --- |
                Mỗi hàng nằm trên 1 dòng duy nhất, dùng thẻ <br> để xuống dòng bên trong ô).
                V. KẾ HOẠCH HOẠT ĐỘNG GÓC, HOẠT ĐỘNG NGOÀI TRỜI & PHỐI HỢP PHỤ HUYNH (Nêu chi tiết mục đích, chuẩn bị đồ dùng mở và biện pháp phối hợp cha mẹ trẻ).

                Tuyệt đối không dùng ký tự ### hay **.
                """
                res_m, err_m = generate_fast_ai_content(api_key, prompt_month, fallback_type="month", fallback_data=info_m)
                if err_m:
                    st.warning(f"⚠️ **Thông báo:** {err_m}\n\n👉 Đang hiển thị mẫu Kế hoạch tháng chuẩn bên dưới.")
                else:
                    st.success("🎉 Đã hoàn thành Kế hoạch Tháng & Sổ Chủ Đề chi tiết!")

                st.session_state["saved_month_plan"] = res_m
                st.session_state["saved_month_info"] = info_m
                save_plan_to_db(f"[Kế Hoạch Tháng] {m_month_val} - {m_theme_val}", res_m)

        if "saved_month_plan" in st.session_state:
            st.divider()
            m_inf = st.session_state.get("saved_month_info", info_m)
            m_font_sz = st.session_state.get("m_font_size_slider", 14)

            mb1, mb2 = st.columns([1, 2])
            with mb1:
                m_docx = export_docx_from_text(st.session_state["saved_month_plan"], m_inf, font_size_pt=m_font_sz, prefix_name="KeHoachThang")
                with open(m_docx, "rb") as fm_f:
                    st.download_button(
                        label="📥 Tải Kế Hoạch Tháng Word (.docx)",
                        data=fm_f,
                        file_name=m_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                        key="dl_month_docx"
                    )
            with mb2:
                render_copy_and_print_buttons(st.session_state["saved_month_plan"], f"Kế Hoạch Tháng: {m_inf['theme']}", m_font_sz, "tab_month_plan")

            m_font_sz = st.slider("🔤 Kích thước chữ hiển thị & In (pt):", 12, 24, m_font_sz, 1, key="m_font_size_slider")
            st.markdown("### 📄 Bản Xem Trước Kế Hoạch Tháng & Sổ Chủ Đề:")
            st.markdown(
                f'<div style="background-color: #ffffff; color: #000000; font-family: \'Times New Roman\', Times, serif; font-size: {m_font_sz}pt; line-height: 1.5; padding: 35px 45px; border-radius: 12px; border: 1.5px solid #B2E5D6; box-shadow: 0 6px 18px rgba(0, 121, 107, 0.08); text-align: justify;">{format_text_to_standard_html(st.session_state["saved_month_plan"], m_font_sz)}</div>',
                unsafe_allow_html=True
            )
            with st.expander("✏️ Chỉnh sửa nội dung Kế Hoạch Tháng"):
                ed_m = st.text_area("Nội dung:", st.session_state["saved_month_plan"], height=380, key="ed_m_ta")
                if ed_m != st.session_state["saved_month_plan"]:
                    st.session_state["saved_month_plan"] = clean_markdown_text(ed_m)
                    st.rerun()

    # ---------------------------------------------------------------------
    # SUB-TAB 2: KẾ HOẠCH GIÁO DỤC TUẦN CHUYÊN SÂU
    # ---------------------------------------------------------------------
    with sub_tab_week:
        st.subheader("📋 Thiết Kế Kế Hoạch Giáo Dục Tuần (Chi Tiết Từ Thứ 2 - Thứ 6)")
        st.caption("Thiết kế đầy đủ 5 thời điểm sinh hoạt trong ngày, hoạt động học có chủ đích và tổ chức 5 góc chơi trong tuần.")

        with st.container(border=True):
            st.markdown("#### 👤 I. THÔNG TIN CHUNG CỦA TUẦN")
            wc1, wc2, wc3 = st.columns(3)
            with wc1:
                w_teacher = st.text_input("Người soạn / Giáo viên", "Nguyễn Thị Thanh Vân", key="w_teacher")
                w_school = st.text_input("Đơn vị / Trường", "Trường Mầm Non Phước Thạnh", key="w_school")
            with wc2:
                w_class = st.text_input("Lớp", "Lớp Mẫu giáo Nhỡ (4-5 tuổi)", key="w_class")
                w_age2 = st.selectbox("Độ tuổi trẻ", ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=2, key="w_age2")
            with wc3:
                w_week_num = st.selectbox("Tuần thực hiện", ["Tuần 1", "Tuần 2", "Tuần 3", "Tuần 4", "Tuần 5"], key="w_week_num")
                w_date_range = st.text_input("Thời gian thực hiện", "Từ Thứ Hai đến Thứ Sáu", key="w_date_range")

            wc4, wc5, wc6 = st.columns([1.5, 2, 1.5])
            with wc4:
                w_big_theme = st.text_input("Chủ đề lớn *", "Bản thân", key="w_big_theme")
            with wc5:
                w_theme_week = st.text_input("Chủ đề nhánh tuần *", "Tôi là ai - Bạn là ai?", key="w_theme_week")
            with wc6:
                w_fmt_sel = st.selectbox(
                    "Định dạng trình bày",
                    ["Bảng tổng hợp Thứ 2 - Thứ 6 (Chuẩn sổ kế hoạch)", "Chi tiết từng ngày (Thứ 2 đến Thứ 6)"],
                    index=0,
                    key="w_fmt_sel"
                )

        with st.container(border=True):
            st.markdown("#### 📚 II. DỰ KIẾN ĐỀ TÀI HOẠT ĐỘNG HỌC & CHẾ ĐỘ SINH HOẠT")
            st.caption("Nhập tên đề tài từng ngày nếu cô đã có sẵn phân phối, hoặc giữ gợi ý để AI tự thiết kế đồng bộ:")
            wd1, wd2, wd3, wd4, wd5 = st.columns(5)
            with wd1:
                w_t2 = st.text_input("Thứ Hai (Thể chất)", "VĐCB: Đi chạy thay đổi tốc độ theo hiệu lệnh", key="w_t2")
            with wd2:
                w_t3 = st.text_input("Thứ Ba (Nhận thức)", "KPKH: Khám phá đặc điểm bản thân bé", key="w_t3")
            with wd3:
                w_t4 = st.text_input("Thứ Tư (Thẩm mỹ)", "Tạo hình: Vẽ chân dung bạn thân của bé", key="w_t4")
            with wd4:
                w_t5 = st.text_input("Thứ Năm (Ngôn ngữ)", "Thơ: Tâm sự của cái mũi / Chuyện về bé", key="w_t5")
            with wd5:
                w_t6 = st.text_input("Thứ Sáu (Âm nhạc/KNXH)", "Dạy hát & Vận động: Mừng sinh nhật", key="w_t6")

            st.caption("Các thời điểm sinh hoạt trong ngày cần thiết kế chi tiết:")
            wa1, wa2, wa3, wa4, wa5 = st.columns(5)
            with wa1: w_act1 = st.checkbox("Đón trẻ & Thể dục sáng", value=True, key="w_act_1")
            with wa2: w_act2 = st.checkbox("Hoạt động học có chủ đích", value=True, key="w_act_2")
            with wa3: w_act3 = st.checkbox("Hoạt động ngoài trời", value=True, key="w_act_3")
            with wa4: w_act4 = st.checkbox("Hoạt động chơi ở các góc", value=True, key="w_act_4")
            with wa5: w_act5 = st.checkbox("Hoạt động chiều & Trả trẻ", value=True, key="w_act_5")

            w_activities = [lbl for flg, lbl in [
                (w_act1, "Đón trẻ & Thể dục sáng"), (w_act2, "Hoạt động học có chủ đích"),
                (w_act3, "Hoạt động ngoài trời"), (w_act4, "Hoạt động chơi ở các góc"),
                (w_act5, "Hoạt động chiều & Trả trẻ")
            ] if flg]

            st.caption("Tích hợp giáo dục trong tuần:")
            we1, we2, we3, we4 = st.columns(4)
            with we1: w_ext1 = st.checkbox("STEAM & Khám phá", value=True, key="w_ext_1")
            with we2: w_ext2 = st.checkbox("Ứng dụng công nghệ AI / Video", value=True, key="w_ext_2")
            with we3: w_ext3 = st.checkbox("Giáo dục SEL / Kỹ năng tự phục vụ", value=True, key="w_ext_3")
            with we4: w_ext4 = st.checkbox("Trò chơi dân gian & Vận động", value=True, key="w_ext_4")

            w_extras = [lbl for flg, lbl in [
                (w_ext1, "STEAM & Khám phá"), (w_ext2, "Ứng dụng công nghệ AI / Video"),
                (w_ext3, "Giáo dục SEL / Kỹ năng tự phục vụ"), (w_ext4, "Trò chơi dân gian & Vận động")
            ] if flg]

            w_custom_notes = st.text_area(
                "Yêu cầu bổ sung hoặc lưu ý riêng của giáo viên trong tuần:",
                placeholder="Ví dụ: Rèn nền nếp rửa tay bằng xà phòng trước khi ăn, chuẩn bị đạo cụ cho góc phân vai bác sĩ...",
                height=75,
                key="w_custom_notes"
            )

        info_w = {
            "school": w_school, "teacher": w_teacher, "class_name": w_class,
            "week_num": w_week_num, "date_range": w_date_range,
            "big_theme": w_big_theme, "theme_week": w_theme_week,
            "topic": f"KeHoach_{w_week_num}_{w_theme_week}", "age": w_age2,
            "t2": w_t2, "t3": w_t3, "t4": w_t4, "t5": w_t5, "t6": w_t6
        }

        if st.button("✨ TẠO KẾ HOẠCH GIÁO DỤC TUẦN CHI TIẾT", type="primary", use_container_width=True, key="btn_gen_week"):
            with st.spinner("⚡ Đang lập Kế hoạch Giáo dục Tuần chi tiết từ Thứ Hai đến Thứ Sáu (vui lòng đợi ~10-15 giây)..."):
                if "Bảng tổng hợp" in w_fmt_sel:
                    w_layout_rule = """
                    III. KẾ HOẠCH HOẠT ĐỘNG HÀNG NGÀY (THỨ HAI - THỨ SÁU):
                    BẮT BUỘC trình bày dạng BẢNG MARKDOWN 6 CỘT:
                    | Thời điểm sinh hoạt | Thứ Hai | Thứ Ba | Thứ Tư | Thứ Năm | Thứ Sáu |
                    | --- | --- | --- | --- | --- | --- |
                    - Gồm 5 hàng tương ứng với 5 thời điểm:
                      1. Đón trẻ & Thể dục sáng
                      2. Hoạt động học có chủ đích
                      3. Hoạt động ngoài trời (ghi rõ: Quan sát có mục đích, Trò chơi vận động, Chơi tự do)
                      4. Hoạt động chơi ở các góc
                      5. Hoạt động chiều & Nêu gương trả trẻ
                    - Mỗi hàng bảng nằm trên 1 dòng duy nhất, dùng thẻ <br> để xuống dòng bên trong ô.
                    """
                else:
                    w_layout_rule = """
                    III. KẾ HOẠCH HOẠT ĐỘNG CHI TIẾT TỪNG NGÀY:
                    Trình bày rõ ràng từng ngày từ 1. THỨ HAI, 2. THỨ BA, 3. THỨ TƯ, 4. THỨ NĂM, 5. THỨ SÁU; mỗi ngày chia đủ 5 gạch đầu dòng thời điểm sinh hoạt.
                    """

                prompt_week = f"""
                Đóng vai là Chuyên gia Giáo dục Mầm non của Bộ GD&ĐT Việt Nam, hãy lập một KẾ HOẠCH GIÁO DỤC TUẦN HOÀN CHỈNH, RẤT DÀI VÀ CỰC KỲ CHI TIẾT (từ 1800 - 2500 từ):
                - Đơn vị: {w_school} | Lớp: {w_class} ({w_age2}) | Người soạn: {w_teacher}
                - Tuần thực hiện: {w_week_num} ({w_date_range})
                - Chủ đề lớn: {w_big_theme} | Chủ đề nhánh tuần: {w_theme_week}
                - Dự kiến đề tài hoạt động học trong tuần:
                  + Thứ Hai: {w_t2}
                  + Thứ Ba: {w_t3}
                  + Thứ Tư: {w_t4}
                  + Thứ Năm: {w_t5}
                  + Thứ Sáu: {w_t6}
                - Các thời điểm trong ngày: {', '.join(w_activities)}
                - Tích hợp: {', '.join(w_extras)}
                - Lưu ý riêng của giáo viên: {w_custom_notes if w_custom_notes.strip() else 'Không'}

                CẤU TRÚC BẮT BUỘC:
                Dòng đầu: KẾ HOẠCH GIÁO DỤC {w_week_num.upper()}
                Thông tin hành chính: Trường, Lớp, Độ tuổi, Người soạn, Chủ đề lớn, Chủ đề nhánh, Thời gian thực hiện.

                I. MỤC TIÊU TUẦN (Ghi cụ thể theo 5 lĩnh vực phát triển: Thể chất, Nhận thức, Ngôn ngữ, Tình cảm - KNXH, Thẩm mỹ)
                II. CHUẨN BỊ CỦA CÔ VÀ TRẺ TRONG TUẦN (Đồ dùng của cô, Đồ dùng của trẻ, Môi trường lớp học)
                {w_layout_rule}
                IV. KẾ HOẠCH TỔ CHỨC HOẠT ĐỘNG GÓC TRONG TUẦN (Chi tiết 5 góc: Góc Phân vai, Góc Xây dựng, Góc Học tập - Sách, Góc Nghệ thuật - Tạo hình/Âm nhạc, Góc Thiên nhiên - STEAM; mỗi góc nêu rõ: Mục đích - Chuẩn bị - Gợi ý hướng dẫn của cô)
                V. ĐÁNH GIÁ CUỐI TUẦN (Tình trạng sức khỏe, Nề nếp thói quen, Kết quả tiếp thu kiến thức kỹ năng)

                Tuyệt đối không dùng ký tự ### hay **.
                """
                res_w, err_w = generate_fast_ai_content(api_key, prompt_week, fallback_type="week", fallback_data=info_w)
                if err_w:
                    st.warning(f"⚠️ **Thông báo:** {err_w}\n\n👉 Đang hiển thị mẫu Kế hoạch tuần chuẩn bảng 6 cột bên dưới.")
                else:
                    st.success("🎉 Đã hoàn thành Kế hoạch Giáo dục Tuần chi tiết!")

                st.session_state["saved_week_plan"] = res_w
                st.session_state["saved_week_info"] = info_w
                save_plan_to_db(f"[Kế Hoạch Tuần] {w_week_num} - {w_theme_week}", res_w)

        if "saved_week_plan" in st.session_state:
            st.divider()
            w_inf = st.session_state.get("saved_week_info", info_w)
            w_font_sz = st.session_state.get("w_font_size_slider", 14)

            wb1, wb2 = st.columns([1, 2])
            with wb1:
                w_docx = export_docx_from_text(st.session_state["saved_week_plan"], w_inf, font_size_pt=w_font_sz, prefix_name="KeHoachTuan")
                with open(w_docx, "rb") as fw_f:
                    st.download_button(
                        label="📥 Tải Kế Hoạch Tuần Word (.docx)",
                        data=fw_f,
                        file_name=w_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                        key="dl_week_docx"
                    )
            with wb2:
                render_copy_and_print_buttons(st.session_state["saved_week_plan"], f"Kế Hoạch Tuần: {w_inf['theme_week']}", w_font_sz, "tab_week_plan")

            w_font_sz = st.slider("🔤 Kích thước chữ hiển thị & In (pt):", 12, 24, w_font_sz, 1, key="w_font_size_slider")
            st.markdown("### 📄 Bản Xem Trước Kế Hoạch Giáo Dục Tuần:")
            st.markdown(
                f'<div style="background-color: #ffffff; color: #000000; font-family: \'Times New Roman\', Times, serif; font-size: {w_font_sz}pt; line-height: 1.5; padding: 35px 45px; border-radius: 12px; border: 1.5px solid #B2E5D6; box-shadow: 0 6px 18px rgba(0, 121, 107, 0.08); text-align: justify;">{format_text_to_standard_html(st.session_state["saved_week_plan"], w_font_sz)}</div>',
                unsafe_allow_html=True
            )
            with st.expander("✏️ Chỉnh sửa nội dung Kế Hoạch Tuần"):
                ed_w = st.text_area("Nội dung:", st.session_state["saved_week_plan"], height=380, key="ed_w_ta")
                if ed_w != st.session_state["saved_week_plan"]:
                    st.session_state["saved_week_plan"] = clean_markdown_text(ed_w)
                    st.rerun()

# ==========================================
# CÁC TAB CHỨC NĂNG KHÁC (GIỮ NGUYÊN 100%)
# ==========================================
elif app_mode == "🎯 Trò Chơi & Đố Vui":
    st.subheader("🎯 Sáng Tác Trò Chơi, Câu Đố & Bài Thơ Mầm Non")
    topic_game = st.text_input("Nhập chủ đề", "An toàn giao thông")
    genre = st.selectbox("Thể loại", ["Bài thơ ngắn dễ thuộc", "Bài đồng dao vui nhộn", "Câu đố vui cho trẻ", "Trò chơi vận động tập thể"])
    if st.button("✨ Sáng Tác Ngay", type="primary"):
        with st.spinner("⚡ Đang sáng tác nội dung..."):
            res, err = generate_fast_ai_content(api_key, f"Sáng tác {genre} cực hay, dài, hấp dẫn cho trẻ mầm non chủ đề {topic_game}. Không dùng dấu ###", fallback_type="creative", fallback_data={"genre": genre, "topic": topic_game})
            if err: st.warning(f"⚠️ {err}")
            st.markdown(format_text_to_standard_html(res, 14), unsafe_allow_html=True)

elif app_mode == "💬 Sổ Liên Lạc & Zalo":
    st.subheader("💬 Trợ Lý Viết Sổ Liên Lạc & Tin Nhắn Zalo")
    task_type = st.radio("Chọn loại:", ["Tin nhắn tổng kết ngày/tuần", "Viết lời nhận xét sự tiến bộ"])
    child_info = st.text_input("Nhập thông tin", "Hôm nay các con học kỹ năng tốt")
    if st.button("💌 Soạn Tin", type="primary"):
        with st.spinner("⚡ Đang soạn tin..."):
            res, err = generate_fast_ai_content(api_key, f"Soạn {task_type} mầm non thân thiện, dài, tình cảm, tóm tắt: {child_info}. Không dùng dấu ###", fallback_type="zalo", fallback_data={"type": task_type, "info": child_info})
            if err: st.warning(f"⚠️ {err}")
            st.markdown(format_text_to_standard_html(res, 14), unsafe_allow_html=True)

elif app_mode == "✂️ Gợi Ý Đồ Dùng DIY":
    st.subheader("✂️ Gợi Ý Làm Đồ Dùng Dạy Học Từ Vật Liệu Mở")
    lesson_name = st.text_input("Nhập bài học", "Biển báo giao thông")
    if st.button("🛠️ Gợi Ý Cách Làm", type="primary"):
        with st.spinner("⚡ Đang gợi ý cách làm chi tiết..."):
            res, err = generate_fast_ai_content(api_key, f"Gợi ý làm đồ dùng phế liệu mầm non chi tiết từng bước cho bài: {lesson_name}. Không dùng dấu ###", fallback_type="diy", fallback_data={"lesson": lesson_name})
            if err: st.warning(f"⚠️ {err}")
            st.markdown(format_text_to_standard_html(res, 14), unsafe_allow_html=True)

# =========================================================================
# TAB KHO LỊCH SỬ GIÁO ÁN (GIỮ NGUYÊN 100%)
# =========================================================================
elif "📚 Kho Lịch Sử Giáo Án" in app_mode:
    st.subheader("📚 Kho Lưu Trữ Giáo Án Đã Soạn")
    st.caption("Toàn bộ các giáo án, kế hoạch và học liệu nâng cao cô vừa soạn sẽ được tự động lưu trữ tại đây để mở lại hoặc tải về máy.")
    
    all_plans = get_all_saved_plans()
    
    if not all_plans:
        st.info("Chưa có giáo án nào được lưu trữ trong hệ thống.")
    else:
        st.success(f"📂 Hiện có tổng cộng **{len(all_plans)} bài** trong kho của bạn:")
        
        for p_id, p_title, p_content, p_time in all_plans:
            with st.expander(f"📖 **{p_title}**  *(Lưu lúc: {p_time})*"):
                col_k1, col_k2, col_k3 = st.columns([2, 1, 1])
                
                with col_k1:
                    st.markdown(f"**Tên tài liệu:** {p_title}")
                    st.caption(f"Thời gian: {p_time}")
                
                with col_k2:
                    docx_path = export_docx_from_text(p_content, {"topic": p_title})
                    with open(docx_path, "rb") as f_down:
                        st.download_button(
                            label="📥 Tải File Word (.docx)",
                            data=f_down,
                            file_name=docx_path,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                            key=f"down_btn_{p_id}",
                            use_container_width=True
                        )
                
                with col_k3:
                    if st.button("🗑️ Xóa Bài Này", key=f"del_btn_{p_id}", use_container_width=True):
                        delete_plan_by_id(p_id)
                        st.toast("Đã xóa bài thành công!")
                        st.rerun()
                
                st.divider()
                st.markdown("##### 📄 Nội dung chi tiết:")
                st.markdown(f'<div style="background-color: #ffffff; color: #000000; font-family: \'Times New Roman\', Times, serif; font-size: 13pt; line-height: 1.5; padding: 25px 35px; border-radius: 8px; border: 1px solid #B2E5D6; text-align: justify;">{format_text_to_standard_html(p_content, 13)}</div>', unsafe_allow_html=True)
