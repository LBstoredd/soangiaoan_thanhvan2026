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
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pydantic import BaseModel
from datetime import datetime, timedelta
import random
import string

# ==========================================
# CẤU HÌNH GIAO DIỆN & KHỞI TẠO HỆ THỐNG
# ==========================================
st.set_page_config(page_title="Phần Mềm Trợ Lý Mầm Non Chuyên Sâu", layout="wide")

# KHỞI TẠO BIẾN HỆ THỐNG (Chống lỗi KeyError)
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "is_admin" not in st.session_state:
    st.session_state["is_admin"] = False
if "api_key_saved" not in st.session_state:
    st.session_state["api_key_saved"] = ""

# THEME XANH MINT HIỆN ĐẠI
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(180deg, #F4FBF9 0%, #FFFFFF 100%);
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        color: #1E293B;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #E6F7F2 0%, #D4F1E8 100%) !important;
        border-right: 1.5px solid #B2E5D6;
    }
    [data-testid="stSidebar"] * {
        color: #0F5247 !important;
    }
    h1 {
        color: #00796B !important;
        font-weight: 800 !important;
        letter-spacing: -0.3px;
        text-shadow: 0 2px 4px rgba(0, 121, 107, 0.08);
    }
    h2, h3, h4, h5 {
        color: #00695C !important;
        font-weight: 700 !important;
    }
    div.stButton > button[kind="primary"],
    div.stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, #00B894 0%, #00897B 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 10px 24px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 12px rgba(0, 184, 148, 0.28) !important;
        transition: all 0.25s ease !important;
    }
    div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0, 184, 148, 0.42) !important;
        background: linear-gradient(135deg, #00CBA4 0%, #00796B 100%) !important;
    }
    div.stButton > button[kind="secondary"] {
        background-color: #FFFFFF !important;
        color: #00796B !important;
        border: 1.5px solid #80CBC4 !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        padding: 8px 18px !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button[kind="secondary"]:hover {
        background-color: #E0F2F1 !important;
        border-color: #00897B !important;
        transform: translateY(-1px);
    }
    .stTextInput > div > div > input, 
    .stTextArea > div > div > textarea,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div {
        border-radius: 10px !important;
        border: 1.5px solid #B2DFDB !important;
        background-color: #FFFFFF !important;
        color: #1E293B !important;
        transition: all 0.25s ease;
    }
    .stTextInput > div > div > input:focus, 
    .stTextArea > div > div > textarea:focus,
    .stNumberInput > div > div > input:focus {
        border-color: #00B894 !important;
        box-shadow: 0 0 0 3px rgba(0, 184, 148, 0.18) !important;
    }
    [data-testid="stExpander"], div[data-testid="stVerticalBlock"] > div[style*="border"] {
        border-radius: 14px !important;
        border: 1px solid #B2E5D6 !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 4px 14px rgba(0, 121, 107, 0.04) !important;
    }
    button[data-baseweb="tab"] {
        font-weight: 700 !important;
        font-size: 15px !important;
        color: #00796B !important;
    }
    .stSlider > div > div > div > div {
        background-color: #00B894 !important;
    }
    div[data-testid="stAlert"] {
        border-radius: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# BỘ LỌC VÀ ĐỊNH DẠNG VĂN BẢN 
# ==========================================
def clean_markdown_text(raw_text: str) -> str:
    if not raw_text:
        return ""
    lines = raw_text.split('\n')
    cleaned_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append("")
            continue
        if re.match(r'^[-*_]{3,}$', stripped):
            continue
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
    'GIÁO ÁN', 'KẾ HOẠCH', 'ĐỀ TÀI', 'CHỦ ĐỀ', 'Hoạt động 1', 'Hoạt động 2',
    'Hoạt động 3', '1. THỨ', '2. THỨ', '3. THỨ', '4. THỨ', '5. THỨ',
    'BÀI THƠ:', 'BÀI ĐỒNG DAO:', 'HỆ THỐNG CÂU ĐỐ', 'TRÒ CHƠI VẬN ĐỘNG',
    'MẪU TIN NHẮN', 'LỜI NHẬN XÉT', 'Ý TƯỞNG LÀM ĐỒ DÙNG'
)

LABEL_BOLD_PREFIXES = (
    '- Tình trạng sức khỏe:', '- Trạng thái cảm xúc, thái độ:', '-Trạng thái cảm xúc, thái độ:',
    '- Trạng thái cảm xúc:', '- Kiến thức, kỹ năng:', '- Kiến thức, kỹ năng & Năng lực số:',
    '1. Kiến thức:', '2. Kỹ năng:', '3. Thái độ:', '2. Kỹ năng & Năng lực số:', '1. Mục tiêu trọng tâm & Năng lực số:',
    '1. Đồ dùng của cô:', '2. Đồ dùng của trẻ:', '3. Không gian tổ chức:',
    '1. Mục đích - Yêu cầu:', '2. Chuẩn bị:', '3. Cách chơi:', '4. Luật chơi:',
    '- Bước 1:', '- Bước 2:', '- Bước 3:', '- Bước 4:',
    '* Bước 1:', '* Bước 2:', '* Bước 3:', '* Bước 4:',
    '- Khởi động:', '- Trọng động:', '- Hồi tĩnh:',
    '* Khởi động:', '* Trọng động:', '* Hồi tĩnh:',
    'a) ', 'b) ', 'c) ',
    '- Đón trẻ & Thể dục sáng:', '- Hoạt động học:', '- Hoạt động ngoài trời:',
    '- Hoạt động chơi ở các góc:', '- Hoạt động chiều:',
    '- Thể chất:', '- Nhận thức:', '- Ngôn ngữ:', '- Tình cảm - KNXH:', '- Thẩm mỹ:',
    '- Góc Phân vai:', '- Góc Xây dựng', '- Góc Nghệ thuật', '- Góc Tạo hình:',
    '- Góc Học tập', '- Góc Âm nhạc:', '- Góc Khám phá',
    '- Tuần 1:', '- Tuần 2:', '- Tuần 3:', '- Tuần 4:', '- Tuần 5:',
    'Lớp:', 'Chủ đề:', 'Người soạn:', 'Thời gian thực hiện:', 'Loại hoạt động:', 'Hướng tiếp cận:'
)

def is_full_bold_heading(line_s: str) -> bool:
    if re.match(r'^(I|II|III|IV|V|VI)\.\s', line_s):
        return True
    if line_s.startswith(FULL_BOLD_PREFIXES):
        return True
    if line_s.isupper() and len(line_s) > 5:
        return True
    return False

def should_bold_label_before_colon(line_s: str) -> bool:
    if ':' not in line_s:
        return False
    if line_s.startswith(LABEL_BOLD_PREFIXES):
        return True
    if re.match(r'^(\d+\.|[-+])\s+[^:]{2,38}:', line_s):
        return True
    return False

def format_text_to_standard_html(text_content: str, font_size_pt: int = 14) -> str:
    html_paragraphs = []
    for raw_line in text_content.split('\n'):
        line_s = raw_line.strip()
        if not line_s:
            html_paragraphs.append('<div style="height: 6pt;"></div>')
            continue

        p_style = (
            f"font-family: 'Times New Roman', Times, serif; "
            f"font-size: {font_size_pt}pt; "
            f"line-height: 1.5; "
            f"text-align: justify; "
            f"margin: 0 0 4pt 0; "
            f"color: #000000;"
        )

        if is_full_bold_heading(line_s):
            escaped = html.escape(line_s)
            html_paragraphs.append(f'<p style="{p_style}"><b>{escaped}</b></p>')
        elif should_bold_label_before_colon(line_s):
            idx = line_s.find(':')
            label_part = html.escape(line_s[:idx + 1])
            rest_part = html.escape(line_s[idx + 1:])
            html_paragraphs.append(f'<p style="{p_style}"><b>{label_part}</b>{rest_part}</p>')
        else:
            escaped = html.escape(line_s)
            html_paragraphs.append(f'<p style="{p_style}">{escaped}</p>')

    return "".join(html_paragraphs)

def export_docx_from_text(text_content: str, info: dict, custom_template=None, font_size_pt=13):
    if custom_template:
        doc = Document(custom_template)
    else:
        doc = Document()

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(font_size_pt)

    for raw_line in text_content.split('\n'):
        line_s = raw_line.strip()
        
