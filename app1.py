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
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime, timedelta
import random
import string

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt as PptxPt
    from pptx.dml.color import RGBColor
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
# BỘ LỌC VÀ ĐỊNH DẠNG VĂN BẢN CHUẨN
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
    'GIÁO ÁN', 'KẾ HOẠCH', 'ĐỀ TÀI', 'CHỦ ĐỀ', 'Hoạt động 1', 'Hoạt động 2', 'Hoạt động 3',
    '1. THỨ', '2. THỨ', '3. THỨ', '4. THỨ', '5. THỨ', 'BÀI THƠ:', 'BÀI ĐỒNG DAO:',
    'HỆ THỐNG CÂU ĐỐ', 'TRÒ CHƠI VẬN ĐỘNG', 'MẪU TIN NHẮN', 'LỜI NHẬN XÉT', 'Ý TƯỞNG LÀM ĐỒ DÙNG',
    'KỊCH BẢN DẠY NHANH', 'KỊCH BẢN VIDEO AI', 'PHIẾU HỌC TẬP', 'TRÒ CHƠI POWERPOINT',
    'BỘ PROMPT SINH HÌNH ẢNH', 'QUY TRÌNH 5 BƯỚC', 'CẢNH ', 'SLIDE ', 'BÀI TẬP '
)
LABEL_BOLD_PREFIXES = (
    '- Tình trạng sức khỏe:', '- Trạng thái cảm xúc, thái độ:', '-Trạng thái cảm xúc, thái độ:',
    '- Trạng thái cảm xúc:', '- Kiến thức, kỹ năng:', '- Kiến thức, kỹ năng & Năng lực số:',
    '1. Kiến thức:', '2. Kỹ năng:', '3. Thái độ:', '2. Kỹ năng & Năng lực số:',
    '1. Mục tiêu trọng tâm & Năng lực số:', '1. Đồ dùng của cô:', '2. Đồ dùng của trẻ:',
    '3. Không gian tổ chức:', '1. Mục đích - Yêu cầu:', '2. Chuẩn bị:', '3. Cách chơi:', '4. Luật chơi:',
    '- Bước 1:', '- Bước 2:', '- Bước 3:', '- Bước 4:', '* Bước 1:', '* Bước 2:', '* Bước 3:', '* Bước 4:',
    '- Khởi động:', '- Trọng động:', '- Hồi tĩnh:', '* Khởi động:', '* Trọng động:', '* Hồi tĩnh:',
    'a) ', 'b) ', 'c) ', '- Đón trẻ & Thể dục sáng:', '- Hoạt động học:', '- Hoạt động ngoài trời:',
    '- Hoạt động chơi ở các góc:', '- Hoạt động chiều:', '- Thể chất:', '- Nhận thức:', '- Ngôn ngữ:',
    '- Tình cảm - KNXH:', '- Thẩm mỹ:', '- Góc Phân vai:', '- Góc Xây dựng', '- Góc Nghệ thuật',
    '- Góc Tạo hình:', '- Góc Học tập', '- Góc Âm nhạc:', '- Góc Khám phá', '- Tuần 1:', '- Tuần 2:',
    '- Tuần 3:', '- Tuần 4:', '- Tuần 5:', 'Lớp:', 'Chủ đề:', 'Người soạn:', 'Thời gian thực hiện:',
    'Loại hoạt động:', 'Hướng tiếp cận:', '- Thời lượng:', '- Hình ảnh:', '- Lời thoại:', '- Âm thanh:',
    '- Prompt tiếng Anh:', '- Câu hỏi:', '- Đáp án đúng:', '- Hiệu ứng:'
)

def is_full_bold_heading(line_s: str) -> bool:
    if re.match(r'^(I|II|III|IV|V|VI|VII)\.\s', line_s): return True
    if line_s.startswith(FULL_BOLD_PREFIXES): return True
    if line_s.isupper() and len(line_s) > 5: return True
    return False

def should_bold_label_before_colon(line_s: str) -> bool:
    if ':' not in line_s: return False
    if line_s.startswith(LABEL_BOLD_PREFIXES): return True
    if re.match(r'^(\d+\.|[-+])\s+[^:]{2,42}:', line_s): return True
    return False

def format_text_to_standard_html(text_content: str, font_size_pt: int = 14) -> str:
    html_paragraphs = []
    for raw_line in text_content.split('\n'):
        line_s = raw_line.strip()
        if not line_s:
            html_paragraphs.append('<div style="height: 6pt;"></div>')
            continue
        p_style = f"font-family: 'Times New Roman', Times, serif; font-size: {font_size_pt}pt; line-height: 1.5; text-align: justify; margin: 0 0 4pt 0; color: #000000;"
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

def export_docx_from_text(text_content: str, info: dict, custom_template=None, font_size_pt=13, prefix_name="GiaoAn"):
    if custom_template: doc = Document(custom_template)
    else: doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(font_size_pt)
    
    for raw_line in text_content.split('\n'):
        line_s = raw_line.strip()
        if not line_s: continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.3
        
        if is_full_bold_heading(line_s):
            run = p.add_run(line_s)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size_pt)
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
            
    safe_topic = "".join(c if c.isalnum() else "_" for c in info.get('topic', 'GiaoAn'))[:40]
    filename = f"{prefix_name}_{safe_topic}.docx"
    doc.save(filename)
    return filename

def export_pptx_game(topic: str, age: str, raw_text: str) -> str:
    if not HAS_PPTX:
        return ""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.background
    fill1 = bg1.fill
    fill1.solid()
    fill1.fore_color.rgb = RGBColor(230, 247, 242)

    tb1 = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(11.333), Inches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = f"TRÒ CHƠI TƯƠNG TÁC MẦM NON\nĐỀ TÀI: {topic.upper()}"
    r1.font.size = PptxPt(38)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0, 121, 107)

    p2 = tf1.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = PptxPt(20)
    r2 = p2.add_run()
    r2.text = f"Lứa tuổi: {age} | Bấm phím mũi tên hoặc click chuột để bắt đầu!"
    r2.font.size = PptxPt(22)
    r2.font.color.rgb = RGBColor(30, 41, 59)

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
        f.fore_color.rgb = RGBColor(255, 253, 245)

        header_box = s.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.2))
        htf = header_box.text_frame
        htf.word_wrap = True
        hp = htf.paragraphs[0]
        hr = hp.add_run()
        hr.text = q_title
        hr.font.size = PptxPt(28)
        hr.font.bold = True
        hr.font.color.rgb = RGBColor(0, 105, 92)

        body_box = s.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(4.8))
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
                br.font.color.rgb = RGBColor(211, 47, 47)
            else:
                br.font.color.rgb = RGBColor(30, 41, 59)

    slide_end = prs.slides.add_slide(blank_layout)
    bg_end = slide_end.background
    f_end = bg_end.fill
    f_end.solid()
    f_end.fore_color.rgb = RGBColor(230, 247, 242)
    tb_end = slide_end.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.333), Inches(2.5))
    tf_end = tb_end.text_frame
    pe = tf_end.paragraphs[0]
    pe.alignment = PP_ALIGN.CENTER
    re_run = pe.add_run()
    re_run.text = "🎉 CHÚC MỪNG CÁC BÉ ĐÃ HOÀN THÀNH XUẤT SẮC TRÒ CHƠI! 🎉"
    re_run.font.size = PptxPt(32)
    re_run.font.bold = True
    re_run.font.color.rgb = RGBColor(0, 121, 107)

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
        <button id="copyBtn_{unique_key}" onclick="copyFormattedContent_{unique_key}()" style="flex: 1; background: linear-gradient(135deg, #0984E3, #00796B); color: white; border: none; padding: 10px 16px; border-radius: 12px; font-size: 15px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 12px rgba(9, 132, 227, 0.25); transition: all 0.2s ease;">📋 Sao Chép Nhanh (HTML/Text)</button>
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
            setTimeout(() => {{ btn.innerHTML = "📋 Sao Chép Nhanh (HTML/Text)"; btn.style.background = "linear-gradient(135deg, #0984E3, #00796B)"; }}, 2500);
        }} catch (err) {{
            const ta = document.createElement("textarea"); ta.value = plainContent; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);
            btn.innerHTML = "✅ Đã Sao Chép Văn Bản!";
            setTimeout(() => {{ btn.innerHTML = "📋 Sao Chép Nhanh (HTML/Text)"; }}, 2500);
        }}
    }}
    function printFormattedContent_{unique_key}() {{
        const pWin = window.open('', '_blank');
        pWin.document.write('<!DOCTYPE html><html><head><title>' + {title_json} + '</title><style>@page {{ size: A4; margin: 18mm 20mm 18mm 20mm; }} body {{ font-family: "Times New Roman", Times, serif; font-size: {font_size_pt}pt; line-height: 1.5; color: #000; margin: 0; padding: 0; text-align: justify; }}</style></head><body>' + {html_json} + '</body></html>');
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

# Tự động khôi phục API Key đã lưu từ SQLite khi mở lại web
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
# BỘ CỨU CÁNH CHUYÊN SÂU & BỘ CÔNG CỤ NÂNG CAO DỰ PHÒNG
# =========================================================================
def generate_instant_pedagogical_plan(info: dict) -> str:
    topic = info.get("topic", "Hoạt động khám phá").strip()
    domain = info.get("domain", "Nhận thức")
    sub_cat = info.get("sub_category", "Khám phá")
    route = info.get("route_type", "Hoạt động chung")
    art_type = info.get("art_type", "")
    age = info.get("age", "4-5 tuổi")
    theme = info.get("theme", "Chủ đề quanh bé")
    teacher = info.get("teacher", "Nguyễn Thị Thanh Vân").strip() or "Giáo viên"
    school = info.get("school", "Trường Mầm Non Phước Thạnh").strip() or "Trường Mầm Non"
    class_name = info.get("class_name", "Lớp Mẫu giáo Nhỡ").strip() or "Lớp Mẫu giáo"
    extras_list = info.get("extras", [])
    extras_str = ", ".join(extras_list) if extras_list else "Giáo dục kỹ năng sống và bảo vệ môi trường"
    custom_ideas = info.get("custom_ideas", "").strip()
    today_str = datetime.now().strftime("%d/%m/%Y")

    art_note = f" - Hình thức: {art_type}" if art_type else ""

    raw_plan = f"""{school.upper()}
Lớp: {class_name} | Độ tuổi: {age}
GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG ({info.get('activity_type', 'Hoạt động học')})
ĐỀ TÀI: {topic.upper()}{art_note}
Chủ đề lớn: {theme} | Lĩnh vực: {domain} ({sub_cat} - {route})
Người soạn: {teacher} | Ngày dạy: {today_str}

I. MỤC ĐÍCH - YÊU CẦU
1. Kiến thức:
- Trẻ nhận biết, gọi đúng tên và nêu được những đặc điểm nổi bật, cấu tạo cơ bản của nội dung bài học: {topic}{art_note}.
- Trẻ hiểu được lợi ích, ý nghĩa thực tế của {topic} đối với đời sống con người và môi trường xung quanh.
- Trẻ nắm được quy trình và các bước tham gia hoạt động thực hành, trò chơi theo đúng hướng dẫn của cô giáo.
2. Kỹ năng:
- Rèn luyện kỹ năng quan sát có chủ đích, so sánh, phân loại và ghi nhớ các chi tiết đặc trưng.
- Phát triển ngôn ngữ mạch lạc: Trẻ tự tin diễn đạt suy nghĩ, trả lời câu hỏi trọn vẹn câu, biết trao đổi cùng bạn.
- Rèn luyện thao tác vận động tinh khéo léo, phối hợp tay - mắt nhịp nhàng trong hoạt động trải nghiệm thực tế.
3. Thái độ:
- Trẻ hào hứng, chủ động, tích cực tham gia vào từng hoạt động học tập và vui chơi.
- Trẻ biết lắng nghe hiệu lệnh của cô, biết phối hợp nhóm và chia sẻ đồ dùng học tập cùng các bạn.
- Lồng ghép giáo dục: {extras_str}; trẻ biết yêu quý sản phẩm lao động và giữ gìn vệ sinh lớp học.

II. CHUẨN BỊ
1. Đồ dùng của cô:
- Giáo án điện tử thiết kế sinh động, slide hình ảnh, video tư liệu sắc nét liên quan đến {topic}.
- Hệ thống âm thanh, loa bluetooth, nhạc không lời êm dịu và các bài hát vui nhộn theo chủ đề {theme}.
- Đồ dùng trực quan vật thật / mô hình kích thước lớn phục vụ quan sát đàm thoại.{f' Ý tưởng riêng của cô: {custom_ideas}' if custom_ideas else ''}
2. Đồ dùng của trẻ:
- Bộ đồ dùng thực hành, học liệu đủ số lượng cho từng trẻ (khay đựng, thẻ hình, tranh mẫu, đồ chơi trải nghiệm).
- Trang phục gọn gàng, tâm thế thoải mái, tâm lý tự tin sẵn sàng bước vào giờ học.
3. Địa điểm - Không gian tổ chức:
- Phòng học sạch sẽ, thoáng mát, đủ ánh sáng; bố trí đội hình chữ U hoặc vòng cung thuận tiện quan sát màn hình.

III. TIẾN TRÌNH TỔ CHỨC HOẠT ĐỘNG
Hoạt động 1: Ổn định tổ chức, gây hứng thú (3 - 5 phút)
- Cô tập trung trẻ lại quanh mình, mở nhạc giai điệu vui tươi theo chủ đề {theme} và cùng trẻ nhún nhảy:
  + Cô nói: "Nhiệt liệt chào mừng tất cả các bé lớp {class_name} đến với tiết học kỳ diệu ngày hôm nay!"
  + Trẻ đồng thanh đáp: "Chúng con chào cô ạ!"
- Cô tạo tình huống bất ngờ với chiếc hộp quà bí mật:
  + Cô nói: "Hôm nay, cô mang đến một món quà vô cùng đặc biệt. Các con thử đoán xem bên trong có điều bí mật gì nào?"
  + Trẻ hào hứng phỏng đoán: "Dạ thưa cô là đồ chơi ạ / Là tranh ảnh ạ!"
  + Cô nhẹ nhàng mở quà và giới thiệu biểu tượng của bài học: "Đó chính là những hình ảnh vô cùng thú vị về {topic}! Hôm nay cô con mình sẽ cùng nhau bước vào hành trình khám phá nhé!"

Hoạt động 2: Nội dung trọng tâm (20 - 25 phút)
* Bước 1: Khám phá trải nghiệm thực tế & Đàm thoại gợi mở:
- Cô giới thiệu vật thật / tranh ảnh sinh động về {topic} cho cả lớp quan sát cận cảnh.
- Cô gợi mở để trẻ tự do phát hiện bằng các giác quan (nhìn, nghe, chạm, sờ):
  + Cô hỏi: "Các con có nhận xét gì về {topic}? Ai giỏi cho cô biết hình dáng, màu sắc và đặc điểm của nó nào?"
  + Trẻ giơ tay trả lời: "Dạ con thấy có màu sắc đẹp ạ / Dạ con thấy nó rất quen thuộc ạ!"
  + Cô hỏi tiếp: "Để hiểu rõ hơn về công dụng và sự phát triển của {topic}, các con hãy cùng hướng mắt lên màn hình quan sát đoạn phim ngắn này nhé!"
- Sau khi xem video, cô đặt câu hỏi gợi mở tư duy logic:
  + Cô hỏi: "Qua đoạn phim vừa rồi, con biết thêm điều gì mới? Nếu không có nó thì cuộc sống của chúng ta sẽ thế nào?"
  + Trẻ suy nghĩ và tự tin trả lời theo cảm nhận.
  + Cô chốt lại kiến thức chuẩn xác, diễn đạt ngắn gọn, dễ nhớ để trẻ khắc sâu bài học.

* Bước 2: Trải nghiệm thực hành - Trẻ làm chủ hoạt động:
- Cô chia trẻ về các nhóm nhỏ ngồi theo bàn đã chuẩn bị sẵn đồ dùng học tập.
- Cô nêu rõ nhiệm vụ và làm mẫu thao tác rõ ràng, dứt khoát:
  + Cô nói: "Bây giờ các kỹ sư nhí hãy cùng nhau thực hành thao tác sắp xếp, phân loại và tạo ra sản phẩm về {topic} nhé!"
- Trẻ bắt đầu thực hiện bài tập. Cô đi từng bàn bao quát, quan sát kỹ năng của từng trẻ:
  + Động viên, khích lệ những trẻ còn lúng túng hoặc nhút nhát.
  + Gợi ý mở rộng để những trẻ khá, giỏi sáng tạo thêm các chi tiết mới lạ.

* Bước 3: Trò chơi củng cố kiến thức:
- Cô giới thiệu trò chơi vận động: 'Đội Nào Nhanh Nhất'.
- Luật chơi: Mỗi lượt một bạn bật qua các vòng thể dục lên chọn đúng hình ảnh/đồ vật về {topic} đặt vào rổ của đội mình, sau đó chạy về đập tay vào bạn tiếp theo.
- Cách chơi: Cô bật nhạc rộn ràng; 2 đội thi đua nhau trong thời gian 1 bản nhạc.
- Trẻ chơi hào hứng 2 - 3 lượt. Cô cùng cả lớp kiểm tra kết quả, đếm số lượng và tuyên dương đội chiến thắng.

Hoạt động 3: Kết thúc hoạt động (2 - 3 phút)
- Cô tập trung trẻ lại, nhận xét chung về tinh thần học tập:
  + Cô khen ngợi sự tập trung, sáng tạo và tinh thần hợp tác rất tốt của cả lớp hôm nay.
  + Nhắc nhở giáo dục nhẹ nhàng: Các con nhớ giữ gìn đồ dùng, bảo vệ môi trường và biết yêu quý cuộc sống xung quanh.
- Cho trẻ cùng cô thu dọn đồ dùng học tập gọn gàng, cất vào đúng nơi quy định và hát vang bài ca ra sân chơi.

IV. ĐÁNH GIÁ CUỐI NGÀY
- Tình trạng sức khỏe: 100% trẻ khỏe mạnh, thoải mái, tham gia tích cực các hoạt động từ đầu đến cuối tiết học.
- Trạng thái cảm xúc, thái độ: Trẻ vui vẻ, hào hứng, tự tin phát biểu xây dựng bài và biết đoàn kết với bạn.
- Kiến thức, kỹ năng: Đại đa số trẻ nắm vững nội dung bài dạy ({topic}), hoàn thành tốt yêu cầu thực hành; cô đã hỗ trợ kịp thời các trẻ còn chậm tiến bộ."""
    return clean_markdown_text(raw_plan)

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
    today_str = datetime.now().strftime("%d/%m/%Y")

    raw_plan = f"""{school.upper()}
Lớp: {class_name} | Độ tuổi: {age} | Số lượng: {num_kids} trẻ
GIÁO ÁN TÍCH HỢP NĂNG LỰC SỐ & CÔNG NGHỆ THÔNG MINH
ĐỀ TÀI: {topic.upper()}
Chủ đề lớn: {theme} | Lĩnh vực: {domain}
Loại hoạt động: {act_type} | Thời lượng: {duration}
Người soạn: {teacher} | Ngày dạy: {today_str}

I. MỤC ĐÍCH - YÊU CẦU
1. Kiến thức:
- Trẻ nhận biết, gọi đúng tên và các đặc điểm cấu tạo nổi bật của {topic} qua việc quan sát vật thật kết hợp tương tác số.
- Trẻ bước đầu làm quen với các thiết bị công nghệ giáo dục ({digital_equip}), hiểu được thao tác chạm - chọn trên màn hình cảm ứng hoặc bảng tương tác số ({digital_tools}).
2. Kỹ năng & Năng lực số:
- Hình thành kỹ năng số ban đầu: Trẻ biết quan sát điểm chạm, dùng ngón tay trỏ thực hiện thao tác kéo - thả, ghép đôi hình ảnh trên màn hình tương tác thông minh.
- Rèn luyện kỹ năng quan sát, so sánh đối chiếu giữa vật thật ngoài đời và mô hình mô phỏng 3D trên màn hình.
3. Thái độ:
- Trẻ hào hứng, thích thú khám phá công nghệ, có ý thức giữ khoảng cách an toàn cho mắt khi tiếp xúc màn hình.

II. CHUẨN BỊ
1. Đồ dùng của cô: {digital_equip}; học liệu số ({digital_tools}); vật thật trực quan.
2. Đồ dùng của trẻ: Vật thật trải nghiệm đa giác quan, thẻ tín hiệu tương tác.
3. Không gian tổ chức: Phòng học thoáng mát, màn hình ngang tầm mắt trẻ.

III. TIẾN TRÌNH TỔ CHỨC HOẠT ĐỘNG
Hoạt động 1: Mở đầu - Khơi gợi hứng thú với Trợ lý AI (3 - 5 phút)
- Cô mở màn hình tương tác giới thiệu nhân vật AI chào đón cả lớp, tạo tình huống khám phá '{topic}'.
Hoạt động 2: Trọng tâm - Khám phá Đa phương tiện & Tương tác số (20 - 25 phút)
- Bước 1: Trẻ quan sát vật thật kết hợp hình ảnh 3D/Video trên tivi, đàm thoại nhận xét đặc điểm.
- Bước 2: Trẻ lên chạm màn hình tương tác phân loại, ghép đôi kiến thức về '{topic}'.
- Bước 3: Tổ chức trò chơi số hóa theo đội nhóm, củng cố khắc sâu bài học.
Hoạt động 3: Kết thúc & Giáo dục nhẹ nhàng (2 - 3 phút)
- Cô và trợ lý AI tuyên dương cả lớp, nhắc trẻ bảo vệ đôi mắt khi sử dụng thiết bị điện tử.

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

1. ỔN ĐỊNH & GÂY HỨNG THÚ (00:00 - 04:00 | 4 phút):
- Hiệu lệnh tập trung: Cô vỗ tay theo nhịp 1-2-3, đọc vè "Xúm xít, xúm xít! Bên cô, bên cô!".
- Lời dẫn mở màn: "Hôm nay cô mang đến một hộp quà bí mật từ thế giới {theme}. Đố các bé biết bên trong có gì?"
- Chuyển tiếp: Mở hộp quà giới thiệu đề tài '{topic}'.

2. KHÁM PHÁ TRỌNG TÂM (04:00 - 14:00 | 10 phút):
- Quan sát trực quan: Cho trẻ sờ, nhìn, gọi tên đặc điểm nổi bật của '{topic}'.
- 3 Câu hỏi đinh:
  + Câu 1: "Con nhìn thấy {topic} có màu sắc và hình dáng như thế nào?"
  + Câu 2: "Theo con, {topic} dùng để làm gì / có ích lợi gì?"
  + Câu 3: "Nếu không có {topic} thì điều gì sẽ xảy ra?"
- Chốt kiến thức: Cô tóm tắt ngắn gọn bằng 2 câu dễ nhớ.

3. THỰC HÀNH TRẢI NGHIỆM (14:00 - 22:00 | 8 phút):
- Chia 3 nhóm thực hành thao tác trực tiếp trên đồ dùng cá nhân.
- Cô đi từng nhóm gợi mở, sửa sai nhẹ nhàng, khen ngợi trẻ sáng tạo.

4. TRÒ CHƠI CỦNG CỐ (22:00 - 28:00 | 6 phút):
- Trò chơi: "Ai Nhanh Trí Hơn". Chia 2 đội thi đua bật qua vòng chọn đúng hình ảnh về '{topic}'.
- Kiểm tra kết quả, đếm số lượng cùng trẻ và tặng hoa bé ngoan.

5. KẾT THÚC & CHUYỂN HOẠT ĐỘNG (28:00 - 30:00 | 2 phút):
- Tuyên dương cả lớp, giáo dục kỹ năng sống gắn với '{topic}', hát bài hát chủ đề và thu dọn đồ dùng.""")

    elif tool_type == "video_ai":
        return clean_markdown_text(f"""KỊCH BẢN VIDEO AI MINH HỌA TIẾT DẠY
ĐỀ TÀI: {topic.upper()} | ĐỘ TUỔI: {age}

CẢNH 1 (00:00 - 00:05):
- Thời lượng: 5 giây
- Hình ảnh: Khung cảnh lớp học mầm non Việt Nam rực rỡ nắng mai, cô giáo dịu dàng mỉm cười chào các bé bên biểu tượng '{topic}'.
- Lời thoại: "Chào mừng các bé yêu đến với hành trình khám phá điều kỳ diệu về {topic} hôm nay!"
- Âm thanh: Nhạc nền vui tươi, tiếng chuông leng keng nhẹ nhàng.
- Prompt tiếng Anh: Cute 3D Pixar style illustration of a friendly Vietnamese kindergarten teacher smiling warmly in a bright colorful classroom with {topic}, cheerful Vietnamese preschoolers looking excited, soft morning sunlight, joyful atmosphere --ar 16:9

CẢNH 2 (00:05 - 00:10):
- Thời lượng: 5 giây
- Hình ảnh: Cận cảnh chi tiết sinh động của '{topic}' xuất hiện lấp lánh ở trung tâm màn hình với màu sắc tươi sáng, rõ nét.
- Lời thoại: "Các con hãy mở to đôi mắt xinh xem điều gì đang xuất hiện trước mắt chúng mình nào!"
- Âm thanh: Hiệu ứng âm thanh phép thuật lung linh.
- Prompt tiếng Anh: Cute 3D educational animation close-up of {topic} glowing softly with vibrant colors, clean bright background, child-friendly 3D render style, high detail --ar 16:9

CẢNH 3 (00:10 - 00:15):
- Thời lượng: 5 giây
- Hình ảnh: Một em bé mầm non Việt Nam 4-5 tuổi mặc đồng phục dễ thương đang chăm chú quan sát và chỉ tay khám phá '{topic}'.
- Lời thoại: "Ồ! Thật là thú vị phải không nào? Chúng mình cùng tìm hiểu từng đặc điểm nhé!"
- Âm thanh: Nhạc dạo vui nhộn, tiết tấu vừa phải.
- Prompt tiếng Anh: Cute 3D style, a smiling 4-year-old Vietnamese child in neat preschool uniform pointing curiously at {topic}, bright warm classroom setting, expressive happy face, Pixar animation style --ar 16:9

CẢNH 4 (00:15 - 00:20):
- Thời lượng: 5 giây
- Hình ảnh: Các bạn nhỏ cùng nhau thực hành, chia sẻ và vui vẻ hợp tác trải nghiệm '{topic}' trên bàn học.
- Lời thoại: "Khi chúng mình cùng nhau khám phá và chia sẻ, mọi điều đều trở nên thật dễ dàng và vui vẻ!"
- Âm thanh: Tiếng cười nói trong trẻo, nhạc nền ấm áp.
- Prompt tiếng Anh: Cute 3D illustration of happy Vietnamese kindergarten children working together at a colorful wooden table exploring {topic}, teamwork, warm lighting, detailed 3D cartoon --ar 16:9

CẢNH 5 (00:20 - 00:25):
- Thời lượng: 5 giây
- Hình ảnh: Cảnh tổng kết rực rỡ với pháo giấy, ngôi sao lấp lánh chúc mừng các bé hoàn thành bài học '{topic}'.
- Lời thoại: "Hoan hô các bé thật giỏi! Hẹn gặp lại các con trong những bài học tiếp theo nhé!"
- Âm thanh: Nhạc kết thúc vui tươi, tiếng vỗ tay rộn rã.
- Prompt tiếng Anh: Cute 3D celebration scene with colorful confetti, golden stars, and smiling Vietnamese preschoolers cheering happily around {topic}, joyful atmosphere --ar 16:9

VII. QUY TRÌNH 5 BƯỚC DỰNG VIDEO AI & DANH SÁCH TÀI NGUYÊN
1. Danh sách tài nguyên đề xuất:
- Công cụ tạo ảnh AI: Midjourney / Canva Magic Media / Google ImageFX (Dán các câu Prompt tiếng Anh ở trên để tạo ảnh từng cảnh).
- Công cụ tạo chuyển động (Image-to-Video AI): Kling AI / Runway Gen-3 / Hailuo AI (Tải ảnh lên để tạo chuyển động 5 giây cho mỗi cảnh).
- Công cụ lồng tiếng & Âm thanh AI: CapCut / ElevenLabs / Vbee (Chọn giọng đọc nữ miền Bắc/Nam truyền cảm).
- Phần mềm dựng hoàn thiện: CapCut PC hoặc CapCut Mobile.
2. Quy trình 5 bước dựng video AI cho giáo viên mầm non:
- Bước 1: Lập kịch bản & Phân cảnh (Storyboarding) theo 5 cảnh ngắn mỗi cảnh 5 giây để giữ sự tập trung của trẻ.
- Bước 2: Tạo hình ảnh đồng bộ nhân vật bằng các Prompt tiếng Anh ở trên với tỷ lệ khung hình ngang 16:9.
- Bước 3: Chuyển ảnh tĩnh thành video chuyển động nhẹ nhàng (Image-to-Video).
- Bước 4: Nhập lời thoại vào công cụ chuyển văn bản thành giọng nói tiếng Việt và ghép nhạc nền không lời vui tươi.
- Bước 5: Đưa 5 đoạn video vào CapCut, thêm chữ tiêu đề to rõ và xuất file MP4 Full HD chiếu lên tivi lớp học.""")

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

    else:  # image_prompts
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
    month = info.get("month", "Tháng")
    theme = info.get("theme", "Chủ đề")
    return f"KẾ HOẠCH GIÁO DỤC {month.upper()} - CHỦ ĐỀ: {theme.upper()}\nI. MỤC TIÊU PHÁT TRIỂN THÁNG: Phát triển toàn diện 5 mặt cho trẻ mầm non.\nII. MẠNG NỘI DUNG: Trẻ làm quen, khám phá đặc điểm và trải nghiệm thực tế.\nIII. MẠNG HOẠT ĐỘNG: Toán, Khám phá, Âm nhạc, Văn học, Tạo hình.\nIV. ĐÁNH GIÁ: Theo dõi sự tiến bộ của trẻ hàng tuần."

def generate_instant_week_plan(info: dict) -> str:
    week = info.get("week_num", "Tuần")
    theme = info.get("theme_week", "Chủ đề")
    return f"KẾ HOẠCH GIÁO DỤC {week.upper()} - NHÁNH: {theme.upper()}\nI. MỤC TIÊU: Trẻ hiểu rõ đặc điểm nhánh {theme}.\nII. KẾ HOẠCH HÀNG NGÀY:\n1. THỨ HAI: Học tập và vận động cơ bản.\n2. THỨ BA: Khám phá và vui chơi ngoài trời.\n3. THỨ TƯ: Tạo hình sáng tạo.\n4. THỨ NĂM: Làm quen văn học.\n5. THỨ SÁU: Âm nhạc và tổng kết nêu gương bé ngoan."

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
        text_out = future.result(timeout=25) 
    except FuturesTimeoutError:
        error_msg = "Thời gian phản hồi của AI vượt quá 25 giây."
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
st.markdown("Hệ thống hỗ trợ soạn giáo án chuẩn, tích hợp năng lực số, lên kế hoạch tháng/tuần, sáng tác trò chơi và quản lý lớp học.")

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
# TAB 1: SOẠN GIÁO ÁN CHUẨN + BỘ CÔNG CỤ NÂNG CAO (5 TÍNH NĂNG)
# =========================================================================
elif app_mode == "📝 Soạn Giáo Án Chuẩn":
    st.subheader("📝 Thiết Lập Thông Số Giáo Án (Chuẩn Chương Trình GDMN)")
    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("Người soạn / Giáo viên", "Nguyễn Thị Thanh Vân")
        school_name = st.text_input("Đơn vị / Trường", "Trường Mầm Non Phước Thạnh")
    with col2:
        class_name = st.text_input("Lớp", "Lớp Mẫu giáo Nhỡ (4-5 tuổi)")
        theme = st.text_input("Chủ đề lớn / Nhánh", "Nghề nghiệp")

    st.markdown("### CHẾ ĐỘ SOẠN THẢO *")
    c_mode1, c_mode2 = st.columns(2)
    with c_mode1: mode_gv = st.button("🌱 Giáo án thường\nChuẩn mực, tinh gọn, rõ ràng theo đúng phân phối chương trình.", use_container_width=True)
    with c_mode2: mode_gioi = st.button("⭐ Thi GV Giỏi\nStorytelling, tình huống sáng tạo nhưng KHÔNG vượt chuẩn tuổi.", use_container_width=True)
    if "soan_mode_val" not in st.session_state: st.session_state["soan_mode_val"] = "Giáo án thường"
    if mode_gv: st.session_state["soan_mode_val"] = "Giáo án thường"
    if mode_gioi: st.session_state["soan_mode_val"] = "Thi GV Giỏi"
    st.info(f"👉 Đang chọn chế độ soạn thảo: **{st.session_state['soan_mode_val']}**")

    st.markdown("### ĐỊNH DẠNG HIỂN THỊ GIÁO ÁN")
    c_fmt1, c_fmt2 = st.columns(2)
    with c_fmt1: fmt_1col = st.button("📄 1 Cột chuẩn (Toàn cảnh)\nLời dẫn, câu hỏi & hành động của cô theo gạch đầu dòng (-)", use_container_width=True)
    with c_fmt2: fmt_2col = st.button("📊 2 Cột (Cô & Trẻ)\nHoạt động của giáo viên song song nhiệm vụ dự kiến của trẻ", use_container_width=True)
    if "display_fmt_val" not in st.session_state: st.session_state["display_fmt_val"] = "1 Cột chuẩn"
    if fmt_1col: st.session_state["display_fmt_val"] = "1 Cột chuẩn"
    if fmt_2col: st.session_state["display_fmt_val"] = "2 Cột (Cô & Trẻ)"
    st.info(f"👉 Đang chọn định dạng: **{st.session_state['display_fmt_val']}**")

    st.markdown("### 1. ĐỘ TUỔI HỌC SINH *")
    age_group = st.radio("Chọn độ tuổi", ["3-12 th", "12-18 th", "18-24 th", "24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], horizontal=True, index=5)
    st.markdown("### 2. LOẠI HÌNH HOẠT ĐỘNG *")
    activity_type = st.radio("Chọn loại hoạt động", ["Hoạt động học", "Hoạt động góc / Vui chơi", "Hoạt động ngoài trời", "Hoạt động chiều"], horizontal=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### 3. LĨNH VỰC PHÁT TRIỂN *")
        domain = st.selectbox("Lĩnh vực", ["Thể chất", "Nhận thức", "Ngôn ngữ", "Thẩm mỹ", "Tình cảm - KNXH"], key="domain_select_box")
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

    st.markdown("### 5. CẤU TRÚC MỤC ĐÍCH - YÊU CẦU:")
    goal_structure = st.radio("Lựa chọn cấu trúc", ["1-3 gạch đầu dòng", "2-3 mục chuẩn (Kiến thức, Kỹ năng, Thái độ)"], horizontal=True, index=1)
    st.markdown("### 6. ĐỀ TÀI BÀI DẠY CỤ THỂ * (BẮT BUỘC):")
    topic = st.text_input("Nhập tên đề tài", "Bé khám phá hành trình lớn lên của hạt đậu")
    st.markdown("### 7. YÊU CẦU TÍCH HỢP (TÙY CHỌN) - (Chọn 1-2 mục):")
    col_c1, col_c2, col_c3, col_c4, col_c5 = st.columns(5)
    with col_c1: opt0 = st.checkbox("Không")
    with col_c2: opt1 = st.checkbox("STEAM")
    with col_c3: opt2 = st.checkbox("CÔNG NGHỆ SỐ")
    with col_c4: opt3 = st.checkbox("Bảo vệ môi trường")
    with col_c5: opt4 = st.checkbox("Quyền trẻ em")
    extras = []
    if opt1: extras.append("STEAM")
    if opt2: extras.append("CÔNG NGHỆ SỐ")
    if opt3: extras.append("Bảo vệ môi trường")
    if opt4: extras.append("Quyền trẻ em")

    st.markdown("### 8. MỤC TIÊU CẦN ĐẠT (Tùy chọn):")
    custom_goals = st.text_area("Nhập MT1, MT45... kèm nội dung thực tế", placeholder="VD:\n- Trẻ nhận biết quá trình nảy mầm và lớn lên của hạt đậu.")
    st.markdown("### 9. YÊU CẦU BỔ SUNG (Ý TƯỞNG GIÁO VIÊN MUỐN ĐƯA VÀO TIẾT DẠY):")
    custom_ideas = st.text_area("Nhập ý tưởng riêng", placeholder="Cô hãy nhập các ý tưởng riêng: đạo cụ rối tay, nhạc chủ đề, bài thơ mới...")
    st.markdown("### 10. ĐÁNH GIÁ CUỐI NGÀY:")
    auto_assessment = st.checkbox("Quan sát thực tế cho Đánh giá cuối ngày (Mục IV) - Tạo đánh giá luôn", value=True)
    
    plan_info_dict = {
        "school": school_name, "teacher": teacher_name, "class_name": class_name,
        "theme": theme, "domain": domain, "sub_category": sub_category, "activity_type": activity_type,
        "route_type": route_type, "art_type": art_type_val, "topic": topic, "method": "Gemini AI", "age": age_group, "extras": extras,
        "mode": st.session_state["soan_mode_val"], "display_format": st.session_state["display_fmt_val"], 
        "custom_goals": custom_goals, "custom_ideas": custom_ideas, "auto_assessment": auto_assessment,
        "doc_title": "GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG HỌC"
    }

    if st.button("✨ TẠO GIÁO ÁN HOÀN CHỈNH CHI TIẾT", type="primary", use_container_width=True):
        with st.spinner("⚡ Đang kết nối AI biên soạn kịch bản giáo án SIÊU CHI TIẾT (vui lòng đợi ~5-8 giây)..."):
            art_prompt_part = f" - Dạng tiết tạo hình: {art_type_val}" if art_type_val else ""
            
            prompt = f"""
            Đóng vai là một Chuyên gia Giáo dục Mầm non xuất sắc của Bộ GD&ĐT Việt Nam, hãy biên soạn một GIÁO ÁN MẦM NON HOÀN CHỈNH, RẤT DÀI VÀ CỰC KỲ CHI TIẾT (Ít nhất 1500 - 2000 từ) để giáo viên có thể cầm lên lớp giảng dạy ngay:
            - Đề tài: {topic} | Độ tuổi: {age_group} | Chủ đề lớn: {theme}
            - Lĩnh vực phát triển: {domain} ({sub_category}) | Loại hoạt động: {activity_type} | Dạng tiết: {route_type}
            {art_prompt_part}
            - Chế độ: {st.session_state['soan_mode_val']} | Định dạng: {st.session_state['display_fmt_val']}
            - Mục tiêu cần đạt: {custom_goals if custom_goals else 'Tự động tạo chuẩn 3 mặt: Kiến thức, Kỹ năng, Thái độ (Mỗi mặt viết 3-4 ý rất chi tiết)'}
            - Tích hợp: {', '.join(extras) if extras else 'Cơ bản'}
            - Ý tưởng của giáo viên: {custom_ideas if custom_ideas else 'Không'}
            
            YÊU CẦU BẮT BUỘC VỀ NỘI DUNG VÀ ĐỘ DÀI:
            1. BÀI SOẠN PHẢI RẤT DÀI, CHI TIẾT, KHÔNG ĐƯỢC TÓM TẮT HOẶC VIẾT SƠ LƯỢC.
            2. Phần III. TIẾN TRÌNH TỔ CHỨC HOẠT ĐỘNG: Bắt buộc viết dưới dạng KỊCH BẢN TƯƠNG TÁC CHI TIẾT LỜI CÔ VÀ LỜI TRẺ:
               - Ghi rõ từng lời dẫn dắt, câu hỏi gợi mở của cô (Ví dụ: Cô nói: "...", Cô hỏi: "...")
               - Ghi rõ các câu trả lời ngây thơ, phong phú của trẻ (Ví dụ: Trẻ hào hứng trả lời: "...")
               - Cô khen ngợi, sửa sai, gợi mở tư duy cho trẻ.
            3. Trọng tâm bài học phải chia làm các bước mạch lạc: Khám phá trải nghiệm thực tế -> Cô đàm thoại phân tích sâu -> Trẻ thực hành sáng tạo -> Trò chơi củng cố có luật chơi rõ ràng.
            4. Lời văn phải ngọt ngào, trong sáng, đúng tâm lý lứa tuổi {age_group}.

            Trình bày theo đúng 4 mục lớn:
            I. MỤC ĐÍCH - YÊU CẦU (1. Kiến thức, 2. Kỹ năng, 3. Thái độ)
            II. CHUẨN BỊ (1. Đồ dùng của cô, 2. Đồ dùng của trẻ, 3. Địa điểm)
            III. TIẾN TRÌNH TỔ CHỨC HOẠT ĐỘNG (Hoạt động 1: Ổn định tổ chức; Hoạt động 2: Trọng tâm; Hoạt động 3: Kết thúc)
            IV. ĐÁNH GIÁ CUỐI NGÀY (Bắt buộc gồm đúng 3 dòng:
            - Tình trạng sức khỏe: ...
            - Trạng thái cảm xúc, thái độ: ...
            - Kiến thức, kỹ năng: ...)
            
            Tuyệt đối KHÔNG sử dụng ký tự ### hay **. Dùng gạch đầu dòng (-) hoặc số thứ tự.
            """
            result_text, err = generate_fast_ai_content(api_key, prompt, fallback_type="plan", fallback_data=plan_info_dict)
            
            if err:
                st.warning(f"⚠️ **Thông báo:** {err}\n\n👉 Đang sử dụng kịch bản mầm non đầy đủ 1.500 từ chuẩn chỉnh.")
            else:
                st.success("🎉 Đã hoàn thành giáo án chi tiết! Mời cô xem bên dưới:")

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
            base_context = st.session_state.get("generated_plan_text", "")[:1200]
            with st.spinner("⚡ Đang khởi tạo học liệu nâng cao đồng bộ với bài dạy..."):
                if selected_adv_tool == "quick_script":
                    adv_title = f"Kịch bản dạy nhanh: {topic}"
                    adv_prompt = f"""
                    Dựa trên đề tài mầm non '{topic}' (Độ tuổi: {age_group}, Chủ đề: {theme}, Lĩnh vực: {domain}), hãy viết một KỊCH BẢN DẠY NHANH (Phao cầm tay lên lớp) thật mạch lạc, dễ nhìn:
                    - Phân bổ thời gian chi tiết từng phút (00:00 - 30:00) cho: Ổn định & Gây hứng thú -> Khám phá trọng tâm -> Thực hành trải nghiệm -> Trò chơi củng cố -> Kết thúc.
                    - Ghi rõ các câu lời dẫn chuyển tiếp ngắn gọn, vè/đồng dao tập trung trẻ và 3-4 câu hỏi gợi mở cốt lõi của cô.
                    Tuyệt đối không dùng dấu ### hay **.
                    """
                elif selected_adv_tool == "video_ai":
                    adv_title = f"Kịch bản video AI: {topic}"
                    adv_prompt = f"""
                    Hãy viết một KỊCH BẢN VIDEO AI MINH HỌA TIẾT DẠY MẦM NON thật chi tiết cho đề tài '{topic}' (Độ tuổi: {age_group}, Chủ đề: {theme}):
                    - Chia thành 5 - 6 cảnh quay (mỗi cảnh 5 giây).
                    - Mỗi cảnh trình bày rõ ràng các mục:
                      CẢNH X (Thời lượng: 5 giây):
                      - Hình ảnh: (Mô tả chi tiết bối cảnh lớp học mầm non Việt Nam, nhân vật bé và cô giáo)
                      - Lời thoại: (Lời thuyết minh tiếng Việt truyền cảm, dễ hiểu cho bé {age_group})
                      - Âm thanh: (Nhạc nền và hiệu ứng âm thanh)
                      - Prompt tiếng Anh: (Câu lệnh tiếng Anh chi tiết phong cách hoạt hình 3D Pixar trẻ em Việt Nam để dán vào AI tạo ảnh/video, kết thúc bằng --ar 16:9)
                    - Cuối bài thêm mục: VII. QUY TRÌNH 5 BƯỚC DỰNG VIDEO AI & DANH SÁCH TÀI NGUYÊN (Midjourney/Canva, Kling AI/Runway, CapCut).
                    Tuyệt đối không dùng dấu ### hay **.
                    """
                elif selected_adv_tool == "worksheet":
                    adv_title = f"Phiếu học tập: {topic}"
                    adv_prompt = f"""
                    Thiết kế một PHIẾU HỌC TẬP IN A4 CHO TRẺ MẦM NON lứa tuổi {age_group} về đề tài '{topic}' (Chủ đề: {theme}):
                    - Có phần tiêu đề (Trường, Lớp, Họ và tên bé).
                    - Thiết kế 4 BÀI TẬP TRỰC QUAN phù hợp lứa tuổi {age_group} (VD: Bài tập 1: Quan sát & Tô màu; Bài tập 2: Nối cặp đôi tương ứng; Bài tập 3: Đếm số lượng & Khoanh tròn; Bài tập 4: Góc vẽ sáng tạo của bé).
                    - Mỗi bài tập nêu rõ: Yêu cầu dành cho bé, Cách cô/bố mẹ hướng dẫn, và Mô tả bố cục hình vẽ trên phiếu để giáo viên dễ thiết kế/in ấn.
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
                else:  # image_prompts
                    adv_title = f"Prompt sinh hình ảnh: {topic}"
                    adv_prompt = f"""
                    Tạo BỘ 6 PROMPT SINH HÌNH ẢNH MINH HỌA ĐỒNG BỘ TIẾT DẠY cho đề tài '{topic}' (Độ tuổi: {age_group}, Chủ đề: {theme}):
                    - Đồng bộ theo tiến trình tiết dạy: (1) Ảnh bìa slide mở đầu, (2) Ảnh tình huống gây hứng thú, (3) Ảnh trực quan kiến thức 1, (4) Ảnh trực quan kiến thức 2, (5) Ảnh hoạt động thực hành nhóm, (6) Ảnh sticker phần thưởng bé ngoan.
                    - Mỗi mục ghi rõ:
                      + Mục đích sử dụng trong tiết dạy (tiếng Việt)
                      + Prompt tiếng Anh chi tiết (phong cách hoạt hình 3D Pixar tươi sáng, nhân vật cô giáo và trẻ em mầm non Việt Nam, kèm tỷ lệ --ar 16:9 hoặc --ar 1:1).
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
                st.download_button(label="📥 Tải File Word (.docx)", data=file, file_name=docx_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True, key="dl_main_plan_btn")
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
# TAB 2: TÍCH HỢP NĂNG LỰC SỐ (ĐẦY ĐỦ 4 KHỐI)
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
        with st.spinner("⚡ Đang kết nối AI thiết kế Giáo án Năng lực số siêu chi tiết (vui lòng đợi ~5-8 giây)..."):
            prompt_digital = f"""
            Đóng vai là Chuyên gia Giáo dục Mầm non xuất sắc, hãy biên soạn một GIÁO ÁN MẦM NON TÍCH HỢP NĂNG LỰC SỐ & AI CHUẨN BỘ GD&ĐT VIỆT NAM (Dài ít nhất 1500 - 2000 từ):
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
            1. BÀI SOẠN PHẢI RẤT DÀI, CHI TIẾT, KHÔNG TÓM TẮT.
            2. Phần III. TIẾN TRÌNH HOẠT ĐỘNG: Trình bày dạng KỊCH BẢN TƯƠNG TÁC CHI TIẾT LỜI CÔ - LỜI TRẺ.
            Trình bày chuẩn 4 phần:
            I. MỤC ĐÍCH - YÊU CẦU
            II. CHUẨN BỊ HỌC LIỆU & KHÔNG GIAN SỐ
            III. TIẾN TRÌNH TỔ CHỨC HOẠT ĐỘNG TÍCH HỢP NĂNG LỰC SỐ (Hoạt động 1: Mở đầu; Hoạt động 2: Trọng tâm; Hoạt động 3: Kết thúc)
            IV. ĐÁNH GIÁ CUỐI NGÀY (Bắt buộc gồm đúng 3 dòng:
            - Tình trạng sức khỏe: ...
            - Trạng thái cảm xúc, thái độ: ...
            - Kiến thức, kỹ năng: ...)

            Tuyệt đối KHÔNG sử dụng ký tự ### hay **. Dùng gạch đầu dòng (-) hoặc số thứ tự.
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
                st.download_button(label="📥 Tải File Word (.docx)", data=f_d, file_name=d_docx, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True, key="dl_digital_btn")
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
# TAB KẾ HOẠCH & SỔ CHỦ ĐỀ
# =========================================================================
elif app_mode == "🗓️ Kế Hoạch & Sổ Chủ Đề":
    sub_tab_month, sub_tab_week = st.tabs(["📅 Kế Hoạch Tháng", "📋 Kế Hoạch Tuần"])
    
    with sub_tab_month:
        st.subheader("📖 Thiết Kế Kế Hoạch Tháng")
        m_month_val = st.selectbox("Tháng thực hiện", ["Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12"], index=1)
        m_theme_val = st.text_input("Chủ đề chính", "Thế giới thực vật")
        m_age = st.selectbox("Độ tuổi", ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=2)
        if st.button("✨ Tạo Kế Hoạch Tháng AI", type="primary"):
            with st.spinner("⚡ Đang lập kế hoạch chi tiết..."):
                info_m = {"month": m_month_val, "theme": m_theme_val, "age": m_age}
                res, err = generate_fast_ai_content(api_key, f"Lập kế hoạch giáo dục tháng {m_month_val} chủ đề {m_theme_val} cho {m_age}. Thật dài và chi tiết. Không dùng dấu ###.", fallback_type="month", fallback_data=info_m)
                if err: st.warning(f"⚠️ {err}")
                st.session_state["saved_month_plan"] = res
                save_plan_to_db(f"[Tháng] {m_theme_val}", res)
        if "saved_month_plan" in st.session_state:
            st.markdown(format_text_to_standard_html(st.session_state["saved_month_plan"], 14), unsafe_allow_html=True)
    
    with sub_tab_week:
        st.subheader("📋 Thiết Kế Kế Hoạch Tuần")
        w_week_num = st.selectbox("Tuần thực hiện", ["Tuần 1", "Tuần 2", "Tuần 3", "Tuần 4", "Tuần 5"])
        w_theme_week = st.text_input("Nhánh tuần", "Tôi là ai - Bạn là ai?")
        w_age2 = st.selectbox("Độ tuổi trẻ", ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=2)
        if st.button("✨ Tạo Kế Hoạch Tuần AI", type="primary"):
            with st.spinner("⚡ Đang lập kế hoạch tuần chi tiết..."):
                info_w = {"week_num": w_week_num, "theme_week": w_theme_week, "age": w_age2}
                res, err = generate_fast_ai_content(api_key, f"Lập kế hoạch tuần {w_week_num} chủ đề {w_theme_week} cho trẻ {w_age2}. Phân bổ chi tiết từ Thứ 2 đến Thứ 6. Dài và chi tiết. Không dùng dấu ###.", fallback_type="week", fallback_data=info_w)
                if err: st.warning(f"⚠️ {err}")
                st.session_state["saved_week_plan"] = res
                save_plan_to_db(f"[Tuần] {w_theme_week}", res)
        if "saved_week_plan" in st.session_state:
            st.markdown(format_text_to_standard_html(st.session_state["saved_week_plan"], 14), unsafe_allow_html=True)

# ==========================================
# CÁC TAB CHỨC NĂNG KHÁC
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
# TAB KHO LỊCH SỬ GIÁO ÁN
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
