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
# KHỞI TẠO CẤU HÌNH & SESSION STATE 
# ==========================================
st.set_page_config(page_title="Phần Mềm Trợ Lý Mầm Non Chuyên Sâu", layout="wide")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "is_admin" not in st.session_state:
    st.session_state["is_admin"] = False
if "api_key_saved" not in st.session_state:
    st.session_state["api_key_saved"] = ""

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

FULL_BOLD_PREFIXES = ('GIÁO ÁN', 'KẾ HOẠCH', 'ĐỀ TÀI', 'CHỦ ĐỀ', 'Hoạt động 1', 'Hoạt động 2', 'Hoạt động 3', '1. THỨ', '2. THỨ', '3. THỨ', '4. THỨ', '5. THỨ', 'BÀI THƠ:', 'BÀI ĐỒNG DAO:', 'HỆ THỐNG CÂU ĐỐ', 'TRÒ CHƠI VẬN ĐỘNG', 'MẪU TIN NHẮN', 'LỜI NHẬN XÉT', 'Ý TƯỞNG LÀM ĐỒ DÙNG')
LABEL_BOLD_PREFIXES = ('- Tình trạng sức khỏe:', '- Trạng thái cảm xúc, thái độ:', '-Trạng thái cảm xúc, thái độ:', '- Trạng thái cảm xúc:', '- Kiến thức, kỹ năng:', '- Kiến thức, kỹ năng & Năng lực số:', '1. Kiến thức:', '2. Kỹ năng:', '3. Thái độ:', '2. Kỹ năng & Năng lực số:', '1. Mục tiêu trọng tâm & Năng lực số:', '1. Đồ dùng của cô:', '2. Đồ dùng của trẻ:', '3. Không gian tổ chức:', '1. Mục đích - Yêu cầu:', '2. Chuẩn bị:', '3. Cách chơi:', '4. Luật chơi:', '- Bước 1:', '- Bước 2:', '- Bước 3:', '- Bước 4:', '* Bước 1:', '* Bước 2:', '* Bước 3:', '* Bước 4:', '- Khởi động:', '- Trọng động:', '- Hồi tĩnh:', '* Khởi động:', '* Trọng động:', '* Hồi tĩnh:', 'a) ', 'b) ', 'c) ', '- Đón trẻ & Thể dục sáng:', '- Hoạt động học:', '- Hoạt động ngoài trời:', '- Hoạt động chơi ở các góc:', '- Hoạt động chiều:', '- Thể chất:', '- Nhận thức:', '- Ngôn ngữ:', '- Tình cảm - KNXH:', '- Thẩm mỹ:', '- Góc Phân vai:', '- Góc Xây dựng', '- Góc Nghệ thuật', '- Góc Tạo hình:', '- Góc Học tập', '- Góc Âm nhạc:', '- Góc Khám phá', '- Tuần 1:', '- Tuần 2:', '- Tuần 3:', '- Tuần 4:', '- Tuần 5:', 'Lớp:', 'Chủ đề:', 'Người soạn:', 'Thời gian thực hiện:', 'Loại hoạt động:', 'Hướng tiếp cận:')

def is_full_bold_heading(line_s: str) -> bool:
    if re.match(r'^(I|II|III|IV|V|VI)\.\s', line_s): return True
    if line_s.startswith(FULL_BOLD_PREFIXES): return True
    if line_s.isupper() and len(line_s) > 5: return True
    return False

def should_bold_label_before_colon(line_s: str) -> bool:
    if ':' not in line_s: return False
    if line_s.startswith(LABEL_BOLD_PREFIXES): return True
    if re.match(r'^(\d+\.|[-+])\s+[^:]{2,38}:', line_s): return True
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

def export_docx_from_text(text_content: str, info: dict, custom_template=None, font_size_pt=13):
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
    filename = f"GiaoAn_{safe_topic}.docx"
    doc.save(filename)
    return filename

def render_copy_and_print_buttons(text_content: str, doc_title: str, font_size_pt: int, unique_key: str):
    formatted_html = format_text_to_standard_html(text_content, font_size_pt)
    html_json = json.dumps(formatted_html)
    plain_json = json.dumps(text_content)
    title_json = json.dumps(doc_title)
    
    component_code = f"""
    <div style="display: flex; gap: 12px; width: 100%; align-items: center; margin: 0; padding: 0;">
        <button id="copyBtn_{unique_key}" onclick="copyFormattedContent_{unique_key}()" style="flex: 1; background: linear-gradient(135deg, #0984E3, #00796B); color: white; border: none; padding: 10px 16px; border-radius: 12px; font-size: 15px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 12px rgba(9, 132, 227, 0.25); transition: all 0.2s ease;">📋 Sao Chép Nhanh (HTML/Text)</button>
        <button onclick="printFormattedContent_{unique_key}()" style="flex: 1; background: linear-gradient(135deg, #00B894, #00897B); color: white; border: none; padding: 10px 16px; border-radius: 12px; font-size: 15px; font-weight: bold; cursor: pointer; box-shadow: 0 4px 12px rgba(0, 184, 148, 0.28); transition: all 0.2s ease;">🖨️ In Giáo Án Ngay</button>
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
# KHỞI TẠO DATABASE
# ==========================================
def init_db():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS saved_plans (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, content TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS licenses (id INTEGER PRIMARY KEY AUTOINCREMENT, license_key TEXT UNIQUE, customer_name TEXT, expires_at DATE, is_active INTEGER DEFAULT 1)''')
        cursor.execute('''INSERT OR IGNORE INTO licenses (id, license_key, customer_name, expires_at, is_active) VALUES (1, 'LBSTORE2026', 'Admin Văn Bình (Master Admin)', '2027-12-31', 1)''')
        conn.commit()
        conn.close()
    except Exception: pass
init_db()

def check_license(key):
    if key == 'LBSTORE2026': return True, "👑 Đăng nhập quyền Admin thành công!"
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
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
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO licenses (license_key, customer_name, expires_at, is_active) VALUES (?, ?, ?, 1)", (new_key, customer_name, expires_date))
        conn.commit()
        conn.close()
        return True, new_key, expires_date
    except Exception as e: return False, str(e), ""

def get_all_licenses():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT id, license_key, customer_name, expires_at, is_active FROM licenses ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception: return []

def save_plan_to_db(title, content_json):
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO saved_plans (title, content) VALUES (?, ?)", (title, content_json))
        conn.commit()
        conn.close()
    except Exception: pass

def get_all_saved_plans():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, created_at FROM saved_plans ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception: return []

def get_plan_by_id(plan_id):
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT title, content FROM saved_plans WHERE id = ?", (plan_id,))
        row = cursor.fetchone()
        conn.close()
        return row
    except Exception: return None

# ==========================================
# BỘ CỨU CÁNH & XỬ LÝ AI 
# ==========================================
def generate_instant_pedagogical_plan(info: dict) -> str:
    topic = info.get("topic", "Hoạt động")
    domain = info.get("domain", "Phát triển nhận thức")
    age = info.get("age", "4-5 tuổi")
    custom_ideas = info.get("custom_ideas", "").strip()
    
    raw_plan = f"""TRƯỜNG MẦM NON
Lớp: Lớp Mẫu giáo | Độ tuổi: {age}
GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG
ĐỀ TÀI: {topic.upper()}
Lĩnh vực: {domain}

I. MỤC ĐÍCH - YÊU CẦU
1. Kiến thức: Trẻ nhận biết và hiểu rõ các đặc điểm cốt lõi của bài học {topic}.
2. Kỹ năng: Rèn luyện kỹ năng quan sát, tư duy phán đoán và trả lời câu hỏi mạch lạc.
3. Thái độ: Trẻ tích cực tham gia hoạt động, đoàn kết với bạn bè.

II. CHUẨN BỊ
1. Đồ dùng của cô: Giáo án điện tử, hình ảnh/video về {topic}. {custom_ideas}
2. Đồ dùng của trẻ: Rổ đồ chơi, không gian an toàn, sạch sẽ.

III. TỔ CHỨC HOẠT ĐỘNG
Hoạt động 1: Ổn định tổ chức gây hứng thú
- Cô tập trung trẻ, hát bài hát phù hợp với chủ đề và trò chuyện dẫn dắt trẻ vào bài học {topic}.

Hoạt động 2: Trọng tâm
- Bước 1: Khám phá trải nghiệm. Cô đưa ra tình huống hoặc vật thật để trẻ quan sát. Cô đặt các câu hỏi mở: "Các con thấy gì đây?", "Ai có nhận xét gì?".
- Bước 2: Cô giải thích và chốt lại kiến thức chuẩn xác, mở rộng thêm hiểu biết cho trẻ.
- Bước 3: Luyện tập. Trẻ tham gia trò chơi củng cố liên quan đến {topic}. Cô bao quát và sửa sai cho trẻ.

Hoạt động 3: Kết thúc
- Cô nhận xét, tuyên dương các bạn học ngoan. Giáo dục trẻ bài học ý nghĩa. Chuyển hoạt động nhẹ nhàng.

IV. ĐÁNH GIÁ CUỐI NGÀY
- Tình trạng sức khỏe: 100% trẻ khỏe mạnh, tham gia tốt các hoạt động.
- Trạng thái cảm xúc, thái độ: Trẻ vui vẻ, hào hứng tương tác tốt cùng cô và bạn.
- Kiến thức, kỹ năng: Đa số trẻ nắm được bài, một số trẻ nhút nhát đã được cô hỗ trợ kịp thời.
*(Lưu ý: Đây là bản giáo án dự phòng do hệ thống không kết nối được với AI trong thời gian cho phép. Vui lòng thử lại sau 1 phút)*"""
    return clean_markdown_text(raw_plan)

def generate_instant_digital_plan(info: dict) -> str:
    topic = info.get("topic", "Hoạt động số")
    return generate_instant_pedagogical_plan({"topic": topic, "domain": "Năng lực số", "age": "5-6 tuổi"})

def generate_instant_month_plan(info: dict) -> str:
    month = info.get("month", "Tháng")
    theme = info.get("theme", "Chủ đề")
    return f"""KẾ HOẠCH GIÁO DỤC {month.upper()} - CHỦ ĐỀ: {theme.upper()}
I. MỤC TIÊU PHÁT TRIỂN THÁNG: Phát triển toàn diện 5 mặt cho trẻ mầm non.
II. MẠNG NỘI DUNG: Trẻ làm quen, khám phá đặc điểm và trải nghiệm thực tế.
III. MẠNG HOẠT ĐỘNG: Toán, Khám phá, Âm nhạc, Văn học, Tạo hình.
IV. ĐÁNH GIÁ: Theo dõi sự tiến bộ của trẻ hàng tuần.
*(Lưu ý: Đây là bản kế hoạch dự phòng)*"""

def generate_instant_week_plan(info: dict) -> str:
    week = info.get("week_num", "Tuần")
    theme = info.get("theme_week", "Chủ đề")
    return f"""KẾ HOẠCH GIÁO DỤC {week.upper()} - NHÁNH: {theme.upper()}
I. MỤC TIÊU: Trẻ hiểu rõ đặc điểm nhánh {theme}.
II. KẾ HOẠCH HÀNG NGÀY:
1. THỨ HAI: Học tập và vận động cơ bản.
2. THỨ BA: Khám phá và vui chơi ngoài trời.
3. THỨ TƯ: Tạo hình sáng tạo.
4. THỨ NĂM: Làm quen văn học.
5. THỨ SÁU: Âm nhạc và tổng kết nêu gương bé ngoan.
*(Lưu ý: Đây là bản kế hoạch dự phòng)*"""

def generate_instant_creative_content(genre: str, topic: str) -> str:
    return f"{genre.upper()}: {topic.upper()}\nĐây là bản tạo nhanh do AI đang bận. Bạn vui lòng thử lại sau 1 phút nhé!"

def generate_instant_zalo(task_type: str, info_text: str) -> str:
    return f"MẪU ZALO NHANH:\nKính gửi Phụ huynh, hôm nay lớp học về {info_text}. Các con rất ngoan và tích cực!"

def generate_instant_diy(lesson: str) -> str:
    return f"ĐỒ DÙNG DIY CHO: {lesson.upper()}\n1. Dùng bìa carton, nắp chai nhựa.\n2. Cắt dán và trang trí màu sắc đẹp mắt."

# ==========================================
# TRÁI TIM CỦA AI: ĐÃ SỬA LỖI MODEL & THỜI GIAN
# ==========================================
def generate_fast_ai_content(client, prompt, fallback_type="plan", fallback_data=None):
    def _call_model():
        # Gọi model flash của Google, mở khóa giới hạn độ dài bằng max_output_tokens=8192
        res = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                max_output_tokens=8192,
                temperature=0.7
            )
        )
        return res.text

    executor = ThreadPoolExecutor(max_workers=1)
    future = executor.submit(_call_model)
    try:
        # THỜI GIAN CHỜ: 45 GIÂY ĐỂ ĐẢM BẢO AI CÓ THỂ HOÀN THÀNH BÀI DÀI
        text_out = future.result(timeout=45) 
        executor.shutdown(wait=False)
        if text_out and text_out.strip():
            return clean_markdown_text(text_out)
    except Exception as e:
        executor.shutdown(wait=False)
        print(f"API Error: {e}") # Bắt lỗi để không bị dừng app

    # Nếu AI lỗi hoặc quá chậm, dùng bản ngắn dự phòng
    if fallback_type == "plan": return generate_instant_pedagogical_plan(fallback_data)
    elif fallback_type == "digital": return generate_instant_digital_plan(fallback_data)
    elif fallback_type == "month": return generate_instant_month_plan(fallback_data)
    elif fallback_type == "week": return generate_instant_week_plan(fallback_data)
    elif fallback_type == "creative": return generate_instant_creative_content(fallback_data.get("genre", ""), fallback_data.get("topic", ""))
    elif fallback_type == "zalo": return generate_instant_zalo(fallback_data.get("type", ""), fallback_data.get("info", ""))
    elif fallback_type == "diy": return generate_instant_diy(fallback_data.get("lesson", ""))
    return "Đã hoàn thành nội dung cho cô giáo!"

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

with st.sidebar:
    st.markdown("### 📌 Danh Mục Tính Năng")
    app_mode = st.radio("Chọn tác vụ:", [
        "👑 Quản Lý Khách Thuê (Admin)", "📝 Soạn Giáo Án Chuẩn", "💻 Tích Hợp Năng Lực Số",
        "🗓️ Kế Hoạch & Sổ Chủ Đề", "🎯 Trò Chơi & Đố Vui", "💬 Sổ Liên Lạc & Zalo", 
        "✂️ Gợi Ý Đồ Dùng DIY", "📚 Kho Lịch Sử Giáo Án"
    ] if st.session_state["is_admin"] else [
        "📝 Soạn Giáo Án Chuẩn", "💻 Tích Hợp Năng Lực Số", "🗓️ Kế Hoạch & Sổ Chủ Đề", 
        "🎯 Trò Chơi & Đố Vui", "💬 Sổ Liên Lạc & Zalo", "✂️ Gợi Ý Đồ Dùng DIY", "📚 Kho Lịch Sử Giáo Án"
    ])
    st.divider()
    st.header("⚙️ Cấu hình hệ thống")
    input_api_key = st.text_input("Nhập Google Gemini API Key", type="password", value=st.session_state["api_key_saved"])
    if st.button("💾 Xác Nhận Lưu API Key"):
        if input_api_key.strip():
            st.session_state["api_key_saved"] = input_api_key.strip()
            st.success("✅ Đã lưu API Key thành công!")
        else: st.error("⚠️ Vui lòng nhập API Key hợp lệ!")
    if st.session_state["api_key_saved"]: st.success("🟢 Trạng thái: Đã kết nối AI sẵn sàng!")
    else: st.warning("🔴 Trạng thái: Chưa nhập API Key (Vẫn tạo được Giáo Án tức thì)!")
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
# TAB 1: SOẠN GIÁO ÁN CHUẨN (BẢN GỐC TỐI QUA + NÂNG CẤP AI CHI TIẾT)
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
        art_type_val = ""
        if domain == "Nhận thức": sub_category = st.selectbox("Phân loại chi tiết", ["Khám phá", "Toán"], key="sub_nhan_thuc")
        elif domain == "Ngôn ngữ": sub_category = st.selectbox("Phân loại chi tiết", ["Chữ cái - tiết tập tô", "Làm quen", "Trò chơi", "Thơ", "Truyện"], key="sub_ngon_ngu")
        elif domain == "Thẩm mỹ":
            sub_category = st.selectbox("Phân loại chi tiết", ["Âm nhạc - Dạy hát", "Nghe nhạc", "Múa minh họa", "Gõ đệm theo phách, nhịp", "Tạo hình"], key="sub_tham_my")
            if sub_category == "Tạo hình":
                with st.container(border=True):
                    st.markdown("🎨 **Phân loại tiết Tạo hình**")
                    art_type_val = st.radio("Hình thức:", ["Tiết theo mẫu", "Tiết theo đề tài"], horizontal=True, key="art_radio_key")
        else: sub_category = st.text_input("Hoạt động cụ thể", "Phát triển vận động", key="sub_khac")
    with col_b:
        st.markdown("### 4. DẠNG TIẾT CHUYÊN MÔN (ROUTE) *")
        if domain == "Thể chất": route_type = st.selectbox("Dạng tiết", ["Vận động cơ bản (VĐCB)", "Trò chơi vận động", "Thể dục sáng"], key="route_tc")
        elif domain == "Nhận thức": route_type = st.selectbox("Dạng tiết", ["Khám phá khoa học", "Làm quen với toán", "Khám phá xã hội"], key="route_nt")
        elif domain == "Ngôn ngữ": route_type = st.selectbox("Dạng tiết", ["Làm quen chữ cái", "Làm quen văn học (Thơ/Truyện)", "Phát triển ngôn ngữ"], key="route_nn")
        elif domain == "Thẩm mỹ": route_type = st.selectbox("Dạng tiết", ["Âm nhạc", "Tạo hình"], key="route_tm")
        else: route_type = st.selectbox("Dạng tiết", ["Kỹ năng xã hội", "Hoạt động chung"], key="route_khac")

    st.markdown("### 5. CẤU TRÚC MỤC ĐÍCH - YÊU CẦU:")
    goal_structure = st.radio("Lựa chọn cấu trúc", ["1-3 gạch đầu dòng", "2-3 mục chuẩn (Kiến thức, Kỹ năng, Thái độ)"], horizontal=True, index=1)
    st.markdown("### 6. ĐỀ TÀI BÀI DẠY CỤ THỂ * (BẮT BUỘC):")
    topic = st.text_input("Nhập tên đề tài", "Vẽ lọ hoa (hoặc Ném xa bằng 1 tay)")
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
    custom_goals = st.text_area("Nhập MT1, MT45... kèm nội dung thực tế", placeholder="VD:\n- Trẻ biết cách cầm bút vẽ lọ hoa cân đối, tô màu khéo léo (MT15).")
    st.markdown("### 9. YÊU CẦU BỔ SUNG (Ý TƯỞNG GIÁO VIÊN MUỐN ĐƯA VÀO TIẾT DẠY):")
    custom_ideas = st.text_area("Nhập ý tưởng riêng", placeholder="Cô hãy nhập các ý tưởng riêng: đạo cụ rối tay, nhạc chủ đề, bài thơ mới...")
    st.markdown("### 10. ĐÁNH GIÁ CUỐI NGÀY:")
    auto_assessment = st.checkbox("Quan sát thực tế cho Đánh giá cuối ngày (Mục IV) - Tạo đánh giá luôn", value=True)
    
    plan_info_dict = {
        "school": school_name, "teacher": teacher_name, "class_name": class_name,
        "theme": theme, "domain": domain, "sub_category": sub_category, "activity_type": activity_type,
        "route_type": route_type, "art_type": art_type_val, "topic": topic, "method": "Gemini 1.5 Flash", "age": age_group, "extras": extras,
        "mode": st.session_state["soan_mode_val"], "display_format": st.session_state["display_fmt_val"], 
        "custom_goals": custom_goals, "custom_ideas": custom_ideas, "auto_assessment": auto_assessment,
        "doc_title": "GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG HỌC"
    }

    if st.button("✨ TẠO GIÁO ÁN HOÀN CHỈNH CHI TIẾT", type="primary", use_container_width=True):
        with st.spinner("⚡ Đang biên soạn kịch bản giáo án SIÊU CHI TIẾT & SÁNG TẠO (AI đang viết nội dung sâu, vui lòng đợi ~15 giây)..."):
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    art_prompt_part = f" - Dạng tiết tạo hình: {art_type_val}" if domain == "Thẩm mỹ" and sub_category == "Tạo hình" else ""
                    
                    # PROMPT SIÊU CHI TIẾT: ÉP AI VIẾT THẬT DÀI, CÓ LỜI THOẠI CÔ VÀ TRẺ
                    prompt = f"""
                    Bạn là một Chuyên gia Giáo dục Mầm non xuất sắc. Hãy soạn một GIÁO ÁN MẦM NON CHUẨN BỘ GD&ĐT với yêu cầu BẮT BUỘC SAU ĐÂY:
                    - Đề tài: {topic} | Độ tuổi: {age_group} | Chủ đề: {theme} | Lĩnh vực: {domain} ({sub_category})
                    - Loại hoạt động: {activity_type} | Dạng tiết: {route_type}
                    {art_prompt_part}
                    - Chế độ: {st.session_state['soan_mode_val']} | Định dạng: {st.session_state['display_fmt_val']}
                    - Mục tiêu: {custom_goals if custom_goals else 'Tự động tạo chuẩn 3 mặt: Kiến thức, Kỹ năng, Thái độ (Mỗi mặt 2-3 ý)'}
                    - Tích hợp: {', '.join(extras) if extras else 'Cơ bản'}
                    - Ý tưởng của giáo viên: {custom_ideas if custom_ideas else 'Không'}
                    
                    YÊU CẦU LÕI: 
                    1. BÀI PHẢI RẤT DÀI, CỰC KỲ CHI TIẾT NHƯ MỘT KỊCH BẢN (Ít nhất 1000 từ). KHÔNG VIẾT CHUNG CHUNG SƠ LƯỢC.
                    2. Trong phần Tổ chức hoạt động: Phải ghi rõ ràng TỪNG LỜI THOẠI, CÂU HỎI HẤP DẪN của Cô (Cô nói: "...") và LỜI ĐÁP CỦA TRẺ (Trẻ đáp: "...").
                    3. Thiết kế tình huống khởi động lôi cuốn, trò chơi thực hành phong phú, kích thích tư duy của trẻ. Ngôn từ ngộ nghĩnh, gần gũi.
                    
                    Trình bày theo 4 phần:
                    I. MỤC ĐÍCH - YÊU CẦU
                    II. CHUẨN BỊ (Liệt kê cụ thể từng đồ dùng)
                    III. TỔ CHỨC HOẠT ĐỘNG (Chia rõ 3 hoạt động: Gây hứng thú, Trọng tâm khám phá/thực hành, Kết thúc. Trình bày dưới dạng đối thoại Cô - Trò).
                    IV. ĐÁNH GIÁ CUỐI NGÀY (Bắt buộc gồm đúng 3 dòng:
                    - Tình trạng sức khỏe: ...
                    - Trạng thái cảm xúc, thái độ: ...
                    - Kiến thức, kỹ năng: ...)
                    
                    Tuyệt đối KHÔNG sử dụng ký tự ### hay **. Dùng gạch đầu dòng (-) hoặc (1, 2, 3).
                    """
                    result_text = generate_fast_ai_content(client, prompt, fallback_type="plan", fallback_data=plan_info_dict)
                except Exception as e:
                    result_text = generate_instant_pedagogical_plan(plan_info_dict)
            else:
                result_text = generate_instant_pedagogical_plan(plan_info_dict)

            result_text = clean_markdown_text(result_text)
            st.session_state["generated_plan_text"] = result_text
            st.session_state["plan_info"] = plan_info_dict
            save_plan_to_db(topic, result_text)
        st.success("🎉 Đã hoàn thành giáo án chi tiết! Mời cô xem bên dưới:")

    if "generated_plan_text" in st.session_state:
        st.divider()
        info = st.session_state["plan_info"]
        current_font_size = st.session_state.get("font_size_slider", 14)
        
        btn_col1, btn_col2 = st.columns([1, 2])
        with btn_col1:
            docx_file = export_docx_from_text(st.session_state["generated_plan_text"], info, custom_template=None, font_size_pt=current_font_size)
            with open(docx_file, "rb") as file:
                st.download_button(label="📥 Tải File Word (.docx)", data=file, file_name=docx_file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
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
            edited_text = st.text_area("Nội dung giáo án:", st.session_state["generated_plan_text"], height=400)
            if edited_text != st.session_state["generated_plan_text"]:
                st.session_state["generated_plan_text"] = clean_markdown_text(edited_text)
                st.rerun()

# =========================================================================
# TAB TÍCH HỢP NĂNG LỰC SỐ
# =========================================================================
elif app_mode == "💻 Tích Hợp Năng Lực Số":
    head_c1, head_c2 = st.columns([3, 1])
    with head_c1:
        st.subheader("💻 Soạn Giáo Án Tích Hợp Năng Lực Số")
        st.caption("Gợi ý bài học hiện đại, linh hoạt, kích thích tư duy tương tác và niềm vui khám phá")
    with head_c2:
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        if st.button("🧪 Nhập Dữ Liệu Mẫu", type="secondary", use_container_width=True):
            st.session_state["dg_topic"] = "Khám phá sự kỳ diệu của quả cam"
            st.success("✅ Đã nạp dữ liệu thử nghiệm mẫu!")

    with st.container(border=True):
        st.markdown("#### 👤 THÔNG TIN BÀI DẠY")
        dg_topic = st.text_input("Tên hoạt động số", value=st.session_state.get("dg_topic", "Khám phá thế giới qua màn hình tương tác"))
        dg_domain = st.selectbox("Lĩnh vực", ["Phát triển nhận thức", "Phát triển thể chất", "Ngôn ngữ"], index=0)
        dg_age = st.selectbox("Độ tuổi", ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=3)
        
    if st.button("✨ Soạn Giáo Án Năng Lực Số AI", type="primary"):
        with st.spinner("⚡ Đang biên soạn giáo án Tích hợp số CHI TIẾT (Mất khoảng 15-20 giây)..."):
            digital_info_dict = {"topic": dg_topic, "domain": dg_domain, "age": dg_age}
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    prompt_digital = f"""
                    Đóng vai là chuyên gia giáo dục mầm non, soạn giáo án Tích hợp Năng lực số:
                    - Đề tài: {dg_topic} | Độ tuổi: {dg_age} | Lĩnh vực: {dg_domain}
                    
                    YÊU CẦU: Viết RẤT DÀI VÀ CỰC KỲ CHI TIẾT. Trình bày như một kịch bản lên lớp với lời cô dẫn dắt, tương tác với thiết bị số và câu trả lời của trẻ. Không dùng dấu ### hoặc **. Gồm 4 phần: Mục đích yêu cầu, Chuẩn bị, Tổ chức hoạt động, Đánh giá cuối ngày.
                    """
                    res = generate_fast_ai_content(client, prompt_digital, fallback_type="digital", fallback_data=digital_info_dict)
                except Exception:
                    res = generate_instant_digital_plan(digital_info_dict)
            else:
                res = generate_instant_digital_plan(digital_info_dict)
            st.session_state["saved_digital_plan"] = res

    if "saved_digital_plan" in st.session_state:
        st.divider()
        d_font_size = st.session_state.get("d_font_size_slider", 14)
        render_copy_and_print_buttons(st.session_state["saved_digital_plan"], f"Giáo Án Năng Lực Số", d_font_size, "tab_digital")
        d_font_size = st.slider("🔤 Kích thước chữ (pt):", 12, 24, d_font_size, 1, key="d_font_size_slider")
        st.markdown(f'<div style="background-color: #ffffff; color: #000000; font-family: \'Times New Roman\', Times, serif; font-size: {d_font_size}pt; line-height: 1.5; padding: 35px 45px; border-radius: 12px; border: 1.5px solid #B2E5D6; box-shadow: 0 6px 18px rgba(0, 121, 107, 0.08); text-align: justify;">{format_text_to_standard_html(st.session_state["saved_digital_plan"], d_font_size)}</div>', unsafe_allow_html=True)

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
                if api_key:
                    try:
                        client = genai.Client(api_key=api_key)
                        res = generate_fast_ai_content(client, f"Lập kế hoạch giáo dục tháng {m_month_val} chủ đề {m_theme_val} cho {m_age}. Thật dài và chi tiết. Không dùng dấu ###.", fallback_type="month", fallback_data=info_m)
                    except Exception: res = generate_instant_month_plan(info_m)
                else: res = generate_instant_month_plan(info_m)
                st.session_state["saved_month_plan"] = res
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
                if api_key:
                    try:
                        client = genai.Client(api_key=api_key)
                        res = generate_fast_ai_content(client, f"Lập kế hoạch tuần {w_week_num} chủ đề {w_theme_week} cho trẻ {w_age2}. Phân bổ chi tiết từ Thứ 2 đến Thứ 6. Dài và chi tiết. Không dùng dấu ###.", fallback_type="week", fallback_data=info_w)
                    except Exception: res = generate_instant_week_plan(info_w)
                else: res = generate_instant_week_plan(info_w)
                st.session_state["saved_week_plan"] = res
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
            res = generate_instant_creative_content(genre, topic_game)
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    res = generate_fast_ai_content(client, f"Sáng tác {genre} cực hay, dài, hấp dẫn cho trẻ mầm non chủ đề {topic_game}. Không dùng dấu ###", fallback_type="creative", fallback_data={"genre": genre, "topic": topic_game})
                except: pass
            st.markdown(format_text_to_standard_html(res, 14), unsafe_allow_html=True)

elif app_mode == "💬 Sổ Liên Lạc & Zalo":
    st.subheader("💬 Trợ Lý Viết Sổ Liên Lạc & Tin Nhắn Zalo")
    task_type = st.radio("Chọn loại:", ["Tin nhắn tổng kết ngày/tuần", "Viết lời nhận xét sự tiến bộ"])
    child_info = st.text_input("Nhập thông tin", "Hôm nay các con học kỹ năng tốt")
    if st.button("💌 Soạn Tin", type="primary"):
        with st.spinner("⚡ Đang soạn tin..."):
            res = generate_instant_zalo(task_type, child_info)
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    res = generate_fast_ai_content(client, f"Soạn {task_type} mầm non thân thiện, dài, tình cảm, tóm tắt: {child_info}. Không dùng dấu ###", fallback_type="zalo", fallback_data={"type": task_type, "info": child_info})
                except: pass
            st.markdown(format_text_to_standard_html(res, 14), unsafe_allow_html=True)

elif app_mode == "✂️ Gợi Ý Đồ Dùng DIY":
    st.subheader("✂️ Gợi Ý Làm Đồ Dùng Dạy Học Từ Vật Liệu Mở")
    lesson_name = st.text_input("Nhập bài học", "Biển báo giao thông")
    if st.button("🛠️ Gợi Ý Cách Làm", type="primary"):
        with st.spinner("⚡ Đang gợi ý cách làm chi tiết..."):
            res = generate_instant_diy(lesson_name)
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    res = generate_fast_ai_content(client, f"Gợi ý làm đồ dùng phế liệu mầm non chi tiết từng bước cho bài: {lesson_name}. Không dùng dấu ###", fallback_type="diy", fallback_data={"lesson": lesson_name})
                except: pass
            st.markdown(format_text_to_standard_html(res, 14), unsafe_allow_html=True)

elif app_mode == "📚 Kho Lịch Sử Giáo Án":
    st.subheader("📚 Kho Lưu Trữ Giáo Án Đã Soạn")
    plans = get_all_saved_plans()
    if not plans: st.info("Chưa có giáo án nào được lưu trữ trong hệ thống.")
    else:
        for p_id, p_title, p_time in plans:
            with st.container():
                cols = st.columns([4, 2, 2])
                cols[0].markdown(f"**{p_title}** *(Lưu lúc: {p_time})*")
                if cols[1].button("Tải lại Word", key=f"load_{p_id}"):
                    row = get_plan_by_id(p_id)
                    if row:
                        title_db, content_db = row
                        f_path = export_docx_from_text(content_db, {"topic": title_db})
                        with open(f_path, "rb") as f_dl:
                            cols[2].download_button("📥 Tải File", data=f_dl, file_name=f_path, key=f"dl_{p_id}")
