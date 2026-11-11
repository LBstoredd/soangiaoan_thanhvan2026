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

# --- CẤU HÌNH GIAO DIỆN STREAMLIT & THEME XANH MINT HIỆN ĐẠI ---
st.set_page_config(page_title="Phần Mềm Trợ Lý Mầm Non Chuyên Sâu", layout="wide")

st.markdown("""
<style>
    /* Nền tổng thể Xanh Mint Trắng dịu mắt & Phông chữ hiện đại */
    .stApp {
        background: linear-gradient(180deg, #F4FBF9 0%, #FFFFFF 100%);
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        color: #1E293B;
    }

    /* Thanh menu bên trái (Sidebar) tông Xanh Mint Pastel */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #E6F7F2 0%, #D4F1E8 100%) !important;
        border-right: 1.5px solid #B2E5D6;
    }
    [data-testid="stSidebar"] * {
        color: #0F5247 !important;
    }

    /* Tiêu đề ứng dụng sắc nét */
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

    /* Nút bấm chính (Primary Buttons) & Nút Tải File Word */
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

    /* Nút tạo giáo án màu loang (Tím-Xanh) */
    .btn-create-magic button {
        background: linear-gradient(135deg, #8A2BE2 0%, #4B0082 100%) !important;
        color: #FFFFFF !important;
        border-radius: 14px !important;
        font-size: 16px !important;
        font-weight: 800 !important;
        padding: 12px 30px !important;
        box-shadow: 0 4px 15px rgba(138, 43, 226, 0.4) !important;
    }
    .btn-create-magic button:hover {
        background: linear-gradient(135deg, #9932CC 0%, #8A2BE2 100%) !important;
        transform: translateY(-2px);
        box-shadow: 0 6px 18px rgba(138, 43, 226, 0.5) !important;
    }

    /* Nút AI màu xanh dương */
    .btn-ai-suggest button {
        background: linear-gradient(135deg, #0984E3 0%, #005A9E 100%) !important;
        color: #FFFFFF !important;
        border-radius: 10px !important;
        font-size: 14px !important;
        font-weight: 700 !important;
        box-shadow: 0 3px 10px rgba(9, 132, 227, 0.3) !important;
    }

    /* Khung nhập liệu (Input, Textarea, Selectbox, NumberInput) */
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

    /* Thẻ chứa thông tin, Khung mở rộng & Tabs */
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

# --- BỘ LỌC LÀM SẠCH TỰ ĐỘNG ---
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

# --- BỘ ĐỊNH DẠNG HTML & XUẤT WORD CHUYÊN NGHIỆP ---
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
        if not line_s:
            continue
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
        <button id="copyBtn_{unique_key}" onclick="copyFormattedContent_{unique_key}()" style="
            flex: 1;
            background: linear-gradient(135deg, #0984E3, #00796B);
            color: white;
            border: none;
            padding: 10px 16px;
            border-radius: 12px;
            font-size: 15px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(9, 132, 227, 0.25);
            transition: all 0.2s ease;
        ">
            📋 Sao Chép Nhanh (HTML/Text)
        </button>
        <button onclick="printFormattedContent_{unique_key}()" style="
            flex: 1;
            background: linear-gradient(135deg, #00B894, #00897B);
            color: white;
            border: none;
            padding: 10px 16px;
            border-radius: 12px;
            font-size: 15px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(0, 184, 148, 0.28);
            transition: all 0.2s ease;
        ">
            🖨️ In Giáo Án Ngay
        </button>
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
            }} else {{
                await navigator.clipboard.writeText(plainContent);
            }}
            btn.innerHTML = "✅ Đã Sao Chép Chuẩn Word/Zalo!";
            btn.style.background = "linear-gradient(135deg, #2ECC71, #27AE60)";
            setTimeout(() => {{
                btn.innerHTML = "📋 Sao Chép Nhanh (HTML/Text)";
                btn.style.background = "linear-gradient(135deg, #0984E3, #00796B)";
            }}, 2500);
        }} catch (err) {{
            const ta = document.createElement("textarea");
            ta.value = plainContent;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
            btn.innerHTML = "✅ Đã Sao Chép Văn Bản!";
            setTimeout(() => {{
                btn.innerHTML = "📋 Sao Chép Nhanh (HTML/Text)";
            }}, 2500);
        }}
    }}

    function printFormattedContent_{unique_key}() {{
        const htmlBody = {html_json};
        const docTitle = {title_json};
        const pWin = window.open('', '_blank');
        pWin.document.write('<!DOCTYPE html><html><head><title>' + docTitle + '</title>');
        pWin.document.write('<style>');
        pWin.document.write('@page {{ size: A4; margin: 18mm 20mm 18mm 20mm; }}');
        pWin.document.write('body {{ font-family: "Times New Roman", Times, serif; font-size: {font_size_pt}pt; line-height: 1.5; color: #000; margin: 0; padding: 0; text-align: justify; }}');
        pWin.document.write('</style></head><body>');
        pWin.document.write(htmlBody);
        pWin.document.write('</body></html>');
        pWin.document.close();
        pWin.focus();
        setTimeout(function() {{
            pWin.print();
            pWin.close();
        }}, 350);
    }}
    </script>
    """
    components.html(component_code, height=48)

# --- KHỞI TẠO DATABASE ---
def init_db():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS saved_plans (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, content TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        cursor.execute('''CREATE TABLE IF NOT EXISTS licenses (id INTEGER PRIMARY KEY AUTOINCREMENT, license_key TEXT UNIQUE, customer_name TEXT, expires_at DATE, is_active INTEGER DEFAULT 1)''')
        cursor.execute('''INSERT OR IGNORE INTO licenses (id, license_key, customer_name, expires_at, is_active) VALUES (1, 'LBSTORE2026', 'Admin Văn Bình (Master Admin)', '2027-12-31', 1)''')
        conn.commit()
        conn.close()
    except Exception:
        pass
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

def get_all_licenses():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT id, license_key, customer_name, expires_at, is_active FROM licenses ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception: return []

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

# --- HÀM TẠO GIÁO ÁN NHANH ---
def generate_instant_pedagogical_plan(info: dict) -> str:
    today_str = datetime.now().strftime("%Y-%m-%d")
    topic = info.get("topic", "Hoạt động khám phá")
    domain = info.get("domain", "Phát triển nhận thức")
    sub_cat = info.get("activity_type", "Học có chủ đích")
    age = info.get("age", "4-5 tuổi")
    theme = info.get("theme", "Thế giới thực vật")
    teacher = info.get("teacher", "Giáo viên")
    school = info.get("school", "Trường Mầm Non")
    class_name = info.get("class_name", "Lớp Mẫu giáo")
    custom_goals = info.get("custom_goals", "").strip()
    custom_ideas = info.get("custom_ideas", "").strip()
    
    if custom_goals: goals_block = custom_goals
    else: goals_block = (f"1. Kiến thức:\n- Trẻ nhận biết, gọi đúng tên và nắm được đặc điểm cốt lõi của nội dung bài học: {topic}.\n2. Kỹ năng:\n- Rèn luyện kỹ năng quan sát, chú ý ghi nhớ có chủ định.\n3. Thái độ:\n- Trẻ hứng thú, tích cực tham gia hoạt động.")
    idea_note = f"\n- Đạo cụ & ý tưởng của giáo viên: {custom_ideas}" if custom_ideas else ""

    raw_plan = f"""{school.upper()}
Lớp: {class_name} | Độ tuổi: {age}
GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG
ĐỀ TÀI: {topic.upper()}
Chủ đề: {theme} | Lĩnh vực: {domain}
Người soạn: {teacher} | Ngày dạy: {today_str}

I. MỤC ĐÍCH - YÊU CẦU
{goals_block}

II. CHUẨN BỊ
1. Đồ dùng của cô:
- Giáo án điện tử, học liệu trực quan: {topic}.{idea_note}
2. Đồ dùng của trẻ:
- Trang phục gọn gàng, học liệu đủ cho mỗi trẻ.
3. Không gian tổ chức: Trong lớp học an toàn, sạch sẽ.

III. TỔ CHỨC HOẠT ĐỘNG
Hoạt động 1: Ổn định tổ chức, gây hứng thú
- Cô trò chuyện gợi mở dẫn dắt vào đề tài '{topic}'.

Hoạt động 2: Trọng tâm
- Bước 1: Khám phá và đàm thoại về {topic}.
- Bước 2: Trẻ thực hành và trải nghiệm trực tiếp.
- Bước 3: Tổ chức trò chơi củng cố kiến thức.

Hoạt động 3: Kết thúc
- Cô nhận xét, khen ngợi và chuyển hoạt động nhẹ nhàng.

IV. ĐÁNH GIÁ CUỐI NGÀY
- Tình trạng sức khỏe: Đa số trẻ khỏe mạnh, tham gia đầy đủ các hoạt động.
- Trạng thái cảm xúc, thái độ: Trẻ hào hứng, tích cực.
- Kiến thức, kỹ năng: Trẻ thực hiện tốt các yêu cầu bài học {topic}."""
    return clean_markdown_text(raw_plan)

def generate_fast_ai_content(client, prompt, fallback_type="plan", fallback_data=None):
    def _call_model():
        res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
        return res.text

    executor = ThreadPoolExecutor(max_workers=1)
    future = executor.submit(_call_model)
    try:
        text_out = future.result(timeout=12)
        executor.shutdown(wait=False)
        if text_out and text_out.strip():
            return clean_markdown_text(text_out)
    except Exception:
        executor.shutdown(wait=False)

    if fallback_type == "plan": return generate_instant_pedagogical_plan(fallback_data)
    return "Đã hoàn thành nội dung cho cô giáo!"

# --- KHỞI TẠO SESSION STATE ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if "is_admin" not in st.session_state: st.session_state["is_admin"] = False
if "api_key_saved" not in st.session_state: st.session_state["api_key_saved"] = ""

# --- MÀN HÌNH ĐĂNG NHẬP ---
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

# --- GIAO DIỆN CHÍNH ---
st.title("🧸 Trợ Lý Ảo Toàn Năng - Giáo Viên Mầm Non (LB Store)")
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
# TAB 1: SOẠN GIÁO ÁN CHUẨN (NÂNG CẤP GIAO DIỆN THEO ẢNH)
# ==========================================
if app_mode == "📝 Soạn Giáo Án Chuẩn":
    # KHỐI 1: THÔNG TIN BÀI DẠY
    with st.container(border=True):
        st.markdown("#### 1️⃣ Thông tin bài dạy")
        st.caption("Điền những thông tin cốt lõi để AI xây dựng giáo án đúng đối tượng.")
        
        info_c1, info_c2, info_c3 = st.columns(3)
        with info_c1:
            age_group = st.selectbox("Độ tuổi *", ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=2, key="g_age")
        with info_c2:
            theme = st.selectbox("Chủ đề giáo dục *", ["Thế giới thực vật", "Thế giới động vật", "Gia đình", "Trường mầm non", "Nghề nghiệp", "Giao thông", "Hiện tượng tự nhiên", "Quê hương đất nước"], index=0, key="g_theme")
        with info_c3:
            domain = st.selectbox("Lĩnh vực phát triển *", ["Phát triển nhận thức", "Phát triển thể chất", "Phát triển ngôn ngữ", "Phát triển thẩm mỹ", "Phát triển tình cảm & KNXH"], index=0, key="g_domain")
            
        info_c4, info_c5 = st.columns([1, 2])
        with info_c4:
            activity_type = st.selectbox("Loại hoạt động *", ["Học có chủ đích", "Hoạt động góc", "Hoạt động ngoài trời", "Hoạt động chiều"], key="g_type")
        with info_c5:
            topic = st.text_input("Tên hoạt động cụ thể *", value="Bé khám phá hành trình lớn lên của hạt đậu", key="g_topic")

        info_c6, info_c7, info_c8, info_c9 = st.columns(4)
        with info_c6:
            edu_model = st.selectbox("Mô hình tổ chức hoạt động", ["Không áp dụng", "Mô hình 5E", "Mô hình EDP", "Mô hình 5E + EDP"], index=1, key="g_model")
        with info_c7:
            plan_purpose = st.selectbox("Mục đích sử dụng giáo án", ["Giáo án hàng ngày", "Giáo án dự thi/thao giảng", "Giáo án chuyên đề", "Giáo án STEAM"], index=0, key="g_purp")
        with info_c8:
            kids_level = st.selectbox("Mức độ của trẻ", ["Đồng đều", "Có trẻ đặc biệt", "Có trẻ tiếp thu chậm"], index=0, key="g_level")
        with info_c9:
            duration = st.selectbox("Thời lượng", ["15 phút", "20 phút", "25 phút", "30 phút", "35 phút"], index=3, key="g_time")

        teacher_name = st.text_input("Tên giáo viên (Tùy chọn)", "Nguyễn Thị Thanh Vân", key="g_gv")
        school_name = st.text_input("Tên trường / Lớp (Tùy chọn)", "Trường Mầm Non Phước Thạnh - Lớp Chồi 1", key="g_tr")

    # KHỐI 2: MỨC ĐỘ TÍCH HỢP
    with st.container(border=True):
        st.markdown("#### 2️⃣ Mức độ tích hợp")
        st.caption("Chọn tối đa 3 nội dung. 'Cơ bản' không kết hợp với lựa chọn khác.")
        
        tc1, tc2, tc3, tc4, tc5 = st.columns(5)
        with tc1:
            t_coban = st.checkbox("Cơ bản", value=True, key="t_cb")
            t_ngonngu = st.checkbox("Tăng cường tiếng Việt", key="t_nn")
        with tc2:
            t_steam = st.checkbox("Tích hợp STEAM", key="t_st")
            t_knxh = st.checkbox("Kỹ năng xã hội", key="t_kn")
        with tc3:
            t_ai = st.checkbox("Tích hợp AI", key="t_ai")
            t_vd = st.checkbox("Tăng cường vận động", key="t_vd")
        with tc4:
            t_steam_nangcao = st.checkbox("STEAM + Kỹ năng cao", key="t_stnc")
        with tc5:
            t_antoan = st.checkbox("An toàn giao thông", key="t_at")

        extras = []
        if t_steam: extras.append("Tích hợp STEAM")
        if t_ai: extras.append("Tích hợp AI")
        if t_steam_nangcao: extras.append("STEAM + Kỹ năng cao")
        if t_antoan: extras.append("An toàn giao thông")
        if t_ngonngu: extras.append("Tăng cường tiếng Việt")
        if t_knxh: extras.append("Kỹ năng xã hội")
        if t_vd: extras.append("Tăng cường vận động")

    # KHỐI 3: MỤC TIÊU BÀI DẠY
    with st.container(border=True):
        m_c1, m_c2 = st.columns([4, 1])
        with m_c1:
            st.markdown("#### 3️⃣ Mục tiêu bài dạy")
            st.caption("AI có thể đề xuất mục tiêu chuẩn, giáo viên vẫn có thể chỉnh sửa trước khi tạo giáo án.")
        with m_c2:
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            if st.button("✨ AI gợi ý mục tiêu", key="btn_ai_sug"):
                st.session_state["g_goals"] = f"1. Kiến thức: Trẻ biết quá trình phát triển của {topic.lower()}.\n2. Kỹ năng: Rèn kỹ năng quan sát, so sánh và thực hành thao tác.\n3. Thái độ: Trẻ yêu thiên nhiên, tích cực tham gia hoạt động."
                st.rerun()

        custom_goals = st.text_area(
            "Mục tiêu đề xuất",
            value=st.session_state.get("g_goals", ""),
            placeholder="AI sẽ đề xuất mục tiêu theo Kiến thức - Kỹ năng - Thái độ - Năng lực hình thành...",
            height=120,
            key="g_goals_input"
        )

    # KHỐI 4 & 5: YÊU CẦU RIÊNG
    with st.container(border=True):
        st.markdown("#### 4️⃣ Yêu cầu riêng của giáo viên")
        st.caption("Ghi rõ nội dung bạn muốn nhấn mạnh hoặc lồng ghép thêm vào giáo án.")
        custom_ideas = st.text_area("Yêu cầu bổ sung", placeholder="Ví dụ: Thêm trò chơi đóng vai, sử dụng đạo cụ rối tay...", height=80, key="g_req")

    # NÚT BẤM TRUNG TÂM
    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)
    btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])
    with btn_col2:
        st.markdown("<div class='btn-create-magic'>", unsafe_allow_html=True)
        btn_create = st.button("🚀 SẴN SÀNG TẠO HỌC LIỆU - TẠO GIÁO ÁN HOÀN CHỈNH", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    plan_info_dict = {
        "school": school_name, "teacher": teacher_name, "class_name": "Lớp mầm non",
        "theme": theme, "domain": domain, "activity_type": activity_type,
        "route_type": "Hoạt động chung", "topic": topic, "method": edu_model,
        "age": age_group, "extras": extras, "custom_goals": custom_goals, 
        "custom_ideas": custom_ideas, "display_format": "1 Cột chuẩn", "auto_assessment": True
    }

    if btn_create:
        with st.spinner("⚡ Đang biên soạn giáo án 5E chuẩn sư phạm thông minh..."):
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    prompt = f"""
                    Soạn giáo án mầm non chuẩn Bộ GD&ĐT Việt Nam:
                    - Đề tài: {topic} | Độ tuổi: {age_group} | Chủ đề: {theme} | Lĩnh vực: {domain}
                    - Loại: {activity_type} | Mô hình: {edu_model} | Trẻ: {kids_level}
                    - Mục tiêu: {custom_goals if custom_goals else 'Tự động tạo chuẩn 3 mặt (Kiến thức, Kỹ năng, Thái độ)'}
                    - Tích hợp: {', '.join(extras) if extras else 'Cơ bản'}
                    - Yêu cầu riêng: {custom_ideas if custom_ideas else 'Không'}
                    
                    Trình bày chuẩn 4 mục lớn:
                    I. MỤC ĐÍCH - YÊU CẦU
                    II. CHUẨN BỊ (Cô & Trẻ)
                    III. TỔ CHỨC HOẠT ĐỘNG (Theo các bước {edu_model} hoặc 3 bước chuẩn)
                    IV. ĐÁNH GIÁ CUỐI NGÀY (Bắt buộc gồm đúng 3 dòng: - Tình trạng sức khỏe: ... ; - Trạng thái cảm xúc, thái độ: ... ; - Kiến thức, kỹ năng: ... )
                    Tuyệt đối không dùng dấu ### hoặc **.
                    """
                    result_text = generate_fast_ai_content(client, prompt, fallback_type="plan", fallback_data=plan_info_dict)
                except Exception:
                    result_text = generate_instant_pedagogical_plan(plan_info_dict)
            else:
                result_text = generate_instant_pedagogical_plan(plan_info_dict)

            result_text = clean_markdown_text(result_text)
            st.session_state["generated_plan_text"] = result_text
            st.session_state["plan_info"] = plan_info_dict
            save_plan_to_db(topic, result_text)
        st.success("🎉 Đã hoàn thành giáo án!")

    # HIỂN THỊ KẾT QUẢ VÀ XUẤT BẢN WORD
    if "generated_plan_text" in st.session_state:
        st.divider()
        info = st.session_state["plan_info"]
        current_font_size = st.session_state.get("font_size_slider", 14)
        
        st.markdown("### 🛠️ KẾT QUẢ AI & XUẤT BẢN")
        
        btn_c1, btn_c2 = st.columns([1, 2])
        with btn_c1:
            docx_file = export_docx_from_text(st.session_state["generated_plan_text"], info, font_size_pt=current_font_size)
            with open(docx_file, "rb") as file:
                st.download_button(
                    label="📥 Tải Xuất Word (.docx)",
                    data=file,
                    file_name=docx_file,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
        with btn_c2:
            render_copy_and_print_buttons(
                st.session_state["generated_plan_text"],
                f"Giáo Án: {info['topic']}",
                current_font_size,
                "tab1_main"
            )

        font_size_val = st.slider("🔤 Kích thước chữ hiển thị & In (pt):", min_value=12, max_value=24, value=current_font_size, step=1, key="font_size_slider")
        
        st.markdown("### 📄 Bản Xem Trước Giáo Án (Chuẩn In Ấn):")
        formatted_preview_html = format_text_to_standard_html(st.session_state["generated_plan_text"], font_size_val)
        st.markdown(
            f"""
            <div style="
                background-color: #ffffff;
                color: #000000;
                font-family: 'Times New Roman', Times, serif;
                font-size: {font_size_val}pt;
                line-height: 1.5;
                padding: 35px 45px;
                border-radius: 12px;
                border: 1.5px solid #B2E5D6;
                box-shadow: 0 6px 18px rgba(0, 121, 107, 0.08);
                text-align: justify;
            ">{formatted_preview_html}</div>
            """,
            unsafe_allow_html=True
        )

        with st.expander("✏️ Chỉnh sửa nội dung văn bản (Bấm vào để mở khung sửa chữ)"):
            edited_text = st.text_area("Nội dung giáo án:", st.session_state["generated_plan_text"], height=400)
            if edited_text != st.session_state["generated_plan_text"]:
                st.session_state["generated_plan_text"] = clean_markdown_text(edited_text)
                st.rerun()

# (GIỮ NGUYÊN CÁC TAB 2,3,4,5,6 TRƯỚC ĐÓ Ở PHẦN BÊN DƯỚI...)
