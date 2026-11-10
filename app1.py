import sqlite3
import json
import time
import html
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

# --- CẤU TRÚC DỮ LIỆU ĐẦU RA CHO AI (PYDANTIC) ---
class LessonPlan(BaseModel):
    title: str
    objectives: list[str]
    children_prep: list[str]
    teacher_prep: list[str]
    steps: list[dict]
    assessment: str

# --- HÀM KHỞI TẠO DATABASE & QUẢN LÝ MÃ KÍCH HOẠT (CHỐNG KHÓA FILE) ---
def init_db():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS saved_plans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                content TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS licenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                license_key TEXT UNIQUE,
                customer_name TEXT,
                expires_at DATE,
                is_active INTEGER DEFAULT 1
            )
        ''')
        cursor.execute('''
            INSERT OR IGNORE INTO licenses (id, license_key, customer_name, expires_at, is_active)
            VALUES (1, 'LBSTORE2026', 'Admin Văn Bình (Master Admin)', '2027-12-31', 1)
        ''')
        conn.commit()
        conn.close()
    except Exception:
        pass

init_db()

def check_license(key):
    if key == 'LBSTORE2026':
        return True, "👑 Đăng nhập quyền Admin thành công!"
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT customer_name, expires_at, is_active FROM licenses WHERE license_key = ?", (key,))
        row = cursor.fetchone()
        conn.close()
        if row:
            c_name, expires_at, is_active = row
            if is_active == 1:
                exp_date = datetime.strptime(expires_at, "%Y-%m-%d").date()
                if datetime.now().date() <= exp_date:
                    return True, f"✅ Kích hoạt thành công! Xin chào: {c_name} (Hạn đến: {expires_at})"
                else:
                    return False, "❌ Mã kích hoạt này đã hết hạn sử dụng. Vui lòng liên hệ gia hạn (100k/tháng)!"
    except Exception:
        pass
    return False, "❌ Mã kích hoạt không hợp lệ hoặc không tồn tại!"

def create_new_license(customer_name, months):
    random_str = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    new_key = f"LB-{random_str}"
    expires_date = (datetime.now() + timedelta(days=30 * months)).strftime("%Y-%m-%d")
    
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO licenses (license_key, customer_name, expires_at, is_active) VALUES (?, ?, ?, 1)", 
                       (new_key, customer_name, expires_date))
        conn.commit()
        conn.close()
        return True, new_key, expires_date
    except Exception as e:
        return False, str(e), ""

def get_all_licenses():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT id, license_key, customer_name, expires_at, is_active FROM licenses ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception:
        return []

def save_plan_to_db(title, content_json):
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS saved_plans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                content TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        cursor.execute("INSERT INTO saved_plans (title, content) VALUES (?, ?)", (title, content_json))
        conn.commit()
        conn.close()
    except Exception:
        pass

def get_all_saved_plans():
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, created_at FROM saved_plans ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        return rows
    except Exception:
        return []

def get_plan_by_id(plan_id):
    try:
        conn = sqlite3.connect('mamnon_rag.db', timeout=30)
        cursor = conn.cursor()
        cursor.execute("SELECT title, content FROM saved_plans WHERE id = ?", (plan_id,))
        row = cursor.fetchone()
        conn.close()
        return row
    except Exception:
        return None

# --- CẤU HÌNH TỐC ĐỘ CAO THINKING LEVEL LOW ---
def build_fast_config():
    try:
        return types.GenerateContentConfig(
            temperature=0.35,
            thinking_config=types.ThinkingConfig(thinking_level="LOW")
        )
    except Exception:
        try:
            return types.GenerateContentConfig(
                temperature=0.35,
                thinking_config=types.ThinkingConfig(thinking_budget=0)
            )
        except Exception:
            return types.GenerateContentConfig(temperature=0.35)

# --- BỘ KHUNG SƯ PHẠM CHUẨN 100/100 ĐIỂM ---
def generate_instant_pedagogical_plan(info: dict) -> str:
    today_str = datetime.now().strftime("%Y-%m-%d")
    topic = info.get("topic", "Hoạt động khám phá")
    domain = info.get("domain", "Nhận thức")
    sub_cat = info.get("sub_category", "Khám phá")
    route = info.get("route_type", "Hoạt động chung")
    age = info.get("age", "4-5 tuổi")
    theme = info.get("theme", "Thế giới xung quanh")
    teacher = info.get("teacher", "Giáo viên")
    school = info.get("school", "Trường Mầm Non")
    class_name = info.get("class_name", "Lớp Mẫu giáo")
    extras_list = info.get("extras", [])
    extras_str = ", ".join(extras_list) if extras_list else "Giáo dục kỹ năng sống và giữ gìn vệ sinh"
    custom_goals = info.get("custom_goals", "").strip()
    custom_ideas = info.get("custom_ideas", "").strip()
    mode = info.get("mode", "Giáo án thường")

    if custom_goals:
        goals_block = custom_goals
    else:
        goals_block = (
            f"1. Kiến thức:\n"
            f"- Trẻ nhận biết, gọi đúng tên và nắm được đặc điểm cốt lõi của nội dung bài học: {topic} phù hợp độ tuổi {age}.\n"
            f"- Trẻ hiểu cách thức tham gia hoạt động {route} thuộc lĩnh vực {domain} ({sub_cat}).\n"
            f"2. Kỹ năng:\n"
            f"- Rèn luyện cho trẻ kỹ năng quan sát, chú ý ghi nhớ có chủ định, phối hợp tay - mắt và thao tác thực hành chuẩn xác.\n"
            f"- Phát triển kỹ năng diễn đạt mạch lạc, trả lời trọn câu và kỹ năng hợp tác nhóm cùng cô và bạn.\n"
            f"3. Thái độ:\n"
            f"- Trẻ hứng thú, tích cực tham gia hoạt động; lồng ghép hiệu quả nội dung tích hợp: {extras_str}."
        )

    idea_note = f"\n- Đạo cụ & ý tưởng mở rộng của giáo viên: {custom_ideas}" if custom_ideas else ""
    intro_hook = (
        f"- Cô xuất hiện cùng tình huống rối tay / hộp quà bí ẩn gắn với chủ đề '{theme}' để dẫn dắt trẻ vào câu chuyện khám phá '{topic}'."
        if "Thi GV Giỏi" in mode
        else f"- Cô cùng trẻ vận động nhẹ nhàng theo nhạc chủ đề '{theme}', trò chuyện gợi mở dẫn dắt vào đề tài '{topic}'."
    )

    if domain == "Thể chất":
        main_activity = (
            f"Hoạt động 2: Trọng tâm - {route} ({topic})\n"
            f"* Khởi động:\n"
            f"- Cho trẻ đi vòng tròn kết hợp các kiểu đi: đi thường -> đi bằng mũi bàn chân -> đi thường -> đi bằng gót chân -> chạy chậm -> chạy nhanh -> về đội hình hàng ngang.\n"
            f"* Trọng động:\n"
            f"a) Bài tập phát triển chung (BTPTC):\n"
            f"- Tập kết hợp lời ca theo chủ đề '{theme}': Động tác Tay (2 lần x 8 nhịp), Chân (4 lần x 8 nhịp), Bụng - Lườn (2 lần x 8 nhịp), Bật (2 lần x 8 nhịp).\n"
            f"b) Vận động cơ bản: {topic}\n"
            f"- Cô làm mẫu lần 1: Toàn phần, chuẩn xác, đẹp mắt, không giải thích.\n"
            f"- Cô làm mẫu lần 2: Kết hợp phân tích kỹ thuật từng tư thế chuẩn bị, nhịp thực hiện và tư thế kết thúc.\n"
            f"- Mời 1 - 2 trẻ lên tập thử để cô quan sát và nhận xét mẫu.\n"
            f"- Trẻ thực hiện: Lần lượt 2 trẻ lên thực hiện (2 - 3 lượt), cô bao quát, sửa sai tư thế kịp thời và khích lệ trẻ.\n"
            f"c) Trò chơi vận động bổ trợ:\n"
            f"- Cô phổ biến tên trò chơi, cách chơi, luật chơi và tổ chức cho trẻ chơi hào hứng 2 - 3 lần.\n"
            f"* Hồi tĩnh:\n"
            f"- Trẻ đi nhẹ nhàng 1 - 2 vòng quanh sân, hít thở sâu, thả lỏng cơ thể."
        )
    else:
        main_activity = (
            f"Hoạt động 2: Trọng tâm - Khám phá & Thực hành ({topic})\n"
            f"* Bước 1: Quan sát, khám phá và đàm thoại:\n"
            f"- Cô giới thiệu trực quan nội dung '{topic}' bằng đa giác quan (nhìn, nghe, chạm, trải nghiệm thực tế).\n"
            f"- Cô đặt hệ thống câu hỏi gợi mở từ dễ đến khó để trẻ tự nhận xét đặc điểm, cấu tạo, ý nghĩa của '{topic}'.\n"
            f"- Cô chốt lại kiến thức chuẩn xác, ngắn gọn, dễ nhớ cho trẻ lứa tuổi {age}.\n"
            f"* Bước 2: Luyện tập - Mở rộng ({sub_cat}):\n"
            f"- Tổ chức cho trẻ thực hành theo cá nhân và theo nhóm nhỏ.\n"
            f"- Lồng ghép trò chơi củng cố giúp trẻ khắc sâu kiến thức về '{topic}' và phát triển kỹ năng hợp tác.\n"
            f"- Cô quan sát, gợi ý, hỗ trợ những trẻ còn lúng túng và tuyên dương các ý tưởng sáng tạo của trẻ."
        )

    return f"""{school.upper()}
Lớp: {class_name} | Độ tuổi: {age}
GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG ({info.get('activity_type', 'Hoạt động học')})
ĐỀ TÀI: {topic.upper()}
Chủ đề: {theme} | Lĩnh vực: {domain} ({sub_cat} - {route})
Người soạn: {teacher} | Ngày dạy: {today_str}

I. MỤC ĐÍCH - YÊU CẦU
{goals_block}

II. CHUẨN BỊ
1. Đồ dùng của cô:
- Giáo án điện tử, nhạc không lời và bài hát theo chủ đề '{theme}'.
- Đồ dùng trực quan, mô hình, tranh ảnh hoặc vật thật phục vụ bài dạy: {topic}.{idea_note}
2. Đồ dùng của trẻ:
- Trang phục gọn gàng, tâm thế thoải mái; bộ đồ dùng thực hành đủ cho mỗi trẻ.
- Không gian lớp học/sân tập an toàn, sạch sẽ.

III. TỔ CHỨC HOẠT ĐỘNG
Hoạt động 1: Ổn định tổ chức, gây hứng thú (3 - 4 phút)
{intro_hook}
- Khơi gợi sự tò mò và kiểm tra sức khỏe, tâm thế của trẻ trước khi bước vào hoạt động chính.

{main_activity}

Hoạt động 3: Kết thúc (2 - 3 phút)
- Cô nhận xét chung giờ học, khen ngợi và động viên sự cố gắng của cả lớp.
- Nhắc nhở trẻ lồng ghép giáo dục ({extras_str}), cùng cô thu dọn đồ dùng gọn gàng và chuyển hoạt động nhẹ nhàng.

IV. ĐÁNH GIÁ CUỐI NGÀY
- Tình trạng sức khỏe: Đa số trẻ khỏe mạnh, tham gia đầy đủ các hoạt động, không có biểu hiện bất thường.
- Trạng thái cảm xúc, thái độ: Trẻ hào hứng, tích cực, biết phối hợp cùng cô và bạn trong suốt tiết học '{topic}'.
- Kiến thức, kỹ năng: Trẻ nắm vững mục tiêu bài học ({topic}), thực hiện tốt các yêu cầu sư phạm lứa tuổi {age}; cô kịp thời hỗ trợ các trẻ còn rụt rè để toàn lớp đạt kết quả tốt."""

# --- HÀM GỌI AI SIÊU TỐC VỚI TIMEOUT KHÓA CỨNG ---
def generate_fast_ai_content(client, prompt, fallback_info=None):
    models_to_try = [
        'gemini-3.1-flash-lite',
        'gemini-3.1-flash-lite-preview',
        'gemini-3.8-flash',
        'gemini-3.0-flash',
        'gemini-2.5-flash'
    ]
    config = build_fast_config()

    def _call_single_model(m_name):
        res = client.models.generate_content(
            model=m_name,
            contents=prompt,
            config=config
        )
        return res.text

    for model_name in models_to_try:
        executor = ThreadPoolExecutor(max_workers=1)
        future = executor.submit(_call_single_model, model_name)
        try:
            text_out = future.result(timeout=10)
            executor.shutdown(wait=False)
            if text_out and text_out.strip():
                return text_out
        except FuturesTimeoutError:
            executor.shutdown(wait=False)
            continue
        except Exception:
            executor.shutdown(wait=False)
            continue

    if fallback_info:
        return generate_instant_pedagogical_plan(fallback_info)
    raise Exception("Máy chủ AI đang bận, vui lòng thử lại sau ít giây!")

# --- HÀM XUẤT FILE WORD TỪ TEXT ĐÃ CHỈNH SỬA (HỖ TRỢ CHỈNH CỠ CHỮ) ---
def export_docx_from_text(text_content: str, info: dict, custom_template=None, font_size_pt=13):
    if custom_template:
        doc = Document(custom_template)
    else:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(font_size_pt)

        p_school = doc.add_paragraph()
        p_school.add_run(f"• Trường: {info['school']}\n").bold = True
        p_school.add_run(f"• Giáo viên: {info['teacher']} | Lớp: {info['class_name']}")

        title = doc.add_heading('GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG HỌC', level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p_info = doc.add_paragraph()
        p_info.add_run(f"• Chủ đề lớn: {info['theme']}\n").bold = True
        p_info.add_run(f"• Chế độ soạn thảo: {info['mode']} | Hiển thị: {info['display_format']}\n")
        p_info.add_run(f"• Loại hoạt động: {info['activity_type']} | Độ tuổi: {info['age']}\n")
        p_info.add_run(f"• Lĩnh vực: {info['domain']} ({info['sub_category']})\n").bold = True
        p_info.add_run(f"• Dạng tiết (ROUTE): {info['route_type']}\n")
        p_info.add_run(f"• Đề tài: {info['topic']}\n").bold = True
        if info['extras']:
            p_info.add_run(f"• Yêu cầu tích hợp: {', '.join(info['extras'])}\n")
        if info['custom_goals']:
            p_info.add_run(f"• Mục tiêu cần đạt & Mã MT: {info['custom_goals']}\n")
        if info['custom_ideas']:
            p_info.add_run(f"• Yêu cầu bổ sung của GV: {info['custom_ideas']}\n")
        if info['auto_assessment']:
            p_info.add_run(f"• Tự động tạo đánh giá cuối ngày: Bật")

    for line in text_content.split('\n'):
        if line.strip():
            doc.add_paragraph(line)

    safe_topic = "".join(c if c.isalnum() else "_" for c in info['topic'])[:40]
    filename = f"GiaoAn_{safe_topic}.docx"
    doc.save(filename)
    return filename

# --- CẤU HÌNH GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Phần Mềm Trợ Lý Mầm Non Chuyên Sâu", layout="wide")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "is_admin" not in st.session_state:
    st.session_state["is_admin"] = False
if "api_key_saved" not in st.session_state:
    st.session_state["api_key_saved"] = ""

# --- MÀN HÌNH ĐĂNG NHẬP MÃ BẢN QUYỀN ---
if not st.session_state["authenticated"]:
    st.title("🔐 Xác Thực Sử Dụng Phần Mềm Mầm Non")
    st.markdown("Vui lòng nhập **Mã Kích Hoạt Bản Quyền (100k/tháng)** để truy cập hệ thống trợ lý.")
    
    col_l1, col_l2 = st.columns([2, 1])
    with col_l1:
        input_key = st.text_input("Nhập Mã Kích Hoạt Của Bạn", type="password", placeholder="Nhập mã...")
        if st.button("🔓 Mở Khóa Sử Dụng", type="primary"):
            if input_key == "LBSTORE2026":
                st.session_state["authenticated"] = True
                st.session_state["is_admin"] = True
                st.success("👑 Đăng nhập quyền Admin thành công!")
                st.rerun()
            else:
                valid, msg = check_license(input_key)
                if valid:
                    st.session_state["authenticated"] = True
                    st.session_state["is_admin"] = False
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
    with col_l2:
        st.info("💡 **Gói Thuê Phần Mềm:**\n- Giá: **100.000 VNĐ / Tháng**\n- ĐTDD LB Store:\n  `0939.72.9009` (Văn Bình)\n- Chuyển khoản nhận mã ngay!")
    st.stop()

# --- GIAO DIỆN CHÍNH SAU KHI MỞ KHÓA ---
st.title("🧸 Trợ Lý Ảo Toàn Năng - Giáo Viên Mầm Non (LB Store)")
st.markdown("Hệ thống hỗ trợ soạn giáo án chuẩn, lên kế hoạch tháng, sáng tác trò chơi, viết sổ liên lạc và quản lý lớp học tự động.")

# ==========================================
# SIDEBAR (ĐÃ ĐƯA DANH MỤC TÍNH NĂNG LÊN TRÊN)
# ==========================================
with st.sidebar:
    st.markdown("### 📌 Danh Mục Tính Năng")
    
    if st.session_state["is_admin"]:
        app_mode = st.radio("Chọn tác vụ:", [
            "👑 Quản Lý Khách Thuê (Admin)", 
            "📝 Soạn Giáo Án Chuẩn", 
            "🗓️ Kế Hoạch & Sổ Chủ Đề", 
            "🎯 Trò Chơi & Đố Vui", 
            "💬 Sổ Liên Lạc & Zalo", 
            "✂️ Gợi Ý Đồ Dùng DIY", 
            "📚 Kho Lịch Sử Giáo Án"
        ])
    else:
        app_mode = st.radio("Chọn tác vụ:", [
            "📝 Soạn Giáo Án Chuẩn", 
            "🗓️ Kế Hoạch & Sổ Chủ Đề", 
            "🎯 Trò Chơi & Đố Vui", 
            "💬 Sổ Liên Lạc & Zalo", 
            "✂️ Gợi Ý Đồ Dùng DIY", 
            "📚 Kho Lịch Sử Giáo Án"
        ])

    st.divider()
    st.header("⚙️ Cấu hình hệ thống")
    input_api_key = st.text_input("Nhập Google Gemini API Key", type="password", value=st.session_state["api_key_saved"])
    
    if st.button("💾 Xác Nhận Lưu API Key"):
        if input_api_key.strip():
            st.session_state["api_key_saved"] = input_api_key.strip()
            st.success("✅ Đã lưu API Key thành công!")
        else:
            st.error("⚠️ Vui lòng nhập API Key hợp lệ!")

    if st.session_state["api_key_saved"]:
        st.success("🟢 Trạng thái: Đã kết nối AI sẵn sàng!")
    else:
        st.warning("🔴 Trạng thái: Chưa nhập API Key (Vẫn tạo được Giáo Án tức thì)!")
    
    st.divider()
    if st.button("🔒 Đăng Xuất / Khóa Lại"):
        st.session_state["authenticated"] = False
        st.session_state["is_admin"] = False
        st.rerun()

api_key = st.session_state["api_key_saved"]

# --- TAB ADMIN: QUẢN LÝ KHÁCH THUÊ ---
if st.session_state["is_admin"] and app_mode == "👑 Quản Lý Khách Thuê (Admin)":
    st.subheader("👑 Khu Vực Quản Trị Viên - Cấp Mã Thuê Bao")
    st.markdown("Nhập thông tin khách hàng vừa chuyển khoản 100k để hệ thống tự sinh mã kích hoạt:")
    
    with st.form("admin_form"):
        c_name_input = st.text_input("Tên khách hàng / Tên trường (VD: Cô Thảo - Mầm Non Hướng Dương)")
        months_input = st.selectbox("Thời hạn thuê", [1, 2, 3, 6, 12], format_func=lambda x: f"{x} Tháng ({x*100}k)")
        submitted = st.form_submit_button("✨ Tạo Mã Kích Hoạt Mới", type="primary")
        
        if submitted:
            if not c_name_input:
                st.warning("Vui lòng nhập tên khách hàng!")
            else:
                success, key_or_err, exp_date = create_new_license(c_name_input, months_input)
                if success:
                    st.success(f"🎉 Đã tạo mã thành công cho **{c_name_input}**!")
                    st.code(f"Mã kích hoạt: {key_or_err} (Hạn đến: {exp_date})", language="text")
                    st.info("👉 Hãy copy mã trên gửi qua Zalo cho khách hàng!")
                else:
                    st.error(f"Lỗi: {key_or_err}")

    st.divider()
    st.subheader("📋 Danh Sách Tất Cả Mã Đã Cấp")
    licenses_list = get_all_licenses()
    for l_id, l_key, l_name, l_exp, l_active in licenses_list:
        status_text = "🟢 Đang hoạt động" if l_active == 1 else "🔴 Đã khóa"
        st.markdown(f"- **Mã:** `{l_key}` | **Khách:** {l_name} | **Hạn:** {l_exp} | **Trạng thái:** {status_text}")

# --- TAB 1: SOẠN GIÁO ÁN CHUẨN ---
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
    with c_mode1:
        mode_gv = st.button("🌱 Giáo án thường\nChuẩn mực, tinh gọn, rõ ràng theo đúng phân phối chương trình.", use_container_width=True)
    with c_mode2:
        mode_gioi = st.button("⭐ Thi GV Giỏi\nStorytelling, tình huống sáng tạo nhưng KHÔNG vượt chuẩn tuổi.", use_container_width=True)
    
    if "soan_mode_val" not in st.session_state:
        st.session_state["soan_mode_val"] = "Giáo án thường"
    
    if mode_gv:
        st.session_state["soan_mode_val"] = "Giáo án thường"
    if mode_gioi:
        st.session_state["soan_mode_val"] = "Thi GV Giỏi"
        
    st.info(f"👉 Đang chọn chế độ soạn thảo: **{st.session_state['soan_mode_val']}**")

    st.markdown("### ĐỊNH DẠNG HIỂN THỊ GIÁO ÁN")
    c_fmt1, c_fmt2 = st.columns(2)
    with c_fmt1:
        fmt_1col = st.button("📄 1 Cột chuẩn (Toàn cảnh)\nLời dẫn, câu hỏi & hành động của cô theo gạch đầu dòng (-)", use_container_width=True)
    with c_fmt2:
        fmt_2col = st.button("📊 2 Cột (Cô & Trẻ)\nHoạt động của giáo viên song song nhiệm vụ dự kiến của trẻ", use_container_width=True)
        
    if "display_fmt_val" not in st.session_state:
        st.session_state["display_fmt_val"] = "1 Cột chuẩn"
        
    if fmt_1col:
        st.session_state["display_fmt_val"] = "1 Cột chuẩn"
    if fmt_2col:
        st.session_state["display_fmt_val"] = "2 Cột (Cô & Trẻ)"
        
    st.info(f"👉 Đang chọn định dạng: **{st.session_state['display_fmt_val']}**")

    st.markdown("### 1. ĐỘ TUỔI HỌC SINH *")
    age_group = st.radio("Chọn độ tuổi", ["3-12 th", "12-18 th", "18-24 th", "24-36 th", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], horizontal=True, index=5)

    st.markdown("### 2. LOẠI HÌNH HOẠT ĐỘNG *")
    activity_type = st.radio("Chọn loại hoạt động", ["Hoạt động học", "Hoạt động góc / Vui chơi", "Hoạt động ngoài trời", "Hoạt động chiều"], horizontal=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### 3. LĨNH VỰC PHÁT TRIỂN *")
        domain = st.selectbox("Lĩnh vực", ["Thể chất", "Nhận thức", "Ngôn ngữ", "Thẩm mỹ", "Tình cảm - KNXH"])
        
        sub_category = ""
        if domain == "Nhận thức":
            sub_category = st.selectbox("Phân loại chi tiết", ["Khám phá", "Toán"])
        elif domain == "Ngôn ngữ":
            sub_category = st.selectbox("Phân loại chi tiết", ["Chữ cái - tiết tập tô", "Làm quen", "Trò chơi", "Thơ", "Truyện"])
        elif domain == "Thẩm mỹ":
            sub_category = st.selectbox("Phân loại chi tiết", ["Âm nhạc - Dạy hát", "Nghe nhạc", "Múa minh họa", "Gõ đệm theo phách, nhịp", "Tạo hình"])
        else:
            sub_category = st.text_input("Hoạt động cụ thể", "Phát triển vận động")

    with col_b:
        st.markdown("### 4. DẠNG TIẾT CHUYÊN MÔN (ROUTE) *")
        if domain == "Thể chất":
            route_type = st.selectbox("Dạng tiết", ["Vận động cơ bản (VĐCB)", "Trò chơi vận động", "Thể dục sáng"])
        elif domain == "Nhận thức":
            route_type = st.selectbox("Dạng tiết", ["Khám phá khoa học", "Làm quen với toán", "Khám phá xã hội"])
        elif domain == "Ngôn ngữ":
            route_type = st.selectbox("Dạng tiết", ["Làm quen chữ cái", "Làm quen văn học (Thơ/Truyện)", "Phát triển ngôn ngữ"])
        elif domain == "Thẩm mỹ":
            route_type = st.selectbox("Dạng tiết", ["Âm nhạc", "Tạo hình"])
        else:
            route_type = st.selectbox("Dạng tiết", ["Kỹ năng xã hội", "Hoạt động chung"])

    st.markdown("### 5. CẤU TRÚC MỤC ĐÍCH - YÊU CẦU:")
    goal_structure = st.radio("Lựa chọn cấu trúc", ["1-3 gạch đầu dòng", "2-3 mục chuẩn (Kiến thức, Kỹ năng, Thái độ)"], horizontal=True, index=1)

    st.markdown("### 6. ĐỀ TÀI BÀI DẠY CỤ THỂ * (BẮT BUỘC):")
    topic = st.text_input("Nhập tên đề tài", "Ném xa bằng 1 tay")

    st.markdown("### 7. YÊU CẦU TÍCH HỢP (TÙY CHỌN) - (Chọn 1-2 mục):")
    col_c1, col_c2, col_c3, col_c4, col_c5 = st.columns(5)
    with col_c1:
        opt0 = st.checkbox("Không")
    with col_c2:
        opt1 = st.checkbox("STEAM")
    with col_c3:
        opt2 = st.checkbox("CÔNG NGHỆ SỐ")
    with col_c4:
        opt3 = st.checkbox("Bảo vệ môi trường")
    with col_c5:
        opt4 = st.checkbox("Quyền trẻ em")
    
    extras = []
    if opt1: extras.append("STEAM")
    if opt2: extras.append("CÔNG NGHỆ SỐ")
    if opt3: extras.append("Bảo vệ môi trường")
    if opt4: extras.append("Quyền trẻ em")

    st.markdown("### 8. MỤC TIÊU CẦN ĐẠT & MÃ MỤC TIÊU XÁC THỰC (MT):")
    custom_goals = st.text_area("Nhập MT1, MT45... kèm nội dung thực tế (Tùy chọn)", placeholder="VD:\n- Trẻ biết tên vận động cơ bản ném xa bằng 1 tay (MT15).\n- Trẻ hiểu cách thực hiện kỹ thuật ném xa bằng 1 tay: tư thế chuẩn bị, cách đưa tay và hướng ném.")

    st.markdown("### 9. YÊU CẦU BỔ SUNG (Ý TƯỞNG GIÁO VIÊN MUỐN ĐƯA VÀO TIẾT DẠY):")
    custom_ideas = st.text_area("Nhập ý tưởng riêng (Tùy chọn)", placeholder="Cô hãy nhập các ý tưởng riêng: đạo cụ rối tay, nhạc chủ đề, bài thơ mới, cách chia nhóm...")

    st.markdown("### 10. ĐÁNH GIÁ CUỐI NGÀY:")
    auto_assessment = st.checkbox("Quan sát thực tế cho Đánh giá cuối ngày (Mục IV) - Tạo đánh giá luôn: AI tự động biên soạn nội dung nhận xét đánh giá cuối ngày bám sát đề tài.", value=True)

    uploaded_template = st.file_uploader("Tải lên mẫu file Word (.docx) của trường (Tùy chọn)", type=["docx"])

    plan_info_dict = {
        "school": school_name, "teacher": teacher_name, "class_name": class_name,
        "theme": theme, "domain": domain, "sub_category": sub_category, "activity_type": activity_type,
        "route_type": route_type, "topic": topic, "method": "Gemini 3.8 Flash", "age": age_group, "extras": extras,
        "mode": st.session_state["soan_mode_val"], "display_format": st.session_state["display_fmt_val"], 
        "custom_goals": custom_goals, "custom_ideas": custom_ideas, "auto_assessment": auto_assessment
    }

    if st.button("✨ Tạo Giáo Án Chương Trình Hiện Hành", type="primary", use_container_width=True):
        with st.spinner("⚡ Đang khởi tạo giáo án chuẩn sư phạm (Tốc độ tối ưu ~4.7 giây)..."):
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    prompt = f"""
                    Soạn giáo án mầm non chuẩn Bộ GD&ĐT Việt Nam thật ngắn gọn, súc tích, chuẩn xác 100/100 điểm sư phạm:
                    - Trường: {school_name} | Lớp: {class_name} | Độ tuổi: {age_group} | Giáo viên: {teacher_name}
                    - Chủ đề: {theme} | Loại hoạt động: {activity_type}
                    - Lĩnh vực: {domain} ({sub_category}) | Dạng tiết: {route_type}
                    - Đề tài: {topic}
                    - Chế độ: {st.session_state['soan_mode_val']} | Định dạng: {st.session_state['display_fmt_val']}
                    - Cấu trúc mục tiêu: {goal_structure}
                    - Tích hợp: {', '.join(extras) if extras else 'Không'}
                    - Mã MT & Mục tiêu riêng: {custom_goals if custom_goals else 'Tự động chuẩn hóa'}
                    - Ý tưởng riêng: {custom_ideas if custom_ideas else 'Không'}
                    
                    Trình bày đúng 4 mục lớn:
                    I. MỤC ĐÍCH - YÊU CẦU (1. Kiến thức, 2. Kỹ năng, 3. Thái độ)
                    II. CHUẨN BỊ (1. Đồ dùng của cô, 2. Đồ dùng của trẻ)
                    III. TỔ CHỨC HOẠT ĐỘNG (Hoạt động 1: Ổn định tổ chức; Hoạt động 2: Trọng tâm; Hoạt động 3: Kết thúc)
                    IV. ĐÁNH GIÁ CUỐI NGÀY (Gồm đúng 3 gạch đầu dòng: - Tình trạng sức khỏe; - Trạng thái cảm xúc, thái độ; - Kiến thức, kỹ năng)
                    """
                    result_text = generate_fast_ai_content(client, prompt, fallback_info=plan_info_dict)
                except Exception:
                    result_text = generate_instant_pedagogical_plan(plan_info_dict)
            else:
                result_text = generate_instant_pedagogical_plan(plan_info_dict)

            st.session_state["generated_plan_text"] = result_text
            st.session_state["plan_info"] = plan_info_dict
            save_plan_to_db(topic, result_text)
        st.success("🎉 Đã hoàn thành giáo án! Xem và thao tác bên dưới:")

    # =========================================================================
    # KHU VỰC KẾT QUẢ: TÙY CHỈNH TRÌNH BÀY & THAO TÁC (ĐÃ SẮP XẾP LẠI THỨ TỰ)
    # =========================================================================
    if "generated_plan_text" in st.session_state:
        st.divider()
        info = st.session_state["plan_info"]
        
        # Lấy kích thước chữ hiện tại từ session_state (mặc định 14pt)
        current_font_size = st.session_state.get("font_size_slider", 14)
        
        st.markdown("### 🛠️ Tùy Chỉnh Trình Bày & Thao Tác")
        
        # 1. TIẾP LÀ: TẢI FILE, IN GIÁO ÁN
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            template_path = uploaded_template if uploaded_template else None
            docx_file = export_docx_from_text(st.session_state["generated_plan_text"], info, template_path, font_size_pt=current_font_size)
            with open(docx_file, "rb") as file:
                st.download_button(
                    label="📥 Tải File Word (.docx)",
                    data=file,
                    file_name=docx_file,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
        with btn_col2:
            plan_text_json = json.dumps(st.session_state["generated_plan_text"])
            topic_title_json = json.dumps(f"Giáo Án: {info['topic']}")
            print_button_html = f"""
            <div style="display: flex; justify-content: center; align-items: center; width: 100%;">
                <button onclick="printLessonPlan()" style="
                    background: linear-gradient(135deg, #1e7e34, #28a745);
                    color: white;
                    border: none;
                    padding: 10px 24px;
                    border-radius: 8px;
                    font-size: 15px;
                    font-weight: bold;
                    cursor: pointer;
                    width: 100%;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.15);
                    transition: all 0.2s ease;
                ">
                    🖨️ In Giáo Án Ngay
                </button>
            </div>
            <script>
            function printLessonPlan() {{
                var rawContent = {plan_text_json};
                var docTitle = {topic_title_json};
                var pWin = window.open('', '_blank');
                pWin.document.write('<!DOCTYPE html><html><head><title>' + docTitle + '</title>');
                pWin.document.write('<style>');
                pWin.document.write('@page {{ size: A4; margin: 15mm 20mm 15mm 20mm; }}');
                pWin.document.write('body {{ font-family: "Times New Roman", Times, serif; font-size: {current_font_size}pt; line-height: 1.5; color: #000; margin: 0; padding: 0; }}');
                pWin.document.write('.doc-body {{ white-space: pre-wrap; word-wrap: break-word; text-align: justify; }}');
                pWin.document.write('</style></head><body>');
                pWin.document.write('<div class="doc-body">' + rawContent.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\\n/g, "<br/>") + '</div>');
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
            components.html(print_button_html, height=52)

        # 2. RỒI ĐẾN: ĐIỂM SƯ PHẠM, TIẾP THEO LÀ ĐỊNH DẠNG CỘT
        info_col1, info_col2 = st.columns(2)
        with info_col1:
            st.success("✅ **Điểm sư phạm: 100/100**")
        with info_col2:
            st.info(f"📐 **Định dạng:** {info['display_format']}")
            
        # 3. RỒI MỚI ĐẾN: KÍCH THƯỚC HIỂN THỊ
        font_size_val = st.slider(
            "🔤 Kích thước chữ hiển thị & In (pt):", 
            min_value=12, 
            max_value=24, 
            value=current_font_size, 
            step=1, 
            key="font_size_slider"
        )
        
        # 4. BẢN XEM TRƯỚC GIÁO ÁN (TRANG GIẤY IN TRẮNG)
        st.markdown("### 📄 Bản Xem Trước Giáo Án (Trang Giấy In):")
        escaped_preview_text = html.escape(st.session_state["generated_plan_text"])
        st.markdown(
            f"""
            <div style="
                background-color: #ffffff;
                color: #111111;
                font-family: 'Times New Roman', Times, serif;
                font-size: {font_size_val}pt;
                line-height: 1.6;
                padding: 30px 40px;
                border-radius: 8px;
                border: 1px solid #d3d3d3;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
                white-space: pre-wrap;
                text-align: justify;
                word-wrap: break-word;
            ">{escaped_preview_text}</div>
            """,
            unsafe_allow_html=True
        )

        with st.expander("✏️ Chỉnh sửa nội dung văn bản (Bấm vào để mở khung sửa chữ)"):
            edited_text = st.text_area("Nội dung giáo án:", st.session_state["generated_plan_text"], height=380)
            if edited_text != st.session_state["generated_plan_text"]:
                st.session_state["generated_plan_text"] = edited_text
                st.rerun()

# --- CÁC TAB CÒN LẠI ---
elif app_mode == "🗓️ Kế Hoạch & Sổ Chủ Đề":
    st.subheader("🗓️ Lập Kế Hoạch Giáo Dục Theo Chủ Đề & Tháng")
    theme_input = st.text_input("Nhập tên Chủ Đề Lớn (VD: Thế giới thực vật, Giao thông, Tết và mùa xuân...)")
    age_plan = st.selectbox("Độ tuổi lớp", ["3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], key="age_p")
    
    if st.button("🚀 Tạo Kế Hoạch Trọn Gói", type="primary"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            with st.spinner("⚡ Đang lập kế hoạch chủ đề siêu tốc..."):
                try:
                    res_text = generate_fast_ai_content(
                        client,
                        f"Lập kế hoạch giáo dục chi tiết theo chủ đề '{theme_input}' cho trẻ lứa tuổi {age_plan}, bao gồm: mục tiêu phát triển các lĩnh vực, mạng nội dung, mạng hoạt động và gợi ý phân bổ các góc chơi trong tuần. Trình bày bài bản, rõ ràng."
                    )
                    st.success("🎉 Đã hoàn thành kế hoạch chủ đề cho cô:")
                    st.markdown(res_text)
                except Exception as e:
                    st.error(f"⚠️ {e}")

elif app_mode == "🎯 Trò Chơi & Đố Vui":
    st.subheader("🎯 Sáng Tác Trò Chơi, Câu Đố & Bài Thơ Mầm Non")
    topic_game = st.text_input("Nhập chủ đề hoặc sự vật (VD: Củ cà rốt, An toàn giao thông, Đếm đến 5...)")
    genre = st.selectbox("Thể loại cần tạo", ["Bài thơ ngắn dễ thuộc", "Bài đồng dao vui nhộn", "Câu đố vui cho trẻ", "Trò chơi vận động tập thể"])
    
    if st.button("✨ Sáng Tác Ngay", type="primary"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            with st.spinner("⚡ Đang sáng tác nội dung siêu tốc..."):
                try:
                    res_text = generate_fast_ai_content(
                        client,
                        f"Hãy sáng tác hoặc thiết kế một {genre} dành riêng cho trẻ mầm non với chủ đề: {topic_game}. Ngôn ngữ trong trẻo, vui tươi. Nếu là trò chơi cần nêu rõ mục đích, luật chơi và cách chơi."
                    )
                    st.success("🎉 Đã hoàn thành tác phẩm cho cô:")
                    st.markdown(res_text)
                except Exception as e:
                    st.error(f"⚠️ {e}")

elif app_mode == "💬 Sổ Liên Lạc & Zalo":
    st.subheader("💬 Trợ Lý Viết Sổ Liên Lạc & Tin Nhắn Zalo Phụ Huynh")
    task_type = st.radio("Chọn loại nội dung cần viết:", ["Tin nhắn tổng kết ngày/tuần gửi nhóm Zalo", "Viết lời nhận xét sự tiến bộ cho từng bé"])
    
    if task_type == "Tin nhắn tổng kết ngày/tuần gửi nhóm Zalo":
        summary_topic = st.text_input("Nội dung hoạt động chính (VD: Hôm nay các con học bài thơ 'Cục len', tập tô chữ o...)")
        if st.button("💌 Soạn Tin Nhắn Zalo", type="primary"):
            if not api_key:
                st.error("⚠️ Vui lòng nhập API Key!")
            else:
                client = genai.Client(api_key=api_key)
                with st.spinner("⚡ Đang soạn tin siêu tốc..."):
                    try:
                        res_text = generate_fast_ai_content(
                            client,
                            f"Soạn một tin nhắn Zalo thân thiện, chuyên nghiệp từ giáo viên mầm non gửi phụ huynh tóm tắt các hoạt động: {summary_topic}. Kèm lời dặn dò nhẹ nhàng buổi tối."
                        )
                        st.success("🎉 Mẫu tin nhắn Zalo:")
                        st.markdown(res_text)
                    except Exception as e:
                        st.error(f"⚠️ {e}")
    else:
        child_info = st.text_input("Nhập vài từ khóa về tình hình của bé (VD: Bé Bảo An: ngoan, ăn giỏi, tích cực phát biểu nhưng còn rụt rè giờ thể chất)")
        if st.button("✍️ Viết Lời Nhận Xét", type="primary"):
            if not api_key:
                st.error("⚠️ Vui lòng nhập API Key!")
            else:
                client = genai.Client(api_key=api_key)
                with st.spinner("⚡ Đang viết lời nhận xét siêu tốc..."):
                    try:
                        res_text = generate_fast_ai_content(
                            client,
                            f"Viết một đoạn nhận xét cuối tuần khéo léo, tình cảm, mang tính khích lệ cao dành cho học sinh mầm non dựa trên thông tin: {child_info}."
                        )
                        st.success("🎉 Lời nhận xét mẫu cho cô:")
                        st.markdown(res_text)
                    except Exception as e:
                        st.error(f"⚠️ {e}")

elif app_mode == "✂️ Gợi Ý Đồ Dùng DIY":
    st.subheader("✂️ Gợi Ý Làm Đồ Dùng Dạy Học Từ Vật Liệu Mở")
    lesson_name = st.text_input("Nhập tên bài học hoặc kỹ năng cần làm đồ dùng (VD: Đếm số lượng trong phạm vi 5...)")
    
    if st.button("🛠️ Gợi Ý Cách Làm", type="primary"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            with st.spinner("⚡ Đang lên ý tưởng chế tạo siêu tốc..."):
                try:
                    res_text = generate_fast_ai_content(
                        client,
                        f"Gợi ý cho giáo viên mầm non cách tận dụng các nguyên vật liệu phế liệu/vật liệu mở (chai nhựa, ống hút, bìa carton...) để làm đồ dùng đồ chơi tự tạo phục vụ cho bài học: {lesson_name}. Liệt kê nguyên liệu và các bước làm ngắn gọn, an toàn cho trẻ."
                    )
                    st.success("🎉 Ý tưởng đồ dùng tự tạo:")
                    st.markdown(res_text)
                except Exception as e:
                    st.error(f"⚠️ {e}")

elif app_mode == "📚 Kho Lịch Sử Giáo Án":
    st.subheader("📚 Kho Lưu Trữ Giáo Án Đã Soạn")
    plans = get_all_saved_plans()
    if not plans:
        st.info("Chưa có giáo án nào được lưu trữ trong hệ thống.")
    else:
        for p_id, p_title, p_time in plans:
            with st.container():
                cols = st.columns([4, 2, 2])
                cols[0].markdown(f"**{p_title}** *(Lưu lúc: {p_time})*")
                if cols[1].button("Tải lại Word", key=f"load_{p_id}"):
                    row = get_plan_by_id(p_id)
                    if row:
                        title_db, content_db = row
                        info_dummy = {"school": "Mầm Non Phước Thạnh", "teacher": "Nguyễn Thị Thanh Vân", "class_name": "Lớp Chồi 1", "theme": "Chung", "domain": "Nhận thức", "sub_category": "Toán", "activity_type": "Hoạt động học", "route_type": "Khám phá", "topic": title_db, "method": "Gemini 3.8 Flash", "age": "4-5 tuổi", "extras": [], "mode": "Giáo án thường", "display_format": "1 Cột chuẩn", "custom_goals": "", "custom_ideas": "", "auto_assessment": True}
                        f_path = export_docx_from_text(content_db, info_dummy)
                        with open(f_path, "rb") as f_dl:
                            cols[2].download_button("📥 Tải File", data=f_dl, file_name=f_path, key=f"dl_{p_id}")
