import sqlite3
import json
import streamlit as st
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

# --- HÀM KHỞI TẠO DATABASE & QUẢN LÝ MÃ KÍCH HOẠT ---
def init_db():
    conn = sqlite3.connect('mamnon_rag.db')
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

init_db()

def check_license(key):
    conn = sqlite3.connect('mamnon_rag.db')
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
    return False, "❌ Mã kích hoạt không hợp lệ hoặc không tồn tại!"

def create_new_license(customer_name, months):
    random_str = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
    new_key = f"LB-{random_str}"
    expires_date = (datetime.now() + timedelta(days=30 * months)).strftime("%Y-%m-%d")
    
    conn = sqlite3.connect('mamnon_rag.db')
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO licenses (license_key, customer_name, expires_at, is_active) VALUES (?, ?, ?, 1)", 
                       (new_key, customer_name, expires_date))
        conn.commit()
        conn.close()
        return True, new_key, expires_date
    except Exception as e:
        conn.close()
        return False, str(e), ""

def get_all_licenses():
    conn = sqlite3.connect('mamnon_rag.db')
    cursor = conn.cursor()
    cursor.execute("SELECT id, license_key, customer_name, expires_at, is_active FROM licenses ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def save_plan_to_db(title, content_json):
    conn = sqlite3.connect('mamnon_rag.db')
    cursor = conn.cursor()
    cursor.execute("INSERT INTO saved_plans (title, content) VALUES (?, ?)", (title, content_json))
    conn.commit()
    conn.close()

def get_all_saved_plans():
    conn = sqlite3.connect('mamnon_rag.db')
    cursor = conn.cursor()
    cursor.execute("SELECT id, title, created_at FROM saved_plans ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_plan_by_id(plan_id):
    conn = sqlite3.connect('mamnon_rag.db')
    cursor = conn.cursor()
    cursor.execute("SELECT title, content FROM saved_plans WHERE id = ?", (plan_id,))
    row = cursor.fetchone()
    conn.close()
    return row

# --- HÀM XUẤT FILE WORD TỪ TEXT ĐÃ CHỈNH SỬA ---
def export_docx_from_text(text_content: str, info: dict, custom_template=None):
    if custom_template:
        doc = Document(custom_template)
    else:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(13)

        p_school = doc.add_paragraph()
        p_school.add_run(f"• Trường: {info['school']}\n").bold = True
        p_school.add_run(f"• Giáo viên: {info['teacher']} | Lớp: {info['class_name']}")

        title = doc.add_heading('GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG HỌC', level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p_info = doc.add_paragraph()
        p_info.add_run(f"• Chủ đề lớn: {info['theme']}\n").bold = True
        p_info.add_run(f"• Loại hoạt động: {info['activity_type']} | Độ tuổi: {info['age']}\n")
        p_info.add_run(f"• Lĩnh vực: {info['domain']} ({info['sub_category']})\n").bold = True
        p_info.add_run(f"• Dạng tiết (ROUTE): {info['route_type']}\n")
        p_info.add_run(f"• Đề tài: {info['topic']} (Phương pháp: {info['method']})\n").bold = True
        if info['extras']:
            p_info.add_run(f"• Tích hợp: {', '.join(info['extras'])}")

    for line in text_content.split('\n'):
        if line.strip():
            doc.add_paragraph(line)

    filename = f"GiaoAn_{info['topic'].replace(' ', '_')}.docx"
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

with st.sidebar:
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
        st.warning("🔴 Trạng thái: Chưa nhập API Key!")
    
    st.divider()
    if st.button("🔒 Đăng Xuất / Khóa Lại"):
        st.session_state["authenticated"] = False
        st.session_state["is_admin"] = False
        st.rerun()
        
    st.divider()
    st.markdown("### 📌 Danh Mục Tính Năng")
    
    if st.session_state["is_admin"]:
        app_mode = st.radio("Chọn tác vụ:", [
            "👑 Quản Lý Khách Thuê (Admin)", 
            "📝 Soạn Giáo Án Chuẩn", 
            "🗓️ Kế Hoạch & Sổ Chủ Đề", 
            "🎯 Trò Chơi & Đố Vui", 
            "💌 Sổ Liên Lạc & Zalo", 
            "✂️ Gợi Ý Đồ Dùng DIY", 
            "📚 Kho Lịch Sử Giáo Án"
        ])
    else:
        app_mode = st.radio("Chọn tác vụ:", [
            "📝 Soạn Giáo Án Chuẩn", 
            "🗓️ Kế Hoạch & Sổ Chủ Đề", 
            "🎯 Trò Chơi & Đố Vui", 
            "💌 Sổ Liên Lạc & Zalo", 
            "✂️ Gợi Ý Đồ Dùng DIY", 
            "📚 Kho Lịch Sử Giáo Án"
        ])

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

    st.divider("📋 Danh Sách Tất Cả Mã Đã Cấp")
    licenses_list = get_all_licenses()
    for l_id, l_key, l_name, l_exp, l_active in licenses_list:
        status_text = "🟢 Đang hoạt động" if l_active == 1 else "🔴 Đã khóa"
        st.markdown(f"- **Mã:** `{l_key}` | **Khách:** {l_name} | **Hạn:** {l_exp} | **Trạng thái:** {status_text}")

# --- TAB 1: SOẠN GIÁO ÁN CHUẨN ---
elif app_mode == "📝 Soạn Giáo Án Chuẩn":
    st.subheader("📝 Thiết Lập Hệ Thống Soạn Giáo Án")
    
    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("Họ và tên giáo viên", "Nguyễn Thị Thanh Vân")
        school_name = st.text_input("Tên trường mầm non", "Mầm Non Phước Thạnh")
    with col2:
        class_name = st.text_input("Tên lớp", "Lớp Chồi 1 (4-5 tuổi)")
        theme = st.text_input("Chủ đề lớn / Nhánh", "Thế giới thực vật")

    st.markdown("### 1. Độ tuổi học sinh:")
    age_group = st.radio("Chọn độ tuổi", ["3-12 th", "12-18 th", "18-24 th", "24-36 th", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], horizontal=True, index=5)

    st.markdown("### 2. Loại hình hoạt động:")
    activity_type = st.radio("Chọn loại hoạt động", ["Hoạt động học", "Hoạt động góc / Vui chơi", "Hoạt động ngoài trời", "Hoạt động chiều"], horizontal=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### 3. Lĩnh vực phát triển & Phân loại:")
        domain = st.selectbox("Lĩnh vực", ["Thể chất", "Nhận thức", "Ngôn ngữ", "Thẩm mỹ", "Tình cảm - KNXH"])
        
        sub_category = ""
        if domain == "Nhận thức":
            sub_category = st.selectbox("Phân loại chi tiết", ["Toán", "Khám phá"])
        elif domain == "Ngôn ngữ":
            sub_category = st.selectbox("Phân loại chi tiết", ["Chữ cái - tiết tập tô", "Làm quen", "Trò chơi", "Thơ", "Truyện"])
        elif domain == "Thẩm mỹ":
            sub_category = st.selectbox("Phân loại chi tiết", ["Âm nhạc - Dạy hát", "Nghe nhạc", "Múa minh họa", "Gõ đệm theo phách, nhịp", "Tạo hình"])
        else:
            sub_category = st.text_input("Hoạt động cụ thể", "Hoạt động chung")

    with col_b:
        st.markdown("### 4. Dạng tiết chuyên môn (ROUTE):")
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

    st.markdown("### 5. Cấu trúc mục tiêu - yêu cầu:")
    goal_structure = st.radio("Lựa chọn cấu trúc", ["1-3 gạch đầu dòng", "2-3 mục chuẩn"], horizontal=True)

    st.markdown("### 6. Đề tài bài dạy cụ thể:")
    topic = st.text_input("Nhập tên đề tài", "Học chữ cái o, ô, ơ")
    method = st.selectbox("Phương pháp giảng dạy", ["Truyền thống", "STEAM 5E", "STEAM EDP"])

    st.markdown("### 7. Tích hợp giáo dục STEAM & Chuyên đề cốt lõi:")
    col_c1, col_c2, col_c3, col_c4 = st.columns(4)
    with col_c1:
        opt1 = st.checkbox("Bảo vệ môi trường")
    with col_c2:
        opt2 = st.checkbox("An toàn giao thông")
    with col_c3:
        opt3 = st.checkbox("Quyền trẻ em")
    with col_c4:
        opt4 = st.checkbox("Công nghệ số")
    
    extras = []
    if opt1: extras.append("Bảo vệ môi trường")
    if opt2: extras.append("An toàn giao thông")
    if opt3: extras.append("Quyền trẻ em")
    if opt4: extras.append("Công nghệ số")

    uploaded_template = st.file_uploader("Tải lên mẫu file Word (.docx) của trường (Tùy chọn)", type=["docx"])

    if st.button("🚀 Tạo Bản Án Chuẩn Chương Trình Hiện Hành", type="primary"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập và bấm nút 'Xác Nhận Lưu API Key' ở menu bên trái trước khi soạn giáo án!")
        else:
            try:
                client = genai.Client(api_key=api_key)
                prompt = f"""
                Bạn là chuyên gia giáo dục mầm non chuẩn Bộ GD&ĐT Việt Nam. Hãy soạn một giáo án chi tiết chuẩn xác bằng tiếng Việt với các thông tin sau:
                - Tên bài dạy: {topic}
                - Loại hình hoạt động: {activity_type}
                - Lĩnh vực phát triển: {domain} (Phân loại chi tiết: {sub_category})
                - Dạng tiết (ROUTE): {route_type}
                - Chủ đề lớn / Nhánh: {theme}
                - Độ tuổi: {age_group}
                - Phương pháp: {method}
                - Cấu trúc mục tiêu: {goal_structure}
                - Tích hợp thêm: {', '.join(extras) if extras else 'Không'}
                
                Yêu cầu trình bày rõ ràng các phần chuẩn sư phạm mầm non: Mục đích yêu cầu, Chuẩn bị (đồ dùng của cô và trẻ), Cách tiến hành chi tiết từng bước (Hoạt động của cô và Hoạt động của trẻ), Đánh giá sau giờ dạy.
                """
                with st.spinner("AI đang phân tích và biên soạn giáo án chuẩn mầm non..."):
                    response = client.models.generate_content(
                        model='gemini-3.8-flash',
                        contents=prompt,
                    )
                    
                    st.session_state["generated_plan_text"] = response.text
                    st.session_state["plan_info"] = {
                        "school": school_name, "teacher": teacher_name, "class_name": class_name,
                        "theme": theme, "domain": domain, "sub_category": sub_category, "activity_type": activity_type,
                        "route_type": route_type, "topic": topic, "method": method, "age": age_group, "extras": extras
                    }
                    
                    save_plan_to_db(topic, response.text)
                st.success("🎉 Đã soạn giáo án thành công! Cô có thể xem và chỉnh sửa trực tiếp bên dưới.")
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi gọi AI: {e}")

    if "generated_plan_text" in st.session_state:
        st.divider()
        st.markdown("### ✏️ Chỉnh sửa nội dung giáo án trực tiếp:")
        st.info("💡 Cô có thể sửa chữ trực tiếp trong ô dưới đây theo ý muốn trước khi tải về file Word!")
        
        edited_text = st.text_area("Nội dung giáo án:", st.session_state["generated_plan_text"], height=400)
        
        info = st.session_state["plan_info"]
        template_path = uploaded_template if uploaded_template else None
        docx_file = export_docx_from_text(edited_text, info, template_path)

        with open(docx_file, "rb") as file:
            st.download_button(
                label="📥 Tải xuống File Word (.docx) đã chỉnh sửa",
                data=file,
                file_name=docx_file,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )

# --- TAB 2: KẾ HOẠCH & SỔ CHỦ ĐỀ ---
elif app_mode == "🗓️ Kế Hoạch & Sổ Chủ Đề":
    st.subheader("🗓️ Lập Kế Hoạch Giáo Dục Theo Chủ Đề & Tháng")
    theme_input = st.text_input("Nhập tên Chủ Đề Lớn (VD: Thế giới thực vật, Giao thông, Tết và mùa xuân...)")
    age_plan = st.selectbox("Độ tuổi lớp", ["3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], key="age_p")
    
    if st.button("🚀 Tạo Kế Hoạch Trọn Gói", type="primary"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            prompt = f"Lập kế hoạch giáo dục chi tiết theo chủ đề '{theme_input}' cho trẻ lứa tuổi {age_plan}, bao gồm: mục tiêu phát triển các lĩnh vực, mạng nội dung, mạng hoạt động và gợi ý phân bổ các góc chơi trong tuần. Trình bày bài bản, rõ ràng."
            with st.spinner("Đang lập kế hoạch chủ đề..."):
                res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                st.success("Đã hoàn thành kế hoạch chủ đề cho cô:")
                st.markdown(res.text, unsafe_allow_html=True)

# --- TAB 3: TRÒ CHƠI & ĐỐ VUI ---
elif app_mode == "🎯 Trò Chơi & Đố Vui":
    st.subheader("🎯 Sáng Tác Trò Chơi, Câu Đố & Bài Thơ Mầm Non")
    topic_game = st.text_input("Nhập chủ đề hoặc sự vật (VD: Củ cà rốt, An toàn giao thông, Đếm đến 5...)")
    genre = st.selectbox("Thể loại cần tạo", ["Bài thơ ngắn dễ thuộc", "Bài đồng dao vui nhộn", "Câu đố vui cho trẻ", "Trò chơi vận động tập thể"])
    
    if st.button("✨ Sáng Tác Ngay", type="primary"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            prompt = f"Hãy sáng tác hoặc thiết kế một {genre} dành riêng cho trẻ mầm non với chủ đề: {topic_game}. Ngôn ngữ trong trẻo, vui tươi. Nếu là trò chơi cần nêu rõ mục đích, luật chơi và cách chơi."
            with st.spinner("Đang sáng tác nội dung..."):
                res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                st.success("Đã hoàn thành tác phẩm cho cô:")
                st.markdown(res.text, unsafe_allow_html=True)

# --- TAB 4: SỔ LIÊN LẠC & ZALO ---
elif app_mode == "💬 Sổ Liên Lạc & Zalo":
    st.subheader("💬 Trợ Lý Viết Sổ Liên Lạc & Tin Nhắn Zalo Phụ Huynh")
    task_type = st.radio("Chọn loại nội dung cần viết:", ["Tin nhắn tổng kết ngày/tuần gửi nhóm Zalo", "Viết lời nhận xét sự tiến bộ cho từng bé"])
    
    if task_type == "Tin nhắn tổng kết ngày/tuần gửi nhóm Zalo":
        summary_topic = st.text_input("Nội dung hoạt động chính (VD: Hôm nay các con học bài thơ 'Cục len', tập tô chữ o...)")
        if st.button("💌 Soạn Tin Nhắn Zalo", type="primary"):
            if not api_key:
                st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
            else:
                client = genai.Client(api_key=api_key)
                prompt = f"Soạn một tin nhắn Zalo thân thiện, chuyên nghiệp từ giáo viên mầm non gửi phụ huynh tóm tắt các hoạt động: {summary_topic}. Kèm lời dặn dò nhẹ nhàng buổi tối."
                with st.spinner("Đang soạn tin..."):
                    res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                    st.success("Mẫu tin nhắn Zalo:")
                    st.markdown(res.text, unsafe_allow_html=True)
    else:
        child_info = st.text_input("Nhập vài từ khóa về tình hình của bé (VD: Bé Bảo An: ngoan, ăn giỏi, tích cực phát biểu nhưng còn rụt rè giờ thể chất)")
        if st.button("✍️ Viết Lời Nhận Xét", type="primary"):
            if not api_key:
                st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
            else:
                client = genai.Client(api_key=api_key)
                prompt = f"Viết một đoạn nhận xét cuối tuần khéo léo, tình cảm, mang tính khích lệ cao dành cho học sinh mầm non dựa trên thông tin: {child_info}."
                with st.spinner("Đang viết lời nhận xét..."):
                    res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                    st.success("Lời nhận xét mẫu cho cô:")
                    st.markdown(res.text, unsafe_allow_html=True)

# --- TAB 5: GỢI Ý ĐỒ DÙNG DIY ---
elif app_mode == "✂️ Gợi Ý Đồ Dùng DIY":
    st.subheader("✂️ Gợi Ý Làm Đồ Dùng Dạy Học Từ Vật Liệu Mở")
    lesson_name = st.text_input("Nhập tên bài học hoặc kỹ năng cần làm đồ dùng (VD: Đếm số lượng trong phạm vi 5...)")
    
    if st.button("🛠️ Gợi Ý Cách Làm", type="primary"):
        if not api_key:
            st.error("⚠️ Vui lòng nhập và lưu API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            prompt = f"Gợi ý cho giáo viên mầm non cách tận dụng các nguyên vật liệu phế liệu/vật liệu mở (chai nhựa, ống hút, bìa carton...) để làm đồ dùng đồ chơi tự tạo phục vụ cho bài học: {lesson_name}. Liệt kê nguyên liệu và các bước làm ngắn gọn, an toàn cho trẻ."
            with st.spinner("Đang lên ý tưởng chế tạo..."):
                res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                st.success("Ý tưởng đồ dùng tự tạo:")
                st.markdown(res.text, unsafe_allow_html=True)

# --- TAB 6: KHO LỊCH SỬ GIÁO ÁN ---
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
                        info_dummy = {"school": "Mầm Non Phước Thạnh", "teacher": "Nguyễn Thị Thanh Vân", "class_name": "Lớp Chồi 1", "theme": "Chung", "domain": "Nhận thức", "sub_category": "Toán", "activity_type": "Hoạt động học", "route_type": "Khám phá", "topic": title_db, "method": "Truyền thống", "age": "4-5 tuổi", "extras": []}
                        f_path = export_docx_from_text(content_db, info_dummy)
                        with open(f_path, "rb") as f_dl:
                            cols[2].download_button("📥 Tải File", data=f_dl, file_name=f_path, key=f"dl_{p_id}")
