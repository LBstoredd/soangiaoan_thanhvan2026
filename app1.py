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
    # Bảng lưu giáo án
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS saved_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Bảng quản lý mã kích hoạt thuê tháng
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS licenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            license_key TEXT UNIQUE,
            customer_name TEXT,
            expires_at DATE,
            is_active INTEGER DEFAULT 1
        )
    ''')
    
    # Tạo sẵn một mã dùng thử mặc định cho anh test: MÃ LÀ "LBSTORE2026" (hạn 1 năm)
    cursor.execute('''
        INSERT OR IGNORE INTO licenses (id, license_key, customer_name, expires_at, is_active)
        VALUES (1, 'LBSTORE2026', 'Admin Văn Bình (Master Key)', '2027-12-31', 1)
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
            # Kiểm tra hạn sử dụng
            exp_date = datetime.strptime(expires_at, "%Y-%m-%d").date()
            if datetime.now().date() <= exp_date:
                return True, f"✅ Kích hoạt thành công! Xin chào: {c_name} (Hạn đến: {expires_at})"
            else:
                return False, "❌ Mã kích hoạt này đã hết hạn sử dụng. Vui lòng liên hệ gia hạn (100k/tháng)!"
    return False, "❌ Mã kích hoạt không hợp lệ hoặc không tồn tại!"

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

# --- HÀM XUẤT FILE WORD (.DOCX) ---
def export_docx(plan: LessonPlan, info: dict, custom_template=None):
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
        p_info.add_run(f"• Lĩnh vực: {info['domain']}\n")
        p_info.add_run(f"• Đề tài: {info['topic']} (Phương pháp: {info['method']})\n").bold = True
        p_info.add_run(f"• Đối tượng: {info['age']}")

    doc.add_heading('I. Mục đích yêu cầu', level=2)
    for obj in plan.objectives:
        doc.add_paragraph(f"- {obj}", style='List Bullet')

    doc.add_heading('II. Chuẩn bị', level=2)
    doc.add_paragraph("Đồ dùng của cô:", style='Normal').runs[0].bold = True
    for item in plan.teacher_prep:
        doc.add_paragraph(f"- {item}", style='List Bullet')
    doc.add_paragraph("Đồ dùng của trẻ:", style='Normal').runs[0].bold = True
    for item in plan.children_prep:
        doc.add_paragraph(f"- {item}", style='List Bullet')

    doc.add_heading('III. Cách tiến hành', level=2)
    for idx, step in enumerate(plan.steps, 1):
        doc.add_paragraph(f"Bước {idx}: {step.get('step_name', '')}", style='Normal').runs[0].bold = True
        doc.add_paragraph(f"Hoạt động của cô: {step.get('teacher_action', '')}")
        doc.add_paragraph(f"Hoạt động của trẻ: {step.get('children_action', '')}")

    doc.add_heading('IV. Đánh giá sau giờ dạy', level=2)
    doc.add_paragraph(plan.assessment)

    filename = f"GiaoAn_{info['topic'].replace(' ', '_')}.docx"
    doc.save(filename)
    return filename

# --- CẤU HÌNH GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Phần Mềm Soạn Giáo Án Mầm Non Tự Động", layout="wide")

# Kiểm tra trạng thái đăng nhập qua mã bản quyền trong session_state
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# --- MÀN HÌNH KHÓA / NHẬP MÃ BẢN QUYỀN ---
if not st.session_state["authenticated"]:
    st.title("🔐 Xác Thực Sử Dụng Phần Mềm Mầm Non")
    st.markdown("Phần mềm hỗ trợ soạn giáo án chuẩn chuyên sâu. Vui lòng nhập **Mã Kích Hoạt Bản Quyền (100k/tháng)** để truy cập hệ thống.")
    
    col_l1, col_l2 = st.columns([2, 1])
    with col_l1:
        input_key = st.text_input("Nhập Mã Kích Hoạt Của Bạn", type="password", placeholder="Nhập mã do admin cấp...")
        if st.button("🔓 Mở Khóa Sử Dụng", type="primary"):
            valid, msg = check_license(input_key)
            if valid:
                st.session_state["authenticated"] = True
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)
    with col_l2:
        st.info("💡 **Gói Cước Thuê Thêm:**\n- Giá cước: **100.000 VNĐ / Tháng**\n- Liên hệ Zalo/Hotline Admin:\n  `0982.155.559` (Văn Bình)\n- Chuyển khoản nhận mã ngay lập tức!")
    
    st.stop() # Dừng lại không cho hiển thị phần mềm bên dưới nếu chưa nhập đúng mã

# --- GIAO DIỆN CHÍNH SAU KHI ĐÃ ĐƯỢC MỞ KHÓA ---
st.title("🧸 Phần Mềm Soạn Giáo Án Mầm Non Tự Động (Thanh Vân 2026)")
st.markdown("Hỗ trợ soạn giáo án đa phương pháp, sáng tác thơ/đồng dao, gợi ý đồ chơi DIY và tóm tắt tin nhắn phụ huynh.")

# Thanh sidebar cấu hình
with st.sidebar:
    st.header("⚙️ Cấu hình hệ thống")
    api_key = st.text_input("Nhập Google Gemini API Key", type="password")
    
    st.divider()
    if st.button("🔒 Đăng Xuất / Khóa Lại"):
        st.session_state["authenticated"] = False
        st.rerun()
        
    st.divider()
    st.info("💡 **Tiện ích bổ sung cho cô giáo:**")
    app_mode = st.radio("Chọn chức năng:", ["📝 Soạn Giáo Án Chuẩn", "📚 Kho Lịch Sử Giáo Án", "🎨 Sáng Tác Thơ & Đồng Dao", "✂️ Gợi Ý Đồ Dùng DIY", "💬 Tóm Tắt Tin Nhắn Zalo Phụ Huynh"])

# --- TAB 1: SOẠN GIÁO ÁN CHUẨN ---
if app_mode == "📝 Soạn Giáo Án Chuẩn":
    col1, col2 = st.columns(2)
    with col1:
        teacher_name = st.text_input("Họ và tên giáo viên", "Nguyễn Thị Thanh Vân")
        school_name = st.text_input("Tên trường mầm non", "Mầm Non Phước Thạnh")
        class_name = st.text_input("Tên lớp", "Lớp Chồi 1 (4-5 tuổi)")
        age_group = st.selectbox("Độ tuổi", ["24-36 tháng", "3-4 tuổi", "4-5 tuổi", "5-6 tuổi"])
    with col2:
        domain = st.selectbox("Lĩnh vực phát triển", ["Thể chất", "Nhận thức", "Ngôn ngữ", "Thẩm mỹ", "Tình cảm - KNXH"])
        theme = st.text_input("Chủ đề lớn", "Trường mầm non")
        method = st.selectbox("Phương pháp giảng dạy chuyên sâu", ["Truyền thống", "STEAM 5E", "STEAM EDP"])
        topic = st.text_input("Tên đề tài bài dạy", "Học chữ cái")

    uploaded_template = st.file_uploader("Tải lên mẫu file Word (.docx) của trường (Tùy chọn)", type=["docx"])

    if st.button("🚀 Soạn Giáo Án Theo Phương Pháp Chuyên Sâu", type="primary"):
        if not api_key:
            st.error("Vui lòng nhập Google Gemini API Key ở thanh bên trái!")
        else:
            try:
                client = genai.Client(api_key=api_key)
                prompt = f"""
                Bạn là chuyên gia giáo dục mầm non chuẩn Bộ GD&ĐT Việt Nam. Hãy soạn một giáo án chi tiết chuẩn xác bằng tiếng Việt với các thông tin sau:
                - Tên bài dạy: {topic}
                - Lĩnh vực: {domain}
                - Chủ đề: {theme}
                - Độ tuổi: {age_group}
                - Phương pháp: {method}
                
                Yêu cầu văn phong sư phạm mầm non chuẩn mực (xưng hô cô - trẻ, có hoạt động gây hứng thú, đàm thoại, thực hành và nhận xét). Trả về đúng định dạng cấu trúc dữ liệu JSON.
                """
                with st.spinner("AI đang phân tích và biên soạn giáo án chuẩn mầm non..."):
                    response = client.models.generate_content(
                        model='gemini-3.8-flash',
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=LessonPlan,
                            temperature=0.7,
                        ),
                    )
                    plan = LessonPlan.model_validate_json(response.text)
                    
                    info = {
                        "school": school_name, "teacher": teacher_name, "class_name": class_name,
                        "theme": theme, "domain": domain, "topic": topic, "method": method, "age": age_group
                    }
                    
                    # Lưu tự động vào Database lịch sử
                    save_plan_to_db(topic, response.text)

                st.success("🎉 Đã soạn giáo án thành công và tự động lưu vào kho lịch sử!")
                
                # Hiển thị kết quả trực quan
                st.subheader(f"📖 {plan.title}")
                st.markdown(f"**🎯 Mục đích yêu cầu:**")
                for obj in plan.objectives:
                    st.markdown(f"- {obj}")
                
                with st.expander("Xem chi tiết Cách tiến hành theo phương pháp mới", expanded=True):
                    for step in plan.steps:
                        st.markdown(f"**{step.get('step_name')}**")
                        c_col1, c_col2 = st.columns(2)
                        with c_col1:
                            st.info(f"**Hoạt động của cô:**\n\n{step.get('teacher_action')}")
                        with c_col2:
                            st.success(f"**Hoạt động của trẻ:**\n\n{step.get('children_action')}")

                template_path = uploaded_template if uploaded_template else None
                docx_file = export_docx(plan, info, template_path)

                with open(docx_file, "rb") as file:
                    st.download_button(
                        label="📥 Tải xuống File Word (.docx) chuẩn",
                        data=file,
                        file_name=docx_file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi gọi AI: {e}")

# --- TAB 2: KHO LỊCH SỬ GIÁO ÁN ---
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
                        title_db, json_db = row
                        plan_obj = LessonPlan.model_validate_json(json_db)
                        info_dummy = {"school": "Mầm Non Phước Thạnh", "teacher": "Nguyễn Thị Thanh Vân", "class_name": "Lớp Chồi 1", "theme": "Chung", "domain": "Chung", "topic": title_db, "method": "Truyền thống", "age": "4-5 tuổi"}
                        f_path = export_docx(plan_obj, info_dummy)
                        with open(f_path, "rb") as f_dl:
                            cols[2].download_button("📥 Tải File", data=f_dl, file_name=f_path, key=f"dl_{p_id}")

# --- TAB 3: SÁNG TÁC THƠ & ĐỒNG DAO ---
elif app_mode == "🎨 Sáng Tác Thơ & Đồng Dao":
    st.subheader("🎨 Sáng Tác Thơ, Câu Đố & Đồng Dao Mầm Non")
    topic_poetry = st.text_input("Nhập chủ đề hoặc tên sự vật (VD: Củ cà rốt, Chú gà con, Đi vệ sinh sạch sẽ...)")
    genre = st.selectbox("Thể loại cần sáng tác", ["Bài thơ ngắn (4-8 câu)", "Bài đồng dao vui nhộn", "Câu đố vui cho trẻ"])
    
    if st.button("✨ Sáng Tác Ngay", type="primary"):
        if not api_key:
            st.error("Vui lòng nhập API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            prompt = f"Hãy sáng tác một {genre} dành riêng cho trẻ em lứa tuổi mầm non với chủ đề: {topic_poetry}. Ngôn ngữ trong trẻo, dễ thuộc, vui nhộn."
            with st.spinner("Đang sáng tác..."):
                res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                st.success("Đã hoàn thành tác phẩm cho cô:")
                st.markdown(res.text)

# --- TAB 4: GỢI Ý ĐỒ DÙNG DIY ---
elif app_mode == "✂️ Gợi Ý Đồ Dùng DIY":
    st.subheader("✂️ Gợi Ý Làm Đồ Dùng Dạy Học Từ Vật Liệu Mở")
    lesson_name = st.text_input("Nhập tên bài học hoặc kỹ năng cần làm đồ dùng (VD: Đếm số lượng trong phạm vi 5, Phân loại động vật nuôi...)")
    
    if st.button("🛠️ Gợi Ý Cách Làm", type="primary"):
        if not api_key:
            st.error("Vui lòng nhập API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            prompt = f"Gợi ý cho giáo viên mầm non cách tận dụng các nguyên vật liệu phế liệu/vật liệu mở (chai nhựa, ống hút, bìa carton, vải vụn...) để làm đồ dùng đồ chơi tự tạo phục vụ cho bài học: {lesson_name}. Liệt kê nguyên liệu và các bước làm ngắn gọn, an toàn cho trẻ."
            with st.spinner("Đang lên ý tưởng chế tạo..."):
                res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                st.success("Ý tưởng đồ dùng tự tạo:")
                st.markdown(res.text)

# --- TAB 5: TÓM TẮT TIN NHẮN ZALO PHỤ HUYNH ---
elif app_mode == "💬 Tóm Tắt Tin Nhắn Zalo Phụ Hyuynh":
    st.subheader("💬 Tạo Tin Nhắn Tóm Tắt Gửi Nhóm Zalo Phụ Huynh")
    summary_topic = st.text_input("Nội dung chính trong ngày/tuần (VD: Hôm nay các con học bài thơ 'Cục len', tập tô chữ o, ăn hết suất...)")
    
    if st.button("💌 Soạn Tin Nhắn Gửi Phụ Huynh", type="primary"):
        if not api_key:
            st.error("Vui lòng nhập API Key ở menu bên trái!")
        else:
            client = genai.Client(api_key=api_key)
            prompt = f"Soạn một đoạn tin nhắn Zalo thân thiện, lịch sự, chuyên nghiệp từ giáo viên mầm non gửi cho phụ huynh học sinh để tóm tắt các hoạt động: {summary_topic}. Có kèm lời dặn dò nhẹ nhàng và chúc phụ huynh buổi tối vui vẻ."
            with st.spinner("Đang soạn tin nhắn..."):
                res = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
                st.success("Mẫu tin nhắn Zalo:")
                st.markdown(res.text)
