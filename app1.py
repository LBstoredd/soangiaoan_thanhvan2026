import sqlite3
import json
import streamlit as st
from google import genai
from google.genai import types
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. Định nghĩa hàm khởi tạo database ở trên cùng (Tự động tạo bảng & bổ sung đầy đủ cột)
def init_db():
    conn = sqlite3.connect('mamnon_rag.db')
    cursor = conn.cursor()
    
    # Tạo bảng saved_plans nếu chưa có
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS saved_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Kiểm tra và tự động thêm các cột còn thiếu nếu dùng database cũ
    cursor.execute("PRAGMA table_info(saved_plans)")
    columns = [column[1] for column in cursor.fetchall()]
    
    if 'topic_title' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN topic_title TEXT")
    if 'school_name' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN school_name TEXT")
    if 'teacher_name' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN teacher_name TEXT")
    if 'class_name' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN class_name TEXT")
    if 'age_group' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN age_group TEXT")
    if 'domain' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN domain TEXT")
    if 'sub_domain' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN sub_domain TEXT")
    if 'method' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN method TEXT")
    if 'full_plan_json' not in columns:
        cursor.execute("ALTER TABLE saved_plans ADD COLUMN full_plan_json TEXT")
        
    conn.commit()
    conn.close()

# 2. Gọi hàm khởi tạo ngay sau khi đã định nghĩa xong
init_db()

# --- CẤU TRÚC ÉP KIỂU AI TRẢ VỀ CHUẨN JSON ---
class ActivityStep(BaseModel):
    step_name: str
    teacher_action: str
    children_action: str

class LessonPlan(BaseModel):
    knowledge_goal: list[str]
    skill_goal: list[str]
    attitude_goal: list[str]
    teacher_prep: str
    children_prep: str
    steps: list[ActivityStep]

# --- HÀM KẾT NỐI DATABASE (RAG) ---
def get_sample_from_db(domain, age, method):
    conn = sqlite3.connect("mamnon_rag.db")
    cursor = conn.cursor()
    
    cursor.execute("SELECT knowledge_standard, skill_standard, attitude_standard FROM age_standards WHERE domain = ? AND (age_group = ? OR age_group = 'Mọi độ tuổi') LIMIT 1", (domain, age))
    std = cursor.fetchone()
    
    cursor.execute("SELECT topic_title, goals_json, prep_json, steps_json FROM sample_plans WHERE method = ? LIMIT 1", (method,))
    sample = cursor.fetchone()
    if not sample:
        cursor.execute("SELECT topic_title, goals_json, prep_json, steps_json FROM sample_plans LIMIT 1")
        sample = cursor.fetchone()
        
    conn.close()
    
    std_text = f"Kiến thức chuẩn: {std[0]}\nKỹ năng chuẩn: {std[1]}\nThái độ chuẩn: {std[2]}" if std else "Bám sát chuẩn chương trình mầm non."
    sample_text = f"Phương pháp mẫu: {method}\nCấu trúc mẫu: {sample[3]}" if sample else "Cấu trúc chuẩn 3 phần."
    
    return std_text, sample_text

def get_random_activity_from_db(theme):
    conn = sqlite3.connect("mamnon_rag.db")
    cursor = conn.cursor()
    cursor.execute("SELECT title, content FROM activities_bank WHERE main_theme = ? ORDER BY RANDOM() LIMIT 1", (theme,))
    act = cursor.fetchone()
    conn.close()
    return act if act else ("Trò chơi vận động sáng tạo", "Cô tổ chức cho trẻ vận động theo nhạc vui tươi kết hợp trò chơi phản ứng nhanh.")

def save_plan_to_db(info, plan):
    conn = sqlite3.connect("mamnon_rag.db")
    cursor = conn.cursor()
    full_json = json.dumps(plan.model_dump(), ensure_ascii=False)
    cursor.execute('''
        INSERT INTO saved_plans (teacher_name, class_name, school_name, age_group, domain, sub_domain, method, topic_title, full_plan_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        info['teacher'], info['class_name'], info['school'], info['age'], 
        info['domain'], 'Hoạt động học', info['method'], info['topic'], full_json
    ))
    conn.commit()
    conn.close()

def load_saved_plans():
    conn = sqlite3.connect("mamnon_rag.db")
    cursor = conn.cursor()
    cursor.execute("SELECT id, topic_title, created_at, school_name FROM saved_plans ORDER BY id DESC LIMIT 10")
    rows = cursor.fetchall()
    conn.close()
    return rows

# --- HÀM GỌI GEMINI AI TẠO GIÁO ÁN THEO PHƯƠNG PHÁP CHUYÊN SÂU ---
def generate_ai_plan(api_key, topic, domain, age, theme, method):
    client = genai.Client(api_key=api_key)
    std_text, sample_text = get_sample_from_db(domain, age, method)
    
    method_instruction = ""
    if method == "STEAM 5E":
        method_instruction = "BẮT BUỘC phần Cách tiến hành phải chia đúng 5 bước chuẩn STEAM 5E: E1: Gắn kết (Engage), E2: Khám phá (Explore), E3: Giải thích (Explain), E4: Áp dụng / Mở rộng (Elaborate), E5: Đánh giá (Evaluate)."
    elif method == "STEAM EDP":
        method_instruction = "BẮT BUỘC phần Cách tiến hành phải chia đúng 5 bước quy trình thiết kế kỹ thuật EDP: Bước 1: Hỏi/Xác định vấn đề (Ask), Bước 2: Tưởng tượng (Imagine), Bước 3: Lập kế hoạch/Vẽ bản thiết kế (Plan), Bước 4: Chế tạo (Create), Bước 5: Thử nghiệm và Cải tiến (Improve)."
    else:
        method_instruction = "Phần Cách tiến hành chia theo 3 hoạt động truyền thống: Hoạt động 1: Gây hứng thú, Hoạt động 2: Nội dung trọng tâm, Hoạt động 3: Kết thúc."

    prompt = f"""
    Bạn là chuyên gia thiết kế chương trình giáo dục mầm non tại Việt Nam.
    Hãy soạn giáo án chi tiết và chuẩn xác theo các thông tin sau:
    - Độ tuổi: {age}
    - Chủ đề: {theme}
    - Lĩnh vực phát triển: {domain}
    - Đề tài cụ thể: {topic}
    - Phương pháp giảng dạy: {method}
    
    QUY CHUẨN BẮT BUỘC:
    1. Xưng hô: Dùng "Cô" và "Trẻ" (Tuyệt đối không dùng Giáo viên / Học sinh).
    2. Mục đích - Yêu cầu phải bám sát chuẩn độ tuổi: {std_text}
    3. {method_instruction}
    4. Chi tiết lời dẫn của cô, câu hỏi đàm thoại và dự kiến phản hồi của trẻ.
    
    THAM KHẢO CẤU TRÚC MẪU:
    {sample_text}
    """

    response = client.models.generate_content(
        model='gemini-2.0-flash',
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=LessonPlan,
            temperature=0.7,
        ),
    )
    return LessonPlan.model_validate_json(response.text)

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
        p_school.add_run(f"Trường: {info['school']}\n").bold = True
        p_school.add_run(f"Giáo viên: {info['teacher']} | Lớp: {info['class_name']}")
        
        title = doc.add_heading('\nGIÁO ÁN TỔ CHỨC HOẠT ĐỘNG HỌC', level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        p_info = doc.add_paragraph()
        p_info.add_run(f"• Chủ đề lớn: {info['theme']}\n").bold = True
        p_info.add_run(f"• Lĩnh vực: {info['domain']}\n")
        p_info.add_run(f"• Đề tài: {info['topic']} (Phương pháp: {info['method']})\n").bold = True
        p_info.add_run(f"• Đối tượng: {info['age']}\n")

        doc.add_heading('I. MỤC ĐÍCH - YÊU CẦU', level=2)
        doc.add_paragraph('1. Kiến thức:', style='List Number')
        for i in plan.knowledge_goal: doc.add_paragraph(f"- {i}")
        doc.add_paragraph('2. Kỹ năng:', style='List Number')
        for i in plan.skill_goal: doc.add_paragraph(f"- {i}")
        doc.add_paragraph('3. Thái độ:', style='List Number')
        for i in plan.attitude_goal: doc.add_paragraph(f"- {i}")

        doc.add_heading('II. CHUẨN BỊ', level=2)
        doc.add_paragraph(f"1. Đồ dùng của cô: {plan.teacher_prep}")
        doc.add_paragraph(f"2. Đồ dùng của trẻ: {plan.children_prep}")

        doc.add_heading('III. CÁCH TIẾN HÀNH', level=2)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        hdr = table.rows[0].cells
        hdr[0].text = 'Hoạt động của cô'
        hdr[1].text = 'Hoạt động của trẻ'

        for step in plan.steps:
            row = table.add_row().cells
            row[0].text = f"{step.step_name}\n\n{step.teacher_action}"
            row[1].text = step.children_action

    filename = f"GiaoAn_{info['topic'][:10].strip()}.docx"
    doc.save(filename)
    return filename

# --- GIAO DIỆN STREAMLIT (CHUYÊN SÂU) ---
st.set_page_config(page_title="Phần Mềm Soạn Giáo Án Mầm Non AI (Pro by Thanh Vân)", layout="wide")

st.title("👶 Phần Mềm Soạn Giáo Án Mầm Non Tự Động (Phiên bản Chuyên Sâu Pro-Thanh vân 2026)")
st.write("Hỗ trợ soạn giáo án đa phương pháp (Truyền thống, STEAM 5E, STEAM EDP) và tùy chỉnh theo mẫu file Word của trường.")

with st.sidebar:
    st.header("🔑 Cấu hình & Mẫu trường")
    api_key = st.text_input("Nhập Google Gemini API Key:", type="password")
    
    st.markdown("---")
    st.subheader("📂 Tải lên File Mẫu của Trường")
    uploaded_template = st.file_uploader("Chọn file Word (.docx) mẫu chuẩn:", type=["docx"])
    if uploaded_template:
        with open("temp_template.docx", "wb") as f:
            f.write(uploaded_template.getbuffer())
        st.success("Đã tải mẫu Word thành công!")

    st.markdown("---")
    st.subheader("📚 Lịch sử giáo án gần đây")
    saved_list = load_saved_plans()
    if saved_list:
        for item in saved_list:
            school_display = item[3] if len(item) > 3 and item[3] else "Chưa rõ"
            st.text(f"• [{item[2][:10]}] {item[1]} ({school_display})")

if not api_key:
    st.warning("⚠️ Vui lòng nhập Google Gemini API Key ở thanh bên trái để bắt đầu sử dụng.")
else:
    col1, col2 = st.columns(2)
    
    with col1:
        teacher = st.text_input("Họ và tên giáo viên:", "Nguyễn Thị Thanh Vân")
        school = st.text_input("Tên trường mầm non:", "Trường Mầm Non Phước Thạnh")
        class_name = st.text_input("Tên lớp:", "Lớp Chồi 1 (4-5 tuổi)")
        age = st.selectbox("Độ tuổi:", ['24-36 tháng', '3-4 tuổi', '4-5 tuổi', '5-6 tuổi'])

    with col2:
        domain = st.selectbox("Lĩnh vực phát triển:", ['Thể chất', 'Nhận thức', 'Ngôn ngữ', 'TC-KNXH', 'Thẩm mỹ'])
        theme = st.selectbox("Chủ đề lớn:", [
            'Trường mầm non', 'Bản thân', 'Gia đình', 'Nghề nghiệp', 
            'Thế giới động vật', 'Thế giới thực vật', 'Giao thông', 
            'Hiện tượng tự nhiên', 'Quê hương - Đất nước - Bác Hồ'
        ])
        method = st.selectbox("Phương pháp giảng dạy chuyên sâu:", ['Truyền thống', 'STEAM 5E', 'STEAM EDP', 'Montessori'])
        topic = st.text_input("Tên đề tài bài dạy:", "Dự án STEAM: Chế tạo ô tô chạy bằng sức gió...")

    st.markdown("---")
    
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        generate_clicked = st.button("🚀 Soạn Giáo Án Theo Phương Pháp Chuyên Sâu", type="primary", use_container_width=True)
    with col_btn2:
        random_act_clicked = st.button("🎲 Đổi Trò Chơi", use_container_width=True)

    if generate_clicked:
        if not topic:
            st.error("Vui lòng nhập tên đề tài bài dạy!")
        else:
            with st.spinner(f"🤖 AI đang áp dụng phương pháp **{method}** để soạn giáo án, vui lòng đợi trong giây lát..."):
                try:
                    info = {
                        "teacher": teacher, "school": school, "class_name": class_name,
                        "age": age, "domain": domain, "theme": theme, "method": method, "topic": topic
                    }
                    plan = generate_ai_plan(api_key, topic, domain, age, theme, method)
                    
                    save_plan_to_db(info, plan)
                    st.session_state['current_plan'] = plan
                    st.session_state['current_info'] = info
                    st.success("🎉 Đã hoàn thành giáo án chuyên sâu!")
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi khi gọi AI: {e}")

    if random_act_clicked:
        act_title, act_content = get_random_activity_from_db(theme)
        st.toast(f"✨ Gợi ý trò chơi mới:\n\n**{act_title}**\n{act_content}", icon="🎲")

    if 'current_plan' in st.session_state:
        plan = st.session_state['current_plan']
        info = st.session_state['current_info']
        
        st.markdown("### 📋 Xem trước nội dung giáo án chuyên sâu:")
        
        with st.expander("Xem chi tiết Mục đích - Yêu cầu & Chuẩn bị", expanded=True):
            st.markdown(f"**I. MỤC ĐÍCH - YÊU CẦU**")
            st.write("1. Kiến thức:", plan.knowledge_goal)
            st.write("2. Kỹ năng:", plan.skill_goal)
            st.write("3. Thái độ:", plan.attitude_goal)
            st.markdown(f"**II. CHUẨN BỊ**")
            st.write(f"- Đồ dùng của cô: {plan.teacher_prep}")
            st.write(f"- Đồ dùng của trẻ: {plan.children_prep}")

        with st.expander("Xem chi tiết Cách tiến hành theo phương pháp mới", expanded=True):
            for step in plan.steps:
                st.markdown(f"**{step.step_name}**")
                c_col1, c_col2 = st.columns(2)
                with c_col1:
                    st.info(f"**Hoạt động của cô:**\n\n{step.teacher_action}")
                with c_col2:
                    st.success(f"**Hoạt động của trẻ:**\n\n{step.children_action}")

        # Xuất file (dùng mẫu tùy chỉnh nếu có)
        template_path = "temp_template.docx" if uploaded_template else None
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
