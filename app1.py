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

# --- BỘ LỌC LÀM SẠCH TỰ ĐỘNG (GỠ SẠCH DẤU *, **, ###, ---) ---
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

# --- DATABASE KHUNG 9 THÁNG MẦM NON MẪU ---
SAMPLE_9_MONTHS = {
    "Tháng 9": {
        "theme": "Trường mầm non thân yêu của bé",
        "goals": "- Thể chất: Trẻ thích nghi với chế độ sinh hoạt ở trường, thực hiện các vận động đi, chạy nhẹ nhàng.\n- Nhận thức: Trẻ biết tên trường, lớp, tên cô giáo và một số bạn thân thiết trong lớp.\n- Ngôn ngữ: Trẻ biết chào hỏi lễ phép, nói trọn câu khi bày tỏ mong muốn.\n- Tình cảm - KNXH: Trẻ tự tin khi đến lớp, biết cất đồ chơi gọn gàng đúng nơi quy định.\n- Thẩm mỹ: Trẻ hào hứng hát múa các bài ca về trường lớp mầm non.",
        "content_net": "- Tuần 1: Trường mầm non của bé (Làm quen trường, lớp và góc chơi).\n- Tuần 2: Cô giáo và các bạn thân yêu.\n- Tuần 3: Đồ dùng, đồ chơi quen thuộc trong lớp học của bé.\n- Tuần 4: Ngày hội trăng rằm của bé (Tết Trung Thu rộn ràng).",
        "activities": "Khám phá trường lớp; Thơ: Bạn mới; Hát: Trường chúng cháu là trường mầm non; VĐCB: Đi theo đường hẹp; Vẽ bóng bay tặng bạn."
    },
    "Tháng 10": {
        "theme": "Bản thân và các giác quan kỳ diệu",
        "goals": "- Thể chất: Trẻ phối hợp tay mắt khéo léo, nhận biết các món ăn có lợi cho sức khỏe.\n- Nhận thức: Trẻ nhận biết giới tính của bản thân, chức năng chính của 5 giác quan.\n- Ngôn ngữ: Trẻ biết giới thiệu tên, tuổi, sở thích và dùng từ ngữ chỉ các bộ phận cơ thể.\n- Tình cảm - KNXH: Trẻ biết giữ gìn vệ sinh thân thể, có ý thức tự phục vụ bản thân.\n- Thẩm mỹ: Trẻ biết tạo hình khuôn mặt, tô màu các trang phục yêu thích.",
        "content_net": "- Tuần 1: Tôi là ai? (Tên gọi, tuổi, giới tính và sở thích của bé).\n- Tuần 2: Cơ thể bé và 5 giác quan kỳ diệu.\n- Tuần 3: Bé cần gì để lớn lên khỏe mạnh? (Dinh dưỡng và vệ sinh).\n- Tuần 4: Trang phục đáng yêu của bé theo mùa.",
        "activities": "Toán: Nhận biết phía phải - phía trái bản thân; Truyện: Tay phải tay trái; Tạo hình: Cắt dán bạn trai bạn gái; Dạy hát: Cái mũi."
    },
    "Tháng 11": {
        "theme": "Những nghề nghiệp quen thuộc trong xã hội",
        "goals": "- Thể chất: Rèn luyện sức bền, sự nhanh nhẹn thông qua các thao tác mô phỏng nghề nghiệp.\n- Nhận thức: Trẻ biết tên gọi, trang phục, dụng cụ và sản phẩm của một số nghề gần gũi.\n- Ngôn ngữ: Sử dụng đúng từ ngữ chuyên ngành nghề, đọc thơ kể chuyện truyền cảm.\n- Tình cảm - KNXH: Trẻ biết ơn các thầy cô giáo, trân trọng người lao động và yêu quý sản phẩm lao động.\n- Thẩm mỹ: Thể hiện tình cảm qua bài hát, tranh vẽ tri ân cô giáo nhân ngày 20/11.",
        "content_net": "- Tuần 1: Nghề giáo viên cao quý (Chào mừng ngày Nhà giáo Việt Nam 20/11).\n- Tuần 2: Nghề bác sĩ, y tá chăm sóc sức khỏe cộng đồng.\n- Tuần 3: Nghề xây dựng và kiến trúc sư tài ba.\n- Tuần 4: Nghề nông và các nghề sản xuất dịch vụ khác.",
        "activities": "Khám phá: Dụng cụ nghề y; Thơ: Cô giáo của em; Hát: Cháu yêu cô chú công nhân; Tạo hình: Làm bưu thiếp tặng cô."
    },
    "Tháng 12": {
        "theme": "Thế giới động vật quanh bé",
        "goals": "- Thể chất: Trẻ bắt chước các động tác vận động của con vật: bò, nhảy, bật xa nhịp nhàng.\n- Nhận thức: Trẻ phân biệt được đặc điểm sinh sống, thức ăn, tiếng kêu của các nhóm con vật.\n- Ngôn ngữ: Bắt chước âm thanh, kể chuyện sáng tạo về các con vật nuôi và thú rừng.\n- Tình cảm - KNXH: Trẻ có tình yêu thương động vật, không trêu chọc thú dữ, bảo vệ môi trường sống.\n- Thẩm mỹ: Trẻ nặn, xé dán con vật và múa minh họa các bài hát về muôn thú.",
        "content_net": "- Tuần 1: Những con vật nuôi gần gũi trong gia đình (Chó, mèo, gà, vịt).\n- Tuần 2: Những con thú hoang dã trong rừng sâu (Voi, hổ, khỉ, gấu).\n- Tuần 3: Động vật bơi lội dưới nước (Cá, tôm, cua, mực).\n- Tuần 4: Côn trùng bé nhỏ và chim muông biết bay.",
        "activities": "VĐCB: Bật qua vật cản 15-20cm; Khám phá: Vòng đời của bướm; Thơ: Đàn gà con; Hát: Chú voi con ở Bản Đôn; Nặn con cá nhỏ."
    },
    "Tháng 1": {
        "theme": "Thế giới thực vật - Tết và Mùa xuân",
        "goals": "- Thể chất: Rèn luyện sự khéo léo của các ngón tay qua hoạt động chăm sóc cây và gói bánh.\n- Nhận thức: Nhận biết các loại hoa quả đặc trưng mùa xuân, phong tục đón Tết cổ truyền.\n- Ngôn ngữ: Trẻ biết nói lời chúc Tết ông bà, bố mẹ lễ phép, trôi chảy.\n- Tình cảm - KNXH: Trẻ tự hào về văn hóa truyền thống, yêu thiên nhiên cỏ cây hoa lá.\n- Thẩm mỹ: Cảm nhận vẻ đẹp rực rỡ của hoa mai, hoa đào mùa xuân qua nét vẽ, lời ca.",
        "content_net": "- Tuần 1: Các loài hoa rực rỡ đón xuân (Hoa mai, hoa đào, hoa cúc).\n- Tuần 2: Cây xanh và các loại quả ngọt mùa xuân.\n- Tuần 3: Phong tục Tết Nguyên Đán cổ truyền và món bánh chưng xanh.\n- Tuần 4: Lễ hội mùa xuân và các trò chơi dân gian rộn ràng.",
        "activities": "Tạo hình: Trang trí cành mai ngày Tết; Khám phá: Sự phát triển của hạt giống; Toán: Đếm số lượng quả trong phạm vi 5; Hát: Sắp đến Tết rồi."
    },
    "Tháng 2": {
        "theme": "Giao thông và an toàn trên đường",
        "goals": "- Thể chất: Phát triển định hướng không gian, đi giữ thăng bằng theo tín hiệu giao thông.\n- Nhận thức: Phân biệt các loại phương tiện giao thông (đường bộ, sắt, thủy, hàng không) và biển báo cơ bản.\n- Ngôn ngữ: Nhận diện và gọi đúng tên biển báo, trả lời nhanh các tình huống luật giao thông.\n- Tình cảm - KNXH: Trẻ có ý thức tuân thủ luật an toàn giao thông, đội mũ bảo hiểm khi ngồi xe máy.\n- Thẩm mỹ: Gấp, vẽ, dán các phương tiện giao thông sáng tạo.",
        "content_net": "- Tuần 1: Phương tiện giao thông đường bộ quen thuộc (Xe đạp, xe máy, ô tô).\n- Tuần 2: Phương tiện giao thông đường sắt và đường hàng không.\n- Tuần 3: Phương tiện giao thông đường thủy lướt sóng.\n- Tuần 4: Bé học luật giao thông và đèn tín hiệu đường phố.",
        "activities": "VĐCB: Bò chui qua cổng hẹp; Khám phá: Đèn tín hiệu giao thông; Thơ: Cô dạy con; Tạo hình: Cắt dán thuyền buồm trên biển; Hát: Em đi qua ngã tư đường phố."
    },
    "Tháng 3": {
        "theme": "Nước và các hiện tượng tự nhiên",
        "goals": "- Thể chất: Tăng cường vận động ngoài trời, rèn thói quen uống nước ấm sạch.\n- Nhận thức: Hiểu vòng tuần hoàn của nước, ích lợi của nước và tác hại của thời tiết cực đoan.\n- Ngôn ngữ: Dùng từ miêu tả hiện tượng: nắng, mưa, gió, bão, cầu vồng rực rỡ.\n- Tình cảm - KNXH: Trẻ có ý thức tiết kiệm nguồn nước sạch, giữ gìn vệ sinh môi trường sống.\n- Thẩm mỹ: Cảm thụ âm thanh tự nhiên của tiếng mưa rơi, tiếng suối chảy qua âm nhạc.",
        "content_net": "- Tuần 1: Nguồn nước sạch và ích lợi của nước đối với sự sống.\n- Tuần 2: Các hiện tượng thời tiết: Nắng, mưa, gió, mây.\n- Tuần 3: Bốn mùa kỳ diệu trong năm (Xuân, Hạ, Thu, Đông).\n- Tuần 4: Bảo vệ nguồn nước và môi trường sống của bé.",
        "activities": "Thí nghiệm: Sự hòa tan và chìm nổi; Toán: So sánh dung tích 2 bình nước; Truyện: Giọt nước tí xíu; Dạy hát: Cho tôi đi làm mưa với."
    },
    "Tháng 4": {
        "theme": "Quê hương - Đất nước - Bác Hồ kính yêu",
        "goals": "- Thể chất: Rèn luyện tính kiên trì, dẻo dai qua các trò chơi vận động mô phỏng danh lam thắng cảnh.\n- Nhận thức: Trẻ biết tên nước Việt Nam, thủ đô Hà Nội, nhận biết lăng Bác và cờ đỏ sao vàng.\n- Ngôn ngữ: Đọc thơ, hát rõ lời, tự hào giới thiệu về quê hương em đang sống.\n- Tình cảm - KNXH: Trẻ thể hiện lòng kính yêu Bác Hồ, tự hào là người con đất Việt.\n- Thẩm mỹ: Vẽ hoa dâng Bác, biểu diễn văn nghệ truyền thống dân tộc.",
        "content_net": "- Tuần 1: Quê hương tươi đẹp và danh lam thắng cảnh địa phương.\n- Tuần 2: Thủ đô Hà Nội ngàn năm văn hiến và bản đồ Việt Nam.\n- Tuần 3: Bác Hồ kính yêu với thiếu nhi mầm non.\n- Tuần 4: Lễ hội giỗ tổ Hùng Vương và truyền thống uống nước nhớ nguồn.",
        "activities": "Khám phá: Lăng Bác Hồ; Thơ: Ảnh Bác; Hát: Nhớ ơn Bác; Tạo hình: Xé dán cờ Tổ quốc; VĐCB: Chạy nhanh 15m."
    },
    "Tháng 5": {
        "theme": "Trường tiểu học - Tạm biệt mầm non",
        "goals": "- Thể chất: Chuẩn bị thể lực tốt, khả năng ngồi thẳng lưng và vận động tinh khéo léo.\n- Nhận thức: Làm quen đồ dùng học tập lớp 1: bút chì, thước kẻ, sách giáo khoa; nhận biết môi trường tiểu học.\n- Ngôn ngữ: Phát âm chuẩn 29 chữ cái, diễn đạt ý nghĩ tự tin trước tập thể.\n- Tình cảm - KNXH: Tâm thế hào hứng bước vào lớp 1, lưu luyến nhớ ơn các cô giáo mầm non.\n- Thẩm mỹ: Làm album kỷ niệm lưu giữ khoảnh khắc tuổi mầm non.",
        "content_net": "- Tuần 1: Trường tiểu học của em (Làm quen quang cảnh và phòng học lớp 1).\n- Tuần 2: Đồ dùng học tập của học sinh lớp 1.\n- Tuần 3: Mừng ngày sinh nhật Bác Hồ 19/5.\n- Tuần 4: Tạm biệt trường mầm non - Chào lớp 1 thân yêu.",
        "activities": "Khám phá: Chiếc cặp sách của học sinh lớp 1; Thơ: Bé vào lớp 1; Dạy hát: Tạm biệt búp bê thân yêu; Tạo hình: Làm bưu thiếp kỷ niệm."
    }
}

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

# --- BỘ KHUNG SƯ PHẠM CHUẨN 100/100 ĐIỂM (TAB 1: GIÁO ÁN NGÀY) ---
def generate_instant_pedagogical_plan(info: dict) -> str:
    today_str = datetime.now().strftime("%Y-%m-%d")
    topic = info.get("topic", "Hoạt động khám phá")
    domain = info.get("domain", "Nhận thức")
    sub_cat = info.get("sub_category", "Khám phá")
    route = info.get("route_type", "Hoạt động chung")
    art_type = info.get("art_type", "")
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

    art_note = f" (Dạng tiết: {art_type})" if domain == "Thẩm mỹ" and sub_cat == "Tạo hình" and art_type else ""

    if custom_goals:
        goals_block = custom_goals
    else:
        goals_block = (
            f"1. Kiến thức:\n"
            f"- Trẻ nhận biết, gọi đúng tên và nắm được đặc điểm cốt lõi của nội dung bài học: {topic}{art_note} phù hợp độ tuổi {age}.\n"
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
    elif domain == "Thẩm mỹ" and sub_cat == "Tạo hình":
        if art_type == "Tiết theo mẫu":
            main_activity = (
                f"Hoạt động 2: Trọng tâm - Tạo hình: {topic} (Tiết theo mẫu)\n"
                f"* Bước 1: Quan sát và đàm thoại mẫu:\n"
                f"- Cô cho trẻ quan sát bức tranh/sản phẩm mẫu của cô. Đàm thoại về màu sắc, bố cục, đường nét và cách thể hiện '{topic}'.\n"
                f"- Cô hỏi ý định của trẻ về cách thực hiện.\n"
                f"* Bước 2: Cô làm mẫu:\n"
                f"- Cô làm mẫu cho trẻ quan sát từng bước (chọn màu, cách cầm bút/cầm kéo/cách nặn, bố cục sắp xếp trên giấy).\n"
                f"- Nhấn mạnh các điểm trọng tâm để sản phẩm đẹp và sáng tạo.\n"
                f"* Bước 3: Trẻ thực hiện:\n"
                f"- Nhắc nhở trẻ tư thế ngồi, cách cầm dụng cụ.\n"
                f"- Trẻ thực hiện bài vẽ/nặn/cắt dán theo mẫu của cô. Cô bao quát lớp, gợi ý giúp đỡ trẻ yếu và khuyến khích trẻ sáng tạo thêm chi tiết phụ.\n"
                f"* Bước 4: Trưng bày và nhận xét sản phẩm:\n"
                f"- Trẻ trưng bày sản phẩm lên giá treo.\n"
                f"- Cô mời trẻ lên nhận xét bài của mình, bài của bạn thích nhất và vì sao; cô nhận xét chung, động viên khích lệ."
            )
        else:
            main_activity = (
                f"Hoạt động 2: Trọng tâm - Tạo hình: {topic} (Tiết theo đề tài)\n"
                f"* Bước 1: Gợi mở và thỏa thuận ý tưởng:\n"
                f"- Cô trò chuyện cùng trẻ về đề tài '{topic}', gợi ý để trẻ tự nhớ lại đặc điểm hình dáng, màu sắc.\n"
                f"- Hỏi ý định của từng trẻ: Con sẽ vẽ/nặn/cắt dán như thế nào? Dùng màu gì? Sắp xếp bố cục ra sao?\n"
                f"* Bước 2: Trẻ thực hiện:\n"
                f"- Trẻ tiến hành sáng tạo tác phẩm theo đề tài '{topic}' của riêng mình.\n"
                f"- Cô bao quát lớp, khích lệ sự sáng tạo, độc đáo của trẻ; hỗ trợ trẻ còn lúng túng trong kỹ năng thực hiện.\n"
                f"* Bước 3: Trưng bày và chia sẻ cảm xúc:\n"
                f"- Trưng bày toàn bộ sản phẩm của lớp.\n"
                f"- Cô cho trẻ tự giới thiệu về bức tranh/sản phẩm sáng tạo của mình; cô nhận xét tuyên dương các ý tưởng mới lạ, sáng tạo."
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

    raw_plan = f"""{school.upper()}
Lớp: {class_name} | Độ tuổi: {age}
GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG ({info.get('activity_type', 'Hoạt động học')})
ĐỀ TÀI: {topic.upper()}{art_note}
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
    return clean_markdown_text(raw_plan)

# --- BỘ CỨU CÁNH KẾ HOẠCH THÁNG (VIDEO 1) ---
def generate_instant_month_plan(info: dict) -> str:
    school = info.get("school", "Trường Mầm Non")
    teacher = info.get("teacher", "Giáo viên")
    class_name = info.get("class_name", "Lớp Mẫu giáo")
    age = info.get("age", "4-5 tuổi")
    month = info.get("month", "Tháng 10")
    theme = info.get("theme", "Thế giới xung quanh")
    goals = info.get("goals", "").strip()
    content_net = info.get("content_net", "").strip()
    activities = info.get("activities", "").strip()
    notes = info.get("notes", "").strip()

    if not goals:
        goals = (
            "- Thể chất: Phát triển cơ lớn qua các bài tập vận động cơ bản, hoàn thiện kỹ năng phối hợp tay mắt và rèn luyện thói quen tự phục vụ.\n"
            "- Nhận thức: Mở rộng hiểu biết về các hiện tượng, sự vật đặc trưng của chủ đề; hình thành biểu tượng toán học sơ đẳng.\n"
            "- Ngôn ngữ: Mở rộng vốn từ vựng, tự tin biểu đạt cảm xúc và lời nói mạch lạc qua các bài thơ, câu đố và câu chuyện kể.\n"
            "- Tình cảm - KNXH: Hình thành thái độ thân thiện, hợp tác cùng bạn bè, yêu quý thiên nhiên và biết tôn trọng các quy định chung.\n"
            "- Thẩm mỹ: Cảm thụ vẻ đẹp rực rỡ của môi trường xung quanh, tích cực tham gia các hoạt động tạo hình và biểu diễn âm nhạc."
        )

    if not content_net:
        content_net = (
            f"- Tuần 1: Khám phá mở đầu về '{theme}'.\n"
            f"- Tuần 2: Các đặc điểm cấu tạo, môi trường sống và hoạt động nổi bật.\n"
            f"- Tuần 3: Trải nghiệm thực hành, ứng dụng sáng tạo gắn với cuộc sống của bé.\n"
            f"- Tuần 4: Tổng kết chủ đề, triển lãm sản phẩm bé ngoan và chuẩn bị chuyển tiếp chủ đề mới."
        )

    if not activities:
        activities = (
            "- Khám phá khoa học: Quan sát vật thật, đàm thoại và làm thí nghiệm trải nghiệm.\n"
            "- Làm quen với toán: Đếm số lượng, so sánh kích thước, phân biệt hình khối và định hướng không gian.\n"
            "- Âm nhạc: Dạy hát đúng cao độ, rèn kỹ năng gõ đệm theo tiết tấu và múa minh họa.\n"
            "- Văn học: Kể chuyện theo tranh, đọc thơ diễn cảm và tập đóng kịch phân vai.\n"
            "- Tạo hình: Rèn kỹ năng vẽ nét cong tròn, cắt dán theo mẫu, nặn và làm đồ dùng tự tạo từ vật liệu mở."
        )

    notes_text = f"\nVI. LƯU Ý & YÊU CẦU BỔ SUNG ĐẶC BIỆT:\n{notes}" if notes else ""

    text = f"""{school.upper()}
Lớp: {class_name} | Độ tuổi: {age}
KẾ HOẠCH GIÁO DỤC THÁNG: {month.upper()}
CHỦ ĐỀ CHÍNH: {theme.upper()}
Giáo viên thực hiện: {teacher}

I. MỤC TIÊU PHÁT TRIỂN THÁNG (5 MẶT CHUẨN BỘ GD&ĐT):
{goals}

II. MẠNG NỘI DUNG PHÂN BỔ THEO TUẦN:
{content_net}

III. MẠNG HOẠT ĐỘNG HỌC CHỦ ĐẠO TRONG THÁNG:
{activities}

IV. KẾ HOẠCH HOẠT ĐỘNG GÓC THÁNG:
- Góc Phân vai: Bé đóng vai bác sĩ, cô giáo, người bán hàng, nội trợ gia đình gắn liền với chủ đề '{theme}'.
- Góc Xây dựng - Lắp ghép: Xây dựng khu vườn, trang trại, công viên hoặc các công trình mô phỏng chủ đề.
- Góc Nghệ thuật - Tạo hình: Cắt dán bưu thiếp, vẽ tranh hoa lá, nặn các chi tiết và làm đồ chơi tự tạo từ phế liệu sạch.
- Góc Học tập - Sách: Nhận diện chữ cái, chữ số, chơi lô tô, ghép tranh theo chủ đề và xem truyện tranh mầm non.
- Góc Khám phá thiên nhiên: Chăm sóc cây xanh, lau lá, tưới nước và quan sát các thí nghiệm khoa học đơn giản.

V. PHỐI HỢP CÙNG PHỤ HUYNH TRONG THÁNG:
- Thông báo tới phụ huynh nội dung chủ đề qua bảng tuyên truyền và nhóm Zalo của lớp.
- Vận động phụ huynh ủng hộ các nguyên vật liệu mở, phế liệu sạch để cùng cô làm đồ dùng học tập cho các con.
- Trao đổi định kỳ tình hình sức khỏe, thói quen ăn ngủ và tiến bộ tâm lý của trẻ.{notes_text}"""
    return clean_markdown_text(text)

# --- BỘ CỨU CÁNH KẾ HOẠCH TUẦN (VIDEO 2: MA TRẬN 5 NGÀY) ---
def generate_instant_week_plan(info: dict) -> str:
    school = info.get("school", "Trường Mầm Non")
    teacher = info.get("teacher", "Giáo viên")
    class_name = info.get("class_name", "Lớp Mẫu giáo")
    age = info.get("age", "4-5 tuổi")
    week_num = info.get("week_num", "Tuần 1")
    theme_month = info.get("theme_month", "Chủ đề tháng")
    theme_week = info.get("theme_week", "Nhánh tuần")
    time_range = info.get("time_range", "Thứ Hai đến Thứ Sáu")
    matrix = info.get("matrix", {})
    notes = info.get("notes", "").strip()

    def get_cell(day, field, default_val):
        val = matrix.get(day, {}).get(field, "").strip()
        return val if val else default_val

    t2_hoc = get_cell("t2", "hoc", "VĐCB: Đi trong đường hẹp đầu đội túi cát")
    t2_troi = get_cell("t2", "troi", "Quan sát thời tiết buổi sáng; TCVĐ: Mèo đuổi chuột")
    t2_goc = get_cell("t2", "goc", "Góc phân vai: Bác sĩ khám bệnh; Góc xây dựng: Lắp ghép công viên")
    t2_chieu = get_cell("t2", "chieu", "Ôn luyện bài tập buổi sáng; Làm quen bài thơ mới")

    t3_hoc = get_cell("t3", "hoc", "Khám phá: Tìm hiểu về đặc điểm của các sự vật quanh bé")
    t3_troi = get_cell("t3", "troi", "Dạo chơi vườn trường; Nhặt lá cây rụng làm đồ chơi")
    t3_goc = get_cell("t3", "goc", "Góc học tập: Xếp tương ứng 1-1; Góc tạo hình: Tô màu tranh vẽ")
    t3_chieu = get_cell("t3", "chieu", "Hướng dẫn kỹ năng rửa tay bằng xà phòng đúng cách")

    t4_hoc = get_cell("t4", "hoc", "Làm quen với toán: Đếm đến 3, nhận biết chữ số 3")
    t4_troi = get_cell("t4", "troi", "Lắng nghe âm thanh thiên nhiên; TCDG: Kéo co")
    t4_goc = get_cell("t4", "goc", "Góc âm nhạc: Biểu diễn bài hát chủ đề; Góc thiên nhiên: Tưới cây")
    t4_chieu = get_cell("t4", "chieu", "Kể chuyện theo tranh: Bài học tình bạn; Chơi tự do")

    t5_hoc = get_cell("t5", "hoc", "Văn học: Truyện kể truyền cảm hứng cho bé")
    t5_troi = get_cell("t5", "troi", "Vẽ phấn trên sân trường; Chơi với cát và nước")
    t5_goc = get_cell("t5", "goc", "Góc thư viện: Xem sách tranh mầm non; Góc đóng vai bán hàng")
    t5_chieu = get_cell("t5", "chieu", "Rèn kỹ năng gấp quần áo, cất dọn đồ chơi ngăn nắp")

    t6_hoc = get_cell("t6", "hoc", "Tạo hình / Âm nhạc: Dạy hát bài ca vui nhộn")
    t6_troi = get_cell("t6", "troi", "Giao lưu thể thao nhẹ giữa các tổ; Chơi tự do")
    t6_goc = get_cell("t6", "goc", "Tổng kết các góc chơi; Triển lãm sản phẩm tạo hình trong tuần")
    t6_chieu = get_cell("t6", "chieu", "Nêu gương bé ngoan cuối tuần; Trao cờ bé chăm ngoan")

    notes_text = f"\nIV. LƯU Ý BỔ SUNG TRONG TUẦN:\n{notes}" if notes else ""

    text = f"""{school.upper()}
Lớp: {class_name} | Độ tuổi: {age}
KẾ HOẠCH GIÁO DỤC TUẦN: {week_num.upper()}
Chủ đề lớn: {theme_month} | Nhánh tuần: {theme_week.upper()}
Thời gian thực hiện: {time_range} | Giáo viên: {teacher}

I. MỤC TIÊU TRỌNG TÂM TRONG TUẦN:
- Rèn luyện kỹ năng vận động, giữ thăng bằng và phản xạ nhanh nhẹn trong các hoạt động tập thể.
- Khắc sâu kiến thức nhánh tuần '{theme_week}', rèn luyện cách phát âm và diễn đạt câu hoàn chỉnh.
- Thực hiện nghiêm túc nề nếp sinh hoạt, đoàn kết chia sẻ đồ dùng đồ chơi và giữ gìn vệ sinh cá nhân.

II. KẾ HOẠCH TỔ CHỨC HOẠT ĐỘNG HÀNG NGÀY (THỨ HAI ĐẾN THỨ SÁU):

1. THỨ HAI:
- Đón trẻ & Thể dục sáng: Đón trẻ ân cần, kiểm tra thân nhiệt; tập bài tập phát triển chung nhịp nhàng.
- Hoạt động học: {t2_hoc}
- Hoạt động ngoài trời: {t2_troi}
- Hoạt động chơi ở các góc: {t2_goc}
- Hoạt động chiều: {t2_chieu}

2. THỨ
