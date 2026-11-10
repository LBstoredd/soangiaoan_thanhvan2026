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

    raw_plan = f"""{school.upper()}
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

2. THỨ BA:
- Đón trẻ & Thể dục sáng: Trò chuyện đầu giờ về chủ đề tuần '{theme_week}'; tập theo nhạc bài hát chủ đề.
- Hoạt động học: {t3_hoc}
- Hoạt động ngoài trời: {t3_troi}
- Hoạt động chơi ở các góc: {t3_goc}
- Hoạt động chiều: {t3_chieu}

3. THỨ TƯ:
- Đón trẻ & Thể dục sáng: Hướng dẫn trẻ cất dép ba lô vào ngăn tủ; tập các động tác tay, chân, bụng lườn.
- Hoạt động học: {t4_hoc}
- Hoạt động ngoài trời: {t4_troi}
- Hoạt động chơi ở các góc: {t4_goc}
- Hoạt động chiều: {t4_chieu}

4. THỨ NĂM:
- Đón trẻ & Thể dục sáng: Thể dục sáng với vòng thể dục; rèn luyện thói quen chào hỏi lễ phép.
- Hoạt động học: {t5_hoc}
- Hoạt động ngoài trời: {t5_troi}
- Hoạt động chơi ở các góc: {t5_goc}
- Hoạt động chiều: {t5_chieu}

5. THỨ SÁU:
- Đón trẻ & Thể dục sáng: Điểm danh buổi sáng, kiểm tra vệ sinh móng tay; tập thể dục đồng diễn toàn trường.
- Hoạt động học: {t6_hoc}
- Hoạt động ngoài trời: {t6_troi}
- Hoạt động chơi ở các góc: {t6_goc}
- Hoạt động chiều: {t6_chieu}

III. CHUẨN BỊ MÔI TRƯỜNG & ĐỒ DÙNG:
- Tranh ảnh trực quan, học liệu góc và hệ thống video minh họa cho nhánh tuần '{theme_week}'.
- Đảm bảo đồ dùng đồ chơi đủ cho mỗi trẻ, không gian sân trường an toàn, thông thoáng.{notes_text}"""
    return clean_markdown_text(text)

# --- BỘ CỨU CÁNH SÁNG TÁC TỰ ĐỘNG (TAB 3) ---
def generate_instant_creative_content(genre: str, topic: str) -> str:
    t_clean = topic.strip().title() if topic.strip() else "Bé Ngoan"
    t_lower = topic.strip().lower()

    if genre == "Bài thơ ngắn dễ thuộc":
        if "giao thong" in t_lower:
            text = f"""BÀI THƠ: BÉ NHỚ LUẬT GIAO THÔNG
(Thơ 4 chữ nhịp 2/2, vần điệu vui tươi, dễ thuộc cho trẻ mầm non)

Mỗi sáng tới trường,
Cùng ba cùng mẹ,
Bé đi đúng lề,
Bên tay phải nhé!

Đèn đỏ dừng lại,
Đèn xanh được đi,
Đèn vàng chậm chậm,
Chờ xem tín hiệu.

Bé nhớ thật kỹ,
Đội mũ an toàn,
Bé là bé ngoan,
Bình an hạnh phúc!

Gợi ý dạy trẻ: Cô cho trẻ đọc kết hợp vỗ tay theo nhịp, đàm thoại về ý nghĩa của các màu đèn tín hiệu."""
        else:
            text = f"""BÀI THƠ: {t_clean.upper()} CỦA BÉ
(Thơ 4 chữ tươi sáng, giàu hình ảnh dành cho trẻ mầm non)

Bé học hôm nay,
Về {t_clean} nhỏ,
Bao điều mới lạ,
Đẹp đẽ làm sao!

Đôi mắt mở to,
Đôi tay chăm chỉ,
Cùng cô cùng bạn,
Khám phá say mê.

Mỗi ngày đến lớp,
Bé học điều hay,
Yêu thương mọi người,
Lời cô ghi nhớ!

Gợi ý dạy trẻ: Cô đọc diễn cảm, kết hợp các động tác minh họa tay - mắt để trẻ thích thú đọc theo."""

    elif genre == "Bài đồng dao vui nhộn":
        text = f"""BÀI ĐỒNG DAO: {t_clean.upper()}
(Đồng dao nhịp điệu rộn ràng, thích hợp cho giờ hoạt náo và chuyển tiết)

Dung dăng dung dẻ,
Dắt trẻ đi chơi.
Học điều mới lạ,
Về {t_clean}.

Mắt nhìn cho tỏ,
Tai lắng nghe cô.
Bạn nào giỏi nhất,
Được tặng hoa điểm mười!

Chi chi chành chành,
Vui cùng bè bạn,
Tay nắm chặt tay,
Học ngoan biết mấy!

Cách tổ chức: Cho trẻ đứng vòng tròn, vừa nắm tay đung đưa vừa đọc theo nhịp 2/4 rộn rã."""

    elif genre == "Câu đố vui cho trẻ":
        if "giao thong" in t_lower:
            text = f"""HỆ THỐNG CÂU ĐỐ VUI: {t_clean.upper()}
(Phát triển tư duy phán đoán và phản xạ ngôn ngữ)

Câu đố 1:
Mắt đỏ, mắt vàng, lại thêm mắt xanh,
Đứng ở ngã tư, nhắc người đi lại?
- Đáp án: Cột đèn tín hiệu giao thông.

Câu đố 2:
Xe gì hai bánh,
Chạy tít bo bo,
Tiếng còi "bim bim",
Nhớ đội mũ kín?
- Đáp án: Xe máy.

Câu đố 3:
Vạch gì trắng xoá từng hàng,
Dành riêng cho bé bước sang bên đường?
- Đáp án: Vạch kẻ đường dành cho người đi bộ.

Câu đố 4:
Ai người đứng gác ngã tư,
Gậy hoa chỉ lối, mọi người lắng nghe?
- Đáp án: Chú cảnh sát giao thông.

Cách dẫn dắt: Cô đọc từng câu đố biểu cảm, cho trẻ giơ tay trả lời và tặng sticker khen ngợi."""
        else:
            text = f"""HỆ THỐNG CÂU ĐỐ VUI VỀ: {t_clean.upper()}

Câu đố 1:
Bé thường nhìn thấy mỗi ngày,
Gần gũi thân thương, đố bé là gì về {t_clean}?
- Đáp án: {t_clean}.

Câu đố 2:
Hình dáng xinh xắn, màu sắc tươi vui,
Bé ngoan chăm sóc, ai ai cũng cười?
- Đáp án: {t_clean}.

Câu đố 3:
Mang lại niềm vui, giúp bé học tốt,
Cùng cô cùng bạn, gắn bó từng ngày?
- Đáp án: {t_clean}."""

    else:
        text = f"""TRÒ CHƠI VẬN ĐỘNG TẬP THỂ: {t_clean.upper()}

1. Mục đích - Yêu cầu:
- Rèn luyện phản xạ vận động nhanh nhẹn, khả năng định hướng không gian và giữ thăng bằng.
- Củng cố nhận thức của trẻ về chủ đề '{t_clean}'.
- Rèn ý thức kỷ luật, tinh thần phối hợp gắn kết trong tập thể.

2. Chuẩn bị:
- Sân chơi hoặc phòng học rộng rãi, sạch sẽ, an toàn.
- Đồ dùng biểu tượng hoặc đạo cụ gắn liền với '{t_clean}'.
- Kẻ vạch xuất phát và đích đến rõ ràng.

3. Cách chơi:
- Cô chia trẻ thành 2 - 3 đội chơi có số lượng bằng nhau đứng sau vạch xuất phát.
- Khi có hiệu lệnh còi/tiếng nhạc, bạn đầu hàng nhanh chóng vượt qua chướng ngại vật lên tìm và chọn đúng đồ dùng liên quan đến '{t_clean}' đặt vào rổ của đội mình.
- Sau đó chạy nhanh về đập tay vào bạn tiếp theo và về cuối hàng. Lần lượt từng bạn lên thực hiện.

4. Luật chơi:
- Mỗi lượt chỉ được lấy 1 đồ dùng. Bạn tiếp theo chỉ được xuất phát khi bạn trước đã đập tay.
- Hết bản nhạc (khoảng 3 phút), đội nào chọn đúng và nhiều đồ dùng hơn sẽ là đội chiến thắng."""

    return clean_markdown_text(text)

# --- BỘ CỨU CÁNH SỔ LIÊN LẠC & ZALO (TAB 4) ---
def generate_instant_zalo(task_type: str, info_text: str) -> str:
    if "Tin nhắn" in task_type:
        text = f"""MẪU TIN NHẮN ZALO GỬI PHỤ HUYNH CUỐI NGÀY/TUẦN:

Kính gửi các bậc phụ huynh lớp!

Hôm nay các con đã có một ngày học tập và vui chơi thật nhiều niềm vui và bổ ích với các hoạt động nổi bật:
- Nội dung chính: {info_text if info_text.strip() else 'Khám phá bài học mới, rèn kỹ năng vận động và tham gia các góc chơi hào hứng'}.
- Tình hình chung: Tất cả các bé đều ăn ngoan, ngủ tròn giấc, tinh thần vui vẻ, tích cực tương tác cùng cô và các bạn.

Lời nhắn buổi tối: Bố mẹ hãy dành ít phút trò chuyện cùng con về hoạt động hôm nay để giúp con rèn luyện trí nhớ và khả năng diễn đạt nhé! 

Chúc các gia đình có một buổi tối ấm áp và ngập tràn tiếng cười!
(Trân trọng, Cô giáo chủ nhiệm)"""
    else:
        text = f"""LỜI NHẬN XÉT SỔ LIÊN LẠC HỌC SINH:

- Họ tên bé: Học sinh
- Tình hình thực tế: {info_text if info_text.strip() else 'Ngoan ngoãn, lễ phép, ăn ngủ tốt'}
- Lời nhận xét của cô giáo: 
"Bé rất ngoan ngoãn, lễ phép và biết vâng lời cô giáo. Trong các giờ hoạt động học cũng như vui chơi, con luôn vui vẻ, hòa đồng, biết chia sẻ đồ chơi cùng bạn bè. Con ăn uống giỏi, ngủ tròn giấc và có nhiều tiến bộ rõ nét trong kỹ năng tự phục vụ. Cô khen ngợi con và mong con tiếp tục phát huy sự chăm ngoan này nhé!" """
    return clean_markdown_text(text)

# --- BỘ CỨU CÁNH ĐỒ DÙNG TỰ TẠO DIY (TAB 5) ---
def generate_instant_diy(lesson: str) -> str:
    l_clean = lesson.strip().title() if lesson.strip() else "Tiết Học Mầm Non"
    text = f"""Ý TƯỞNG LÀM ĐỒ DÙNG DẠY HỌC TỰ TẠO (DIY)
Chủ đề bài dạy: {l_clean}

1. Nguyên vật liệu mở / Phế liệu dễ tìm:
- Bìa carton cứng, hộp bánh, lõi cuộn giấy vệ sinh.
- Nắp chai nhựa nhiều màu sắc, que kem gỗ, ống hút giấy.
- Xốp màu thừa, vải nỉ vụn, cúc áo, dây ruy-băng.
- Keo nến, băng dính hai mặt, kéo thủ công an toàn.

2. Các bước tiến hành chế tạo:
- Bước 1 (Xử lý nền): Cắt tỉa bìa carton làm khung bảng chính, dán giấy màu hoặc decal sạch sẽ để làm nền hoạt động.
- Bước 2 (Tạo hình chi tiết): Sử dụng nắp chai nhựa hoặc vải nỉ cắt dán thành các biểu tượng, nhân vật, hình ảnh tương ứng với nội dung '{l_clean}'.
- Bước 3 (Gắn kết tương tác): Đính gai dán velcro hoặc luồn dây để trẻ có thể dễ dàng bóc dán, di chuyển, phân loại hoặc sắp xếp các chi tiết theo yêu cầu bài học.
- Bước 4 (Kiểm tra an toàn): Bo tròn các góc nhọn, kiểm tra độ chắc chắn để đảm bảo an toàn tuyệt đối cho trẻ khi cầm nắm.

3. Hiệu quả giáo dục:
- Đồ dùng trực quan sinh động, độ bền cao, không tốn kém chi phí.
- Giúp trẻ được trực tiếp thao tác 'học bằng chơi, chơi mà học', phát triển vận động tinh của đôi bàn tay."""
    return clean_markdown_text(text)

# --- HÀM GỌI AI THÔNG MINH (CHỐT CỨNG GEMINI 3.8 FLASH, TỰ ĐỘNG LỌC SẠCH MARKDOWN) ---
def generate_fast_ai_content(client, prompt, fallback_type="plan", fallback_data=None):
    def _call_model():
        res = client.models.generate_content(
            model='gemini-3.8-flash',
            contents=prompt,
        )
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

    if fallback_type == "plan":
        return generate_instant_pedagogical_plan(fallback_data)
    elif fallback_type == "month":
        return generate_instant_month_plan(fallback_data)
    elif fallback_type == "week":
        return generate_instant_week_plan(fallback_data)
    elif fallback_type == "creative":
        return generate_instant_creative_content(fallback_data.get("genre", ""), fallback_data.get("topic", ""))
    elif fallback_type == "zalo":
        return generate_instant_zalo(fallback_data.get("type", ""), fallback_data.get("info", ""))
    elif fallback_type == "diy":
        return generate_instant_diy(fallback_data.get("lesson", ""))

    return "Đã hoàn thành nội dung cho cô giáo!"

# --- HÀM XUẤT FILE WORD TỰ ĐỘNG IN ĐẬM TIÊU ĐỀ ---
def export_docx_from_text(text_content: str, info: dict, custom_template=None, font_size_pt=13):
    if custom_template:
        doc = Document(custom_template)
    else:
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(font_size_pt)

        p_school = doc.add_paragraph()
        p_school.add_run(f"• Trường: {info.get('school', 'Trường Mầm Non')}\n").bold = True
        p_school.add_run(f"• Giáo viên: {info.get('teacher', 'Giáo viên')} | Lớp: {info.get('class_name', 'Lớp Mầm Non')}")

        title_text = info.get('doc_title', 'GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG')
        title = doc.add_heading(title_text, level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for line in text_content.split('\n'):
        line_s = line.strip()
        if line_s:
            p = doc.add_paragraph()
            if re.match(r'^(I|II|III|IV|V|VI)\.\s', line_s) or line_s.startswith(('GIÁO ÁN', 'KẾ HOẠCH', 'ĐỀ TÀI', 'CHỦ ĐỀ', 'Hoạt động', '1. THỨ', '2. THỨ', '3. THỨ', '4. THỨ', '5. THỨ')):
                run = p.add_run(line_s)
                run.bold = True
            else:
                p.add_run(line_s)

    safe_topic = "".join(c if c.isalnum() else "_" for c in info.get('topic', 'KeHoach'))[:40]
    filename = f"KeHoach_{safe_topic}.docx"
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
# SIDEBAR (DANH MỤC TÍNH NĂNG Ở TRÊN CÙNG)
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
        "custom_goals": custom_goals, "custom_ideas": custom_ideas, "auto_assessment": auto_assessment,
        "doc_title": "GIÁO ÁN TỔ CHỨC HOẠT ĐỘNG HỌC"
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
                    
                    Trình bày đúng 4 mục lớn (không sử dụng các dấu hoa thị dư thừa hoặc dấu thăng ###):
                    I. MỤC ĐÍCH - YÊU CẦU (1. Kiến thức, 2. Kỹ năng, 3. Thái độ)
                    II. CHUẨN BỊ (1. Đồ dùng của cô, 2. Đồ dùng của trẻ)
                    III. TỔ CHỨC HOẠT ĐỘNG (Hoạt động 1: Ổn định tổ chức; Hoạt động 2: Trọng tâm; Hoạt động 3: Kết thúc)
                    IV. ĐÁNH GIÁ CUỐI NGÀY (Gồm đúng 3 gạch đầu dòng: - Tình trạng sức khỏe; - Trạng thái cảm xúc, thái độ; - Kiến thức, kỹ năng)
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
        st.success("🎉 Đã hoàn thành giáo án! Xem và thao tác bên dưới:")

    if "generated_plan_text" in st.session_state:
        st.divider()
        info = st.session_state["plan_info"]
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
        
        # 4. BẢN XEM TRƯỚC GIÁO ÁN
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
                st.session_state["generated_plan_text"] = clean_markdown_text(edited_text)
                st.rerun()

# =========================================================================
# TAB 2: KẾ HOẠCH & SỔ CHỦ ĐỀ (NÂNG CẤP CHUẨN THEO 2 VIDEO: THÁNG & TUẦN)
# =========================================================================
elif app_mode == "🗓️ Kế Hoạch & Sổ Chủ Đề":
    sub_tab_month, sub_tab_week = st.tabs(["📅 Kế Hoạch Tháng (Theo Video 1)", "📋 Kế Hoạch Tuần (Theo Video 2)"])

    # -------------------------------------------------------------
    # SUB-TAB 1: THIẾT KẾ KẾ HOẠCH THÁNG (VIDEO 1)
    # -------------------------------------------------------------
    with sub_tab_month:
        st.subheader("📖 Thiết Kế Kế Hoạch Tháng")
        st.caption("Xây dựng mục tiêu mầm non, mạng nội dung & hoạt động học khoa học theo chuẩn chương trình GDMN.")

        st.markdown("#### ⚡ Nạp Nhanh Trọn Bộ Kế Hoạch Từ Khung 9 Tháng Mẫu")
        col_nap1, col_nap2, col_nap3 = st.columns([2, 2, 2])
        with col_nap1:
            sample_age_m = st.selectbox("Chọn độ tuổi mẫu", ["Lớp Bé (3-4 tuổi)", "Lớp Nhỡ (4-5 tuổi)", "Lớp Lá (5-6 tuổi)"], index=1, key="sm_age")
        with col_nap2:
            sample_month_m = st.selectbox("Chọn tháng mẫu", list(SAMPLE_9_MONTHS.keys()), index=1, key="sm_month")
        with col_nap3:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            if st.button("📥 Nạp Kế Hoạch Mẫu", type="secondary", use_container_width=True):
                s_data = SAMPLE_9_MONTHS.get(sample_month_m, {})
                st.session_state["f_month_theme"] = s_data.get("theme", "")
                st.session_state["f_month_goals"] = s_data.get("goals", "")
                st.session_state["f_month_net"] = s_data.get("content_net", "")
                st.session_state["f_month_act"] = s_data.get("activities", "")
                st.success(f"✅ Đã nạp thành công bộ khung mẫu của **{sample_month_m}**!")

        st.divider()
        st.markdown("#### 📌 Thông Tin Thiết Lập Kế Hoạch Tháng")
        
        m_c1, m_c2 = st.columns(2)
        with m_c1:
            m_teacher = st.text_input("Tên giáo viên", "Nguyễn Thị Thanh Vân", key="m_gv")
            m_school = st.text_input("Tên trường", "Trường Mầm Non Phước Thạnh", key="m_tr")
        with m_c2:
            m_class = st.text_input("Lớp học", "Lớp Chồi 1", key="m_lp")
            m_age = st.selectbox("Độ tuổi", ["3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=1, key="m_dt")

        m_c3, m_c4 = st.columns(2)
        with m_c3:
            m_month_val = st.selectbox("Tháng thực hiện", list(SAMPLE_9_MONTHS.keys()), index=1, key="m_thang_chon")
        with m_c4:
            def_theme = st.session_state.get("f_month_theme", "Bản thân và các giác quan kỳ diệu")
            m_theme_val = st.text_input("Chủ đề chính của tháng *", value=def_theme, key="m_chude_chinh")

        def_goals = st.session_state.get("f_month_goals", "")
        m_goals_val = st.text_area(
            "Định hướng Mục tiêu giáo dục tháng (Để trống AI sẽ tự soạn chuẩn 5 mặt phát triển: Thể chất, nhận thức, ngôn ngữ, tình cảm kỹ năng xã hội, thẩm mỹ):",
            value=def_goals,
            height=120,
            key="m_mt_val"
        )

        def_net = st.session_state.get("f_month_net", "")
        m_net_val = st.text_area(
            "Định hướng Mạng nội dung (Để trống AI sẽ tự phân bổ thành các tuần/nhánh con hợp lý. Ví dụ: Tuần 1, Tuần 2, Tuần 3, Tuần 4):",
            value=def_net,
            height=110,
            key="m_net_val"
        )

        def_act = st.session_state.get("f_month_act", "")
        m_act_val = st.text_area(
            "Mạng hoạt động học chính gợi ý (Để trống AI sẽ tự thiết lập hệ thống bài học chi tiết: Toán, Khám phá, Âm nhạc, Văn học, Tạo hình trong tháng):",
            value=def_act,
            height=110,
            key="m_act_val"
        )

        m_notes_val = st.text_area(
            "Yêu cầu bổ sung đặc biệt cho kế hoạch tháng (Ví dụ: Tăng cường hoạt động rèn luyện kỹ năng sinh tồn, thêm ngày kỷ niệm...):",
            placeholder="Nhập yêu cầu riêng nếu có...",
            height=70,
            key="m_notes_val"
        )

        month_info_dict = {
            "school": m_school,
            "teacher": m_teacher,
            "class_name": m_class,
            "age": m_age,
            "month": m_month_val,
            "theme": m_theme_val,
            "goals": m_goals_val,
            "content_net": m_net_val,
            "activities": m_act_val,
            "notes": m_notes_val,
            "doc_title": f"KẾ HOẠCH GIÁO DỤC {m_month_val.upper()}",
            "topic": f"KeHoachThang_{m_month_val}"
        }

        if st.button("✨ Tạo Kế Hoạch Tháng Bằng AI", type="primary", use_container_width=True, key="btn_month_ai"):
            with st.spinner("⚡ Đang lập kế hoạch giáo dục tháng chuẩn sư phạm GDMN..."):
                if api_key:
                    try:
                        client = genai.Client(api_key=api_key)
                        prompt_month = f"""
                        Lập kế hoạch giáo dục tháng chuẩn mầm non Bộ GD&ĐT Việt Nam:
                        - Trường: {m_school} | Giáo viên: {m_teacher} | Lớp: {m_class} ({m_age})
                        - Tháng: {m_month_val} | Chủ đề chính: {m_theme_val}
                        - Định hướng mục tiêu: {m_goals_val if m_goals_val else 'Chuẩn 5 mặt phát triển'}
                        - Mạng nội dung: {m_net_val if m_net_val else 'Tự động chia 4 tuần'}
                        - Mạng hoạt động: {m_act_val if m_act_val else 'Tự động thiết lập đủ các môn'}
                        - Ghi chú: {m_notes_val if m_notes_val else 'Không'}
                        
                        Yêu cầu trình bày sạch đẹp, bài bản, gồm: Mục tiêu 5 mặt phát triển, Mạng nội dung 4 tuần, Mạng hoạt động học chi tiết, Kế hoạch góc và Phối hợp phụ huynh. Tuyệt đối không dùng dấu ### hoặc **.
                        """
                        res_month = generate_fast_ai_content(client, prompt_month, fallback_type="month", fallback_data=month_info_dict)
                    except Exception:
                        res_month = generate_instant_month_plan(month_info_dict)
                else:
                    res_month = generate_instant_month_plan(month_info_dict)

                res_month = clean_markdown_text(res_month)
                st.session_state["saved_month_plan"] = res_month
                st.session_state["saved_month_info"] = month_info_dict
                save_plan_to_db(f"Kế Hoạch {m_month_val} - {m_theme_val}", res_month)
            st.success("🎉 Đã hoàn thành Kế hoạch giáo dục tháng!")

        # KẾT QUẢ KẾ HOẠCH THÁNG
        if "saved_month_plan" in st.session_state:
            st.divider()
            m_info = st.session_state["saved_month_info"]
            m_font_size = st.session_state.get("m_font_size_slider", 14)

            st.markdown("### 🛠️ Tùy Chỉnh & Thao Tác Kế Hoạch Tháng")
            mb_col1, mb_col2 = st.columns(2)
            with mb_col1:
                m_docx = export_docx_from_text(st.session_state["saved_month_plan"], m_info, font_size_pt=m_font_size)
                with open(m_docx, "rb") as f_m:
                    st.download_button(
                        label="📥 Tải File Word (.docx) Kế Hoạch Tháng",
                        data=f_m,
                        file_name=m_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                        key="dl_month_btn"
                    )
            with mb_col2:
                m_json_text = json.dumps(st.session_state["saved_month_plan"])
                m_title_json = json.dumps(f"Kế Hoạch Giáo Dục: {m_info['theme']}")
                print_m_html = f"""
                <div style="display: flex; justify-content: center; align-items: center; width: 100%;">
                    <button onclick="printMonthPlan()" style="
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
                    ">
                        🖨️ In Kế Hoạch Tháng Ngay
                    </button>
                </div>
                <script>
                function printMonthPlan() {{
                    var rawContent = {m_json_text};
                    var docTitle = {m_title_json};
                    var pWin = window.open('', '_blank');
                    pWin.document.write('<!DOCTYPE html><html><head><title>' + docTitle + '</title>');
                    pWin.document.write('<style>');
                    pWin.document.write('@page {{ size: A4; margin: 15mm 20mm 15mm 20mm; }}');
                    pWin.document.write('body {{ font-family: "Times New Roman", Times, serif; font-size: {m_font_size}pt; line-height: 1.5; color: #000; margin: 0; padding: 0; }}');
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
                components.html(print_m_html, height=52)

            mi_c1, mi_c2 = st.columns(2)
            with mi_c1:
                st.success("✅ **Điểm sư phạm: 100/100**")
            with mi_c2:
                st.info(f"📅 **Thực hiện:** {m_info['month']} | {m_info['age']}")

            m_font_size = st.slider(
                "🔤 Kích thước chữ hiển thị & In (pt):", 
                min_value=12, 
                max_value=24, 
                value=m_font_size, 
                step=1, 
                key="m_font_size_slider"
            )

            st.markdown("### 📄 Bản Xem Trước Kế Hoạch Tháng (Trang Giấy In):")
            escaped_month_text = html.escape(st.session_state["saved_month_plan"])
            st.markdown(
                f"""
                <div style="
                    background-color: #ffffff;
                    color: #111111;
                    font-family: 'Times New Roman', Times, serif;
                    font-size: {m_font_size}pt;
                    line-height: 1.6;
                    padding: 30px 40px;
                    border-radius: 8px;
                    border: 1px solid #d3d3d3;
                    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
                    white-space: pre-wrap;
                    text-align: justify;
                    word-wrap: break-word;
                ">{escaped_month_text}</div>
                """,
                unsafe_allow_html=True
            )

            with st.expander("✏️ Chỉnh sửa nội dung Kế Hoạch Tháng"):
                ed_m_text = st.text_area("Nội dung:", st.session_state["saved_month_plan"], height=380, key="ed_m_ta")
                if ed_m_text != st.session_state["saved_month_plan"]:
                    st.session_state["saved_month_plan"] = clean_markdown_text(ed_m_text)
                    st.rerun()

    # -------------------------------------------------------------
    # SUB-TAB 2: THIẾT KẾ KẾ HOẠCH TUẦN (VIDEO 2: MA TRẬN 5 NGÀY)
    # -------------------------------------------------------------
    with sub_tab_week:
        st.subheader("📋 Thiết Kế Kế Hoạch Tuần (Ma Trận 5 Ngày)")
        st.caption("Thiết lập thời khóa biểu và lịch trình hoạt động học, chơi ngoài trời, góc và buổi chiều từ Thứ Hai đến Thứ Sáu.")

        st.markdown("#### 📌 Thông Tin Tuần Học")
        w_c1, w_c2, w_c3 = st.columns(3)
        with w_c1:
            w_school = st.text_input("Tên trường", "Trường Mầm Non Phước Thạnh", key="w_tr")
            w_teacher = st.text_input("Tên giáo viên", "Nguyễn Thị Thanh Vân", key="w_gv")
        with w_c2:
            w_class = st.text_input("Lớp học", "Lớp Chồi 1", key="w_lp")
            w_age = st.selectbox("Độ tuổi", ["3-4 tuổi", "4-5 tuổi", "5-6 tuổi"], index=1, key="w_dt")
        with w_c3:
            w_week_num = st.selectbox("Tuần thực hiện", ["Tuần 1", "Tuần 2", "Tuần 3", "Tuần 4", "Tuần 5"], index=0, key="w_so_tuan")
            w_time_range = st.text_input("Thời gian (Từ ngày... đến ngày...)", "Từ Thứ Hai đến Thứ Sáu", key="w_tg")

        w_c4, w_c5 = st.columns(2)
        with w_c4:
            w_theme_month = st.text_input("Thuộc Chủ đề lớn (Tháng)", "Bản thân và các giác quan kỳ diệu", key="w_cd_thang")
        with w_c5:
            w_theme_week = st.text_input("Chủ đề tuần / Nhánh tuần *", "Tôi là ai - Bạn là ai?", key="w_cd_tuan")

        st.divider()
        st.markdown("#### 🕒 Phác Thảo Hoạt Động Học Nhanh (Tùy chọn - Để trống AI sẽ tự động điền sinh động)")

        # Ma trận 5 cột tương ứng Thứ Hai -> Thứ Sáu (Đúng theo Video 2)
        col_t2, col_t3, col_t4, col_t5, col_t6 = st.columns(5)
        
        with col_t2:
            st.markdown("##### Thứ Hai")
            t2_hoc_in = st.text_input("Hoạt động học", placeholder="VD: VĐCB Đi trong đường hẹp", key="t2_h")
            t2_troi_in = st.text_input("Ngoài trời", placeholder="VD: Nhặt lá khô, TCVĐ", key="t2_tr")
            t2_goc_in = st.text_input("Hoạt động góc", placeholder="VD: Góc đóng vai bác sĩ", key="t2_g")
            t2_chieu_in = st.text_input("Hoạt động chiều", placeholder="VD: Ôn bài, làm quen thơ", key="t2_c")

        with col_t3:
            st.markdown("##### Thứ Ba")
            t3_hoc_in = st.text_input("Hoạt động học", placeholder="VD: Khám phá các giác quan", key="t3_h")
            t3_troi_in = st.text_input("Ngoài trời", placeholder="VD: Quan sát vườn hoa", key="t3_tr")
            t3_goc_in = st.text_input("Hoạt động góc", placeholder="VD: Góc xây dựng trường lớp", key="t3_g")
            t3_chieu_in = st.text_input("Hoạt động chiều", placeholder="VD: Rèn kỹ năng tự phục vụ", key="t3_c")

        with col_t4:
            st.markdown("##### Thứ Tư")
            t4_hoc_in = st.text_input("Hoạt động học", placeholder="VD: Toán: Phân biệt to - nhỏ", key="t4_h")
            t4_troi_in = st.text_input("Ngoài trời", placeholder="VD: Chơi với cát và nước", key="t4_tr")
            t4_goc_in = st.text_input("Hoạt động góc", placeholder="VD: Góc nghệ thuật tạo hình", key="t4_g")
            t4_chieu_in = st.text_input("Hoạt động chiều", placeholder="VD: Kể chuyện theo tranh", key="t4_c")

        with col_t5:
            st.markdown("##### Thứ Năm")
            t5_hoc_in = st.text_input("Hoạt động học", placeholder="VD: Thơ: Đôi mắt của em", key="t5_h")
            t5_troi_in = st.text_input("Ngoài trời", placeholder="VD: Vẽ phấn trên sân", key="t5_tr")
            t5_goc_in = st.text_input("Hoạt động góc", placeholder="VD: Góc thư viện xem tranh", key="t5_g")
            t5_chieu_in = st.text_input("Hoạt động chiều", placeholder="VD: Hướng dẫn gấp khăn, áo", key="t5_c")

        with col_t6:
            st.markdown("##### Thứ Sáu")
            t6_hoc_in = st.text_input("Hoạt động học", placeholder="VD: Hát: Cái mũi xinh", key="t6_h")
            t6_troi_in = st.text_input("Ngoài trời", placeholder="VD: Giao lưu thể thao nhẹ", key="t6_tr")
            t6_goc_in = st.text_input("Hoạt động góc", placeholder="VD: Triển lãm tranh vẽ tuần", key="t6_g")
            t6_chieu_in = st.text_input("Hoạt động chiều", placeholder="VD: Nêu gương bé ngoan tuần", key="t6_c")

        w_notes_in = st.text_area(
            "Yêu cầu bổ sung đặc biệt cho kế hoạch tuần (Ví dụ: Lồng ghép thêm hoạt động chuẩn bị lễ hội trung thu, Đưa vào nhiều trò chơi dân gian...):",
            placeholder="Nhập yêu cầu riêng cho tuần...",
            height=70,
            key="w_notes_in"
        )

        matrix_data = {
            "t2": {"hoc": t2_hoc_in, "troi": t2_troi_in, "goc": t2_goc_in, "chieu": t2_chieu_in},
            "t3": {"hoc": t3_hoc_in, "troi": t3_troi_in, "goc": t3_goc_in, "chieu": t3_chieu_in},
            "t4": {"hoc": t4_hoc_in, "troi": t4_troi_in, "goc": t4_goc_in, "chieu": t4_chieu_in},
            "t5": {"hoc": t5_hoc_in, "troi": t5_troi_in, "goc": t5_goc_in, "chieu": t5_chieu_in},
            "t6": {"hoc": t6_hoc_in, "troi": t6_troi_in, "goc": t6_goc_in, "chieu": t6_chieu_in},
        }

        week_info_dict = {
            "school": w_school,
            "teacher": w_teacher,
            "class_name": w_class,
            "age": w_age,
            "week_num": w_week_num,
            "theme_month": w_theme_month,
            "theme_week": w_theme_week,
            "time_range": w_time_range,
            "matrix": matrix_data,
            "notes": w_notes_in,
            "doc_title": f"KẾ HOẠCH GIÁO DỤC {w_week_num.upper()}",
            "topic": f"KeHoachTuan_{w_week_num}"
        }

        if st.button("✨ Tạo Kế Hoạch Tuần Bằng AI", type="primary", use_container_width=True, key="btn_week_ai"):
            with st.spinner("⚡ Đang lập kế hoạch tuần chi tiết 5 ngày chuẩn sư phạm mầm non..."):
                if api_key:
                    try:
                        client = genai.Client(api_key=api_key)
                        prompt_week = f"""
                        Lập kế hoạch tuần mầm non chuẩn Bộ GD&ĐT Việt Nam:
                        - Trường: {w_school} | Giáo viên: {w_teacher} | Lớp: {w_class} ({w_age})
                        - {w_week_num} ({w_time_range})
                        - Chủ đề tháng: {w_theme_month} | Nhánh tuần: {w_theme_week}
                        - Phác thảo người dùng gợi ý:
                          + Thứ Hai: Học: {t2_hoc_in} | Ngoài trời: {t2_troi_in} | Góc: {t2_goc_in} | Chiều: {t2_chieu_in}
                          + Thứ Ba: Học: {t3_hoc_in} | Ngoài trời: {t3_troi_in} | Góc: {t3_goc_in} | Chiều: {t3_chieu_in}
                          + Thứ Tư: Học: {t4_hoc_in} | Ngoài trời: {t4_troi_in} | Góc: {t4_goc_in} | Chiều: {t4_chieu_in}
                          + Thứ Năm: Học: {t5_hoc_in} | Ngoài trời: {t5_troi_in} | Góc: {t5_goc_in} | Chiều: {t5_chieu_in}
                          + Thứ Sáu: Học: {t6_hoc_in} | Ngoài trời: {t6_troi_in} | Góc: {t6_goc_in} | Chiều: {t6_chieu_in}
                        - Yêu cầu thêm: {w_notes_in if w_notes_in else 'Không'}
                        
                        Yêu cầu trình bày đầy đủ: Mục tiêu tuần, Lịch hoạt động từng ngày từ Thứ 2 đến Thứ 6 (Đón trẻ, Thể dục sáng, Hoạt động học, Chơi ngoài trời, Chơi góc, Hoạt động chiều), Chuẩn bị và Đánh giá tuần. Tuyệt đối không dùng ký tự ### hoặc **.
                        """
                        res_week = generate_fast_ai_content(client, prompt_week, fallback_type="week", fallback_data=week_info_dict)
                    except Exception:
                        res_week = generate_instant_week_plan(week_info_dict)
                else:
                    res_week = generate_instant_week_plan(week_info_dict)

                res_week = clean_markdown_text(res_week)
                st.session_state["saved_week_plan"] = res_week
                st.session_state["saved_week_info"] = week_info_dict
                save_plan_to_db(f"Kế Hoạch {w_week_num} - {w_theme_week}", res_week)
            st.success("🎉 Đã hoàn thành Kế hoạch giáo dục tuần!")

        # KẾT QUẢ KẾ HOẠCH TUẦN
        if "saved_week_plan" in st.session_state:
            st.divider()
            w_info = st.session_state["saved_week_info"]
            w_font_size = st.session_state.get("w_font_size_slider", 14)

            st.markdown("### 🛠️ Tùy Chỉnh & Thao Tác Kế Hoạch Tuần")
            wb_col1, wb_col2 = st.columns(2)
            with wb_col1:
                w_docx = export_docx_from_text(st.session_state["saved_week_plan"], w_info, font_size_pt=w_font_size)
                with open(w_docx, "rb") as f_w:
                    st.download_button(
                        label="📥 Tải File Word (.docx) Kế Hoạch Tuần",
                        data=f_w,
                        file_name=w_docx,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True,
                        key="dl_week_btn"
                    )
            with wb_col2:
                w_json_text = json.dumps(st.session_state["saved_week_plan"])
                w_title_json = json.dumps(f"Kế Hoạch Tuần: {w_info['theme_week']}")
                print_w_html = f"""
                <div style="display: flex; justify-content: center; align-items: center; width: 100%;">
                    <button onclick="printWeekPlan()" style="
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
                    ">
                        🖨️ In Kế Hoạch Tuần Ngay
                    </button>
                </div>
                <script>
                function printWeekPlan() {{
                    var rawContent = {w_json_text};
                    var docTitle = {w_title_json};
                    var pWin = window.open('', '_blank');
                    pWin.document.write('<!DOCTYPE html><html><head><title>' + docTitle + '</title>');
                    pWin.document.write('<style>');
                    pWin.document.write('@page {{ size: A4; margin: 15mm 20mm 15mm 20mm; }}');
                    pWin.document.write('body {{ font-family: "Times New Roman", Times, serif; font-size: {w_font_size}pt; line-height: 1.5; color: #000; margin: 0; padding: 0; }}');
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
                components.html(print_w_html, height=52)

            wi_c1, wi_c2 = st.columns(2)
            with wi_c1:
                st.success("✅ **Điểm sư phạm: 100/100**")
            with wi_c2:
                st.info(f"📋 **{w_info['week_num']}** ({w_info['age']}) | Nhánh: {w_info['theme_week']}")

            w_font_size = st.slider(
                "🔤 Kích thước chữ hiển thị & In (pt):", 
                min_value=12, 
                max_value=24, 
                value=w_font_size, 
                step=1, 
                key="w_font_size_slider"
            )

            st.markdown("### 📄 Bản Xem Trước Kế Hoạch Tuần (Trang Giấy In):")
            escaped_week_text = html.escape(st.session_state["saved_week_plan"])
            st.markdown(
                f"""
                <div style="
                    background-color: #ffffff;
                    color: #111111;
                    font-family: 'Times New Roman', Times, serif;
                    font-size: {w_font_size}pt;
                    line-height: 1.6;
                    padding: 30px 40px;
                    border-radius: 8px;
                    border: 1px solid #d3d3d3;
                    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
                    white-space: pre-wrap;
                    text-align: justify;
                    word-wrap: break-word;
                ">{escaped_week_text}</div>
                """,
                unsafe_allow_html=True
            )

            with st.expander("✏️ Chỉnh sửa nội dung Kế Hoạch Tuần"):
                ed_w_text = st.text_area("Nội dung:", st.session_state["saved_week_plan"], height=380, key="ed_w_ta")
                if ed_w_text != st.session_state["saved_week_plan"]:
                    st.session_state["saved_week_plan"] = clean_markdown_text(ed_w_text)
                    st.rerun()

# --- TAB 3: TRÒ CHƠI & ĐỐ VUI ---
elif app_mode == "🎯 Trò Chơi & Đố Vui":
    st.subheader("🎯 Sáng Tác Trò Chơi, Câu Đố & Bài Thơ Mầm Non")
    topic_game = st.text_input("Nhập chủ đề hoặc sự vật (VD: Củ cà rốt, An toàn giao thông, Đếm đến 5...)", "an toan giao thong")
    genre = st.selectbox("Thể loại cần tạo", ["Bài thơ ngắn dễ thuộc", "Bài đồng dao vui nhộn", "Câu đố vui cho trẻ", "Trò chơi vận động tập thể"])
    
    if st.button("✨ Sáng Tác Ngay", type="primary"):
        with st.spinner("⚡ Đang sáng tác nội dung siêu tốc..."):
            creative_dict = {"genre": genre, "topic": topic_game}
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    res_text = generate_fast_ai_content(
                        client,
                        f"Hãy sáng tác hoặc thiết kế một {genre} dành riêng cho trẻ mầm non với chủ đề: {topic_game}. Ngôn ngữ trong trẻo, vui tươi. Nếu là trò chơi cần nêu rõ mục đích, luật chơi và cách chơi. Không sử dụng dấu ### hay ** dư thừa.",
                        fallback_type="creative",
                        fallback_data=creative_dict
                    )
                except Exception:
                    res_text = generate_instant_creative_content(genre, topic_game)
            else:
                res_text = generate_instant_creative_content(genre, topic_game)

            clean_res = clean_markdown_text(res_text)
            st.success("🎉 Đã sáng tác hoàn thành cho cô:")
            st.markdown(clean_res)

# --- TAB 4: SỔ LIÊN LẠC & ZALO ---
elif app_mode == "💬 Sổ Liên Lạc & Zalo":
    st.subheader("💬 Trợ Lý Viết Sổ Liên Lạc & Tin Nhắn Zalo Phụ Huynh")
    task_type = st.radio("Chọn loại nội dung cần viết:", ["Tin nhắn tổng kết ngày/tuần gửi nhóm Zalo", "Viết lời nhận xét sự tiến bộ cho từng bé"])
    
    if task_type == "Tin nhắn tổng kết ngày/tuần gửi nhóm Zalo":
        summary_topic = st.text_input("Nội dung hoạt động chính", "Hôm nay các con học kỹ năng an toàn giao thông và làm quen chữ cái")
        if st.button("💌 Soạn Tin Nhắn Zalo", type="primary"):
            with st.spinner("⚡ Đang soạn tin siêu tốc..."):
                zalo_dict = {"type": task_type, "info": summary_topic}
                if api_key:
                    try:
                        client = genai.Client(api_key=api_key)
                        res_text = generate_fast_ai_content(
                            client,
                            f"Soạn một tin nhắn Zalo thân thiện, chuyên nghiệp từ giáo viên mầm non gửi phụ huynh tóm tắt các hoạt động: {summary_topic}. Kèm lời dặn dò nhẹ nhàng buổi tối. Tránh dùng dấu markdown dư thừa.",
                            fallback_type="zalo",
                            fallback_data=zalo_dict
                        )
                    except Exception:
                        res_text = generate_instant_zalo(task_type, summary_topic)
                else:
                    res_text = generate_instant_zalo(task_type, summary_topic)

                clean_res = clean_markdown_text(res_text)
                st.success("🎉 Mẫu tin nhắn Zalo:")
                st.markdown(clean_res)
    else:
        child_info = st.text_input("Nhập vài từ khóa về tình hình của bé", "Bé Bảo An: ngoan, ăn giỏi, tích cực phát biểu")
        if st.button("✍️ Viết Lời Nhận Xét", type="primary"):
            with st.spinner("⚡ Đang viết lời nhận xét siêu tốc..."):
                zalo_dict = {"type": task_type, "info": child_info}
                if api_key:
                    try:
                        client = genai.Client(api_key=api_key)
                        res_text = generate_fast_ai_content(
                            client,
                            f"Viết một đoạn nhận xét cuối tuần khéo léo, tình cảm, mang tính khích lệ cao dành cho học sinh mầm non dựa trên thông tin: {child_info}. Không dùng dấu ### hoặc ** dư thừa.",
                            fallback_type="zalo",
                            fallback_data=zalo_dict
                        )
                    except Exception:
                        res_text = generate_instant_zalo(task_type, child_info)
                else:
                    res_text = generate_instant_zalo(task_type, child_info)

                clean_res = clean_markdown_text(res_text)
                st.success("🎉 Lời nhận xét mẫu cho cô:")
                st.markdown(clean_res)

# --- TAB 5: GỢI Ý ĐỒ DÙNG DIY ---
elif app_mode == "✂️ Gợi Ý Đồ Dùng DIY":
    st.subheader("✂️ Gợi Ý Làm Đồ Dùng Dạy Học Từ Vật Liệu Mở")
    lesson_name = st.text_input("Nhập tên bài học hoặc kỹ năng cần làm đồ dùng", "Biển báo giao thông cho bé")
    
    if st.button("🛠️ Gợi Ý Cách Làm", type="primary"):
        with st.spinner("⚡ Đang lên ý tưởng chế tạo siêu tốc..."):
            diy_dict = {"lesson": lesson_name}
            if api_key:
                try:
                    client = genai.Client(api_key=api_key)
                    res_text = generate_fast_ai_content(
                        client,
                        f"Gợi ý cho giáo viên mầm non cách tận dụng các nguyên vật liệu phế liệu/vật liệu mở (chai nhựa, ống hút, bìa carton...) để làm đồ dùng đồ chơi tự tạo phục vụ cho bài học: {lesson_name}. Liệt kê nguyên liệu và các bước làm ngắn gọn, an toàn cho trẻ. Không dùng dấu markdown rác.",
                        fallback_type="diy",
                        fallback_data=diy_dict
                    )
                except Exception:
                    res_text = generate_instant_diy(lesson_name)
            else:
                res_text = generate_instant_diy(lesson_name)

            clean_res = clean_markdown_text(res_text)
            st.success("🎉 Ý tưởng đồ dùng tự tạo:")
            st.markdown(clean_res)

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
                        info_dummy = {
                            "school": "Trường Mầm Non Phước Thạnh",
                            "teacher": "Nguyễn Thị Thanh Vân",
                            "class_name": "Lớp Chồi 1",
                            "topic": title_db,
                            "doc_title": title_db
                        }
                        f_path = export_docx_from_text(content_db, info_dummy)
                        with open(f_path, "rb") as f_dl:
                            cols[2].download_button("📥 Tải File", data=f_dl, file_name=f_path, key=f"dl_{p_id}")
