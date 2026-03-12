import streamlit as st
import pandas as pd
import random
import io
import json
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="EduTest Pro - CV7991 (Form 2025)", page_icon="🎓", layout="wide")

st.markdown("""
<style>
    .stMetric { background-color: #f0f2f6; padding: 10px; border-radius: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 4px 4px 0 0; padding-top: 10px; padding-bottom: 10px; }
    .stTabs [aria-selected="true"] { background-color: #e6f0ff; border-bottom: 3px solid #0066cc; font-weight: bold;}
</style>
""", unsafe_allow_html=True)

# --- KHỞI TẠO SESSION STATE ---
# Thêm cột Loai_cau_hoi (Trắc nghiệm, Đúng/Sai, Tự luận)
if 'db_df' not in st.session_state:
    st.session_state.db_df = pd.DataFrame(columns=['Chuong', 'Muc_do', 'Loai_cau_hoi', 'Noi_dung', 'A', 'B', 'C', 'D', 'Dap_an_dung'])
if 'matrix_df' not in st.session_state:
    st.session_state.matrix_df = pd.DataFrame()

# --- HÀM TIỆN ÍCH ---
def style_text(paragraph, text, bold=False, italic=False):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    return run

# --- HÀM GỌI GEMINI AI TẠO CÂU HỎI (NÂNG CẤP ĐA ĐỊNH DẠNG) ---
def generate_questions_with_ai(api_key, subject, chapter, q_tn, q_ds, q_tl):
    prompt = f"""
    Bạn là một chuyên gia giáo dục tại Việt Nam. Hãy soạn câu hỏi môn {subject}, phần/chủ đề "{chapter}" bám sát cấu trúc đề thi mới nhất (Form 2025).
    Số lượng yêu cầu:
    - {q_tn} câu Trắc nghiệm 4 lựa chọn (Loại: Trắc nghiệm)
    - {q_ds} câu Trắc nghiệm Đúng/Sai (Loại: Đúng/Sai) - Mỗi câu gồm 1 đoạn dẫn chung và 4 ý a, b, c, d.
    - {q_tl} câu Tự luận / Trả lời ngắn (Loại: Tự luận)
    
    Quy tắc cấu trúc JSON trả về (TUYỆT ĐỐI KHÔNG chứa markdown như ```json):
    [
        {{
            "Chuong": "{chapter}", "Muc_do": "NB", "Loai_cau_hoi": "Trắc nghiệm",
            "Noi_dung": "Nội dung câu hỏi...", "A": "...", "B": "...", "C": "...", "D": "...", "Dap_an_dung": "A"
        }},
        {{
            "Chuong": "{chapter}", "Muc_do": "VD", "Loai_cau_hoi": "Đúng/Sai",
            "Noi_dung": "Đoạn dẫn/Đề bài chung...", "A": "Ý a...", "B": "Ý b...", "C": "Ý c...", "D": "Ý d...", "Dap_an_dung": "Đ, S, Đ, S"
        }},
        {{
            "Chuong": "{chapter}", "Muc_do": "VDC", "Loai_cau_hoi": "Tự luận",
            "Noi_dung": "Câu hỏi tự luận...", "A": "", "B": "", "C": "", "D": "", "Dap_an_dung": "Đáp án chi tiết..."
        }}
    ]
    """
    
    headers = {'Content-Type': 'application/json'}
    data = {
        "contents": [{"parts":[{"text": prompt}]}],
        "generationConfig": {"response_mime_type": "application/json"}
    }
    
    models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-flash-8b", "gemini-1.5-pro"]
    last_error = ""
    for model in models:
        url = f"[https://generativelanguage.googleapis.com/v1beta/models/](https://generativelanguage.googleapis.com/v1beta/models/){model}:generateContent?key={api_key}"
        try:
            response = requests.post(url, headers=headers, json=data)
            if response.status_code == 200:
                result = response.json()['candidates'][0]['content']['parts'][0]['text']
                result = result.replace('```json', '').replace('```', '').strip()
                return json.loads(result)
            elif response.status_code == 404:
                continue
            else:
                last_error = response.text
        except Exception as e:
            last_error = str(e)
            continue
            
    st.error(f"❌ Lỗi kết nối AI: {last_error}")
    return None

# --- HÀM TRỘN CÂU HỎI (Chỉ trộn trắc nghiệm thường) ---
def shuffle_question(row):
    loai = str(row.get('Loai_cau_hoi', 'Trắc nghiệm')).strip()
    
    # Nếu là Đúng/Sai hoặc Tự luận thì KHÔNG trộn phương án (giữ nguyên a,b,c,d)
    if loai != 'Trắc nghiệm':
        return {
            'Noi_dung': row['Noi_dung'],
            'Options': {'A': row.get('A',''), 'B': row.get('B',''), 'C': row.get('C',''), 'D': row.get('D','')},
            'Correct': row['Dap_an_dung'],
            'Muc_do': row['Muc_do'],
            'Loai': loai,
            'Chuong': row.get('Chuong', 'Chung')
        }
        
    # Xử lý trộn câu trắc nghiệm 4 lựa chọn
    options = [('A', str(row['A'])), ('B', str(row['B'])), ('C', str(row['C'])), ('D', str(row['D']))]
    correct_content = str(row[row['Dap_an_dung']])
    random.shuffle(options)
    
    new_options = {}
    new_correct_label = ""
    labels = ['A', 'B', 'C', 'D']
    for i in range(4):
        new_label = labels[i]
        content = options[i][1]
        new_options[new_label] = content
        if content == correct_content:
            new_correct_label = new_label
            
    return {
        'Noi_dung': row['Noi_dung'],
        'Options': new_options,
        'Correct': new_correct_label,
        'Muc_do': row['Muc_do'],
        'Loai': loai,
        'Chuong': row.get('Chuong', 'Chung')
    }

# --- HÀM XUẤT WORD CHIA 3 PHẦN ---
def export_full_exam(exam_list, info):
    doc = Document()
    
    # Header
    header_table = doc.add_table(rows=1, cols=2)
    header_table.width = Inches(6)
    left_cell = header_table.cell(0, 0).paragraphs[0]
    style_text(left_cell, f"{info['school'].upper()}\n", bold=True)
    style_text(left_cell, f"GV: {info['teacher']}", italic=True)
    
    right_cell = header_table.cell(0, 1).paragraphs[0]
    right_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_text(right_cell, f"ĐỀ KIỂM TRA MÔN {info['subject'].upper()}\n", bold=True)
    style_text(right_cell, f"Thời gian làm bài: {info['time']} phút")

    doc.add_paragraph("_" * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"ĐỀ KIỂM TRA MÔN {info['subject'].upper()}").runs[0].bold = True
    
    # Tách danh sách câu hỏi theo loại
    list_tn = [q for q in exam_list if q['Loai'] == 'Trắc nghiệm']
    list_ds = [q for q in exam_list if q['Loai'] == 'Đúng/Sai']
    list_tl = [q for q in exam_list if q['Loai'] == 'Tự luận']

    # --- PHẦN I: TRẮC NGHIỆM ---
    if list_tn:
        doc.add_paragraph("\nPHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN.").runs[0].bold = True
        for i, q in enumerate(list_tn):
            p = doc.add_paragraph()
            style_text(p, f"Câu {i+1}: ", bold=True)
            style_text(p, str(q['Noi_dung']))
            for label in ['A', 'B', 'C', 'D']:
                opt_p = doc.add_paragraph()
                opt_p.paragraph_format.left_indent = Inches(0.3)
                style_text(opt_p, f"{label}. ", bold=True)
                style_text(opt_p, str(q['Options'][label]))

    # --- PHẦN II: ĐÚNG/SAI ---
    if list_ds:
        doc.add_paragraph("\nPHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI.").runs[0].bold = True
        for i, q in enumerate(list_ds):
            p = doc.add_paragraph()
            style_text(p, f"Câu {i+1}: ", bold=True)
            style_text(p, str(q['Noi_dung']))
            for label, sub in zip(['A', 'B', 'C', 'D'], ['a', 'b', 'c', 'd']):
                if q['Options'][label] and str(q['Options'][label]).strip() not in ['nan', 'None', '']:
                    opt_p = doc.add_paragraph()
                    opt_p.paragraph_format.left_indent = Inches(0.3)
                    style_text(opt_p, f"{sub}) ", bold=True)
                    style_text(opt_p, str(q['Options'][label]))

    # --- PHẦN III: TỰ LUẬN ---
    if list_tl:
        doc.add_paragraph("\nPHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN / TỰ LUẬN.").runs[0].bold = True
        for i, q in enumerate(list_tl):
            p = doc.add_paragraph()
            style_text(p, f"Câu {i+1}: ", bold=True)
            style_text(p, str(q['Noi_dung']))
            doc.add_paragraph("\n\n") # Khoảng trống để học sinh làm bài

    # --- BẢNG ĐÁP ÁN ---
    doc.add_page_break()
    doc.add_paragraph("BẢNG ĐÁP ÁN & HƯỚNG DẪN CHẤM").runs[0].bold = True
    
    if list_tn:
        doc.add_paragraph("Phần I. Trắc nghiệm").runs[0].italic = True
        ans_table = doc.add_table(rows=2, cols=len(list_tn))
        ans_table.style = 'Table Grid'
        for i, q in enumerate(list_tn):
            ans_table.cell(0, i).text = str(i+1)
            ans_table.cell(1, i).text = str(q['Correct'])

    if list_ds:
        doc.add_paragraph("\nPhần II. Đúng/Sai").runs[0].italic = True
        ans_table_ds = doc.add_table(rows=5, cols=len(list_ds) + 1)
        ans_table_ds.style = 'Table Grid'
        ans_table_ds.cell(0, 0).text = "Câu"
        ans_table_ds.cell(1, 0).text = "Ý a"
        ans_table_ds.cell(2, 0).text = "Ý b"
        ans_table_ds.cell(3, 0).text = "Ý c"
        ans_table_ds.cell(4, 0).text = "Ý d"
        for i, q in enumerate(list_ds):
            ans_table_ds.cell(0, i+1).text = str(i+1)
            # Tách chuỗi "Đ, S, Đ, S"
            ans_arr = str(q['Correct']).replace(' ', '').split(',')
            for j in range(4):
                if j < len(ans_arr):
                    ans_table_ds.cell(j+1, i+1).text = ans_arr[j]

    if list_tl:
        doc.add_paragraph("\nPhần III. Tự luận/Trả lời ngắn").runs[0].italic = True
        for i, q in enumerate(list_tl):
            doc.add_paragraph(f"Câu {i+1}: {q['Correct']}")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- UI CHÍNH ---
st.title("🎓 Hệ Thống Tạo Đề Đa Định Dạng (Form 2025)")

with st.sidebar:
    st.image("[https://cdn-icons-png.flaticon.com/512/3074/3074058.png](https://cdn-icons-png.flaticon.com/512/3074/3074058.png)", width=100)
    st.header("⚙️ Thông tin chung")
    school = st.text_input("Tên trường", "THCS Đa Phước")
    teacher = st.text_input("Tên giáo viên", "Tống Phước Khải")
    subject = st.text_input("Môn học", "Toán")
    duration = st.number_input("Thời gian (phút)", value=90)
    
    st.divider()
    st.markdown("🔑 **Khóa API AI (Tùy chọn)**")
    api_key = st.text_input("Nhập Gemini API Key", type="password")

tab1, tab2, tab3 = st.tabs(["📁 1. Dữ liệu", "✨ 2. Sinh Câu Hỏi AI", "🖨️ 3. Lọc & Xuất Đề"])

with tab1:
    st.subheader("Tải lên Ngân hàng câu hỏi (Excel)")
    st.caption("File cần có các cột: Chuong, Muc_do, Loai_cau_hoi, Noi_dung, A, B, C, D, Dap_an_dung")
    uploaded_file = st.file_uploader("", type=["xlsx"])
    if st.button("Tải vào hệ thống", type="primary"):
        if uploaded_file:
            new_df = pd.read_excel(uploaded_file)
            # Xử lý tương thích ngược: Nếu Excel cũ không có cột Loai_cau_hoi thì mặc định là Trắc nghiệm
            if 'Loai_cau_hoi' not in new_df.columns:
                new_df['Loai_cau_hoi'] = 'Trắc nghiệm'
            new_df['Loai_cau_hoi'] = new_df['Loai_cau_hoi'].fillna('Trắc nghiệm')
            
            st.session_state.db_df = pd.concat([st.session_state.db_df, new_df], ignore_index=True).drop_duplicates()
            st.success("Đã bổ sung câu hỏi vào Kho dữ liệu chung!")
            
    if not st.session_state.db_df.empty:
        st.info(f"📚 Tổng: **{len(st.session_state.db_df)}** câu")
        st.dataframe(st.session_state.db_df[['Loai_cau_hoi', 'Muc_do', 'Noi_dung']].head(5))
        if st.button("🗑️ Xóa sạch bộ nhớ", type="secondary"):
            st.session_state.db_df = pd.DataFrame(columns=['Chuong', 'Muc_do', 'Loai_cau_hoi', 'Noi_dung', 'A', 'B', 'C', 'D', 'Dap_an_dung'])
            st.rerun()

with tab2:
    st.subheader("Trợ lý AI - Tạo mới câu hỏi đa định dạng")
    
    chapter_ai = st.text_input("Nhập tên Chủ đề / Chương:")
    
    col_tn, col_ds, col_tl = st.columns(3)
    ai_tn = col_tn.number_input("Số câu Trắc nghiệm", min_value=0, value=2)
    ai_ds = col_ds.number_input("Số câu Đúng/Sai", min_value=0, value=1)
    ai_tl = col_tl.number_input("Số câu Tự luận", min_value=0, value=1)
        
    if st.button("✨ XUẤT XƯỞNG CÂU HỎI MỚI", type="primary"):
        if not api_key:
            st.error("❌ Nhập API Key ở Menu trái trước.")
        elif not chapter_ai:
            st.error("❌ Nhập tên Chủ đề cần tạo.")
        else:
            with st.spinner('🤖 AI đang soạn thảo 3 loại câu hỏi...'):
                gen_data = generate_questions_with_ai(api_key, subject, chapter_ai, ai_tn, ai_ds, ai_tl)
                if gen_data:
                    new_ai_df = pd.DataFrame(gen_data)
                    st.session_state.db_df = pd.concat([st.session_state.db_df, new_ai_df], ignore_index=True)
                    st.success("🎉 Đã tạo thành công đa định dạng!")
                    st.dataframe(new_ai_df)

with tab3:
    if st.session_state.db_df.empty:
        st.warning("⚠️ Kho dữ liệu trống.")
    else:
        st.subheader("Bốc đề & Xuất bản (Cấu trúc 2025)")
        st.markdown("Chọn số lượng câu hỏi bạn muốn lấy từ kho lưu trữ để tạo thành 1 đề hoàn chỉnh:")
        
        c1, c2, c3 = st.columns(3)
        # Đếm số lượng hiện có trong kho
        co_tn = len(st.session_state.db_df[st.session_state.db_df['Loai_cau_hoi'] == 'Trắc nghiệm'])
        co_ds = len(st.session_state.db_df[st.session_state.db_df['Loai_cau_hoi'] == 'Đúng/Sai'])
        co_tl = len(st.session_state.db_df[st.session_state.db_df['Loai_cau_hoi'] == 'Tự luận'])
        
        lay_tn = c1.number_input(f"Câu Trắc nghiệm (Kho có: {co_tn})", min_value=0, max_value=co_tn, value=min(12, co_tn))
        lay_ds = c2.number_input(f"Câu Đúng/Sai (Kho có: {co_ds})", min_value=0, max_value=co_ds, value=min(2, co_ds))
        lay_tl = c3.number_input(f"Câu Tự luận (Kho có: {co_tl})", min_value=0, max_value=co_tl, value=min(2, co_tl))

        st.divider()
        if st.button("🚀 BỐC ĐỀ & TẠO FILE WORD", type="primary", use_container_width=True):
            with st.spinner('Đang trộn phương án và xuất file...'):
                exam_list = []
                
                # Bốc ngẫu nhiên từng loại
                if lay_tn > 0:
                    pool_tn = st.session_state.db_df[st.session_state.db_df['Loai_cau_hoi'] == 'Trắc nghiệm'].sample(n=lay_tn)
                    for _, row in pool_tn.iterrows(): exam_list.append(shuffle_question(row))
                        
                if lay_ds > 0:
                    pool_ds = st.session_state.db_df[st.session_state.db_df['Loai_cau_hoi'] == 'Đúng/Sai'].sample(n=lay_ds)
                    for _, row in pool_ds.iterrows(): exam_list.append(shuffle_question(row))
                        
                if lay_tl > 0:
                    pool_tl = st.session_state.db_df[st.session_state.db_df['Loai_cau_hoi'] == 'Tự luận'].sample(n=lay_tl)
                    for _, row in pool_tl.iterrows(): exam_list.append(shuffle_question(row))
                
                # Trộn thứ tự các câu hỏi bên trong mỗi phần (không trộn lẫn các phần với nhau)
                # Nhưng hàm xuất Word đã tự động gom nhóm theo Loai_cau_hoi nên ta cứ lưu thẳng
                info = {'school': school, 'teacher': teacher, 'subject': subject, 'time': duration}
                docx_file = export_full_exam(exam_list, info)
                
                st.balloons()
                st.success("✅ Đã tạo thành công Đề thi 3 Phần!")
                st.download_button(
                    label="📥 TẢI ĐỀ KIỂM TRA MỚI (.docx)",
                    data=docx_file,
                    file_name=f"De_Kiem_Tra_Form2025_{subject}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
